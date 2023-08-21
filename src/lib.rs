mod error;
mod numbering;
mod types;
mod utils;

use wasm_bindgen;
use wasm_bindgen::prelude::*;

use base64::{engine::general_purpose, Engine};
use docx_rs::{
    BreakType, DocumentChild, DrawingData, HyperlinkData, Paragraph, ParagraphChild,
    ParagraphProperty, RunChild, RunProperty,
};

use error::DocError;
use numbering::{ListRelation, NumberingMeta, NumberingType};
use serde::Serialize;
use types::{Chunk, ChunkType, Properties, Px, Spacing};

#[derive(Serialize)]
#[wasm_bindgen]
pub struct DocxDocument {
    chunks: Vec<Chunk>,
    #[serde(skip_serializing)]
    docx: docx_rs::Docx,
    id_counter: usize,
}

use gloo_utils::format::JsValueSerdeExt;

#[wasm_bindgen]
impl DocxDocument {
    pub fn new(bytes: Vec<u8>) -> DocxDocument {
        utils::set_panic_hook();

        let docx = docx_rs::read_docx(&bytes).unwrap();
        let mut doc = DocxDocument::from_docx(docx);
        doc.build().unwrap();

        doc
    }

    pub fn get_chunks(self) -> JsValue {
        utils::set_panic_hook();
        JsValue::from_serde(&self.chunks).unwrap()
    }

    fn from_docx(docx: docx_rs::Docx) -> DocxDocument {
        DocxDocument {
            chunks: vec![],
            docx: docx,
            id_counter: 0,
        }
    }

    fn build(&mut self) -> Result<(), DocError> {
        let mut i = 0;

        while i < self.docx.document.children.len() {
            self.parse_index(&mut i)?;
            i += 1;
        }

        Ok(())
    }

    fn parse_index(&mut self, i: &mut usize) -> Result<(), DocError> {
        let chunks = match self.docx.document.children[*i].to_owned() {
            DocumentChild::Paragraph(p) => {
                let (chunk_type, numbering_meta) = self.define_chunk_type(&p)?;

                match chunk_type {
                    ChunkType::Paragraph => self.parse_paragraph(&p, ChunkType::Paragraph)?,
                    ChunkType::Li => {
                        self.parse_numbering_sequence(0, i, &numbering_meta.unwrap(), *i)?
                    }
                    _ => return Err(DocError::new("unknown chunk type")),
                }
            }
            _ => return Ok(()),
        };

        self.chunks.extend(chunks);

        Ok(())
    }

    fn parse_paragraph(&mut self, p: &Paragraph, t: ChunkType) -> Result<Vec<Chunk>, DocError> {
        let mut style_props = Properties::new();
        if let Some(v) = &p.property.style {
            // load properties from the style
            self.load_props_from_style(&v.val, &mut style_props);
        }

        // load properties from inline para style
        let mut para_props = self.override_para_properties(&style_props, &p.property);
        // load properties from inline run style
        let base_run_props = self.override_run_properties(&para_props, &p.property.run_property);

        let mut paragraph = Chunk::new(self.id(), t);

        let (paragraph_children, max_sz) =
            self.parse_block_content(&base_run_props, p.children.iter())?;

        if let Some(s) = &para_props.spacing {
            para_props.line_height = s.calc_line_spacing(max_sz);
        }
        paragraph.set_props(para_props);

        let mut chunks: Vec<Chunk> = vec![];
        chunks.push(paragraph);
        chunks.extend(paragraph_children);
        chunks.push(Chunk::new(self.id(), ChunkType::End));

        Ok(chunks)
    }

    fn parse_numbering_sequence(
        &mut self,
        l: usize,
        from: &mut usize,
        parent_meta: &NumberingMeta,
        from_cp: usize,
    ) -> Result<Vec<Chunk>, DocError> {
        let mut chunks: Vec<Chunk> = vec![];

        while *from < self.docx.document.children.len() {
            let p = match self.docx.document.children[*from].to_owned() {
                DocumentChild::Paragraph(p) => p,
                _ => break,
            };

            let mut num_meta = parent_meta.to_owned();
            let mut chunk_type = ChunkType::Li;

            if *from > from_cp {
                let (t, m) = self.define_chunk_type(&p)?;
                if t != ChunkType::Li {
                    // child is not a list
                    *from -= 1;
                    chunks.push(Chunk::new(self.id(), ChunkType::End));

                    return Ok(chunks);
                }
                num_meta = m.unwrap();
                chunk_type = t;
            } else {
                // push root list chunk (ol/ul)
                chunks.push(Chunk::new(self.id(), parent_meta.to_chunk_type()));
            }

            match numbering::define_list_relation(&num_meta, &parent_meta) {
                ListRelation::ListItem => {
                    let list_chunks = self.parse_paragraph(&p, chunk_type)?;
                    chunks.extend(list_chunks);

                    *from += 1;
                }
                ListRelation::SiblingList => {
                    if l == 0 {
                        // if nesting level is zero, then return root
                        *from -= 1;
                        chunks.push(Chunk::new(self.id(), ChunkType::End));
                        return Ok(chunks);
                    }

                    let list_type = num_meta.to_owned().to_chunk_type();
                    let sibling_list = Chunk::new(self.id(), list_type);
                    let list_children =
                        self.parse_numbering_sequence(l, from, &num_meta, from.clone())?;

                    chunks.push(sibling_list);
                    chunks.extend(list_children);
                    chunks.push(Chunk::new(self.id(), ChunkType::End));
                }
                ListRelation::NestedList => {
                    let list_children =
                        self.parse_numbering_sequence(l + 1, from, &num_meta, from.clone())?;
                    chunks.extend(list_children);
                }
                ListRelation::End => {
                    chunks.push(Chunk::new(self.id(), ChunkType::End));
                    return Ok(chunks);
                }
            }
        }

        return Ok(chunks);
    }

    fn parse_block_content<'a, T>(
        &mut self,
        base_props: &Properties,
        children_iter: T,
    ) -> Result<(Vec<Chunk>, usize), DocError>
    where
        T: Iterator<Item = &'a ParagraphChild>,
    {
        let mut chunks: Vec<Chunk> = vec![];
        let mut max_font_size: usize = 0;

        for child in children_iter {
            match child {
                ParagraphChild::Hyperlink(v) => match &v.link {
                    HyperlinkData::External { rid, path: _ } => {
                        let url = self.find_hyperlink_url(rid);

                        let mut hyperlink = Chunk::new(self.id(), ChunkType::Link);
                        hyperlink.set_url(url.unwrap_or(String::default()));

                        let (children, max_sz) =
                            self.parse_block_content(base_props, v.children.iter())?;

                        chunks.push(hyperlink);
                        chunks.extend(children);
                        chunks.push(Chunk::new(self.id(), ChunkType::End));

                        if max_sz > max_font_size {
                            max_font_size = max_sz;
                        }
                    }
                    _ => (),
                },
                ParagraphChild::Run(run) => {
                    let run_props = self.override_run_properties(base_props, &run.run_property);

                    if let Some(sz) = &run_props.font_size {
                        if sz.get_val() as usize > max_font_size {
                            max_font_size = sz.get_val() as usize;
                        }
                    }

                    for run_child in run.children.iter() {
                        match run_child {
                            RunChild::Drawing(d) => {
                                // check if the element is image
                                if let Some(d) = &d.data {
                                    match d {
                                        DrawingData::Pic(pic) => {
                                            // parse image

                                            let url = self
                                                .find_image_source(&pic.id)
                                                .unwrap_or(String::default());

                                            let w = utils::emu_to_px(pic.size.0 as i32);
                                            let h = utils::emu_to_px(pic.size.1 as i32);

                                            let mut image = Chunk::new(self.id(), ChunkType::Image);
                                            image.set_url(url);
                                            image.set_size(w, h);

                                            chunks.push(image);
                                        }
                                        _ => (),
                                    }
                                }
                            }
                            RunChild::Text(t) => {
                                // parse text

                                let mut text = Chunk::new(self.id(), ChunkType::Text);
                                text.set_props(run_props.to_owned());
                                text.set_text(t.text.clone());

                                chunks.push(text);
                            }
                            RunChild::Tab(_) => {
                                // parse tab
                                let mut tab = Chunk::new(self.id(), ChunkType::Text);
                                tab.set_props(run_props.to_owned());
                                tab.set_text("\t".to_owned());

                                chunks.push(tab);
                            }
                            RunChild::Break(b) => {
                                if b.break_type == BreakType::TextWrapping {
                                    chunks.push(Chunk::new(self.id(), ChunkType::Break));
                                }
                            }
                            _ => (),
                        }
                    }
                }
                _ => (),
            }
        }

        Ok((chunks, max_font_size))
    }

    fn load_props_from_style(&self, id: &String, dest: &mut Properties) {
        let st = self.docx.styles.find_style_by_id(&id).unwrap();
        if let Some(based_on) = &st.based_on {
            self.load_props_from_style(&based_on.val, dest);
        }

        let para_props = self.override_para_properties(dest, &st.paragraph_property);
        let run_props = self.override_run_properties(dest, &st.run_property);

        dest.align = para_props.align;
        dest.indent = para_props.indent;
        dest.spacing = para_props.spacing;

        dest.color = run_props.color;
        dest.background = run_props.background;
        dest.font_size = run_props.font_size;
        dest.font_family = run_props.font_family;
        dest.bold = run_props.bold;
        dest.italic = run_props.italic;
        dest.underline = run_props.underline;
    }

    fn override_para_properties(&self, base: &Properties, props: &ParagraphProperty) -> Properties {
        let mut new_props = base.to_owned();

        if let Some(j) = &props.alignment {
            new_props.align = Some(j.val.to_owned());
        }
        if let Some(indent) = &props.indent {
            if let Some(start_emu) = indent.start {
                let px = utils::indent_to_px(start_emu);
                new_props.indent = Some(Px::new(px));
            }
        }
        if let Some(s) = &props.line_spacing {
            // FIXME adjust spacing or line-height formula
            let mut spacing = Spacing::new(0, 0);
            if let Some(v) = s.before {
                spacing.set_before(utils::indent_to_px(v as i32) as usize);
            }
            if let Some(v) = s.after {
                spacing.set_after(utils::indent_to_px(v as i32) as usize);
            }
            new_props.spacing = Some(spacing);
        }

        new_props
    }

    fn override_run_properties(&self, base: &Properties, props: &RunProperty) -> Properties {
        let mut new_props = base.to_owned();
        if let Some(bold) = &props.bold {
            new_props.bold = Some(bold.val);
        }
        if let Some(italic) = &props.italic {
            new_props.italic = Some(italic.val);
        }
        if let Some(underline) = &props.underline {
            // FIXME check underline kind
            new_props.underline = Some(!underline.val.is_empty());
        }
        if let Some(sz) = &props.sz {
            let sz_px = utils::docx_pt_to_px(sz.val as i32);
            new_props.font_size = Some(Px::new(sz_px));
        }
        if let Some(color) = &props.color {
            new_props.color = Some(color.val.to_owned());
        }
        // if let Some(background) = &props.shading {
        //     // FIXME Not supported by the docx-rs library. Need to add this ability manually
        // }
        if let Some(fonts) = &props.fonts {
            if let Some(v) = &fonts.ascii {
                new_props.font_family = Some(v.to_owned());
            }
        }

        new_props
    }

    fn find_hyperlink_url(&self, rid: &String) -> Option<String> {
        for h in self.docx.hyperlinks.iter() {
            if &h.0 == rid {
                return Some(h.1.to_owned());
            }
        }
        None
    }

    fn find_image_source(&self, rid: &String) -> Option<String> {
        for h in self.docx.images.iter() {
            if &h.0 == rid {
                let data = &h.3 .0;
                if !data.is_empty() {
                    let s = general_purpose::STANDARD.encode(&data);
                    return Some(s);
                }
            }
        }
        None
    }

    fn define_chunk_type(
        &self,
        para: &Paragraph,
    ) -> core::result::Result<(ChunkType, Option<NumberingMeta>), DocError> {
        if !para.has_numbering {
            return Ok((ChunkType::Paragraph, None));
        }

        let num_props = para.property.numbering_property.as_ref().unwrap();

        let id = num_props.id.as_ref().unwrap().id;
        let level = num_props.level.as_ref().unwrap().val;

        match self
            .docx
            .numberings
            .numberings
            .binary_search_by(|x| x.id.cmp(&id))
        {
            Ok(index) => {
                let num = &self.docx.numberings.numberings[index];

                match self
                    .docx
                    .numberings
                    .abstract_nums
                    .binary_search_by(|x| x.id.cmp(&num.abstract_num_id))
                {
                    Ok(index) => {
                        let abstr_num = &self.docx.numberings.abstract_nums[index];
                        let level_props = &abstr_num.levels[level];

                        let marker = match level_props.format.val.as_str() {
                            "decimal" => NumberingType::Ordered,
                            "bullet" => NumberingType::Unordered,
                            _ => {
                                return Ok((ChunkType::Paragraph, None));
                            }
                        };

                        Ok((ChunkType::Li, Some(NumberingMeta::new(id, level, marker))))
                    }
                    Err(_) => {
                        return Err(DocError::new(&format!(
                            "abstract numbering with id {} not found",
                            num.abstract_num_id.to_string()
                        )))
                    }
                }
            }
            Err(_) => {
                return Err(DocError::new(&format!(
                    "numbering with id {} not found",
                    id.to_string()
                )))
            }
        }
    }

    fn id(&mut self) -> usize {
        self.id_counter += 1;
        self.id_counter
    }
}

#[cfg(test)]
mod tests {
    use docx_rs::{
        AbstractNumbering, AlignmentType, BreakType, Docx, Hyperlink, HyperlinkType, IndentLevel,
        Level, LevelJc, LevelText, LineSpacing, NumberFormat, Numbering, NumberingId, Paragraph,
        Run, RunFonts, Start, Style, StyleType,
    };

    use crate::{
        types::{Chunk, ChunkType, Properties, Px},
        DocxDocument,
    };

    struct T {
        id_counter: usize,
    }

    impl T {
        pub fn new() -> T {
            T { id_counter: 0 }
        }

        fn id(&mut self) -> usize {
            self.id_counter += 1;
            self.id_counter
        }

        fn text(&mut self, text: &str, props: Properties) -> Chunk {
            let mut ch = Chunk::new(self.id(), ChunkType::Text);
            ch.set_props(props);
            ch.set_text(text.to_owned());
            ch
        }
        fn hyperlink(&mut self, url: &str, props: Properties) -> Chunk {
            let mut ch = Chunk::new(self.id(), ChunkType::Link);
            ch.set_props(props);
            ch.set_url(url.to_owned());
            ch
        }
        fn para(&mut self, props: Properties) -> Chunk {
            let mut ch = Chunk::new(self.id(), ChunkType::Paragraph);
            ch.set_props(props);
            ch
        }
        fn end(&mut self) -> Chunk {
            Chunk::new(self.id(), ChunkType::End)
        }
        fn ol(&mut self, props: Properties) -> Chunk {
            let mut ch = Chunk::new(self.id(), ChunkType::Ol);
            ch.set_props(props);
            ch
        }
        fn ul(&mut self, props: Properties) -> Chunk {
            let mut ch = Chunk::new(self.id(), ChunkType::Ul);
            ch.set_props(props);
            ch
        }
        fn li(&mut self, props: Properties) -> Chunk {
            let mut ch = Chunk::new(self.id(), ChunkType::Li);
            ch.set_props(props);
            ch
        }
        fn br(&mut self) -> Chunk {
            Chunk::new(self.id(), ChunkType::Break)
        }
    }

    fn to_json_from_docx(docx: docx_rs::Docx) -> String {
        let mut d = DocxDocument::from_docx(docx);
        d.build().unwrap();
        serde_json::to_string(&d.chunks).unwrap()
    }
    fn px_to_docx_points(px: i32) -> i32 {
        let pt = px_to_pt(px);
        pt * 2
    }
    fn px_to_pt(px: i32) -> i32 {
        (px as f32 * 0.75).round() as i32
    }
    fn px_to_indent(px: i32) -> i32 {
        px * 15
    }
    use std::{fs::File, io::Read};

    #[test]
    fn test_read() {
        let path = "./example/all.docx".to_owned();
        let mut f = File::open(path).unwrap();
        let mut buf = Vec::<u8>::new();
        f.read_to_end(&mut buf).unwrap();

        let mut d = DocxDocument::new(buf);
        d.build().unwrap();
    }

    #[test]
    fn test_from_docx() {
        let mut t = T::new();
        let expected_chunks = vec![
            t.para(Properties {
                indent: Some(Px::new(60)),
                ..Default::default()
            }),
            t.text(
                "Hello World!",
                Properties {
                    bold: Some(true),
                    ..Default::default()
                },
            ),
            t.end(),
            t.para(Properties::default()),
            t.hyperlink("", Properties::default()),
            t.text("Click me!", Properties::default()),
            t.end(),
            t.end(),
        ];
        let expexted_json = serde_json::to_string(&expected_chunks).unwrap();

        let docx = Docx::new()
            .add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_text("Hello World!"))
                    .indent(Some(px_to_indent(60)), None, None, None)
                    .bold(),
            )
            .add_paragraph(
                Paragraph::new().add_hyperlink(
                    // empty url value because the hyperlink map is processed during build stage
                    Hyperlink::new("".to_owned(), HyperlinkType::External)
                        .add_run(Run::new().add_text("Click me!")),
                ),
            );
        let actual_json = to_json_from_docx(docx);

        assert_eq!(expexted_json, actual_json);
    }

    #[test]
    fn test_one_level_numbering() {
        let mut t = T::new();
        let expected_chunks = vec![
            t.ul(Properties::default()),
            /**/ t.li(Properties::default()),
            /*  */ t.text("Hello", Properties::default()),
            /**/ t.end(),
            /**/ t.li(Properties::default()),
            /*  */ t.text("World!", Properties::default()),
            /**/ t.end(),
            t.end(),
            t.para(Properties::default()),
            /**/ t.text("some text 1", Properties::default()),
            t.end(),
            t.para(Properties::default()),
            /**/ t.text("some text 2", Properties::default()),
            t.end(),
        ];
        let expected_json = serde_json::to_string(&expected_chunks).unwrap();

        let docx = Docx::new()
            .add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_text("Hello"))
                    .numbering(NumberingId::new(2), IndentLevel::new(0)),
            )
            .add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_text("World!"))
                    .numbering(NumberingId::new(2), IndentLevel::new(0)),
            )
            .add_paragraph(Paragraph::new().add_run(Run::new().add_text("some text 1")))
            .add_paragraph(Paragraph::new().add_run(Run::new().add_text("some text 2")))
            .add_abstract_numbering(AbstractNumbering::new(2).add_level(Level::new(
                0,
                Start::new(1),
                NumberFormat::new("bullet"),
                LevelText::new("Section %1."),
                LevelJc::new("left"),
            )))
            .add_numbering(Numbering::new(2, 2));
        let actual_json = to_json_from_docx(docx);

        assert_eq!(expected_json, actual_json);
    }

    #[test]
    fn test_three_levels_numbering() {
        let mut t = T::new();
        let expected_chunks = vec![
            t.ol(Properties::default()),
            /**/ t.li(Properties::default()),
            /*   */ t.text("level 1", Properties::default()),
            /**/ t.end(),
            /**/ t.ol(Properties::default()),
            /*   */ t.li(Properties::default()),
            /*      */ t.text("level 2", Properties::default()),
            /*   */ t.end(),
            /*   */ t.ol(Properties::default()),
            /*      */ t.li(Properties::default()),
            /*         */ t.text("level 3", Properties::default()),
            /*      */ t.end(),
            /*   */ t.end(),
            /**/ t.end(),
            /**/ t.li(Properties::default()),
            /*   */ t.text("level 1", Properties::default()),
            /**/ t.end(),
            t.end(),
            t.para(Properties::default()),
            /**/ t.text("Test", Properties::default()),
            t.end(),
        ];
        let expected_json = serde_json::to_string(&expected_chunks).unwrap();

        let docx = Docx::new()
            .add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_text("level 1"))
                    .numbering(NumberingId::new(2), IndentLevel::new(0)),
            )
            .add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_text("level 2"))
                    .numbering(NumberingId::new(2), IndentLevel::new(1)),
            )
            .add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_text("level 3"))
                    .numbering(NumberingId::new(2), IndentLevel::new(2)),
            )
            .add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_text("level 1"))
                    .numbering(NumberingId::new(2), IndentLevel::new(0)),
            )
            .add_paragraph(Paragraph::new().add_run(Run::new().add_text("Test")))
            .add_abstract_numbering(
                AbstractNumbering::new(2)
                    .add_level(Level::new(
                        0,
                        Start::new(1),
                        NumberFormat::new("decimal"),
                        LevelText::new("Section %1."),
                        LevelJc::new("left"),
                    ))
                    .add_level(Level::new(
                        1,
                        Start::new(1),
                        NumberFormat::new("decimal"),
                        LevelText::new("Section %2."),
                        LevelJc::new("left"),
                    ))
                    .add_level(Level::new(
                        2,
                        Start::new(1),
                        NumberFormat::new("decimal"),
                        LevelText::new("Section %3."),
                        LevelJc::new("left"),
                    )),
            )
            .add_numbering(Numbering::new(2, 2));
        let actual_json = to_json_from_docx(docx);

        assert_eq!(expected_json, actual_json);
    }

    #[test]
    fn test_inline_styles() {
        let mut t = T::new();
        let expected_chunks = vec![
            t.para(Properties {
                align: Some("center".to_owned()),
                ..Default::default()
            }),
            t.text(
                "Hello Word!",
                Properties {
                    bold: Some(true),
                    italic: Some(true),
                    ..Default::default()
                },
            ),
            t.text(
                "Hello Rust!",
                Properties {
                    bold: Some(true),
                    underline: Some(true),
                    ..Default::default()
                },
            ),
            t.end(),
        ];
        let expected_json = serde_json::to_string(&expected_chunks).unwrap();

        let docx = Docx::new().add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("Hello Word!").italic())
                .add_run(Run::new().add_text("Hello Rust!").underline("single"))
                .bold()
                .align(AlignmentType::Center),
        );
        let actual_json = to_json_from_docx(docx);

        assert_eq!(expected_json, actual_json);
    }

    #[test]
    fn test_styles() {
        let mut t = T::new();
        let expected_chunks = vec![
            t.para(Properties {
                align: Some("end".to_owned()),
                ..Default::default()
            }),
            t.text(
                "Hello Word!",
                Properties {
                    font_size: Some(Px::new(32)),
                    font_family: Some("Times".to_owned()),
                    color: Some("#888".to_owned()),
                    ..Default::default()
                },
            ),
            t.text(
                "Hello Rust!",
                Properties {
                    font_size: Some(Px::new(8)),
                    font_family: Some("Times".to_owned()),
                    color: Some("#888".to_owned()),
                    bold: Some(true),
                    ..Default::default()
                },
            ),
            t.end(),
        ];
        let expected_json = serde_json::to_string(&expected_chunks).unwrap();

        let st_base = Style::new("Test_base", StyleType::Paragraph)
            .align(AlignmentType::End)
            .color("#FFF")
            .size(px_to_docx_points(32) as usize);

        let st = Style::new("Test", StyleType::Paragraph)
            .based_on("Test_base")
            .color("#888")
            .fonts(RunFonts::new().ascii("Times"));

        let docx = Docx::new()
            .add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_text("Hello Word!"))
                    .add_run(
                        Run::new()
                            .add_text("Hello Rust!")
                            .bold()
                            .size(px_to_docx_points(8) as usize),
                    )
                    .style("Test"),
            )
            .add_style(st_base)
            .add_style(st);

        let actual_json = to_json_from_docx(docx);

        assert_eq!(expected_json, actual_json);
    }

    #[test]
    fn test_break() {
        let mut t = T::new();
        let expected_chunks = vec![
            t.para(Properties::default()),
            t.text("Hello", Properties::default()),
            t.br(),
            t.text("Rust!", Properties::default()),
            t.end(),
        ];
        let expected_json = serde_json::to_string(&expected_chunks).unwrap();

        let docx = Docx::new().add_paragraph(
            Paragraph::new().add_run(
                Run::new()
                    .add_text("Hello")
                    .add_break(BreakType::TextWrapping)
                    .add_text("Rust!"),
            ),
        );

        let actual_json = to_json_from_docx(docx);

        assert_eq!(expected_json, actual_json);
    }

    #[test]
    fn test_spacing() {
        let mut t = T::new();
        let expected_chunks = vec![
            t.para(Properties {
                line_height: Some("3".to_owned()),
                ..Default::default()
            }),
            t.text(
                "Hello Rust!",
                Properties {
                    font_size: Some(Px::new(10)),
                    ..Default::default()
                },
            ),
            t.end(),
        ];
        let expexted_json = serde_json::to_string(&expected_chunks).unwrap();

        let docx = Docx::new().add_paragraph(
            Paragraph::new()
                .add_run(
                    Run::new()
                        .add_text("Hello Rust!")
                        .size(px_to_docx_points(10) as usize),
                )
                .line_spacing(
                    LineSpacing::new()
                        .after(px_to_indent(10) as u32)
                        .before(px_to_indent(10) as u32),
                ),
        );
        let actual_json = to_json_from_docx(docx);

        assert_eq!(expexted_json, actual_json);
    }
}
