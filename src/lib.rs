mod error;
mod types;

use wasm_bindgen;
use wasm_bindgen::prelude::*;

use base64::{engine::general_purpose, Engine};
use docx_rs::{
    DocumentChild, DrawingData, HyperlinkData, Paragraph, ParagraphChild, ParagraphProperty,
    RunChild, RunProperty,
};

use error::DocError;
use serde::Serialize;
use types::{Chunk, ChunkType, Properties};

#[derive(Clone, PartialEq)]
enum NumberingType {
    Unordered,
    Ordered,
}

#[derive(Clone)]
struct NumberingMeta {
    pub id: usize,
    pub level: usize,
    pub marker: NumberingType,
}

impl NumberingMeta {
    fn new(id: usize, level: usize, marker: NumberingType) -> NumberingMeta {
        NumberingMeta {
            id: id,
            level: level,
            marker: marker,
        }
    }

    fn to_chunk_type(&self) -> ChunkType {
        match self.marker {
            NumberingType::Ordered => ChunkType::Ol,
            NumberingType::Unordered => ChunkType::Ul,
        }
    }
}

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
        let docx = docx_rs::read_docx(&bytes).unwrap();
        let mut doc = DocxDocument::from_docx(docx);
        doc.build().unwrap();

        doc
    }

    pub fn get_chunks(self) -> JsValue {
        JsValue::from_serde(&self.chunks).unwrap()
    }

    fn from_docx(docx: docx_rs::Docx) -> DocxDocument {
        DocxDocument {
            chunks: vec![],
            docx: docx,
            id_counter: 0,
        }
    }

    fn id(&mut self) -> usize {
        self.id_counter += 1;
        self.id_counter
    }

    fn build(&mut self) -> Result<(), DocError> {
        let mut i = 0;

        while i < self.docx.document.children.len() {
            match self.docx.document.children[i].to_owned() {
                DocumentChild::Paragraph(p) => {
                    let (chunk_type, numbering_meta) = self.define_chunk_type(&p)?;

                    match chunk_type {
                        ChunkType::Paragraph => self.parse_para(&p, ChunkType::Paragraph)?,
                        ChunkType::Li => {
                            let meta = &numbering_meta.unwrap();
                            let i_copy = i.to_owned();

                            self.parse_numbering(0, &mut i, &meta, i_copy)?
                        }
                        _ => return Err(DocError::new("unknown chunk type")),
                    };
                }
                _ => (),
            }
            i += 1;
        }

        Ok(())
    }

    fn append_chunk(&mut self, ch: Chunk) {
        self.chunks.push(ch);
    }

    fn parse_block_content<'a, T>(
        &mut self,
        base_props: &Properties,
        children_iter: T,
    ) -> Result<(), DocError>
    where
        T: Iterator<Item = &'a ParagraphChild>,
    {
        for child in children_iter {
            match child {
                ParagraphChild::Hyperlink(v) => match &v.link {
                    HyperlinkData::External { rid, path: _ } => {
                        let url = self.find_hyperlink_url(rid);

                        let mut hyperlink = Chunk::new(self.id(), ChunkType::Link);
                        hyperlink.set_url(url.unwrap_or(String::default()));

                        self.append_chunk(hyperlink);

                        self.parse_block_content(base_props, v.children.iter())?;

                        let end = Chunk::new(self.id(), ChunkType::End);
                        self.append_chunk(end);
                    }
                    _ => (),
                },
                ParagraphChild::Run(run) => {
                    let run_props = self.override_run_properties(base_props, &run.run_property);

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

                                            let mut image = Chunk::new(self.id(), ChunkType::Image);
                                            image.set_url(url);

                                            self.append_chunk(image);
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

                                self.append_chunk(text);
                            }
                            _ => (),
                        }
                    }
                }
                _ => (),
            }
        }

        Ok(())
    }

    fn parse_para(&mut self, p: &Paragraph, t: ChunkType) -> Result<(), DocError> {
        // get properties from the style
        let mut style_props = Properties::new();
        if let Some(v) = &p.property.style {
            self.load_props_from_style(&v.val, &mut style_props);
        }

        // get properties from inline run style
        let base_run_props = self.override_run_properties(&style_props, &p.property.run_property);
        // get properties from inline para style
        let para_props = self.override_para_properties(&style_props, &p.property);

        // parse paragraph content

        let mut paragraph = Chunk::new(self.id(), t);
        paragraph.set_props(para_props);

        self.append_chunk(paragraph);

        self.parse_block_content(&base_run_props, p.children.iter())?;

        let end = Chunk::new(self.id(), ChunkType::End);
        self.append_chunk(end);

        Ok(())
    }

    fn parse_numbering(
        &mut self,
        l: usize,
        from: &mut usize,
        parent_meta: &NumberingMeta,
        from_cp: usize,
    ) -> Result<(), DocError> {
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
                    // child is not list
                    *from -= 1;
                    let end = Chunk::new(self.id(), ChunkType::End);
                    self.append_chunk(end);

                    return Ok(());
                }
                num_meta = m.unwrap();
                chunk_type = t;
            } else {
                let list = Chunk::new(self.id(), parent_meta.to_chunk_type());
                self.append_chunk(list);
            }

            if num_meta.id == parent_meta.id
                && num_meta.level == parent_meta.level
                && num_meta.marker == parent_meta.marker
            {
                //
                // the same list
                //
                self.parse_para(&p, chunk_type)?;

                *from += 1;
            } else if num_meta.level > parent_meta.level {
                //
                // nested list
                //

                self.parse_numbering(l + 1, from, &num_meta, from.clone())?;
            } else if num_meta.level == parent_meta.level {
                //
                // sibling list
                //
                if l == 0 {
                    // if nesting level is zero, then return root list
                    let end = Chunk::new(self.id(), ChunkType::End);
                    self.append_chunk(end);
                    return Ok(());
                }
                let list_type = num_meta.to_owned().to_chunk_type();
                let sibling_root = Chunk::new(self.id(), list_type);

                self.append_chunk(sibling_root);

                self.parse_numbering(l, from, &num_meta, from.clone())?;
                let end = Chunk::new(self.id(), ChunkType::End);
                self.append_chunk(end);
            } else {
                //
                // list of the level above
                //
                let end = Chunk::new(self.id(), ChunkType::End);
                self.append_chunk(end);

                return Ok(());
            }
        }

        return Ok(());
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
        dest.line_height = para_props.line_height;

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
        if let Some(_) = &props.indent {
            // TODO parse indent
        }
        if let Some(_) = &props.line_spacing {
            // TODO parse line height
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
            new_props.underline = Some(!underline.val.is_empty());
        }
        if let Some(sz) = &props.sz {
            // FIXME parse twips/pt to px
            let size = sz.val;
            new_props.font_size = Some(format!("{}px", size.to_string()));
        }
        if let Some(color) = &props.color {
            new_props.color = Some(color.val.to_owned());
        }
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
}

#[cfg(test)]
mod tests {
    use docx_rs::{
        AbstractNumbering, AlignmentType, Docx, Hyperlink, HyperlinkType, IndentLevel, Level,
        LevelJc, LevelText, NumberFormat, Numbering, NumberingId, Paragraph, Run, RunFonts, Start,
        Style, StyleType,
    };

    use crate::{
        types::{Chunk, ChunkType, Properties},
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
            ch.props = props;
            ch.set_text(text.to_owned());
            ch
        }
        fn hyperlink(&mut self, url: &str, props: Properties) -> Chunk {
            let mut ch = Chunk::new(self.id(), ChunkType::Link);
            ch.props = props;
            ch.set_url(url.to_owned());
            ch
        }
        fn para(&mut self, props: Properties) -> Chunk {
            let mut ch = Chunk::new(self.id(), ChunkType::Paragraph);
            ch.props = props;
            ch
        }
        fn end(&mut self) -> Chunk {
            Chunk::new(self.id(), ChunkType::End)
        }
        fn ol(&mut self, props: Properties) -> Chunk {
            let mut ch = Chunk::new(self.id(), ChunkType::Ol);
            ch.props = props;
            ch
        }
        fn ul(&mut self, props: Properties) -> Chunk {
            let mut ch = Chunk::new(self.id(), ChunkType::Ul);
            ch.props = props;
            ch
        }
        fn li(&mut self, props: Properties) -> Chunk {
            let mut ch = Chunk::new(self.id(), ChunkType::Li);
            ch.props = props;
            ch
        }
    }

    fn to_json_from_docx(docx: docx_rs::Docx) -> String {
        let mut d = DocxDocument::from_docx(docx);
        d.build().unwrap();
        serde_json::to_string(&d.chunks).unwrap()
    }

    #[test]
    fn test_from_docx() {
        let mut t = T::new();
        let expected_chunks = vec![
            t.para(Properties::default()),
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
                    font_size: Some("32px".to_owned()),
                    font_family: Some("Times".to_owned()),
                    color: Some("#888".to_owned()),
                    ..Default::default()
                },
            ),
            t.text(
                "Hello Rust!",
                Properties {
                    font_size: Some("8px".to_owned()),
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
            .size(32);

        let st = Style::new("Test", StyleType::Paragraph)
            .based_on("Test_base")
            .color("#888")
            .fonts(RunFonts::new().ascii("Times"));

        let docx = Docx::new()
            .add_paragraph(
                Paragraph::new()
                    .add_run(Run::new().add_text("Hello Word!"))
                    .add_run(Run::new().add_text("Hello Rust!").bold().size(8))
                    .style("Test"),
            )
            .add_style(st_base)
            .add_style(st);

        let actual_json = to_json_from_docx(docx);

        assert_eq!(expected_json, actual_json);
    }
}
