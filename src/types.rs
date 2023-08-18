use serde::{Serialize, Serializer};

#[derive(PartialEq, Clone, Copy)]
pub enum ChunkType {
    Paragraph = 2 | 0x4000 | 0x2000,
    Text = 3 | 0x8000,
    Image = 5 | 0x8000,
    Link = 6 | 0x2000 | 0x8000,
    Ul = 8 | 0x2000 | 0x4000,
    Ol = 9 | 0x2000 | 0x4000,
    Li = 10 | 0x2000 | 0x4000,
    End = 0x1fff,
}

#[derive(Serialize, Clone)]
pub struct Chunk {
    id: usize,
    chunk_type: ChunkType,
    #[serde(skip_serializing_if = "Option::is_none")]
    text: Option<String>,
    #[serde(skip_serializing_if = "Properties::is_empty")]
    pub props: Properties,
}

#[derive(Serialize, Debug, Clone, Default)]
pub struct Properties {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub url: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub background: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_height: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<String>,
}

impl Chunk {
    pub fn new(id: usize, t: ChunkType) -> Chunk {
        Chunk {
            id: id,
            chunk_type: t,
            props: Properties::new(),
            text: None,
        }
    }

    pub fn set_text(&mut self, text: String) {
        if self.chunk_type == ChunkType::Text {
            self.text = Some(text);
        }
    }

    pub fn set_url(&mut self, url: String) {
        if self.chunk_type == ChunkType::Image || self.chunk_type == ChunkType::Link {
            self.props.url = Some(url);
        }
    }

    pub fn set_size(&mut self, w: usize, h: usize) {
        if self.chunk_type == ChunkType::Image {
            self.props.width = Some(format!("{}px", w.to_string()));
            self.props.height = Some(format!("{}px", h.to_string()));
        }
    }

    pub fn set_props(&mut self, props: Properties) {
        match self.chunk_type {
            ChunkType::Paragraph
            | ChunkType::Li
            | ChunkType::Ol
            | ChunkType::Ul
            | ChunkType::Link => {
                self.props.indent = props.indent;
                self.props.align = props.align;
                self.props.line_height = props.line_height;
            }
            ChunkType::Text => {
                self.props.color = props.color;
                self.props.background = props.background;
                self.props.font_size = props.font_size;
                self.props.font_family = props.font_family;
                self.props.bold = props.bold;
                self.props.italic = props.italic;
                self.props.underline = props.underline;
            }
            _ => (),
        }
    }
}

impl Properties {
    pub fn new() -> Properties {
        Default::default()
    }

    pub fn is_empty(&self) -> bool {
        self.url == None
            && self.color == None
            && self.background == None
            && self.font_size == None
            && self.font_family == None
            && self.bold == None
            && self.italic == None
            && self.underline == None
            && self.align == None
            && self.indent == None
            && self.line_height == None
    }
}

impl Serialize for ChunkType {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        match self {
            ChunkType::Paragraph => serializer.serialize_u32(2 | 0x4000 | 0x2000),
            ChunkType::Text => serializer.serialize_u32(3 | 0x8000),
            ChunkType::Image => serializer.serialize_i32(5 | 0x8000),
            ChunkType::Link => serializer.serialize_i32(6 | 0x2000 | 0x8000),
            ChunkType::Ul => serializer.serialize_i32(8 | 0x2000 | 0x4000),
            ChunkType::Ol => serializer.serialize_i32(9 | 0x2000 | 0x4000),
            ChunkType::Li => serializer.serialize_i32(10 | 0x2000 | 0x4000),
            ChunkType::End => serializer.serialize_i32(0x1fff),
        }
    }
}
