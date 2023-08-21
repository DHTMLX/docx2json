use serde::{Serialize, Serializer};

use crate::utils;

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
    Break = 11 | 0x4000,
}

#[derive(Serialize, Clone)]
pub struct Chunk {
    id: usize,
    chunk_type: ChunkType,
    #[serde(skip_serializing_if = "Option::is_none")]
    text: Option<String>,
    #[serde(skip_serializing_if = "Properties::is_empty")]
    props: Properties,
}

#[derive(Debug, Clone)]
pub struct Px {
    val: i32,
}

#[derive(Debug, Clone)]
pub struct Spacing {
    after: usize,
    before: usize,
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
    pub font_size: Option<Px>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub underline: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub strike: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indent: Option<Px>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub line_height: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub width: Option<Px>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub height: Option<Px>,

    #[serde(skip_serializing)]
    pub spacing: Option<Spacing>,
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

    pub fn set_size(&mut self, w: i32, h: i32) {
        if self.chunk_type == ChunkType::Image {
            self.props.width = Some(Px::new(w));
            self.props.height = Some(Px::new(h));
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
                self.props.strike = props.strike;
            }
            ChunkType::Image => {
                self.props.width = props.width;
                self.props.height = props.height;
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
            && self.font_size.is_none()
            && self.font_family == None
            && self.bold == None
            && self.italic == None
            && self.underline == None
            && self.align == None
            && self.indent.is_none()
            && self.line_height == None
            && self.width.is_none()
            && self.height.is_none()
            && self.strike == None
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
            ChunkType::Break => serializer.serialize_i32(11 | 0x4000),
        }
    }
}

impl Px {
    pub fn new(val: i32) -> Px {
        Px { val: val }
    }
    pub fn get_str(&self) -> String {
        format!("{}px", self.val)
    }
    pub fn get_val(&self) -> i32 {
        self.val
    }
}

impl Serialize for Px {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        serializer.serialize_str(&self.get_str())
    }
}

impl Spacing {
    pub fn new(after_px: usize, before_px: usize) -> Spacing {
        Spacing {
            after: after_px,
            before: before_px,
        }
    }
    pub fn set_after(&mut self, px: usize) {
        self.after = px;
    }
    pub fn set_before(&mut self, px: usize) {
        self.before = px;
    }
    pub fn calc_line_spacing(&self, sz_px: usize) -> Option<String> {
        if self.after == 0 && self.before == 0 {
            return None;
        }

        let sz = if sz_px == 0 {
            utils::DEFAULT_SZ_PX
        } else {
            sz_px
        };

        let line_height = self.after + self.before + sz;

        Some((line_height as f32 / sz as f32).to_string())
    }
}
