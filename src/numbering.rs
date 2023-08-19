use crate::types::ChunkType;

#[derive(Clone, PartialEq)]
pub enum NumberingType {
    Unordered,
    Ordered,
}

pub enum ListRelation {
    ListItem,
    SiblingList,
    NestedList,
    End,
}

#[derive(Clone)]
pub struct NumberingMeta {
    pub id: usize,
    pub level: usize,
    pub marker: NumberingType,
}

impl NumberingMeta {
    pub fn new(id: usize, level: usize, marker: NumberingType) -> NumberingMeta {
        NumberingMeta {
            id: id,
            level: level,
            marker: marker,
        }
    }

    pub fn to_chunk_type(&self) -> ChunkType {
        match self.marker {
            NumberingType::Ordered => ChunkType::Ol,
            NumberingType::Unordered => ChunkType::Ul,
        }
    }
}

pub fn define_list_relation(cur: &NumberingMeta, next: &NumberingMeta) -> ListRelation {
    if cur.id == next.id && cur.level == next.level && cur.marker == next.marker {
        return ListRelation::ListItem;
    } else if cur.level > next.level {
        return ListRelation::NestedList;
    } else if cur.level == next.level {
        return ListRelation::SiblingList;
    } else {
        return ListRelation::End;
    }
}
