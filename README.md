## Docx2json

Docx2json is a Rust and WebAssembly-based library that allows easily converting files in Docx format to JSON files.

### How to start

#### 1. Install wasm pack

```
cargo install wasm-pack
```

#### 2. Run App

```
yarn 
yarn build
yarn dev
```

### Output format

The document converts to the array of chunks. Each chunk represents a piece of data

```ts
type Chunk = {
	id: number;
	type: ChunkType;
	text?: string;
	props?: TextProperties | ImageProperties | LinkProperties;
};

// LINK
type LinkProperties {
	url: string;
}

// IMAGE
type ImageProperties {
	url: string;
}

// TEXT
type TextProperties {
	color?: string,
	background?: string,
	fontSize?: string,
	fontFamily?: string,
	bold?: boolean,
	italic?: boolean,
	underline?: boolean,
}

// PARAGRAPH, LI, OL
type BlockProperties {
	align?: string,
	indent?: string,
	lineHeight?: string
}

const PARAGRAPH: ChunkType = 2 | 0x4000 | 0x2000;
const TEXT: ChunkType = 3 | 0x8000;
const IMAGE: ChunkType = 5 | 0x8000;
const LINK: ChunkType = 6 | 0x2000 | 0x8000;
const NEWLINE: ChunkType = 7 | 0x8000;
const UL: ChunkType = 8 | 0x2000 | 0x4000;
const OL: ChunkType = 9 | 0x2000 | 0x4000;
const LI: ChunkType = 10 | 0x2000 | 0x4000;
const END: ChunkType = 0x1fff;
```
