import { DocxDocument } from "../pkg/docx2json.js";

export function convertArray(data) {
  const docx = DocxDocument.new(data);
  const chunks = docx.get_chunks();
  return chunks;
}

export function convert(data) {
  if (data instanceof File) {
    return new Promise((res) => {
      const reader = new FileReader();
      reader.readAsArrayBuffer(data);
      reader.onload = (e) => {
        res(convertArray(new Int8Array(e.target.result)));
      };
    });
  }
}
