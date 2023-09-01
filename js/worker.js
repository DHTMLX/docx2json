import { DocxDocument } from "../pkg/docx2json";

onmessage = function (e) {
  const config = e.data;

  if (e.data.type === "convert") {
    const data = e.data.data;
    if (data instanceof File) {
      const reader = new FileReader();
      reader.readAsArrayBuffer(data);
      reader.onload = (e) => doConvert(new Int8Array(e.target.result));
    } else {
      doConvert(data);
    }
  }
};

function doConvert(input) {
  const docx = DocxDocument.new(input);
  const chunks = docx.get_chunks();

  postMessage({
    uid: new Date().valueOf(),
    type: "ready",
    data: chunks,
    styles,
  });
}

postMessage({ type: "init" });
