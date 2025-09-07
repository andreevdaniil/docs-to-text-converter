import express from "express";
import mammoth from "mammoth";      // для DOCX
import iconv from "iconv-lite";     // для корректной кириллицы в RTF

const app = express();
const PORT = process.env.PORT || 3000;

// Принимаем «сырое» тело любого типа (распознаем сами)
app.post(
  "/convert",
  express.raw({ type: () => true, limit: "20mb" }),
  async (req, res) => {
    try {
      const buf = Buffer.isBuffer(req.body) ? req.body : Buffer.from(req.body || []);

      if (isDocx(buf)) {
        // DOCX -> text
        const { value } = await mammoth.extractRawText({ buffer: buf });
        return res.json({ type: "docx", text: value });
      }

      if (isRtf(buf)) {
        // RTF -> text
        const text = rtfToText(buf);
        return res.json({ type: "rtf", text });
      }

      return res.status(415).json({ error: "Unsupported file type" });
    } catch (err) {
      res.status(500).json({ error: String(err) });
    }
  }
);

// healthcheck
app.get("/health", (_, res) => res.send("ok"));

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});

/* ---------- helpers ---------- */

// DOCX = zip: 'PK\x03\x04' в начале
function isDocx(buf) {
  return buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4B && buf[2] === 0x03 && buf[3] === 0x04;
}

// RTF начинается с "{\rtf"
function isRtf(buf) {
  const head = buf.slice(0, 5).toString("ascii");
  return head.startsWith("{\\rtf");
}

// Преобразование RTF -> текст с учётом \ansicpgXXXX и \'hh
function rtfToText(rtfBuf) {
  // читаем как binary, чтобы не потерять байты для \'hh
  let rtf = rtfBuf.toString("binary");

  // кодовая страница (часто cp1251)
  const m = rtf.match(/\\ansicpg(\d+)/);
  const codepage = m ? `cp${m[1]}` : "cp1251";

  // убрать служебные группы
  rtf = rtf
    .replace(/\{\\\*[^{}]*\}/g, "")
    .replace(/\{\\fonttbl[^{}]*\}/g, "")
    .replace(/\{\\colortbl[^{}]*\}/g, "")
    .replace(/\{\\stylesheet[^{}]*\}/g, "");

  // базовые управлялки
  rtf = rtf.replace(/\\par[d]?\b/g, "\n").replace(/\\tab\b/g, "\t");

  // \'hh -> символ нужной кодировки
  rtf = rtf.replace(/\\'([0-9a-fA-F]{2})/g, (_, h) =>
    iconv.decode(Buffer.from(h, "hex"), codepage)
  );

  // убрать прочие управляющие слова (\wordN / \word)
  rtf = rtf.replace(/\\[a-z]+-?\d*(?:\s|(?=[\\{}]))/gi, "");

  // убрать группирующие скобки и лишние переводы
  rtf = rtf
    .replace(/[{}]/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  return rtf;
}