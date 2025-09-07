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
function isRtf(buf) {
  const head = buf.slice(0, 5).toString("ascii");
  return head.startsWith("{\\rtf");
}

// Walk the RTF and drop whole *nested* groups whose first control word
// matches one of the ignoreKeys (e.g. "fonttbl", "colortbl", "stylesheet", "*", "pict", "info", "header", "footer")
function stripGroups(rtf, ignoreKeys) {
  const N = rtf.length;
  let out = "";
  const stack = []; // [{ignore:boolean}]
  let i = 0;

  const isLetter = c => (c >= "a" && c <= "z") || (c >= "A" && c <= "Z");

  while (i < N) {
    const ch = rtf[i];

    if (ch === "{") {
      // new group: decide if it's ignorable
      // lookahead control word right after '{' and optional backslash
      let j = i + 1;
      let ignore = false;

      if (rtf[j] === "\\") j++;

      // special destination group: {\* ...}
      if (rtf[j] === "*") {
        ignore = true;
        j++;
      } else {
        // read first control word
        if (rtf[j] === "\\") j++;
        let k = j;
        while (k < N && isLetter(rtf[k])) k++;
        const cw = rtf.slice(j, k).toLowerCase();
        if (ignoreKeys.has(cw)) ignore = true;
      }

      stack.push({ ignore });
      if (!stack.some(s => s.ignore)) out += "{";
      i++;
      continue;
    }

    if (ch === "}") {
      const g = stack.pop() || { ignore: false };
      if (!stack.some(s => s.ignore) && !g.ignore) out += "}";
      i++;
      continue;
    }

    // inside pict group we skip everything until its closing '}' (handled by stack)
    // skipping happens automatically because stack.some(ignore) is true

    // normal char
    if (!stack.some(s => s.ignore)) out += ch;
    i++;
  }
  return out;
}

// Basic RTF -> text
function rtfToText(buf) {
  // work on a binary string to preserve \'hh
  let rtf = buf.toString("binary");

  // detect code page
  const cpMatch = rtf.match(/\\ansicpg(\d+)/);
  const codepage = cpMatch ? `cp${cpMatch[1]}` : "cp1251";

  // 1) drop nested ignorable groups completely
  rtf = stripGroups(
    rtf,
    new Set(["fonttbl", "colortbl", "stylesheet", "info", "pict", "header", "footer"])
  );
  // also drop any {\* …} groups
  rtf = stripGroups(rtf, new Set(["*"]));

  // 2) paragraph / tabs
  rtf = rtf.replace(/\\par[d]?\b/g, "\n").replace(/\\tab\b/g, "\t");

  // 3) hex escapes \'hh -> decode with codepage
  rtf = rtf.replace(/\\'([0-9a-fA-F]{2})/g, (_, h) =>
    iconv.decode(Buffer.from(h, "hex"), codepage)
  );

  // 4) remove remaining control words like \b, \fs24, \f0, \u-123?
  rtf = rtf.replace(/\\[a-z]+-?\d*(?:\s|(?=[\\{}]))/gi, "");

  // 5) strip braces and collapse whitespace
  rtf = rtf
    .replace(/[{}]/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  return rtf;
}

export { isRtf, rtfToText };