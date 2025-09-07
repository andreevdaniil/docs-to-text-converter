// index.js
import express from "express";
import mammoth from "mammoth";
import iconv from "iconv-lite";

const app = express();
const PORT = process.env.PORT || 3000;

// ===== RTF helpers =====

// Определяем RTF по началу "{\rtf"
function isRtfBuffer(buf) {
  const head = buf.slice(0, 5).toString("ascii");
  return head.startsWith("{\\rtf");
}

// Удаляем вложенные группы полностью: fonttbl/colortbl/stylesheet/info/pict/header/footer и любые {\* ...}
function stripGroups(rtf, ignoreKeys) {
  const N = rtf.length;
  let out = "";
  const stack = []; // [{ignore:boolean}]
  let i = 0;

  const isLetter = (c) =>
    (c >= "a" && c <= "z") || (c >= "A" && c <= "Z");

  const groupIgnored = () => stack.some((s) => s.ignore);

  while (i < N) {
    const ch = rtf[i];

    if (ch === "{") {
      // выясняем, нужно ли игнорить группу
      let j = i + 1;
      let ignore = false;

      if (rtf[j] === "\\") j++;

      if (rtf[j] === "*") {
        ignore = true; // {\* ...}
        j++;
      } else {
        if (rtf[j] === "\\") j++;
        let k = j;
        while (k < N && isLetter(rtf[k])) k++;
        const cw = rtf.slice(j, k).toLowerCase();
        if (ignoreKeys.has(cw)) ignore = true;
      }

      stack.push({ ignore });
      if (!groupIgnored()) out += "{";
      i++;
      continue;
    }

    if (ch === "}") {
      const g = stack.pop() || { ignore: false };
      if (!groupIgnored() && !g.ignore) out += "}";
      i++;
      continue;
    }

    if (!groupIgnored()) out += ch;
    i++;
  }
  return out;
}

// Конвертация RTF -> текст
function rtfToText(buf) {
  // binary-строка, чтобы сохранить \'hh
  let rtf = buf.toString("binary");

  // кодовая страница
  const cpMatch = rtf.match(/\\ansicpg(\d+)/);
  const codepage = cpMatch ? `cp${cpMatch[1]}` : "cp1251";

  // сносим вложенные служебные группы
  rtf = stripGroups(
    rtf,
    new Set([
      "fonttbl",
      "colortbl",
      "stylesheet",
      "info",
      "pict",
      "header",
      "footer",
    ])
  );
  // и любые {\* ...}
  rtf = stripGroups(rtf, new Set(["*"]));

  // абзацы/табы
  rtf = rtf.replace(/\\par[d]?\b/g, "\n").replace(/\\tab\b/g, "\t");

  // hex-эскейпы: \'hh
  rtf = rtf.replace(/\\'([0-9a-fA-F]{2})/g, (_, h) =>
    iconv.decode(Buffer.from(h, "hex"), codepage)
  );

  // вычищаем управляющие слова \b \fs24 \f0 \u-123? и т.п.
  rtf = rtf.replace(
    /\\[a-z]+-?\d*(?:\s|(?=[\\{}]))/gi,
    ""
  );

  // без скобок и лишних переносов
  rtf = rtf
    .replace(/[{}]/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  return rtf;
}

// ===== HTTP =====

// Принимаем «сырое» тело любого типа, будем определять формат сами
app.post(
  "/convert",
  express.raw({ type: "*/*", limit: "20mb" }),
  async (req, res) => {
    try {
      const buf = Buffer.isBuffer(req.body)
        ? req.body
        : Buffer.from(req.body);

      // 1) RTF по сигнатуре
      if (isRtfBuffer(buf)) {
        const text = rtfToText(buf);
        return res.json({ type: "rtf", text });
      }

      // 2) DOCX через mammoth
      const { value } = await mammoth.extractRawText({ buffer: buf });
      return res.json({ type: "docx", text: value });
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: String(err) });
    }
  }
);

// healthcheck
app.get("/health", (_, res) => res.send("ok"));

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});