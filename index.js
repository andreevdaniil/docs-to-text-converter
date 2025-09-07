import express from "express";
import mammoth from "mammoth";
import iconv from "iconv-lite";

const app = express();
const PORT = process.env.PORT || 3000;

// ---- helpers: sniff type ----
function looksLikeRtf(buf) {
  return buf?.slice(0, 5).toString("ascii") === "{\\rtf";
}
function looksLikeDocx(buf) {
  // zip magic "PK" — DOCX это zip-контейнер
  return buf?.slice(0, 2).toString("ascii") === "PK";
}

// ---- RTF -> plain text ----
function rtfToText(rtfBuf) {
  let rtf = rtfBuf.toString("latin1"); // не ломаем байты
  // 1) вырезаем картинки, вложенные объекты и скрытые группы
  rtf = rtf
    .replace(/\{\\\*\\[^{}]+?\{[^{}]*\}[^{}]*\}/gs, "") // спец. группы {\*\...{...}...}
    .replace(/\{\\pict[\s\S]*?\}/gi, "");              // картинки {\pict ...}

  // 2) вытащим кодовую страницу, по умолчанию 1252
  const mCpg = rtf.match(/\\ansicpg(\d{3,4})/i);
  const cp = mCpg ? `cp${mCpg[1]}` : "cp1252";

  // 3) сначала unicode: \uNNNN?  (RTF: signed 16-bit)
  rtf = rtf.replace(/\\u(-?\d+)\??/g, (_, u) => {
    let code = parseInt(u, 10);
    if (code < 0) code = 65536 + code; // signed -> unsigned
    return String.fromCharCode(code);
  });

  // 4) потом байтовые последовательности \'hh -> собрать в Buffer и декодировать в нужной кодировке
  rtf = rtf.replace(/\\'([0-9a-fA-F]{2})/g, (_, hh) =>
    iconv.decode(Buffer.from(hh, "hex"), cp)
  );

  // 5) убрать управляющие слова/экранирования \\, \{, \}
  rtf = rtf
    .replace(/\\[a-z]+-?\d* ?/gi, "") // команды \b, \par, \fs20 и т.п.
    .replace(/\\[\{\}\\]/g, m => m.slice(1)) // \{ \} \\ -> { } \
    .replace(/[{}]/g, ""); // скобочные группы

  // 6) нормализация пробелов/переводов строк
  return rtf
    .replace(/\r\n?/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]{2,}/g, " ")
    .trim();
}

// ---- middleware: принимем сырой бинарь (octet-stream, docx, rtf) ----
app.post(
  "/convert",
  express.raw({
    type: [
      "application/octet-stream",
      "application/rtf",
      "text/rtf",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ],
    limit: "20mb",
  }),
  async (req, res) => {
    try {
      const buf = Buffer.from(req.body);

      let text = "";
      if (looksLikeDocx(buf)) {
        const { value } = await mammoth.extractRawText({ buffer: buf });
        text = value || "";
      } else if (looksLikeRtf(buf)) {
        text = rtfToText(buf);
      } else {
        // fallback: попробуем как RTF, часто приходят с generic content-type
        text = rtfToText(buf);
      }

      // лёгкая пост-очистка
      text = text.replace(/\u0000/g, "").trim();
      res.json({ type: looksLikeDocx(buf) ? "docx" : "rtf", text });
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