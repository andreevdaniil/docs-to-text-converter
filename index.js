import express from "express";
import mammoth from "mammoth";

const app = express();
const PORT = process.env.PORT || 3000;

const isDocxOrOctet = (req) =>
  /application\/(vnd\.openxmlformats-officedocument\.wordprocessingml\.document|octet-stream)/i
    .test(req.headers["content-type"] || "");

app.post(
  "/convert",
  express.raw({ type: isDocxOrOctet, limit: "20mb" }),
  async (req, res) => {
    try {
      const { value } = await mammoth.extractRawText({ buffer: req.body });
      res.json({ text: value });
    } catch (err) {
      res.status(500).json({ error: String(err) });
    }
  }
);

// healthcheck, чтобы платформа видела, что жив
app.get("/health", (_, res) => res.send("ok"));

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});