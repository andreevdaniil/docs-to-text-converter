import express from "express";
import mammoth from "mammoth";

const app = express();
const PORT = process.env.PORT || 3000;

app.post(
  "/convert",
  express.raw({
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    limit: "20mb",
  }),
  async (req, res) => {
    try {
      const { value } = await mammoth.extractRawText({ buffer: req.body });
      res.json({ text: value });
    } catch (err) {
      res.status(500).json({ error: String(err) });
    }
  }
);

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});