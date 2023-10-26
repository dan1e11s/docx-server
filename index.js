const express = require('express');
const cors = require('cors');
const multer = require('multer');
const mammoth = require('mammoth');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3003;

const storage = multer.memoryStorage();
const upload = multer({ storage });

const replacements = []; // Массив для хранения замен

app.use(
  cors({
    origin: '*',
    methods: 'GET,HEAD,PUT,PATCH,POST,DELETE',
    credentials: true,
  })
);

app.use(express.json());

// Замена слов в документе
app.post('/add-replacement', (req, res) => {
  const keyword = req.body.keyword;
  const replacement = req.body.replacement;
  replacements.push({ keyword, replacement });
  res.status(200).json({ message: 'Replacement added.' });
});

// Применение всех замен и отправка документа клиенту
app.post('/apply-replacements', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded.' });
  }

  const fileBuffer = req.file.buffer;

  mammoth
    .extractRawText({ buffer: fileBuffer })
    .then((result) => {
      const text = result.value;
      let replacedText = text;

      // Применяем все замены
      replacements.forEach((replace) => {
        const { keyword, replacement } = replace;
        replacedText = replacedText.replace(
          new RegExp(keyword, 'g'),
          replacement
        );
      });

      const encodedFileName = encodeURIComponent('новый-документ.docx');
      const fullFileName = 'новый-документ.docx';

      // Отправляем измененный документ как бинарные данные
      const buffer = Buffer.from(replacedText, 'utf8');
      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      );
      res.setHeader(
        'Content-Disposition',
        `attachment; filename="${encodedFileName}"`
      );
      res.setHeader('Content-Length', buffer.length); // Устанавливаем длину контента
      res.send(buffer);

      replacements.length = 0; // Очищаем массив замен
    })
    .catch((error) => {
      console.error(error);
      res.status(500).json({ error: 'Failed to process the document.' });
    });
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
