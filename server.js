const express = require('express');
const multer = require('multer');
const mammoth = require('mammoth');
const path = require('path');
const fs = require('fs');
const AdmZip = require('adm-zip');

const app = express();
const PORT = process.env.PORT || 3000;

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = path.join(require('os').tmpdir(), 'gift-uploads');
    if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const uniqueName = `${Date.now()}-${file.originalname}`;
    cb(null, uniqueName);
  }
});

const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext !== '.docx') {
      return cb(new Error('Only .docx files are accepted'), false);
    }
    cb(null, true);
  },
  limits: { fileSize: 10 * 1024 * 1024 } // 10MB per file
});

// Serve static frontend
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// ─── TEXT NORMALIZATION ───────────────────────────────────────────────
function normalizeText(text) {
  return text
    .replace(/\u201c/g, '"')   // left smart double quote
    .replace(/\u201d/g, '"')   // right smart double quote
    .replace(/\u2018/g, "'")   // left smart single quote
    .replace(/\u2019/g, "'")   // right smart single quote
    .replace(/\u2013/g, '-')   // en dash
    .replace(/\u2014/g, '-')   // em dash
    .replace(/\u00a0/g, ' ')   // non-breaking space
    .replace(/\u200b/g, '');   // zero-width space
}

// ─── DOCX PARSING (RAW XML) ──────────────────────────────────────────
// We parse the raw XML to preserve the exact paragraph + newline structure.
// Breaks (<w:br/>) inside runs are converted to \n so we get the same
// line-split structure that python-docx exposes via paragraph.text.
async function parseDocxRaw(filePath) {
  const AdmZipLocal = require('adm-zip');
  const zip = new AdmZipLocal(filePath);
  const docXml = zip.readAsText('word/document.xml');

  const paragraphs = [];
  const paraRegex = /<w:p[ >][\s\S]*?<\/w:p>/g;
  let paraMatch;

  while ((paraMatch = paraRegex.exec(docXml)) !== null) {
    const paraXml = paraMatch[0];

    // Detect heading style
    const styleMatch = paraXml.match(/<w:pStyle w:val="([^"]+)"/);
    const style = styleMatch ? styleMatch[1] : 'normal';

    // Walk through <w:br/> and <w:t> elements in document order.
    // This correctly handles breaks that live inside <w:r> runs.
    let text = '';
    const elementRegex = /<w:br[^>]*\/?>|<w:t[^>]*>([^<]*)<\/w:t>/g;
    let elMatch;

    while ((elMatch = elementRegex.exec(paraXml)) !== null) {
      const tag = elMatch[0];
      if (tag.startsWith('<w:br')) {
        text += '\n';
      } else {
        text += elMatch[1]; // captured text content
      }
    }

    if (text.trim()) {
      paragraphs.push({ style, text: normalizeText(text) });
    }
  }

  return paragraphs;
}

// ─── MCQ EXTRACTION ──────────────────────────────────────────────────
function extractMCQs(paragraphs) {
  const questions = [];
  let questionNumber = 0;

  for (const para of paragraphs) {
    // Skip headings — they're module/topic names
    if (para.style && (para.style.startsWith('Heading') || para.style.match(/^heading/i))) {
      continue;
    }

    const text = para.text.trim();
    if (!text) continue;

    // Split by newline to separate question, options, and answer
    const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    if (lines.length < 3) continue; // Need at least question + 2 options + answer

    // First line is the question text
    const questionText = lines[0];

    // Extract options and answer
    const options = {};
    let answerLetter = null;

    for (let i = 1; i < lines.length; i++) {
      const line = lines[i];

      // Match option lines: a) ... , b) ... etc.
      const optMatch = line.match(/^([a-d])\)\s*(.+)$/);
      if (optMatch) {
        options[optMatch[1]] = optMatch[2].trim();
        continue;
      }

      // Match answer line: Answer: x) ...
      const ansMatch = line.match(/^Answer:\s*([a-d])\)/);
      if (ansMatch) {
        answerLetter = ansMatch[1];
      }
    }

    // Validate: we need at least 2 options and an answer
    const optionLetters = Object.keys(options);
    if (optionLetters.length < 2 || !answerLetter) continue;
    if (!options[answerLetter]) continue; // answer letter must match an option

    questionNumber++;
    questions.push({
      number: questionNumber,
      text: questionText,
      options,
      answerLetter,
      optionOrder: ['a', 'b', 'c', 'd'].filter(l => l in options)
    });
  }

  return questions;
}

// ─── GIFT FORMAT GENERATOR ───────────────────────────────────────────
function generateGIFT(questions) {
  const blocks = questions.map(q => {
    const lines = [];
    lines.push(`::Question ${q.number}::`);
    lines.push(q.text);
    lines.push('{');

    for (const letter of q.optionOrder) {
      const prefix = letter === q.answerLetter ? '=' : '~';
      lines.push(`${prefix}${q.options[letter]}`);
    }

    lines.push('}');
    return lines.join('\n');
  });

  return blocks.join('\n\n') + '\n';
}

// ─── API: UPLOAD & CONVERT ───────────────────────────────────────────
app.post('/api/upload', upload.array('files', 30), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }

    const results = [];
    const errors = [];

    for (const file of req.files) {
      try {
        const paragraphs = await parseDocxRaw(file.path);
        const questions = extractMCQs(paragraphs);

        if (questions.length === 0) {
          errors.push({
            filename: file.originalname,
            error: 'No valid MCQ questions found in document'
          });
          continue;
        }

        const giftContent = generateGIFT(questions);
        const outputName = file.originalname.replace(/\.docx$/i, '_gift.txt');

        results.push({
          filename: file.originalname,
          outputName,
          questionCount: questions.length,
          questions,
          giftContent
        });
      } catch (err) {
        errors.push({
          filename: file.originalname,
          error: err.message
        });
      } finally {
        // Clean up uploaded file
        try { fs.unlinkSync(file.path); } catch (e) {}
      }
    }

    res.json({ results, errors });
  } catch (err) {
    res.status(500).json({ error: 'Server error: ' + err.message });
  }
});

// ─── API: DOWNLOAD SINGLE FILE ───────────────────────────────────────
app.post('/api/download', express.json(), (req, res) => {
  const { content, filename } = req.body;
  if (!content || !filename) {
    return res.status(400).json({ error: 'Missing content or filename' });
  }

  res.setHeader('Content-Type', 'text/plain; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.send(content);
});

// ─── API: DOWNLOAD MERGED FILE ───────────────────────────────────────
app.post('/api/download-merged', express.json(), (req, res) => {
  const { contents, filename } = req.body;
  if (!contents || !Array.isArray(contents)) {
    return res.status(400).json({ error: 'Missing contents array' });
  }

  const merged = contents.join('\n\n');
  res.setHeader('Content-Type', 'text/plain; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename="${filename || 'merged_gift.txt'}"`);
  res.send(merged);
});

// ─── START SERVER ────────────────────────────────────────────────────
app.listen(PORT, '0.0.0.0', () => {
  console.log(`GIFT Converter running on port ${PORT}`);
});
