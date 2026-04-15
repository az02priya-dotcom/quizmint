const express = require('express');
const multer = require('multer');
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

const ALLOWED_EXTENSIONS = ['.docx', '.doc', '.txt', '.pdf', '.xlsx', '.xls', '.csv'];

const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (!ALLOWED_EXTENSIONS.includes(ext)) {
      return cb(new Error(`Unsupported file type: ${ext}. Accepted: ${ALLOWED_EXTENSIONS.join(', ')}`), false);
    }
    cb(null, true);
  },
  limits: { fileSize: 10 * 1024 * 1024 } // 10MB per file
});

// Serve static frontend
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json({ limit: '5mb' }));

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
    .replace(/\u200b/g, '')    // zero-width space
    .replace(/\r\n/g, '\n')    // normalize line endings
    .replace(/\r/g, '\n');
}

// ═══════════════════════════════════════════════════════════════════════
//  FILE PARSERS — Extract raw text from each file format
// ═══════════════════════════════════════════════════════════════════════

// ─── DOCX PARSER (RAW XML) ──────────────────────────────────────────
async function parseDocx(filePath) {
  const zip = new AdmZip(filePath);
  const docXml = zip.readAsText('word/document.xml');

  const paragraphs = [];
  const paraRegex = /<w:p[ >][\s\S]*?<\/w:p>/g;
  let paraMatch;

  while ((paraMatch = paraRegex.exec(docXml)) !== null) {
    const paraXml = paraMatch[0];
    const styleMatch = paraXml.match(/<w:pStyle w:val="([^"]+)"/);
    const style = styleMatch ? styleMatch[1] : 'normal';

    let text = '';
    const elementRegex = /<w:br[^>]*\/?>|<w:t[^>]*>([^<]*)<\/w:t>/g;
    let elMatch;
    while ((elMatch = elementRegex.exec(paraXml)) !== null) {
      if (elMatch[0].startsWith('<w:br')) {
        text += '\n';
      } else {
        text += elMatch[1];
      }
    }

    if (text.trim()) {
      paragraphs.push({ style, text: normalizeText(text) });
    }
  }

  return paragraphs;
}

// ─── PDF PARSER ──────────────────────────────────────────────────────
async function parsePdf(filePath) {
  const pdfParse = require('pdf-parse');
  const buffer = fs.readFileSync(filePath);
  const data = await pdfParse(buffer);
  const text = normalizeText(data.text);
  return [{ style: 'normal', text }];
}

// ─── TXT PARSER ──────────────────────────────────────────────────────
async function parseTxt(filePath) {
  const raw = fs.readFileSync(filePath, 'utf-8');
  const text = normalizeText(raw);
  return [{ style: 'normal', text }];
}

// ─── EXCEL / CSV PARSER ─────────────────────────────────────────────
async function parseExcel(filePath) {
  const XLSX = require('xlsx');
  const workbook = XLSX.readFile(filePath);
  const allText = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    if (rows.length === 0) continue;

    // Detect if first row is a header
    const firstRow = rows[0].map(c => String(c).toLowerCase().trim());
    const hasHeader = firstRow.some(h =>
      ['question', 'q', 'quiz', 'options', 'answer', 'correct'].includes(h) ||
      h.includes('question') || h.includes('option') || h.includes('answer')
    );

    if (hasHeader) {
      // Structured Excel: find column indices
      const qIdx = firstRow.findIndex(h => h.includes('question') || h === 'q' || h === 'quiz');
      const optAIdx = firstRow.findIndex(h => h === 'a' || h === 'option a' || h === 'option 1' || h === 'a)');
      const optBIdx = firstRow.findIndex(h => h === 'b' || h === 'option b' || h === 'option 2' || h === 'b)');
      const optCIdx = firstRow.findIndex(h => h === 'c' || h === 'option c' || h === 'option 3' || h === 'c)');
      const optDIdx = firstRow.findIndex(h => h === 'd' || h === 'option d' || h === 'option 4' || h === 'd)');
      const ansIdx = firstRow.findIndex(h => h.includes('answer') || h.includes('correct') || h === 'ans');

      // Also check for "options" column (all options in one cell)
      const optsIdx = firstRow.findIndex(h => h === 'options');

      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        const question = String(row[qIdx >= 0 ? qIdx : 0] || '').trim();
        if (!question) continue;

        let text = question + '\n';

        if (optAIdx >= 0) {
          // Separate columns for each option
          const optA = String(row[optAIdx] || '').trim();
          const optB = String(row[optBIdx >= 0 ? optBIdx : optAIdx + 1] || '').trim();
          const optC = String(row[optCIdx >= 0 ? optCIdx : optAIdx + 2] || '').trim();
          const optD = String(row[optDIdx >= 0 ? optDIdx : optAIdx + 3] || '').trim();
          if (optA) text += ` a) ${optA}\n`;
          if (optB) text += ` b) ${optB}\n`;
          if (optC) text += ` c) ${optC}\n`;
          if (optD) text += ` d) ${optD}\n`;
        } else if (optsIdx >= 0) {
          // All options in one cell, separated by newlines or semicolons
          const opts = String(row[optsIdx] || '').split(/[;\n]/).map(o => o.trim()).filter(Boolean);
          opts.forEach((o, i) => {
            const letter = String.fromCharCode(97 + i); // a, b, c, d...
            text += ` ${letter}) ${o}\n`;
          });
        }

        const answer = String(row[ansIdx >= 0 ? ansIdx : row.length - 1] || '').trim();
        if (answer) text += ` Answer: ${answer}\n`;

        allText.push({ style: 'normal', text: normalizeText(text) });
      }
      return allText;
    } else {
      // Unstructured: just concatenate all cells as text
      const lines = rows.map(row => row.map(c => String(c)).join(' ')).join('\n');
      return [{ style: 'normal', text: normalizeText(lines) }];
    }
  }

  return allText;
}

// ─── PASTE TEXT PARSER ───────────────────────────────────────────────
function parsePastedText(rawText) {
  const text = normalizeText(rawText);
  return [{ style: 'normal', text }];
}

// ═══════════════════════════════════════════════════════════════════════
//  UNIVERSAL MCQ EXTRACTOR
//  Handles many question formats from any source
// ═══════════════════════════════════════════════════════════════════════
function extractMCQsUniversal(paragraphs) {
  const questions = [];
  let questionNumber = 0;

  // First try the structured approach (original docx-style: each paragraph = 1 question)
  // Only use this when we have multiple paragraphs (like from a Word doc)
  if (paragraphs.length > 1) {
    for (const para of paragraphs) {
      if (para.style && (para.style.startsWith('Heading') || para.style.match(/^heading/i))) {
        continue;
      }

      const text = para.text.trim();
      if (!text) continue;

      const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
      if (lines.length < 3) continue;

      // Try structured extraction
      const extracted = tryExtractFromLines(lines);
      if (extracted) {
        questionNumber++;
        extracted.number = questionNumber;
        questions.push(extracted);
      }
    }

    // If structured approach found questions, return them
    if (questions.length > 0) return questions;
  }

  // For single-paragraph input (paste, txt, pdf) or when structured failed,
  // join all text and use block-based extraction
  const fullText = paragraphs.map(p => p.text).join('\n');
  return extractFromFreeText(fullText);
}

// ─── TRY EXTRACT FROM STRUCTURED LINES ──────────────────────────────
function tryExtractFromLines(lines) {
  let questionText = null;
  const options = {};
  let answerLetter = null;
  let answerText = null;

  for (const line of lines) {
    // Match answer line first (various formats)
    const ansMatch = line.match(/^(?:Answer|Ans|Correct(?:\s*Answer)?)\s*[:=\-]\s*(?:([a-dA-D])\s*[).\]:]?\s*)?(.*)$/i);
    if (ansMatch) {
      if (ansMatch[1]) {
        answerLetter = ansMatch[1].toLowerCase();
      }
      if (ansMatch[2]) {
        answerText = ansMatch[2].trim();
      }
      continue;
    }

    // Match option lines: a), A), a., A., (a), (A) — letters only
    const optMatch = line.match(/^(?:\(?([a-dA-D])\s*[).\]:]\s*)(.+)$/);
    if (optMatch) {
      const letter = optMatch[1].toLowerCase();
      options[letter] = optMatch[2].trim();
      continue;
    }

    // If no question text yet, this is the question
    if (!questionText) {
      // Strip leading question number: "1.", "Q1.", "Question 1:", "1)", etc.
      const cleaned = line.replace(/^(?:Q(?:uestion)?\s*\.?\s*)?(\d+)\s*[.):\-]\s*/i, '').trim();
      questionText = cleaned || line;
    }
  }

  // Resolve answer
  if (!answerLetter && answerText) {
    // Try to match answer text to an option
    for (const [letter, text] of Object.entries(options)) {
      if (text.toLowerCase() === answerText.toLowerCase()) {
        answerLetter = letter;
        break;
      }
    }
    // Try partial match
    if (!answerLetter) {
      for (const [letter, text] of Object.entries(options)) {
        if (text.toLowerCase().includes(answerText.toLowerCase()) ||
            answerText.toLowerCase().includes(text.toLowerCase())) {
          answerLetter = letter;
          break;
        }
      }
    }
    // Try if answerText is just a letter
    if (!answerLetter && /^[a-dA-D]$/.test(answerText)) {
      answerLetter = answerText.toLowerCase();
    }
  }

  const optionLetters = Object.keys(options);
  if (!questionText || optionLetters.length < 2 || !answerLetter || !options[answerLetter]) {
    return null;
  }

  return {
    number: 0,
    text: questionText,
    options,
    answerLetter,
    optionOrder: ['a', 'b', 'c', 'd'].filter(l => l in options)
  };
}

// ─── EXTRACT FROM FREE-FORM TEXT ────────────────────────────────────
function extractFromFreeText(fullText) {
  const questions = [];
  let questionNumber = 0;

  // Split into question blocks using various separators
  // Look for patterns: "1.", "Q1.", "Question 1:", numbered questions, or double newlines
  const blocks = splitIntoQuestionBlocks(fullText);

  for (const block of blocks) {
    const lines = block.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    if (lines.length < 3) continue;

    const extracted = tryExtractFromLines(lines);
    if (extracted) {
      questionNumber++;
      extracted.number = questionNumber;
      questions.push(extracted);
    }
  }

  return questions;
}

// ─── SPLIT TEXT INTO QUESTION BLOCKS ────────────────────────────────
function splitIntoQuestionBlocks(text) {
  // Strategy 1: Split by question number patterns
  const numberedPattern = /(?:^|\n)\s*(?:Q(?:uestion)?\s*\.?\s*)?\d+\s*[.):\-]\s*/gi;
  const matches = [...text.matchAll(numberedPattern)];

  if (matches.length >= 2) {
    const blocks = [];
    for (let i = 0; i < matches.length; i++) {
      const start = matches[i].index;
      const end = i + 1 < matches.length ? matches[i + 1].index : text.length;
      blocks.push(text.substring(start, end).trim());
    }
    return blocks;
  }

  // Strategy 2: Split by double newlines (blank line between questions)
  const doubleNewlineBlocks = text.split(/\n\s*\n/).filter(b => b.trim().length > 0);
  if (doubleNewlineBlocks.length >= 2) {
    return doubleNewlineBlocks;
  }

  // Strategy 3: Return as single block
  return [text];
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

// ═══════════════════════════════════════════════════════════════════════
//  API ROUTES
// ═══════════════════════════════════════════════════════════════════════

// ─── API: UPLOAD & CONVERT (files) ──────────────────────────────────
app.post('/api/upload', upload.array('files', 30), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }

    const results = [];
    const errors = [];

    for (const file of req.files) {
      try {
        const ext = path.extname(file.originalname).toLowerCase();
        let paragraphs;

        // Route to correct parser based on file type
        switch (ext) {
          case '.docx':
          case '.doc':
            paragraphs = await parseDocx(file.path);
            break;
          case '.pdf':
            paragraphs = await parsePdf(file.path);
            break;
          case '.txt':
            paragraphs = await parseTxt(file.path);
            break;
          case '.xlsx':
          case '.xls':
          case '.csv':
            paragraphs = await parseExcel(file.path);
            break;
          default:
            throw new Error(`Unsupported file type: ${ext}`);
        }

        const questions = extractMCQsUniversal(paragraphs);

        if (questions.length === 0) {
          errors.push({
            filename: file.originalname,
            error: 'No valid MCQ questions found. Make sure the file contains questions with options (a/b/c/d) and marked answers.'
          });
          continue;
        }

        const giftContent = generateGIFT(questions);
        const baseName = file.originalname.replace(/\.[^.]+$/, '');
        const outputName = `${baseName}_gift.txt`;

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
        try { fs.unlinkSync(file.path); } catch (e) {}
      }
    }

    res.json({ results, errors });
  } catch (err) {
    res.status(500).json({ error: 'Server error: ' + err.message });
  }
});

// ─── API: PASTE & CONVERT (text) ────────────────────────────────────
app.post('/api/paste', (req, res) => {
  try {
    const { text } = req.body;
    if (!text || !text.trim()) {
      return res.status(400).json({ error: 'No text provided' });
    }

    const paragraphs = parsePastedText(text);
    const questions = extractMCQsUniversal(paragraphs);

    if (questions.length === 0) {
      return res.status(400).json({
        error: 'No valid MCQ questions found. Make sure the text contains questions with options (a/b/c/d) and marked answers.'
      });
    }

    const giftContent = generateGIFT(questions);

    res.json({
      results: [{
        filename: 'Pasted Text',
        outputName: 'pasted_gift.txt',
        questionCount: questions.length,
        questions,
        giftContent
      }],
      errors: []
    });
  } catch (err) {
    res.status(500).json({ error: 'Server error: ' + err.message });
  }
});

// ─── API: DOWNLOAD ──────────────────────────────────────────────────
app.post('/api/download', express.json(), (req, res) => {
  const { content, filename } = req.body;
  if (!content || !filename) {
    return res.status(400).json({ error: 'Missing content or filename' });
  }
  res.setHeader('Content-Type', 'text/plain; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.send(content);
});

// ─── API: DOWNLOAD MERGED ───────────────────────────────────────────
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

// ─── START SERVER ───────────────────────────────────────────────────
app.listen(PORT, '0.0.0.0', () => {
  console.log(`QuizMint running on port ${PORT}`);
});
