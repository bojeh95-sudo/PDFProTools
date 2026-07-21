import express from 'express';
import path from 'path';
import { promises as fs } from 'node:fs';
import crypto from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { GoogleGenAI } from '@google/genai';
import pkg from 'pdf-encrypt-decrypt';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType } from 'docx';
import { PDFDocument, rgb } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import { createRequire } from 'node:module';
import bidiFactory from 'bidi-js';
import reshaper from 'arabic-persian-reshaper';

const { ArabicShaper } = reshaper;

const { encryptPDF, PDFPermission } = pkg;

const requireFn = createRequire(path.join(process.cwd(), 'server.js'));
const rawPdfParse = requireFn('pdf-parse');

function detectArabicOrientation(text: string): 'reversed' | 'logical' | 'unknown' {
  const words = text.split(/\s+/).map(w => w.trim()).filter(w => /[\u0600-\u06FF]/.test(w));
  if (words.length === 0) return 'unknown';

  let reversedCount = 0;
  let logicalCount = 0;

  for (const word of words) {
    if (word.startsWith('ال') || word.startsWith('بال') || word.startsWith('وال') || word.startsWith('لل') || word.startsWith('في') || word.startsWith('من') || word.startsWith('على')) {
      logicalCount++;
    }
    if (word.endsWith('لا') || word.endsWith('لاب') || word.endsWith('لاو') || word.endsWith('لل') || word.endsWith('يف') || word.endsWith('نم') || word.endsWith('ىلع')) {
      reversedCount++;
    }
  }

  if (reversedCount > logicalCount) return 'reversed';
  if (logicalCount > reversedCount) return 'logical';
  return 'unknown';
}

function fixArabicTextOrder(text: string): string {
  if (!text) return '';
  const arabicRegex = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;
  const arabicRunRegex = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+/g;

  const lines = text.split('\n');
  const correctedLines = lines.map(line => {
    const trimmed = line.trim();
    if (!trimmed || !arabicRegex.test(trimmed)) {
      return line;
    }
    
    const orientation = detectArabicOrientation(trimmed);
    if (orientation === 'logical') {
      return line;
    }
    
    const tokens = line.split(/(\s+)/);
    const reversedTokens = tokens.reverse();
    const correctedTokens = reversedTokens.map(token => {
      if (token.trim() && arabicRegex.test(token)) {
        return token.replace(arabicRunRegex, match => match.split('').reverse().join(''));
      }
      return token;
    });
    
    return correctedTokens.join('');
  });
  
  return correctedLines.join('\n');
}

function cleanPageDecorations(text: string): string {
  if (!text) return '';
  const lines = text.split('\n');
  const cleanedLines = lines.map(line => {
    const trimmed = line.trim();
    const isPageNum = 
      /^\s*-\s*\d+\s*-\s*$/i.test(trimmed) || 
      /^\s*--\s*\d+\s*(of|من)\s*\d+\s*--\s*$/i.test(trimmed) ||
      /^\s*page\s+\d+\s*(of|من)\s*\d+\s*$/i.test(trimmed) ||
      /^\s*page\s+\d+\s*$/i.test(trimmed) ||
      /^\s*صفحة\s+\d+\s*(من|of)\s*\d+\s*$/i.test(trimmed) ||
      /^\s*صفحة\s+\d+\s*$/i.test(trimmed) ||
      /^\s*\[\s*Page\s+\d+\s*\]\s*$/i.test(trimmed) ||
      /^\s*\[\s*صفحة\s+\d+\s*\]\s*$/i.test(trimmed) ||
      /^\s*\d+\s*\/\s*\d+\s*$/i.test(trimmed);
    
    if (isPageNum) {
      return '';
    }
    return line;
  });
  return cleanedLines.filter(line => line !== '').join('\n');
}

function cleanCsvOfPageDecorations(csvText: string): string {
  if (!csvText) return '';
  const lines = csvText.split('\n');
  const filtered = lines.filter(line => {
    const normalized = line.replace(/["']/g, '').trim();
    const isPageNum = 
      /^\s*-\s*\d+\s*-\s*$/i.test(normalized) || 
      /^\s*--\s*\d+\s*(of|من)\s*\d+\s*--\s*$/i.test(normalized) ||
      /^\s*page\s+\d+\s*(of|من)\s*\d+\s*$/i.test(normalized) ||
      /^\s*page\s+\d+\s*$/i.test(normalized) ||
      /^\s*صفحة\s+\d+\s*(من|of)\s*\d+\s*$/i.test(normalized) ||
      /^\s*صفحة\s+\d+\s*$/i.test(normalized) ||
      /^\s*\[\s*Page\s+\d+\s*\]\s*$/i.test(normalized) ||
      /^\s*\[\s*صفحة\s+\d+\s*\]\s*$/i.test(normalized) ||
      /^\s*\d+\s*\/\s*\d+\s*$/i.test(normalized);
      
    return !isPageNum;
  });
  return filtered.join('\n');
}

function parseCSV(csvText: string): string[][] {
  const lines: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let inQuotes = false;
  
  for (let i = 0; i < csvText.length; i++) {
    const char = csvText[i];
    const nextChar = csvText[i + 1];
    
    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        cell += '"';
        i++; // skip next quote
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      row.push(cell);
      cell = '';
    } else if ((char === '\r' || char === '\n') && !inQuotes) {
      if (char === '\r' && nextChar === '\n') {
        i++; // skip \n
      }
      row.push(cell);
      lines.push(row);
      row = [];
      cell = '';
    } else {
      cell += char;
    }
  }
  
  if (cell || row.length > 0) {
    row.push(cell);
    lines.push(row);
  }
  
  return lines;
}

function stringifyCSV(rows: string[][]): string {
  return rows.map(row => {
    return row.map(cell => {
      if (cell.includes(',') || cell.includes('"') || cell.includes('\n') || cell.includes('\r')) {
        const escaped = cell.replace(/"/g, '""');
        return `"${escaped}"`;
      }
      return cell;
    }).join(',');
  }).join('\n');
}

function processExcelCSV(csvText: string): string {
  if (!csvText) return '';
  
  const rows = parseCSV(csvText);
  const processedRows = rows.map(row => {
    return row.map(cell => {
      const trimmedCell = cell.trim();
      const hasArabic = /[\u0600-\u06FF]/.test(trimmedCell);
      if (!hasArabic) {
        return cell;
      }
      
      const orientation = detectArabicOrientation(trimmedCell);
      if (orientation === 'logical') {
        return cell;
      }
      
      const tokens = cell.split(/(\s+)/);
      const reversedTokens = tokens.reverse();
      const correctedTokens = reversedTokens.map(token => {
        const arabicRegex = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;
        const arabicRunRegex = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+/g;
        if (token.trim() && arabicRegex.test(token)) {
          return token.replace(arabicRunRegex, match => match.split('').reverse().join(''));
        }
        return token;
      });
      
      return correctedTokens.join('');
    });
  });
  
  return stringifyCSV(processedRows);
}

async function extractTextFromPdfBuffer(buffer: Buffer): Promise<string> {
  const { PDFParse } = rawPdfParse as any;
  if (!PDFParse) {
    throw new Error('PDFParse constructor not found in pdf-parse module.');
  }
  const parser = new PDFParse({ data: buffer });
  const textResult = await parser.getText();
  const rawText = textResult?.text || '';
  const fixedText = fixArabicTextOrder(rawText);
  return cleanPageDecorations(fixedText);
}



export const app = express();

async function startServer() {
  const PORT = 3000;

  // Support large base64 uploads for PDFs and images
  app.use(express.json({ limit: '50mb' }));
  app.use(express.urlencoded({ limit: '50mb', extended: true }));

  // API Routes
  app.post('/api/gemini/process', async (req, res) => {
    try {
      const { 
        task, fileData, mimeType, text, targetLang, question, secondFileData, secondMimeType,
        extractPagesMode, extractPagesRange, extractFormat,
        summaryLength, summaryFormat,
        preserveLayout,
        ocrLanguage, ocrOutputType,
        formatAutoAlign, formatUnifyFonts,
        redactPiiEmail, redactPiiPhone, redactPiiId, redactPiiAddress, redactKeywords,
        compareMode, diffView
      } = req.body;
      
      const apiKey = process.env.GEMINI_API_KEY;
      
      if (!apiKey) {
        return res.status(500).json({
          error: 'GEMINI_API_KEY environment variable is not configured. Please define it in Settings > Secrets.'
        });
      }

      const ai = new GoogleGenAI({
        apiKey,
        httpOptions: {
          headers: {
            'User-Agent': 'aistudio-build',
          }
        }
      });

      let userPrompt = '';
      switch (task) {
        case 'summarize':
          userPrompt = `Please provide a professional, clean, highly readable, structured summary of this document. 
- Summary Length: ${summaryLength || 'medium'}
- Display Format: ${summaryFormat === 'bullets' ? 'bullet points with clear headlines' : 'connected fluid paragraphs'}
Make it elegant, structured, and visually polished using Markdown. Respect the document's original language (e.g., summarize in Arabic if original is Arabic).`;
          break;
        case 'translate':
          userPrompt = `Please translate the entire content of this document into the following language/locale: ${targetLang || 'English'}. 
- Preserve Original Formatting & Font Scale locations: ${preserveLayout ? 'Yes, strictly translate strings while referencing original visual alignments and page coordinates.' : 'No, optimize the translated output into a fluid markdown document.'}
Format the translated output into beautiful, organized Markdown with clear headings.`;
          break;
        case 'extract':
          userPrompt = `Please extract key data and text from this document. 
- Target Range Selection: ${extractPagesMode === 'range' && extractPagesRange ? `Only focus on page range "${extractPagesRange}".` : 'Extract from all pages.'}
- Target Structure: ${extractFormat === 'structured' ? 'structured bullet points, tables, and bold section headers' : 'plain, clean raw text block'}
Ensure all extracted numbers, emails, addresses, dates, or contact info are listed clearly in the original language.`;
          break;
        case 'ocr':
          userPrompt = `Perform ultra-accurate OCR text recognition on this scanned document or image. 
- Primary Document Language: ${ocrLanguage || 'Arabic'}
- Output Mode: ${ocrOutputType === 'searchable_pdf' ? 'Re-create a structured, coordinates-aware text outline or markdown' : 'Extract clean raw text paragraphs with perfect spacing'}
Ensure every character, sign, and word is perfectly recognized and formatted.`;
          break;
        case 'format':
          userPrompt = `Analyze this document, resolve any visual layout flaws, and structure it into a beautiful, highly professional, publication-ready layout. 
- Auto-Align Paragraphs and text direction: ${formatAutoAlign ? 'Yes, auto-align text blocks and direction based on paragraph language (LTR/RTL).' : 'No'}
- Unify Document Fonts & missing character shapes: ${formatUnifyFonts ? 'Yes, recommend a cohesive font selection (such as Inter for sans, JetBrains Mono for codes) and format clean symbols.' : 'No'}
Return a pristine Markdown layout with consistent padding, headings, and lists.`;
          break;
        case 'pdf2word':
          userPrompt = 'Analyze the provided PDF file. Extract its complete content, maintaining the correct reading order, sections, headings, lists, and tables. Output the extracted elements as a valid JSON array of block objects. Each block object must have the following structure: - "type": Choose from "title", "heading1", "heading2", "paragraph", "bullet", "numbered", "table". - "text": For title, heading1, heading2, paragraph, bullet, numbered. - "rows": For "table", an array of rows, where each row is an array of cell strings (e.g. [["Col1", "Col2"], ["Val1", "Val2"]]). - "alignment": "left" | "center" | "right" (optional, default left). - "bold": boolean (optional). Return ONLY the raw JSON array. Do not put markdown formatting or backticks around it. If you must use formatting, do not use backticks outside of the JSON. Make sure the JSON is fully valid and parseable.';
          break;
        case 'chat':
          userPrompt = question 
            ? `Answer the following question about the provided document: "${question}". 
Crucial Requirement: Since the user is viewing the PDF pages visually, if the answer is found on a specific page, you MUST explicitly cite the page number as "[Page X]" or "[صفحة X]" so they can click it and highlight that page. Respond in the same language as the user's question or the document.` 
            : 'Please provide a warm, professional, brief summary of this document and ask how you can help me understand it better. Make sure to cite any relevant page numbers as "[Page X]".';
          break;
        case 'redact':
          userPrompt = `Scan this document for sensitive, personally identifiable information (PII) and redact it completely by replacing it with "[REDACTED]". 
Target types to redact:
${redactPiiEmail ? '- Email addresses (e.g., test@example.com)\n' : ''}${redactPiiPhone ? '- Telephone, mobile, and fax numbers\n' : ''}${redactPiiId ? '- National IDs, residency cards, passports, driver licenses\n' : ''}${redactPiiAddress ? '- Postal codes, home addresses, coordinates, cities\n' : ''}${redactKeywords ? `- Custom terms or names specified here: "${redactKeywords}"\n` : ''}
Ensure all other non-sensitive content remains completely intact and word-for-word identical.`;
          break;
        case 'compare':
          userPrompt = `You are a professional document auditor. Carefully compare the first document/image with the second document/image (both provided). 
- Comparison Mode: ${compareMode === 'visual' ? 'Visual page-by-page changes and alignments' : 'Textual sentence-level analysis and word differences'}
- Filter View: ${diffView === 'all' ? 'All additions and deletions' : diffView === 'additions' ? 'Only additions and new insertions' : 'Only deletions and removals'}
Provide a clean, highly detailed, and structured Markdown report summarizing the differences with side-by-side notes or tables.`;
          break;
        case 'pdf2excel':
          userPrompt = `You are a high-fidelity tabular data extractor. Analyze the provided document and locate all tables, row data, and key-value matrices.
Reconstruct the extracted tabular data into a valid Comma-Separated Values (CSV) structure.
- Columns must be separated by commas, rows by standard Unix newlines.
- Cells containing commas, double-quotes, or newlines must be enclosed in double quotes (RFC 4180 format).
- Keep all table columns in their exact original order from left to right as they appear visually in the document. Do NOT reverse the order of columns or rows, even if the document contains Arabic text or is a mixed-language statement. The leftmost column in the physical document must remain the first column in the CSV, and the rightmost column must be the last column in the CSV. Purely English columns, dates, and numeric columns (e.g., Balance, Credit, Debit, Date, Transaction ID, Reference) must keep their original position and standard left-to-right order.
- Do NOT reverse the Arabic letters or words yourself in the CSV output. Keep them in their natural, readable spelling and reading order. Let the system handle any letter shaping or direction processing. The cells should contain standard, readable UTF-8 text.
- Preserve original text encoding, especially Arabic characters or special symbols (keep them native).
- If the document is primarily regular text, a letter, or an essay without clear structured tables, do not arbitrarily split sentences across multiple columns or invent tables. Instead, place each logical paragraph/line in the first column (one column per row) sequentially, properly quoted.
- Remove all document page numbers, page headers, or page footers (e.g. "Page X of Y", "-- 1 --") so they do not pollute the output.
- Return ONLY the raw CSV text. Do not put markdown headers, introductory text, or code block triple backticks around the CSV.`;
          break;
        case 'pdf2excel_old_disabled':
          userPrompt = `You are a high-fidelity tabular data extractor. Analyze the provided document and locate all tables, row data, and key-value matrices.
Reconstruct the extracted tabular data into a valid Comma-Separated Values (CSV) structure.
- Columns must be separated by commas, rows by standard Unix newlines.
- Cells containing commas, double-quotes, or newlines must be enclosed in double quotes (RFC 4180 format).
- Preserve original text encoding, especially Arabic characters or special symbols (keep them native).
- If the document is primarily regular text, a letter, or an essay without clear structured tables, do not arbitrarily split sentences across multiple columns or invent tables. Instead, place each logical paragraph/line in the first column (one column per row) sequentially, properly quoted.
- Remove all document page numbers, page headers, or page footers (e.g. "Page X of Y", "-- 1 --") so they do not pollute the output.
Return ONLY the raw CSV text. Do not put markdown headers, introductory text, or code block triple backticks around the CSV.`;
          break;
        case 'pdf2ppt':
          userPrompt = `You are an expert presentation planner and content writer. Translate the provided document into a highly professional, slide-by-slide PowerPoint presentation design outline.
Each slide block must contain:
- Slide Number and a clear, descriptive Title
- Beautifully structured bullet points of the key concepts and data
- Detailed Speaker Notes / Presenter guidelines summarizing statistics or deeper context
Ensure it is fully customized to the document's content and written in a highly clear, professional, and elegant tone. Respect the original document language.`;
          break;
        default:
          return res.status(400).json({ error: 'Invalid task requested.' });
      }

      // Build contents parts
      const parts: any[] = [];
      
      if (fileData && mimeType) {
        parts.push({
          inlineData: {
            data: fileData,
            mimeType: mimeType
          }
        });
      }

      if (secondFileData && secondMimeType) {
        parts.push({
          inlineData: {
            data: secondFileData,
            mimeType: secondMimeType
          }
        });
      }

      let pdfText = text || '';
      if (fileData && mimeType === 'application/pdf' && (!pdfText || pdfText.trim().length < 50)) {
        try {
          const pdfBuffer = Buffer.from(fileData, 'base64');
          pdfText = await extractTextFromPdfBuffer(pdfBuffer);
        } catch (pdfErr) {
          console.error('Server-side PDF text extraction failed:', pdfErr);
        }
      }

      if (pdfText) {
        parts.push({
          text: `Here is the first document text to process:\n\n${pdfText}`
        });
      }

      // Append prompt
      parts.push({
        text: userPrompt
      });

      const response = await ai.models.generateContent({
        model: 'gemini-3.5-flash',
        contents: { parts: parts }
      });

      let resultText = response.text || '';

      if (task === 'pdf2excel') {
        console.log('[pdf2excel] Extracted PDF Text length:', pdfText ? pdfText.length : 0);
        console.log('[pdf2excel] Extracted PDF Text sample:', pdfText ? pdfText.substring(0, 500) : 'EMPTY');
        console.log('[pdf2excel] Raw Gemini response length:', resultText ? resultText.length : 0);
        console.log('[pdf2excel] Raw Gemini response sample:', resultText ? resultText.substring(0, 500) : 'EMPTY');

        // 1. Clean the result: strip any markdown code blocks
        let cleanedResult = resultText;
        cleanedResult = cleanedResult.replace(/^```[a-zA-Z]*\s*/gm, '');
        cleanedResult = cleanedResult.replace(/```\s*$/gm, '');
        cleanedResult = cleanedResult.trim();

        // 2. Check if a valid table / CSV structure is present
        const lines = cleanedResult.split('\n').map(l => l.trim()).filter(Boolean);
        const hasCommas = lines.some(line => line.includes(',') || line.includes(';') || line.includes('\t'));
        const looksLikeConversation = cleanedResult.toLowerCase().includes('sorry') || 
                                     cleanedResult.toLowerCase().includes('i cannot') || 
                                     cleanedResult.toLowerCase().includes('no table') || 
                                     cleanedResult.includes('لا يوجد') || 
                                     cleanedResult.includes('لا أستطيع');

        let finalResult = '';

        if (!cleanedResult || !hasCommas || looksLikeConversation || lines.length < 2) {
          console.log('[pdf2excel] Detected empty, conversational, or invalid CSV response. Triggering robust fallback...');
          // Trigger fallback: place each line of pdfText in a separate row (A1, A2...)
          if (pdfText && pdfText.trim()) {
            const pdfLines = pdfText.split('\n')
              .map(line => line.trim())
              .filter(Boolean);

            const csvRows = pdfLines.map(line => {
              // Escape double quotes and enclose in double quotes
              const escaped = line.replace(/"/g, '""');
              return `"${escaped}"`;
            });
            finalResult = csvRows.join('\n');
            console.log('[pdf2excel] Fallback generated CSV with', pdfLines.length, 'rows.');
          } else {
            // If even pdfText is empty (e.g. Scanned/OCR fallback empty), return a readable layout row so it is never blank
            finalResult = `"Document Content / محتوى المستند"\n"No extractable text or tables found in this PDF file. / لم يتم العثور على نصوص أو جداول قابلة للاستخراج في هذا الملف."`;
            console.log('[pdf2excel] No PDF text available. Generated default notice row.');
          }
        } else {
          finalResult = cleanedResult;
        }

        // Clean finalResult from any remaining page decorations / headers / footers
        finalResult = cleanCsvOfPageDecorations(finalResult);

        // Crucial: Fix Arabic letters/words reversal so it matches SVG and shows correctly in Excel
        finalResult = processExcelCSV(finalResult);

        console.log('[pdf2excel] Final CSV Result length:', finalResult.length);
        resultText = finalResult;
      }

      res.json({ result: resultText });
    } catch (err: any) {
      console.error('Error calling Gemini API:', err);
      res.status(500).json({ error: err.message || 'Failed to process document with AI' });
    }
  });

  // Secure PDF Protection route using pdf-encrypt-decrypt (pure JS wrapper of Go-based shared library)
  app.post('/api/pdf/protect', async (req, res) => {
    try {
      const { fileData, password } = req.body;
      if (!fileData || !password) {
        return res.status(400).json({ error: 'Missing fileData or password.' });
      }

      // Convert base64 data to a Buffer
      const buffer = Buffer.from(fileData, 'base64');

      // Encrypt PDF using the pdf-encrypt-decrypt library with all permissions enabled by default
      const securedBuffer = encryptPDF(buffer, password, password, [PDFPermission.ALL]);

      // Convert the encrypted Buffer back to base64
      const securedBase64 = securedBuffer.toString('base64');

      const format = req.query.format || req.headers['accept'];
      const wantsBinary = format === 'binary' || (format && format.toString().includes('octet-stream'));

      if (wantsBinary) {
        res.setHeader('Content-Type', 'application/pdf');
        const filename = 'protected.pdf';
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"; filename*=UTF-8''${encodeURIComponent(filename)}`);
        res.send(securedBuffer);
      } else {
        res.json({ securedData: securedBase64 });
      }
    } catch (err: any) {
      console.error('Error in /api/pdf/protect:', err);
      res.status(500).json({ error: err.message || 'Failed to encrypt PDF' });
    }
  });

  // Real PDF to Word converter using pdf-parse and docx
  app.post('/api/pdf/pdf2word', async (req, res) => {
    try {
      const { fileData } = req.body;
      if (!fileData) {
        return res.status(400).json({ error: 'Missing fileData' });
      }

      // Convert base64 data to a Buffer
      const buffer = Buffer.from(fileData, 'base64');

      // Parse PDF using pdf-parse
      let extractedText = '';
      try {
        extractedText = await extractTextFromPdfBuffer(buffer);
      } catch (parseErr: any) {
        console.error('pdf-parse failed:', parseErr);
        return res.status(400).json({ error: 'Failed to read PDF file content: ' + parseErr.message });
      }
      if (!extractedText.trim()) {
        return res.status(400).json({ error: 'The uploaded PDF does not contain any readable text.' });
      }

      const isArabic = /[\u0600-\u06FF]/.test(extractedText);

      let blocks: any[] = [];
      const hasGemini = !!process.env.GEMINI_API_KEY;

      if (hasGemini) {
        try {
          const ai = new GoogleGenAI({
            apiKey: process.env.GEMINI_API_KEY!,
            httpOptions: {
              headers: {
                'User-Agent': 'aistudio-build',
              }
            }
          });

          // Truncate text to avoid model limits
          const textSample = extractedText.substring(0, 40000);

          const response = await ai.models.generateContent({
            model: 'gemini-3.5-flash',
            contents: {
              parts: [
                {
                  text: `Analyze the provided text extracted from a PDF. Reconstruct the document structure (titles, headings, paragraphs, lists, tables) as a valid JSON array of block objects.
Each block object must have the following structure:
- "type": "title" | "heading1" | "heading2" | "paragraph" | "bullet" | "numbered" | "table"
- "text": string (for non-table blocks)
- "rows": array of arrays of strings (for tables)
- "alignment": "left" | "center" | "right" (optional, default "left")
- "bold": boolean (optional)

Example output:
[
  {"type": "title", "text": "Document Title", "alignment": "center"},
  {"type": "heading1", "text": "Section 1"},
  {"type": "paragraph", "text": "This is a paragraph of the section."}
]

Return ONLY the raw JSON array. Do not put markdown formatting or backticks around it. If you must use formatting, do not use backticks outside of the JSON. Make sure the JSON is fully valid and parseable.

Here is the extracted text:
${textSample}`
                }
              ]
            }
          });

          const rawResult = response.text || '';
          let cleaned = rawResult.trim();
          if (cleaned.startsWith('```')) {
            cleaned = cleaned.replace(/^```json\s*/i, '');
            cleaned = cleaned.replace(/```\s*$/, '');
            cleaned = cleaned.trim();
          }
          blocks = JSON.parse(cleaned);
        } catch (aiErr) {
          console.error('Gemini failed to structure PDF text, using local parser:', aiErr);
        }
      }

      // Fallback local parser if Gemini is absent or fails
      if (!blocks || blocks.length === 0) {
        blocks = [];
        const lines = extractedText.split('\n');
        let currentParagraph = '';

        for (let i = 0; i < lines.length; i++) {
          const line = lines[i].trim();
          if (!line) {
            if (currentParagraph) {
              blocks.push({ type: 'paragraph', text: currentParagraph });
              currentParagraph = '';
            }
            continue;
          }

          // Check for lists
          const bulletMatch = line.match(/^[\u2022\-\*]\s*(.*)/);
          const numMatch = line.match(/^([0-9]+)[\.\)]\s*(.*)/);

          if (bulletMatch) {
            if (currentParagraph) {
              blocks.push({ type: 'paragraph', text: currentParagraph });
              currentParagraph = '';
            }
            blocks.push({ type: 'bullet', text: bulletMatch[1] });
          } else if (numMatch) {
            if (currentParagraph) {
              blocks.push({ type: 'paragraph', text: currentParagraph });
              currentParagraph = '';
            }
            blocks.push({ type: 'numbered', text: numMatch[2] });
          } else if (line.length < 80 && (line.match(/^[A-Z0-9\s]+$/) || line.endsWith(':'))) {
            // Likely a heading
            if (currentParagraph) {
              blocks.push({ type: 'paragraph', text: currentParagraph });
              currentParagraph = '';
            }
            blocks.push({ type: 'heading1', text: line });
          } else {
            if (currentParagraph) {
              currentParagraph += ' ' + line;
            } else {
              currentParagraph = line;
            }
          }
        }
        if (currentParagraph) {
          blocks.push({ type: 'paragraph', text: currentParagraph });
        }
      }

      // Filter empty blocks
      blocks = blocks.filter(b => b && (b.text || (b.rows && b.rows.length > 0)));

      // Generate the Docx document binary using docx package
      const docChildren: any[] = [];

      for (const b of blocks) {
        if (b.type === 'table') {
          if (b.rows && b.rows.length > 0) {
            const tableRows = b.rows.map((row: string[]) => {
              return new TableRow({
                children: row.map((cellText: string) => {
                  return new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: cellText || '',
                            size: 22, // 11pt
                          }),
                        ],
                        spacing: { before: 80, after: 80 },
                        alignment: isArabic ? AlignmentType.RIGHT : AlignmentType.LEFT,
                      }),
                    ],
                    shading: {
                      fill: 'F8FAFC',
                    },
                  });
                }),
              });
            });

            docChildren.push(
              new Table({
                rows: tableRows,
                width: {
                  size: 100,
                  type: WidthType.PERCENTAGE,
                },
              })
            );
            docChildren.push(new Paragraph({ spacing: { after: 200 } }));
          }
        } else {
          let heading: any = undefined;
          let size = 24; // 12pt
          let isBold = b.bold || false;
          let color = '1E293B'; // Slate-800

          if (b.type === 'title') {
            heading = HeadingLevel.TITLE;
            size = 48; // 24pt
            isBold = true;
          } else if (b.type === 'heading1') {
            heading = HeadingLevel.HEADING_1;
            size = 36; // 18pt
            isBold = true;
            color = '4F46E5'; // Indigo-600
          } else if (b.type === 'heading2') {
            heading = HeadingLevel.HEADING_2;
            size = 28; // 14pt
            isBold = true;
            color = '4F46E5';
          }

          let alignment: any = isArabic ? AlignmentType.RIGHT : AlignmentType.LEFT;
          if (b.alignment === 'center') {
            alignment = AlignmentType.CENTER;
          } else if (b.alignment === 'right') {
            alignment = AlignmentType.RIGHT;
          } else if (b.alignment === 'left') {
            alignment = AlignmentType.LEFT;
          }

          const paraOptions: any = {
            children: [
              new TextRun({
                text: b.text || '',
                size,
                bold: isBold,
                color,
              }),
            ],
            alignment,
            spacing: {
              before: b.type === 'title' ? 0 : 180,
              after: 120,
              line: 276, // 1.15 line height
            },
          };

          if (b.type === 'bullet') {
            paraOptions.bullet = { level: 0 };
          }

          if (heading) {
            paraOptions.heading = heading;
          }

          docChildren.push(new Paragraph(paraOptions));
        }
      }

      const doc = new Document({
        sections: [
          {
            properties: {},
            children: docChildren,
          },
        ],
      });

      const docxBuffer = await Packer.toBuffer(doc);

      // Support both JSON and direct binary responses
      const format = req.query.format || req.headers['accept'];
      const wantsBinary = format === 'binary' || (format && (format.toString().includes('octet-stream') || format.toString().includes('wordprocessingml')));

      if (wantsBinary) {
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        const docxFilename = 'converted-document.docx';
        res.setHeader('Content-Disposition', `attachment; filename="${docxFilename}"; filename*=UTF-8''${encodeURIComponent(docxFilename)}`);
        res.send(docxBuffer);
      } else {
        const base64Docx = docxBuffer.toString('base64');
        res.json({
          docxData: base64Docx,
          blocks: blocks,
          blocksCount: blocks.length
        });
      }

    } catch (err: any) {
      console.error('Error in /api/pdf/pdf2word:', err);
      res.status(500).json({ error: err.message || 'Failed to convert PDF to Word' });
    }
  });

  // Helper to wrap text according to available content width in pdf-lib
  function wrapText(text: string, width: number, font: any, fontSize: number): string[] {
    const words = text.split(/\s+/);
    const lines: string[] = [];
    let currentLine = '';

    for (const word of words) {
      const testLine = currentLine ? `${currentLine} ${word}` : word;
      const testWidth = font.widthOfTextAtSize(testLine, fontSize);
      if (testWidth > width) {
        if (currentLine) {
          lines.push(currentLine);
        }
        currentLine = word;
      } else {
        currentLine = testLine;
      }
    }
    if (currentLine) {
      lines.push(currentLine);
    }
    return lines;
  }

  const bidi = bidiFactory();

  // Format Arabic words & letters to visually render RTL correctly in pdf-lib (using ArabicShaper & bidi-js)
  function formatArabicVisualLine(line: string): string {
    if (!/[\u0600-\u06FF]/.test(line)) return line;
    try {
      // Step 1: Reshape the letters so they have the correct contextual shapes (initial, medial, final, isolated)
      const reshaped = ArabicShaper.convertArabic(line);
      // Step 2: Apply the Unicode Bidirectional Algorithm to reorder the characters for visual rendering
      const embeddingLevels = bidi.getEmbeddingLevels(reshaped);
      const visual = bidi.getReorderedString(reshaped, embeddingLevels);
      return visual;
    } catch (err) {
      console.error('Error shaping/bidi Arabic line, falling back:', err);
      const words = line.split(' ');
      const formattedWords = words.map(w => {
        if (/[\u0600-\u06FF]/.test(w)) {
          return w.split('').reverse().join('');
        }
        return w;
      });
      return formattedWords.reverse().join(' ');
    }
  }

  let cachedCairoFont: Buffer | null = null;
  async function getCairoFont(): Promise<Buffer> {
    if (cachedCairoFont) return cachedCairoFont;
    try {
      const res = await fetch('https://raw.githubusercontent.com/google/fonts/main/ofl/cairo/Cairo-Regular.ttf');
      if (!res.ok) throw new Error('Failed to fetch Cairo font');
      const arrayBuffer = await res.arrayBuffer();
      cachedCairoFont = Buffer.from(arrayBuffer);
      return cachedCairoFont;
    } catch (err) {
      console.error('Error fetching Cairo font, falling back to LiberationSans:', err);
      try {
        const fallbackPath = path.join(process.cwd(), 'node_modules', 'pdfjs-dist', 'standard_fonts', 'LiberationSans-Regular.ttf');
        return await fs.readFile(fallbackPath);
      } catch {
        throw err;
      }
    }
  }

  async function getRegularFont(): Promise<Buffer> {
    const localPath = path.join(process.cwd(), 'node_modules', 'pdfjs-dist', 'standard_fonts', 'LiberationSans-Regular.ttf');
    try {
      return await fs.readFile(localPath);
    } catch (err) {
      console.error('LiberationSans not found, fetching Inter regular:', err);
      const res = await fetch('https://raw.githubusercontent.com/google/fonts/main/ofl/inter/Inter-Regular.ttf');
      const arrayBuffer = await res.arrayBuffer();
      return Buffer.from(arrayBuffer);
    }
  }

  async function createPdfFromBlocks(blocks: any[], isArabic: boolean): Promise<Buffer> {
    const pdfDoc = await PDFDocument.create();
    pdfDoc.registerFontkit(fontkit);

    const fontBuffer = isArabic ? await getCairoFont() : await getRegularFont();
    const customFont = await pdfDoc.embedFont(fontBuffer);

    // A4 Dimensions (595.27 x 841.89 points)
    const PAGE_WIDTH = 595.27;
    const PAGE_HEIGHT = 841.89;
    const MARGIN_LEFT = 50;
    const MARGIN_RIGHT = 50;
    const MARGIN_TOP = 50;
    const MARGIN_BOTTOM = 50;
    const CONTENT_WIDTH = PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT;

    let currentPage = pdfDoc.addPage([PAGE_WIDTH, PAGE_HEIGHT]);
    let y = PAGE_HEIGHT - MARGIN_TOP;

    function addNewPage() {
      currentPage = pdfDoc.addPage([PAGE_WIDTH, PAGE_HEIGHT]);
      y = PAGE_HEIGHT - MARGIN_TOP;
    }

    for (const block of blocks) {
      if (!block) continue;

      if (block.type === 'table') {
        const rows = block.rows;
        if (!rows || rows.length === 0) continue;

        const numCols = rows[0].length;
        if (numCols === 0) continue;

        const colWidth = CONTENT_WIDTH / numCols;
        const cellPadding = 6;
        const fontSize = 10;
        const rowSpacing = 4;

        for (const row of rows) {
          let maxLines = 1;
          for (const cellText of row) {
            const textStr = String(cellText || '');
            const wrapped = wrapText(textStr, colWidth - cellPadding * 2, customFont, fontSize);
            if (wrapped.length > maxLines) {
              maxLines = wrapped.length;
            }
          }

          const rowHeight = maxLines * (fontSize + 4) + cellPadding * 2;

          if (y - rowHeight < MARGIN_BOTTOM) {
            addNewPage();
          }

          let x = MARGIN_LEFT;
          for (let c = 0; c < row.length; c++) {
            const cellText = String(row[c] || '');

            currentPage.drawRectangle({
              x: x,
              y: y - rowHeight,
              width: colWidth,
              height: rowHeight,
              color: rgb(0.97, 0.98, 0.99),
              borderColor: rgb(0.88, 0.90, 0.93),
              borderWidth: 1,
            });

            const wrapped = wrapText(cellText, colWidth - cellPadding * 2, customFont, fontSize);
            let cellY = y - cellPadding - fontSize;
            for (const line of wrapped) {
              const visualLine = formatArabicVisualLine(line);
              let textX = x + cellPadding;
              const cellHasArabic = /[\u0600-\u06FF]/.test(line);
              if (cellHasArabic) {
                textX = x + colWidth - cellPadding - customFont.widthOfTextAtSize(visualLine, fontSize);
              }
              currentPage.drawText(visualLine, {
                x: textX,
                y: cellY,
                size: fontSize,
                font: customFont,
                color: rgb(0.12, 0.16, 0.23),
              });
              cellY -= (fontSize + 4);
            }

            x += colWidth;
          }

          y -= rowHeight;
          y -= rowSpacing;
        }

        y -= 10;
      } else {
        const blockText = String(block.text || '');
        if (!blockText.trim()) continue;

        let fontSize = 11;
        let isBold = block.bold || false;
        let color = rgb(0.12, 0.16, 0.23);
        let spacingBefore = 6;
        let spacingAfter = 6;
        let lineSpacing = 4;

        if (block.type === 'title') {
          fontSize = 24;
          isBold = true;
          color = rgb(0.06, 0.09, 0.16);
          spacingBefore = 15;
          spacingAfter = 15;
        } else if (block.type === 'heading1') {
          fontSize = 18;
          isBold = true;
          color = rgb(0.31, 0.27, 0.90);
          spacingBefore = 12;
          spacingAfter = 10;
        } else if (block.type === 'heading2') {
          fontSize = 14;
          isBold = true;
          color = rgb(0.31, 0.27, 0.90);
          spacingBefore = 10;
          spacingAfter = 8;
        } else if (block.type === 'bullet' || block.type === 'numbered') {
          fontSize = 11;
          spacingBefore = 4;
          spacingAfter = 4;
        }

        y -= spacingBefore;

        let textToRender = blockText;
        if (block.type === 'bullet') {
          textToRender = `•  ${blockText}`;
        } else if (block.type === 'numbered') {
          textToRender = `1.  ${blockText}`;
        }

        const wrappedLines = wrapText(textToRender, CONTENT_WIDTH, customFont, fontSize);

        for (const line of wrappedLines) {
          if (y - fontSize - lineSpacing < MARGIN_BOTTOM) {
            addNewPage();
          }

          const visualLine = formatArabicVisualLine(line);
          let x = MARGIN_LEFT;

          const lineHasArabic = /[\u0600-\u06FF]/.test(line);
          if (block.alignment === 'center') {
            const textW = customFont.widthOfTextAtSize(visualLine, fontSize);
            x = MARGIN_LEFT + (CONTENT_WIDTH - textW) / 2;
          } else if (block.alignment === 'right' || (lineHasArabic && block.alignment !== 'left' && block.alignment !== 'center')) {
            const textW = customFont.widthOfTextAtSize(visualLine, fontSize);
            x = PAGE_WIDTH - MARGIN_RIGHT - textW;
          }

          currentPage.drawText(visualLine, {
            x,
            y: y - fontSize,
            size: fontSize,
            font: customFont,
            color,
          });

          y -= (fontSize + lineSpacing);
        }

        y -= spacingAfter;
      }
    }

    const pdfBytes = await pdfDoc.save();
    return Buffer.from(pdfBytes);
  }

  // Modern Convert to PDF Endpoint
  app.post('/api/pdf/convert2pdf', async (req, res) => {
    try {
      const { fileData, fileName, mimeType, url, toolId } = req.body;
      if (!toolId) {
        return res.status(400).json({ error: 'Missing toolId' });
      }

      let contentToProcess = '';
      let targetMimeType = mimeType || 'text/plain';

      if (url) {
        // HTML via URL Conversion
        try {
          const fetchRes = await fetch(url);
          if (!fetchRes.ok) {
            return res.status(400).json({ error: `Failed to fetch URL content: ${fetchRes.statusText}` });
          }
          contentToProcess = await fetchRes.text();
          targetMimeType = 'text/html';
        } catch (fetchErr: any) {
          return res.status(400).json({ error: `Could not reach URL: ${fetchErr.message}` });
        }
      } else if (fileData) {
        contentToProcess = fileData;
      } else {
        return res.status(400).json({ error: 'Please upload a file or enter a valid URL' });
      }

      const apiKey = process.env.GEMINI_API_KEY;
      if (!apiKey) {
        return res.status(500).json({
          error: 'GEMINI_API_KEY is not configured. Please add your key in Settings > Secrets.'
        });
      }

      const ai = new GoogleGenAI({
        apiKey,
        httpOptions: {
          headers: {
            'User-Agent': 'aistudio-build',
          }
        }
      });

      let systemPrompt = `You are a high-fidelity document parsing model. Analyze the provided file/content and reconstruct its exact content, tables, lists, and structure as a valid JSON array of block objects.
Each block object must strictly have this structure:
- "type": "title" | "heading1" | "heading2" | "paragraph" | "bullet" | "numbered" | "table"
- "text": string (only for title, heading1, heading2, paragraph, bullet, numbered)
- "rows": array of arrays of strings (only for "table")
- "alignment": "left" | "center" | "right" (optional)
- "bold": boolean (optional)

Ensure the text is extracted exactly as it appears in the source, retaining headers, bullet lists, numeric items, tabular figures, and structure.
Return ONLY the raw JSON array. Do not wrap it in markdown or backticks. If you must use formatting, do not use backticks outside of the JSON. Make sure the JSON is fully valid and parseable.`;

      const parts: any[] = [];
      if (url) {
        parts.push({
          text: `Here is the HTML source code of the webpage to convert:\n\n${contentToProcess}`
        });
      } else {
        parts.push({
          inlineData: {
            data: contentToProcess,
            mimeType: targetMimeType
          }
        });
      }
      parts.push({ text: systemPrompt });

      const response = await ai.models.generateContent({
        model: 'gemini-3.5-flash',
        contents: { parts: parts }
      });

      const rawResult = response.text || '[]';
      let cleaned = rawResult.trim();
      if (cleaned.startsWith('```')) {
        cleaned = cleaned.replace(/^```json\s*/i, '');
        cleaned = cleaned.replace(/```\s*$/, '');
        cleaned = cleaned.trim();
      }

      let blocks: any[] = [];
      try {
        blocks = JSON.parse(cleaned);
      } catch (parseErr) {
        console.error('Failed to parse blocks JSON from Gemini:', rawResult);
        blocks = [{ type: 'paragraph', text: rawResult }];
      }

      // Check if text is Arabic for Cairo Font loading
      const allText = JSON.stringify(blocks);
      const isArabic = /[\u0600-\u06FF]/.test(allText);

      // Generate pristine PDF
      const pdfBuffer = await createPdfFromBlocks(blocks, isArabic);

      // Set explicit professional response headers
      res.setHeader('Content-Type', 'application/pdf');
      const targetFileName = fileName 
        ? `${fileName.substring(0, fileName.lastIndexOf('.')) || fileName}.pdf`
        : 'converted.pdf';
      const safeFileName = targetFileName.replace(/["\\]/g, '');
      res.setHeader('Content-Disposition', `attachment; filename="${safeFileName}"; filename*=UTF-8''${encodeURIComponent(targetFileName)}`);
      res.send(pdfBuffer);

    } catch (err: any) {
      console.error('Error in convert2pdf API:', err);
      res.status(500).json({ error: err.message || 'Failed to convert document to PDF' });
    }
  });

  // Tool Names and translations for server-side SEO tag generation
  const toolNames: Record<string, Record<string, string>> = {
    merge: {
      en: 'Merge PDF files',
      ar: 'دمج ملفات PDF',
      es: 'Unir archivos PDF',
      fr: 'Fusionner des fichiers PDF'
    },
    split: {
      en: 'Split PDF documents',
      ar: 'تقسيم مستندات PDF',
      es: 'Dividir documentos PDF',
      fr: 'Diviser documents PDF'
    },
    compress: {
      en: 'Compress PDF size',
      ar: 'ضغط حجم PDF',
      es: 'Comprimir tamaño de PDF',
      fr: 'Compresser la taille du PDF'
    },
    pdf2word: {
      en: 'PDF to Word converter',
      ar: 'تحويل PDF إلى Word',
      es: 'Convertidor de PDF a Word',
      fr: 'Convertisseur PDF en Word'
    },
    pdf2excel: {
      en: 'PDF to Excel converter',
      ar: 'تحويل PDF إلى Excel',
      es: 'Convertidor de PDF a Excel',
      fr: 'Convertisseur PDF en Excel'
    },
    pdf2ppt: {
      en: 'PDF to PowerPoint converter',
      ar: 'تحويل PDF إلى PowerPoint',
      es: 'Convertidor de PDF a PowerPoint',
      fr: 'Convertisseur PDF en PowerPoint'
    },
    img2pdf: {
      en: 'Image to PDF converter',
      ar: 'تحويل الصور إلى PDF',
      es: 'Convertidor de imagen a PDF',
      fr: 'Convertisseur Image en PDF'
    },
    watermark: {
      en: 'Add watermarks to PDF',
      ar: 'إضافة علامات مائية',
      es: 'Añadir marcas de agua',
      fr: 'Ajouter des filigranes'
    },
    protect: {
      en: 'PDF protection (Password Protect)',
      ar: 'حماية وتشفير PDF',
      es: 'Protección de PDF con contraseña',
      fr: 'Protection PDF par mot de passe'
    },
    rotate: {
      en: 'Rotate PDF pages',
      ar: 'تدوير صفحات PDF',
      es: 'Rotar páginas de PDF',
      fr: 'Faire pivoter les pages PDF'
    },
    delete_pages: {
      en: 'Delete PDF pages',
      ar: 'حذف صفحات من PDF',
      es: 'Eliminar páginas de PDF',
      fr: 'Supprimer des pages PDF'
    },
    reorder: {
      en: 'Reorder PDF pages',
      ar: 'إعادة ترتيب صفحات PDF',
      es: 'Reordenar páginas de PDF',
      fr: 'Réorganiser les pages PDF'
    },
    unlock: {
      en: 'Unlock PDF (Remove password)',
      ar: 'فك حماية PDF (إزالة كلمة السر)',
      es: 'Desbloquear PDF (Eliminar contraseña)',
      fr: 'Déverrouiller PDF (Supprimer le mot de passe)'
    },
    pdf2img: {
      en: 'PDF to JPG converter',
      ar: 'تحويل PDF إلى صور JPG',
      es: 'Convertidor de PDF a JPG',
      fr: 'Convertisseur PDF en JPG'
    },
    crop: {
      en: 'Crop PDF pages',
      ar: 'قص واقتصاص صفحات PDF',
      es: 'Cortar páginas de PDF',
      fr: 'Rogner les pages PDF'
    },
    extract: {
      en: 'Smart text extraction (AI)',
      ar: 'استخراج نصوص ذكي بالذكاء الاصطناعي',
      es: 'Extracción inteligente de texto con IA',
      fr: 'Extraction intelligente de texte par IA'
    },
    summarize: {
      en: 'Automatic PDF summarization (AI)',
      ar: 'تلخيص PDF تلقائي بالذكاء الاصطناعي',
      es: 'Resumen automático de PDF con IA',
      fr: 'Résumé automatique de PDF par IA'
    },
    translate: {
      en: 'PDF translation (AI)',
      ar: 'ترجمة ملفات PDF بالذكاء الاصطناعي',
      es: 'Traducción de PDF con IA',
      fr: 'Traduction de PDF par IA'
    },
    ocr: {
      en: 'OCR for scanned documents (AI)',
      ar: 'التعرف الضوئي على الحروف للمستندات بالذكاء الاصطناعي',
      es: 'OCR para documentos escaneados con IA',
      fr: 'OCR pour documents numérisés par IA'
    },
    format: {
      en: 'Intelligent PDF formatting (AI)',
      ar: 'تنسيق ذكي بالذكاء الاصطناعي',
      es: 'Formateo inteligente de PDF con IA',
      fr: 'Formatage intelligent de PDF par IA'
    },
    chat: {
      en: 'Chat with PDF (AI Ask & Answer)',
      ar: 'محادثة مع ملف PDF بالذكاء الاصطناعي',
      es: 'Chatear con PDF con IA',
      fr: 'Discuter avec un PDF par IA'
    },
    redact: {
      en: 'Auto-redact sensitive info (AI)',
      ar: 'إخفاء وتنقية المعلومات الحساسة بالذكاء الاصطناعي',
      es: 'Ocultar información sensible con IA',
      fr: 'Masquage des informations sensibles par IA'
    },
    compare: {
      en: 'Compare PDF versions with AI',
      ar: 'مقارنة نسختين من PDF بالذكاء الاصطناعي',
      es: 'Comparar versiones de PDF con IA',
      fr: 'Comparer des versions PDF par IA'
    }
  };

  function getSEOContent(lang: string, toolId?: string) {
    const isAr = lang === 'ar';
    const isEs = lang === 'es';
    const isFr = lang === 'fr';

    // Base configurations
    let siteName = 'PDFProTools';
    let defaultTitle = 'PDFProTools - Free Online PDF Tools Powered by AI';
    let defaultDesc = 'Free, high-quality, and professional online PDF tools: merge, split, compress, convert PDF to Word, and leverage advanced AI tools for summarizing, translating, and chatting with documents.';

    if (isAr) {
      defaultTitle = 'PDFProTools - أدوات PDF ذكية ومحترفة مجاناً عبر الإنترنت';
      defaultDesc = 'أدوات PDF مجانية ومحترفة أونلاين: دمج، تقسيم، ضغط، تحويل PDF إلى Word ومستندات أخرى، بالإضافة لأدوات ذكية مدعومة بالذكاء الاصطناعي كالتلخيص والترجمة والدردشة مع المستندات.';
    } else if (isEs) {
      defaultTitle = 'PDFProTools - Herramientas PDF en línea gratuitas e inteligentes';
      defaultDesc = 'Herramientas PDF en línea gratuitas e inteligentes: unir, dividir, comprimir, convertir PDF a Word y herramientas avanzadas de IA para resumir, traducir y chatear con documentos.';
    } else if (isFr) {
      defaultTitle = 'PDFProTools - Outils PDF en ligne gratuits et intelligents';
      defaultDesc = 'Outils PDF en ligne gratuits et intelligents : fusionner, diviser, compresser, convertir PDF en Word et outils IA avancés pour résumer, traduire et discuter avec des documents.';
    }

    if (!toolId) {
      return { title: defaultTitle, description: defaultDesc };
    }

    // If a tool is requested
    const toolNameMap = toolNames[toolId];
    if (!toolNameMap) {
      return { title: defaultTitle, description: defaultDesc };
    }

    const toolName = toolNameMap[lang] || toolNameMap['en'];

    let title = '';
    let description = '';

    if (isAr) {
      title = `${toolName} أونلاين مجاناً - ${siteName}`;
      description = `استخدم أداة ${toolName} أونلاين مجاناً وبشكل آمن تماماً وبسرعة فائقة. أدوات PDF احترافية خالية من العلامات المائية ولا تتطلب التسجيل.`;
    } else if (isEs) {
      title = `${toolName} gratis en línea - ${siteName}`;
      description = `Utilice la herramienta ${toolName} gratis en línea de forma segura y ultra rápida. Herramientas PDF profesionales sin marcas de agua y sin necesidad de registro.`;
    } else if (isFr) {
      title = `${toolName} gratuit en ligne - ${siteName}`;
      description = `Utilisez l'outil ${toolName} gratuit en ligne de manière sécurisée et ultra rapide. Outils PDF professionnels sans filigrane et sans inscription requise.`;
    } else {
      title = `${toolName} Online Free - ${siteName}`;
      description = `Use ${toolName} online for free, securely, and ultra-fast. Professional PDF tools with no watermarks and no registration required.`;
    }

    return { title, description };
  }

  // --- Reviews & Ratings Persistent Store ---
  interface Review {
    id: string;
    stars: number;
    comment: string;
    date: string;
  }

  const REVIEWS_FILE = path.join(process.cwd(), 'reviews.json');

  async function loadReviews(): Promise<Review[]> {
    const seedReviews: Review[] = [
      { id: 'seed-1', stars: 5, comment: 'أفضل موقع لتعديل ملفات PDF مجاناً وبدون قيود! وسرعة التحويل ممتازة.', date: '2026-07-18T12:00:00.000Z' },
      { id: 'seed-2', stars: 5, comment: 'Really handy tool. Reordering and merging pages took me less than 10 seconds. Highly recommend!', date: '2026-07-17T15:30:00.000Z' },
      { id: 'seed-3', stars: 4, comment: 'La mejor herramienta para comprimir mis archivos PDF. Muy fácil de usar.', date: '2026-07-16T09:15:00.000Z' },
      { id: 'seed-4', stars: 5, comment: "Excellent service ! J'ai pu numéroter mes documents PDF très rapidement.", date: '2026-07-15T18:45:00.000Z' },
      { id: 'seed-5', stars: 5, comment: 'مفيد جداً خصوصاً دمج الملفات وأدوات الذكاء الاصطناعي لتلخيص المستندات العربية.', date: '2026-07-15T11:20:00.000Z' }
    ];

    try {
      const data = await fs.readFile(REVIEWS_FILE, 'utf-8');
      return JSON.parse(data);
    } catch (error) {
      try {
        await fs.writeFile(REVIEWS_FILE, JSON.stringify(seedReviews, null, 2), 'utf-8');
      } catch (writeErr) {
        console.error('Error writing seed reviews:', writeErr);
      }
      return seedReviews;
    }
  }

  async function saveReviews(reviews: Review[]): Promise<void> {
    await fs.writeFile(REVIEWS_FILE, JSON.stringify(reviews, null, 2), 'utf-8');
  }

  app.get('/api/reviews', async (req, res) => {
    try {
      const allReviews = await loadReviews();
      const totalReviews = allReviews.length;
      const sum = allReviews.reduce((acc, r) => acc + r.stars, 0);
      const averageRating = totalReviews > 0 ? parseFloat((sum / totalReviews).toFixed(1)) : 0;
      const latestReviews = [...allReviews].sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime()).slice(0, 10);
      
      res.json({
        averageRating,
        totalReviews,
        reviews: latestReviews
      });
    } catch (e: any) {
      console.error('Error in GET /api/reviews:', e);
      res.status(500).json({ error: 'Failed to load reviews' });
    }
  });

  app.post('/api/reviews', async (req, res) => {
    try {
      const { stars, comment } = req.body;
      const rating = parseInt(stars, 10);
      if (isNaN(rating) || rating < 1 || rating > 5) {
        return res.status(400).json({ error: 'Invalid rating. Stars must be 1 to 5.' });
      }

      const cleanComment = (comment || '').trim().substring(0, 500);
      const newReview: Review = {
        id: crypto.randomUUID(),
        stars: rating,
        comment: cleanComment,
        date: new Date().toISOString()
      };

      const allReviews = await loadReviews();
      allReviews.push(newReview);
      await saveReviews(allReviews);

      const totalReviews = allReviews.length;
      const sum = allReviews.reduce((acc, r) => acc + r.stars, 0);
      const averageRating = totalReviews > 0 ? parseFloat((sum / totalReviews).toFixed(1)) : 0;
      const latestReviews = [...allReviews].sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime()).slice(0, 10);

      res.status(201).json({
        success: true,
        averageRating,
        totalReviews,
        reviews: latestReviews
      });
    } catch (e: any) {
      console.error('Error in POST /api/reviews:', e);
      res.status(500).json({ error: 'Failed to submit review' });
    }
  });



  // SEO page handler for navigation requests (Production only)
  if (process.env.NODE_ENV === 'production') {
    app.get(['/', '/:lang', '/:lang/:tool', '/:tool'], async (req, res, next) => {
      const url = req.originalUrl || req.url;
      // Skip static assets, vite HMR, or API routes
      if (url.startsWith('/api') || url.includes('.') || req.xhr) {
        return next();
      }

      const langParam = req.params.lang;
      const toolParam = req.params.tool;
      const lang = typeof langParam === 'string' ? langParam : '';
      const tool = typeof toolParam === 'string' ? toolParam : '';

      let targetLang = 'en';
      let targetToolId = '';

      if (lang) {
        if (['en', 'ar', 'es', 'fr'].includes(lang.toLowerCase())) {
          targetLang = lang.toLowerCase();
          if (tool) {
            targetToolId = tool;
          }
        } else {
          // First parameter is actually a tool ID (e.g., /merge)
          targetToolId = lang;
        }
      }

      try {
        const templatePath = path.join(process.cwd(), 'dist', 'index.html');
        let template = '';
        try {
          template = await fs.readFile(templatePath, 'utf-8');
        } catch (err) {
          // Fallback to reading root if dist/index.html is not compiled yet
          const fallbackPath = path.join(process.cwd(), 'index.html');
          template = await fs.readFile(fallbackPath, 'utf-8');
        }

        const seoData = getSEOContent(targetLang, targetToolId);
        const hreflangTags = [
          `    <link rel="alternate" hreflang="x-default" href="https://pdfprotools.com/" />`,
          `    <link rel="alternate" hreflang="en" href="https://pdfprotools.com/en${targetToolId ? '/' + targetToolId : ''}" />`,
          `    <link rel="alternate" hreflang="ar" href="https://pdfprotools.com/ar${targetToolId ? '/' + targetToolId : ''}" />`,
          `    <link rel="alternate" hreflang="es" href="https://pdfprotools.com/es${targetToolId ? '/' + targetToolId : ''}" />`,
          `    <link rel="alternate" hreflang="fr" href="https://pdfprotools.com/fr${targetToolId ? '/' + targetToolId : ''}" />`
        ].join('\n');

        const ogUrl = `https://pdfprotools.com/${targetLang}${targetToolId ? '/' + targetToolId : ''}`;

        const seoHeadContent = `
    <title>${seoData.title}</title>
    <meta name="description" content="${seoData.description}" />
    <meta name="robots" content="index, follow" />
    
    <!-- Open Graph / Facebook -->
    <meta property="og:type" content="website" />
    <meta property="og:url" content="${ogUrl}" />
    <meta property="og:title" content="${seoData.title}" />
    <meta property="og:description" content="${seoData.description}" />
    <meta property="og:image" content="https://pdfprotools.com/favicon.png" />
    
    <!-- Twitter -->
    <meta property="twitter:card" content="summary" />
    <meta property="twitter:url" content="${ogUrl}" />
    <meta property="twitter:title" content="${seoData.title}" />
    <meta property="twitter:description" content="${seoData.description}" />
    <meta property="twitter:image" content="https://pdfprotools.com/favicon.png" />
    
    <!-- Multi-language Hreflang Tags -->
${hreflangTags}
        `;

        let modifiedHtml = template;
        // Remove default <title> tag if present
        modifiedHtml = modifiedHtml.replace(/<title>.*?<\/title>/, '');
        // Insert customized head tags right after <head>
        modifiedHtml = modifiedHtml.replace('<head>', `<head>\n${seoHeadContent}`);

        res.status(200).set({ 'Content-Type': 'text/html' }).end(modifiedHtml);
      } catch (e: any) {
        console.error('Error in SEO Page handler:', e);
        res.status(500).end(e.stack || 'Internal Server Error');
      }
    });
  }

  // Static assets serving in production
  if (process.env.NODE_ENV === "production") {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*splat', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  if (!process.env.VERCEL) {
    app.listen(PORT, "0.0.0.0", () => {
      console.log(`Server running on http://0.0.0.0:${PORT}`);
    });
  }
}

startServer();
