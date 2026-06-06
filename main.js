import { Actor } from 'apify';
import axios from 'axios';
import crypto from 'crypto';
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx';
import PDFDocument from 'pdfkit';

await Actor.init();

// ========== HELPERS ==========

function calculateHash(originalFeedback, improvedFeedback, timestamp) {
  const data = `${originalFeedback}|${improvedFeedback}|${timestamp}`;
  return crypto.createHash('sha256').update(data).digest('hex');
}

async function generateDOCX(recipientName, improvedFeedback, auditHash) {
  const doc = new Document({
    sections: [{
      properties: {},
      children: [
        new Paragraph({
          text: `For: ${recipientName || 'Recipient'}`,
          heading: HeadingLevel.HEADING_2,
          spacing: { after: 120 },
        }),
        new Paragraph({
          text: improvedFeedback,
          spacing: { after: 300 },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: `Verification Code: ${auditHash}`,
              italics: true,
              size: 18,
              color: '888888',
            }),
          ],
        }),
      ],
    }],
  });
  return await Packer.toBuffer(doc);
}

async function generatePDF(recipientName, improvedFeedback, auditHash) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ margin: 50 });
    const chunks = [];
    doc.on('data', chunk => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    doc.fontSize(14).text(`SCDFT – Stech Clarity-Driven Feedback Transformer`, { align: 'center' });
    doc.moveDown(0.5);
    doc.fontSize(11).text(`Addressed to: ${recipientName || 'Recipient'}`);
    doc.fontSize(11).text(`Date: ${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}`);
    doc.moveDown(0.5);
    doc.fontSize(11).text(improvedFeedback, { align: 'justify', lineGap: 4 });
    doc.moveDown(1);
    doc.fontSize(7).text(`Verification Code: ${auditHash}`, { align: 'center' });
    doc.end();
  });
}

async function saveFileToKVS(filename, buffer, contentType) {
  const store = await Actor.openKeyValueStore();
  await store.setValue(filename, buffer, { contentType });
  const runId = process.env.APIFY_RUN_ID || 'default';
  const baseUrl = `https://api.apify.com/v2/key-value-stores/${store.id}/records/${filename}?disableRedirect=true`;
  return baseUrl;
}

function removeSubjectFromBody(body, subject) {
  if (!subject || !body) return body;
  const escapedSubject = subject.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const patterns = [
    new RegExp(`\\*\\*${escapedSubject}\\*\\*\\s*\\n*`, 'gi'),
    new RegExp(`#\\s*${escapedSubject}\\s*\\n*`, 'gi'),
    new RegExp(`^${escapedSubject}\\s*\\n*`, 'gim'),
    new RegExp(`^Subject\\s*:\\s*${escapedSubject}\\s*\\n*`, 'gim'),
  ];
  for (const pattern of patterns) {
    body = body.replace(pattern, '').trim();
  }
  return body;
}

// ========== INPUT CONFIGURATION ==========

const input = await Actor.getInput();
const {
  csvFile,
  feedbacks,
  columnMapping,
  rejectionTemplate,
  maxConcurrency = 5,
  timeout = 60,
} = input;

const SCDFT_PROXY_SECRET = process.env.SCDFT_PROXY_SECRET;
if (!SCDFT_PROXY_SECRET) {
  throw new Error('SCDFT_PROXY_SECRET environment variable is missing');
}

const API_URL = 'https://stech-api.sheradogilang.workers.dev/scdft';

// ========== CSV PARSER ==========

function parseCSV(content) {
  const lines = content.trim().split(/\r?\n/);
  if (lines.length === 0) return [];
  const headers = [];
  let inQuote = false;
  let current = '';
  const firstLine = lines[0];
  for (let ch of firstLine) {
    if (ch === '"') inQuote = !inQuote;
    else if (ch === ',' && !inQuote) {
      headers.push(current.trim().replace(/^"|"$/g, ''));
      current = '';
    } else {
      current += ch;
    }
  }
  headers.push(current.trim().replace(/^"|"$/g, ''));

  const result = [];
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i];
    if (!line.trim()) continue;
    const values = [];
    inQuote = false;
    current = '';
    for (let ch of line) {
      if (ch === '"') inQuote = !inQuote;
      else if (ch === ',' && !inQuote) {
        values.push(current.trim().replace(/^"|"$/g, ''));
        current = '';
      } else {
        current += ch;
      }
    }
    values.push(current.trim().replace(/^"|"$/g, ''));
    const row = {};
    headers.forEach((h, idx) => {
      row[h] = values[idx] || '';
    });
    result.push(row);
  }
  return result;
}

function applyMappingAndTemplate(row, mapping, template) {
  if (!template || !mapping) return null;
  let filled = template;
  for (const [placeholder, columnName] of Object.entries(mapping)) {
    const value = row[columnName] || '';
    filled = filled.replace(new RegExp(`{${placeholder}}`, 'g'), value);
  }
  return filled;
}

// ========== BUILD FEEDBACK LIST ==========

let feedbackList = [];

if (csvFile && typeof csvFile === 'string') {
  let fileContent = null;
  if (csvFile.startsWith('FILE-UPLOAD:')) {
    const fileKey = csvFile.replace('FILE-UPLOAD:', '');
    const fileBuffer = await Actor.getFile(fileKey);
    fileContent = fileBuffer.toString();
  } else if (csvFile.startsWith('http://') || csvFile.startsWith('https://')) {
    const response = await axios.get(csvFile, { responseType: 'text' });
    fileContent = response.data;
  } else {
    throw new Error('Invalid csvFile format. Must be a FILE-UPLOAD: key or a public URL.');
  }
  const rows = parseCSV(fileContent);
  if (rows.length === 0) {
    throw new Error('CSV file is empty or could not be parsed.');
  }

  if (columnMapping && rejectionTemplate && Object.keys(columnMapping).length > 0) {
    feedbackList = rows.map(row => ({
      originalFeedback: applyMappingAndTemplate(row, columnMapping, rejectionTemplate),
      additionalInstructions: row.additionalInstructions || '',
      originalSubject: row.originalSubject || row.subject || '',
      recipientName: row.recipientName || row.recipient_name || row.recipient || row.name || '',
      senderName: row.senderName || row.sender_name || row.sender || '',
      recipientEmail: row.recipientEmail || row.recipient_email || row.email || '',
    })).filter(item => item.originalFeedback);
  } else {
    feedbackList = rows.filter(row => row.originalFeedback).map(row => ({
      originalFeedback: row.originalFeedback,
      additionalInstructions: row.additionalInstructions || '',
      originalSubject: row.originalSubject || row.subject || '',
      recipientName: row.recipientName || row.recipient_name || row.recipient || row.name || '',
      senderName: row.senderName || row.sender_name || row.sender || '',
      recipientEmail: row.recipientEmail || row.recipient_email || row.email || '',
    }));
  }
} else if (Array.isArray(feedbacks) && feedbacks.length > 0) {
  feedbackList = feedbacks;
} else {
  throw new Error('No input provided. Please either upload a CSV file or provide an array of feedbacks.');
}

if (feedbackList.length === 0) {
  throw new Error('No valid feedback entries found. Check your input data.');
}

// ========== PROCESS EACH FEEDBACK ==========

async function processFeedback(item, index) {
  const originalFeedback = item.originalFeedback;
  if (!originalFeedback) {
    return {
      originalFeedback: null,
      improvedFeedback: "",
      status: 'error',
      error: 'Missing originalFeedback field',
      timestamp: new Date().toISOString(),
      auditHash: '',
    };
  }

  const additional = item.additionalInstructions || '';
  const originalSubject = item.originalSubject || '';
  const recipientName = item.recipientName || '';
  const senderName = item.senderName || '';
  const recipientEmail = item.recipientEmail || '';

 let personalization = '';
if (recipientName) {
  personalization += ` The feedback concerns "${recipientName}".`;
}
if (senderName) {
  personalization += ` The feedback is from "${senderName}".`;
}

  let prompt = `Articulate the following customer feedback. 
CRITICAL RULES:
1. Write strictly from the perspective of the feedback giver (use "I").
2. Do not turn it into a company reply.
3. DO NOT include any conversational filler, introductions, or explanations (e.g., do not say "Here is the revised feedback" or "Here's my take"). Output ONLY the articulated feedback.${personalization}`;

if (additional) prompt += `\nAdditional instructions: ${additional}`;
if (originalSubject) {
  prompt += `\nThe feedback subject is "${originalSubject}". Keep the subject unchanged.`;
}
prompt += `\n\nOriginal feedback:\n${originalFeedback}`;

  try {
    const response = await axios.post(API_URL, { message: prompt }, {
      headers: { 'X-Stech-Actor-Secret': SCDFT_PROXY_SECRET },
      timeout: timeout * 1000,
    });
    let improvedFeedback = response.data.response?.trim() || '';
    if (originalSubject) {
      improvedFeedback = removeSubjectFromBody(improvedFeedback, originalSubject);
    }

    const timestamp = new Date().toISOString();
    const auditHash = calculateHash(originalFeedback, improvedFeedback, timestamp);

    // ---- MULTI-FORMAT FILE GENERATION ----
    const docxBuffer = await generateDOCX(recipientName, improvedFeedback, auditHash);
    const pdfBuffer = await generatePDF(recipientName, improvedFeedback, auditHash);

    const docxFilename = `feedback_${index + 1}_${auditHash.substring(0, 8)}.docx`;
    const pdfFilename = `feedback_${index + 1}_${auditHash.substring(0, 8)}.pdf`;

    const docxUrl = await saveFileToKVS(docxFilename, docxBuffer, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    const pdfUrl = await saveFileToKVS(pdfFilename, pdfBuffer, 'application/pdf');

    return {
      originalFeedback,
      improvedFeedback,
      status: 'success',
      timestamp,
      auditHash,
      download_docx: docxUrl,
      download_pdf: pdfUrl,
      ...(originalSubject && { originalSubject }),
      ...(recipientName && { recipientName }),
      ...(senderName && { senderName }),
      ...(recipientEmail && { recipientEmail }),
    };
  } catch (err) {
    return {
      originalFeedback: originalFeedback || "",
      improvedFeedback: "",
      status: 'error',
      error: err.message,
      timestamp: new Date().toISOString(),
      auditHash: '',
      ...(originalSubject && { originalSubject }),
      ...(recipientName && { recipientName }),
      ...(senderName && { senderName }),
      ...(recipientEmail && { recipientEmail }),
    };
  }
}

// ========== PARALLEL EXECUTION ==========

const results = [];
const running = new Set();
const queue = [...feedbackList];
let nextIndex = 0;

while (queue.length > 0 || running.size > 0) {
  while (running.size < maxConcurrency && queue.length > 0) {
    const item = queue.shift();
    const index = nextIndex++;
    const promise = processFeedback(item, index).then(res => {
      running.delete(promise);
      results.push(res);
    });
    running.add(promise);
  }
  if (running.size > 0) {
    await Promise.race(running);
  }
}

results.sort((a, b) => a.index - b.index);
const finalOutput = results.map(({ index, ...rest }) => rest);

await Actor.pushData(finalOutput);
console.log(`Processed ${finalOutput.length} feedbacks. Success: ${finalOutput.filter(r => r.status === 'success').length}, Errors: ${finalOutput.filter(r => r.status === 'error').length}`);

await Actor.exit();
