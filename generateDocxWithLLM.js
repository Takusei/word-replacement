import "dotenv/config";
import fs from "fs";
import PizZip from "pizzip";
import OpenAI from "openai";

/* ======================================================
 * MAIN ENTRY
 * ====================================================== */

/**
 * Generate a Word document by letting LLM fill ALL placeholders.
 * LLM considers:
 *   - placeholder name
 *   - sentence context
 *   - provided key-value data
 *
 * @param {string} filePath - Path to Word template (.docx)
 * @param {Record<string, string>} valueMap - Extracted data
 * @param {"openai" | "azure"} provider - LLM provider
 * @returns {Promise<Buffer>} - Generated docx buffer
 */
export async function generateDocxWithLLM(
  filePath,
  valueMap,
  provider = "openai"
) {
  // 1. Load Word file
  const content = fs.readFileSync(filePath, "binary");
  const zip = new PizZip(content);

  const xmlPath = "word/document.xml";
  let xml = zip.file(xmlPath).asText();

  // 2. Normalize placeholders → [PH_1], [PH_2], ...
  const PLACEHOLDER_REGEX = /\[([^\[\]]+)\]/g;
  let counter = 1;
  const placeholders = [];

  xml = xml.replace(PLACEHOLDER_REGEX, (_, raw) => {
    const id = `PH_${counter++}`;
    placeholders.push({ id, raw });
    return `[${id}]`;
  });

  // 3. Extract plain text for context
  const plainText = stripXml(xml);

  // 4. Resolve each placeholder via LLM
  const resolved = {};

  for (const ph of placeholders) {
    const marker = `[${ph.id}]`;
    const sentence = extractSentence(plainText, marker);

    const prompt = buildPrompt({
      placeholderName: ph.raw,
      marker,
      sentence,
      valueMap,
    });

    const value = await callLLM(prompt, provider);
    resolved[ph.id] = value;
  }

  // 5. Replace placeholders with LLM output
  for (const [id, value] of Object.entries(resolved)) {
    xml = xml.replace(`[${id}]`, value);
  }

  // 6. Write back XML and generate final docx
  zip.file(xmlPath, xml);

  return zip.generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
}

/* ======================================================
 * LLM CALLS
 * ====================================================== */

async function callLLM(prompt, provider) {
  if (provider === "azure") {
    return callAzureOpenAI(prompt);
  }
  return callOpenAI(prompt);
}

/* ---------- OpenAI ---------- */

async function callOpenAI(prompt) {
  const client = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY,
  });

  const res = await client.chat.completions.create({
    model: "gpt-4.1-mini",
    messages: [{ role: "user", content: prompt }],
    temperature: 0,
  });

  return res.choices[0].message.content.trim();
}

/* ---------- Azure OpenAI ---------- */

async function callAzureOpenAI(prompt) {
  const client = new OpenAI({
    apiKey: process.env.AZURE_OPENAI_API_KEY,
    baseURL: `${process.env.AZURE_OPENAI_ENDPOINT}/openai/deployments/${process.env.AZURE_OPENAI_DEPLOYMENT}`,
    defaultQuery: {
      "api-version": process.env.AZURE_OPENAI_API_VERSION,
    },
    defaultHeaders: {
      "api-key": process.env.AZURE_OPENAI_API_KEY,
    },
  });

  const res = await client.chat.completions.create({
    model: process.env.AZURE_OPENAI_DEPLOYMENT,
    messages: [{ role: "user", content: prompt }],
    temperature: 0,
  });

  return res.choices[0].message.content.trim();
}

/* ======================================================
 * PROMPT + TEXT UTILITIES
 * ====================================================== */

function buildPrompt({ placeholderName, marker, sentence, valueMap }) {
  return `
You are filling placeholders in a Word document with the best matching data.

Placeholder name:
${placeholderName}

Placeholder marker in text:
${marker}

Sentence from Word:
"${sentence}"

Available data fields (key : value preview):
${Object.entries(valueMap)
  .map(([k, v]) => `- ${k}: ${String(v).slice(0, 60)}`)
  .join("\n")}

Return ONLY the final replacement value for ${marker}.
If nothing is appropriate, return an empty string.
`.trim();
}

function stripXml(xml) {
  return xml
    .replace(/<[^>]+>/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function extractSentence(text, marker) {
  const idx = text.indexOf(marker);
  if (idx === -1) return "";

  const leftBound = Math.max(text.lastIndexOf("。", idx), text.lastIndexOf(".", idx));
  const start = leftBound === -1 ? 0 : leftBound + 1;

  const jpEnd = text.indexOf("。", idx + marker.length);
  const enEnd = text.indexOf(".", idx + marker.length);
  const ends = [jpEnd, enEnd].filter((i) => i !== -1);
  const end = ends.length ? Math.min(...ends) + 1 : text.length;

  return text.slice(start, end).trim();
}
