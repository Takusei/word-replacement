import "dotenv/config";
import fs from "fs";
import PizZip from "pizzip";
import OpenAI from "openai";
import { DefaultAzureCredential, getBearerTokenProvider } from "@azure/identity";

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
 * @param {Record<string, string> | string} valueMap - Extracted data map or raw text
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

  // 4. Resolve all placeholders via a single LLM call
  const placeholderContext = placeholders.map((ph) => {
    const marker = `[${ph.id}]`;
    return {
      id: ph.id,
      raw: ph.raw,
      marker,
      sentence: extractSentence(plainText, marker),
      contextWindow: extractContextWindow(plainText, marker, 200),
    };
  });

  const batchPrompt = buildBatchPrompt({
    placeholders: placeholderContext,
    valueMap,
  });

  const batchResponse = await callLLM(batchPrompt, provider);
  const resolved = parseBatchResponse(batchResponse, placeholderContext);

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
  const credential = new DefaultAzureCredential();
  const scope = "https://cognitiveservices.azure.com/.default";
  const azureADTokenProvider = getBearerTokenProvider(credential, scope);
  const token = await azureADTokenProvider();

  const client = new OpenAI({
    apiKey: token,
    baseURL: `${process.env.AZURE_OPENAI_ENDPOINT}/openai/deployments/${process.env.AZURE_OPENAI_DEPLOYMENT_NAME}`,
    defaultQuery: {
      "api-version": process.env.AZURE_OPENAI_API_VERSION,
    },
  });

  const res = await client.chat.completions.create({
    model: process.env.AZURE_OPENAI_DEPLOYMENT_NAME,
    messages: [{ role: "user", content: prompt }],
    temperature: 0,
  });

  return res.choices[0].message.content.trim();
}

/* ======================================================
 * PROMPT + TEXT UTILITIES
 * ====================================================== */

function buildBatchPrompt({ placeholders, valueMap }) {
  const valueMapText = formatValueMapForPrompt(valueMap);
  return `
You are filling placeholders in a Word document with the best matching data.

Placeholders:
${placeholders
  .map(
    (ph, i) => `
${i + 1}.
- id: ${ph.id}
- marker: ${ph.marker}
- placeholder name: ${ph.raw}
- sentence: "${ph.sentence}"
- nearby context: "${ph.contextWindow}"
`.trim()
  )
  .join("\n\n")}

Available data:
${valueMapText}

Instructions:
- Resolve each placeholder independently using its sentence and nearby context.
- Prefer exact matches found in Available data.
- Do not infer from other placeholders.
- If the placeholder looks like a date field, search Available data for a date and return it as-is.
- If nothing is appropriate, return an empty string.

Return ONLY a valid JSON object mapping placeholder id to replacement string.
Example:
{"PH_1":"Acme Corp","PH_2":"2026-01-29"}
`.trim();
}

function formatValueMapForPrompt(valueMap) {
  if (typeof valueMap === "string") {
    return valueMap.trim() || "(empty)";
  }

  if (!valueMap || typeof valueMap !== "object" || Array.isArray(valueMap)) {
    return "(empty)";
  }

  const entries = Object.entries(valueMap);
  if (!entries.length) return "(empty)";

  return entries
    .map(([k, v]) => `- ${k}: ${String(v).slice(0, 60)}`)
    .join("\n");
}

function parseBatchResponse(text, placeholderContext) {
  const json = extractFirstJsonObject(text);
  const parsed = JSON.parse(json);

  const resolved = {};
  for (const ph of placeholderContext) {
    const value = parsed[ph.id];
    resolved[ph.id] = typeof value === "string" ? value : "";
  }
  return resolved;
}

function extractFirstJsonObject(text) {
  const start = text.indexOf("{");
  if (start === -1) {
    throw new Error("LLM response did not contain JSON object.");
  }

  let depth = 0;
  for (let i = start; i < text.length; i++) {
    const ch = text[i];
    if (ch === "{") depth++;
    if (ch === "}") depth--;
    if (depth === 0) {
      return text.slice(start, i + 1);
    }
  }

  throw new Error("LLM response JSON object was incomplete.");
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

function extractContextWindow(text, marker, windowSize = 200) {
  const idx = text.indexOf(marker);
  if (idx === -1) return "";

  const start = Math.max(0, idx - windowSize);
  const end = Math.min(text.length, idx + marker.length + windowSize);
  return text.slice(start, end).replace(/\s+/g, " ").trim();
}
