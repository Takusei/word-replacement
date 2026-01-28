import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { generateDocxWithLLM } from "../generateDocxWithLLM.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const [templatePathArg, outputPathArg, valueMapPathArg] = process.argv.slice(2);

const templatePath = templatePathArg || process.env.TEMPLATE_PATH;
const outputPath = outputPathArg || process.env.OUTPUT_PATH || path.join(__dirname, "..", "output.docx");
const valueMapPath = valueMapPathArg || process.env.VALUE_MAP_PATH;
const provider = process.env.LLM_PROVIDER || "openai";

if (!templatePath) {
  console.error("Missing template path. Provide as arg or TEMPLATE_PATH env var.");
  process.exit(1);
}

if (!valueMapPath) {
  console.error("Missing value map JSON path. Provide as arg or VALUE_MAP_PATH env var.");
  process.exit(1);
}

const rawMap = fs.readFileSync(valueMapPath, "utf8");
const valueMap = JSON.parse(rawMap);

const buffer = await generateDocxWithLLM(templatePath, valueMap, provider);
fs.writeFileSync(outputPath, buffer);

console.log(`Wrote ${outputPath}`);
