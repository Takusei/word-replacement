import { generateDocxWithLLM } from "../generateDocxWithLLM.js";

const ok = typeof generateDocxWithLLM === "function";
console.log(ok ? "smoke:ok" : "smoke:fail");
process.exit(ok ? 0 : 1);
