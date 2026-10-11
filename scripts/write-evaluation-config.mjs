import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const scriptDirectory = dirname(fileURLToPath(import.meta.url));
const repositoryRoot = resolve(scriptDirectory, "..");
const outputPath = resolve(repositoryRoot, "evaluation", "config.js");

const config = {
  supabaseUrl: process.env.EVALUATION_SUPABASE_URL || "",
  supabaseAnonKey: process.env.EVALUATION_SUPABASE_ANON_KEY || "",
};

await mkdir(dirname(outputPath), { recursive: true });
await writeFile(outputPath, `window.__EVALUATION_CONFIG__ = ${JSON.stringify(config)};\n`, "utf8");

if (!config.supabaseUrl || !config.supabaseAnonKey) {
  console.warn("Evaluation database secrets are missing; the deployed form will allow download but not online submission.");
} else {
  console.log("Evaluation database configuration generated.");
}

