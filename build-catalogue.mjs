import { readFile, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = dirname(fileURLToPath(import.meta.url));
const cataloguePath = join(root, "materials.json");
const catalogue = JSON.parse(await readFile(cataloguePath, "utf8"));
const content = {};

for (const material of catalogue.materials) {
  if (!material.file) continue;
  content[material.id] = await readFile(join(root, material.file), "utf8");
}

const output = [
  `window.ELECTROBAZA_CATALOG = ${JSON.stringify(catalogue)};`,
  `window.ELECTROBAZA_CONTENT = ${JSON.stringify(content)};`,
  "",
].join("\n");

await writeFile(join(root, "catalogue-data.js"), output, "utf8");
console.log(`Готово: каталог и ${Object.keys(content).length} материалов собраны для index.html.`);
