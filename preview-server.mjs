import { spawn } from "node:child_process";
import { watch } from "node:fs";
import { readFile, stat } from "node:fs/promises";
import { createServer } from "node:http";
import { execFileSync } from "node:child_process";
import { extname, dirname, join, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";

const root = dirname(fileURLToPath(import.meta.url));
const host = "127.0.0.1";
const port = 4173;
const mime = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
};
const clients = new Set();
let reloadTimer;

execFileSync(process.execPath, [join(root, "build-catalogue.mjs")], { stdio: "inherit" });

function notifyReload() {
  for (const response of clients) response.write("data: reload\n\n");
}

const server = createServer(async (request, response) => {
  const requestUrl = new URL(request.url, `http://${host}:${port}`);
  if (requestUrl.pathname === "/__events") {
    response.writeHead(200, {
      "Content-Type": "text/event-stream",
      "Cache-Control": "no-cache",
      Connection: "keep-alive",
    });
    response.write("retry: 1000\n\n");
    clients.add(response);
    request.on("close", () => clients.delete(response));
    return;
  }

  let pathname;
  try {
    pathname = decodeURIComponent(requestUrl.pathname);
  } catch {
    response.writeHead(400).end("Bad request");
    return;
  }
  const requested = pathname === "/" ? "index.html" : pathname.replace(/^\/+/, "");
  const file = resolve(root, requested);
  if (file !== root && !file.startsWith(root + sep)) {
    response.writeHead(403).end("Forbidden");
    return;
  }

  try {
    const fileStat = await stat(file);
    if (!fileStat.isFile()) throw new Error("Not a file");
    let body = await readFile(file);
    if (file === join(root, "index.html")) {
      const reloadScript = '<script>new EventSource("/__events").onmessage=()=>location.reload();</script>';
      body = Buffer.from(body.toString("utf8").replace("</body>", `${reloadScript}</body>`));
    }
    response.writeHead(200, {
      "Content-Type": mime[extname(file).toLowerCase()] || "application/octet-stream",
      "Cache-Control": "no-store",
    });
    response.end(body);
  } catch {
    response.writeHead(404).end("Not found");
  }
});

watch(root, { recursive: true }, (_event, filename) => {
  const relative = String(filename || "").replaceAll("\\", "/");
  if (relative.startsWith(".git/") || relative === "preview-server.mjs") return;
  clearTimeout(reloadTimer);
  reloadTimer = setTimeout(() => {
    if (relative === "materials.json" || relative.startsWith("content/")) {
      try {
        execFileSync(process.execPath, [join(root, "build-catalogue.mjs")], { stdio: "inherit" });
      } catch (error) {
        console.error("Не удалось пересобрать каталог:", error.message);
      }
    }
    notifyReload();
  }, 350);
});

server.listen(port, host, () => {
  const url = `http://${host}:${port}/`;
  console.log(`Локальный предпросмотр: ${url}`);
  console.log("Сохраняйте файлы проекта — страница обновится автоматически. Для остановки нажмите Ctrl+C.");
  spawn("cmd.exe", ["/c", "start", "", url], { detached: true, stdio: "ignore" }).unref();
});
