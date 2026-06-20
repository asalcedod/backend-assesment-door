/**
 * cardGenerator.js
 * Lee el Excel, descarga fotos de Google Drive, genera HTMLs, los convierte
 * a PDF con Playwright y los empaqueta en un ZIP en memoria.
 */

const XLSX         = require("xlsx");
const axios        = require("axios");
const archiver     = require("archiver");
const { chromium } = require("playwright");
const fs           = require("fs");
const path         = require("path");
const os           = require("os");
const http         = require("http");

// ── Extensiones de imagen soportadas ─────────────────────────────────────────
const IMAGE_EXTS = new Set([".jpg", ".jpeg", ".png", ".gif", ".webp"]);

// ── HTML template (se lee una vez al arrancar el servidor) ───────────────────
const TEMPLATE_PATH = path.join(__dirname, "template.html");
const TEMPLATE      = fs.readFileSync(TEMPLATE_PATH, "utf8");

// ── Helper: convierte un Buffer de imagen a data-URI base64 ──────────────────
function toDataUri(buffer, ext) {
  const mime = {
    ".jpg":  "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png":  "image/png",
    ".gif":  "image/gif",
    ".webp": "image/webp",
  }[ext] ?? "image/jpeg";
  return `data:${mime};base64,${buffer.toString("base64")}`;
}

// ── Helper: construye el HTML del grid de fotos ───────────────────────────────
function buildPhotoHtml(photos) {
  if (!photos || photos.length === 0) {
    return `<div class="w-full min-h-[120px] border-2 border-dashed border-gray-300
              flex items-center justify-center bg-gray-50 text-gray-400 italic text-sm">
              No photo available</div>`;
  }

  const cols  = photos.length === 1 ? "grid-cols-1" : "grid-cols-2";
  const items = photos.map((p, i) => `
    <div class="relative flex items-center justify-center bg-gray-50 border border-gray-200 rounded overflow-hidden">
      <img src="${p.dataUri}" alt="Photo ${i + 1}"
           style="max-width:100%; max-height:120px; width:auto; height:auto; object-fit:contain; display:block;">
      <span class="absolute bottom-1 left-1 bg-[#0a3a70] text-white
                   text-[8px] font-bold px-1 rounded opacity-80">${i + 1}</span>
    </div>`).join("\n");

  return `<div class="grid ${cols} gap-2">${items}</div>`;
}

// ── Helper: limpia valores vacíos / NaN del Excel ─────────────────────────────
function clean(val) {
  if (val === undefined || val === null) return "";
  const s = String(val).trim();
  return s.toLowerCase() === "nan" ? "" : s;
}

// ── Google Drive: convierte link de carpeta pública a API list ────────────────
// Soporta links del tipo:
//   https://drive.google.com/drive/folders/<FOLDER_ID>
//   https://drive.google.com/drive/folders/<FOLDER_ID>?usp=sharing
function parseFolderId(driveUrl) {
  const match = driveUrl.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/**
 * Descarga todas las imágenes de una sub-carpeta de Drive cuyo nombre
 * coincide con doorId. Devuelve un array de { dataUri } o [] si no encuentra.
 *
 * Usa la API pública (sin autenticación) de Google Drive a través del
 * endpoint de exportación de archivos públicos.
 */
async function downloadDrivePhotos(rootFolderId, doorId) {
  try {
    // 1. Listar contenido de la carpeta raíz para encontrar la sub-carpeta del doorId
    const listUrl = `https://www.googleapis.com/drive/v3/files`
      + `?q='${rootFolderId}'+in+parents+and+name='${doorId}'+and+mimeType='application/vnd.google-apps.folder'+and+trashed=false`
      + `&fields=files(id,name)`
      + `&key=${process.env.GOOGLE_API_KEY}`;

    const listRes  = await axios.get(listUrl, { timeout: 15000 });
    const folders  = listRes.data.files ?? [];
    if (folders.length === 0) return [];

    const subFolderId = folders[0].id;

    // 2. Listar imágenes dentro de esa sub-carpeta
    const imgUrl = `https://www.googleapis.com/drive/v3/files`
      + `?q='${subFolderId}'+in+parents+and+trashed=false`
      + `&fields=files(id,name,mimeType)`
      + `&key=${process.env.GOOGLE_API_KEY}`;

    const imgRes = await axios.get(imgUrl, { timeout: 15000 });
    const files  = (imgRes.data.files ?? []).filter(f =>
      IMAGE_EXTS.has(path.extname(f.name).toLowerCase())
    );

    // 3. Descargar cada imagen como buffer
    const photos = [];
    for (const file of files) {
      const dlUrl  = `https://www.googleapis.com/drive/v3/files/${file.id}?alt=media&key=${process.env.GOOGLE_API_KEY}`;
      const dlRes  = await axios.get(dlUrl, { responseType: "arraybuffer", timeout: 30000 });
      const ext    = path.extname(file.name).toLowerCase();
      photos.push({ dataUri: toDataUri(Buffer.from(dlRes.data), ext) });
    }
    return photos;

  } catch (err) {
    console.warn(`⚠ No se pudieron descargar fotos de Drive para ${doorId}:`, err.message);
    return [];
  }
}

// ── Función principal exportada ───────────────────────────────────────────────
async function generateCards({
  excelBuffer,
  sheetName,
  headerRow,
  projectName,
  generalContractor,
  subcontractor,
  dateOfPrep,
  driveUrl,
}) {
  // 1. Parsear Excel
  const wb       = XLSX.read(excelBuffer, { type: "buffer" });
  const ws       = wb.Sheets[sheetName];
  if (!ws) throw new Error(`No se encontró la hoja "${sheetName}" en el Excel.`);

  // XLSX.utils.sheet_to_json con defval:"" y range para saltar filas de encabezado extra
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "", range: headerRow });

  // 2. Google Drive folder ID
  const rootFolderId = driveUrl ? parseFolderId(driveUrl) : null;

  // 3. Generar HTMLs en un directorio temporal
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "door-cards-"));
  const htmlDir = path.join(tmpDir, "html");
  fs.mkdirSync(htmlDir);

  const htmlFiles = []; // [{ doorId, filePath }]

  for (const row of rows) {
    const doorId = clean(row["ID Door"]);
    if (!doorId) continue;

    // Fotos de Google Drive
    const photos = rootFolderId
      ? await downloadDrivePhotos(rootFolderId, doorId)
      : [];

    const replacements = {
      // Encabezado
      "{{ PROJECT_NAME }}":                         projectName,
      "{{ GENERAL_CONTRACTOR }}":                   generalContractor,
      "{{ SUBCONTRACTOR }}":                        subcontractor,
      "{{ DATE_OF_PREPARATION }}":                  dateOfPrep,
      "{{ ID_DOOR }}":                              doorId,
      // 01. Location
      "{{ INSPECTION_LEVEL }}":                     clean(row["Inspection Level"]),
      "{{ ROOM_NAME }}":                            clean(row["Room Name"]),
      "{{ ACCESS_STATUS }}":                        clean(row["Access Status"]),
      "{{ DOOR_ORIGIN }}":                          clean(row["Door Origin"]),
      // 02. Field Notes
      "{{ DOOR_MATERIAL }}":                        clean(row["Door Material"]),
      "{{ DOOR_TYPE }}":                            clean(row["Door Type"]),
      "{{ DOOR_NOTES }}":                           clean(row["Door - Notes"]),
      "{{ FRAME_NOTES }}":                          clean(row["Frame - Notes"]),
      "{{ TRANSOM_NOTES }}":                        clean(row["Transom - Notes"]),
      // 03. Assessment
      "{{ TECHNICAL_NOTES }}":                      clean(row["Technical Notes"]),
      "{{ PRIORITY }}":                             clean(row["Priority"]),
      "{{ EVALUATION_SCORE }}":                     clean(row["Evaluation score"]),
      // 04. Photo
      "{{ PHOTO }}":                                buildPhotoHtml(photos),
      // 05-10. Components
      "{{ DOOR_CURRENT_STATUS }}":                  clean(row["Door - Current Status"]),
      "{{ DOOR_RECOMMENDED_ACTION }}":              clean(row["Door - Recommended Action"]),
      "{{ FRAME_CURRENT_STATUS }}":                 clean(row["Frame - Current Status"]),
      "{{ FRAME_RECOMMENDED_ACTION }}":             clean(row["Frame - Recommended Action"]),
      "{{ HINGES_CURRENT_STATUS }}":                clean(row["Hinges - Current Status"]),
      "{{ HINGES_RECOMMENDED_ACTION }}":            clean(row["Hinges - Recommended Action"]),
      "{{ TRANSOM_CURRENT_STATUS }}":               clean(row["Transom - Current Status"]),
      "{{ TRANSOM_RECOMMENDED_ACTION }}":           clean(row["Transom - Recommended Action"]),
      "{{ TRANSOM_FUNCTION }}":                     clean(row["Transom - Function"]),
      "{{ DOOR_HANDLE_CURRENT_STATUS }}":           clean(row["Door handle - Current Status"]),
      "{{ DOOR_HANDLE_RECOMMENDED_ACTION }}":       clean(row["Door handle - Recommended Action"]),
      "{{ AUTOMATIC_CLOSING_ARM_CURRENT_STATUS }}": clean(row["Automatic closing arm - Current Status"]),
    };

    let html = TEMPLATE;
    for (const [key, val] of Object.entries(replacements)) {
      html = html.replaceAll(key, val);
    }

    const filePath = path.join(htmlDir, `${doorId}_card.html`);
    fs.writeFileSync(filePath, html, "utf8");
    htmlFiles.push({ doorId, filePath });
    console.log(`✓ HTML  ${doorId}`);
  }

  // 4. Convertir HTMLs → PDFs con Playwright
  const pdfDir = path.join(tmpDir, "pdf");
  fs.mkdirSync(pdfDir);

  // Levantar un servidor HTTP local para servir los HTMLs
  // (file:// está bloqueado en Chromium headless dentro de contenedores)
  const server = http.createServer((req, res) => {
    const filePath = path.join(htmlDir, req.url.slice(1)); // quitar el / inicial
    if (!fs.existsSync(filePath)) { res.writeHead(404); res.end(); return; }
    const html = fs.readFileSync(filePath, "utf8");
    res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
    res.end(html);
  });

  // Puerto aleatorio disponible
  await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
  const { port } = server.address();
  console.log(`🌐  Servidor HTML local en puerto ${port}`);

  const browser = await chromium.launch({
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });
  const page = await browser.newPage();

  for (const { doorId } of htmlFiles) {
    const url     = `http://127.0.0.1:${port}/${doorId}_card.html`;
    const pdfPath = path.join(pdfDir, `${doorId}_card.pdf`);

    await page.goto(url, { waitUntil: "networkidle", timeout: 30000 });
    await page.pdf({
      path:            pdfPath,
      format:          "Letter",
      printBackground: true,
      margin: { top: "10mm", bottom: "10mm", left: "10mm", right: "10mm" },
    });
    console.log(`✓ PDF   ${doorId}`);
  }

  await browser.close();
  await new Promise(resolve => server.close(resolve));

  // 5. Empaquetar PDFs en un ZIP en memoria
  const zipBuffer = await new Promise((resolve, reject) => {
    const archive = archiver("zip", { zlib: { level: 6 } });
    const chunks  = [];
    archive.on("data",  chunk  => chunks.push(chunk));
    archive.on("end",   ()     => resolve(Buffer.concat(chunks)));
    archive.on("error", reject);
    archive.directory(pdfDir, false);
    archive.finalize();
  });

  // 6. Limpiar directorio temporal
  fs.rmSync(tmpDir, { recursive: true, force: true });

  return zipBuffer;
}

module.exports = { generateCards };
