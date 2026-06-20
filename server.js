const express = require("express");
const cors = require("cors");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const { generateCards } = require("./cardGenerator");

const app = express();
const PORT = process.env.PORT || 3001;

// ── Multer: recibe el Excel en memoria ────────────────────────────────────────
const upload = multer({
  storage: multer.memoryStorage(),
  fileFilter: (req, file, cb) => {
    const ok = [".xlsx", ".xls"].includes(path.extname(file.originalname).toLowerCase());
    cb(ok ? null : new Error("Solo se aceptan archivos .xlsx o .xls"), ok);
  },
  limits: { fileSize: 20 * 1024 * 1024 }, // 20 MB máx
});

app.use(cors());
app.use(express.json());

// ── GET /health — Railway lo usa para verificar que el server está vivo ───────
app.get("/health", (req, res) => res.json({ status: "ok" }));

// ── POST /generate — recibe Excel + metadatos, devuelve ZIP con los PDFs ──────
app.post("/generate", upload.single("excel"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "Falta el archivo Excel." });

  const {
    sheetName       = "Door",
    headerRow       = "4",           // índice base-0 de la fila de encabezados
    projectName     = "",
    generalContractor = "",
    subcontractor   = "",
    dateOfPrep      = "",
    driveUrl        = "",            // carpeta pública de Google Drive
  } = req.body;

  try {
    const zipBuffer = await generateCards({
      excelBuffer:   req.file.buffer,
      sheetName,
      headerRow:     parseInt(headerRow, 10),
      projectName,
      generalContractor,
      subcontractor,
      dateOfPrep,
      driveUrl,
    });

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", 'attachment; filename="door_cards.zip"');
    res.send(zipBuffer);
  } catch (err) {
    console.error("Error generando cards:", err);
    res.status(500).json({ error: err.message || "Error interno del servidor." });
  }
});

app.listen(PORT, () => console.log(`🚀  Backend corriendo en puerto ${PORT}`));
