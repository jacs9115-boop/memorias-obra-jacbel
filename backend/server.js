require("dotenv").config();
const express = require("express");
const cors = require("cors");
const path = require("path");

const PORT = process.env.PORT || 3000;
const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL;

const app = express();
app.use(cors());
app.use(express.json({ limit: "15mb" }));
app.use(express.static(path.join(__dirname, "..", "frontend")));

function requireAppsScriptUrl() {
  if (!APPS_SCRIPT_URL) throw new Error("Falta APPS_SCRIPT_URL");
}

async function llamarAppsScript(url, options) {
  const r = await fetch(url, options);
  return r.json();
}

async function llamarAppsScriptPost(body) {
  return llamarAppsScript(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
}

app.get("/api/health", (req, res) => res.json({ ok: true }));

app.get("/api/obras", async (req, res) => {
  try {
    requireAppsScriptUrl();
    res.json(await llamarAppsScript(`${APPS_SCRIPT_URL}?obras=1`));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.get("/api/presupuestos", async (req, res) => {
  try {
    requireAppsScriptUrl();
    res.json(await llamarAppsScript(`${APPS_SCRIPT_URL}?presupuestos=1`));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.get("/api/presupuesto-preview", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const { fileId } = req.query;
    if (!fileId) return res.status(400).json({ error: "Falta fileId" });
    const data = await llamarAppsScript(`${APPS_SCRIPT_URL}?previsualizar=${encodeURIComponent(fileId)}`);
    if (data && data.error) return res.status(400).json({ error: data.error });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.post("/api/obras", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const { fileId, nombreObra } = req.body;
    if (!fileId || !nombreObra) return res.status(400).json({ error: "Falta fileId o nombreObra" });
    const data = await llamarAppsScriptPost({ accion: "crear_obra", fileId, nombreObra });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.get("/api/obras/:obraId", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const data = await llamarAppsScript(`${APPS_SCRIPT_URL}?obra=${encodeURIComponent(req.params.obraId)}`);
    if (data && data.error) return res.status(400).json({ error: data.error });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.delete("/api/obras/:obraId", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const data = await llamarAppsScriptPost({ accion: "borrar_obra", obraId: req.params.obraId });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.post("/api/medidas", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const data = await llamarAppsScriptPost({ accion: "guardar_medida", ...req.body });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.put("/api/medidas/:medidaId", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const data = await llamarAppsScriptPost({ accion: "editar_medida", medidaId: req.params.medidaId, ...req.body });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.delete("/api/medidas/:medidaId", async (req, res) => {
  try {
    requireAppsScriptUrl();
    const { obraId } = req.body;
    if (!obraId) return res.status(400).json({ error: "Falta obraId" });
    const data = await llamarAppsScriptPost({ accion: "borrar_medida", medidaId: req.params.medidaId, obraId });
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Error inesperado" });
  }
});

app.listen(PORT, () => console.log(`Memorias JACBEL escuchando en puerto ${PORT}`));
