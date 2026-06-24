import { useState, useRef } from "react";
import { buildDocxBlob } from "./docx.js";

const C = {
  yellow: "#f7c948", dark: "#0f1117", mid: "#1a1f2e", surface: "#141824",
  border: "#2d3748", text: "#e2e8f0", muted: "#718096",
  green: "#68d391", blue: "#63b3ed", red: "#fc8181", orange: "#f6ad55",
};

function decodeBuffer(buf) {
  const b = new Uint8Array(buf);
  // BOM UTF-16 LE: FF FE
  if (b[0] === 0xFF && b[1] === 0xFE) return new TextDecoder("utf-16le").decode(buf);
  // BOM UTF-16 BE: FE FF
  if (b[0] === 0xFE && b[1] === 0xFF) return new TextDecoder("utf-16be").decode(buf);
  // No BOM — detect UTF-16 LE by null bytes pattern (every odd byte = 0x00 for ASCII/Latin)
  // Check first 20 bytes: if even-indexed bytes are printable and odd-indexed are 0x00 → UTF-16 LE
  let nullOdds = 0;
  const check = Math.min(b.length, 32);
  for (let i = 1; i < check; i += 2) { if (b[i] === 0x00) nullOdds++; }
  if (nullOdds >= check / 4) return new TextDecoder("utf-16le").decode(buf);
  // Default UTF-8
  return new TextDecoder("utf-8").decode(buf);
}

async function loadScript(src) {
  if (document.querySelector(`script[src="${src}"]`)) return;
  return new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = src; s.onload = res; s.onerror = rej;
    document.head.appendChild(s);
  });
}

async function parseFile(file) {
  const result = {
    fileName: file.name,
    sources: [], tables: [], measures: [],
    relationships: [], pages: [], roles: [], warnings: []
  };

  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js");
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/pako/2.1.0/pako.min.js");

  const arrayBuffer = await file.arrayBuffer();
  const zip = await window.JSZip.loadAsync(arrayBuffer);

  // Connections
  const connFile = zip.file("Connections");
  if (connFile) {
    try {
      const raw = decodeBuffer(await connFile.async("arraybuffer")).replace(/^\uFEFF/, "");
      result.sources = (JSON.parse(raw).Connections || []).map(c => ({
        name: c.ConnectionString || "Sin nombre", type: c.ConnectionType || "?"
      }));
    } catch (e) { result.warnings.push(`Connections: ${e.message}`); }
  }

  // DataModelSchema
  const modelFile = zip.file("DataModelSchema") || zip.file("DataModel");
  if (modelFile) {
    try {
      const buf = await modelFile.async("arraybuffer");
      let raw = decodeBuffer(buf).replace(/^\uFEFF/, "");

      // Fallback pako if still not JSON
      if (!raw.trim().startsWith("{") && !raw.trim().startsWith("[")) {
        try { raw = window.pako.inflate(new Uint8Array(buf), { to: "string" }).replace(/^\uFEFF/, ""); } catch (_) {}
      }

      const parsed = JSON.parse(raw);
      // .pbit = array [{model:{...}}] OR object {model:{...}}
      const db = Array.isArray(parsed) ? (parsed[0]?.model || parsed[0]) : (parsed.model || parsed);

      for (const t of (db.tables || [])) {
        if (t.name?.startsWith("DateTableTemplate") || t.name?.startsWith("LocalDateTable")) continue;
        const entry = { name: t.name, isHidden: !!t.isHidden, columns: [], measures: [] };
        for (const col of (t.columns || [])) {
          if (col.type === "rowNumber") continue;
          entry.columns.push({ name: col.name, dataType: col.dataType || "?", isHidden: !!col.isHidden, calculated: !!col.expression });
        }
        for (const m of (t.measures || [])) {
          const me = { name: m.name, expression: m.expression || "", formatString: m.formatString || "", table: t.name };
          entry.measures.push(me);
          result.measures.push(me);
        }
        result.tables.push(entry);
      }
      for (const r of (db.relationships || [])) {
        result.relationships.push({ from: `${r.fromTable}[${r.fromColumn}]`, to: `${r.toTable}[${r.toColumn}]`, cardinality: r.crossFilteringBehavior || "SingleDirection", active: r.isActive !== false });
      }
      for (const role of (db.roles || [])) {
        result.roles.push({ name: role.name, tableFilters: (role.tablePermissions || []).map(p => ({ table: p.name, filter: p.filterExpression || "" })) });
      }
    } catch (e) { result.warnings.push(`DataModelSchema: ${e.message}`); }
  } else {
    result.warnings.push("DataModelSchema no encontrado.");
  }

  // Report/Layout
  const layoutFile = zip.file("Report/Layout");
  if (layoutFile) {
    try {
      const buf = await layoutFile.async("arraybuffer");
      let raw = decodeBuffer(buf).replace(/^\uFEFF/, "");
      if (!raw.trim().startsWith("{")) {
        try { raw = window.pako.inflate(new Uint8Array(buf), { to: "string" }).replace(/^\uFEFF/, ""); } catch (_) {}
      }
      const layout = JSON.parse(raw);
      for (const section of (layout.sections || [])) {
        const page = { name: section.displayName || section.name || "Página", visuals: [] };
        for (const vc of (section.visualContainers || [])) {
          try {
            const cfg = typeof vc.config === "string" ? JSON.parse(vc.config) : vc.config;
            page.visuals.push({ type: cfg?.singleVisual?.visualType || "visual", title: cfg?.singleVisual?.vcObjects?.title?.[0]?.properties?.text?.expr?.Literal?.Value?.replace(/^'|'$/g, "") || "" });
          } catch (_) {}
        }
        result.pages.push(page);
      }
    } catch (e) { result.warnings.push(`Layout: ${e.message}`); }
  }

  return result;
}

async function groqFetch(body, groqKey, onLog) {
  const MAX_RETRIES = 4;
  for (let attempt = 0; ; attempt++) {
    const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${groqKey}` },
      body: JSON.stringify(body),
    });
    if (res.ok) return res.json();
    const err = await res.json().catch(() => ({}));
    const msg = err?.error?.message || `Groq ${res.status}`;
    if (res.status === 429) {
      // TPD (tokens per day) agotado: esperar minutos es inútil en sesión → señal para fallback de modelo.
      if (/per day|TPD/i.test(msg)) { const e = new Error(msg); e.isDailyLimit = true; throw e; }
      // TPM/RPM (por minuto): backoff y reintento.
      if (attempt < MAX_RETRIES) {
        const ra = parseFloat(res.headers.get("retry-after"));
        const waitMs = Number.isFinite(ra) ? ra * 1000 + 500 : Math.min(2000 * 2 ** attempt, 30000);
        onLog?.(`Rate limit por minuto — reintentando en ${Math.ceil(waitMs / 1000)}s (intento ${attempt + 1}/${MAX_RETRIES})...`, "warn");
        await new Promise(r => setTimeout(r, waitMs));
        continue;
      }
    }
    throw new Error(msg);
  }
}

async function analyzeWithGroq(parsed, groqKey, onLog) {
  const summary = {
    fileName: parsed.fileName,
    sources: parsed.sources,
    tables: parsed.tables.map(t => ({
      name: t.name, isHidden: t.isHidden,
      columns: t.columns.map(c => ({ name: c.name, type: c.dataType, calculated: c.calculated })),
      measures: t.measures.map(m => ({ name: m.name, expression: m.expression.substring(0, 200) })),
    })),
    relationships: parsed.relationships,
    pages: parsed.pages,
    roles: parsed.roles,
  };

  const prompt = `Eres experto en Power BI. Analiza este modelo y genera documentación técnica interna en español.

MODELO:
${JSON.stringify(summary, null, 2).substring(0, 14000)}

Responde SOLO con JSON válido (sin markdown), estructura exacta:
{
  "titulo": "Documentación Técnica: [nombre sin extension]",
  "fecha": "${new Date().toLocaleDateString("es-CO", { year: "numeric", month: "long", day: "numeric" })}",
  "resumen_ejecutivo": "3-4 oraciones sobre propósito, arquitectura y alcance.",
  "introduccion": "2-3 oraciones presentando el modelo, su dominio de negocio y para qué sirve.",
  "conclusiones": "2-3 oraciones de cierre: estado del modelo, fortalezas y recomendaciones.",
  "secciones": [
    { "id": "fuentes_datos", "titulo": "Fuentes de Datos", "contenido": "...", "items": [{"nombre":"...","tipo":"...","descripcion":"..."}] },
    { "id": "arquitectura_modelo", "titulo": "Arquitectura del Modelo", "contenido": "Patrón (estrella/copo/etc), convenciones, capas.", "items": [] },
    { "id": "tablas_columnas", "titulo": "Tablas y Columnas", "contenido": "...", "items": [{"tabla":"...","tipo":"Hecho/Dimensión/Staging/Parámetro","columnas_clave":["..."],"descripcion":"..."}] },
    { "id": "medidas_dax", "titulo": "Medidas DAX", "contenido": "...", "items": [{"nombre":"...","tabla":"...","expresion":"...","proposito":"..."}] },
    { "id": "relaciones", "titulo": "Relaciones del Modelo", "contenido": "...", "items": [{"desde":"...","hasta":"...","cardinalidad":"...","activa":true}] },
    { "id": "paginas_reporte", "titulo": "Páginas del Reporte", "contenido": "...", "items": [{"pagina":"...","visuales":["..."],"proposito":"..."}] }
  ]
}
Incluye solo secciones con datos reales. En medidas DAX explica propósito de negocio.`;

  const body = {
    max_tokens: 4000,
    temperature: 0.2,
    messages: [{ role: "user", content: prompt }],
    response_format: { type: "json_object" },
  };

  // 70b: mejor calidad pero solo 100K TPD. 8b-instant: ~500K TPD → respaldo cuando se agota el día.
  const MODELS = ["llama-3.3-70b-versatile", "llama-3.1-8b-instant"];
  let data;
  for (let i = 0; i < MODELS.length; i++) {
    try {
      if (i > 0) onLog?.(`Presupuesto diario (TPD) de ${MODELS[i - 1]} agotado. Cambiando a ${MODELS[i]}...`, "warn");
      data = await groqFetch({ ...body, model: MODELS[i] }, groqKey, onLog);
      if (i > 0) onLog?.(`Generado con modelo de respaldo ${MODELS[i]} (calidad menor que 70b).`, "ok");
      break;
    } catch (e) {
      if (e.isDailyLimit && i < MODELS.length - 1) continue;
      if (e.isDailyLimit) throw new Error("Presupuesto diario agotado en todos los modelos. Reintenta tras el reseteo (medianoche UTC).");
      throw e;
    }
  }

  return JSON.parse(data.choices?.[0]?.message?.content || "{}");
}

const inp = { background: "#0a0d13", border: `1px solid #2d3748`, borderRadius: 8, color: "#e2e8f0", fontSize: 13, padding: "10px 14px", width: "100%", boxSizing: "border-box", fontFamily: "monospace", outline: "none" };

export default function App() {
  const [groqKey, setGroqKey] = useState(localStorage.getItem("groq_key") || "");
  const [showKey, setShowKey] = useState(false);
  const [files, setFiles] = useState([]);
  const [logs, setLogs] = useState([]);
  const [progress, setProgress] = useState(0);
  const [running, setRunning] = useState(false);
  const [result, setResult] = useState(null);
  const [dragging, setDragging] = useState(false);
  const logRef = useRef(null);
  const lc = { info: C.blue, ok: C.green, warn: C.orange, error: C.red, section: C.yellow };

  const saveKey = k => { setGroqKey(k); localStorage.setItem("groq_key", k); };
  const addLog = (msg, type = "info") => {
    setLogs(p => [...p, { msg, type, time: new Date().toLocaleTimeString("es-CO") }]);
    setTimeout(() => logRef.current?.scrollTo(0, logRef.current.scrollHeight), 50);
  };
  const addFiles = fs => {
    const valid = fs.filter(f => f.name.endsWith(".pbix") || f.name.endsWith(".pbit"));
    setFiles(p => { const ex = new Set(p.map(f => f.name)); return [...p, ...valid.filter(f => !ex.has(f.name))]; });
  };

  const run = async () => {
    if (!files.length || running) return;
    if (!groqKey.trim()) { addLog("Ingresa tu Groq API key.", "error"); return; }
    setRunning(true); setLogs([]); setProgress(0); setResult(null);
    try {
      for (const file of files) {
        addLog(`Procesando: ${file.name}`, "section"); setProgress(5);
        addLog("Extrayendo estructura...", "info");
        const parsed = await parseFile(file);
        setProgress(35);
        parsed.warnings.forEach(w => addLog(w, "warn"));
        addLog(`✓ Tablas: ${parsed.tables.length} · Medidas: ${parsed.measures.length} · Relaciones: ${parsed.relationships.length} · Páginas: ${parsed.pages.length}`, "ok");
        addLog("Analizando con Groq llama-3.3-70b...", "info");
        const docData = await analyzeWithGroq(parsed, groqKey.trim(), addLog);
        setProgress(85); addLog("✓ Análisis completado", "ok");
        addLog("Generando documento Word...", "info");
        const blob = await buildDocxBlob(docData, parsed);
        const url = URL.createObjectURL(blob);
        setProgress(100); addLog("✅ Listo", "ok");
        setResult({ url, name: file.name.replace(/\.(pbix|pbit)$/, "_documentacion.docx"), docData, parsed });
      }
    } catch (e) { addLog(`Error: ${e.message}`, "error"); }
    finally { setRunning(false); }
  };

  return (
    <div style={{ background: C.dark, minHeight: "100vh", fontFamily: "'Segoe UI',system-ui,sans-serif", color: C.text }}>
      <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, padding: "14px 24px", display: "flex", alignItems: "center", gap: 12 }}>
        <span style={{ fontSize: 22 }}>📊</span>
        <div>
          <div style={{ fontWeight: 700, fontSize: 15, color: C.yellow }}>PBIX / PBIT Documenter</div>
          <div style={{ fontSize: 11, color: C.muted }}>Groq · llama-3.3-70b-versatile · Gratis</div>
        </div>
      </div>
      <div style={{ maxWidth: 760, margin: "0 auto", padding: "20px" }}>
        <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10, padding: 16, marginBottom: 16 }}>
          <div style={{ fontSize: 12, color: C.muted, marginBottom: 8, fontWeight: 600 }}>
            🔑 GROQ API KEY
            <a href="https://console.groq.com/keys" target="_blank" rel="noreferrer" style={{ color: C.blue, marginLeft: 10, fontSize: 11 }}>→ console.groq.com</a>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <input type={showKey ? "text" : "password"} placeholder="gsk_..." value={groqKey} onChange={e => saveKey(e.target.value)} style={{ ...inp, flex: 1 }} />
            <button onClick={() => setShowKey(p => !p)} style={{ background: C.mid, border: `1px solid ${C.border}`, borderRadius: 8, color: C.muted, padding: "0 14px", cursor: "pointer" }}>{showKey ? "🙈" : "👁"}</button>
          </div>
          {groqKey && <div style={{ fontSize: 11, color: C.green, marginTop: 6 }}>✓ Key guardada en este navegador</div>}
        </div>

        <div onDragOver={e => { e.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)}
          onDrop={e => { e.preventDefault(); setDragging(false); addFiles([...e.dataTransfer.files]); }}
          onClick={() => document.getElementById("fi").click()}
          style={{ border: `2px dashed ${dragging ? C.yellow : C.border}`, borderRadius: 10, padding: "32px 24px", textAlign: "center", cursor: "pointer", background: dragging ? C.mid : C.surface, transition: "all 0.2s" }}>
          <input id="fi" type="file" accept=".pbix,.pbit" multiple style={{ display: "none" }} onChange={e => addFiles([...e.target.files])} />
          <div style={{ fontSize: 28, marginBottom: 8 }}>📁</div>
          <div style={{ color: C.muted, fontSize: 14 }}>Arrastra .pbix / .pbit o haz clic</div>
        </div>

        {files.length > 0 && (
          <div style={{ marginTop: 14 }}>
            {files.map((f, i) => (
              <div key={i} style={{ background: C.mid, border: `1px solid ${C.border}`, borderRadius: 8, padding: "10px 14px", display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                <span>📊</span>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13, fontWeight: 600 }}>{f.name}</div>
                  <div style={{ fontSize: 11, color: C.muted }}>{(f.size / 1024 / 1024).toFixed(2)} MB</div>
                </div>
                <button onClick={() => setFiles(p => p.filter((_, j) => j !== i))} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 16 }}>✕</button>
              </div>
            ))}
            <button onClick={run} disabled={running} style={{ width: "100%", marginTop: 4, padding: 13, background: running ? "#5a4a10" : "linear-gradient(135deg,#f7c948,#f59e0b)", color: "#0f1117", fontWeight: 700, fontSize: 14, border: "none", borderRadius: 10, cursor: running ? "not-allowed" : "pointer" }}>
              {running ? "⏳ Procesando..." : "⚡ Generar Documentación"}
            </button>
            {(running || progress > 0) && (
              <div style={{ height: 4, background: C.border, borderRadius: 2, overflow: "hidden", marginTop: 10 }}>
                <div style={{ height: "100%", width: `${progress}%`, background: "linear-gradient(90deg,#f7c948,#f59e0b)", transition: "width 0.4s", borderRadius: 2 }} />
              </div>
            )}
          </div>
        )}

        {logs.length > 0 && (
          <div ref={logRef} style={{ background: "#0a0d13", border: `1px solid ${C.border}`, borderRadius: 8, padding: 14, marginTop: 16, maxHeight: 220, overflowY: "auto", fontFamily: "monospace", fontSize: 12 }}>
            {logs.map((l, i) => <div key={i} style={{ color: lc[l.type] || C.text, padding: "1px 0" }}>[{l.time}] {l.msg}</div>)}
          </div>
        )}

        {result && (
          <div style={{ background: "#0d1f0d", border: "1px solid #276227", borderRadius: 10, padding: 20, marginTop: 16 }}>
            <div style={{ color: C.green, fontWeight: 700, fontSize: 15, marginBottom: 12 }}>✅ Documentación generada</div>
            <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
              {[["Tablas", result.parsed.tables.length], ["Medidas", result.parsed.measures.length], ["Relaciones", result.parsed.relationships.length], ["Páginas", result.parsed.pages.length]].map(([l, n]) => (
                <div key={l} style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: "10px 16px", textAlign: "center", minWidth: 70 }}>
                  <div style={{ color: C.yellow, fontWeight: 700, fontSize: 20 }}>{n}</div>
                  <div style={{ color: C.muted, fontSize: 11 }}>{l}</div>
                </div>
              ))}
            </div>
            <div style={{ background: "#0a0d13", border: `1px solid ${C.border}`, borderRadius: 8, padding: 14, marginBottom: 14, maxHeight: 180, overflowY: "auto", fontSize: 12, color: C.muted, whiteSpace: "pre-wrap", lineHeight: 1.6 }}>
              {result.docData.resumen_ejecutivo}{"\n\n"}{result.docData.secciones?.map(s => `## ${s.titulo}\n${s.contenido || ""}`).join("\n\n").substring(0, 600)}...
            </div>
            <a href={result.url} download={result.name} style={{ display: "inline-block", padding: "10px 22px", background: "#276227", color: C.green, border: "1px solid #276227", borderRadius: 8, fontWeight: 600, fontSize: 13, textDecoration: "none" }}>
              ⬇ Descargar {result.name}
            </a>
          </div>
        )}
      </div>
    </div>
  );
}
