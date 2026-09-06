import React, { useState, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  UploadCloud,
  FileJson2,
  X,
  Download,
  ClipboardPaste,
  Trash2,
  CheckCircle2,
  AlertTriangle,
  FileSpreadsheet,
  FlaskConical,
} from "lucide-react";

/* ----------------------------------------------------------------
   Paleta y tipografía
   - Fondo: papel frío azulado (no el crema/terracota genérico)
   - Tinta: azul-marino casi negro
   - Acento: dorado sello oficial, para la acción principal
   ---------------------------------------------------------------- */
const C = {
  page: "#EDF1F5",
  surface: "#FFFFFF",
  surfaceSoft: "#F5F7FA",
  ink: "#152238",
  inkSoft: "#4B5872",
  inkFaint: "#8792A6",
  line: "#D9DFE7",
  lineStrong: "#C2CAD6",
  accent: "#B9862E",
  accentDeep: "#8E6720",
  accentSoft: "#F4E9D3",
  navyTint: "#DEE6F3",
  navyText: "#28407A",
  goldTint: "#F6E9C9",
  goldText: "#7A5A12",
  greyTint: "#E6E8EC",
  greyText: "#525C6E",
  danger: "#B0413E",
  dangerTint: "#F6E1DF",
  success: "#2E7457",
  successTint: "#DFEEE6",
};

const serif = "Georgia, 'Iowan Old Style', 'Times New Roman', serif";
const sans =
  "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif";

const card = {
  background: C.surface,
  border: `1px solid ${C.line}`,
  borderRadius: 10,
};

/* ----------------------------------------------------------------
   Catálogos DTE (Ministerio de Hacienda, El Salvador)
   ---------------------------------------------------------------- */
const TIPO_DTE = {
  "01": "Factura",
  "03": "Comprobante de Crédito Fiscal",
  "04": "Nota de Remisión",
  "05": "Nota de Crédito",
  "06": "Nota de Débito",
  "07": "Comprobante de Retención",
  "08": "Comprobante de Liquidación",
  "09": "Documento Contable de Liquidación",
  11: "Factura de Exportación",
  14: "Factura de Sujeto Excluido",
  15: "Comprobante de Donación",
};
const CONDICION_OP = { 1: "Contado", 2: "Crédito", 3: "Otro" };

function bucketFor(tipoDte) {
  if (tipoDte === "03") return "Crédito Fiscal";
  if (tipoDte === "01") return "Consumidor Final";
  return "Otros DTE";
}
const BUCKET_STYLE = {
  "Crédito Fiscal": { bg: C.navyTint, fg: C.navyText },
  "Consumidor Final": { bg: C.goldTint, fg: C.goldText },
  "Otros DTE": { bg: C.greyTint, fg: C.greyText },
  "JSON genérico": { bg: C.greyTint, fg: C.greyText },
};

const num = (v) => {
  const n = typeof v === "string" ? parseFloat(v) : v;
  return typeof n === "number" && !Number.isNaN(n) ? n : "";
};

function flattenObject(obj, prefix = "", res = {}) {
  Object.entries(obj || {}).forEach(([k, v]) => {
    const key = prefix ? `${prefix}.${k}` : k;
    if (v && typeof v === "object" && !Array.isArray(v)) {
      flattenObject(v, key, res);
    } else if (Array.isArray(v)) {
      res[key] =
        v.length && typeof v[0] === "object"
          ? JSON.stringify(v)
          : v.join(", ");
    } else {
      res[key] = v ?? "";
    }
  });
  return res;
}

function normalizeDte(json, filename) {
  const ident = json.identificacion || {};
  const emisor = json.emisor || {};
  const receptor = json.receptor || {};
  const resumen = json.resumen || {};
  const items = Array.isArray(json.cuerpoDocumento) ? json.cuerpoDocumento : [];
  const tipoDte = String(ident.tipoDte ?? "");
  const bucket = bucketFor(tipoDte);

  let iva = "";
  if (typeof resumen.totalIva === "number") iva = resumen.totalIva;
  else if (Array.isArray(resumen.tributos)) {
    const t = resumen.tributos.find((t) => t.codigo === "20");
    if (t) iva = t.valor;
  }

  return {
    kind: "dte",
    bucket,
    tipoLabel: TIPO_DTE[tipoDte] || (tipoDte ? `Otro (${tipoDte})` : "Desconocido"),
    archivo: filename,
    codigoGeneracion: ident.codigoGeneracion || "",
    numeroControl: ident.numeroControl || "",
    fechaEmision: ident.fecEmi || "",
    horaEmision: ident.horEmi || "",
    emisorNombre: emisor.nombre || emisor.nombreComercial || "",
    emisorNit: emisor.nit || "",
    emisorNrc: emisor.nrc || "",
    receptorNombre:
      receptor.nombre || (bucket === "Consumidor Final" ? "Consumidor Final" : ""),
    receptorDoc: receptor.nit || receptor.numDocumento || receptor.dui || "",
    ventaGravada: num(resumen.totalGravada ?? resumen.subTotal),
    ventaExenta: num(resumen.totalExenta),
    ventaNoSuj: num(resumen.totalNoSuj),
    descuento: num(resumen.totalDescu),
    iva: num(iva),
    ivaRetenido: num(resumen.ivaRete1),
    rentaRetenida: num(resumen.reteRenta),
    total: num(resumen.montoTotalOperacion ?? resumen.totalPagar),
    condicionOperacion: CONDICION_OP[resumen.condicionOperacion] || "",
    sello: json.selloRecibido || json.sello_recibido || "",
    items: items.map((it, idx) => ({
      archivo: filename,
      numeroControl: ident.numeroControl || "",
      numItem: it.numItem ?? idx + 1,
      codigo: it.codigo || "",
      descripcion: it.descripcion || "",
      cantidad: num(it.cantidad),
      unidad: it.uniMedida ?? it.unidadMedida ?? "",
      precioUni: num(it.precioUni),
      descu: num(it.montoDescu),
      ventaNoSuj: num(it.ventaNoSuj),
      ventaExenta: num(it.ventaExenta),
      ventaGravada: num(it.ventaGravada),
    })),
  };
}

function normalizeGeneric(json, filename) {
  const flat = { Archivo: filename, ...flattenObject(json) };
  return { kind: "generic", bucket: "JSON genérico", archivo: filename, flat };
}

function parseAny(json, filename) {
  return json && typeof json === "object" && json.identificacion
    ? normalizeDte(json, filename)
    : normalizeGeneric(json, filename);
}

/* ----------------------------------------------------------------
   Construcción del Excel
   ---------------------------------------------------------------- */
const RESUMEN_COLS = [
  ["archivo", "Archivo"],
  ["tipoLabel", "Tipo de documento"],
  ["codigoGeneracion", "Código de generación"],
  ["numeroControl", "Número de control"],
  ["fechaEmision", "Fecha de emisión"],
  ["horaEmision", "Hora de emisión"],
  ["emisorNombre", "Emisor"],
  ["emisorNit", "NIT emisor"],
  ["emisorNrc", "NRC emisor"],
  ["receptorNombre", "Receptor"],
  ["receptorDoc", "NIT / DUI receptor"],
  ["ventaGravada", "Venta gravada"],
  ["ventaExenta", "Venta exenta"],
  ["ventaNoSuj", "Venta no sujeta"],
  ["descuento", "Descuento"],
  ["iva", "IVA"],
  ["ivaRetenido", "IVA retenido"],
  ["rentaRetenida", "Renta retenida"],
  ["total", "Total a pagar"],
  ["condicionOperacion", "Condición de operación"],
  ["sello", "Sello de recibido"],
];
const MONEY_KEYS = new Set([
  "ventaGravada",
  "ventaExenta",
  "ventaNoSuj",
  "descuento",
  "iva",
  "ivaRetenido",
  "rentaRetenida",
  "total",
  "precioUni",
  "descu",
]);
const ITEM_COLS = [
  ["archivo", "Archivo"],
  ["numeroControl", "Número de control"],
  ["numItem", "N° item"],
  ["codigo", "Código"],
  ["descripcion", "Descripción"],
  ["cantidad", "Cantidad"],
  ["unidad", "Unidad"],
  ["precioUni", "Precio unitario"],
  ["descu", "Descuento"],
  ["ventaNoSuj", "Venta no sujeta"],
  ["ventaExenta", "Venta exenta"],
  ["ventaGravada", "Venta gravada"],
];

function sheetFromCols(rows, cols) {
  const header = cols.map(([, label]) => label);
  const aoa = [header, ...rows.map((r) => cols.map(([key]) => r[key] ?? ""))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  cols.forEach(([key], c) => {
    if (MONEY_KEYS.has(key)) {
      for (let r = 1; r <= rows.length; r++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (cell && typeof cell.v === "number") cell.z = "#,##0.00";
      }
    }
  });
  ws["!cols"] = cols.map(([, label]) => ({
    wch: Math.max(12, Math.min(32, label.length + 4)),
  }));
  if (rows.length) {
    ws["!autofilter"] = {
      ref: XLSX.utils.encode_range(
        { r: 0, c: 0 },
        { r: rows.length, c: cols.length - 1 }
      ),
    };
  }
  return ws;
}

function sheetFromGeneric(flatRows) {
  const keySet = [];
  flatRows.forEach((r) =>
    Object.keys(r).forEach((k) => {
      if (!keySet.includes(k)) keySet.push(k);
    })
  );
  const aoa = [keySet, ...flatRows.map((r) => keySet.map((k) => r[k] ?? ""))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = keySet.map((k) => ({ wch: Math.max(12, Math.min(32, k.length + 4)) }));
  return ws;
}

function addSheet(wb, name, ws) {
  let n = name.slice(0, 31);
  let i = 2;
  const used = wb.SheetNames;
  while (used.includes(n)) {
    const suffix = ` (${i++})`;
    n = name.slice(0, 31 - suffix.length) + suffix;
  }
  XLSX.utils.book_append_sheet(wb, ws, n);
}

function buildWorkbook(docs, { separarPorTipo, incluirDetalle }) {
  const dte = docs.filter((d) => d.kind === "dte");
  const generic = docs.filter((d) => d.kind === "generic");
  const wb = XLSX.utils.book_new();

  if (dte.length) {
    if (separarPorTipo) {
      ["Crédito Fiscal", "Consumidor Final", "Otros DTE"].forEach((b) => {
        const rows = dte.filter((d) => d.bucket === b);
        if (rows.length) addSheet(wb, `Resumen - ${b}`, sheetFromCols(rows, RESUMEN_COLS));
      });
    } else {
      addSheet(wb, "Resumen DTE", sheetFromCols(dte, RESUMEN_COLS));
    }

    if (incluirDetalle) {
      if (separarPorTipo) {
        ["Crédito Fiscal", "Consumidor Final", "Otros DTE"].forEach((b) => {
          const items = dte.filter((d) => d.bucket === b).flatMap((d) => d.items);
          if (items.length) addSheet(wb, `Detalle - ${b}`, sheetFromCols(items, ITEM_COLS));
        });
      } else {
        const items = dte.flatMap((d) => d.items);
        if (items.length) addSheet(wb, "Detalle de items", sheetFromCols(items, ITEM_COLS));
      }
    }
  }

  if (generic.length) {
    addSheet(wb, "JSON genérico", sheetFromGeneric(generic.map((d) => d.flat)));
  }

  if (!wb.SheetNames.length) {
    addSheet(wb, "Vacío", XLSX.utils.aoa_to_sheet([["No hay datos para exportar"]]));
  }
  return wb;
}

/* ----------------------------------------------------------------
   Datos de ejemplo (para probar la app sin archivos propios)
   ---------------------------------------------------------------- */
const EJEMPLOS = [
  {
    filename: "ejemplo-consumidor-final.json",
    json: {
      identificacion: {
        tipoDte: "01",
        codigoGeneracion: "6E4F1C2A-0000-4E00-9A00-000000000001",
        numeroControl: "DTE-01-M001P001-000000000000123",
        fecEmi: "2026-08-14",
        horEmi: "10:32:00",
        tipoMoneda: "USD",
      },
      emisor: { nombre: "Comercial San Isidro, S.A. de C.V.", nit: "0614-010190-101-2", nrc: "123456" },
      receptor: { nombre: "" },
      cuerpoDocumento: [
        { numItem: 1, codigo: "A-100", descripcion: "Cemento gris 42.5 kg", cantidad: 10, uniMedida: "Unidad", precioUni: 6.5, ventaGravada: 65 },
        { numItem: 2, codigo: "A-220", descripcion: "Varilla de hierro 3/8", cantidad: 20, uniMedida: "Unidad", precioUni: 4.25, ventaGravada: 85 },
      ],
      resumen: { totalGravada: 150, totalExenta: 0, totalNoSuj: 0, totalDescu: 0, totalIva: 19.5, montoTotalOperacion: 169.5, condicionOperacion: 1 },
    },
  },
  {
    filename: "ejemplo-credito-fiscal.json",
    json: {
      identificacion: {
        tipoDte: "03",
        codigoGeneracion: "9B2A3D4E-0000-4E00-9A00-000000000002",
        numeroControl: "DTE-03-M001P001-000000000000045",
        fecEmi: "2026-08-20",
        horEmi: "15:05:00",
        tipoMoneda: "USD",
      },
      emisor: { nombre: "Distribuidora Cuscatlán, S.A. de C.V.", nit: "0614-020295-102-3", nrc: "654321" },
      receptor: { nombre: "Ferretería El Progreso, S.A. de C.V.", nit: "0614-030300-103-4" },
      cuerpoDocumento: [
        { numItem: 1, codigo: "B-010", descripcion: "Pintura acrílica blanca 5 gal", cantidad: 4, uniMedida: "Cubeta", precioUni: 45, ventaGravada: 180 },
        { numItem: 2, codigo: "B-045", descripcion: "Brocha profesional 4 pulg", cantidad: 15, uniMedida: "Unidad", precioUni: 3.2, ventaGravada: 48 },
      ],
      resumen: { totalGravada: 228, totalExenta: 0, totalNoSuj: 0, totalDescu: 0, tributos: [{ codigo: "20", descripcion: "IVA", valor: 29.64 }], montoTotalOperacion: 257.64, condicionOperacion: 2 },
    },
  },
];

/* ----------------------------------------------------------------
   Componente principal
   ---------------------------------------------------------------- */
export default function JsonToExcelApp() {
  const [entries, setEntries] = useState([]); // {id, name, status, doc?, errorMsg?}
  const [separarPorTipo, setSepararPorTipo] = useState(true);
  const [incluirDetalle, setIncluirDetalle] = useState(true);
  const [dragOver, setDragOver] = useState(false);
  const [pasteOpen, setPasteOpen] = useState(false);
  const [pasteText, setPasteText] = useState("");
  const [pasteError, setPasteError] = useState("");
  const [activeTab, setActiveTab] = useState("resumen");
  const fileInputRef = useRef(null);

  const addEntry = useCallback((name, text) => {
    const id = `${name}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
    try {
      const json = JSON.parse(text);
      const doc = parseAny(json, name);
      setEntries((prev) => [...prev, { id, name, status: "ok", doc }]);
    } catch (e) {
      setEntries((prev) => [
        ...prev,
        { id, name, status: "error", errorMsg: "No se pudo interpretar como JSON válido" },
      ]);
    }
  }, []);

  const handleFiles = useCallback(
    (fileList) => {
      Array.from(fileList).forEach((file) => {
        const reader = new FileReader();
        reader.onload = (e) => addEntry(file.name, String(e.target.result || ""));
        reader.onerror = () =>
          setEntries((prev) => [
            ...prev,
            { id: `${file.name}-${Date.now()}`, name: file.name, status: "error", errorMsg: "No se pudo leer el archivo" },
          ]);
        reader.readAsText(file);
      });
    },
    [addEntry]
  );

  const onDrop = useCallback(
    (e) => {
      e.preventDefault();
      setDragOver(false);
      handleFiles(e.dataTransfer.files);
    },
    [handleFiles]
  );

  const removeEntry = (id) => setEntries((prev) => prev.filter((d) => d.id !== id));
  const clearAll = () => setEntries([]);

  const loadEjemplos = () => {
    EJEMPLOS.forEach((ex) => addEntry(ex.filename, JSON.stringify(ex.json)));
  };

  const submitPaste = () => {
    if (!pasteText.trim()) return;
    try {
      JSON.parse(pasteText);
      addEntry(`pegado-${entries.length + 1}.json`, pasteText);
      setPasteText("");
      setPasteError("");
      setPasteOpen(false);
    } catch (e) {
      setPasteError("Ese texto no es un JSON válido. Revisa las comillas y comas.");
    }
  };

  const okEntries = entries.filter((e) => e.status === "ok");
  const errorEntries = entries.filter((e) => e.status === "error");
  const docs = okEntries.map((e) => e.doc);
  const dteDocs = docs.filter((d) => d.kind === "dte");

  const stats = useMemo(() => {
    const totalMonto = dteDocs.reduce((a, d) => a + (typeof d.total === "number" ? d.total : 0), 0);
    const cf = dteDocs.filter((d) => d.bucket === "Crédito Fiscal").length;
    const cons = dteDocs.filter((d) => d.bucket === "Consumidor Final").length;
    const otros = dteDocs.filter((d) => d.bucket === "Otros DTE").length;
    return { totalMonto, cf, cons, otros, totalItems: dteDocs.reduce((a, d) => a + d.items.length, 0) };
  }, [dteDocs]);

  const previewRows = useMemo(() => {
    if (activeTab === "resumen") return dteDocs.slice(0, 8);
    return dteDocs.flatMap((d) => d.items).slice(0, 10);
  }, [activeTab, dteDocs]);

  const download = () => {
    if (!docs.length) return;
    const wb = buildWorkbook(docs, { separarPorTipo, incluirDetalle });
    const today = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `dte_el_salvador_${today}.xlsx`);
  };

  const money = (v) =>
    typeof v === "number" ? v.toLocaleString("es-SV", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : "—";

  return (
    <div style={{ background: C.page, minHeight: "100vh", fontFamily: sans, color: C.ink }}>
      {/* Encabezado */}
      <header style={{ borderBottom: `1px solid ${C.line}`, background: C.surface }}>
        <div style={{ maxWidth: 1120, margin: "0 auto", padding: "28px 24px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <div
              style={{
                width: 42,
                height: 42,
                borderRadius: 9,
                background: C.ink,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                flexShrink: 0,
              }}
            >
              <FileSpreadsheet size={22} color={C.accentSoft} />
            </div>
            <div>
              <h1 style={{ fontFamily: serif, fontSize: 26, margin: 0, letterSpacing: 0.2 }}>
                De DTE a Excel
              </h1>
              <p style={{ margin: "2px 0 0", color: C.inkSoft, fontSize: 14.5 }}>
                Reúne tus Comprobantes de Crédito Fiscal y Facturas de Consumidor Final en un solo libro, listo para revisar.
              </p>
            </div>
          </div>
        </div>
      </header>

      <main style={{ maxWidth: 1120, margin: "0 auto", padding: "28px 24px 64px" }}>
        <div className="app-grid">
          {/* Columna izquierda */}
          <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
            {/* Dropzone */}
            <div
              style={{ ...card, padding: 18 }}
              onDragOver={(e) => {
                e.preventDefault();
                setDragOver(true);
              }}
              onDragLeave={() => setDragOver(false)}
              onDrop={onDrop}
            >
              <div
                onClick={() => fileInputRef.current?.click()}
                style={{
                  border: `1.5px dashed ${dragOver ? C.accent : C.lineStrong}`,
                  borderRadius: 8,
                  padding: "26px 16px",
                  textAlign: "center",
                  cursor: "pointer",
                  background: dragOver ? C.accentSoft : C.surfaceSoft,
                  transition: "background 120ms, border-color 120ms",
                }}
              >
                <UploadCloud size={26} color={dragOver ? C.accentDeep : C.inkFaint} style={{ marginBottom: 8 }} />
                <div style={{ fontSize: 14.5, fontWeight: 600 }}>Suelta uno o más archivos .json</div>
                <div style={{ fontSize: 13, color: C.inkFaint, marginTop: 2 }}>o haz clic para elegirlos</div>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".json,application/json"
                  multiple
                  style={{ display: "none" }}
                  onChange={(e) => {
                    handleFiles(e.target.files);
                    e.target.value = "";
                  }}
                />
              </div>

              <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                <button
                  onClick={() => setPasteOpen((v) => !v)}
                  style={{
                    flex: 1,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    gap: 6,
                    padding: "8px 10px",
                    borderRadius: 7,
                    border: `1px solid ${C.line}`,
                    background: C.surface,
                    color: C.inkSoft,
                    fontSize: 13.5,
                    cursor: "pointer",
                  }}
                >
                  <ClipboardPaste size={15} /> Pegar JSON
                </button>
                <button
                  onClick={loadEjemplos}
                  style={{
                    flex: 1,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    gap: 6,
                    padding: "8px 10px",
                    borderRadius: 7,
                    border: `1px solid ${C.line}`,
                    background: C.surface,
                    color: C.inkSoft,
                    fontSize: 13.5,
                    cursor: "pointer",
                  }}
                >
                  <FlaskConical size={15} /> Probar con ejemplo
                </button>
              </div>

              {pasteOpen && (
                <div style={{ marginTop: 12 }}>
                  <textarea
                    value={pasteText}
                    onChange={(e) => setPasteText(e.target.value)}
                    placeholder='Pega aquí el contenido JSON, por ejemplo: { "identificacion": { ... } }'
                    rows={6}
                    style={{
                      width: "100%",
                      border: `1px solid ${C.line}`,
                      borderRadius: 7,
                      padding: 10,
                      fontSize: 12.5,
                      fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace",
                      resize: "vertical",
                      boxSizing: "border-box",
                    }}
                  />
                  {pasteError && (
                    <div style={{ color: C.danger, fontSize: 12.5, marginTop: 4 }}>{pasteError}</div>
                  )}
                  <button
                    onClick={submitPaste}
                    style={{
                      marginTop: 8,
                      width: "100%",
                      padding: "8px 10px",
                      borderRadius: 7,
                      border: "none",
                      background: C.ink,
                      color: C.surface,
                      fontSize: 13.5,
                      cursor: "pointer",
                    }}
                  >
                    Agregar a la lista
                  </button>
                </div>
              )}
            </div>

            {/* Lista de archivos */}
            <div style={{ ...card, padding: 16 }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                <span style={{ fontSize: 13.5, fontWeight: 700, color: C.inkSoft }}>
                  Documentos cargados ({entries.length})
                </span>
                {entries.length > 0 && (
                  <button
                    onClick={clearAll}
                    style={{ border: "none", background: "none", color: C.inkFaint, fontSize: 12.5, cursor: "pointer", display: "flex", alignItems: "center", gap: 4 }}
                  >
                    <Trash2 size={13} /> Vaciar
                  </button>
                )}
              </div>

              {entries.length === 0 ? (
                <div style={{ fontSize: 13, color: C.inkFaint, padding: "10px 2px" }}>
                  Aún no has agregado ningún archivo. Arrastra tus DTE en formato JSON arriba.
                </div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 6, maxHeight: 280, overflowY: "auto" }}>
                  {entries.map((e) => {
                    const bucket = e.status === "ok" ? e.doc.bucket : null;
                    const style = bucket ? BUCKET_STYLE[bucket] : { bg: C.dangerTint, fg: C.danger };
                    return (
                      <div
                        key={e.id}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 8,
                          padding: "8px 9px",
                          borderRadius: 7,
                          background: C.surfaceSoft,
                          border: `1px solid ${C.line}`,
                        }}
                      >
                        {e.status === "ok" ? (
                          <CheckCircle2 size={15} color={C.success} style={{ flexShrink: 0 }} />
                        ) : (
                          <AlertTriangle size={15} color={C.danger} style={{ flexShrink: 0 }} />
                        )}
                        <div style={{ minWidth: 0, flex: 1 }}>
                          <div style={{ fontSize: 12.5, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                            {e.name}
                          </div>
                          <div style={{ fontSize: 11, color: e.status === "ok" ? C.inkFaint : C.danger }}>
                            {e.status === "ok" ? e.doc.tipoLabel || "JSON genérico" : e.errorMsg}
                          </div>
                        </div>
                        {bucket && (
                          <span
                            style={{
                              fontSize: 10.5,
                              padding: "2px 7px",
                              borderRadius: 999,
                              background: style.bg,
                              color: style.fg,
                              whiteSpace: "nowrap",
                              flexShrink: 0,
                            }}
                          >
                            {bucket}
                          </span>
                        )}
                        <button
                          onClick={() => removeEntry(e.id)}
                          style={{ border: "none", background: "none", cursor: "pointer", color: C.inkFaint, flexShrink: 0, padding: 2 }}
                        >
                          <X size={14} />
                        </button>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            {/* Opciones */}
            <div style={{ ...card, padding: 16 }}>
              <div style={{ fontSize: 13.5, fontWeight: 700, color: C.inkSoft, marginBottom: 10 }}>
                Cómo organizar el Excel
              </div>
              <label style={{ display: "flex", alignItems: "flex-start", gap: 9, fontSize: 13, marginBottom: 12, cursor: "pointer" }}>
                <input
                  type="checkbox"
                  checked={separarPorTipo}
                  onChange={(e) => setSepararPorTipo(e.target.checked)}
                  style={{ marginTop: 2 }}
                />
                <span>
                  Separar por tipo de documento
                  <div style={{ color: C.inkFaint, fontSize: 12 }}>
                    Una hoja de resumen para Crédito Fiscal y otra para Consumidor Final.
                  </div>
                </span>
              </label>
              <label style={{ display: "flex", alignItems: "flex-start", gap: 9, fontSize: 13, cursor: "pointer" }}>
                <input
                  type="checkbox"
                  checked={incluirDetalle}
                  onChange={(e) => setIncluirDetalle(e.target.checked)}
                  style={{ marginTop: 2 }}
                />
                <span>
                  Incluir detalle de productos
                  <div style={{ color: C.inkFaint, fontSize: 12 }}>
                    Agrega hoja(s) con cada línea de producto o servicio.
                  </div>
                </span>
              </label>
            </div>

            <button
              onClick={download}
              disabled={!docs.length}
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                gap: 8,
                padding: "13px 16px",
                borderRadius: 9,
                border: "none",
                background: docs.length ? C.accent : C.greyTint,
                color: docs.length ? "#2A1C05" : C.inkFaint,
                fontSize: 15,
                fontWeight: 700,
                cursor: docs.length ? "pointer" : "not-allowed",
              }}
            >
              <Download size={18} /> Descargar Excel
            </button>
          </div>

          {/* Columna derecha: resumen + vista previa */}
          <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
            {/* Tarjetas de estadísticas */}
            <div className="stats-grid">
              {[
                { label: "Documentos", value: dteDocs.length, tint: C.surface },
                { label: "Crédito Fiscal", value: stats.cf, tint: C.navyTint, fg: C.navyText },
                { label: "Consumidor Final", value: stats.cons, tint: C.goldTint, fg: C.goldText },
                { label: "Total operado", value: `$${money(stats.totalMonto)}`, tint: C.successTint, fg: C.success },
              ].map((s) => (
                <div key={s.label} style={{ ...card, padding: "14px 16px", background: s.tint }}>
                  <div style={{ fontSize: 12, color: s.fg || C.inkFaint, fontWeight: 600 }}>{s.label}</div>
                  <div style={{ fontFamily: serif, fontSize: 22, marginTop: 3, color: s.fg || C.ink }}>{s.value}</div>
                </div>
              ))}
            </div>

            {errorEntries.length > 0 && (
              <div style={{ ...card, padding: 12, background: C.dangerTint, borderColor: "transparent", display: "flex", gap: 8, alignItems: "flex-start" }}>
                <AlertTriangle size={16} color={C.danger} style={{ marginTop: 1, flexShrink: 0 }} />
                <div style={{ fontSize: 13, color: C.danger }}>
                  {errorEntries.length === 1
                    ? "1 archivo no se pudo leer y no se incluirá en el Excel."
                    : `${errorEntries.length} archivos no se pudieron leer y no se incluirán en el Excel.`}
                </div>
              </div>
            )}

            {/* Vista previa */}
            <div style={{ ...card, overflow: "hidden" }}>
              <div style={{ display: "flex", borderBottom: `1px solid ${C.line}` }}>
                {[
                  { id: "resumen", label: "Resumen" },
                  { id: "detalle", label: "Detalle de items" },
                ].map((t) => (
                  <button
                    key={t.id}
                    onClick={() => setActiveTab(t.id)}
                    style={{
                      padding: "12px 18px",
                      border: "none",
                      background: "none",
                      cursor: "pointer",
                      fontSize: 13.5,
                      fontWeight: 700,
                      color: activeTab === t.id ? C.ink : C.inkFaint,
                      borderBottom: activeTab === t.id ? `2px solid ${C.accent}` : "2px solid transparent",
                    }}
                  >
                    {t.label}
                  </button>
                ))}
              </div>

              {dteDocs.length === 0 ? (
                <div style={{ padding: "48px 20px", textAlign: "center", color: C.inkFaint }}>
                  <FileJson2 size={30} style={{ marginBottom: 10, opacity: 0.6 }} />
                  <div style={{ fontSize: 14 }}>
                    Carga tus documentos JSON o prueba con el ejemplo para ver aquí una vista previa.
                  </div>
                </div>
              ) : activeTab === "resumen" ? (
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                    <thead>
                      <tr style={{ background: C.surfaceSoft }}>
                        {["Tipo", "N° control", "Emisor", "Receptor", "Total"].map((h) => (
                          <th key={h} style={{ textAlign: "left", padding: "9px 12px", color: C.inkSoft, fontWeight: 700, borderBottom: `1px solid ${C.line}`, whiteSpace: "nowrap" }}>
                            {h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {previewRows.map((d, i) => (
                        <tr key={i} style={{ borderBottom: `1px solid ${C.line}` }}>
                          <td style={{ padding: "9px 12px" }}>
                            <span style={{ fontSize: 10.5, padding: "2px 7px", borderRadius: 999, background: BUCKET_STYLE[d.bucket].bg, color: BUCKET_STYLE[d.bucket].fg }}>
                              {d.bucket}
                            </span>
                          </td>
                          <td style={{ padding: "9px 12px", color: C.inkSoft, fontFamily: "ui-monospace, monospace", fontSize: 11.5 }}>{d.numeroControl || "—"}</td>
                          <td style={{ padding: "9px 12px" }}>{d.emisorNombre || "—"}</td>
                          <td style={{ padding: "9px 12px" }}>{d.receptorNombre || "—"}</td>
                          <td style={{ padding: "9px 12px", textAlign: "right", fontVariantNumeric: "tabular-nums" }}>${money(d.total)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {dteDocs.length > previewRows.length && (
                    <div style={{ padding: "8px 12px", fontSize: 12, color: C.inkFaint }}>
                      Mostrando {previewRows.length} de {dteDocs.length} documentos. El Excel incluirá todos.
                    </div>
                  )}
                </div>
              ) : (
                <div style={{ overflowX: "auto" }}>
                  {stats.totalItems === 0 ? (
                    <div style={{ padding: "36px 20px", textAlign: "center", color: C.inkFaint, fontSize: 13.5 }}>
                      Ninguno de los documentos cargados trae líneas de producto (cuerpoDocumento).
                    </div>
                  ) : (
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                      <thead>
                        <tr style={{ background: C.surfaceSoft }}>
                          {["Descripción", "Cant.", "Precio unit.", "Venta gravada"].map((h) => (
                            <th key={h} style={{ textAlign: h === "Descripción" ? "left" : "right", padding: "9px 12px", color: C.inkSoft, fontWeight: 700, borderBottom: `1px solid ${C.line}` }}>
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {previewRows.map((it, i) => (
                          <tr key={i} style={{ borderBottom: `1px solid ${C.line}` }}>
                            <td style={{ padding: "9px 12px" }}>{it.descripcion || "—"}</td>
                            <td style={{ padding: "9px 12px", textAlign: "right", fontVariantNumeric: "tabular-nums" }}>{it.cantidad}</td>
                            <td style={{ padding: "9px 12px", textAlign: "right", fontVariantNumeric: "tabular-nums" }}>${money(it.precioUni)}</td>
                            <td style={{ padding: "9px 12px", textAlign: "right", fontVariantNumeric: "tabular-nums" }}>${money(it.ventaGravada)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                  {stats.totalItems > previewRows.length && (
                    <div style={{ padding: "8px 12px", fontSize: 12, color: C.inkFaint }}>
                      Mostrando {previewRows.length} de {stats.totalItems} líneas. El Excel incluirá todas.
                    </div>
                  )}
                </div>
              )}
            </div>

            <p style={{ fontSize: 12, color: C.inkFaint, lineHeight: 1.6, margin: 0 }}>
              Todo se procesa en tu navegador: los archivos no se envían a ningún servidor. Si un JSON no
              tiene la estructura de un DTE, igual se incluye en una hoja aparte ("JSON genérico") con sus
              datos aplanados, para que nada se pierda.
            </p>
          </div>
        </div>
      </main>
    </div>
  );
}
