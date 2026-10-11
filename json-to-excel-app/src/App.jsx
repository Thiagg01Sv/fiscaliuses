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
  ExternalLink,
  ShieldCheck,
} from "lucide-react";

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
  warn: "#9A6A00",
  warnTint: "#FBF0D3",
  violetTint: "#E8E3F4",
  violetText: "#4A3B80",
};

const serif = "Georgia, 'Iowan Old Style', 'Times New Roman', serif";
const sans =
  "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif";
const mono = "ui-monospace, SFMono-Regular, Menlo, monospace";

const card = {
  background: C.surface,
  border: `1px solid ${C.line}`,
  borderRadius: 10,
};

/* ----------------------------------------------------------------
   Catálogos DTE (Ministerio de Hacienda, El Salvador)
   ---------------------------------------------------------------- */
const TIPO_DTE = {
  "01": "Factura de Consumidor Final",
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

const BUCKETS = [
  "Crédito Fiscal",
  "Consumidor Final",
  "Nota de Crédito",
  "Retención IVA",
  "Sujeto Excluido",
  "Liquidación",
  "Otros DTE",
];
const BUCKET_BY_TIPO = {
  "01": "Consumidor Final",
  "03": "Crédito Fiscal",
  "05": "Nota de Crédito",
  "07": "Retención IVA",
  "08": "Liquidación",
  "09": "Liquidación",
  "14": "Sujeto Excluido",
};
function bucketFor(tipoDte) {
  return BUCKET_BY_TIPO[tipoDte] || "Otros DTE";
}
const BUCKET_STYLE = {
  "Crédito Fiscal": { bg: C.navyTint, fg: C.navyText },
  "Consumidor Final": { bg: C.goldTint, fg: C.goldText },
  "Nota de Crédito": { bg: C.dangerTint, fg: C.danger },
  "Retención IVA": { bg: C.successTint, fg: C.success },
  "Sujeto Excluido": { bg: C.accentSoft, fg: C.accentDeep },
  Liquidación: { bg: C.violetTint, fg: C.violetText },
  "Otros DTE": { bg: C.greyTint, fg: C.greyText },
  "JSON genérico": { bg: C.greyTint, fg: C.greyText },
};

const num = (v) => {
  const n = typeof v === "string" ? parseFloat(v) : v;
  return typeof n === "number" && !Number.isNaN(n) ? n : "";
};
const txt = (v) => (v === null || v === undefined ? "" : String(v));
const round2 = (n) => Math.round(n * 100) / 100;
const digits = (v) => txt(v).replace(/\D/g, "");

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

/* ----------------------------------------------------------------
   Validación local (estructura, sello, firma y cuadre de montos)
   OJO: esto NO reemplaza la consulta oficial del MH. Para eso se
   genera el enlace a la Consulta Pública de cada documento.
   ---------------------------------------------------------------- */
const UUID_RE = /^[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}$/;
const CONTROL_RE = /^DTE-(\d{2})-[A-Z0-9]{8}-\d{15}$/;
const SELLO_RE = /^[A-Z0-9]{40}$/;

function decodeJws(token) {
  try {
    const part = String(token).split(".")[1];
    const b64 = part.replace(/-/g, "+").replace(/_/g, "/");
    const bin = atob(b64 + "=".repeat((4 - (b64.length % 4)) % 4));
    const bytes = Uint8Array.from(bin, (c) => c.charCodeAt(0));
    return JSON.parse(new TextDecoder().decode(bytes));
  } catch (e) {
    return null;
  }
}

// Compara estructuras; en textos solo mira letras y números (evita falsos
// errores por tildes dañadas, ej. "RETENCI?N" vs "RETENCIÓN").
function deepEq(a, b) {
  if (a === b) return true;
  if (typeof a === "string" && typeof b === "string") {
    const clean = (s) => s.replace(/[^A-Za-z0-9]/g, "");
    return clean(a) === clean(b);
  }
  if (!a || !b || typeof a !== "object" || typeof b !== "object") return false;
  const ka = Object.keys(a);
  const kb = Object.keys(b);
  return ka.length === kb.length && ka.every((k) => deepEq(a[k], b[k]));
}

const close = (a, b, tol = 0.011) => Math.abs(a - b) <= tol;
const sumBy = (arr, k) =>
  round2(arr.reduce((s, x) => s + (typeof num(x[k]) === "number" ? num(x[k]) : 0), 0));

function validateDte(json, tipoDte, bodyArr) {
  const out = [];
  const add = (level, msg) => out.push({ level, msg });
  const id = json.identificacion || {};
  const em = json.emisor || {};
  const rs = json.resumen || {};
  const rec = json.receptor || json.sujetoExcluido || {};

  if (!TIPO_DTE[tipoDte]) add("warn", `Tipo de DTE no reconocido (${tipoDte || "vacío"}).`);

  if (!UUID_RE.test(txt(id.codigoGeneracion)))
    add("error", "El código de generación no tiene formato UUID en mayúsculas.");

  const m = CONTROL_RE.exec(txt(id.numeroControl));
  if (!m) add("error", "El número de control no tiene el formato DTE-TT-XXXXXXXX-000000000000000.");
  else if (m[1] !== tipoDte)
    add("error", `El número de control es del tipo ${m[1]} pero tipoDte dice ${tipoDte}.`);

  const hoy = new Date().toLocaleDateString("sv-SE");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(txt(id.fecEmi))) add("error", "Fecha de emisión inválida (se espera AAAA-MM-DD).");
  else if (id.fecEmi > hoy) add("warn", `La fecha de emisión (${id.fecEmi}) es posterior a hoy.`);

  if (id.ambiente === "00") add("warn", "Documento del ambiente de pruebas (00); no es válido fiscalmente.");
  else if (id.ambiente !== "01") add("warn", `Ambiente desconocido (${txt(id.ambiente)}).`);

  const sello = json.selloRecibido || json.sello_recibido;
  if (!sello) add("error", "No trae sello de recepción: no hay evidencia de que el MH lo haya procesado.");
  else if (!SELLO_RE.test(txt(sello))) add("warn", "El sello de recepción no tiene 40 caracteres alfanuméricos.");

  if (digits(em.nit).length !== 14) add("error", "El NIT del emisor no tiene 14 dígitos.");
  if (rec.tipoDocumento === "36" && digits(rec.numDocumento).length !== 14)
    add("warn", "El receptor indica tipo de documento NIT (36) pero el número no tiene 14 dígitos.");
  if (id.tipoMoneda && id.tipoMoneda !== "USD") add("warn", `Moneda distinta de USD (${id.tipoMoneda}).`);

  // Firma electrónica: se decodifica el contenido firmado y se contrasta
  // con el JSON. No se verifica la firma criptográfica (requiere el
  // certificado del emisor).
  if (!json.firmaElectronica) {
    add("warn", "No incluye firmaElectronica: no se puede contrastar el contenido firmado.");
  } else {
    const p = decodeJws(json.firmaElectronica);
    if (!p) add("warn", "No se pudo decodificar la firma electrónica.");
    else {
      const pid = p.identificacion || {};
      ["codigoGeneracion", "numeroControl", "fecEmi", "horEmi"].forEach((k) => {
        if (pid[k] !== id[k]) add("error", `El contenido firmado trae otro valor en ${k}.`);
      });
      if (digits(p.emisor && p.emisor.nit) !== digits(em.nit))
        add("error", "El NIT del emisor difiere del contenido firmado.");
      if (!deepEq(p.resumen, json.resumen))
        add("error", "El resumen del JSON no coincide con el contenido firmado (posible alteración).");
    }
  }

  // Cuadre de montos
  if (tipoDte === "07") {
    if (!bodyArr.length) add("error", "El comprobante de retención no tiene líneas en cuerpoDocumento.");
    const sIva = sumBy(bodyArr, "ivaRetenido");
    const sMonto = sumBy(bodyArr, "montoSujetoGrav");
    if (typeof rs.totalIVAretenido === "number" && !close(sIva, rs.totalIVAretenido))
      add("error", `La suma del IVA retenido por línea (${sIva}) no coincide con el resumen (${rs.totalIVAretenido}).`);
    if (typeof rs.totalSujetoRetencion === "number" && !close(sMonto, rs.totalSujetoRetencion))
      add("error", `La suma del monto sujeto por línea (${sMonto}) no coincide con el resumen (${rs.totalSujetoRetencion}).`);
    bodyArr.forEach((it, i) => {
      if (
        txt(it.codigoRetencionMH) === "22" &&
        typeof it.montoSujetoGrav === "number" &&
        typeof it.ivaRetenido === "number"
      ) {
        const esperado = round2(it.montoSujetoGrav * 0.01);
        if (!close(esperado, it.ivaRetenido))
          add("warn", `Línea ${it.numItem ?? i + 1}: el 1% de ${it.montoSujetoGrav} es ${esperado}, pero se retuvo ${it.ivaRetenido}.`);
      }
    });
  }

  if (["01", "03", "05"].includes(tipoDte) && bodyArr.length && typeof rs.totalGravada === "number") {
    const s = sumBy(bodyArr, "ventaGravada");
    if (!close(s, rs.totalGravada, 0.02))
      add("warn", `La suma de ventas gravadas por línea (${s}) no coincide con el resumen (${rs.totalGravada}).`);
  }
  if (["03", "05"].includes(tipoDte) && typeof rs.totalGravada === "number" && Array.isArray(rs.tributos)) {
    const t = rs.tributos.find((x) => x.codigo === "20");
    if (t && typeof t.valor === "number") {
      const base = rs.totalGravada - (typeof rs.descuGravada === "number" ? rs.descuGravada : 0);
      if (!close(round2(base * 0.13), t.valor, 0.03))
        add("warn", `El IVA (${t.valor}) no equivale al 13% de la base gravada (${round2(base)}). Revisar.`);
    }
  }
  if (tipoDte === "14" && bodyArr.length && typeof rs.totalCompra === "number") {
    const s = sumBy(bodyArr, "compra");
    if (!close(s, rs.totalCompra, 0.02))
      add("warn", `La suma de compras por línea (${s}) no coincide con el total de compra (${rs.totalCompra}).`);
  }

  return out;
}

const VAL_LABEL = { ok: "Sin observaciones", warn: "Con advertencias", error: "Con errores" };
function statusOf(issues) {
  if (issues.some((i) => i.level === "error")) return "error";
  if (issues.length) return "warn";
  return "ok";
}
const VAL_STYLE = {
  ok: { bg: C.successTint, fg: C.success },
  warn: { bg: C.warnTint, fg: C.warn },
  error: { bg: C.dangerTint, fg: C.danger },
};

/* ----------------------------------------------------------------
   Normalización
   ---------------------------------------------------------------- */
function normalizeDte(json, filename) {
  const ident = json.identificacion || {};
  const emisor = json.emisor || {};
  const receptor = json.receptor || json.sujetoExcluido || {};
  const resumen = json.resumen || {};
  const raw = json.cuerpoDocumento;
  // El Documento Contable de Liquidación (09) trae cuerpoDocumento como objeto
  const bodyArr = Array.isArray(raw) ? raw : raw && typeof raw === "object" ? [raw] : [];
  const b0 = bodyArr[0] || {};
  const tipoDte = String(ident.tipoDte ?? "");
  const bucket = bucketFor(tipoDte);
  const isRet = tipoDte === "07";
  const isDcl = tipoDte === "09";

  let iva = "";
  if (typeof resumen.totalIva === "number") iva = resumen.totalIva;
  else if (Array.isArray(resumen.tributos)) {
    const t = resumen.tributos.find((t) => t.codigo === "20");
    if (t) iva = t.valor;
  } else if (isDcl) iva = b0.iva;

  const issues = validateDte(json, tipoDte, bodyArr);
  const status = statusOf(issues);
  const urlMH =
    ident.codigoGeneracion && ident.fecEmi
      ? `https://admin.factura.gob.sv/consultaPublica?ambiente=${ident.ambiente || "01"}&codGen=${ident.codigoGeneracion}&fechaEmi=${ident.fecEmi}`
      : "";

  const items = bodyArr.map((it, idx) => {
    if (isDcl) {
      return {
        archivo: filename,
        numeroControl: ident.numeroControl || "",
        numItem: idx + 1,
        codigo: txt(it.codLiqui),
        descripcion: `Liquidación ${txt(it.codLiqui)} (${txt(it.periodoLiquidacionFechaInicio)} a ${txt(it.periodoLiquidacionFechaFin)})`,
        cantidad: num(it.cantidadDoc),
        unidad: "",
        precioUni: "",
        descu: "",
        ventaNoSuj: "",
        ventaExenta: "",
        ventaGravada: num(it.valorOperaciones),
        docRel: "",
        fechaDocRel: "",
        montoSujeto: num(it.montoSujetoPercepcion),
        ivaRetItem: "",
        codRetencion: "",
      };
    }
    return {
      archivo: filename,
      numeroControl: ident.numeroControl || "",
      numItem: it.numItem ?? idx + 1,
      codigo: it.codigo || "",
      descripcion: it.descripcion || it.obsItem || "",
      cantidad: num(it.cantidad),
      unidad: it.uniMedida ?? it.unidadMedida ?? "",
      precioUni: num(it.precioUni),
      descu: num(it.montoDescu),
      ventaNoSuj: num(it.ventaNoSuj),
      ventaExenta: num(it.ventaExenta),
      ventaGravada: num(it.ventaGravada ?? it.compra),
      docRel: txt(it.numDocumento ?? it.numeroDocumento),
      fechaDocRel: txt(it.fechaEmision ?? it.fechaGeneracion),
      montoSujeto: num(it.montoSujetoGrav),
      ivaRetItem: num(it.ivaRetenido),
      codRetencion: txt(it.codigoRetencionMH),
    };
  });

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
    emisorNit: txt(emisor.nit),
    emisorNrc: txt(emisor.nrc),
    receptorNombre:
      receptor.nombre || (bucket === "Consumidor Final" ? "Consumidor Final" : ""),
    receptorDoc: txt(receptor.nit || receptor.numDocumento || receptor.dui),
    ventaGravada: isRet
      ? ""
      : num(isDcl ? b0.valorOperaciones : resumen.totalGravada ?? resumen.totalCompra ?? resumen.subTotal),
    ventaExenta: num(resumen.totalExenta),
    ventaNoSuj: num(resumen.totalNoSuj),
    descuento: num(resumen.totalDescu),
    iva: num(iva),
    montoSujeto: num(
      isRet ? resumen.totalSujetoRetencion : isDcl ? b0.montoSujetoPercepcion : resumen.totalSujetoRetencion
    ),
    ivaRetenido: num(isRet ? resumen.totalIVAretenido : resumen.ivaRete1 ?? resumen.totalIVAretenido),
    rentaRetenida: num(resumen.reteRenta),
    total: isRet
      ? ""
      : num(
          isDcl
            ? b0.liquidoApagar
            : resumen.montoTotalOperacion ?? resumen.totalPagar ?? resumen.total ?? resumen.totalCompra
        ),
    condicionOperacion: CONDICION_OP[resumen.condicionOperacion] || "",
    sello: json.selloRecibido || json.sello_recibido || "",
    nDocs: bodyArr.length,
    issues,
    valEstado: VAL_LABEL[status],
    valStatus: status,
    valDetalle: issues
      .map((i) => `${i.level === "error" ? "ERROR" : "Aviso"}: ${i.msg}`)
      .join(" | "),
    urlMH,
    items,
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
  ["fechaEmision", "Fecha de emisión"],
  ["horaEmision", "Hora de emisión"],
  ["codigoGeneracion", "Código de generación"],
  ["sello", "Sello de recepción"],
  ["numeroControl", "Número de control"],
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
  ["montoSujeto", "Monto sujeto"],
  ["ivaRetenido", "IVA retenido (1%)"],
  ["rentaRetenida", "Renta retenida"],
  ["total", "Total a pagar"],
  ["condicionOperacion", "Condición de operación"],
  ["valEstado", "Validación local"],
  ["valDetalle", "Detalle de validación"],
  ["urlMH", "Consulta pública MH"],
];
const RET_COLS = [
  ["archivo", "Archivo"],
  ["fechaEmision", "Fecha de emisión"],
  ["codigoGeneracion", "Código de generación"],
  ["sello", "Sello de recepción"],
  ["numeroControl", "Número de control"],
  ["emisorNit", "NIT emisor (agente de retención)"],
  ["emisorNombre", "Emisor"],
  ["receptorDoc", "NIT receptor (sujeto retenido)"],
  ["receptorNombre", "Receptor"],
  ["nDocs", "N° docs. relacionados"],
  ["montoSujeto", "Monto sujeto a retención"],
  ["ivaRetenido", "IVA retenido (1%)"],
  ["valEstado", "Validación local"],
  ["valDetalle", "Detalle de validación"],
  ["urlMH", "Consulta pública MH"],
];
const MONEY_KEYS = new Set([
  "ventaGravada",
  "ventaExenta",
  "ventaNoSuj",
  "descuento",
  "iva",
  "montoSujeto",
  "ivaRetenido",
  "rentaRetenida",
  "total",
  "precioUni",
  "descu",
  "ivaRetItem",
]);
const ITEM_COLS = [
  ["archivo", "Archivo"],
  ["numeroControl", "Número de control"],
  ["numItem", "N° item"],
  ["codigo", "Código"],
  ["descripcion", "Descripción"],
  ["docRel", "Documento relacionado"],
  ["fechaDocRel", "Fecha doc. relacionado"],
  ["cantidad", "Cantidad"],
  ["unidad", "Unidad"],
  ["precioUni", "Precio unitario"],
  ["descu", "Descuento"],
  ["ventaNoSuj", "Venta no sujeta"],
  ["ventaExenta", "Venta exenta"],
  ["ventaGravada", "Venta gravada"],
  ["montoSujeto", "Monto sujeto"],
  ["ivaRetItem", "IVA retenido"],
  ["codRetencion", "Código retención MH"],
];
const VAL_COLS = [
  ["archivo", "Archivo"],
  ["numeroControl", "Número de control"],
  ["codigoGeneracion", "Código de generación"],
  ["valEstado", "Estado"],
  ["nivel", "Nivel"],
  ["mensaje", "Observación"],
];

function sheetFromCols(rows, cols) {
  const header = cols.map(([, label]) => label);
  const aoa = [header, ...rows.map((r) => cols.map(([key]) => r[key] ?? ""))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  cols.forEach(([key], c) => {
    for (let r = 1; r <= rows.length; r++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (key === "urlMH") {
        const url = rows[r - 1].urlMH;
        if (url) ws[addr] = { t: "s", v: "Verificar en MH", l: { Target: url, Tooltip: url } };
      } else if (MONEY_KEYS.has(key) && cell && typeof cell.v === "number") {
        cell.z = "#,##0.00";
      }
    }
  });
  // Ancho según el contenido, para que los códigos largos salgan completos
  ws["!cols"] = cols.map(([key, label]) => {
    if (key === "urlMH") return { wch: 18 };
    const longest = rows.reduce(
      (m, r) => Math.max(m, String(r[key] ?? "").length),
      label.length
    );
    return { wch: Math.min(70, Math.max(12, longest + 2)) };
  });
  if (rows.length) {
    ws["!autofilter"] = {
      ref: XLSX.utils.encode_range({ r: 0, c: 0 }, { r: rows.length, c: cols.length - 1 }),
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
      BUCKETS.forEach((b) => {
        const rows = dte.filter((d) => d.bucket === b);
        if (rows.length)
          addSheet(wb, `Resumen - ${b}`, sheetFromCols(rows, b === "Retención IVA" ? RET_COLS : RESUMEN_COLS));
      });
    } else {
      addSheet(wb, "Resumen DTE", sheetFromCols(dte, RESUMEN_COLS));
      const ret = dte.filter((d) => d.bucket === "Retención IVA");
      if (ret.length) addSheet(wb, "Retenciones IVA", sheetFromCols(ret, RET_COLS));
    }

    if (incluirDetalle) {
      if (separarPorTipo) {
        BUCKETS.forEach((b) => {
          const items = dte.filter((d) => d.bucket === b).flatMap((d) => d.items);
          if (items.length) addSheet(wb, `Detalle - ${b}`, sheetFromCols(items, ITEM_COLS));
        });
      } else {
        const items = dte.flatMap((d) => d.items);
        if (items.length) addSheet(wb, "Detalle de items", sheetFromCols(items, ITEM_COLS));
      }
    }

    const valRows = dte.flatMap((d) =>
      d.issues.length
        ? d.issues.map((i) => ({
            archivo: d.archivo,
            numeroControl: d.numeroControl,
            codigoGeneracion: d.codigoGeneracion,
            valEstado: d.valEstado,
            nivel: i.level === "error" ? "Error" : "Aviso",
            mensaje: i.msg,
          }))
        : [
            {
              archivo: d.archivo,
              numeroControl: d.numeroControl,
              codigoGeneracion: d.codigoGeneracion,
              valEstado: d.valEstado,
              nivel: "",
              mensaje: "Sin observaciones",
            },
          ]
    );
    addSheet(wb, "Validación", sheetFromCols(valRows, VAL_COLS));
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
  {
    filename: "ejemplo-retencion-iva.json",
    json: {
      selloRecibido: "2026AB123CD456EF789GH012IJ345KL678MN901OP",
      identificacion: {
        ambiente: "01",
        tipoDte: "07",
        numeroControl: "DTE-07-M001P002-000000000000010",
        codigoGeneracion: "7C1D2E3F-0000-4E00-9A00-000000000007",
        fecEmi: "2026-09-04",
        horEmi: "14:12:43",
        tipoMoneda: "USD",
      },
      emisor: { nombre: "Aseguradora Ejemplo, S.A.", nit: "06140101900011", nrc: "100001" },
      receptor: { tipoDocumento: "36", numDocumento: "06140202950022", nombre: "Restaurante Ejemplo, S.A. de C.V." },
      cuerpoDocumento: [
        { numItem: 1, tipoDte: "01", tipoDoc: 2, numDocumento: "A1B2C3D4-0000-4E00-9A00-000000000099", fechaEmision: "2026-07-15", montoSujetoGrav: 142.27, codigoRetencionMH: "22", ivaRetenido: 1.42, descripcion: "RETENCIÓN 1%" },
      ],
      resumen: { totalSujetoRetencion: 142.27, totalIVAretenido: 1.42, totalIVAretenidoLetras: "(UN 42/100 DOLARES)" },
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
      const json = JSON.parse(text.replace(/^\uFEFF/, ""));
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
      JSON.parse(pasteText.replace(/^\uFEFF/, ""));
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
    const sign = (d) => (d.bucket === "Nota de Crédito" ? -1 : 1);
    const totalMonto = round2(
      dteDocs.reduce((a, d) => a + (typeof d.total === "number" ? sign(d) * d.total : 0), 0)
    );
    const montoSujeto = round2(dteDocs.reduce((a, d) => a + (typeof d.montoSujeto === "number" ? d.montoSujeto : 0), 0));
    const ivaRet = round2(dteDocs.reduce((a, d) => a + (typeof d.ivaRetenido === "number" ? d.ivaRetenido : 0), 0));
    const byBucket = {};
    dteDocs.forEach((d) => {
      byBucket[d.bucket] = (byBucket[d.bucket] || 0) + 1;
    });
    const conObs = dteDocs.filter((d) => d.valStatus !== "ok").length;
    return { totalMonto, montoSujeto, ivaRet, byBucket, conObs, totalItems: dteDocs.reduce((a, d) => a + d.items.length, 0) };
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

  const th = { textAlign: "left", padding: "9px 12px", color: C.inkSoft, fontWeight: 700, borderBottom: `1px solid ${C.line}`, whiteSpace: "nowrap" };
  const thR = { ...th, textAlign: "right" };
  const td = { padding: "9px 12px" };
  const tdMono = { ...td, color: C.inkSoft, fontFamily: mono, fontSize: 11.5, whiteSpace: "nowrap" };
  const tdNum = { ...td, textAlign: "right", fontVariantNumeric: "tabular-nums", whiteSpace: "nowrap" };

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
                Reúne tus Créditos Fiscales, Facturas, Notas de Crédito, Retenciones, Sujetos Excluidos y Liquidaciones en un solo libro, con validación.
              </p>
            </div>
          </div>
        </div>
      </header>

      <main style={{ maxWidth: 1120, margin: "0 auto", padding: "28px 24px 64px" }}>
        <div className="app-grid">
          {/* Columna izquierda */}
          <div style={{ display: "flex", flexDirection: "column", gap: 18, minWidth: 0 }}>
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
                      fontFamily: mono,
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
                    Una hoja de resumen por cada tipo: Crédito Fiscal, Consumidor Final, Nota de Crédito, Retención IVA, Sujeto Excluido y Liquidación.
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
                    Agrega hoja(s) con cada línea; en retenciones, cada documento relacionado con su monto sujeto y retención.
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

          {/* Columna derecha: resumen + validación + vista previa */}
          <div style={{ display: "flex", flexDirection: "column", gap: 18, minWidth: 0 }}>
            {/* Tarjetas de estadísticas */}
            <div className="stats-grid">
              {[
                { label: "Documentos", value: dteDocs.length, tint: C.surface },
                { label: "Total operado (NC resta)", value: `$${money(stats.totalMonto)}`, tint: C.successTint, fg: C.success },
                { label: "Monto sujeto", value: `$${money(stats.montoSujeto)}`, tint: C.navyTint, fg: C.navyText },
                { label: "IVA retenido (1%)", value: `$${money(stats.ivaRet)}`, tint: C.goldTint, fg: C.goldText },
              ].map((s) => (
                <div key={s.label} style={{ ...card, padding: "14px 16px", background: s.tint }}>
                  <div style={{ fontSize: 12, color: s.fg || C.inkFaint, fontWeight: 600 }}>{s.label}</div>
                  <div style={{ fontFamily: serif, fontSize: 22, marginTop: 3, color: s.fg || C.ink }}>{s.value}</div>
                </div>
              ))}
            </div>

            {dteDocs.length > 0 && (
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {BUCKETS.filter((b) => stats.byBucket[b]).map((b) => (
                  <span
                    key={b}
                    style={{ fontSize: 12.5, padding: "4px 11px", borderRadius: 999, background: BUCKET_STYLE[b].bg, color: BUCKET_STYLE[b].fg, fontWeight: 600 }}
                  >
                    {b}: {stats.byBucket[b]}
                  </span>
                ))}
              </div>
            )}

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

            {/* Validación */}
            {dteDocs.length > 0 && (
              <div style={{ ...card, padding: 16 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                  <ShieldCheck size={17} color={C.navyText} />
                  <span style={{ fontSize: 14, fontWeight: 700 }}>Validación</span>
                  <span style={{ fontSize: 12.5, color: stats.conObs ? C.warn : C.success, marginLeft: "auto" }}>
                    {stats.conObs
                      ? `${stats.conObs} de ${dteDocs.length} con observaciones`
                      : `${dteDocs.length} de ${dteDocs.length} sin observaciones`}
                  </span>
                </div>
                <p style={{ fontSize: 12, color: C.inkFaint, lineHeight: 1.55, margin: "0 0 10px" }}>
                  Revisa formato, sello, firma y cuadre de montos en tu navegador. La confirmación oficial la da el
                  portal de Consulta Pública de DTE del Ministerio de Hacienda: el botón abre el documento ya cargado.
                </p>
                <div style={{ display: "flex", flexDirection: "column", gap: 6, maxHeight: 260, overflowY: "auto" }}>
                  {dteDocs.map((d, i) => {
                    const vs = VAL_STYLE[d.valStatus];
                    return (
                      <div key={i} style={{ padding: "8px 10px", borderRadius: 7, background: C.surfaceSoft, border: `1px solid ${C.line}` }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          <span style={{ fontSize: 10.5, padding: "2px 8px", borderRadius: 999, background: vs.bg, color: vs.fg, whiteSpace: "nowrap", fontWeight: 600 }}>
                            {d.valEstado}
                          </span>
                          <span style={{ fontSize: 12, fontFamily: mono, color: C.inkSoft, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}>
                            {d.numeroControl || d.archivo}
                          </span>
                          {d.urlMH && (
                            <a
                              href={d.urlMH}
                              target="_blank"
                              rel="noopener noreferrer"
                              style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 12, color: C.navyText, textDecoration: "none", whiteSpace: "nowrap", fontWeight: 600 }}
                            >
                              Verificar en MH <ExternalLink size={12} />
                            </a>
                          )}
                        </div>
                        {d.issues.length > 0 && (
                          <ul style={{ margin: "6px 0 0", paddingLeft: 18, fontSize: 12, color: C.inkSoft, lineHeight: 1.5 }}>
                            {d.issues.map((it, k) => (
                              <li key={k} style={{ color: it.level === "error" ? C.danger : C.warn }}>
                                {it.msg}
                              </li>
                            ))}
                          </ul>
                        )}
                      </div>
                    );
                  })}
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
                        <th style={th}>Tipo</th>
                        <th style={th}>Fecha</th>
                        <th style={th}>Código de generación</th>
                        <th style={th}>Sello de recepción</th>
                        <th style={th}>NIT emisor</th>
                        <th style={th}>Receptor</th>
                        <th style={thR}>Monto sujeto</th>
                        <th style={thR}>IVA ret. 1%</th>
                        <th style={thR}>Total</th>
                      </tr>
                    </thead>
                    <tbody>
                      {previewRows.map((d, i) => (
                        <tr key={i} style={{ borderBottom: `1px solid ${C.line}` }}>
                          <td style={td}>
                            <span style={{ fontSize: 10.5, padding: "2px 7px", borderRadius: 999, background: BUCKET_STYLE[d.bucket].bg, color: BUCKET_STYLE[d.bucket].fg, whiteSpace: "nowrap" }}>
                              {d.bucket}
                            </span>
                          </td>
                          <td style={{ ...td, whiteSpace: "nowrap" }}>{d.fechaEmision || "—"}</td>
                          <td style={tdMono}>{d.codigoGeneracion || "—"}</td>
                          <td style={tdMono}>{d.sello || "—"}</td>
                          <td style={tdMono}>{d.emisorNit || "—"}</td>
                          <td style={{ ...td, whiteSpace: "nowrap" }}>{d.receptorNombre || "—"}</td>
                          <td style={tdNum}>{d.montoSujeto === "" ? "—" : `$${money(d.montoSujeto)}`}</td>
                          <td style={tdNum}>{d.ivaRetenido === "" ? "—" : `$${money(d.ivaRetenido)}`}</td>
                          <td style={tdNum}>{d.total === "" ? "—" : `$${money(d.total)}`}</td>
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
                      Ninguno de los documentos cargados trae líneas (cuerpoDocumento).
                    </div>
                  ) : (
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                      <thead>
                        <tr style={{ background: C.surfaceSoft }}>
                          <th style={th}>Descripción</th>
                          <th style={th}>Doc. relacionado</th>
                          <th style={thR}>Cant.</th>
                          <th style={thR}>Precio unit.</th>
                          <th style={thR}>Venta gravada</th>
                          <th style={thR}>Monto sujeto</th>
                          <th style={thR}>IVA ret.</th>
                        </tr>
                      </thead>
                      <tbody>
                        {previewRows.map((it, i) => (
                          <tr key={i} style={{ borderBottom: `1px solid ${C.line}` }}>
                            <td style={td}>{it.descripcion || "—"}</td>
                            <td style={tdMono}>{it.docRel || "—"}</td>
                            <td style={tdNum}>{it.cantidad}</td>
                            <td style={tdNum}>{it.precioUni === "" ? "—" : `$${money(it.precioUni)}`}</td>
                            <td style={tdNum}>{it.ventaGravada === "" ? "—" : `$${money(it.ventaGravada)}`}</td>
                            <td style={tdNum}>{it.montoSujeto === "" ? "—" : `$${money(it.montoSujeto)}`}</td>
                            <td style={tdNum}>{it.ivaRetItem === "" ? "—" : `$${money(it.ivaRetItem)}`}</td>
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
