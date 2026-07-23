/**
 * parseBrou.ts
 *
 * Limpia un estado de cuenta de BROU (.xls) y devuelve únicamente los
 * movimientos correspondientes a una fecha dada, con las columnas que
 * nos importan para la planilla de bancos.
 *
 * Formato de origen (filas típicas, puede variar un poco de fila pero
 * la cabecera "Fecha | Descripción | | Número de documento | Asunto |
 * Dependencia | Débito | Crédito" siempre está presente):
 *
 *   Fecha       Descripción                 Número doc.        Asunto                                   Dependencia               Débito    Crédito
 *   46223       TRF SPI PAGO PROV.          NBC76584065        COME-Pago de fac. 5-ASOCIACION ...        199 - Casa Matriz                    427.0
 */

import * as XLSX from "xlsx";
import { identificarCuenta } from "./accountIdentifier";
import type { MovimientoLimpio } from "./types";

const ENCABEZADO_ESPERADO = ["Fecha", "Descripción"];

function hojaAMatriz(sheet: XLSX.WorkSheet): unknown[][] {
  // cellDates:false porque acá queremos el número de serie crudo y lo convertimos nosotros;
  // así el comportamiento es idéntico venga o no seteado cellDates en el readFile original.
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" }) as unknown[][];
}

/** Excel guarda las fechas como "número de serie" (días desde 1899-12-30). */
function serialAFecha(serial: number): Date {
  const epoch = new Date(Date.UTC(1899, 11, 30));
  const ms = serial * 24 * 60 * 60 * 1000;
  return new Date(epoch.getTime() + ms);
}

function fechaATextoDDMMYYYY(d: Date): string {
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const yyyy = d.getUTCFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function mismodia(a: Date, b: Date): boolean {
  return (
    a.getUTCFullYear() === b.getUTCFullYear() &&
    a.getUTCMonth() === b.getUTCMonth() &&
    a.getUTCDate() === b.getUTCDate()
  );
}

function aNumero(v: unknown): number {
  if (v === "" || v === undefined || v === null) return 0;
  if (typeof v === "number") return v;
  const n = Number(String(v).replace(/\./g, "").replace(",", "."));
  return Number.isNaN(n) ? 0 : n;
}

/**
 * @param rutaOWorkbook ruta al .xls, o un WorkBook de SheetJS ya leído
 * @param fechaObjetivo la fecha de la que queremos quedarnos con los movimientos
 */
export function parseBrou(
  rutaOWorkbook: string | XLSX.WorkBook,
  fechaObjetivo: Date
): MovimientoLimpio[] {
  const workbook =
    typeof rutaOWorkbook === "string" ? XLSX.readFile(rutaOWorkbook) : rutaOWorkbook;

  const cuenta = identificarCuenta(workbook);

  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const matriz = hojaAMatriz(sheet);

  // 1. Encontrar la fila de encabezado de la tabla de movimientos.
  const idxEncabezado = matriz.findIndex(
    (fila) =>
      String(fila[0]).trim() === ENCABEZADO_ESPERADO[0] &&
      String(fila[1]).trim() === ENCABEZADO_ESPERADO[1]
  );
  if (idxEncabezado === -1) {
    throw new Error(
      "No se encontró la fila de encabezado 'Fecha | Descripción' en el archivo BROU. ¿Cambió el formato?"
    );
  }

  const movimientos: MovimientoLimpio[] = [];

  // 2. Recorrer filas de datos hasta encontrar una fila vacía (fin de la tabla).
  for (let i = idxEncabezado + 1; i < matriz.length; i++) {
    const fila = matriz[i];
    const fechaRaw = fila[0];

    // Fin de la tabla de movimientos: fila sin fecha.
    if (fechaRaw === "" || fechaRaw === undefined) break;

    const fechaSerial = typeof fechaRaw === "number" ? fechaRaw : aNumero(fechaRaw);
    const fecha = serialAFecha(fechaSerial);

    if (!mismodia(fecha, fechaObjetivo)) continue;

    const descripcion = String(fila[1] ?? "").trim();
    const numeroDocumento = String(fila[3] ?? "").trim();
    const asunto = String(fila[4] ?? "").trim();
    const dependencia = String(fila[5] ?? "").trim();
    const debito = aNumero(fila[6]);
    const credito = aNumero(fila[7]);

    if (debito === 0 && credito === 0) continue; // fila rara sin importe, la salteamos

    movimientos.push({
      fecha,
      fechaTexto: fechaATextoDDMMYYYY(fecha),
      banco: "BROU",
      cuentaKey: cuenta.cuentaKey,
      tipo: debito !== 0 ? "debito" : "credito",
      monto: debito !== 0 ? debito : credito,
      descripcion,
      // En BROU el cliente casi siempre está en "Asunto"; si viene vacío,
      // usamos la descripción como respaldo para que la IA igual tenga algo para cotejar.
      textoParaMatchCliente: asunto || descripcion,
      referencia: numeroDocumento || undefined,
      categoriaBanco: dependencia || undefined,
      filaOriginal: i + 1, // 1-based, como lo vería un humano abriendo el Excel
    });
  }

  return movimientos;
}

// --- Uso directo por consola: `ts-node parseBrou.ts archivo.xls 20/07/2026` ---
if (require.main === module) {
  const [archivo, fechaTexto] = process.argv.slice(2);
  if (!archivo || !fechaTexto) {
    console.log("Uso: ts-node parseBrou.ts <archivo.xls> <dd/mm/yyyy>");
    process.exit(1);
  }
  const [dd, mm, yyyy] = fechaTexto.split("/").map(Number);
  const fechaObjetivo = new Date(Date.UTC(yyyy, mm - 1, dd));
  const resultado = parseBrou(archivo, fechaObjetivo);
  console.log(JSON.stringify(resultado, null, 2));
  console.log(`\nTotal movimientos encontrados para ${fechaTexto}: ${resultado.length}`);
}
