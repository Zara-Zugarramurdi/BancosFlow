/**
 * parseSantander.ts
 *
 * Limpia un estado de cuenta de Santander (.xlsx) y devuelve únicamente
 * los movimientos correspondientes a una fecha dada.
 *
 * Formato de origen:
 *   Fecha        Referencia      Tipo Movimiento                     Descripción                                          Débito     Crédito    Saldo
 *   16/07/2026   TT55549584      DEBITO OPERACION EN BANCA DIGITAL   523239TT55549584 Trf. Plaza- Teledata Sa Pesos        -400000.0             581089.44
 *
 * A diferencia de BROU, acá no hay columna "Asunto" separada: el nombre
 * del cliente (cuando está) viene mezclado dentro de "Descripción".
 */

import * as XLSX from "xlsx";
import { identificarCuenta } from "./accountIdentifier";
import type { MovimientoLimpio } from "./types";

function hojaAMatriz(sheet: XLSX.WorkSheet): unknown[][] {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" }) as unknown[][];
}

function fechaATextoDDMMYYYY(d: Date): string {
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const yyyy = d.getUTCFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

/** Santander guarda la fecha como texto "dd/mm/yyyy". */
function textoAFecha(texto: string): Date | null {
  const m = texto.trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return null;
  const [, dd, mm, yyyy] = m;
  return new Date(Date.UTC(Number(yyyy), Number(mm) - 1, Number(dd)));
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
 * @param rutaOWorkbook ruta al .xlsx, o un WorkBook de SheetJS ya leído
 * @param fechaObjetivo la fecha de la que queremos quedarnos con los movimientos
 */
export function parseSantander(
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
      String(fila[0]).trim() === "Fecha" &&
      String(fila[1]).trim() === "Referencia" &&
      String(fila[2]).trim() === "Tipo Movimiento"
  );
  if (idxEncabezado === -1) {
    throw new Error(
      "No se encontró la fila de encabezado 'Fecha | Referencia | Tipo Movimiento' en el archivo Santander. ¿Cambió el formato?"
    );
  }

  const movimientos: MovimientoLimpio[] = [];

  // 2. Recorrer filas de datos hasta encontrar una fila vacía (fin de la tabla).
  for (let i = idxEncabezado + 1; i < matriz.length; i++) {
    const fila = matriz[i];
    const fechaRaw = String(fila[0] ?? "").trim();
    const descripcion = String(fila[3] ?? "").trim();

    // Fin de la tabla: fila sin fecha.
    if (fechaRaw === "" && descripcion === "") break;

    // La primera fila de datos suele ser "Saldo inicial", sin fecha: la salteamos.
    if (fechaRaw === "") continue;

    const fecha = textoAFecha(fechaRaw);
    if (!fecha) continue; // fila rara, no tiene formato de fecha reconocible

    if (!mismodia(fecha, fechaObjetivo)) continue;

    const referencia = String(fila[1] ?? "").trim();
    const tipoMovimiento = String(fila[2] ?? "").trim();
    const debito = aNumero(fila[4]);
    const credito = aNumero(fila[5]);

    if (debito === 0 && credito === 0) continue;

    movimientos.push({
      fecha,
      fechaTexto: fechaATextoDDMMYYYY(fecha),
      banco: "SANTANDER",
      cuentaKey: cuenta.cuentaKey,
      tipo: debito !== 0 ? "debito" : "credito",
      monto: Math.abs(debito !== 0 ? debito : credito),
      descripcion,
      // Santander no separa "asunto": el cliente hay que buscarlo dentro de la propia descripción.
      textoParaMatchCliente: descripcion,
      referencia: referencia || undefined,
      categoriaBanco: tipoMovimiento || undefined,
      filaOriginal: i + 1,
    });
  }

  return movimientos;
}

// --- Uso directo por consola: `ts-node parseSantander.ts archivo.xlsx 20/07/2026` ---
if (require.main === module) {
  const [archivo, fechaTexto] = process.argv.slice(2);
  if (!archivo || !fechaTexto) {
    console.log("Uso: ts-node parseSantander.ts <archivo.xlsx> <dd/mm/yyyy>");
    process.exit(1);
  }
  const [dd, mm, yyyy] = fechaTexto.split("/").map(Number);
  const fechaObjetivo = new Date(Date.UTC(yyyy, mm - 1, dd));
  const resultado = parseSantander(archivo, fechaObjetivo);
  console.log(JSON.stringify(resultado, null, 2));
  console.log(`\nTotal movimientos encontrados para ${fechaTexto}: ${resultado.length}`);
}
