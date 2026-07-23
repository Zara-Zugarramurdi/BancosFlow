/**
 * accountIdentifier.ts
 *
 * Dado un archivo de estado de cuenta (BROU o Santander, .xls o .xlsx),
 * determina a cuál de las 5 cuentas de la empresa corresponde.
 *
 * Funciona leyendo la estructura "de cabecera" de cada formato, que es
 * estable independientemente del día en que se descargue el archivo.
 */

import * as XLSX from "xlsx";
import type { Banco, Moneda, CuentaIdentificada, CuentaKey } from "./types";

/** Tabla de mapeo: (banco, moneda) -> cuenta de la empresa.
 *  Si mañana se agrega una 6ta cuenta, solo hay que sumar una fila acá. */
const MAPA_CUENTAS: Record<string, { cuentaKey: CuentaKey; etiqueta: string }> = {
  "BROU|UYU": { cuentaKey: "BROU_PESOS", etiqueta: "BROU Pesos" },
  "BROU|USD": { cuentaKey: "BROU_DOLARES", etiqueta: "BROU Dólares" },
  "BROU|EUR": { cuentaKey: "BROU_EUROS", etiqueta: "BROU Euros" },
  "SANTANDER|UYU": { cuentaKey: "SANTANDER_PESOS", etiqueta: "Santander Pesos" },
  "SANTANDER|USD": { cuentaKey: "SANTANDER_DOLARES", etiqueta: "Santander Dólares" },
};

/** Convierte el símbolo/código de moneda que usa cada banco a un código ISO estándar. */
function normalizarMoneda(raw: string): Moneda {
  const v = raw.trim().toUpperCase();
  if (v === "UYU" || v === "$" || v.startsWith("$")) return "UYU";
  if (v === "USD" || v === "U$S" || v.includes("USD") || v.includes("U$S")) return "USD";
  if (v === "EUR" || v === "€" || v.includes("EUR") || v.includes("€")) return "EUR";
  throw new Error(`No se pudo reconocer la moneda: "${raw}"`);
}

function leerHoja(workbook: XLSX.WorkBook): XLSX.WorkSheet {
  const nombreHoja = workbook.SheetNames[0];
  return workbook.Sheets[nombreHoja];
}

/** Convierte la hoja a una matriz de filas/columnas (más fácil de recorrer que el objeto crudo de SheetJS). */
function hojaAMatriz(sheet: XLSX.WorkSheet): unknown[][] {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" }) as unknown[][];
}

/** Detecta si el archivo es de BROU o de Santander mirando las primeras filas. */
function detectarBanco(matriz: unknown[][]): Banco {
  const primeras20 = matriz.slice(0, 20).map((f) => f.join(" | ")).join("\n");
  if (primeras20.includes("Saldos y Movimientos")) return "BROU";
  if (primeras20.includes("Banco Santander")) return "SANTANDER";
  throw new Error(
    "No se pudo determinar el banco: el archivo no contiene ni 'Saldos y Movimientos' (BROU) ni 'Banco Santander' (Santander)."
  );
}

function identificarBrou(matriz: unknown[][]): { numeroCuenta: string; moneda: Moneda } {
  // Buscamos la fila que trae "Nº de Cuenta" y "Moneda" (normalmente la fila 6, pero
  // buscamos dinámicamente por robustez ante pequeños cambios de formato).
  for (const fila of matriz) {
    const celdaCuenta = fila.find(
      (c) => typeof c === "string" && c.includes("Nº de Cuenta")
    ) as string | undefined;
    const celdaMoneda = fila.find(
      (c) => typeof c === "string" && c.startsWith("Moneda")
    ) as string | undefined;

    if (celdaCuenta && celdaMoneda) {
      const numeroCuenta = celdaCuenta.split("\n")[1]?.trim() ?? celdaCuenta;
      const monedaRaw = celdaMoneda.split("\n")[1]?.trim() ?? celdaMoneda;
      return { numeroCuenta, moneda: normalizarMoneda(monedaRaw) };
    }
  }
  throw new Error("BROU: no se encontró la fila con 'Nº de Cuenta' / 'Moneda'.");
}

function identificarSantander(matriz: unknown[][]): { numeroCuenta: string; moneda: Moneda } {
  // Buscamos la fila de encabezado "Cuenta | | Moneda | Sucursal", los datos están en la fila siguiente.
  const idx = matriz.findIndex(
    (fila) => fila[0] === "Cuenta" && fila[2] === "Moneda"
  );
  if (idx === -1 || !matriz[idx + 1]) {
    throw new Error("Santander: no se encontró la fila de encabezado 'Cuenta / Moneda / Sucursal'.");
  }
  const filaDatos = matriz[idx + 1];
  const descripcionCuenta = String(filaDatos[0] ?? "");
  const monedaRaw = String(filaDatos[2] ?? "");

  // descripcionCuenta viene como "CTA. PYME PREMIUM PESOS, 000000148245"
  const numeroCuenta = descripcionCuenta.split(",").pop()?.trim() ?? descripcionCuenta;

  return { numeroCuenta, moneda: normalizarMoneda(monedaRaw) };
}

/**
 * Identifica a qué cuenta de la empresa corresponde un archivo de estado de cuenta.
 * Acepta tanto la ruta a un archivo (.xls o .xlsx) como un WorkBook ya leído por SheetJS.
 */
export function identificarCuenta(input: string | XLSX.WorkBook): CuentaIdentificada {
  const workbook = typeof input === "string" ? XLSX.readFile(input) : input;
  const matriz = hojaAMatriz(leerHoja(workbook));

  const banco = detectarBanco(matriz);
  const { numeroCuenta, moneda } =
    banco === "BROU" ? identificarBrou(matriz) : identificarSantander(matriz);

  const clave = `${banco}|${moneda}`;
  const match = MAPA_CUENTAS[clave];
  if (!match) {
    throw new Error(
      `Combinación banco/moneda no reconocida: ${clave}. Revisar MAPA_CUENTAS en accountIdentifier.ts si es una cuenta nueva.`
    );
  }

  return {
    banco,
    moneda,
    numeroCuenta,
    cuentaKey: match.cuentaKey,
    etiqueta: match.etiqueta,
  };
}

// --- Uso directo por consola: `ts-node accountIdentifier.ts archivo1.xls archivo2.xlsx ...` ---
if (require.main === module) {
  const archivos = process.argv.slice(2);
  if (archivos.length === 0) {
    console.log("Uso: ts-node accountIdentifier.ts <archivo1> <archivo2> ...");
    process.exit(1);
  }
  for (const archivo of archivos) {
    try {
      const resultado = identificarCuenta(archivo);
      console.log(archivo, "->", resultado);
    } catch (err) {
      console.error(archivo, "-> ERROR:", (err as Error).message);
    }
  }
}
