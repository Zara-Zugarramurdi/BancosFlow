/**
 * procesarEstadoDeCuenta.ts
 *
 * Punto de entrada único: recibe cualquiera de los 2 formatos (BROU o Santander),
 * detecta cuál es, y devuelve los movimientos limpios del día pedido + a qué cuenta
 * corresponde el archivo. Esto es lo que en la práctica va a llamar el proceso que
 * lee la carpeta de "halcón" todas las mañanas.
 */

import * as XLSX from "xlsx";
import { identificarCuenta } from "./accountIdentifier";
import { parseBrou } from "./parseBrou";
import { parseSantander } from "./parseSantander";
import type { CuentaIdentificada, MovimientoLimpio } from "./types";

export interface ResultadoProcesamiento {
  cuenta: CuentaIdentificada;
  movimientos: MovimientoLimpio[];
}

export function procesarEstadoDeCuenta(
  rutaArchivo: string,
  fechaObjetivo: Date
): ResultadoProcesamiento {
  const workbook = XLSX.readFile(rutaArchivo);
  const cuenta = identificarCuenta(workbook);

  const movimientos =
    cuenta.banco === "BROU"
      ? parseBrou(workbook, fechaObjetivo)
      : parseSantander(workbook, fechaObjetivo);

  return { cuenta, movimientos };
}

// --- Uso directo por consola: `ts-node procesarEstadoDeCuenta.ts archivo.xlsx 20/07/2026` ---
if (require.main === module) {
  const [archivo, fechaTexto] = process.argv.slice(2);
  if (!archivo || !fechaTexto) {
    console.log("Uso: ts-node procesarEstadoDeCuenta.ts <archivo> <dd/mm/yyyy>");
    process.exit(1);
  }
  const [dd, mm, yyyy] = fechaTexto.split("/").map(Number);
  const fechaObjetivo = new Date(Date.UTC(yyyy, mm - 1, dd));

  const { cuenta, movimientos } = procesarEstadoDeCuenta(archivo, fechaObjetivo);

  console.log("Cuenta detectada:", cuenta);
  console.log(`Movimientos del ${fechaTexto}: ${movimientos.length}`);
  console.table(
    movimientos.map((m) => ({
      fecha: m.fechaTexto,
      tipo: m.tipo,
      monto: m.monto,
      textoParaMatchCliente: m.textoParaMatchCliente.slice(0, 60),
    }))
  );
}
