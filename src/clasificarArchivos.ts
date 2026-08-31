/**
 * clasificarArchivos.ts
 *
 * Dado el contenido de la carpeta del día, decide qué es cada archivo:
 * la planilla de bancos, un estado de cuenta, o algo a ignorar.
 *
 * La clasificación se hace por CONTENIDO, nunca por nombre de archivo: los nombres
 * reales varían todo el tiempo ("Copia de Planilla BANCOS desde 12-22 18-08.xlsx",
 * "Detalle_Movimiento_Cuenta (7).xls", "ESTADO_DE_CUENTA5-0-0000...(5).xlsx") y
 * depender de ellos sería frágil.
 */

import * as fs from "fs";
import * as path from "path";
import * as XLSX from "xlsx";
import { identificarCuenta } from "./accountIdentifier";
import { cargarConfig, Config } from "./config";
import type { CuentaIdentificada } from "./types";

/**
 * Hojas que tiene que tener un archivo para considerarse la planilla de bancos.
 * Se exige que estén TODAS: son las 5 cuentas del ledger. Un estado de cuenta
 * tiene una sola hoja, así que no hay forma de confundirlos.
 */
const HOJAS_REQUERIDAS_PLANILLA = ["BROU $", "BROU U$S", "BROU EUROS", "SANTANDER $", "SANTANDER U$S"];

const EXTENSIONES_EXCEL = [".xls", ".xlsx", ".xlsm"];

export interface ArchivoEstadoDeCuenta {
  ruta: string;
  nombre: string;
  cuenta: CuentaIdentificada;
  modificado: Date;
  tamanio: number;
}

export interface ArchivoPlanilla {
  ruta: string;
  nombre: string;
  modificado: Date;
  tamanio: number;
}

export interface ArchivoIgnorado {
  ruta: string;
  nombre: string;
  motivo: string;
}

export interface Clasificacion {
  /** La planilla elegida (la más reciente si hubiera varias). */
  planilla: ArchivoPlanilla | null;
  /** Planillas descartadas por no ser la más reciente. Se borran en el procesamiento. */
  planillasDuplicadas: ArchivoPlanilla[];
  estadosDeCuenta: ArchivoEstadoDeCuenta[];
  ignorados: ArchivoIgnorado[];
  /** Si hay algún archivo de bloqueo, indicando que alguien tiene un libro abierto. */
  hayArchivosAbiertos: boolean;
}

/**
 * ¿Es un archivo de bloqueo, de los que aparecen mientras alguien tiene el libro abierto?
 *
 * Hay dos formatos según el programa, y hay que cubrir los dos:
 *   - Excel:       `~$nombre.xlsx`
 *   - LibreOffice: `.~lock.nombre.xlsx#`
 *
 * El de LibreOffice apareció en la carpeta real de producción: en esa VM la planilla se
 * abre con LibreOffice, no con Excel. Antes se detectaba sólo el de Excel, así que el
 * archivo de LibreOffice caía en "no es un Excel" y el proceso escribía igual sobre una
 * planilla abierta — con el riesgo de que la persona guardara encima de lo insertado.
 */
function esArchivoDeBloqueo(nombre: string): boolean {
  if (nombre.startsWith("~$")) return true;                          // Excel
  if (nombre.startsWith(".~lock.") && nombre.endsWith("#")) return true; // LibreOffice
  return false;
}

/** ¿Tiene el archivo las hojas que caracterizan a la planilla de bancos? */
function esPlanillaDeBancos(ruta: string): boolean {
  try {
    // `bookSheets` lee sólo la lista de hojas: mucho más barato que parsear todo el libro.
    const wb = XLSX.readFile(ruta, { bookSheets: true });
    const hojas = new Set(wb.SheetNames);
    return HOJAS_REQUERIDAS_PLANILLA.every((h) => hojas.has(h));
  } catch {
    return false;
  }
}

/**
 * Clasifica el contenido de una carpeta.
 *
 * El orden de las decisiones importa: primero se descartan los archivos que no son
 * Excel, después se intenta identificar la cuenta (si funciona, es un estado de
 * cuenta), y recién si eso falla se mira si tiene las hojas de la planilla. Así un
 * estado de cuenta nunca se confunde con la planilla ni al revés.
 */
export function clasificarCarpeta(rutaCarpeta: string, configExplicita?: Config): Clasificacion {
  const config = configExplicita ?? cargarConfig();

  if (!fs.existsSync(rutaCarpeta)) {
    throw new Error(`La carpeta no existe: ${rutaCarpeta}`);
  }

  const planillas: ArchivoPlanilla[] = [];
  const estadosDeCuenta: ArchivoEstadoDeCuenta[] = [];
  const ignorados: ArchivoIgnorado[] = [];
  let hayArchivosAbiertos = false;

  for (const nombre of fs.readdirSync(rutaCarpeta)) {
    const ruta = path.join(rutaCarpeta, nombre);

    let stat: fs.Stats;
    try {
      stat = fs.statSync(ruta);
    } catch {
      ignorados.push({ ruta, nombre, motivo: "no se pudo leer" });
      continue;
    }
    if (stat.isDirectory()) {
      ignorados.push({ ruta, nombre, motivo: "es una carpeta" });
      continue;
    }

    if (esArchivoDeBloqueo(nombre)) {
      hayArchivosAbiertos = true;
      ignorados.push({ ruta, nombre, motivo: "archivo de bloqueo: alguien tiene el libro abierto" });
      continue;
    }
    if (nombre === config.nombreArchivoRegistro) {
      ignorados.push({ ruta, nombre, motivo: "archivo de control del proceso" });
      continue;
    }
    if (!EXTENSIONES_EXCEL.includes(path.extname(nombre).toLowerCase())) {
      ignorados.push({ ruta, nombre, motivo: "no es un archivo Excel" });
      continue;
    }

    // ¿Es un estado de cuenta? Lo decide el identificador que ya usamos.
    try {
      const cuenta = identificarCuenta(ruta);
      estadosDeCuenta.push({ ruta, nombre, cuenta, modificado: stat.mtime, tamanio: stat.size });
      continue;
    } catch {
      // No es un estado de cuenta reconocible: seguimos probando.
    }

    if (esPlanillaDeBancos(ruta)) {
      planillas.push({ ruta, nombre, modificado: stat.mtime, tamanio: stat.size });
      continue;
    }

    ignorados.push({
      ruta,
      nombre,
      motivo: "Excel no reconocido (no es un estado de cuenta ni la planilla de bancos)",
    });
  }

  // Si hay más de una planilla, gana la más reciente por fecha de modificación.
  planillas.sort((a, b) => b.modificado.getTime() - a.modificado.getTime());
  const [planilla = null, ...planillasDuplicadas] = planillas;

  // Orden estable de los estados de cuenta, para que los reportes sean predecibles.
  estadosDeCuenta.sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));

  return { planilla, planillasDuplicadas, estadosDeCuenta, ignorados, hayArchivosAbiertos };
}

// --- Uso por consola: node clasificarArchivos.js <carpeta> ---
if (require.main === module) {
  const rutaCarpeta = process.argv[2];
  if (!rutaCarpeta) {
    console.log("Uso: node clasificarArchivos.js <carpeta>");
    process.exit(1);
  }
  try {
    const c = clasificarCarpeta(rutaCarpeta);
    console.log(`Carpeta: ${rutaCarpeta}\n`);
    console.log(`Planilla de bancos: ${c.planilla ? c.planilla.nombre : "(no encontrada)"}`);
    if (c.planillasDuplicadas.length > 0) {
      console.log("Planillas más viejas (se descartarían):");
      c.planillasDuplicadas.forEach((p) => console.log(`   ${p.nombre}  (${p.modificado.toISOString()})`));
    }
    console.log(`\nEstados de cuenta: ${c.estadosDeCuenta.length}`);
    c.estadosDeCuenta.forEach((e) => console.log(`   ${e.cuenta.etiqueta.padEnd(18)} ${e.nombre}`));
    if (c.ignorados.length > 0) {
      console.log(`\nIgnorados: ${c.ignorados.length}`);
      c.ignorados.forEach((i) => console.log(`   ${i.nombre}  -> ${i.motivo}`));
    }
    if (c.hayArchivosAbiertos) {
      console.log("\nAVISO: alguien tiene un libro abierto; conviene esperar antes de procesar.");
    }
  } catch (err) {
    console.error("ERROR:", (err as Error).message);
    process.exit(1);
  }
}
