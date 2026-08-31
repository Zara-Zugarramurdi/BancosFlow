/**
 * registroProcesados.ts
 *
 * Maneja el archivo de control (`.bancosflow.json`) que queda dentro de cada carpeta
 * del día, con qué estados de cuenta ya se procesaron, cuándo y con qué resultado.
 *
 * Se eligió un archivo de control aparte en vez de renombrar los archivos originales
 * (por ejemplo agregándoles `.procesado`) por dos motivos: son archivos que subió una
 * persona y modificarlos es intrusivo, y sobre un montaje SMB un rename puede fallar
 * si alguien tiene el archivo abierto.
 *
 * Se guarda el hash SHA-256 de cada estado de cuenta procesado, no sólo el nombre. Así,
 * si alguien vuelve a subir un archivo corregido con el mismo nombre, el hash cambia y
 * se reprocesa en vez de darlo por hecho.
 *
 * OJO: esto evita trabajo repetido, pero NO es la protección contra duplicados. Esa
 * sigue siendo la deduplicación por fecha+tipo+monto de `actualizarPlanilla`, que
 * funciona aunque el registro se borre o se pierda.
 */

import * as crypto from "crypto";
import * as fs from "fs";

/** Versión del formato del archivo, por si algún día cambia su estructura. */
const VERSION_FORMATO = 1;

export interface EntradaProcesado {
  /** Nombre del archivo dentro de la carpeta del día. */
  archivo: string;
  /** SHA-256 del contenido, para detectar reenvíos corregidos con el mismo nombre. */
  hash: string;
  /** Cuenta detectada, como etiqueta legible. */
  cuenta: string;
  /** Cuándo se procesó (ISO 8601). */
  procesadoEn: string;
  /** Cuántos movimientos se agregaron a la planilla. */
  agregados: number;
  /** Cuántos se omitieron por estar ya cargados. */
  omitidos: number;
  /** Filas de la planilla donde quedaron, para poder auditar. */
  filas?: string;
}

export interface Registro {
  version: number;
  /** Fecha de la carpeta (yyyy-mm-dd), para que el archivo se entienda solo. */
  fecha: string;
  actualizadoEn: string;
  procesados: EntradaProcesado[];
}

export function hashDeArchivo(ruta: string): string {
  return crypto.createHash("sha256").update(fs.readFileSync(ruta)).digest("hex");
}

export function registroVacio(fecha: string): Registro {
  return { version: VERSION_FORMATO, fecha, actualizadoEn: new Date().toISOString(), procesados: [] };
}

/**
 * Lee el registro de una carpeta. Si no existe, o si está corrupto, devuelve uno vacío:
 * perder el registro nunca debe frenar el proceso, porque lo peor que puede pasar es que
 * se reprocese un archivo, y de eso ya protege la deduplicación.
 */
export function leerRegistro(rutaRegistro: string, fecha: string): Registro {
  if (!fs.existsSync(rutaRegistro)) return registroVacio(fecha);

  try {
    const datos = JSON.parse(fs.readFileSync(rutaRegistro, "utf8")) as Registro;
    if (!Array.isArray(datos.procesados)) throw new Error("sin lista de procesados");
    return { ...registroVacio(fecha), ...datos };
  } catch (err) {
    console.warn(
      `AVISO: el registro "${rutaRegistro}" no se pudo leer (${(err as Error).message}). ` +
        `Se sigue con un registro vacío; a lo sumo se reprocesa algún archivo, y la ` +
        `deduplicación evita que se dupliquen movimientos.`
    );
    return registroVacio(fecha);
  }
}

export function guardarRegistro(rutaRegistro: string, registro: Registro): void {
  registro.actualizadoEn = new Date().toISOString();
  // Escritura atómica: se escribe a un temporal y se renombra, para que una
  // interrupción a mitad de camino no deje un JSON truncado.
  const temporal = `${rutaRegistro}.tmp`;
  fs.writeFileSync(temporal, JSON.stringify(registro, null, 2), "utf8");
  fs.renameSync(temporal, rutaRegistro);
}

/** ¿Este archivo, con este contenido exacto, ya se procesó? */
export function yaProcesado(registro: Registro, nombreArchivo: string, hash: string): boolean {
  return registro.procesados.some((p) => p.archivo === nombreArchivo && p.hash === hash);
}

/**
 * Anota un archivo como procesado. Si ya había una entrada para ese nombre (por ejemplo
 * una versión anterior del mismo archivo), se reemplaza en lugar de acumular.
 */
export function anotarProcesado(registro: Registro, entrada: EntradaProcesado): void {
  const indice = registro.procesados.findIndex((p) => p.archivo === entrada.archivo);
  if (indice === -1) registro.procesados.push(entrada);
  else registro.procesados[indice] = entrada;
}

// --- Uso por consola: node registroProcesados.js <ruta del .bancosflow.json> ---
if (require.main === module) {
  const ruta = process.argv[2];
  if (!ruta) {
    console.log("Uso: node registroProcesados.js <ruta del archivo de registro>");
    process.exit(1);
  }
  const r = leerRegistro(ruta, "(desconocida)");
  console.log(`Registro: ${ruta}`);
  console.log(`Fecha: ${r.fecha}  ·  actualizado: ${r.actualizadoEn}`);
  console.log(`Archivos procesados: ${r.procesados.length}`);
  for (const p of r.procesados) {
    console.log(`  ${p.cuenta.padEnd(18)} ${p.archivo}`);
    console.log(`     ${p.agregados} agregados, ${p.omitidos} omitidos${p.filas ? `, filas ${p.filas}` : ""}  ·  ${p.procesadoEn}`);
  }
}
