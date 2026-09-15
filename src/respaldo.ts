/**
 * respaldo.ts
 *
 * Copia la planilla antes de modificarla, y purga las copias que superan la retención
 * configurada.
 *
 * Se respalda **sólo la planilla**: es lo único que el proceso modifica y lo único
 * irreemplazable. Los estados de cuenta quedan igual en el fileserver y, si hiciera
 * falta, se vuelven a descargar del banco.
 *
 * Los respaldos van al disco local de la VM (`rutaBackups`), no al fileserver: si el
 * problema fuera justamente el montaje de red o algo que corrompa esa carpeta, tener
 * la copia en el mismo lugar no serviría de nada.
 */

import * as fs from "fs";
import * as path from "path";
import { cargarConfig, Config } from "./config";

export interface ResultadoRespaldo {
  rutaOrigen: string;
  rutaRespaldo: string;
  simulado: boolean;
  /** Si se salteó por existir ya un respaldo idéntico. */
  omitido?: boolean;
}

/**
 * Marca de tiempo para el nombre del respaldo: `2026-08-31_142530`.
 * Sin `:` porque no es un carácter válido en nombres de archivo en Windows, y estos
 * respaldos podrían terminar copiándose a un share.
 */
function marcaDeTiempo(fecha: Date = new Date()): string {
  const p = (n: number) => String(n).padStart(2, "0");
  return (
    `${fecha.getFullYear()}-${p(fecha.getMonth() + 1)}-${p(fecha.getDate())}` +
    `_${p(fecha.getHours())}${p(fecha.getMinutes())}${p(fecha.getSeconds())}`
  );
}

/**
 * Respalda la planilla. Devuelve `null` en modo simulación sólo si no se pudo calcular
 * la ruta; en simulación normal devuelve la ruta que se usaría sin copiar nada.
 *
 * El llamador debe invocar esto **sólo cuando realmente va a escribir**. Si no hay
 * movimientos nuevos, no tiene sentido acumular copias idénticas.
 */
/**
 * ¿Ya existe un respaldo idéntico al archivo que estamos por copiar?
 *
 * Evita acumular copias repetidas cuando el proceso reintenta. Escenario real: si la
 * publicación falla, el poller reintenta cada ~60s; como el registro no se escribe, cada
 * ciclo vuelve a procesar y a respaldar. Sin este chequeo eso significaba un archivo de
 * ~1,2 MB por minuto (unos 2 GB por día) en un disco de 30 GB.
 *
 * Se comparan tamaño y fecha de modificación del origen: en ese bucle de reintento la
 * planilla se vuelve a descargar del remoto y conserva su fecha, así que se mantiene
 * idéntica ciclo a ciclo y el respaldo se saltea a partir del segundo.
 */
function yaHayRespaldoIdentico(rutaPlanilla: string, config: Config): boolean {
  if (!fs.existsSync(config.rutaBackups)) return false;
  const origen = fs.statSync(rutaPlanilla);
  const base = path.basename(rutaPlanilla, path.extname(rutaPlanilla));

  const candidatos = fs
    .readdirSync(config.rutaBackups)
    .filter((n) => n.startsWith(`${base}__`))
    .map((n) => {
      const ruta = path.join(config.rutaBackups, n);
      return { ruta, stat: fs.statSync(ruta) };
    })
    .sort((a, b) => b.stat.mtimeMs - a.stat.mtimeMs);

  if (candidatos.length === 0) return false;
  const ultimo = candidatos[0];

  // Mismo tamaño y respaldo posterior (o igual) a la última modificación del origen:
  // es la misma versión del archivo, no hace falta duplicarla.
  return ultimo.stat.size === origen.size && ultimo.stat.mtimeMs >= origen.mtimeMs;
}

export function respaldarPlanilla(
  rutaPlanilla: string,
  opciones: { simular?: boolean; config?: Config; fecha?: Date } = {}
): ResultadoRespaldo {
  const config = opciones.config ?? cargarConfig();
  const simular = opciones.simular ?? false;

  if (!fs.existsSync(rutaPlanilla)) {
    throw new Error(`No se puede respaldar: la planilla no existe en ${rutaPlanilla}`);
  }

  if (!simular && !fs.existsSync(config.rutaBackups)) {
    fs.mkdirSync(config.rutaBackups, { recursive: true });
  }

  const extension = path.extname(rutaPlanilla);
  const base = path.basename(rutaPlanilla, extension);
  const nombre = `${base}__${marcaDeTiempo(opciones.fecha)}${extension}`;
  const rutaRespaldo = path.join(config.rutaBackups, nombre);

  if (!simular && yaHayRespaldoIdentico(rutaPlanilla, config)) {
    return { rutaOrigen: rutaPlanilla, rutaRespaldo, simulado: false, omitido: true };
  }

  if (!simular) {
    fs.copyFileSync(rutaPlanilla, rutaRespaldo);
  }

  return { rutaOrigen: rutaPlanilla, rutaRespaldo, simulado: simular, omitido: false };
}

export interface ResultadoPurga {
  eliminados: string[];
  conservados: number;
  simulado: boolean;
}

/**
 * Borra los respaldos más viejos que `retencionBackupsDias`.
 *
 * Sólo mira archivos con extensión de Excel dentro de `rutaBackups`, para no borrar por
 * error algo que alguien haya dejado ahí.
 */
export function purgarRespaldosViejos(
  opciones: { simular?: boolean; config?: Config; ahora?: Date } = {}
): ResultadoPurga {
  const config = opciones.config ?? cargarConfig();
  const simular = opciones.simular ?? false;
  const ahora = opciones.ahora ?? new Date();

  if (!fs.existsSync(config.rutaBackups)) {
    return { eliminados: [], conservados: 0, simulado: simular };
  }

  const limite = ahora.getTime() - config.retencionBackupsDias * 24 * 60 * 60 * 1000;
  const eliminados: string[] = [];
  let conservados = 0;

  for (const nombre of fs.readdirSync(config.rutaBackups)) {
    const ruta = path.join(config.rutaBackups, nombre);
    if (![".xlsx", ".xls", ".xlsm"].includes(path.extname(nombre).toLowerCase())) continue;

    let stat: fs.Stats;
    try {
      stat = fs.statSync(ruta);
    } catch {
      continue;
    }
    if (!stat.isFile()) continue;

    if (stat.mtime.getTime() < limite) {
      eliminados.push(ruta);
      if (!simular) fs.unlinkSync(ruta);
    } else {
      conservados++;
    }
  }

  return { eliminados, conservados, simulado: simular };
}

// --- Uso por consola:
//   node respaldo.js purgar [--simular]
//   node respaldo.js copiar <planilla> [--simular]
if (require.main === module) {
  const args = process.argv.slice(2);
  const accion = args[0];
  const simular = args.includes("--simular");

  try {
    if (accion === "purgar") {
      const r = purgarRespaldosViejos({ simular });
      console.log(`Respaldos conservados: ${r.conservados}`);
      console.log(`${simular ? "Se eliminarían" : "Eliminados"}: ${r.eliminados.length}`);
      r.eliminados.forEach((e) => console.log(`   ${path.basename(e)}`));
    } else if (accion === "copiar" && args[1]) {
      const r = respaldarPlanilla(args[1], { simular });
      console.log(`${simular ? "Se copiaría a" : "Respaldo creado"}: ${r.rutaRespaldo}`);
    } else {
      console.log("Uso:");
      console.log("  node respaldo.js purgar [--simular]");
      console.log("  node respaldo.js copiar <planilla> [--simular]");
      process.exit(1);
    }
  } catch (err) {
    console.error("ERROR:", (err as Error).message);
    process.exit(1);
  }
}
