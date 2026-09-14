/**
 * rclone.ts
 *
 * Envoltorio delgado sobre el binario `rclone`, para hablar con SharePoint sin montar
 * nada en el sistema de archivos.
 *
 * ------------------------------------------------------------------------
 * POR QUÉ NO SE USA `rclone mount`
 * ------------------------------------------------------------------------
 *
 * Se probó primero con un montaje FUSE (`rclone mount`), porque permitía dejar el resto
 * del código sin cambios. No funcionó, y el modo en que falló es peor que el fallo en sí:
 *
 *   - **SharePoint modifica los archivos de Office al recibirlos.** Se suben 1.200.869
 *     bytes y del otro lado quedan 1.208.496 — y en el intento siguiente, 1.208.499.
 *     Distinto cada vez, porque agrega metadatos propios.
 *   - rclone compara el tamaño de origen y destino, no coinciden, concluye
 *     `corrupted on transfer` y **borra lo que acababa de subir**.
 *   - Con el montaje, todo eso pasa de forma asincrónica y silenciosa: nuestro proceso
 *     escribía el archivo, veía éxito, guardaba el respaldo, anotaba el registro... y la
 *     planilla nunca llegaba a SharePoint. Sin mirar los logs de rclone, nadie se enteraba.
 *
 * Se verificó que un archivo binario de 1 MB sube sin problema y que el mismo `.xlsx`
 * falla siempre: no es el tamaño ni la sesión multiparte, es el retoque que SharePoint
 * le hace a los archivos de Office.
 *
 * Con `rclone copy` explícito y `--ignore-size --ignore-checksum` la subida funciona, y
 * —lo más importante— **si falla, falla acá**: el error es de nuestro proceso, sale por
 * el log del servicio y el registro no se escribe. Se acabó el "creí que había terminado".
 */

import { execFile } from "child_process";
import { promisify } from "util";
import { cargarConfig, Config } from "./config";

const ejecutar = promisify(execFile);

export interface ArchivoRemoto {
  nombre: string;
  tamanio: number;
  modificado: Date;
  esDirectorio: boolean;
}

/** Junta el remoto y una ruta: `sharepoint:` + `a/b` -> `sharepoint:a/b` */
export function rutaRemota(config: Config, ...partes: string[]): string {
  const limpio = partes
    .filter((p) => p !== "")
    .map((p) => p.replace(/^\/+|\/+$/g, ""))
    .join("/");
  const base = config.remotoRclone.endsWith(":") ? config.remotoRclone : `${config.remotoRclone}:`;
  return `${base}${limpio}`;
}

async function correrRclone(config: Config, args: string[]): Promise<string> {
  try {
    const { stdout } = await ejecutar(config.rcloneBinario, args, {
      maxBuffer: 32 * 1024 * 1024,
      timeout: config.timeoutRcloneSegundos * 1000,
    });
    return stdout;
  } catch (err) {
    const e = err as { stderr?: string; message: string };
    // El stderr de rclone dice mucho más que el mensaje de error de Node.
    const detalle = (e.stderr ?? "").trim() || e.message;
    throw new Error(`rclone ${args[0]} falló:\n${detalle}`);
  }
}

/** Lista el contenido de una carpeta remota. Devuelve [] si la carpeta no existe. */
export async function listar(rutaCompleta: string, configExplicita?: Config): Promise<ArchivoRemoto[]> {
  const config = configExplicita ?? cargarConfig();
  let salida: string;
  try {
    salida = await correrRclone(config, ["lsjson", rutaCompleta]);
  } catch (err) {
    if (/directory not found|not found/i.test((err as Error).message)) return [];
    throw err;
  }

  const crudo = JSON.parse(salida) as Array<{
    Name: string;
    Size: number;
    ModTime: string;
    IsDir: boolean;
  }>;

  return crudo.map((a) => ({
    nombre: a.Name,
    tamanio: a.Size,
    modificado: new Date(a.ModTime),
    esDirectorio: a.IsDir,
  }));
}

/** ¿Existe este archivo remoto? */
export async function existeArchivo(rutaCompleta: string, configExplicita?: Config): Promise<ArchivoRemoto | null> {
  const config = configExplicita ?? cargarConfig();
  const barra = rutaCompleta.lastIndexOf("/");
  const carpeta = rutaCompleta.slice(0, barra);
  const nombre = rutaCompleta.slice(barra + 1);
  const contenido = await listar(carpeta, config);
  return contenido.find((a) => a.nombre === nombre && !a.esDirectorio) ?? null;
}

/** Crea una carpeta remota (y las intermedias). Idempotente. */
export async function crearCarpeta(rutaCompleta: string, configExplicita?: Config): Promise<void> {
  const config = configExplicita ?? cargarConfig();
  await correrRclone(config, ["mkdir", rutaCompleta]);
}

/** Descarga el contenido de una carpeta remota a una carpeta local. */
export async function descargarCarpeta(
  rutaRemotaCompleta: string,
  destinoLocal: string,
  configExplicita?: Config
): Promise<void> {
  const config = configExplicita ?? cargarConfig();
  await correrRclone(config, ["copy", rutaRemotaCompleta, destinoLocal, ...config.flagsRcloneTransferencia]);
}

/** Descarga un archivo remoto puntual a una ruta local exacta. */
export async function descargarArchivo(
  rutaRemotaCompleta: string,
  destinoLocal: string,
  configExplicita?: Config
): Promise<void> {
  const config = configExplicita ?? cargarConfig();
  await correrRclone(config, ["copyto", rutaRemotaCompleta, destinoLocal, ...config.flagsRcloneTransferencia]);
}

export interface ResultadoSubida {
  rutaRemota: string;
  tamanioLocal: number;
  tamanioRemoto: number;
  modificadoRemoto: Date;
}

/**
 * Sube un archivo local a una ruta remota exacta y **verifica que haya llegado**.
 *
 * La verificación no puede comparar tamaños: SharePoint retoca los archivos de Office y
 * el tamaño del destino nunca coincide con el del origen (por eso hacen falta
 * `--ignore-size --ignore-checksum`). Lo que sí se comprueba es que el archivo exista y
 * que su fecha de modificación sea posterior al momento en que arrancó la subida.
 *
 * Eso alcanza para detectar el modo de falla real que tuvimos: cuando rclone consideraba
 * la transferencia corrupta, **borraba el archivo del destino**. Con esta verificación,
 * ese caso sale como error inmediato en vez de pasar desapercibido.
 */
export async function subirArchivo(
  rutaLocal: string,
  rutaRemotaCompleta: string,
  configExplicita?: Config
): Promise<ResultadoSubida> {
  const config = configExplicita ?? cargarConfig();
  const fs = await import("fs");

  const tamanioLocal = fs.statSync(rutaLocal).size;
  // Margen hacia atrás: el reloj del servidor y el nuestro no tienen por qué coincidir
  // al segundo, y no queremos fallar por eso.
  const momentoInicio = new Date(Date.now() - 5 * 60 * 1000);

  await correrRclone(config, [
    "copyto",
    rutaLocal,
    rutaRemotaCompleta,
    ...config.flagsRcloneTransferencia,
    ...config.flagsRcloneSubida,
  ]);

  const remoto = await existeArchivo(rutaRemotaCompleta, config);
  if (!remoto) {
    throw new Error(
      `La subida de ${rutaRemotaCompleta} terminó sin error pero el archivo no está en el destino. ` +
        `Es el síntoma de que rclone lo consideró corrupto y lo borró; revisar los flags de transferencia.`
    );
  }
  if (remoto.modificado.getTime() < momentoInicio.getTime()) {
    throw new Error(
      `La subida de ${rutaRemotaCompleta} terminó sin error pero el archivo remoto tiene fecha ` +
        `${remoto.modificado.toISOString()}, anterior a esta corrida. Probablemente no se haya escrito.`
    );
  }

  return {
    rutaRemota: rutaRemotaCompleta,
    tamanioLocal,
    tamanioRemoto: remoto.tamanio,
    modificadoRemoto: remoto.modificado,
  };
}

/** Borra un archivo remoto. */
export async function borrarArchivo(rutaRemotaCompleta: string, configExplicita?: Config): Promise<void> {
  const config = configExplicita ?? cargarConfig();
  await correrRclone(config, ["deletefile", rutaRemotaCompleta]);
}

/** Comprueba que el remoto responde. Útil para diagnosticar antes de procesar. */
export async function verificarConexion(configExplicita?: Config): Promise<string[]> {
  const config = configExplicita ?? cargarConfig();
  const salida = await correrRclone(config, ["lsd", rutaRemota(config, "")]);
  return salida.trim().split("\n").filter((l) => l.trim() !== "");
}

// --- Uso por consola: node rclone.js <listar|verificar> [ruta] ---
if (require.main === module) {
  const [accion, ruta] = process.argv.slice(2);
  const config = cargarConfig();

  (async () => {
    try {
      if (accion === "verificar") {
        const lineas = await verificarConexion(config);
        console.log(`Conexión OK con "${config.remotoRclone}". Contenido de la raíz:`);
        lineas.forEach((l) => console.log(`  ${l.trim()}`));
      } else if (accion === "listar") {
        const destino = ruta ?? rutaRemota(config, config.rutaRemotaBase);
        const contenido = await listar(destino, config);
        console.log(`${destino}  (${contenido.length} elementos)`);
        for (const a of contenido) {
          console.log(
            `  ${a.esDirectorio ? "[dir] " : "      "}${a.nombre.padEnd(60)} ${String(a.tamanio).padStart(10)}  ${a.modificado.toISOString()}`
          );
        }
      } else {
        console.log("Uso:");
        console.log("  node rclone.js verificar");
        console.log("  node rclone.js listar [ruta remota]");
        process.exit(1);
      }
    } catch (err) {
      console.error("ERROR:", (err as Error).message);
      process.exit(1);
    }
  })();
}
