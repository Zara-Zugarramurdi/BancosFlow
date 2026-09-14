/**
 * almacenamiento.ts
 *
 * Aísla *de dónde salen* los archivos del resto del proceso.
 *
 * Todo el trabajo pesado (`clasificarArchivos`, los parsers, `actualizarPlanilla`) sigue
 * operando sobre rutas locales comunes y no se entera de si esas rutas son una carpeta
 * montada o una copia que se bajó de SharePoint. Este módulo se encarga de dejar los
 * archivos donde el resto espera encontrarlos, y de publicar la planilla al terminar.
 *
 * Dos modos, según `config.modoAcceso`:
 *
 *   - `sistemaArchivos`: la carpeta ya está montada. No hay nada que bajar ni que subir;
 *     las rutas que se devuelven son las del montaje. Es el comportamiento histórico.
 *   - `rclone`: se baja la carpeta del día y la planilla a `rutaTrabajoLocal`, se trabaja
 *     ahí, y al final se sube la planilla verificando que haya llegado.
 */

import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { cargarConfig, Config } from "./config";
import { calcularRutasDelDia, crearCarpetasDelDia, RutasDelDia } from "./rutasPlanillaBancos";
import * as rclone from "./rclone";

export interface EspacioDeTrabajo {
  /** Rutas "lógicas" del día (año, mes, día y fecha). */
  rutas: RutasDelDia;
  /** Carpeta LOCAL donde están los archivos del día para procesar. */
  carpetaLocalDelDia: string;
  /** Ruta LOCAL de la planilla, o null si todavía no está disponible. */
  planillaLocal: string | null;
  /** Ruta LOCAL del archivo de registro. */
  rutaRegistro: string;
  /** Si hay que subir la planilla al terminar (sólo en modo rclone). */
  requierePublicacion: boolean;
}

/** Ruta del registro cuando se guarda en disco local, separado por fecha. */
function rutaRegistroLocal(config: Config, fecha: string): string {
  const carpeta = path.join(config.rutaTrabajoLocal, "registros");
  fs.mkdirSync(carpeta, { recursive: true });
  return path.join(carpeta, `${fecha}.json`);
}

/**
 * Deja todo listo para procesar el día indicado y devuelve rutas locales.
 *
 * En modo `rclone` crea la carpeta remota si falta, baja su contenido y baja la planilla.
 * En modo `sistemaArchivos` sólo crea las carpetas y devuelve las rutas del montaje.
 */
export async function prepararEspacioDeTrabajo(
  fecha: Date = new Date(),
  opciones: { config?: Config; simular?: boolean } = {}
): Promise<EspacioDeTrabajo> {
  const config = opciones.config ?? cargarConfig();
  const simular = opciones.simular ?? false;
  const rutas = calcularRutasDelDia(fecha, config);

  // ---------- Modo montaje: el comportamiento de siempre ----------
  if (config.modoAcceso === "sistemaArchivos") {
    crearCarpetasDelDia(fecha, { config, simular });
    const planilla =
      config.ubicacionPlanilla === "maestraFija"
        ? fs.existsSync(config.rutaPlanillaMaestra)
          ? config.rutaPlanillaMaestra
          : null
        : null; // en carpetaDelDia la resuelve el clasificador
    return {
      rutas,
      carpetaLocalDelDia: rutas.rutaDia,
      planillaLocal: planilla,
      rutaRegistro:
        config.ubicacionRegistro === "local"
          ? rutaRegistroLocal(config, rutas.fecha)
          : rutas.rutaRegistro,
      requierePublicacion: false,
    };
  }

  // ---------- Modo rclone ----------
  const carpetaLocal = path.join(config.rutaTrabajoLocal, "dia", rutas.fecha);
  // Se limpia y se vuelve a bajar en cada corrida: así nunca se procesa contra una copia
  // vieja si alguien cambió algo en SharePoint.
  fs.rmSync(carpetaLocal, { recursive: true, force: true });
  fs.mkdirSync(carpetaLocal, { recursive: true });

  const remotoDia = rclone.rutaRemota(config, config.rutaRemotaBase, rutas.anio, rutas.mes, rutas.dia);

  // La DESCARGA se hace también en modo simulación: sólo lee del remoto y escribe en
  // nuestra carpeta de trabajo, así que no altera nada de Administración. Hace falta para
  // poder informar qué se agregaría. Lo único que la simulación evita es la subida.
  if (!simular) {
    await rclone.crearCarpeta(remotoDia, config);
  }
  await rclone.descargarCarpeta(remotoDia, carpetaLocal, config);

  // Bajar la planilla, según dónde viva.
  let planillaLocal: string | null = null;
  if (config.ubicacionPlanilla === "maestraFija") {
    const remotoPlanilla = rclone.rutaRemota(config, config.rutaRemotaPlanillaMaestra);
    const existe = (await rclone.existeArchivo(remotoPlanilla, config)) !== null;
    if (existe) {
      planillaLocal = path.join(config.rutaTrabajoLocal, "planilla", path.basename(config.rutaRemotaPlanillaMaestra));
      fs.mkdirSync(path.dirname(planillaLocal), { recursive: true });
      await rclone.descargarArchivo(remotoPlanilla, planillaLocal, config);
    }
  }
  // Con `carpetaDelDia`, la planilla ya vino en la descarga de la carpeta y la ubica el
  // clasificador, igual que en modo montaje.

  return {
    rutas,
    carpetaLocalDelDia: carpetaLocal,
    planillaLocal,
    rutaRegistro:
      config.ubicacionRegistro === "local"
        ? rutaRegistroLocal(config, rutas.fecha)
        : path.join(carpetaLocal, config.nombreArchivoRegistro),
    requierePublicacion: true,
  };
}

export interface ResultadoPublicacion {
  publicado: boolean;
  destino?: string;
  tamanioLocal?: number;
  tamanioRemoto?: number;
  detalle?: string;
}

/**
 * Sube la planilla de vuelta al remoto, si hace falta.
 *
 * En modo montaje no hace nada: el archivo ya se escribió en su lugar definitivo.
 *
 * En modo rclone sube y **verifica**. Si la verificación falla, lanza error: es
 * deliberado, porque el modo de falla que queremos evitar es justamente el silencioso —
 * que el proceso dé por bueno un trabajo que nunca llegó al destino.
 */
export async function publicarPlanilla(
  espacio: EspacioDeTrabajo,
  rutaPlanillaLocal: string,
  opciones: { config?: Config; simular?: boolean } = {}
): Promise<ResultadoPublicacion> {
  const config = opciones.config ?? cargarConfig();
  const simular = opciones.simular ?? false;

  if (!espacio.requierePublicacion) {
    return { publicado: false, detalle: "modo sistemaArchivos: la planilla ya está en su lugar" };
  }

  const destino =
    config.ubicacionPlanilla === "maestraFija"
      ? rclone.rutaRemota(config, config.rutaRemotaPlanillaMaestra)
      : rclone.rutaRemota(
          config,
          config.rutaRemotaBase,
          espacio.rutas.anio,
          espacio.rutas.mes,
          espacio.rutas.dia,
          path.basename(rutaPlanillaLocal)
        );

  if (simular) {
    return { publicado: false, destino, detalle: "simulación: no se subió nada" };
  }

  const r = await rclone.subirArchivo(rutaPlanillaLocal, destino, config);
  return {
    publicado: true,
    destino,
    tamanioLocal: r.tamanioLocal,
    tamanioRemoto: r.tamanioRemoto,
  };
}

/** Sube el registro al remoto, si está configurado para vivir en la carpeta del día. */
export async function publicarRegistro(
  espacio: EspacioDeTrabajo,
  opciones: { config?: Config; simular?: boolean } = {}
): Promise<void> {
  const config = opciones.config ?? cargarConfig();
  const simular = opciones.simular ?? false;
  if (!espacio.requierePublicacion || config.ubicacionRegistro === "local" || simular) return;
  if (!fs.existsSync(espacio.rutaRegistro)) return;

  const destino = rclone.rutaRemota(
    config,
    config.rutaRemotaBase,
    espacio.rutas.anio,
    espacio.rutas.mes,
    espacio.rutas.dia,
    config.nombreArchivoRegistro
  );
  await rclone.subirArchivo(espacio.rutaRegistro, destino, config);
}

/** Borra una planilla duplicada, en el lugar que corresponda según el modo. */
export async function borrarPlanillaDuplicada(
  espacio: EspacioDeTrabajo,
  rutaLocal: string,
  opciones: { config?: Config; simular?: boolean } = {}
): Promise<void> {
  const config = opciones.config ?? cargarConfig();
  const simular = opciones.simular ?? false;
  if (simular) return;

  if (!espacio.requierePublicacion) {
    fs.unlinkSync(rutaLocal);
    return;
  }
  const destino = rclone.rutaRemota(
    config,
    config.rutaRemotaBase,
    espacio.rutas.anio,
    espacio.rutas.mes,
    espacio.rutas.dia,
    path.basename(rutaLocal)
  );
  await rclone.borrarArchivo(destino, config);
  fs.rmSync(rutaLocal, { force: true });
}

/**
 * Huella del contenido de la carpeta del día, para que el poller detecte cambios.
 *
 * En modo rclone se consulta el remoto directamente en vez de mirar la copia local: lo que
 * interesa es si Administración subió algo nuevo, no el estado de nuestra copia de trabajo.
 */
export async function huellaDelDia(
  fecha: Date = new Date(),
  opciones: { config?: Config } = {}
): Promise<string> {
  const config = opciones.config ?? cargarConfig();
  const rutas = calcularRutasDelDia(fecha, config);

  if (config.modoAcceso === "sistemaArchivos") {
    if (!fs.existsSync(rutas.rutaDia)) return "(no existe)";
    const partes: string[] = [];
    for (const nombre of fs.readdirSync(rutas.rutaDia).sort()) {
      if (nombre === config.nombreArchivoRegistro) continue;
      try {
        const s = fs.statSync(path.join(rutas.rutaDia, nombre));
        partes.push(`${nombre}:${s.size}:${Math.floor(s.mtimeMs)}`);
      } catch {
        partes.push(`${nombre}:?`);
      }
    }
    return partes.join("|");
  }

  const remotoDia = rclone.rutaRemota(config, config.rutaRemotaBase, rutas.anio, rutas.mes, rutas.dia);
  const contenido = await rclone.listar(remotoDia, config);
  return contenido
    .filter((a) => !a.esDirectorio && a.nombre !== config.nombreArchivoRegistro)
    .map((a) => `${a.nombre}:${a.tamanio}:${a.modificado.getTime()}`)
    .sort()
    .join("|");
}

/** Carpeta temporal para el modo simulación, que trabaja sobre una copia descartable. */
export function carpetaTemporal(prefijo = "bancosflow-"): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefijo));
}
