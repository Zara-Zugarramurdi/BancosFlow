/**
 * config.ts
 *
 * Único lugar donde viven los parámetros ajustables del proceso automático.
 *
 * Los valores por defecto de acá son los acordados para la VM de producción. Se
 * pueden pisar sin tocar el código, con un archivo JSON: por defecto se busca en
 * `config/bancosflow.config.json` (relativo a la raíz del proyecto), y se puede
 * apuntar a otro lado con la variable de entorno `BANCOSFLOW_CONFIG`.
 *
 * Sólo hace falta escribir en el JSON las claves que se quieran cambiar; el resto
 * toma el valor por defecto.
 */

import * as fs from "fs";
import * as path from "path";

export type UbicacionPlanilla = "carpetaDelDia" | "maestraFija";

/**
 * Cómo se llega a los archivos.
 *
 * - `sistemaArchivos`: la carpeta está montada (SMB del fileserver viejo, o `rclone mount`).
 *   Se lee y escribe directamente sobre ella.
 * - `rclone`: no hay nada montado. Se baja a una carpeta local de trabajo con `rclone copy`,
 *   se procesa ahí, y se sube con `rclone copy` verificando que haya llegado.
 *
 * Se agregó `rclone` porque el montaje FUSE contra SharePoint resultó inviable: SharePoint
 * retoca los archivos de Office al recibirlos, rclone lo interpreta como transferencia
 * corrupta y borra lo subido — todo de forma asincrónica y silenciosa, así que el proceso
 * reportaba éxito y la planilla nunca llegaba. Ver el comentario de `src/rclone.ts`.
 */
export type ModoAcceso = "sistemaArchivos" | "rclone";

export interface Config {
  /** Carpeta raíz sobre la que trabaja el proceso. Es la única que se toca. */
  rutaCarpetaBase: string;

  /**
   * Dónde buscar la planilla de bancos.
   *
   * - `carpetaDelDia`: la sube Administración junto con los estados de cuenta (modo actual).
   * - `maestraFija`: hay una sola planilla en `rutaPlanillaMaestra` y las carpetas del día
   *   contienen únicamente estados de cuenta.
   *
   * En ambos modos las carpetas por día se crean igual, porque los estados de cuenta se
   * suben todos los días de todas formas. Lo único que cambia es de dónde sale la planilla.
   */
  ubicacionPlanilla: UbicacionPlanilla;

  /** Ruta de la planilla maestra. Sólo se usa si `ubicacionPlanilla` es `maestraFija`. */
  rutaPlanillaMaestra: string;

  /** Dónde se guardan los respaldos. Fuera del fileserver, en disco local de la VM. */
  rutaBackups: string;

  /** Días que se conservan los respaldos antes de purgarlos. */
  retencionBackupsDias: number;

  /** Cada cuánto revisa el poller si cambió algo en la carpeta del día. */
  intervaloPollSegundos: number;

  /**
   * Cuánto tiene que estar quieta la carpeta (sin cambios) antes de procesar.
   * Da tiempo a que Administración termine de subir todos los archivos y evita
   * agarrar un archivo a medio copiar.
   */
  esperaSinCambiosSegundos: number;

  /** Hora a la que se crean las carpetas del día, en formato `HH:MM`. */
  horaCreacionCarpetas: string;

  /**
   * Zona horaria con la que se decide "qué día es hoy".
   *
   * Importa: si la VM está en UTC y no se fija esto, entre las 21:00 y las 00:00 de
   * Uruguay el proceso ya estaría creando y mirando la carpeta del día siguiente.
   */
  zonaHoraria: string;

  /** Nombres de los meses para las carpetas, en español. */
  nombresMeses: string[];

  /** Si el día lleva cero adelante (`09`) o no (`9`). Acordado: sin cero. */
  diaConCeroAdelante: boolean;

  /** Nombre del archivo de control de archivos ya procesados. */
  nombreArchivoRegistro: string;

  // --- Acceso a los archivos ---

  /** Cómo se llega a los archivos: carpeta montada o rclone explícito. */
  modoAcceso: ModoAcceso;

  /**
   * Dónde vive el registro de procesados.
   *
   * `local` lo guarda en `rutaTrabajoLocal/registros`, fuera del share. Es lo recomendado:
   * saca una escritura del remoto y elimina el `rename` sobre archivo existente, que es la
   * operación más frágil sobre montajes de red. Administración no necesita verlo.
   */
  ubicacionRegistro: "local" | "carpetaDelDia";

  // --- Sólo para modoAcceso = "rclone" ---

  /** Nombre del remoto configurado en rclone, con o sin los dos puntos. Ej: `sharepoint:` */
  remotoRclone: string;

  /** Ruta de la carpeta base DENTRO del remoto (sin el nombre del remoto). */
  rutaRemotaBase: string;

  /** Carpeta local donde se bajan los archivos para trabajar. Se limpia sola. */
  rutaTrabajoLocal: string;

  /** Ruta de la planilla maestra dentro del remoto. Sólo con `ubicacionPlanilla=maestraFija`. */
  rutaRemotaPlanillaMaestra: string;

  /** Binario de rclone. Se puede poner la ruta completa si no está en el PATH. */
  rcloneBinario: string;

  /** Flags para cualquier transferencia (bajada y subida). */
  flagsRcloneTransferencia: string[];

  /**
   * Flags adicionales sólo para subir.
   *
   * `--ignore-size` y `--ignore-checksum` son **imprescindibles** con SharePoint: retoca los
   * archivos de Office al recibirlos, así que el tamaño del destino nunca coincide con el del
   * origen (verificado: 1.200.869 bytes se convierten en 1.208.496, y en el intento siguiente
   * en 1.208.499). Sin estos flags rclone da `corrupted on transfer` y borra lo que subió.
   */
  flagsRcloneSubida: string[];

  /** Tiempo máximo para un comando de rclone, en segundos. */
  timeoutRcloneSegundos: number;

  /**
   * Archivo donde se registra la salida de TODOS los comandos de rclone, salgan bien o mal.
   *
   * Hasta ahora el `stderr` de rclone sólo se miraba cuando el comando fallaba; en el caso
   * exitoso se descartaba por completo, así que no quedaba constancia de cuántos intentos
   * hicieron falta ni de cuánto tardó. Eso dejó el diagnóstico a ciegas más de una vez.
   * Vacío desactiva el registro.
   */
  rutaLogRclone: string;
}

export const CONFIG_POR_DEFECTO: Config = {
  rutaCarpetaBase: "/media/windowsshare/contable/privado/ADMINISTRACION/PlanillaBancos",
  ubicacionPlanilla: "carpetaDelDia",
  rutaPlanillaMaestra: "",
  rutaBackups: "/home/teledata/backups",
  retencionBackupsDias: 90,
  intervaloPollSegundos: 60,
  esperaSinCambiosSegundos: 60,
  horaCreacionCarpetas: "00:00",
  zonaHoraria: "America/Montevideo",
  nombresMeses: [
    "Enero",
    "Febrero",
    "Marzo",
    "Abril",
    "Mayo",
    "Junio",
    "Julio",
    "Agosto",
    "Septiembre",
    "Octubre",
    "Noviembre",
    "Diciembre",
  ],
  diaConCeroAdelante: false,
  nombreArchivoRegistro: ".bancosflow.json",

  modoAcceso: "sistemaArchivos",
  ubicacionRegistro: "local",

  remotoRclone: "sharepoint:",
  rutaRemotaBase: "contable/privado/ADMINISTRACION/PlanillaBancos",
  rutaTrabajoLocal: "/home/teledata/bancosflow-trabajo",
  rutaRemotaPlanillaMaestra: "contable/privado/ADMINISTRACION/PlanillaBancos/PlanillaBancos.xlsx",
  rcloneBinario: "rclone",
  flagsRcloneTransferencia: [],
  flagsRcloneSubida: ["--ignore-size", "--ignore-checksum", "--ignore-times"],
  timeoutRcloneSegundos: 600,
  rutaLogRclone: "/home/teledata/bancosflow-trabajo/rclone.log",
};

function rutaConfigPorDefecto(): string {
  // `__dirname` apunta a dist/ cuando corre compilado, así que subimos un nivel.
  return path.join(__dirname, "..", "config", "bancosflow.config.json");
}

let cache: Config | null = null;

/**
 * Devuelve la configuración efectiva: los valores por defecto pisados por lo que
 * haya en el archivo JSON, si existe. Se cachea, así que leer la config muchas
 * veces no cuesta nada.
 */
export function cargarConfig(rutaExplicita?: string): Config {
  if (cache && !rutaExplicita) return cache;

  const ruta = rutaExplicita ?? process.env.BANCOSFLOW_CONFIG ?? rutaConfigPorDefecto();
  let config: Config = { ...CONFIG_POR_DEFECTO };

  if (fs.existsSync(ruta)) {
    let crudo: unknown;
    try {
      crudo = JSON.parse(fs.readFileSync(ruta, "utf8"));
    } catch (err) {
      throw new Error(
        `El archivo de configuración "${ruta}" no es JSON válido: ${(err as Error).message}`
      );
    }
    if (typeof crudo !== "object" || crudo === null || Array.isArray(crudo)) {
      throw new Error(`El archivo de configuración "${ruta}" debe contener un objeto JSON.`);
    }

    // Avisar de claves desconocidas: casi siempre son errores de tipeo que, en
    // silencio, dejarían el proceso corriendo con el valor por defecto.
    const conocidas = new Set(Object.keys(CONFIG_POR_DEFECTO));
    const desconocidas = Object.keys(crudo).filter((k) => !conocidas.has(k));
    if (desconocidas.length > 0) {
      console.warn(
        `AVISO: la configuración tiene claves desconocidas que se van a ignorar: ${desconocidas.join(", ")}`
      );
    }

    config = { ...config, ...(crudo as Partial<Config>) };
  }

  validar(config, ruta);
  if (!rutaExplicita) cache = config;
  return config;
}

function validar(config: Config, ruta: string): void {
  const problemas: string[] = [];

  if (!config.rutaCarpetaBase) problemas.push("rutaCarpetaBase no puede estar vacía");
  if (!config.rutaBackups) problemas.push("rutaBackups no puede estar vacía");

  if (config.ubicacionPlanilla !== "carpetaDelDia" && config.ubicacionPlanilla !== "maestraFija") {
    problemas.push(`ubicacionPlanilla debe ser "carpetaDelDia" o "maestraFija"`);
  }
  // Con acceso por sistema de archivos hace falta la ruta local; con rclone, la remota
  // (que se valida más abajo). Por eso el chequeo depende del modo de acceso.
  if (
    config.ubicacionPlanilla === "maestraFija" &&
    config.modoAcceso === "sistemaArchivos" &&
    !config.rutaPlanillaMaestra
  ) {
    problemas.push(
      `con ubicacionPlanilla="maestraFija" y modoAcceso="sistemaArchivos" hay que definir rutaPlanillaMaestra`
    );
  }

  if (!(config.retencionBackupsDias > 0)) problemas.push("retencionBackupsDias debe ser mayor a 0");
  if (!(config.intervaloPollSegundos > 0)) problemas.push("intervaloPollSegundos debe ser mayor a 0");
  if (!(config.esperaSinCambiosSegundos > 0)) problemas.push("esperaSinCambiosSegundos debe ser mayor a 0");

  if (!/^\d{1,2}:\d{2}$/.test(config.horaCreacionCarpetas)) {
    problemas.push(`horaCreacionCarpetas debe tener formato HH:MM (recibido: "${config.horaCreacionCarpetas}")`);
  }
  if (config.nombresMeses.length !== 12) {
    problemas.push(`nombresMeses debe tener exactamente 12 elementos (recibido: ${config.nombresMeses.length})`);
  }
  if (!config.nombreArchivoRegistro) problemas.push("nombreArchivoRegistro no puede estar vacío");

  if (config.modoAcceso !== "sistemaArchivos" && config.modoAcceso !== "rclone") {
    problemas.push(`modoAcceso debe ser "sistemaArchivos" o "rclone"`);
  }
  if (config.ubicacionRegistro !== "local" && config.ubicacionRegistro !== "carpetaDelDia") {
    problemas.push(`ubicacionRegistro debe ser "local" o "carpetaDelDia"`);
  }
  if (config.modoAcceso === "rclone") {
    if (!config.remotoRclone) problemas.push("con modoAcceso=rclone hay que definir remotoRclone");
    if (!config.rutaRemotaBase) problemas.push("con modoAcceso=rclone hay que definir rutaRemotaBase");
    if (!config.rutaTrabajoLocal) problemas.push("con modoAcceso=rclone hay que definir rutaTrabajoLocal");
    if (config.ubicacionPlanilla === "maestraFija" && !config.rutaRemotaPlanillaMaestra) {
      problemas.push("con modoAcceso=rclone y maestraFija hay que definir rutaRemotaPlanillaMaestra");
    }
    if (!(config.timeoutRcloneSegundos > 0)) problemas.push("timeoutRcloneSegundos debe ser mayor a 0");
  }

  // Validar la zona horaria contra el propio motor de Intl, en vez de una lista fija.
  try {
    new Intl.DateTimeFormat("es-UY", { timeZone: config.zonaHoraria });
  } catch {
    problemas.push(`zonaHoraria "${config.zonaHoraria}" no es una zona horaria válida`);
  }

  if (problemas.length > 0) {
    throw new Error(
      `Configuración inválida (${ruta}):\n` + problemas.map((p) => `  - ${p}`).join("\n")
    );
  }
}

// --- Uso por consola: muestra la configuración efectiva, para verificar en la VM ---
if (require.main === module) {
  const ruta = process.env.BANCOSFLOW_CONFIG ?? rutaConfigPorDefecto();
  console.log(`Archivo de configuración: ${ruta}${fs.existsSync(ruta) ? "" : "  (no existe: se usan los valores por defecto)"}`);
  console.log(JSON.stringify(cargarConfig(), null, 2));
}
