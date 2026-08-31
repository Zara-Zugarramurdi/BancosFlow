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

  /** Nombre del archivo de control dentro de cada carpeta del día. */
  nombreArchivoRegistro: string;
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
  if (config.ubicacionPlanilla === "maestraFija" && !config.rutaPlanillaMaestra) {
    problemas.push(`con ubicacionPlanilla="maestraFija" hay que definir rutaPlanillaMaestra`);
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
