/**
 * rutasPlanillaBancos.ts
 *
 * Calcula y crea la estructura de carpetas donde Administración deja los archivos:
 *
 *   <rutaCarpetaBase>/<año>/<Mes>/<día>
 *   ej: .../PlanillaBancos/2026/Agosto/19
 *
 * Las carpetas se crean siempre, independientemente de dónde viva la planilla
 * (`ubicacionPlanilla`), porque los estados de cuenta se suben por día en los dos modos.
 */

import * as fs from "fs";
import * as path from "path";
import { cargarConfig, Config } from "./config";

export interface RutasDelDia {
  /** Fecha a la que corresponden estas rutas, como "yyyy-mm-dd" en la zona configurada. */
  fecha: string;
  anio: string;
  mes: string;
  dia: string;
  /** .../PlanillaBancos/2026 */
  rutaAnio: string;
  /** .../PlanillaBancos/2026/Agosto */
  rutaMes: string;
  /** .../PlanillaBancos/2026/Agosto/19 */
  rutaDia: string;
  /** Archivo de control dentro de la carpeta del día. */
  rutaRegistro: string;
}

/**
 * Devuelve las partes de la fecha (año, mes 1-12, día) en la zona horaria configurada.
 *
 * Se usa `Intl` en vez de los métodos locales de `Date` para no depender de cómo esté
 * configurado el reloj de la VM: si la máquina está en UTC, entre las 21:00 y la
 * medianoche de Uruguay `new Date().getDate()` ya devolvería el día siguiente, y el
 * proceso miraría una carpeta equivocada justo en el horario de cierre.
 */
export function partesDeFecha(fecha: Date, zonaHoraria: string): { anio: number; mes: number; dia: number } {
  const formateador = new Intl.DateTimeFormat("en-CA", {
    timeZone: zonaHoraria,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });
  const partes = formateador.formatToParts(fecha);
  const buscar = (tipo: string) => Number(partes.find((p) => p.type === tipo)?.value);
  return { anio: buscar("year"), mes: buscar("month"), dia: buscar("day") };
}

/** Calcula las rutas para una fecha dada (por defecto, hoy). No crea nada en disco. */
export function calcularRutasDelDia(fecha: Date = new Date(), configExplicita?: Config): RutasDelDia {
  const config = configExplicita ?? cargarConfig();
  const { anio, mes, dia } = partesDeFecha(fecha, config.zonaHoraria);

  const nombreAnio = String(anio);
  const nombreMes = config.nombresMeses[mes - 1];
  const nombreDia = config.diaConCeroAdelante ? String(dia).padStart(2, "0") : String(dia);

  const rutaAnio = path.join(config.rutaCarpetaBase, nombreAnio);
  const rutaMes = path.join(rutaAnio, nombreMes);
  const rutaDia = path.join(rutaMes, nombreDia);

  return {
    fecha: `${nombreAnio}-${String(mes).padStart(2, "0")}-${String(dia).padStart(2, "0")}`,
    anio: nombreAnio,
    mes: nombreMes,
    dia: nombreDia,
    rutaAnio,
    rutaMes,
    rutaDia,
    rutaRegistro: path.join(rutaDia, config.nombreArchivoRegistro),
  };
}

export interface ResultadoCreacion extends RutasDelDia {
  /** Qué carpetas hubo que crear realmente (las que ya existían no se listan). */
  creadas: string[];
  /** En modo simulación no se toca el disco. */
  simulado: boolean;
}

/**
 * Crea la estructura año/mes/día si falta. Es idempotente: si ya existe todo, no hace
 * nada y devuelve `creadas: []`. Pensado para correr todos los días a las 00:00 sin
 * riesgo de pisar nada.
 */
export function crearCarpetasDelDia(
  fecha: Date = new Date(),
  opciones: { simular?: boolean; config?: Config } = {}
): ResultadoCreacion {
  const config = opciones.config ?? cargarConfig();
  const simular = opciones.simular ?? false;
  const rutas = calcularRutasDelDia(fecha, config);

  if (!fs.existsSync(config.rutaCarpetaBase)) {
    throw new Error(
      `La carpeta base no existe o no está accesible: ${config.rutaCarpetaBase}\n` +
        `Si es un montaje de red, verificar que esté montado antes de correr el proceso.`
    );
  }

  const creadas: string[] = [];
  for (const ruta of [rutas.rutaAnio, rutas.rutaMes, rutas.rutaDia]) {
    if (fs.existsSync(ruta)) continue;
    creadas.push(ruta);
    if (!simular) {
      // `recursive: true` para que no falle si otro proceso la creó en el medio.
      fs.mkdirSync(ruta, { recursive: true });
    }
  }

  return { ...rutas, creadas, simulado: simular };
}

// --- Uso por consola:
//   node rutasPlanillaBancos.js [--simular] [--fecha yyyy-mm-dd]
if (require.main === module) {
  const args = process.argv.slice(2);
  const simular = args.includes("--simular");
  const indiceFecha = args.indexOf("--fecha");
  const fecha =
    indiceFecha !== -1 && args[indiceFecha + 1]
      ? new Date(`${args[indiceFecha + 1]}T12:00:00Z`) // mediodía UTC: evita saltos de día por zona horaria
      : new Date();

  try {
    const r = crearCarpetasDelDia(fecha, { simular });
    console.log(`Fecha: ${r.fecha}  ->  ${r.anio} / ${r.mes} / ${r.dia}`);
    console.log(`Carpeta del día: ${r.rutaDia}`);
    if (r.creadas.length === 0) {
      console.log("Todas las carpetas ya existían, no hubo nada que crear.");
    } else {
      console.log(simular ? "Se crearían:" : "Creadas:");
      r.creadas.forEach((c) => console.log(`  ${c}`));
    }
  } catch (err) {
    console.error("ERROR:", (err as Error).message);
    process.exit(1);
  }
}
