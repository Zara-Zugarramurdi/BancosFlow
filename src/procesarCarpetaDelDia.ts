/**
 * procesarCarpetaDelDia.ts
 *
 * Orquestador: junta todas las piezas y procesa la carpeta de un día.
 *
 *   1. Ubica la carpeta del día y la planilla.
 *   2. Clasifica los archivos (planilla / estados de cuenta / ignorados).
 *   3. Descarta los estados de cuenta ya procesados, según el registro.
 *   4. Respalda la planilla (sólo si hay algo para procesar).
 *   5. Corre `actualizarDesdeUltimaFecha` por cada estado de cuenta pendiente.
 *   6. Anota lo hecho en el registro y purga respaldos viejos.
 *
 * Tiene modo simulación (`--dry-run`) que informa exactamente qué haría sin escribir
 * absolutamente nada. Conviene correr así los primeros días.
 */

import * as fs from "fs";
import * as path from "path";
import { cargarConfig, Config } from "./config";
import { calcularRutasDelDia, crearCarpetasDelDia } from "./rutasPlanillaBancos";
import { clasificarCarpeta, Clasificacion } from "./clasificarArchivos";
import {
  leerRegistro,
  guardarRegistro,
  hashDeArchivo,
  yaProcesado,
  anotarProcesado,
} from "./registroProcesados";
import { respaldarPlanilla, purgarRespaldosViejos } from "./respaldo";
import { actualizarDesdeUltimaFecha } from "./actualizarDesdeUltimaFecha";

/** Archivo que, si existe en la carpeta base, frena el proceso sin tocar nada. */
const NOMBRE_ARCHIVO_PAUSA = "PAUSADO";

export type MotivoNoProcesado =
  | "pausado"
  | "carpeta-inexistente"
  | "archivos-abiertos"
  | "sin-planilla"
  | "sin-estados-de-cuenta"
  | "todo-procesado";

export interface ResultadoCuenta {
  archivo: string;
  cuenta: string;
  agregados: number;
  omitidos: number;
  filas?: string;
  avisos: string[];
  error?: string;
}

export interface ResultadoCarpeta {
  fecha: string;
  rutaDia: string;
  procesado: boolean;
  motivo?: MotivoNoProcesado;
  planilla?: string;
  resultados: ResultadoCuenta[];
  rutaRespaldo?: string;
  simulado: boolean;
}

function hayPausa(config: Config): boolean {
  return fs.existsSync(path.join(config.rutaCarpetaBase, NOMBRE_ARCHIVO_PAUSA));
}

/**
 * Ubica la planilla a usar, según la configuración.
 * En modo `maestraFija` no se busca en la carpeta del día: se usa siempre la misma.
 */
function ubicarPlanilla(config: Config, clasificacion: Clasificacion): string | null {
  if (config.ubicacionPlanilla === "maestraFija") {
    return fs.existsSync(config.rutaPlanillaMaestra) ? config.rutaPlanillaMaestra : null;
  }
  return clasificacion.planilla ? clasificacion.planilla.ruta : null;
}

function fechaATexto(d: Date): string {
  return `${String(d.getUTCDate()).padStart(2, "0")}/${String(d.getUTCMonth() + 1).padStart(2, "0")}/${d.getUTCFullYear()}`;
}

export async function procesarCarpetaDelDia(
  opciones: { fecha?: Date; simular?: boolean; config?: Config } = {}
): Promise<ResultadoCarpeta> {
  const config = opciones.config ?? cargarConfig();
  const simular = opciones.simular ?? false;
  const fecha = opciones.fecha ?? new Date();

  const rutas = calcularRutasDelDia(fecha, config);
  const base: ResultadoCarpeta = {
    fecha: rutas.fecha,
    rutaDia: rutas.rutaDia,
    procesado: false,
    resultados: [],
    simulado: simular,
  };

  // Freno de mano: permite parar el proceso sin acceso a la VM, dejando un archivo
  // llamado PAUSADO en la carpeta base.
  if (hayPausa(config)) {
    return { ...base, motivo: "pausado" };
  }

  if (!fs.existsSync(rutas.rutaDia)) {
    return { ...base, motivo: "carpeta-inexistente" };
  }

  const clasificacion = clasificarCarpeta(rutas.rutaDia, config);

  // Si alguien tiene un Excel abierto, esperamos al próximo ciclo: escribir la planilla
  // mientras está abierta puede terminar en que la persona guarde encima de lo insertado.
  if (clasificacion.hayArchivosAbiertos) {
    return { ...base, motivo: "archivos-abiertos" };
  }

  if (clasificacion.estadosDeCuenta.length === 0) {
    return { ...base, motivo: "sin-estados-de-cuenta" };
  }

  const rutaPlanilla = ubicarPlanilla(config, clasificacion);
  if (!rutaPlanilla) {
    // Se espera: procesar contra la planilla equivocada es peor que no procesar.
    return { ...base, motivo: "sin-planilla" };
  }

  // Filtrar lo que ya se procesó, comparando por contenido y no sólo por nombre.
  const registro = leerRegistro(rutas.rutaRegistro, rutas.fecha);
  const pendientes = clasificacion.estadosDeCuenta
    .map((e) => ({ ...e, hash: hashDeArchivo(e.ruta) }))
    .filter((e) => !yaProcesado(registro, e.nombre, e.hash));

  if (pendientes.length === 0) {
    return { ...base, planilla: path.basename(rutaPlanilla), motivo: "todo-procesado" };
  }

  // Respaldo: sólo ahora que sabemos que hay algo real para hacer.
  const respaldo = respaldarPlanilla(rutaPlanilla, { simular, config });

  // Descartar planillas viejas duplicadas, si las hubiera.
  for (const vieja of clasificacion.planillasDuplicadas) {
    if (!simular) fs.unlinkSync(vieja.ruta);
  }

  const resultados: ResultadoCuenta[] = [];
  for (const estado of pendientes) {
    const avisos: string[] = [];
    try {
      if (simular) {
        // En simulación se procesa contra una copia temporal, así se puede informar
        // exactamente qué se agregaría sin tocar la planilla real.
        const copia = path.join(
          fs.mkdtempSync(path.join(require("os").tmpdir(), "bancosflow-")),
          path.basename(rutaPlanilla)
        );
        fs.copyFileSync(rutaPlanilla, copia);
        const r = await actualizarDesdeUltimaFecha(copia, estado.ruta, copia);
        if (r.diasSinCobertura.length > 0) {
          avisos.push(
            `no cubre ${r.diasSinCobertura.length} día(s): ${r.diasSinCobertura.map(fechaATexto).join(", ")}`
          );
        }
        resultados.push({
          archivo: estado.nombre,
          cuenta: estado.cuenta.etiqueta,
          agregados: r.agregados.length,
          omitidos: r.omitidosPorDuplicado.length,
          filas: r.agregados.length > 0 ? `${r.filaInicial}-${r.filaFinal}` : undefined,
          avisos,
        });
        fs.rmSync(path.dirname(copia), { recursive: true, force: true });
        continue;
      }

      const r = await actualizarDesdeUltimaFecha(rutaPlanilla, estado.ruta, rutaPlanilla);
      if (r.diasSinCobertura.length > 0) {
        avisos.push(
          `no cubre ${r.diasSinCobertura.length} día(s): ${r.diasSinCobertura.map(fechaATexto).join(", ")}`
        );
      }
      const filas = r.agregados.length > 0 ? `${r.filaInicial}-${r.filaFinal}` : undefined;

      anotarProcesado(registro, {
        archivo: estado.nombre,
        hash: estado.hash,
        cuenta: estado.cuenta.etiqueta,
        procesadoEn: new Date().toISOString(),
        agregados: r.agregados.length,
        omitidos: r.omitidosPorDuplicado.length,
        filas,
      });

      resultados.push({
        archivo: estado.nombre,
        cuenta: estado.cuenta.etiqueta,
        agregados: r.agregados.length,
        omitidos: r.omitidosPorDuplicado.length,
        filas,
        avisos,
      });
    } catch (err) {
      // Un archivo que falla no debe impedir que se procesen los demás: se anota el
      // error y se sigue. Como no se registra como procesado, se reintenta solo en el
      // próximo ciclo.
      resultados.push({
        archivo: estado.nombre,
        cuenta: estado.cuenta.etiqueta,
        agregados: 0,
        omitidos: 0,
        avisos,
        error: (err as Error).message,
      });
    }
  }

  if (!simular) {
    guardarRegistro(rutas.rutaRegistro, registro);
    purgarRespaldosViejos({ config });
  }

  return {
    ...base,
    procesado: true,
    planilla: path.basename(rutaPlanilla),
    resultados,
    rutaRespaldo: respaldo.rutaRespaldo,
  };
}

const EXPLICACION_MOTIVOS: Record<MotivoNoProcesado, string> = {
  pausado: `hay un archivo "${NOMBRE_ARCHIVO_PAUSA}" en la carpeta base: el proceso está frenado a propósito`,
  "carpeta-inexistente": "la carpeta del día todavía no existe",
  "archivos-abiertos": "hay un archivo abierto en Excel; se espera al próximo ciclo",
  "sin-planilla": "todavía no se subió la planilla de bancos; se espera",
  "sin-estados-de-cuenta": "todavía no se subió ningún estado de cuenta",
  "todo-procesado": "todos los estados de cuenta de la carpeta ya estaban procesados",
};

export function imprimirResultado(r: ResultadoCarpeta): void {
  const etiqueta = r.simulado ? "[SIMULACIÓN] " : "";
  console.log(`${etiqueta}Fecha ${r.fecha}  ·  ${r.rutaDia}`);

  if (!r.procesado) {
    console.log(`Sin procesar: ${EXPLICACION_MOTIVOS[r.motivo!]}`);
    return;
  }

  console.log(`Planilla: ${r.planilla}`);
  if (r.rutaRespaldo) {
    console.log(`Respaldo: ${r.rutaRespaldo}${r.simulado ? "  (no se copió)" : ""}`);
  }
  console.log("");

  let totalAgregados = 0;
  for (const c of r.resultados) {
    if (c.error) {
      console.log(`  ${c.cuenta.padEnd(18)} ERROR: ${c.error}`);
      continue;
    }
    totalAgregados += c.agregados;
    console.log(
      `  ${c.cuenta.padEnd(18)} ${String(c.agregados).padStart(3)} agregados` +
        `${c.filas ? ` (filas ${c.filas})` : ""}, ${c.omitidos} omitidos`
    );
    c.avisos.forEach((a) => console.log(`     AVISO: ${a}`));
  }
  console.log(`\nTotal agregado: ${totalAgregados} movimiento(s)${r.simulado ? " (simulado, no se escribió nada)" : ""}`);
}

// --- Uso por consola:
//   node procesarCarpetaDelDia.js [--dry-run] [--fecha yyyy-mm-dd] [--crear-carpetas]
if (require.main === module) {
  const args = process.argv.slice(2);
  const simular = args.includes("--dry-run") || args.includes("--simular");
  const crearCarpetas = args.includes("--crear-carpetas");
  const i = args.indexOf("--fecha");
  const fecha = i !== -1 && args[i + 1] ? new Date(`${args[i + 1]}T12:00:00Z`) : new Date();

  (async () => {
    try {
      if (crearCarpetas) crearCarpetasDelDia(fecha, { simular });
      imprimirResultado(await procesarCarpetaDelDia({ fecha, simular }));
    } catch (err) {
      console.error("ERROR:", (err as Error).message);
      process.exit(1);
    }
  })();
}
