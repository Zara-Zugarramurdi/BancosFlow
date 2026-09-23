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
import { crearCarpetasDelDia } from "./rutasPlanillaBancos";
import {
  prepararEspacioDeTrabajo,
  publicarPlanilla,
  publicarRegistro,
  borrarPlanillaDuplicada,
  carpetaTemporal,
} from "./almacenamiento";
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

export interface ResultadoPublicacionCarpeta {
  publicado: boolean;
  destino?: string;
  tamanioLocal?: number;
  tamanioRemoto?: number;
}

export interface ResultadoCarpeta {
  fecha: string;
  rutaDia: string;
  procesado: boolean;
  motivo?: MotivoNoProcesado;
  planilla?: string;
  resultados: ResultadoCuenta[];
  rutaRespaldo?: string;
  publicacion?: ResultadoPublicacionCarpeta;
  simulado: boolean;
}

function hayPausa(config: Config): boolean {
  return fs.existsSync(path.join(config.rutaCarpetaBase, NOMBRE_ARCHIVO_PAUSA));
}

/**
 * Ubica la planilla a usar, según la configuración.
 * En modo `maestraFija` no se busca en la carpeta del día: se usa siempre la misma.
 */
function ubicarPlanilla(
  config: Config,
  clasificacion: Clasificacion,
  planillaDelEspacio: string | null
): string | null {
  if (config.ubicacionPlanilla === "maestraFija") {
    // La resuelve `prepararEspacioDeTrabajo`: en modo rclone ya viene descargada.
    return planillaDelEspacio;
  }
  return clasificacion.planilla ? clasificacion.planilla.ruta : null;
}

function fechaATexto(d: Date): string {
  return `${String(d.getUTCDate()).padStart(2, "0")}/${String(d.getUTCMonth() + 1).padStart(2, "0")}/${d.getUTCFullYear()}`;
}

export async function procesarCarpetaDelDia(
  opciones: { fecha?: Date; simular?: boolean; config?: Config; soloFecha?: Date } = {}
): Promise<ResultadoCarpeta> {
  const config = opciones.config ?? cargarConfig();
  const simular = opciones.simular ?? false;
  const fecha = opciones.fecha ?? new Date();

  const espacio = await prepararEspacioDeTrabajo(fecha, { config, simular });
  const rutas = espacio.rutas;
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

  if (!fs.existsSync(espacio.carpetaLocalDelDia)) {
    return { ...base, motivo: "carpeta-inexistente" };
  }

  const clasificacion = clasificarCarpeta(espacio.carpetaLocalDelDia, config);

  // Si alguien tiene el libro abierto (Excel o LibreOffice), esperamos al próximo ciclo:
  // escribir la planilla mientras está abierta puede terminar en que la persona guarde
  // encima de lo que insertamos.
  if (clasificacion.hayArchivosAbiertos) {
    return { ...base, motivo: "archivos-abiertos" };
  }

  if (clasificacion.estadosDeCuenta.length === 0) {
    return { ...base, motivo: "sin-estados-de-cuenta" };
  }

  const rutaPlanilla = ubicarPlanilla(config, clasificacion, espacio.planillaLocal);
  if (!rutaPlanilla) {
    // Se espera: procesar contra la planilla equivocada es peor que no procesar.
    return { ...base, motivo: "sin-planilla" };
  }

  // Filtrar lo que ya se procesó, comparando por contenido y no sólo por nombre.
  const registro = leerRegistro(espacio.rutaRegistro, rutas.fecha);
  // En modo día puntual se ignora el registro: se está cargando a propósito un día
  // atrasado de un archivo que puede figurar como ya procesado. La deduplicación por
  // fecha+tipo+monto sigue evitando que se dupliquen movimientos.
  const pendientes = clasificacion.estadosDeCuenta
    .map((e) => ({ ...e, hash: hashDeArchivo(e.ruta) }))
    .filter((e) => opciones.soloFecha !== undefined || !yaProcesado(registro, e.nombre, e.hash));

  if (pendientes.length === 0) {
    return { ...base, planilla: path.basename(rutaPlanilla), motivo: "todo-procesado" };
  }

  // Respaldo: sólo ahora que sabemos que hay algo real para hacer.
  const respaldo = respaldarPlanilla(rutaPlanilla, { simular, config });

  // Descartar planillas viejas duplicadas, si las hubiera.
  for (const vieja of clasificacion.planillasDuplicadas) {
    await borrarPlanillaDuplicada(espacio, vieja.ruta, { config, simular });
  }

  const resultados: ResultadoCuenta[] = [];
  for (const estado of pendientes) {
    const avisos: string[] = [];
    try {
      if (simular) {
        // En simulación se procesa contra una copia temporal, así se puede informar
        // exactamente qué se agregaría sin tocar la planilla real.
        const copia = path.join(carpetaTemporal(), path.basename(rutaPlanilla));
        fs.copyFileSync(rutaPlanilla, copia);
        const r = await actualizarDesdeUltimaFecha(copia, estado.ruta, copia, { soloFecha: opciones.soloFecha });
        if (r.diasSinCobertura.length > 0) {
          avisos.push(
            `no cubre ${r.diasSinCobertura.length} día(s): ${r.diasSinCobertura.map(fechaATexto).join(", ")}`
          );
        }
        if (r.diasFaltantesAlFinal.length > 0) {
          avisos.push(
            `llega sólo hasta el ${fechaATexto(r.ultimaFechaDelEstado!)}; ` +
              `falta(n) ${r.diasFaltantesAlFinal.length} día(s) hasta hoy`
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

      const r = await actualizarDesdeUltimaFecha(rutaPlanilla, estado.ruta, rutaPlanilla, { soloFecha: opciones.soloFecha });
      if (r.diasSinCobertura.length > 0) {
        avisos.push(
          `no cubre ${r.diasSinCobertura.length} día(s): ${r.diasSinCobertura.map(fechaATexto).join(", ")}`
        );
      }
      if (r.diasFaltantesAlFinal.length > 0) {
        avisos.push(
          `llega sólo hasta el ${fechaATexto(r.ultimaFechaDelEstado!)}; ` +
            `falta(n) ${r.diasFaltantesAlFinal.length} día(s) hasta hoy`
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

  // A partir de acá puede fallar la publicación. La purga de respaldos va en un `finally`
  // más abajo para que se ejecute igual: si quedara sólo en el camino feliz, un remoto
  // caído dejaría de purgar justo cuando más respaldos se están generando.
  try {
  // Publicar la planilla ANTES de anotar el registro: si la subida falla, el registro no
  // se escribe y el próximo ciclo reintenta. Al revés quedaría anotado como hecho algo que
  // nunca llegó al destino, que es exactamente el modo de falla que se quiere evitar.
  let publicacion: ResultadoPublicacionCarpeta | undefined;
  const huboCambios = resultados.some((r) => r.agregados > 0);
  if (huboCambios) {
    const p = await publicarPlanilla(espacio, rutaPlanilla, { config, simular });
    publicacion = {
      publicado: p.publicado,
      destino: p.destino,
      tamanioLocal: p.tamanioLocal,
      tamanioRemoto: p.tamanioRemoto,
    };
  }

  if (!simular) {
    guardarRegistro(espacio.rutaRegistro, registro);
    await publicarRegistro(espacio, { config, simular });
  }

  return {
    ...base,
    procesado: true,
    planilla: path.basename(rutaPlanilla),
    resultados,
    rutaRespaldo: respaldo.omitido ? undefined : respaldo.rutaRespaldo,
    publicacion,
  };
  } finally {
    if (!simular) purgarRespaldosViejos({ config });
  }
}

const EXPLICACION_MOTIVOS: Record<MotivoNoProcesado, string> = {
  pausado: `hay un archivo "${NOMBRE_ARCHIVO_PAUSA}" en la carpeta base: el proceso está frenado a propósito`,
  "carpeta-inexistente": "la carpeta del día todavía no existe",
  "archivos-abiertos": "alguien tiene un libro abierto (Excel o LibreOffice); se espera al próximo ciclo",
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

  if (r.publicacion?.publicado) {
    console.log(
      `Planilla publicada en ${r.publicacion.destino}` +
        ` (local ${r.publicacion.tamanioLocal} bytes, remoto ${r.publicacion.tamanioRemoto} bytes)`
    );
    console.log("El tamaño remoto difiere del local porque SharePoint retoca los archivos de Office; es esperable.");
  }
}

// --- Uso por consola:
//   node procesarCarpetaDelDia.js [--dry-run] [--fecha yyyy-mm-dd] [--crear-carpetas]
if (require.main === module) {
  const args = process.argv.slice(2);
  const simular = args.includes("--dry-run") || args.includes("--simular");
  const crearCarpetas = args.includes("--crear-carpetas");
  const i = args.indexOf("--fecha");
  const fecha = i !== -1 && args[i + 1] ? new Date(`${args[i + 1]}T12:00:00Z`) : new Date();

  // `--cargar-dia dd/mm/yyyy` procesa SÓLO los movimientos de ese día, ignorando hasta
  // dónde está cargada la planilla. Es para recuperar un día que quedó sin cargar: el
  // flujo normal mira hacia adelante y nunca lo levantaría.
  const j = args.indexOf("--cargar-dia");
  let soloFecha: Date | undefined;
  if (j !== -1 && args[j + 1]) {
    const [dd, mm, yyyy] = args[j + 1].split("/").map(Number);
    soloFecha = new Date(Date.UTC(yyyy, mm - 1, dd));
  }

  (async () => {
    try {
      if (crearCarpetas) crearCarpetasDelDia(fecha, { simular });
      imprimirResultado(await procesarCarpetaDelDia({ fecha, simular, soloFecha }));
    } catch (err) {
      console.error("ERROR:", (err as Error).message);
      process.exit(1);
    }
  })();
}
