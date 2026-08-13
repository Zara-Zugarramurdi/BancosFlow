/**
 * actualizarDesdeUltimaFecha.ts
 *
 * Variante de `actualizarPlanilla` que no recibe una fecha: mira hasta qué día
 * está cargada la planilla y agrega todo lo que el estado de cuenta tenga de ahí
 * en adelante, en una sola pasada.
 *
 * Reutiliza el núcleo de inserción de `actualizarPlanilla` (`aplicarMovimientosAPlanilla`),
 * así que hereda tal cual la deduplicación, el cálculo del ancla, el desarmado de
 * fórmulas compartidas, el corrimiento de referencias, la copia de estilos y el
 * resaltado en amarillo. Lo único propio de este archivo es *qué* movimientos elegir
 * y *en qué orden* entregarlos.
 */

import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { identificarCuenta } from "./accountIdentifier";
import { parseBrouConFiltro } from "./parseBrou";
import { parseSantanderConFiltro } from "./parseSantander";
import {
  aplicarMovimientosAPlanilla,
  MAPA_HOJAS,
  encontrarFilaAncla,
  celdaFechaAClave,
  COL,
} from "./actualizarPlanilla";
import type { ResultadoActualizacion } from "./actualizarPlanilla";
import type { CuentaIdentificada, MovimientoLimpio } from "./types";

/**
 * Tolerancia para el chequeo de cordura de la fecha del ancla. Si la última fecha
 * cargada está más de estos días por delante de la fecha anterior del ledger,
 * sospechamos un error de tipeo y frenamos en vez de arriesgarnos.
 *
 * No es paranoia: en la planilla real hay filas con fechas de 2027 (error de tipeo
 * por 2023) y hasta una de 1928. Si una de esas cayera justo en el ancla, sin este
 * chequeo el proceso tomaría esa fecha como "hasta acá está cargado" y se saltearía
 * meses de movimientos en silencio.
 */
const MAX_SALTO_DIAS = 90;

export interface ResultadoActualizacionDesdeUltimaFecha extends ResultadoActualizacion {
  /** Última fecha que ya estaba cargada en la planilla (la de la fila ancla). */
  ultimaFechaEnPlanilla: Date;
  /** Fecha a partir de la cual se tomaron movimientos del estado de cuenta (inclusive). */
  fechaDesde: Date;
  /** Fechas que el estado de cuenta no cubre entre la planilla y su primer movimiento. */
  diasSinCobertura: Date[];
  /** Cantidad de movimientos agregados, desglosada por día. */
  agregadosPorDia: Array<{ fecha: string; cantidad: number }>;
}

function fechaATextoDDMMYYYY(d: Date): string {
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}/${d.getUTCFullYear()}`;
}

function soloDiaUTC(d: Date): Date {
  return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
}

function sumarDias(d: Date, dias: number): Date {
  return new Date(d.getTime() + dias * 24 * 60 * 60 * 1000);
}

function diferenciaEnDias(a: Date, b: Date): number {
  return Math.round((a.getTime() - b.getTime()) / (24 * 60 * 60 * 1000));
}

/**
 * Determina hasta qué fecha está cargada la hoja, y valida que esa fecha sea creíble.
 *
 * Usa la fecha de la FILA ANCLA (la última fila del bloque contiguo del ledger), no
 * el máximo de las fechas. Es deliberado: la posición manda, no el valor. En la
 * planilla real la fecha máxima de `BROU $` es 28/02/2027 — un error de tipeo por
 * 2023, ya que la fila siguiente vuelve a 01/03/2023. Si preguntáramos "cuál es la
 * fecha más alta", el proceso concluiría que la planilla está al día hasta 2027 y no
 * volvería a agregar nada nunca.
 */
function detectarUltimaFecha(hoja: ExcelJS.Worksheet, nombreHoja: string): { fecha: Date; filaAncla: number } {
  const filaAncla = encontrarFilaAncla(hoja);

  const claveAncla = celdaFechaAClave(hoja.getCell(filaAncla, COL.FECHA).value);
  if (!claveAncla) {
    throw new Error(
      `La última fila del ledger de la hoja "${nombreHoja}" (fila ${filaAncla}) no tiene una fecha legible, ` +
        `así que no se puede determinar hasta qué día está cargada la planilla. Revisar esa fila.`
    );
  }
  const [y, m, d] = claveAncla.split("-").map(Number);
  const fecha = new Date(Date.UTC(y, m - 1, d));

  // Cordura 1: la última fecha cargada no puede estar en el futuro.
  const hoy = soloDiaUTC(new Date());
  if (fecha.getTime() > hoy.getTime()) {
    throw new Error(
      `La última fecha cargada en la hoja "${nombreHoja}" (fila ${filaAncla}) es ${fechaATextoDDMMYYYY(fecha)}, ` +
        `posterior a hoy (${fechaATextoDDMMYYYY(hoy)}). Probablemente sea un error de tipeo en esa fila. ` +
        `Corregirla antes de correr el proceso, para no saltearse movimientos por error.`
    );
  }

  // Cordura 2: no puede haber un salto enorme respecto de la fila con fecha anterior.
  for (let fila = filaAncla - 1; fila >= 2; fila--) {
    const clavePrevia = celdaFechaAClave(hoja.getCell(fila, COL.FECHA).value);
    if (!clavePrevia) continue;
    const [py, pm, pd] = clavePrevia.split("-").map(Number);
    const fechaPrevia = new Date(Date.UTC(py, pm - 1, pd));
    const salto = diferenciaEnDias(fecha, fechaPrevia);
    if (salto > MAX_SALTO_DIAS) {
      throw new Error(
        `La última fecha cargada en la hoja "${nombreHoja}" (fila ${filaAncla}) es ${fechaATextoDDMMYYYY(fecha)}, ` +
          `pero la fila anterior con fecha (fila ${fila}) es ${fechaATextoDDMMYYYY(fechaPrevia)}: un salto de ${salto} días. ` +
          `Suele indicar un error de tipeo en la fecha. Revisar antes de continuar.`
      );
    }
    break; // sólo nos interesa la fila con fecha inmediatamente anterior
  }

  return { fecha, filaAncla };
}

/**
 * Ordena los movimientos cronológicamente, de más viejo a más nuevo, manteniendo
 * estable el orden dentro de un mismo día.
 *
 * Hace falta porque BROU entrega sus estados de cuenta con el movimiento MÁS NUEVO
 * PRIMERO (verificado con archivos reales: 10/08 → 07/08 → ... → 06/08). Cuando se
 * procesaba un solo día daba igual; al insertar un rango de varios días, respetar el
 * orden del archivo dejaría el ledger al revés.
 *
 * Dentro de un mismo día se conserva el orden en que vienen en el archivo: es el
 * único criterio disponible, ya que los estados de cuenta no traen una hora.
 */
function ordenarCronologicamente(movimientos: MovimientoLimpio[]): MovimientoLimpio[] {
  return movimientos
    .map((mov, indice) => ({ mov, indice }))
    .sort((a, b) => {
      const diferencia = a.mov.fecha.getTime() - b.mov.fecha.getTime();
      return diferencia !== 0 ? diferencia : a.indice - b.indice;
    })
    .map(({ mov }) => mov);
}

/**
 * Agrega a la planilla todos los movimientos del estado de cuenta desde la última
 * fecha ya cargada en adelante.
 *
 * El rango arranca en la última fecha cargada **inclusive**, no en el día siguiente.
 * Es a propósito: si ese día se había cargado a mitad de jornada, o el banco sumó
 * movimientos tarde, arrancar al día siguiente los perdería para siempre y en
 * silencio. Al incluirlo, la deduplicación (fecha + tipo + monto, contando
 * repeticiones) descarta los que ya están y agrega sólo los que faltaban. Se puede
 * cambiar a estricto con `incluirUltimaFecha: false`.
 */
export async function actualizarDesdeUltimaFecha(
  rutaPlanilla: string,
  rutaEstadoDeCuenta: string,
  rutaSalida: string = rutaPlanilla,
  opciones: { incluirUltimaFecha?: boolean } = {}
): Promise<ResultadoActualizacionDesdeUltimaFecha> {
  const incluirUltimaFecha = opciones.incluirUltimaFecha ?? true;

  // 1. Identificar la cuenta a partir del estado de cuenta.
  const workbookEstado = XLSX.readFile(rutaEstadoDeCuenta);
  const cuenta: CuentaIdentificada = identificarCuenta(workbookEstado);
  const nombreHoja = MAPA_HOJAS[cuenta.cuentaKey];

  // 2. Abrir la planilla sólo para averiguar hasta qué fecha está cargada.
  //    (`aplicarMovimientosAPlanilla` la vuelve a abrir por su cuenta; se prefiere
  //    esa pequeña ineficiencia antes que complicar la firma del núcleo compartido.)
  const workbookLectura = new ExcelJS.Workbook();
  await workbookLectura.xlsx.readFile(rutaPlanilla);
  const hoja = workbookLectura.getWorksheet(nombreHoja);
  if (!hoja) {
    throw new Error(
      `No se encontró la hoja "${nombreHoja}" en la planilla. Hojas disponibles: ${workbookLectura.worksheets
        .map((h) => h.name)
        .join(", ")}`
    );
  }

  const { fecha: ultimaFechaEnPlanilla } = detectarUltimaFecha(hoja, nombreHoja);
  const fechaDesde = incluirUltimaFecha ? ultimaFechaEnPlanilla : sumarDias(ultimaFechaEnPlanilla, 1);

  // 3. Traer del estado de cuenta todo lo que sea de `fechaDesde` en adelante.
  const incluirFecha = (fecha: Date) => fecha.getTime() >= fechaDesde.getTime();
  const movimientosSinOrdenar =
    cuenta.banco === "BROU"
      ? parseBrouConFiltro(workbookEstado, incluirFecha)
      : parseSantanderConFiltro(workbookEstado, incluirFecha);

  const movimientos = ordenarCronologicamente(movimientosSinOrdenar);

  // 4. Detectar días que este estado de cuenta no llega a cubrir. No es un error
  //    —si Administración no descargó esos días, no hay nada que hacer— pero
  //    conviene informarlo para que quede a la vista.
  const diasSinCobertura: Date[] = [];
  if (movimientos.length > 0) {
    const primeraDelEstado = soloDiaUTC(movimientos[0].fecha);
    const primeraEsperada = sumarDias(ultimaFechaEnPlanilla, 1);
    for (
      let dia = primeraEsperada;
      dia.getTime() < primeraDelEstado.getTime();
      dia = sumarDias(dia, 1)
    ) {
      diasSinCobertura.push(dia);
    }
  }

  // 5. Delegar la inserción al núcleo compartido.
  const resultado = await aplicarMovimientosAPlanilla(rutaPlanilla, cuenta, movimientos, rutaSalida);

  // 6. Desglosar lo agregado por día, para el reporte.
  const conteoPorDia = new Map<string, number>();
  for (const mov of resultado.agregados) {
    const clave = fechaATextoDDMMYYYY(mov.fecha);
    conteoPorDia.set(clave, (conteoPorDia.get(clave) ?? 0) + 1);
  }
  const agregadosPorDia = [...conteoPorDia.entries()].map(([fecha, cantidad]) => ({ fecha, cantidad }));

  return {
    ...resultado,
    ultimaFechaEnPlanilla,
    fechaDesde,
    diasSinCobertura,
    agregadosPorDia,
  };
}

// --- Uso directo por consola:
//   node actualizarDesdeUltimaFecha.js <planilla.xlsx> <estadoDeCuenta.xlsx> [salida.xlsx]
if (require.main === module) {
  const [rutaPlanilla, rutaEstado, rutaSalida] = process.argv.slice(2);
  if (!rutaPlanilla || !rutaEstado) {
    console.log(
      "Uso: node actualizarDesdeUltimaFecha.js <planilla.xlsx> <estadoDeCuenta.xlsx> [salida.xlsx]"
    );
    process.exit(1);
  }

  actualizarDesdeUltimaFecha(rutaPlanilla, rutaEstado, rutaSalida)
    .then((r) => {
      console.log(`Cuenta: ${r.cuenta.etiqueta} -> hoja "${r.hoja}"`);
      console.log(`Última fecha cargada en la planilla: ${fechaATextoDDMMYYYY(r.ultimaFechaEnPlanilla)} (fila ${r.filaAncla})`);
      console.log(`Se toman movimientos desde: ${fechaATextoDDMMYYYY(r.fechaDesde)} (inclusive)`);
      console.log(`Movimientos en ese rango dentro del estado de cuenta: ${r.totalMovimientosEnEstadoDeCuenta}`);
      console.log(`Agregados: ${r.agregados.length}` + (r.agregados.length > 0 ? ` (filas ${r.filaInicial} a ${r.filaFinal})` : ""));
      for (const { fecha, cantidad } of r.agregadosPorDia) {
        console.log(`   ${fecha}: ${cantidad}`);
      }
      console.log(`Omitidos por ya existir: ${r.omitidosPorDuplicado.length}`);

      if (r.diasSinCobertura.length > 0) {
        console.log("");
        console.log(
          `AVISO: este estado de cuenta no cubre ${r.diasSinCobertura.length} día(s) posteriores a lo cargado: ` +
            r.diasSinCobertura.map(fechaATextoDDMMYYYY).join(", ")
        );
        console.log("Si hubo movimientos esos días, hay que cargarlos con el estado de cuenta correspondiente.");
      }

      if (r.agregados.length === 0) {
        console.log("No se modificó el archivo (no había nada nuevo para agregar).");
      }
    })
    .catch((err) => {
      console.error("ERROR:", (err as Error).message);
      process.exit(1);
    });
}
