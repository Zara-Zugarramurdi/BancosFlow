/**
 * actualizarPlanilla.ts
 *
 * Toma la planilla maestra "bancos" (.xlsx) y un estado de cuenta ya parseado,
 * e inserta los movimientos nuevos justo después del último movimiento
 * cronológico real de la hoja que corresponda — no al final físico del archivo.
 *
 * ------------------------------------------------------------------------
 * HISTORIA DE ESTA DECISIÓN (por qué no es tan simple como "agregar al final")
 * ------------------------------------------------------------------------
 *
 * La primera versión de este script agregaba las filas nuevas después de la
 * ÚLTIMA FILA CON CUALQUIER DATO de la hoja, para no arriesgarse a insertar en
 * el medio de una planilla con miles de fórmulas encadenadas. Funcionaba, pero
 * probando contra la planilla real (hoja "BROU $") encontramos el problema:
 * después del último movimiento real (fila 7043, 17/07/2026) hay varios bloques
 * que NO son movimientos del día a día — "PENDIENTES DE DEBITO" (proyección a
 * futuro), un bloque viejo con fechas de 2025 que quedó pegado de una versión
 * anterior, una lista fija de "RETENCION JUDICIAL...", y un cronograma de un
 * préstamo ("PROYECTO BOTIJAS") que reutiliza la columna E para otra cosa. La
 * "última fila con datos" terminaba 55 filas más abajo del final real del
 * ledger, así que las filas nuevas quedaban invisibles para quien mira la
 * hoja esperando encontrarlas justo debajo del último movimiento.
 *
 * La solución obvia — "usar la fila con la fecha más reciente" — tampoco es
 * segura: encontramos fechas mal tipeadas a mano en el medio del ledger real
 * (ej. "30/12/2026" en la hoja SANTANDER $ en medio de datos de mediados de
 * 2026, o "29/05/2028" en SANTANDER U$S en medio de datos de 2024) que
 * hubieran hecho que el script insertara en pleno medio de una cadena de
 * fórmulas activa — mucho más peligroso que insertar de más al final.
 *
 * La heurística que sí funciona, verificada en las 5 hojas: el ledger real es
 * un bloque CONTIGUO de filas sin huecos, desde la fila 2 hasta la última fila
 * con datos antes del primer salto de 3+ filas vacías seguidas. No importa qué
 * fecha tenga cada fila individual, sólo que no haya un hueco. Además — y esto
 * terminó de confirmar que la heurística es correcta — en las 5 hojas ese punto
 * coincide EXACTO con el final del último rango de "fórmula compartida" de
 * Excel en la columna SALDO (ver `encontrarFilaAncla`), que es como Excel
 * internamente marca "hasta acá se rellenó esta fórmula para abajo". Cuando
 * ambos métodos no coinciden, el script prefiere frenar con un error antes que
 * adivinar.
 *
 * ------------------------------------------------------------------------
 * POR QUÉ HAY QUE REESCRIBIR FÓRMULAS A MANO AL INSERTAR
 * ------------------------------------------------------------------------
 *
 * Confirmado con un prototipo (no es una suposición): ni ExcelJS ni la
 * mayoría de las librerías de este estilo reescriben las referencias de una
 * fórmula cuando insertás filas en el medio de una hoja. Si la fila que antes
 * era la 7050 (fórmula `=H7049+F7050-G7050`) pasa a ser la 7052 después de
 * insertar 2 filas arriba, la fórmula se queda tal cual dice "H7049+F7050-G7050"
 * — apuntando a la fila equivocada. Además, la planilla real usa "fórmulas
 * compartidas" de Excel (un solo `H7005:H7043` en vez de una fórmula por fila),
 * que hay que tener en cuenta al mapear qué se movió y qué no.
 *
 * Por eso, después de insertar las filas nuevas, el script recorre TODA la
 * hoja modificada (y TODO el resto del libro, por si alguna otra hoja tiene
 * una fórmula que cruza hacia ésta) y reescribe cualquier referencia de celda
 * cuya fila haya quedado desplazada. Se probó explícitamente forzando un
 * recálculo real en LibreOffice (no sólo mirando el texto de la fórmula) para
 * confirmar que los saldos siguen dando los números correctos después de
 * insertar.
 */

import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { identificarCuenta } from "./accountIdentifier";
import { parseBrou } from "./parseBrou";
import { parseSantander } from "./parseSantander";
import type { CuentaIdentificada, CuentaKey, MovimientoLimpio } from "./types";

/** A qué hoja de la planilla maestra corresponde cada cuenta. */
const MAPA_HOJAS: Record<CuentaKey, string> = {
  BROU_PESOS: "BROU $",
  BROU_DOLARES: "BROU U$S",
  BROU_EUROS: "BROU EUROS",
  SANTANDER_PESOS: "SANTANDER $",
  SANTANDER_DOLARES: "SANTANDER U$S",
};

// Columnas fijas dentro de cada hoja (1-based, como las usa ExcelJS).
const COL = {
  FECHA: 1,
  RECIBO: 2,
  NUM_CHEQUE: 3,
  COMPROBANTE: 4,
  CONCEPTO: 5,
  DEBE: 6,
  HABER: 7,
  SALDO: 8,
} as const;

/**
 * Amarillo que ya usa la planilla para resaltar filas (confirmado mirando el
 * color real de celdas resaltadas a mano en varias hojas: `FFFFFF99`). El
 * patrón que se ve en la planilla real es resaltar desde FECHA hasta la
 * columna del monto que corresponda (DEBE o HABER, según el movimiento) —
 * SALDO se deja sin colorear, es el patrón más consistente entre las hojas.
 */
const AMARILLO_RESALTADO: ExcelJS.Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFFFFF99" },
};

export interface ResultadoActualizacion {
  cuenta: CuentaIdentificada;
  hoja: string;
  totalMovimientosEnEstadoDeCuenta: number;
  agregados: MovimientoLimpio[];
  omitidosPorDuplicado: MovimientoLimpio[];
  /** Fila donde estaba el último movimiento real ANTES de insertar (el "ancla"). */
  filaAncla: number;
  filaInicial: number;
  filaFinal: number;
}

function normalizarTexto(v: unknown): string {
  return String(v ?? "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

/** Redondea a centésimos para comparar montos sin problemas de precisión de floats. */
function redondear(n: number): number {
  return Math.round(n * 100) / 100;
}

function fechaSoloDia(d: Date): string {
  // Clave de fecha en UTC (día calendario), ignorando hora.
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(
    d.getUTCDate()
  ).padStart(2, "0")}`;
}

/** Convierte lo que venga en la celda FECHA de la planilla (Date, string o número) a "yyyy-mm-dd". */
function celdaFechaAClave(valor: ExcelJS.CellValue): string | null {
  if (!valor) return null;
  if (valor instanceof Date) return fechaSoloDia(valor);
  if (typeof valor === "number") {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(epoch.getTime() + valor * 24 * 60 * 60 * 1000);
    return fechaSoloDia(d);
  }
  if (typeof valor === "string") {
    // Algunas filas de la planilla real tienen la fecha cargada como TEXTO
    // ("27/07/2026") en vez de como fecha de Excel — pasa cuando se pega desde
    // otro lado. Sin esto, esas filas quedaban invisibles tanto para la detección
    // del ancla como para el chequeo de duplicados.
    const m = valor.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) {
      const [, dd, mm, yyyy] = m;
      const d = new Date(Date.UTC(Number(yyyy), Number(mm) - 1, Number(dd)));
      // Validar que la fecha exista de verdad (evita aceptar cosas como 31/02/2026).
      if (
        d.getUTCFullYear() === Number(yyyy) &&
        d.getUTCMonth() === Number(mm) - 1 &&
        d.getUTCDate() === Number(dd)
      ) {
        return fechaSoloDia(d);
      }
    }
  }
  return null;
}

/**
 * Clave de un movimiento para detectar duplicados: fecha + tipo + monto (sin texto).
 *
 * OJO: deliberadamente NO incluye el texto de concepto/descripción, aunque el
 * pedido original era "misma fecha, misma descripción, mismo textoParaMatchCliente".
 * Probando contra la planilla real esa clave con texto exacto no funciona: lo que
 * un humano tipeó en CONCEPTO ("ROLA") casi nunca coincide texto-a-texto con lo
 * que el banco manda en su descripción cruda ("537806PP EMITIDO Rola Ltda").
 *
 * fecha+tipo+monto sí es un identificador fuerte acá, pero contando repeticiones
 * (ver `construirIndiceExistentes`): es normal que el mismo día haya más de un
 * movimiento con igual monto (ej. 4 "COMISIONES BANCARIAS" de $77.71 el mismo
 * día, una por transferencia), así que cada coincidencia nueva "consume" una
 * fila existente antes de considerarse duplicado.
 */
function claveMovimiento(fechaClave: string, tipo: "debito" | "credito", monto: number): string {
  return `${fechaClave}|${tipo}|${redondear(monto).toFixed(2)}`;
}

/**
 * ¿Esta fila tiene una FECHA válida? Es el criterio que usamos para decidir si una
 * fila es un movimiento real del ledger o no.
 *
 * Sólo miramos FECHA (y no CONCEPTO/DEBE/HABER/SALDO, como hacíamos antes) porque
 * es lo único que distingue de forma confiable un movimiento real del relleno que
 * tienen estas hojas. Probando contra la planilla real encontramos que mirar las
 * otras columnas rompe la detección de dos formas distintas en la misma hoja
 * (`SANTANDER $`):
 *   - El bloque de proyección "PENDIENTES DE DEBITO" tiene CONCEPTO y SALDO pero
 *     no es historia real, y quedaba contado como parte del ledger.
 *   - Después de ese bloque había 5 filas donde alguien arrastró la fórmula de
 *     SALDO de más, sin ningún movimiento: sin fecha, sin concepto, sin importes.
 *     Como "tenían algo" en SALDO, también contaban como ledger.
 * Resultado: el ancla daba la fila 3466 en vez de la 3454 (el último movimiento
 * real), y los movimientos nuevos terminaban insertados DESPUÉS del bloque de
 * pendientes, que se supone va siempre al final.
 *
 * Todo movimiento real tiene fecha; ninguna fila de relleno, proyección o
 * fórmula-fantasma la tiene. Por eso este criterio es más simple y más robusto.
 */
function filaTieneFecha(hoja: ExcelJS.Worksheet, fila: number): boolean {
  return celdaFechaAClave(hoja.getCell(fila, COL.FECHA).value) !== null;
}

/**
 * Encuentra la última fila del bloque CONTIGUO de movimientos reales, arrancando
 * en la fila 2 (debajo del encabezado) y parando en cuanto aparecen 3 o más filas
 * seguidas sin fecha.
 */
function encontrarAnclaPorContiguidad(hoja: ExcelJS.Worksheet): number {
  const ultimaFilaHoja = hoja.actualRowCount || hoja.rowCount;
  let sinFechaSeguidas = 0;
  for (let fila = 2; fila <= ultimaFilaHoja + 3; fila++) {
    const vacia = fila > ultimaFilaHoja || !filaTieneFecha(hoja, fila);
    if (vacia) {
      sinFechaSeguidas++;
      if (sinFechaSeguidas >= 3) {
        return fila - sinFechaSeguidas;
      }
    } else {
      sinFechaSeguidas = 0;
    }
  }
  // No encontramos ningún hueco de 3+ filas: la hoja es un bloque contiguo hasta el final.
  return ultimaFilaHoja;
}

/**
 * Devuelve todos los rangos de "fórmula compartida" de Excel en la columna SALDO
 * (cada uno es, literalmente, hasta dónde Excel rellenó una fórmula hacia abajo
 * la última vez que alguien la arrastró).
 */
function obtenerRangosFormulaCompartida(hoja: ExcelJS.Worksheet): Array<{ inicio: number; fin: number }> {
  const rangos: Array<{ inicio: number; fin: number }> = [];
  hoja.eachRow((fila) => {
    const valor = fila.getCell(COL.SALDO).value;
    if (valor && typeof valor === "object" && "ref" in valor && typeof valor.ref === "string") {
      const match = valor.ref.match(/^[A-Z]+(\d+):[A-Z]+(\d+)$/);
      if (match) {
        rangos.push({ inicio: parseInt(match[1], 10), fin: parseInt(match[2], 10) });
      }
    }
  });
  return rangos;
}

/**
 * Determina la fila "ancla": la última fila que realmente pertenece al ledger de
 * movimientos. Usamos la continuidad de datos como método principal, y los rangos
 * de fórmula compartida como red de seguridad: si el punto de corte cayera EN
 * MEDIO de un rango de fórmula compartida (no al final de uno), es señal de que
 * algo no es lo que esperamos y preferimos frenar con un error antes que insertar
 * ahí. Que el ancla coincida con el FINAL de un rango (el caso normal) o que no
 * haya ningún rango compartido cerca (hojas más chicas, como BROU EUROS) está bien.
 */
function encontrarFilaAncla(hoja: ExcelJS.Worksheet): number {
  const ancla = encontrarAnclaPorContiguidad(hoja);
  const rangos = obtenerRangosFormulaCompartida(hoja);
  const rangoQueCruza = rangos.find((r) => r.inicio <= ancla && ancla < r.fin);

  if (rangoQueCruza) {
    throw new Error(
      `La fila donde terminaría el ledger real de la hoja "${hoja.name}" (${ancla}, calculada por continuidad ` +
        `de datos) cae en medio de un rango de fórmula compartida de la columna SALDO (${rangoQueCruza.inicio}:` +
        `${rangoQueCruza.fin}). Es inesperado y podría indicar que alguien reorganizó la hoja a mano; revisar ` +
        `manualmente antes de correr el script de nuevo.`
    );
  }

  return ancla;
}

/**
 * Construye un multiset (clave -> cantidad de veces que aparece) de los movimientos
 * ya cargados en la hoja, recorriendo TODAS las filas con datos (no sólo las del
 * ledger contiguo: mejor pecar de cauto y mirar toda la hoja, incluidos los bloques
 * de "planificación", por si alguna vez se cargó algo ahí también).
 */
function construirIndiceExistentes(hoja: ExcelJS.Worksheet): Map<string, number> {
  const contador = new Map<string, number>();
  hoja.eachRow((fila, numeroFila) => {
    if (numeroFila === 1) return; // encabezado
    const fechaClave = celdaFechaAClave(fila.getCell(COL.FECHA).value);
    if (!fechaClave) return;
    const debe = fila.getCell(COL.DEBE).value;
    const haber = fila.getCell(COL.HABER).value;
    const debeNum = typeof debe === "number" ? debe : 0;
    const haberNum = typeof haber === "number" ? haber : 0;
    if (debeNum !== 0) {
      const clave = claveMovimiento(fechaClave, "credito", debeNum);
      contador.set(clave, (contador.get(clave) ?? 0) + 1);
    }
    if (haberNum !== 0) {
      const clave = claveMovimiento(fechaClave, "debito", haberNum);
      contador.set(clave, (contador.get(clave) ?? 0) + 1);
    }
  });
  return contador;
}

function validarEncabezado(hoja: ExcelJS.Worksheet): void {
  const filaEncabezado = hoja.getRow(1);
  const columnasCriticas: Array<[number, string]> = [
    [COL.FECHA, "FECHA"],
    [COL.CONCEPTO, "CONCEPTO"],
    [COL.DEBE, "DEBE"],
    [COL.HABER, "HABER"],
    [COL.SALDO, "SALDO"],
  ];
  for (const [col, esperado] of columnasCriticas) {
    const actual = normalizarTexto(filaEncabezado.getCell(col).value);
    if (actual !== esperado) {
      throw new Error(
        `La columna ${col} de la hoja "${hoja.name}" dice "${actual}" y se esperaba "${esperado}". ` +
          `Puede que el layout de la planilla haya cambiado; revisar antes de seguir para no escribir en columnas equivocadas.`
      );
    }
  }
}

/**
 * Reescribe la fórmula de una celda "clon" de un rango de fórmula compartida,
 * a partir de la fórmula del "master" del rango. Las referencias relativas (sin
 * `$`) se desplazan según la distancia entre la fila del clon y la del master;
 * las absolutas (`$fila`) quedan fijas — es el mismo comportamiento que Excel
 * usa al arrastrar una fórmula hacia abajo.
 */
function formulaParaFilaClon(formulaMaster: string, filaMaster: number, filaClon: number): string {
  const delta = filaClon - filaMaster;
  const REF_RE = /((?:'([^']+)'!|[A-Za-z_][A-Za-z0-9_.]*!))?(\$?)([A-Z]{1,3})(\$?)(\d+)\b/g;
  return formulaMaster.replace(REF_RE, (match, prefijo, _q, d1, col, d2, filaTexto) => {
    if (prefijo) return match; // fórmulas compartidas no cruzan hojas; si hubiera prefijo, no lo tocamos igual
    if (d2 === "$") return match; // fila absoluta, no se desliza
    const fila = parseInt(filaTexto, 10);
    return `${d1}${col}${d2}${fila + delta}`;
  });
}

/**
 * "Des-comparte" (convierte a fórmulas literales, una por celda) cualquier rango
 * de fórmula compartida de la hoja cuyo final caiga en `filaLimite` o después.
 *
 * Hace falta hacer esto ANTES de mover filas: hay un problema conocido de ExcelJS
 * (confirmado con un prototipo — no es una suposición) donde `spliceRows` puede
 * dejar la relación "master/clon" de una fórmula compartida en un estado inválido
 * y tira `Shared Formula master must exist above and or left of clone` al guardar.
 * Convirtiendo todo a fórmulas literales antes de mover filas, evitamos el bug
 * por completo: cada celda queda con su propio texto de fórmula, sin depender de
 * ninguna otra celda "master".
 */
function desCompartirFormulasDesde(hoja: ExcelJS.Worksheet, filaLimite: number): void {
  const rangos: Array<{ inicio: number; fin: number; col: number; formula: string; filaMaster: number }> = [];

  hoja.eachRow((fila) => {
    fila.eachCell((celda) => {
      const v = celda.value as
        | { shareType?: string; formula?: string; ref?: string }
        | null
        | undefined;
      if (v && typeof v === "object" && v.shareType === "shared" && typeof v.formula === "string" && typeof v.ref === "string") {
        const m = v.ref.match(/^[A-Z]+(\d+):[A-Z]+(\d+)$/);
        if (m) {
          rangos.push({
            inicio: parseInt(m[1], 10),
            fin: parseInt(m[2], 10),
            col: Number(celda.fullAddress.col),
            formula: v.formula,
            filaMaster: Number(celda.row),
          });
        }
      }
    });
  });

  for (const rango of rangos) {
    if (rango.fin < filaLimite) continue; // todo el rango queda antes del punto de corte, no hace falta tocarlo
    for (let r = rango.inicio; r <= rango.fin; r++) {
      const celda = hoja.getCell(r, rango.col);
      const formulaLiteral =
        r === rango.filaMaster ? rango.formula : formulaParaFilaClon(rango.formula, rango.filaMaster, r);
      celda.value = { formula: formulaLiteral } as ExcelJS.CellFormulaValue;
    }
  }
}

/**
 * Reescribe las referencias de celda de UNA fórmula que apunten a `hojaObjetivo`
 * con fila >= `filaLimite` (numeración ANTES de insertar), sumándoles `delta`.
 * Referencias a otras hojas, o a filas antes del límite, quedan intactas.
 */
function ajustarReferenciasFormula(
  formula: string,
  hojaActual: string,
  hojaObjetivo: string,
  filaLimite: number,
  delta: number
): string {
  // Grupo 1: prefijo de hoja completo (con el '!'), si lo hay.
  // Grupo 2/3: nombre de hoja entre comillas simples / sin comillas.
  // Grupo 4/6: '$' opcional antes de columna/fila. Grupo 5: letras de columna. Grupo 7: número de fila.
  const REF_RE = /((?:'([^']+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?(\$?)([A-Z]{1,3})(\$?)(\d+)\b/g;

  return formula.replace(
    REF_RE,
    (match, prefijo, hojaEntreComillas, hojaSinComillas, d1, col, d2, filaTexto) => {
      const hojaDeLaReferencia = hojaEntreComillas ?? hojaSinComillas ?? hojaActual;
      if (hojaDeLaReferencia !== hojaObjetivo) return match; // no es una referencia a la hoja que estamos ajustando
      const fila = parseInt(filaTexto, 10);
      if (fila < filaLimite) return match; // esta fila no se movió
      return `${prefijo ?? ""}${d1}${col}${d2}${fila + delta}`;
    }
  );
}

/**
 * Recorre TODO el libro (todas las hojas) y reescribe cualquier fórmula que
 * tenga una referencia a `hojaObjetivo` con fila >= `filaLimite`, sumándole
 * `delta`. Hace falta recorrer todo el libro y no sólo la hoja modificada por
 * las (pocas) fórmulas que cruzan de una cuenta a otra, ej. los "totales"
 * generales de pesos/dólares que sí encontramos en la planilla real.
 */
function ajustarFormulasEnTodoElLibro(
  workbook: ExcelJS.Workbook,
  hojaObjetivo: string,
  filaLimite: number,
  delta: number
): void {
  for (const hoja of workbook.worksheets) {
    hoja.eachRow({ includeEmpty: false }, (fila) => {
      fila.eachCell({ includeEmpty: false }, (celda) => {
        const valor = celda.value;
        if (valor && typeof valor === "object" && "formula" in valor && typeof valor.formula === "string") {
          const nuevaFormula = ajustarReferenciasFormula(valor.formula, hoja.name, hojaObjetivo, filaLimite, delta);
          if (nuevaFormula !== valor.formula) {
            celda.value = { formula: nuevaFormula } as ExcelJS.CellFormulaValue;
          }
        }
      });
    });
  }
}

/**
 * Agrega a la planilla los movimientos de un estado de cuenta que todavía no estén
 * cargados, insertándolos justo después del último movimiento real de la hoja
 * (no al final físico del archivo) y reajustando las fórmulas que haga falta.
 * Guarda el resultado en `rutaSalida` (por defecto, sobreescribe `rutaPlanilla`).
 */
export async function actualizarPlanilla(
  rutaPlanilla: string,
  rutaEstadoDeCuenta: string,
  fechaObjetivo: Date,
  rutaSalida: string = rutaPlanilla
): Promise<ResultadoActualizacion> {
  // 1. Parsear el estado de cuenta con la lógica que ya teníamos.
  const workbookEstado = XLSX.readFile(rutaEstadoDeCuenta);
  const cuenta = identificarCuenta(workbookEstado);
  const movimientos =
    cuenta.banco === "BROU"
      ? parseBrou(workbookEstado, fechaObjetivo)
      : parseSantander(workbookEstado, fechaObjetivo);

  // 2. Abrir la planilla maestra y ubicar la hoja que corresponde.
  const nombreHoja = MAPA_HOJAS[cuenta.cuentaKey];
  const workbookPlanilla = new ExcelJS.Workbook();
  await workbookPlanilla.xlsx.readFile(rutaPlanilla);
  const hoja = workbookPlanilla.getWorksheet(nombreHoja);
  if (!hoja) {
    throw new Error(
      `No se encontró la hoja "${nombreHoja}" en la planilla. Hojas disponibles: ${workbookPlanilla.worksheets
        .map((h) => h.name)
        .join(", ")}`
    );
  }
  validarEncabezado(hoja);

  // 3. Armar el índice de lo que ya está cargado, para no duplicar.
  const existentes = construirIndiceExistentes(hoja);

  const porAgregar: MovimientoLimpio[] = [];
  const omitidosPorDuplicado: MovimientoLimpio[] = [];

  for (const mov of movimientos) {
    const fechaClave = fechaSoloDia(mov.fecha);
    const clave = claveMovimiento(fechaClave, mov.tipo, mov.monto);
    const disponibles = existentes.get(clave) ?? 0;
    if (disponibles > 0) {
      existentes.set(clave, disponibles - 1);
      omitidosPorDuplicado.push(mov);
      continue;
    }
    existentes.set(clave, 0);
    porAgregar.push(mov);
  }

  // 4. Encontrar dónde termina el ledger real y ahí insertar (no al final físico).
  const filaAncla = encontrarFilaAncla(hoja);
  const filaInsercion = filaAncla + 1;
  const cantidad = porAgregar.length;

  if (cantidad === 0) {
    return {
      cuenta,
      hoja: nombreHoja,
      totalMovimientosEnEstadoDeCuenta: movimientos.length,
      agregados: [],
      omitidosPorDuplicado,
      filaAncla,
      filaInicial: filaInsercion,
      filaFinal: filaAncla,
    };
  }

  // 5. Antes de mover nada: convertir a fórmulas literales cualquier rango de
  //    fórmula compartida que vaya a quedar afectado por la inserción (ver
  //    comentario de `desCompartirFormulasDesde` — evita un bug de ExcelJS).
  desCompartirFormulasDesde(hoja, filaInsercion);

  // 6. Insertar `cantidad` filas en blanco justo después del ancla.
  hoja.spliceRows(filaInsercion, 0, ...Array.from({ length: cantidad }, () => [] as unknown[]));

  // 7. Reajustar TODAS las fórmulas del libro que quedaron corridas por el insert.
  //    OJO: esto tiene que pasar ANTES de escribir las filas nuevas (paso 8) — si
  //    fuera al revés, este mismo paso terminaría "reajustando" las fórmulas que
  //    recién escribimos (que ya están bien, recién creadas) como si también se
  //    hubieran corrido, rompiéndolas. En este punto las filas nuevas todavía
  //    están vacías, así que el barrido sólo toca fórmulas que realmente vivían
  //    más abajo y se corrieron de verdad.
  ajustarFormulasEnTodoElLibro(workbookPlanilla, nombreHoja, filaInsercion, cantidad);

  // 8. Completar las filas nuevas con los movimientos (con los números de fila
  //    finales, ya no hace falta ajustarlas después).
  const saldoAnclaTieneValor = hoja.getCell(filaAncla, COL.SALDO).value != null;
  let filaActual = filaAncla;
  for (const mov of porAgregar) {
    filaActual += 1;
    const fila = hoja.getRow(filaActual);
    const concepto = mov.textoParaMatchCliente;

    // Copiar el estilo (borde, fuente, alineación) de la fila ancla — así las filas
    // nuevas se ven igual que el resto del ledger, no en blanco/sin borde. Se copia
    // ANTES de asignar los numFmt/fill de abajo, que sí queremos que queden como los
    // definimos explícitamente. OJO: `cell.style = otraCelda.style` en ExcelJS copia
    // una REFERENCIA compartida al mismo objeto de estilo interno, no una copia
    // independiente — si no lo clonamos acá, pintar la fila nueva de amarillo más
    // abajo termina pintando también la fila ancla (y cualquier otra celda que
    // comparta ese mismo estilo), porque en el fondo es el mismo objeto. Lo
    // confirmamos con un prototipo antes de aplicar el fix.
    for (let col = COL.FECHA; col <= COL.SALDO; col++) {
      const celdaOrigen = hoja.getCell(filaAncla, col);
      fila.getCell(col).style = JSON.parse(JSON.stringify(celdaOrigen.style));
    }

    fila.getCell(COL.FECHA).value = mov.fecha;
    fila.getCell(COL.FECHA).numFmt = "mm-dd-yy";
    fila.getCell(COL.CONCEPTO).value = concepto;

    if (mov.tipo === "credito") {
      fila.getCell(COL.DEBE).value = mov.monto; // DEBE = plata que entra (convención de esta planilla)
      fila.getCell(COL.DEBE).numFmt = "#,##0.00";
    } else {
      fila.getCell(COL.HABER).value = mov.monto; // HABER = plata que sale
      fila.getCell(COL.HABER).numFmt = "#,##0.00";
    }

    if (saldoAnclaTieneValor) {
      fila.getCell(COL.SALDO).value = {
        formula: `H${filaActual - 1}+F${filaActual}-G${filaActual}`,
      } as ExcelJS.CellFormulaValue;
      fila.getCell(COL.SALDO).numFmt = "#,##0.00";
    }

    // Resaltar en amarillo (el mismo tono que ya usa la planilla a mano) desde
    // FECHA hasta la columna del monto que corresponda — SALDO queda sin
    // colorear, que es el patrón más consistente que encontramos en la planilla real.
    const ultimaColumnaAResaltar = mov.tipo === "credito" ? COL.DEBE : COL.HABER;
    for (let col = COL.FECHA; col <= ultimaColumnaAResaltar; col++) {
      fila.getCell(col).fill = AMARILLO_RESALTADO;
    }
  }

  // 9. Forzar que Excel/LibreOffice recalculen todo al abrir el archivo, en vez de
  //    mostrar los valores cacheados (que ya no son válidos para las filas corridas).
  workbookPlanilla.calcProperties.fullCalcOnLoad = true;

  await workbookPlanilla.xlsx.writeFile(rutaSalida);

  return {
    cuenta,
    hoja: nombreHoja,
    totalMovimientosEnEstadoDeCuenta: movimientos.length,
    agregados: porAgregar,
    omitidosPorDuplicado,
    filaAncla,
    filaInicial: filaAncla + 1,
    filaFinal: filaActual,
  };
}

// --- Uso directo por consola:
//   node actualizarPlanilla.js <planilla.xlsx> <estadoDeCuenta.xlsx> <dd/mm/yyyy> [salida.xlsx]
if (require.main === module) {
  const [rutaPlanilla, rutaEstado, fechaTexto, rutaSalida] = process.argv.slice(2);
  if (!rutaPlanilla || !rutaEstado || !fechaTexto) {
    console.log(
      "Uso: node actualizarPlanilla.js <planilla.xlsx> <estadoDeCuenta.xlsx> <dd/mm/yyyy> [salida.xlsx]"
    );
    process.exit(1);
  }
  const [dd, mm, yyyy] = fechaTexto.split("/").map(Number);
  const fechaObjetivo = new Date(Date.UTC(yyyy, mm - 1, dd));

  actualizarPlanilla(rutaPlanilla, rutaEstado, fechaObjetivo, rutaSalida)
    .then((resultado) => {
      console.log(`Cuenta: ${resultado.cuenta.etiqueta} -> hoja "${resultado.hoja}"`);
      console.log(`Último movimiento real antes de correr: fila ${resultado.filaAncla}`);
      console.log(`Movimientos en el estado de cuenta para ${fechaTexto}: ${resultado.totalMovimientosEnEstadoDeCuenta}`);
      console.log(`Agregados: ${resultado.agregados.length} (filas ${resultado.filaInicial} a ${resultado.filaFinal})`);
      console.log(`Omitidos por ya existir: ${resultado.omitidosPorDuplicado.length}`);
      if (resultado.omitidosPorDuplicado.length > 0) {
        console.log("Detalle de omitidos:");
        for (const m of resultado.omitidosPorDuplicado) {
          console.log(`  - ${m.fechaTexto} | ${m.tipo} | ${m.monto} | ${m.textoParaMatchCliente}`);
        }
      }
      if (resultado.agregados.length === 0) {
        console.log("No se modificó el archivo (no había nada nuevo para agregar).");
      }
    })
    .catch((err) => {
      console.error("ERROR:", (err as Error).message);
      process.exit(1);
    });
}
