/**
 * Tipos compartidos por todo el flujo de planilla de bancos.
 */

export type Banco = "BROU" | "SANTANDER";

export type Moneda = "UYU" | "USD" | "EUR";

/** Las 5 cuentas que maneja la empresa hoy. Se puede ampliar sin tocar el resto del código. */
export type CuentaKey =
  | "BROU_PESOS"
  | "BROU_DOLARES"
  | "BROU_EUROS"
  | "SANTANDER_PESOS"
  | "SANTANDER_DOLARES";

export interface CuentaIdentificada {
  banco: Banco;
  moneda: Moneda;
  /** Número de cuenta tal cual figura en el archivo, útil para loguear/auditar. */
  numeroCuenta: string;
  /** A qué de las 5 cuentas de la empresa corresponde. */
  cuentaKey: CuentaKey;
  /** Etiqueta linda para mostrar en la planilla. */
  etiqueta: string;
}

export type TipoMovimiento = "debito" | "credito";

export interface MovimientoLimpio {
  /** Fecha del movimiento, sin hora. */
  fecha: Date;
  /** dd/mm/yyyy, para escribir directo en la planilla o loguear. */
  fechaTexto: string;
  banco: Banco;
  cuentaKey: CuentaKey;
  tipo: TipoMovimiento;
  monto: number;
  /** Texto de descripción "cruda" del banco (columna Descripción). */
  descripcion: string;
  /**
   * Texto donde generalmente aparece el cliente: "Asunto" en BROU,
   * o la propia Descripción en Santander (que no tiene columna Asunto separada).
   * Este es el campo que se le va a pasar a la IA para cotejar contra el listado de clientes.
   */
  textoParaMatchCliente: string;
  /** Número de documento / referencia del banco, si existe. */
  referencia?: string;
  /** Dependencia (BROU) o Tipo de Movimiento (Santander), por si sirve de contexto. */
  categoriaBanco?: string;
  /** Fila original en el excel (1-based), útil para debug. */
  filaOriginal: number;
}
