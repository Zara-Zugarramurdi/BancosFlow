/**
 * textoConcepto.ts
 *
 * Criterios compartidos por los dos formatos (BROU y Santander) para decidir si el
 * texto principal de un movimiento alcanza por sí solo, o si hay que completarlo con
 * el campo de respaldo del banco.
 *
 * Vive en su propio archivo porque la regla es la misma para ambos bancos, aunque los
 * campos concretos cambien:
 *
 *   - BROU:      principal = Asunto        · respaldo = Descripción
 *   - Santander: principal = Descripción   · respaldo = Tipo de Movimiento, luego Referencia
 */

/**
 * ¿El texto necesita que lo completemos con el campo de respaldo?
 *
 * Tres criterios, cualquiera alcanza:
 *
 *   1. Está vacío. No hay nada que preservar.
 *   2. Tiene menos de 5 caracteres. Ningún texto tan corto identifica por sí solo a un
 *      cliente de forma confiable.
 *   3. No contiene ninguna letra. Un texto de puros números o signos (`"3507"`,
 *      `"24609930"`, `"-"`, `"--"`) es una referencia interna del banco, no un nombre.
 *
 * El criterio 3 se expresa como "no tiene letras" en lugar de "son todos dígitos" para
 * que también cubra números con separadores (`"3.507"`) y los guiones sueltos, sin
 * necesidad de enumerar cada variante.
 */
export function necesitaComplemento(texto: string): boolean {
  const t = texto.trim();
  if (t === "") return true;
  if (t.length < 5) return true;
  if (!/\p{L}/u.test(t)) return true;
  return false;
}

/**
 * Arma el texto que va a la columna CONCEPTO de la planilla (y que después se le pasa a
 * la IA para cotejar contra el listado de clientes).
 *
 * Si el texto principal se vale por sí solo, se usa tal cual. Si no, se le CONCATENA el
 * primer respaldo con contenido, en vez de reemplazarlo: `"3507"` + `"TRF SPI PAGO PROV."`
 * queda como `"3507 - TRF SPI PAGO PROV."`.
 *
 * Se concatena incluso cuando el principal es un `"-"` (queda `"- - DEPOSITO CHEQUES
 * CLEARING"`). Es una decisión explícita: se prefiere un poco de ruido antes que
 * arriesgarse a descartar información. Un texto corto puede ser justamente el nombre del
 * cliente — en los estados de cuenta reales aparecen `"ITAU"` y `"BBVA"` como asunto, y en
 * la planilla hay cientos de conceptos cargados a mano con siglas de 3 y 4 letras (DUA,
 * BSE, UTE, SMI, OSE, BPS, DGI...).
 *
 * Única excepción: si el principal está vacío no hay nada que preservar, así que va sólo
 * el respaldo (concatenar dejaría un separador colgando).
 */
export function combinarConRespaldo(principal: string, ...respaldos: string[]): string {
  const p = principal.trim();
  if (!necesitaComplemento(p)) return p;

  const respaldo = respaldos.map((r) => r.trim()).find((r) => r !== "");
  if (respaldo === undefined) return principal; // no hay nada mejor: devolvemos el crudo

  return p === "" ? respaldo : `${p} - ${respaldo}`;
}
