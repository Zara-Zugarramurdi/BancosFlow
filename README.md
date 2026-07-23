# Flujo de Planilla de Bancos — Identificación y limpieza

Probado contra los 4 archivos que pasaste (2 BROU .xls, 2 Santander .xlsx). Los 4 se identifican y filtran correctamente.

## Instalación

```bash
npm install
```

Dependencia clave: [`xlsx` (SheetJS)](https://www.npmjs.com/package/xlsx). Es la que permite leer tanto los `.xls` viejos de BROU como los `.xlsx` de Santander con la misma librería.

## Archivos

- **`src/types.ts`** — tipos compartidos (`CuentaIdentificada`, `MovimientoLimpio`, etc).
- **`src/accountIdentifier.ts`** — dado un archivo (cualquiera de los 2 formatos), determina a cuál de las 5 cuentas corresponde:
  - Detecta el banco buscando texto fijo en el archivo (`"Saldos y Movimientos"` = BROU, `"Banco Santander"` = Santander).
  - Lee la moneda y el número de cuenta desde la cabecera de cada formato.
  - Mapea `(banco, moneda) -> cuenta` usando una tabla `MAPA_CUENTAS` fácil de extender si mañana aparece una 6ª cuenta.
- **`src/parseBrou.ts`** — limpia el formato BROU y devuelve solo los movimientos de una fecha dada.
- **`src/parseSantander.ts`** — ídem para el formato Santander.
- **`src/procesarEstadoDeCuenta.ts`** — orquestador: detecta el formato solo y llama al parser que corresponda. **Este es el que probablemente quieras llamar desde el proceso principal.**
- **`src/actualizarPlanilla.ts`** — toma la planilla maestra + un estado de cuenta ya identificado/parseado, e inserta los movimientos nuevos en el lugar correcto de la hoja. Ver sección dedicada más abajo.

## Uso

```bash
npx tsc   # compila a dist/
node dist/procesarEstadoDeCuenta.js "C:\ruta\al\archivo.xlsx" 20/07/2026
```

O importado desde tu propio código:

```ts
import { procesarEstadoDeCuenta } from "./procesarEstadoDeCuenta";

const hoy = new Date(); // o la fecha del día anterior, según cuándo corra el proceso
const { cuenta, movimientos } = procesarEstadoDeCuenta("C:\\ruta\\archivo.xlsx", hoy);

console.log(cuenta.etiqueta); // "BROU Dólares", "Santander Pesos", etc.
for (const m of movimientos) {
  console.log(m.tipo, m.monto, m.textoParaMatchCliente);
}
```

Cada `movimiento` trae el campo **`textoParaMatchCliente`** — es el texto que le pasarías a Hermes/DeepSeek junto con tu listado de clientes para que sugiera a quién corresponde. En BROU es la columna "Asunto" (o la descripción si el asunto viene vacío); en Santander, como no hay columna de asunto separada, es la propia descripción.

## Cómo identificamos cada cuenta

| Banco     | Moneda | Cómo se detecta                                                                 | Cuenta          |
|-----------|--------|-----------------------------------------------------------------------------------|-----------------|
| BROU      | `$`    | Celda `"Moneda\n$"` en la cabecera                                                | BROU Pesos      |
| BROU      | `U$S`  | Celda `"Moneda\nU$S"`                                                             | BROU Dólares    |
| BROU      | `€`    | Celda `"Moneda\n€"` (no tuvimos archivo de ejemplo, pero el patrón es el mismo)   | BROU Euros      |
| Santander | `UYU`  | Columna "Moneda" = `UYU` en la fila de datos de cabecera                          | Santander Pesos |
| Santander | `USD`  | Columna "Moneda" = `USD`                                                          | Santander Dólares |

No dependemos del número de cuenta como clave principal porque en teoría podría cambiar (ej. si el banco reemplaza el número de cuenta como pasó con "N° de Cuenta anterior" en BROU); la combinación banco+moneda es más estable. El número de cuenta sí se guarda en el resultado (`numeroCuenta`) para loguearlo o para validación cruzada extra si más adelante quieren blindar aún más la identificación.

## Sobre "limpiar para que quede solo un día"

Ambos parsers:
1. Ubican dinámicamente la fila de encabezado de la tabla de movimientos (no asumen un número de fila fijo, por si el banco agrega/saca alguna fila de saldos arriba).
2. Recorren fila por fila hasta que la tabla termina (fila sin fecha).
3. Se quedan solo con las filas cuya fecha coincide (día/mes/año) con la `fechaObjetivo` que le pasás.
4. Filas sin débito ni crédito (basura / separadores) se descartan.

Ojo con un detalle real que vimos en los archivos: **el estado de cuenta de BROU trae varios días mezclados** (ej. filas del 20/07 y del 16/07 en el mismo archivo), así que el filtro por fecha es necesario, no opcional. El de Santander en cambio ya viene acotado a un rango corto ("Período: 16/07/2026 - 20/07/2026") pero igual puede traer más de un día, así que aplica el mismo filtro.

## Actualizar la planilla de bancos automáticamente

**`src/actualizarPlanilla.ts`** — toma la planilla maestra (`.xlsx`, la que tiene las 10 hojas: BROU $, BROU U$S, BROU EUROS, SANTANDER $, SANTANDER U$S, etc.) y un estado de cuenta, e **inserta** los movimientos nuevos de la fecha pedida justo después del último movimiento real de la hoja que corresponda.

```bash
npx tsc
node dist/actualizarPlanilla.js "C:\ruta\bancos.xlsx" "C:\ruta\estado.xlsx" 20/07/2026
```

Por defecto sobreescribe la misma planilla. Si preferís no arriesgar el archivo mientras probás, pasale una 4ª ruta y escribe ahí en vez de sobreescribir:

```bash
node dist/actualizarPlanilla.js "C:\ruta\bancos.xlsx" "C:\ruta\estado.xlsx" 20/07/2026 "C:\ruta\bancos-actualizada.xlsx"
```

Internamente:
1. Parsea el estado de cuenta con la misma lógica de `procesarEstadoDeCuenta` (identifica cuenta + filtra por fecha).
2. Abre la planilla con [ExcelJS](https://www.npmjs.com/package/exceljs) (a diferencia de `xlsx`/SheetJS, que uso para *leer* los estados de cuenta, ExcelJS sí sabe escribir de vuelta preservando fórmulas y formato de celdas).
3. Ubica la hoja según la cuenta detectada (`BROU_PESOS` → hoja `"BROU $"`, etc.) y valida que las columnas FECHA/CONCEPTO/DEBE/HABER/SALDO sigan donde esperamos.
4. Arma un índice de lo que ya está cargado y descarta del estado de cuenta lo que ya existe (ver "Duplicados" abajo).
5. Encuentra la fila donde termina el último movimiento real (el "ancla") — ver más abajo.
6. Inserta ahí las filas nuevas, corriendo hacia abajo todo lo que hubiera después (proyecciones, listas de referencia, lo que sea) y reajustando las fórmulas que haga falta.

### Mapeo de campos: banco → planilla

| Movimiento del banco | Columna en la planilla | Significado |
|---|---|---|
| Crédito (entra plata) | `DEBE` | En la jerga de esta planilla, "DEBE" es lo que les entra (típicamente el pago de un cliente) |
| Débito (sale plata) | `HABER` | Lo que sale (pago a proveedores, comisiones, etc.) |
| — | `CONCEPTO` | Se completa con `textoParaMatchCliente` — el mismo texto "limpio" que ya usábamos para el matching de IA. |
| — | `SALDO` | Fórmula `=H{fila anterior}+F{fila}-G{fila}`, encadenada normalmente, sólo si la fila ancla ya tenía algo en SALDO. |

### Dónde se insertan las filas nuevas (y por qué esto cambió durante el desarrollo)

La primera versión de este script agregaba las filas nuevas después de la **última fila con cualquier dato** de la hoja, para no arriesgarse a insertar en el medio de miles de fórmulas encadenadas. Probando contra la hoja real `"BROU $"` encontramos el problema: después del último movimiento real (fila 7043, 17/07/2026) hay varios bloques que **no son movimientos del día a día** — una proyección `"PENDIENTES DE DEBITO"`, un bloque viejo con fechas de 2025 que quedó pegado de una versión anterior, una lista fija de `"RETENCION JUDICIAL..."`, y el cronograma de un préstamo (`"PROYECTO BOTIJAS"`) que reutiliza la columna E para otra cosa. La "última fila con datos" terminaba **55 filas más abajo** del final real del ledger, así que los movimientos nuevos quedaban invisibles para quien mira la hoja esperando encontrarlos justo debajo del último movimiento — el mismo problema que reportaron.

La solución obvia — "usar la fila con la fecha más reciente" — tampoco es segura: hay fechas mal tipeadas a mano en medio del ledger real (`30/12/2026` en `SANTANDER $` en medio de datos de mediados de 2026; `29/05/2028` en `SANTANDER U$S` en medio de datos de 2024) que hubieran hecho que el script insertara en pleno medio de una cadena de fórmulas activa. Mucho más peligroso que el problema original.

La heurística que sí funciona: el ledger real es un **bloque contiguo de filas sin huecos**, desde la fila 2 hasta la última fila con datos antes del primer salto de 3+ filas vacías seguidas. No importa la fecha de cada fila individual, sólo que no haya un hueco. Como confirmación extra, se cruza con una segunda señal independiente: en las 5 hojas, ese punto coincide exacto con el final del último rango de "fórmula compartida" de Excel en la columna SALDO (la forma en que Excel internamente recuerda "hasta acá se arrastró esta fórmula la última vez"). Si el punto de corte cayera en medio de un rango de fórmula compartida (en vez de coincidir con su final), el script frena con un error en vez de adivinar — mejor eso que arriesgarse a insertar en el lugar equivocado.

### Cómo se reajustan las fórmulas al insertar (validado con recálculo real, no sólo mirando el texto)

Confirmado con un prototipo antes de escribir el código final: **ni ExcelJS ni la mayoría de las librerías de este estilo reescriben las referencias de una fórmula cuando insertás filas**. Si la fila que antes era la 7050 (fórmula `=H7049+F7050-G7050`) pasa a ser la 7058 después de insertar 8 filas arriba, la fórmula se queda tal cual dice `H7049+F7050-G7050` — apuntando a la fila equivocada.

Por eso, después de insertar, el script recorre toda la hoja modificada (y el resto del libro, por si alguna otra hoja tiene una fórmula que cruza hacia ésta — encontramos 5 en total, ver más abajo) y reescribe cualquier referencia de celda cuya fila haya quedado desplazada. También hay un paso previo necesario: la planilla real usa "fórmulas compartidas" de Excel (`H7005:H7043` es una sola fórmula aplicada a 39 filas, no 39 fórmulas individuales) y ExcelJS tiene un bug conocido donde mover filas en el medio de un rango así puede dejar el archivo corrupto (`Shared Formula master must exist above and or left of clone`) — así que antes de mover nada, el script convierte a fórmulas literales cualquier rango que vaya a quedar afectado.

Esto **no se validó sólo mirando que el texto de la fórmula "se viera bien"** — se forzó un recálculo real con LibreOffice (`soffice --headless --convert-to`) después de cada prueba y se verificaron los números resultantes a mano. Ejemplo real de la hoja `BROU $` después de insertar 8 movimientos nuevos en la fila 7044: saldo 744.625,34 (el último real) → -463,44 → 744.161,90 → -538.809,72 → 205.352,18 → ... encadenando correctamente hasta la fila 7051, y la fila de "PENDIENTES DE DEBITO" (que quedó corrida a la fila 7057) siguió apuntando exactamente a la fila 7043 como lo hacía antes de mover nada.

También se probaron y confirmaron correctas, después de actualizar las 4 cuentas de una sola vez sobre la misma planilla:
- Las 5 fórmulas que cruzan de una hoja a otra en toda la planilla (2 "totales" generales en `BROU $`, 1 en `BROU U$S` hacia `SANTANDER U$S`, 2 en `SANTANDER $` hacia `SANTANDER U$S`) — todas quedaron exactamente iguales a como estaban (porque las filas que referencian están antes de donde insertamos en cada caso), y sus valores recalculados coinciden con los originales.
- Correr el mismo comando 2 veces seguidas: la segunda vez no agrega nada (el archivo queda byte a byte idéntico — se comparó el hash).

### Duplicados: por qué la clave NO es "misma fecha + misma descripción + mismo textoParaMatchCliente"

Este era el pedido original, pero probando contra la planilla real (que ya tenía cargados a mano los movimientos del 16/07) encontramos que **no funciona**: lo que una persona tipeó en `CONCEPTO` (ej. `"ROLA"`) casi nunca coincide texto-a-texto con la descripción cruda que manda el banco (ej. `"537806PP EMITIDO Rola Ltda"`). Con esa clave, el script no reconocía los movimientos ya cargados y los iba a duplicar.

En su lugar, la clave de duplicado es **fecha + tipo (débito/crédito) + monto**, pero contando repeticiones (no un simple sí/no): es normal que el mismo día haya más de un movimiento con igual monto —vimos 4 "COMISIONES BANCARIAS" de $77.71 el mismo día, una por cada transferencia—, así que cada coincidencia nueva "consume" una fila existente antes de considerarse duplicado. Si ya hay 4 cargadas y el estado de cuenta trae 4, no agrega nada; si trae 5, agrega sólo la 5ª. Esto se probó explícitamente simulando que faltaba una sola comisión de las 4, y el script agregó únicamente esa.

El texto (`CONCEPTO`/`textoParaMatchCliente`) no se usa para la detección de duplicados, sólo para lo que se escribe en la fila nueva.

### Caveats conocidos

- **`BROU EUROS` no se probó con un estado de cuenta real** (no había ninguno entre los archivos de ejemplo). El código es genérico y la hoja mostró el mismo patrón que las otras 4 (bloque contiguo terminando exacto donde termina el último rango de fórmula compartida), así que debería funcionar igual, pero conviene que la primera corrida real se revise a mano.
- **Correcciones con fecha retroactiva después de ya haber cargado fechas posteriores**: si el banco corrige/reenvía un movimiento de una fecha vieja después de que ya cargaron movimientos de fechas más nuevas, el script lo va a agregar al final (después de lo más nuevo), no intercalado cronológicamente en su lugar — porque el "ancla" para ese momento ya es la fila más reciente. Es un caso borde poco frecuente; si llega a pasar, esa fila puntual se puede reubicar a mano.
- **Columna B (`RECIBO...`)**: al re-guardar con ExcelJS, alguna celda numérica de esa columna (que no usamos para nada) puede pasar a guardarse como texto. No afecta plata ni fechas.

## Siguientes pasos sugeridos

1. **Detección de "ayer"**: al llamar `procesarEstadoDeCuenta` / `actualizarPlanilla`, calculen la fecha objetivo como "ayer hábil" (cuidado con fines de semana/feriados: viernes → el lunes hay que traer 3 días si el banco no generó movimiento sábado/domingo, pero igual filtrando por fecha esto no debería romper nada, solo devolvería 0 movimientos si no hubo actividad).
2. **Watcher de la carpeta en Halcón**: un proceso que mire la carpeta compartida, tan pronto aparezcan los 5 archivos del día (o con un timeout), dispare `actualizarPlanilla` para cada uno contra la planilla maestra.
3. **Matching con IA**: antes de escribir en `CONCEPTO`, mandarle a Hermes/DeepSeek el `textoParaMatchCliente` de cada movimiento + el listado de clientes, y usar la sugerencia (limpia) en vez del texto crudo del banco — hoy `actualizarPlanilla` usa `textoParaMatchCliente` tal cual viene del parser. Conviene guardar también un campo de confianza para que administración revise los casos dudosos en vez de confiar ciegamente.
4. **Backup antes de escribir**: dado que `actualizarPlanilla` sobreescribe la planilla maestra por defecto, vale la pena que el proceso automático guarde una copia con fecha/hora antes de cada corrida (o versione el archivo), para poder revertir fácil si algún día pasa algo inesperado.

## Nota sobre los archivos `.xls` de BROU

Aunque tienen extensión `.xls`, internamente fueron generados por WPS Spreadsheet, no por Excel — igual `xlsx` (SheetJS) los lee sin problema porque respeta el formato binario estándar BIFF. Si en algún momento BROU cambia su exportador y el archivo deja de abrir, lo primero a revisar es si sigue siendo `.xls` real (BIFF) o pasó a ser HTML disfrazado de `.xls` (common en bancos) — en ese caso el parseo cambiaría bastante y avisen para adaptar el script.
