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
- **`src/actualizarDesdeUltimaFecha.ts`** — igual que el anterior pero sin pasarle fecha: detecta hasta qué día está cargada la planilla y agrega todo lo posterior. Ver sección dedicada más abajo.

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

Cada `movimiento` trae el campo **`textoParaMatchCliente`** — es el texto que le pasarías a Hermes/DeepSeek junto con tu listado de clientes para que sugiera a quién corresponde, y es también lo que se escribe en la columna CONCEPTO de la planilla.

- **BROU**: la columna "Asunto"; si viene vacía, la Descripción.
- **BROU**: la columna "Asunto"; si no se vale por sí sola, se completa con la Descripción.
- **Santander**: la Descripción (no hay columna de asunto separada); si no se vale por sí sola, se completa con el Tipo de Movimiento, y como último recurso con la Referencia.

El criterio de "no se vale por sí solo" es común a ambos bancos y vive en `src/textoConcepto.ts`. Alcanza con que se cumpla uno de los tres:

1. **Está vacío.** No hay nada que preservar.
2. **Tiene menos de 5 caracteres.** Ningún texto tan corto identifica por sí solo a un cliente de forma confiable.
3. **No contiene ninguna letra.** Un texto de puros números o signos (`"3507"`, `"24609930"`, `"-"`, `"--"`) es una referencia interna del banco, no un nombre. Se expresa como "no tiene letras" en vez de "son todos dígitos" para cubrir también números con separadores (`"3.507"`) sin enumerar cada variante.

En esos casos **se concatena, no se reemplaza**: `"3507"` + `"TRF SPI PAGO PROV."` queda `"3507 - TRF SPI PAGO PROV."`. Se concatena incluso cuando el principal es un `"-"` (queda `"- - DEPOSITO CHEQUES  CLEARING"`). Es una decisión explícita: se prefiere un poco de ruido antes que arriesgarse a descartar información. Un texto corto puede ser justamente el cliente — en los estados de cuenta reales aparecen `"ITAU"` y `"BBVA"` como asunto de BROU, y en la planilla hay cientos de conceptos cargados a mano con siglas de 3 y 4 letras (DUA, BSE, UTE, SMI, OSE, BPS, DGI...). Única excepción: si el principal está vacío va sólo el respaldo, porque concatenar dejaría un separador colgando.

Medido sobre los 4 estados de cuenta del 18/08 (72 movimientos), la regla se activa en 6: los `"ITAU"`/`"BBVA"` y `"3507"` de BROU, y un `"24609930"` de Santander. El campo `descripcion` mantiene siempre el texto crudo del banco, sin tocar.

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

Las filas insertadas se resaltan en amarillo (`FFFFFF99`, el mismo tono que usa la planilla a mano) **en toda su extensión, de FECHA a SALDO** — no sólo las celdas con contenido, para que no queden resaltadas a medias.

### Dónde se insertan las filas nuevas (y por qué esto cambió durante el desarrollo)

La primera versión de este script agregaba las filas nuevas después de la **última fila con cualquier dato** de la hoja, para no arriesgarse a insertar en el medio de miles de fórmulas encadenadas. Probando contra la hoja real `"BROU $"` encontramos el problema: después del último movimiento real (fila 7043, 17/07/2026) hay varios bloques que **no son movimientos del día a día** — una proyección `"PENDIENTES DE DEBITO"`, un bloque viejo con fechas de 2025 que quedó pegado de una versión anterior, una lista fija de `"RETENCION JUDICIAL..."`, y el cronograma de un préstamo (`"PROYECTO BOTIJAS"`) que reutiliza la columna E para otra cosa. La "última fila con datos" terminaba **55 filas más abajo** del final real del ledger, así que los movimientos nuevos quedaban invisibles para quien mira la hoja esperando encontrarlos justo debajo del último movimiento — el mismo problema que reportaron.

La solución obvia — "usar la fila con la fecha más reciente" — tampoco es segura: hay fechas mal tipeadas a mano en medio del ledger real (`30/12/2026` en `SANTANDER $` en medio de datos de mediados de 2026; `29/05/2028` en `SANTANDER U$S` en medio de datos de 2024) que hubieran hecho que el script insertara en pleno medio de una cadena de fórmulas activa. Mucho más peligroso que el problema original.

La heurística que sí funciona: el ledger real es un **bloque contiguo de filas que tienen FECHA**, desde la fila 2 hasta la última antes del primer salto de 3+ filas seguidas sin fecha. El criterio clave es mirar **únicamente la columna FECHA** — no CONCEPTO/DEBE/HABER/SALDO. Todo movimiento real tiene fecha; ninguna fila de relleno, proyección o fórmula arrastrada de más la tiene.

Esto último se aprendió a los golpes: la primera versión consideraba que una fila "tenía datos" si *cualquiera* de esas 5 columnas tenía algo, y eso rompía la detección en `SANTANDER $` de dos formas combinadas:

- Entre el último movimiento real (fila 3454) y el bloque de proyección "PENDIENTES DE DEBITO" (3457) había sólo **2** filas vacías, no 3 — así que el corte no se activaba ahí, y el bloque de proyección (que tiene CONCEPTO y SALDO) quedaba contado como parte del ledger.
- Después de ese bloque, alguien había arrastrado la fórmula de SALDO 5 filas de más (3462-3466), sin fecha, sin concepto y sin importes. Como "tenían algo" en SALDO, también contaban.

Resultado: el ancla daba 3466 en vez de 3454, y los movimientos nuevos terminaban insertados **después** del bloque de pendientes, que se supone va siempre al final. Con el criterio de "sólo FECHA", el corte cae exacto en 3454 sin importar cuántas filas fantasma haya después. Se verificó además que el cambio no altera el resultado en ninguna de las otras 4 hojas, ni en la versión vieja de la planilla (misma fila de ancla que antes en los 5 casos).

Como red de seguridad adicional, el ancla se cruza con los rangos de "fórmula compartida" de Excel en la columna SALDO: si el punto de corte cayera *en medio* de uno de esos rangos, el script frena con un error en vez de adivinar.

**Nota sobre fechas en texto:** algunas filas de la planilla real tienen la fecha cargada como texto (`"27/07/2026"`) en vez de como fecha de Excel — pasa cuando se pega desde otro lado. La función que lee fechas entiende los tres formatos (fecha real, número de serie de Excel, y texto `dd/mm/yyyy` validado), así que esas filas no quedan invisibles ni para la detección del ancla ni para el chequeo de duplicados.

### Cómo se reajustan las fórmulas al insertar (validado con recálculo real, no sólo mirando el texto)

Confirmado con un prototipo antes de escribir el código final: **ni ExcelJS ni la mayoría de las librerías de este estilo reescriben las referencias de una fórmula cuando insertás filas**. Si la fila que antes era la 7050 (fórmula `=H7049+F7050-G7050`) pasa a ser la 7058 después de insertar 8 filas arriba, la fórmula se queda tal cual dice `H7049+F7050-G7050` — apuntando a la fila equivocada.

Por eso, después de insertar, el script recorre toda la hoja modificada (y el resto del libro, por si alguna otra hoja tiene una fórmula que cruza hacia ésta — encontramos 5 en total, ver más abajo) y reescribe cualquier referencia de celda cuya fila haya quedado desplazada. También hay un paso previo necesario: la planilla real usa "fórmulas compartidas" de Excel (`H7005:H7043` es una sola fórmula aplicada a 39 filas, no 39 fórmulas individuales) y ExcelJS tiene un bug conocido donde mover filas en el medio de un rango así puede dejar el archivo corrupto (`Shared Formula master must exist above and or left of clone`) — así que antes de mover nada, el script convierte a fórmulas literales cualquier rango que vaya a quedar afectado.

Esto **no se validó sólo mirando que el texto de la fórmula "se viera bien"** — se forzó un recálculo real con LibreOffice (`soffice --headless --convert-to`) después de cada prueba y se verificaron los números resultantes a mano. Ejemplo real de la hoja `BROU $` después de insertar 8 movimientos nuevos en la fila 7044: saldo 744.625,34 (el último real) → -463,44 → 744.161,90 → -538.809,72 → 205.352,18 → ... encadenando correctamente hasta la fila 7051, y la fila de "PENDIENTES DE DEBITO" (que quedó corrida a la fila 7057) siguió apuntando exactamente a la fila 7043 como lo hacía antes de mover nada.

También se probaron y confirmaron correctas, después de actualizar las 4 cuentas de una sola vez sobre la misma planilla:
- Las 5 fórmulas que cruzan de una hoja a otra en toda la planilla (2 "totales" generales en `BROU $`, 1 en `BROU U$S` hacia `SANTANDER U$S`, 2 en `SANTANDER $` hacia `SANTANDER U$S`) — todas quedaron exactamente iguales a como estaban (porque las filas que referencian están antes de donde insertamos en cada caso), y sus valores recalculados coinciden con los originales.
- Correr el mismo comando 2 veces seguidas: la segunda vez no agrega nada (el archivo queda byte a byte idéntico — se comparó el hash).
- Las hojas que no se tocan en cada corrida (`BROU EUROS`, `SANTANDER CA $U`, `SANTANDER CA U$S`, `SANTANDER EUROS`, `DIFERIDOS TD`, `CUENTAS EMPLEADOS`) quedan con **0 diferencias** celda por celda.

Esta batería se corrió sobre **dos versiones distintas de la planilla real** (la del 20/07 y la del 04/08) con sus respectivos estados de cuenta, para asegurarse de que los cambios de detección de ancla no rompieran el comportamiento que ya funcionaba. En las 5 hojas de la planilla vieja, el ancla calculada es idéntica antes y después del cambio de criterio.

### Estado del bloque "PENDIENTES DE DEBITO" al insertar

Cuando se insertan movimientos nuevos, el bloque de proyección "PENDIENTES DE DEBITO" se corre hacia abajo correctamente, pero su fórmula inicial (`=H{fila}`) sigue apuntando a la fila que apuntaba antes — o sea, al saldo del que era el último movimiento real *antes* de la corrida, no al nuevo. Ej.: tras cargar los movimientos del 03/08 en `SANTANDER $`, el bloque sigue partiendo del saldo del 31/07.

Esto se consultó con administración y se decidió **dejarlo así a propósito**: el script no toca esa referencia. Si en algún momento se quiere que la proyección parta siempre del saldo más reciente, es un cambio chico pero hay que pedirlo explícitamente.

### Duplicados: por qué la clave NO es "misma fecha + misma descripción + mismo textoParaMatchCliente"

Este era el pedido original, pero probando contra la planilla real (que ya tenía cargados a mano los movimientos del 16/07) encontramos que **no funciona**: lo que una persona tipeó en `CONCEPTO` (ej. `"ROLA"`) casi nunca coincide texto-a-texto con la descripción cruda que manda el banco (ej. `"537806PP EMITIDO Rola Ltda"`). Con esa clave, el script no reconocía los movimientos ya cargados y los iba a duplicar.

En su lugar, la clave de duplicado es **fecha + tipo (débito/crédito) + monto**, pero contando repeticiones (no un simple sí/no): es normal que el mismo día haya más de un movimiento con igual monto —vimos 4 "COMISIONES BANCARIAS" de $77.71 el mismo día, una por cada transferencia—, así que cada coincidencia nueva "consume" una fila existente antes de considerarse duplicado. Si ya hay 4 cargadas y el estado de cuenta trae 4, no agrega nada; si trae 5, agrega sólo la 5ª. Esto se probó explícitamente simulando que faltaba una sola comisión de las 4, y el script agregó únicamente esa.

El texto (`CONCEPTO`/`textoParaMatchCliente`) no se usa para la detección de duplicados, sólo para lo que se escribe en la fila nueva.

### Caveats conocidos

- **`BROU EUROS` no se probó con un estado de cuenta real** (no había ninguno entre los archivos de ejemplo). El código es genérico y la hoja mostró el mismo patrón que las otras 4 (bloque contiguo de filas con fecha, terminando de forma limpia), así que debería funcionar igual, pero conviene que la primera corrida real se revise a mano.
- **Correcciones con fecha retroactiva después de ya haber cargado fechas posteriores**: si el banco corrige/reenvía un movimiento de una fecha vieja después de que ya cargaron movimientos de fechas más nuevas, el script lo va a agregar al final (después de lo más nuevo), no intercalado cronológicamente en su lugar — porque el "ancla" para ese momento ya es la fila más reciente. Es un caso borde poco frecuente; si llega a pasar, esa fila puntual se puede reubicar a mano.
- **Columna B (`RECIBO...`)**: al re-guardar con ExcelJS, alguna celda numérica de esa columna (que no usamos para nada) puede pasar a guardarse como texto. No afecta plata ni fechas.

## Poner la planilla al día sin indicar fecha

**`src/actualizarDesdeUltimaFecha.ts`** — variante de `actualizarPlanilla` que no recibe una fecha: mira hasta qué día está cargada la hoja y agrega todo lo que el estado de cuenta tenga de ahí en adelante, en una sola pasada.

```bash
node dist/actualizarDesdeUltimaFecha.js "C:\ruta\bancos.xlsx" "C:\ruta\estado.xlsx" ["C:\ruta\salida.xlsx"]
```

Reutiliza el núcleo de inserción de `actualizarPlanilla` (`aplicarMovimientosAPlanilla`), así que hereda tal cual la deduplicación, el cálculo del ancla, el desarmado de fórmulas compartidas, el corrimiento de referencias, la copia de estilos y el resaltado en amarillo. Lo propio de este archivo es sólo *qué* movimientos elegir y *en qué orden* entregarlos.

### Cómo sabe hasta dónde está cargada la planilla

Usa la fecha de la **fila ancla** (la última fila del bloque contiguo del ledger), **no** el máximo de las fechas. Es deliberado: la posición manda, no el valor. En la planilla real, la fecha máxima de `BROU $` es 28/02/2027 — un error de tipeo por 2023, ya que la fila siguiente vuelve a 01/03/2023; también hay una fila con fecha de 1928. Si el proceso preguntara "cuál es la fecha más alta", concluiría que la planilla está al día hasta 2027 y no volvería a agregar nada nunca.

Como la fila ancla misma podría tener una fecha mal tipeada, hay dos chequeos de cordura que abortan con un mensaje claro en lugar de saltearse movimientos en silencio:

- La última fecha cargada no puede ser posterior a hoy.
- No puede haber un salto de más de 90 días respecto de la fila con fecha inmediatamente anterior.

### El rango arranca en la última fecha cargada, inclusive

No en el día siguiente. Si ese día se había cargado a mitad de jornada, o el banco sumó movimientos tarde, arrancar al día siguiente los perdería para siempre y en silencio. Al incluirlo, la deduplicación (fecha + tipo + monto, contando repeticiones) descarta los que ya están y agrega sólo los que faltaban. Se puede cambiar a estricto con `incluirUltimaFecha: false`.

Esto no es teórico: probando con la planilla del 10/08 y su estado de cuenta de BROU Pesos, el proceso encontró **3 movimientos del 07/08 que faltaban** además del único del 10/08. Con el criterio estricto se habrían perdido.

### Orden de inserción

Los movimientos se ordenan cronológicamente antes de insertar, de más viejo a más nuevo, manteniendo estable el orden dentro de un mismo día. Hace falta porque **BROU entrega sus estados de cuenta con el movimiento más nuevo primero** (verificado con archivos reales: `10/08 → 07/08 → ... → 06/08`), mientras que Santander los entrega en orden ascendente. Cuando se procesaba un solo día daba igual; al insertar un rango de varios días, respetar el orden del archivo dejaría el ledger al revés.

Dentro de un mismo día se conserva el orden del archivo: es el único criterio disponible, ya que los estados de cuenta no traen hora.

### Días que el estado de cuenta no cubre

Si el movimiento más antiguo del estado de cuenta es posterior a la última fecha cargada, quedan días en el medio que nadie va a cargar (ej. planilla al 20/07 y estado de cuenta que arranca el 28/07). El proceso **avisa y continúa** — si Administración no descargó esos días, no hay nada que el script pueda hacer, pero conviene que quede a la vista.

### Pruebas

- **Equivalencia**: procesar un rango con esta función da un resultado idéntico celda por celda a correr `actualizarPlanilla` día por día sobre ese mismo rango.
- **Regresión**: tras extraer el núcleo compartido, `actualizarPlanilla` produce exactamente los mismos resultados que antes del refactor (mismas cantidades, mismas filas) sobre las planillas de julio y agosto.
- **Punta a punta** sobre la planilla del 13/08 con las 4 cuentas, encontrando cada hoja en un estado distinto (dos al día, dos atrasadas): 14 movimientos insertados en total, orden cronológico correcto, idempotencia (mismo hash al repetir), saldos verificados con recálculo real en LibreOffice y 0 diferencias en las 6 hojas no involucradas.

### Limitación conocida

Si la fila ancla tuviera una fecha mal tipeada **hacia atrás** (ej. enero en vez de julio), el proceso re-escanearía desde esa fecha y podría insertar movimientos viejos al final del ledger, fuera de orden cronológico. No se bloquea porque en el ledger real hay varios saltos hacia atrás legítimos y el chequeo saltaría en falso constantemente. Los typos hacia adelante, que son los peligrosos porque harían saltear meses de movimientos, sí están cubiertos.

# Automatización en el servidor (proceso desatendido)

Además de los comandos que se corren a mano, el proyecto incluye las piezas para que el proceso corra solo en la VM, leyendo la carpeta compartida donde Administración deja los archivos.

> **Estado:** en construcción. Esta sección se va completando a medida que se implementa cada pieza.

## Configuración: `src/config.ts`

Todos los parámetros ajustables viven en un solo lugar. Los valores por defecto están en el código y se pisan con un archivo JSON, sin recompilar nada: por defecto `config/bancosflow.config.json`, o donde apunte la variable de entorno `BANCOSFLOW_CONFIG`. Sólo hace falta escribir las claves que se quieran cambiar.

| Parámetro | Por defecto | Qué controla |
|---|---|---|
| `rutaCarpetaBase` | `/media/windowsshare/.../PlanillaBancos` | Carpeta raíz sobre la que trabaja el proceso. Es la única que se toca. |
| `ubicacionPlanilla` | `carpetaDelDia` | De dónde sale la planilla: la carpeta del día, o `maestraFija`. |
| `rutaPlanillaMaestra` | `""` | Ruta de la planilla única. Sólo se usa con `maestraFija`. |
| `rutaBackups` | `/home/teledata/backups` | Dónde se guardan los respaldos: disco local de la VM, **no** el fileserver. |
| `retencionBackupsDias` | `90` | Días que se conservan los respaldos. |
| `intervaloPollSegundos` | `60` | Cada cuánto se revisa si cambió algo. |
| `esperaSinCambiosSegundos` | `60` | Cuánto tiene que estar quieta la carpeta antes de procesar. |
| `horaCreacionCarpetas` | `00:00` | A qué hora se crean las carpetas del día. |
| `zonaHoraria` | `America/Montevideo` | Con qué zona se decide "qué día es hoy". |
| `nombresMeses` | Enero…Diciembre | Nombres de las carpetas de mes. |
| `diaConCeroAdelante` | `false` | Si el día va como `9` o como `09`. |
| `nombreArchivoRegistro` | `.bancosflow.json` | Archivo de control dentro de cada carpeta del día. |

Correr `node dist/config.js` imprime la configuración efectiva; sirve para verificar en la VM que el archivo JSON se está leyendo. Las claves desconocidas se avisan por consola en vez de ignorarse en silencio, porque casi siempre son errores de tipeo que dejarían el proceso corriendo con el valor por defecto.

### Sobre `ubicacionPlanilla`

Las carpetas por día se crean **siempre**, en los dos modos, porque los estados de cuenta se suben por día de todas formas. Lo único que cambia es de dónde sale la planilla: hoy la sube Administración a la carpeta del día (`carpetaDelDia`); si mañana se prefiere una única planilla en un lugar fijo, se cambia esa clave y se completa `rutaPlanillaMaestra`, sin tocar código.

## Estructura de carpetas: `src/rutasPlanillaBancos.ts`

```
<rutaCarpetaBase>/<año>/<Mes>/<día>     ej: .../PlanillaBancos/2026/Agosto/19
```

```bash
node dist/rutasPlanillaBancos.js [--simular] [--fecha yyyy-mm-dd]
```

Es idempotente: si las carpetas ya existen no hace nada. Con `--simular` informa qué crearía sin tocar el disco. Pensado para correr todos los días a las 00:00.

**Por qué la zona horaria importa.** El día se calcula con `Intl` en la zona configurada, no con los métodos locales de `Date`. Si la VM está en UTC y no se fija esto, entre las 21:00 y la medianoche de Uruguay el proceso ya estaría usando la carpeta del día siguiente. Verificado: el instante `2026-08-20T02:30:00Z` da día **19** con `America/Montevideo` y día **20** con `UTC`.

## Clasificación de archivos: `src/clasificarArchivos.ts`

```bash
node dist/clasificarArchivos.js <carpeta>
```

Decide qué es cada archivo **por contenido, nunca por nombre** — los nombres reales varían todo el tiempo (`Copia de Planilla BANCOS desde 12-22 18-08.xlsx`, `Detalle_Movimiento_Cuenta (7).xls`) y depender de ellos sería frágil.

El orden de las decisiones es deliberado:

1. Se descarta lo que no es Excel, el archivo de control y los temporales de Excel (`~$...`).
2. Se intenta identificar la cuenta con `accountIdentifier`. Si funciona, es un estado de cuenta.
3. Si no, se mira si el libro tiene **las 5 hojas** del ledger (`BROU $`, `BROU U$S`, `BROU EUROS`, `SANTANDER $`, `SANTANDER U$S`). Si las tiene, es la planilla.

Un estado de cuenta tiene una sola hoja, así que no hay forma de confundirlos. Para mirar las hojas se usa `bookSheets`, que lee sólo la lista de nombres en vez de parsear el libro entero.

**Archivos abiertos en Excel.** La presencia de un `~$...` indica que alguien tiene el libro abierto; la clasificación lo reporta con `hayArchivosAbiertos` para que el orquestador postergue el procesamiento hasta el próximo ciclo, en vez de arriesgarse a que la persona guarde encima de lo insertado.

**Planillas duplicadas.** Si hay más de una planilla en la carpeta, gana la de fecha de modificación más reciente y las otras quedan listadas en `planillasDuplicadas` para descartarlas.

Probado contra la carpeta real del 18/08 con las 4 cuentas más ruido (un `.txt`, un temporal de Excel y una segunda planilla más vieja): identifica correctamente la planilla, las 4 cuentas, ignora el resto y avisa del archivo abierto.

## Registro de procesados: `src/registroProcesados.ts`

Dentro de cada carpeta del día queda un `.bancosflow.json` con qué estados de cuenta se procesaron, cuándo y con qué resultado.

Se eligió un archivo de control aparte en vez de renombrar los originales (agregarles `.procesado`) por dos motivos: son archivos que subió una persona y modificarlos es intrusivo, y sobre un montaje SMB un rename puede fallar si alguien tiene el archivo abierto.

Se guarda el **SHA-256** de cada archivo, no sólo el nombre: si alguien vuelve a subir un estado de cuenta corregido con el mismo nombre, el hash cambia y se reprocesa en vez de darlo por hecho. Verificado en pruebas.

> Esto evita trabajo repetido, pero **no** es la protección contra duplicados. Esa sigue siendo la deduplicación por fecha+tipo+monto, que funciona aunque el registro se borre. Por eso, si el JSON está corrupto o ilegible, el proceso avisa y sigue con un registro vacío en vez de frenarse.

```bash
node dist/registroProcesados.js <ruta del .bancosflow.json>   # ver el contenido
```

## Respaldo: `src/respaldo.ts`

Antes de escribir la planilla se guarda una copia en `rutaBackups` con fecha y hora, y se purgan las que superan `retencionBackupsDias` (90 por defecto).

Dos decisiones:

- **Sólo se respalda la planilla.** Es lo único que el proceso modifica y lo único irreemplazable; los estados de cuenta quedan en el fileserver y se pueden volver a bajar del banco.
- **Los respaldos van al disco local de la VM, no al fileserver.** Si el problema fuera justamente el montaje de red o algo que corrompa esa carpeta, tener la copia en el mismo lugar no serviría.

```bash
node dist/respaldo.js purgar [--simular]
node dist/respaldo.js copiar <planilla> [--simular]
```

## Orquestador: `src/procesarCarpetaDelDia.ts`

```bash
node dist/procesarCarpetaDelDia.js [--dry-run] [--fecha yyyy-mm-dd] [--crear-carpetas]
```

Secuencia: ubica la carpeta y la planilla → clasifica → descarta lo ya procesado → respalda → corre `actualizarDesdeUltimaFecha` por cada estado de cuenta pendiente → anota en el registro y purga respaldos viejos.

**`--dry-run` informa exactamente qué haría sin escribir nada.** Para poder decir cuántos movimientos entrarían, procesa contra una copia temporal de la planilla que después borra. Conviene correr así los primeros días.

Situaciones en las que **no** procesa, y por qué:

| Motivo | Comportamiento |
|---|---|
| Existe un archivo `PAUSADO` en la carpeta base | Freno de mano: no toca nada |
| La carpeta del día no existe todavía | Espera |
| Hay un `~$...` (Excel abierto) | Espera al próximo ciclo, para no escribir mientras alguien tiene el archivo abierto |
| Todavía no se subió la planilla | Espera: procesar contra la planilla equivocada es peor que no procesar |
| No hay estados de cuenta | Espera |
| Todos los estados de cuenta ya se procesaron | No hace nada |

Si un estado de cuenta falla, se anota el error y **se sigue con los demás**. Como no queda registrado como procesado, se reintenta solo en el próximo ciclo.

## Poller: `src/poller.ts`

```bash
node dist/poller.js               # modo servicio, no termina
node dist/poller.js --ciclos 3    # para probar a mano
```

**Por qué polling y no `inotify`:** `inotify` no funciona sobre montajes CIFS/SMB — el kernel no recibe eventos de escrituras hechas desde otra máquina. La opción "elegante" no está disponible en este escenario. A un ciclo por minuto el costo es despreciable.

**Cómo evita procesar a medias:** toma una huella de la carpeta (nombres + tamaños + fechas). Si cambió respecto del ciclo anterior, anota el momento y no procesa. Recién cuando pasaron `esperaSinCambiosSegundos` sin ningún cambio, procesa. Eso da tiempo a subir todos los archivos y evita agarrar uno a medio copiar por la red.

Probado simulando una subida progresiva (planilla, después un estado de cuenta, después otro): detectó cada cambio, esperó a que se aquietara y procesó **una sola vez** con todo junto. Un archivo que llega más tarde se procesa solo, sin repetir los anteriores.

Otros detalles: al cambiar de día reinicia el seguimiento y crea la carpeta si falta (así se recupera aunque el timer de las 00:00 no haya corrido), y un error puntual —por ejemplo el montaje caído un momento— se registra y se reintenta en el próximo ciclo, sin tirar abajo el servicio.

## Instalación en la VM

```bash
sudo mkdir -p /opt/bancosflow && cd /opt/bancosflow
sudo git clone https://github.com/Zara-Zugarramurdi/BancosFlow.git .
sudo npm install && sudo npx tsc
sudo mkdir -p config && sudo nano config/bancosflow.config.json   # ver más abajo
sudo mkdir -p /home/teledata/backups
sudo chown -R teledata:teledata /opt/bancosflow /home/teledata/backups

sudo cp systemd/*.service systemd/*.timer /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now bancosflow-carpetas.timer
sudo systemctl enable --now bancosflow-poller.service
```

Config mínima de producción (el resto toma los valores por defecto):

```json
{
  "rutaCarpetaBase": "/media/windowsshare/contable/privado/ADMINISTRACION/PlanillaBancos",
  "rutaBackups": "/home/teledata/backups"
}
```

**Antes de habilitar el servicio** conviene correr unos días a mano con `--dry-run` y revisar que decida bien.

Si el share se monta con una unidad de systemd, descomentar las líneas `Requires=` / `After=` en `bancosflow-poller.service` con el nombre correcto de esa unidad, para que el poller no arranque antes de que el montaje esté disponible.

### Control del poller

```bash
sudo systemctl stop bancosflow-poller       # pausar
sudo systemctl start bancosflow-poller      # reanudar
sudo systemctl restart bancosflow-poller    # tras cambiar la configuración
systemctl status bancosflow-poller          # estado
journalctl -u bancosflow-poller -f          # log en vivo
journalctl -u bancosflow-poller --since today
```

Además hay un **freno de mano sin acceso a la VM**: si se crea un archivo llamado `PAUSADO` en la carpeta base del share, el proceso no toca nada. Sirve para que Administración pueda frenarlo, por ejemplo mientras reorganiza la planilla a mano. Se reanuda borrando el archivo.

## Siguientes pasos sugeridos

1. **Detección de "ayer"**: al llamar `procesarEstadoDeCuenta` / `actualizarPlanilla`, calculen la fecha objetivo como "ayer hábil" (cuidado con fines de semana/feriados: viernes → el lunes hay que traer 3 días si el banco no generó movimiento sábado/domingo, pero igual filtrando por fecha esto no debería romper nada, solo devolvería 0 movimientos si no hubo actividad).
2. **Watcher de la carpeta en Halcón**: un proceso que mire la carpeta compartida, tan pronto aparezcan los 5 archivos del día (o con un timeout), dispare `actualizarPlanilla` para cada uno contra la planilla maestra.
3. **Matching con IA**: antes de escribir en `CONCEPTO`, mandarle a Hermes/DeepSeek el `textoParaMatchCliente` de cada movimiento + el listado de clientes, y usar la sugerencia (limpia) en vez del texto crudo del banco — hoy `actualizarPlanilla` usa `textoParaMatchCliente` tal cual viene del parser. Conviene guardar también un campo de confianza para que administración revise los casos dudosos en vez de confiar ciegamente.
4. **Backup antes de escribir**: dado que `actualizarPlanilla` sobreescribe la planilla maestra por defecto, vale la pena que el proceso automático guarde una copia con fecha/hora antes de cada corrida (o versione el archivo), para poder revertir fácil si algún día pasa algo inesperado.

## Nota sobre los archivos `.xls` de BROU

Aunque tienen extensión `.xls`, internamente fueron generados por WPS Spreadsheet, no por Excel — igual `xlsx` (SheetJS) los lee sin problema porque respeta el formato binario estándar BIFF. Si en algún momento BROU cambia su exportador y el archivo deja de abrir, lo primero a revisar es si sigue siendo `.xls` real (BIFF) o pasó a ser HTML disfrazado de `.xls` (common en bancos) — en ese caso el parseo cambiaría bastante y avisen para adaptar el script.
