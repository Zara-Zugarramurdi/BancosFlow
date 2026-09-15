# BancosFlow — contexto para trabajar en este proyecto

Notas para quien retome el proyecto (persona o asistente de IA). Junta lo que se aprendió
rompiendo cosas contra archivos y servicios reales; casi nada de esto se deduce leyendo el
código.

## Qué hace

Automatiza la actualización diaria de una planilla Excel de bancos. Administración sube a
SharePoint los estados de cuenta del día (4 cuentas: BROU Pesos, BROU Dólares, Santander
Pesos, Santander Dólares) y el proceso los parsea e inserta los movimientos nuevos en una
planilla maestra, sin duplicar.

Corre en una VM Rocky Linux (`apps-autadm`), en `/home/teledata/BancosFlow`, como `root`,
con un servicio systemd que vigila la carpeta del día.

## Reglas de trabajo

- **Pedir autorización antes de hacer commit o push.** Escribir y probar código está bien;
  publicarlo lo decide la persona.
- **No inventar rutas, usuarios ni nombres de servidor.** Ya pasó: se asumió
  `/opt/bancosflow` y `User=teledata` y el servicio no arrancaba. Si no está confirmado,
  preguntar.
- **Probar contra los archivos reales antes de dar algo por bueno.** Los casos borde de
  este dominio no se imaginan, aparecen.
- **Verificar antes de afirmar.** Más de una vez se dio por cierto algo que no se había
  comprobado y mandó el diagnóstico para el lado equivocado.

## Cosas que cuestan horas si no se saben

### SharePoint modifica los archivos de Office al recibirlos

Se suben 1.200.869 bytes y del otro lado quedan 1.208.496; al intento siguiente, 1.208.499.
Distinto cada vez, porque agrega metadatos propios.

Consecuencias:

- `rclone` compara tamaños, no coinciden, concluye `corrupted on transfer` y **borra lo que
  acaba de subir**. Por eso `flagsRcloneSubida` incluye `--ignore-size --ignore-checksum`.
  **No sacarlos.**
- Ninguna verificación puede comparar tamaños. Se comprueba que el archivo exista y que su
  fecha sea posterior al inicio de la subida.

### No se usa `rclone mount`, y no hay que volver a intentarlo

Fue lo primero que se probó, para no tocar el código. Falló, y el modo en que falla es peor
que el fallo: con montaje, el borrado descrito arriba ocurre de forma **asincrónica**, fuera
del proceso. El proceso escribía la planilla, veía éxito, guardaba respaldo, anotaba el
registro — y el archivo nunca llegaba a SharePoint. Sin mirar los logs de rclone, nadie se
enteraba.

Hoy `modoAcceso: "rclone"` baja con `rclone copy`, procesa en disco local y sube con
`rclone copyto`, **verificando**. Si falla, falla el proceso: sale por el journal y el
registro no se escribe, así que el ciclo siguiente reintenta.

### `--ignore-times` es necesario

Sin él, rclone compara la fecha de modificación y saltea la subida si coincide. Como
`--ignore-size` e `--ignore-checksum` le quitan las otras dos comparaciones, la fecha queda
como único criterio y basta con que coincida para que no suba nada — sin error. El código
sólo sube cuando hubo cambios reales, así que forzar la subida es lo correcto.

### El 404 `The upload session was not found`

Apareció el 14/09/2026 y bloqueó las subidas de `PlanillaBancos.xlsx` durante horas. Subir
el mismo contenido **con otro nombre** funcionaba; con ese nombre fallaba siempre, aun
después de limpiar toda la caché local y sin montaje de por medio. Parece una sesión de
subida multiparte colgada del lado de SharePoint para ese ítem.

Se destrabó subiendo con otro nombre y renombrando con `rclone moveto` (un rename es
metadatos, no vuelve a tocar la sesión). Si vuelve a pasar, esa es la salida.

### `inotify` no funciona sobre CIFS/SMB

El kernel no recibe eventos de escrituras hechas desde otra máquina. Por eso el poller hace
polling y no vigilancia de eventos. No es pereza.

### Los archivos de BROU son `.xls` binario viejo (BIFF)

Excel para la web no los abre. Es el motivo por el que Power Automate y los Office Scripts
quedaron descartados: no pueden leer la mitad de las cuentas.

### `Bad uncompressed size: NNN != 0`

Ruido de SheetJS leyendo los `.xlsx` de Santander. Sale por `stderr`, es inofensivo y
aparece desde el primer día. No es síntoma de nada.

## Trampas de la planilla real

La planilla se mantiene a mano desde 2022 y tiene de todo:

- **Fechas mal tipeadas.** En `BROU $` hay filas con fecha de 2027 (error por 2023) y una
  de 1928. Por eso "hasta qué día está cargada la hoja" se calcula con la **fila ancla**
  (última fila del bloque contiguo con fecha), **nunca** con `max(fecha)`: con el máximo, el
  proceso creería estar al día hasta 2027 y no agregaría nada nunca.
- **Bloques que no son movimientos.** Después del último movimiento real puede haber
  proyecciones (`PENDIENTES DE DEBITO`), listas de referencia (retenciones judiciales) y
  cronogramas de préstamos que reutilizan las mismas columnas. Por eso el ancla se detecta
  mirando **sólo la columna FECHA**: todo movimiento real tiene fecha, ningún bloque de
  relleno la tiene.
- **Filas fantasma** con la fórmula de SALDO arrastrada de más, sin datos.
- **Fechas guardadas como texto** (`"27/07/2026"`) en vez de fecha de Excel.
- **Fórmulas compartidas** de Excel (un `H7005:H7043` en vez de 39 fórmulas). ExcelJS tiene
  un bug al mover filas dentro de un rango así; por eso se convierten a literales antes de
  insertar.
- **ExcelJS no reajusta referencias** al insertar filas: hay que reescribirlas a mano. Y
  `cell.style = otraCelda.style` copia una **referencia compartida**, no una copia — pintar
  una fila terminaba pintando otra.

## Convenciones del dominio

- En la planilla, **DEBE = plata que entra** y **HABER = plata que sale**. Es la convención
  que ya usaba Administración, al revés de lo que dice el banco.
- La deduplicación es por **fecha + tipo + monto**, contando repeticiones. **No** por texto:
  lo que una persona tipeó (`"ROLA"`) nunca coincide con lo que manda el banco
  (`"537806PP EMITIDO Rola Ltda"`). Es la protección real contra duplicados y funciona
  aunque se pierda el registro de procesados.
- Las filas nuevas se insertan **justo después del último movimiento real**, no al final del
  archivo, y se pintan enteras de amarillo (`FFFFFF99`).
- El texto de CONCEPTO se completa con el campo de respaldo si es corto (menos de 5
  caracteres) o no tiene letras. Se **concatena**, no se reemplaza: un texto corto puede ser
  el cliente (`ITAU`, `BSE`, `UTE`).

## Cómo verificar cambios

No hay tests automatizados (deuda conocida). La verificación es manual y conviene hacerla
siempre:

```bash
npx tsc                                          # compila sin errores
bash scripts/verificar-entorno.sh                # entorno y conexión
node dist/procesarCarpetaDelDia.js --dry-run     # no escribe nada
```

Para cambios que tocan la escritura de la planilla, además:

- Correr dos veces: la segunda debe agregar 0 y dejar el archivo idéntico.
- Verificar los saldos con un recálculo real (LibreOffice
  `soffice --headless --convert-to xlsx`), no mirando el texto de las fórmulas.
- Confirmar que las hojas no involucradas quedan sin diferencias.

Hay un `rclone` simulado útil para probar sin tocar SharePoint: imita los subcomandos y
reproduce el retoque a los archivos de Office. No está en el repo; se puede rehacer.

## Archivos principales

| Archivo | Qué hace |
|---|---|
| `accountIdentifier.ts` | A qué cuenta corresponde un estado de cuenta |
| `parseBrou.ts` / `parseSantander.ts` | Extraen movimientos de cada formato |
| `textoConcepto.ts` | Criterio común para completar el CONCEPTO |
| `actualizarPlanilla.ts` | Núcleo: ancla, dedup, inserción, fórmulas, estilos |
| `actualizarDesdeUltimaFecha.ts` | Variante sin fecha: detecta hasta dónde está cargada |
| `almacenamiento.ts` | Aísla de dónde salen los archivos (montaje o rclone) |
| `rclone.ts` | Envoltorio del binario, con verificación de subida |
| `procesarCarpetaDelDia.ts` | Orquestador |
| `poller.ts` | Servicio que vigila la carpeta |
| `config.ts` | Todos los parámetros ajustables |

## Pendientes conocidos

- **Sin tests automatizados.** Todo se verifica a mano. Es la deuda más grande.
- **Matching de clientes con IA** — el objetivo original del proyecto, nunca empezado.
  Requiere definición de SGSI sobre modelo local vs. externo.
- **BROU Euros** implementado pero nunca probado con un estado de cuenta real.
- **Sin alertas**: si el proceso falla varios días, nadie se entera salvo que mire el journal.
- El documento para SGSI quedó desactualizado tras la migración a SharePoint.
