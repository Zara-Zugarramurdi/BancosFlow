/**
 * poller.ts
 *
 * Revisa periódicamente la carpeta del día y dispara el procesamiento cuando detecta
 * que Administración terminó de subir archivos.
 *
 * ¿Por qué polling y no `inotify`? Porque **`inotify` no funciona sobre montajes
 * CIFS/SMB**: el kernel no recibe eventos de escrituras hechas desde otra máquina, así
 * que la opción "elegante" simplemente no está disponible en este escenario. El polling
 * es la alternativa que sí funciona, y a un ciclo por minuto no consume nada.
 *
 * La lógica de "esperar a que se aquiete" evita dos problemas concretos:
 *   - procesar cuando todavía falta subir archivos (se procesaría de a uno);
 *   - agarrar un archivo a medio copiar por la red.
 *
 * Se toma una huella de la carpeta (nombres + tamaños + fechas). Si cambió respecto del
 * ciclo anterior, se anota el momento y NO se procesa. Recién cuando pasaron
 * `esperaSinCambiosSegundos` sin ningún cambio, se procesa.
 *
 * El estado vive en memoria a propósito: pensado para correr como servicio de systemd
 * de larga vida. Si el servicio se reinicia, en el peor caso espera un ciclo más.
 */

import * as fs from "fs";
import * as path from "path";
import { cargarConfig, Config } from "./config";
import { calcularRutasDelDia, crearCarpetasDelDia } from "./rutasPlanillaBancos";
import { huellaDelDia } from "./almacenamiento";
import { procesarCarpetaDelDia, imprimirResultado } from "./procesarCarpetaDelDia";

function ahoraTexto(): string {
  return new Date().toISOString().replace("T", " ").slice(0, 19);
}

function log(mensaje: string): void {
  // Salida simple a stdout: systemd la recoge en journald con su propia marca de tiempo.
  console.log(`${ahoraTexto()}  ${mensaje}`);
}

export async function correrPoller(opciones: { config?: Config; ciclos?: number } = {}): Promise<void> {
  const config = opciones.config ?? cargarConfig();
  const maxCiclos = opciones.ciclos; // sin límite si no se especifica: modo servicio

  let huellaAnterior: string | null = null;
  let momentoUltimoCambio: number | null = null;
  let fechaCarpetaActual: string | null = null;
  let ciclo = 0;

  log(
    `Poller iniciado. Revisando cada ${config.intervaloPollSegundos}s, ` +
      `procesa tras ${config.esperaSinCambiosSegundos}s sin cambios.`
  );
  log(`Carpeta base: ${config.rutaCarpetaBase}`);

  for (;;) {
    ciclo++;
    try {
      const rutas = calcularRutasDelDia(new Date(), config);

      // Al cambiar de día se reinicia el seguimiento y se crea la carpeta si falta.
      // Así el proceso se recupera solo aunque el timer de las 00:00 no haya corrido.
      if (fechaCarpetaActual !== rutas.fecha) {
        fechaCarpetaActual = rutas.fecha;
        huellaAnterior = null;
        momentoUltimoCambio = null;
        const creacion =
          config.modoAcceso === "sistemaArchivos"
            ? crearCarpetasDelDia(new Date(), { config })
            : { creadas: [] as string[] };
        if (creacion.creadas.length > 0) {
          log(`Día ${rutas.fecha}: carpetas creadas -> ${rutas.rutaDia}`);
        } else {
          log(`Día ${rutas.fecha}: usando ${rutas.rutaDia}`);
        }
      }

      const huella = await huellaDelDia(new Date(), { config });

      if (huellaAnterior === null) {
        huellaAnterior = huella;
        momentoUltimoCambio = Date.now();
      } else if (huella !== huellaAnterior) {
        log("Cambios detectados en la carpeta; esperando a que se aquiete.");
        huellaAnterior = huella;
        momentoUltimoCambio = Date.now();
      } else if (momentoUltimoCambio !== null) {
        const quietoSegundos = (Date.now() - momentoUltimoCambio) / 1000;
        if (quietoSegundos >= config.esperaSinCambiosSegundos) {
          const resultado = await procesarCarpetaDelDia({ config });
          // Sólo se registra cuando efectivamente pasó algo, para no llenar el journal
          // con "sin procesar" cada minuto durante todo el día.
          if (resultado.procesado) {
            imprimirResultado(resultado);
          } else if (resultado.motivo === "pausado" || resultado.motivo === "archivos-abiertos") {
            log(`Sin procesar: ${resultado.motivo}`);
          }
          // Se marca como atendido: no se vuelve a procesar hasta que algo cambie.
          momentoUltimoCambio = null;
        }
      }
    } catch (err) {
      // Nunca dejar caer el servicio por un error puntual (p. ej. el montaje de red
      // caído un momento): se registra y se reintenta en el próximo ciclo.
      log(`ERROR en el ciclo: ${(err as Error).message}`);
    }

    if (maxCiclos !== undefined && ciclo >= maxCiclos) return;
    await new Promise((r) => setTimeout(r, config.intervaloPollSegundos * 1000));
  }
}

// --- Uso por consola:
//   node poller.js              (modo servicio, no termina)
//   node poller.js --ciclos 3   (para probar a mano)
if (require.main === module) {
  const args = process.argv.slice(2);
  const i = args.indexOf("--ciclos");
  const ciclos = i !== -1 && args[i + 1] ? Number(args[i + 1]) : undefined;

  correrPoller({ ciclos }).catch((err) => {
    console.error("ERROR FATAL:", (err as Error).message);
    process.exit(1);
  });
}
