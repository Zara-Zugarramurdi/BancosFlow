#!/usr/bin/env bash
#
# verificar-entorno.sh
#
# Comprueba que todo lo que BancosFlow necesita esté en su lugar, ANTES de procesar.
# Sólo lee: no modifica configuración, ni sube nada, ni toca los servicios.
#
# Uso:
#   bash scripts/verificar-entorno.sh
#
# Existe porque diagnosticar esto a mano llevó horas más de una vez: el token de rclone
# vencido, el remoto mal escrito o la falta de permiso de escritura dan errores muy
# distintos y ninguno evidente. Acá sale en un minuto y con un mensaje claro.

set -uo pipefail

PROYECTO="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$PROYECTO" || exit 1

OK=0
FALLOS=0

titulo() { printf "\n\033[1m== %s ==\033[0m\n" "$1"; }
bien()   { printf "  \033[32mOK\033[0m    %s\n" "$1"; OK=$((OK+1)); }
mal()    { printf "  \033[31mFALLA\033[0m %s\n" "$1"; FALLOS=$((FALLOS+1)); }
aviso()  { printf "  \033[33mAVISO\033[0m %s\n" "$1"; }

# ---------------------------------------------------------------- herramientas
titulo "Herramientas"

if command -v node >/dev/null 2>&1; then
  bien "node $(node -v)"
else
  mal "node no está instalado o no está en el PATH"
fi

# El binario puede estar configurado con ruta completa (rcloneBinario), así que la
# comprobación real se hace más abajo, una vez leída la configuración.
RCLONE=rclone

if [ -d dist ] && [ -f dist/procesarCarpetaDelDia.js ]; then
  if [ "$(find src -name '*.ts' -newer dist/procesarCarpetaDelDia.js 2>/dev/null | head -1)" ]; then
    aviso "hay archivos en src/ más nuevos que dist/ — falta compilar: npx tsc"
  else
    bien "dist/ compilado y al día"
  fi
else
  mal "falta dist/ — compilar con: npx tsc"
fi

# ---------------------------------------------------------------- configuración
titulo "Configuración"

CONFIG_JSON="$(node dist/config.js 2>&1)"
if [ $? -ne 0 ]; then
  mal "la configuración no carga:"
  echo "$CONFIG_JSON" | sed 's/^/        /'
  echo
  printf "\033[31mNo se puede seguir sin una configuración válida.\033[0m\n"
  exit 1
fi
bien "la configuración carga y es válida"

leer() { echo "$CONFIG_JSON" | sed -n "s/.*\"$1\": *\"\([^\"]*\)\".*/\1/p" | head -1; }

MODO="$(leer modoAcceso)"
REMOTO="$(leer remotoRclone)"
BASE_REMOTA="$(leer rutaRemotaBase)"
PLANILLA_REMOTA="$(leer rutaRemotaPlanillaMaestra)"
TRABAJO="$(leer rutaTrabajoLocal)"
BACKUPS="$(leer rutaBackups)"
RCLONE="$(leer rcloneBinario)"; RCLONE="${RCLONE:-rclone}"

if command -v "$RCLONE" >/dev/null 2>&1; then
  bien "rclone $("$RCLONE" version 2>/dev/null | head -1 | awk '{print $2}') ($RCLONE)"
else
  mal "no se encuentra el binario de rclone: '$RCLONE'"
  mal "instalar con: curl https://rclone.org/install.sh | sudo bash"
fi

printf "  modoAcceso = %s\n" "$MODO"

if [ "$MODO" != "rclone" ]; then
  aviso "modoAcceso NO es 'rclone'. Con SharePoint debería serlo; si se dejó en"
  aviso "'sistemaArchivos' el proceso trabaja contra una carpeta montada."
fi

# ---------------------------------------------------------------- disco local
titulo "Carpetas locales"

for d in "$TRABAJO" "$BACKUPS"; do
  [ -z "$d" ] && continue
  if [ -d "$d" ]; then
    if [ -w "$d" ]; then bien "$d (escribible)"; else mal "$d existe pero no es escribible"; fi
  else
    aviso "$d no existe todavía (se crea sola en la primera corrida)"
  fi
done

LIBRE_KB="$(df -Pk "$PROYECTO" | awk 'NR==2 {print $4}')"
LIBRE_GB=$((LIBRE_KB / 1024 / 1024))
if [ "$LIBRE_GB" -lt 3 ]; then
  mal "quedan ${LIBRE_GB} GB libres: poco margen para respaldos"
else
  bien "espacio libre: ${LIBRE_GB} GB"
fi

# ---------------------------------------------------------------- remoto
if [ "$MODO" = "rclone" ]; then
  titulo "Conexión con SharePoint"

  NOMBRE_REMOTO="${REMOTO%:}"
  if "$RCLONE" listremotes 2>/dev/null | grep -q "^${NOMBRE_REMOTO}:$"; then
    bien "el remoto '${NOMBRE_REMOTO}:' está configurado en rclone"
  else
    mal "el remoto '${NOMBRE_REMOTO}:' NO existe. Ver: "$RCLONE" listremotes / rclone config"
  fi

  if SALIDA="$("$RCLONE" lsd "${NOMBRE_REMOTO}:" 2>&1)"; then
    bien "el remoto responde (token vigente)"
  else
    mal "el remoto no responde. Suele ser el token vencido; rehacer con: $RCLONE config reconnect ${NOMBRE_REMOTO}:"
    echo "$SALIDA" | tail -3 | sed 's/^/        /'
  fi

  if SALIDA="$("$RCLONE" lsjson "${NOMBRE_REMOTO}:${BASE_REMOTA}" 2>&1)"; then
    bien "se ve la carpeta base: ${BASE_REMOTA}"
  else
    mal "no se puede listar ${BASE_REMOTA}"
    echo "$SALIDA" | tail -3 | sed 's/^/        /'
  fi

  # El grep tolera espacios alrededor de los dos puntos: distintas versiones de rclone
  # formatean el JSON distinto y un chequeo literal daba falsos negativos.
  if "$RCLONE" lsjson "${NOMBRE_REMOTO}:$(dirname "$PLANILLA_REMOTA")" 2>/dev/null \
      | grep -qE "\"Name\"[[:space:]]*:[[:space:]]*\"$(basename "$PLANILLA_REMOTA")\""; then
    bien "la planilla maestra existe: $(basename "$PLANILLA_REMOTA")"
  else
    mal "NO se encuentra la planilla maestra en ${PLANILLA_REMOTA}"
    aviso "ojo con la extensión doble (PlanillaBancos.xlsx.xlsx): Windows la oculta"
  fi

  # Prueba de escritura real: es lo único que confirma que el permiso alcanza.
  titulo "Permiso de escritura en el remoto"
  TMP="$(mktemp /tmp/bancosflow-prueba-XXXXXX.txt)"
  echo "prueba de escritura $(date)" > "$TMP"
  if "$RCLONE" copyto "$TMP" "${NOMBRE_REMOTO}:${BASE_REMOTA}/.bancosflow-prueba.txt" >/dev/null 2>&1; then
    bien "se puede escribir en el remoto"
    "$RCLONE" deletefile "${NOMBRE_REMOTO}:${BASE_REMOTA}/.bancosflow-prueba.txt" >/dev/null 2>&1 \
      && bien "y borrar (archivo de prueba eliminado)" \
      || aviso "quedó .bancosflow-prueba.txt en el remoto: borrar a mano"
  else
    mal "NO se puede escribir en el remoto"
  fi
  rm -f "$TMP"
fi

# ---------------------------------------------------------------- servicios
titulo "Servicios"

for u in bancosflow-poller.service bancosflow-carpetas.timer; do
  if systemctl list-unit-files "$u" >/dev/null 2>&1 && systemctl cat "$u" >/dev/null 2>&1; then
    ESTADO="$(systemctl is-active "$u" 2>/dev/null)"
    HABIL="$(systemctl is-enabled "$u" 2>/dev/null)"
    bien "$u: $ESTADO / $HABIL"
  else
    aviso "$u no está instalado"
  fi
done

# ---------------------------------------------------------------- restos
titulo "Restos de configuraciones anteriores"

if mount | grep -qi sharepoint; then
  aviso "hay algo montado en sharepoint: con modoAcceso=rclone no hace falta y confunde"
else
  bien "no hay montajes de sharepoint"
fi

if pgrep -f "rclone mount" >/dev/null 2>&1; then
  aviso "hay un proceso 'rclone mount' corriendo: con modoAcceso=rclone no hace falta"
else
  bien "no hay procesos de rclone mount"
fi

# ---------------------------------------------------------------- resumen
printf "\n\033[1m== Resumen ==\033[0m\n"
printf "  %s verificaciones OK, %s fallas\n" "$OK" "$FALLOS"
if [ "$FALLOS" -eq 0 ]; then
  printf "\n\033[32mEntorno listo.\033[0m Probar con: node dist/procesarCarpetaDelDia.js --dry-run\n"
  exit 0
else
  printf "\n\033[31mHay %s problema(s) a resolver antes de procesar.\033[0m\n" "$FALLOS"
  exit 1
fi
