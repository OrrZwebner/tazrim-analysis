#!/usr/bin/env bash
# render_pdf.sh — print the Hebrew RTL report HTML to an A4 PDF with headless Chrome/Chromium.
#
# Usage:
#   render_pdf.sh <report.html> [out.pdf]     render; out.pdf defaults to <report>.pdf
#   render_pdf.sh --probe                     report which engine WOULD be used, render nothing
#   render_pdf.sh -h | --help
#
# Engine search order (first executable wins):
#   $CHROME_BIN (if set) · google-chrome · google-chrome-stable · chromium · chromium-browser · chrome
#   · /Applications/Google Chrome.app · /Applications/Chromium.app · /Applications/Microsoft Edge.app
# Command: <chrome> --headless=new --disable-gpu --no-pdf-header-footer --print-to-pdf=<out> file://<html>
#
# stdout: exactly ONE JSON object. stderr: progress and warnings.
#   {"ok": true,  "engine": "chrome", "binary": "...", "pdf": "...", "pages": 12|null, "bytes": 123456}
#   {"ok": false, "error": "no-pdf-engine", "engines_tried": [...]}
#   {"ok": false, "error": "input-not-found" | "render-failed", ...}
# Exit codes: 0 ok · 1 input missing · 2 usage · 3 no engine found · 4 engine ran but no PDF written.
# Never claims a PDF exists when none was written. Regenerate the HTML right before printing.
set -uo pipefail

usage() {
  sed -n '2,20p' "$0" | sed 's/^# \{0,1\}//'
}

case "${1:-}" in
  ""|-h|--help) usage; [ -z "${1:-}" ] && exit 2 || exit 0 ;;
esac

ENGINES='"google-chrome","google-chrome-stable","chromium","chromium-browser","chrome","/Applications/Google Chrome.app","/Applications/Chromium.app","/Applications/Microsoft Edge.app"'

find_chrome() {
  local c
  if [ -n "${CHROME_BIN:-}" ] && [ -x "${CHROME_BIN}" ]; then printf '%s' "$CHROME_BIN"; return 0; fi
  for c in google-chrome google-chrome-stable chromium chromium-browser chrome; do
    command -v "$c" >/dev/null 2>&1 && { printf '%s' "$(command -v "$c")"; return 0; }
  done
  for c in "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
           "/Applications/Chromium.app/Contents/MacOS/Chromium" \
           "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"; do
    [ -x "$c" ] && { printf '%s' "$c"; return 0; }
  done
  return 1
}

json_str() { printf '%s' "$1" | sed 's/\\/\\\\/g; s/"/\\"/g'; }

if [ "$1" = "--probe" ]; then
  if BIN=$(find_chrome); then
    printf '{"ok": true, "engine": "chrome", "binary": "%s", "probe": true}\n' "$(json_str "$BIN")"
    echo "probe: chrome — $BIN" >&2
    exit 0
  fi
  printf '{"ok": false, "error": "no-pdf-engine", "engines_tried": [%s], "probe": true}\n' "$ENGINES"
  echo "no Chrome/Chromium found — install Google Chrome or Chromium (or set CHROME_BIN)" >&2
  exit 3
fi

SRC=$1
if [ ! -f "$SRC" ]; then
  printf '{"ok": false, "error": "input-not-found", "input": "%s"}\n' "$(json_str "$SRC")"
  echo "error: no such file: $SRC" >&2
  exit 1
fi
# Absolute paths: Chrome needs a file:// URL; a relative one silently renders a blank page.
SRC_ABS=$(cd "$(dirname "$SRC")" && pwd)/$(basename "$SRC")
OUT=${2:-${SRC_ABS%.*}.pdf}
case "$OUT" in
  /*) OUT_ABS=$OUT ;;
  *)  OUT_ABS=$(cd "$(dirname "$OUT")" && pwd)/$(basename "$OUT") ;;
esac
mkdir -p "$(dirname "$OUT_ABS")"

if ! CHROME=$(find_chrome); then
  printf '{"ok": false, "error": "no-pdf-engine", "engines_tried": [%s], "html": "%s"}\n' "$ENGINES" "$(json_str "$SRC_ABS")"
  cat >&2 <<MSG
error: no Chrome/Chromium was found, so NO PDF WAS PRODUCED.
       The HTML is complete at: $SRC_ABS
       Searched: google-chrome, google-chrome-stable, chromium, chromium-browser, chrome,
                 /Applications/Google Chrome.app, /Applications/Chromium.app, /Applications/Microsoft Edge.app
       Install Google Chrome or Chromium, or set CHROME_BIN=/path/to/chrome, and run again.
MSG
  exit 3
fi

echo "render: $CHROME --headless=new" >&2
rm -f "$OUT_ABS"
# --no-pdf-header-footer: otherwise Chrome stamps the file:// URL and the date into every margin,
# which lands on top of the content in an RTL document.
"$CHROME" --headless=new --disable-gpu --no-pdf-header-footer --no-first-run --no-default-browser-check \
          --print-to-pdf="$OUT_ABS" "file://$SRC_ABS" >/dev/null 2>&1
RC=$?
if [ ! -s "$OUT_ABS" ]; then
  printf '{"ok": false, "error": "render-failed", "engine": "chrome", "binary": "%s", "exit": %s, "html": "%s"}\n' \
         "$(json_str "$CHROME")" "$RC" "$(json_str "$SRC_ABS")"
  echo "error: Chrome exited with $RC and wrote no PDF" >&2
  exit 4
fi

PAGES=""
if command -v pdfinfo >/dev/null 2>&1; then
  PAGES=$(pdfinfo "$OUT_ABS" 2>/dev/null | awk '/^Pages:/ {print $2}')
fi
if [ -z "$PAGES" ]; then
  PAGES=$(LC_ALL=C grep -a -c -E '/Type[[:space:]]*/Page([^s]|$)' "$OUT_ABS" 2>/dev/null || true)
fi
case "$PAGES" in ''|0) PAGES="null" ;; esac
BYTES=$(wc -c < "$OUT_ABS" | tr -d ' ')
printf '{"ok": true, "engine": "chrome", "binary": "%s", "pdf": "%s", "pages": %s, "bytes": %s}\n' \
       "$(json_str "$CHROME")" "$(json_str "$OUT_ABS")" "$PAGES" "$BYTES"
echo "render: wrote $OUT_ABS — pages: $PAGES, $BYTES bytes" >&2
[ "$PAGES" = "null" ] && echo "warning: page count unknown — report it as unknown, do not guess" >&2
exit 0
