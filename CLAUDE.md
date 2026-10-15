# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Objetivo

Complemento para Microsoft Word que lee el documento en voz alta con voces neuronales y **marca en el documento la palabra que se está leyendo**. Controles: leer/pausa/detener, velocidad, volumen y voz.

**Usuario objetivo: persona con conocimientos básicos de computación que solo se maneja dentro de Word.** Todo debe funcionar desde el panel: nada de consolas ni configuración manual, y los mensajes de error en lenguaje simple. La UI está en español rioplatense.

## Estructura

```
addin/    Office Add-in (task pane, TypeScript, generado con yo office, manifest XML)
helper/   Servicio local Python (FastAPI en http://127.0.0.1:8766): edge-tts con fallback a Piper
```

El helper existe porque edge-tts y Piper son Python y no pueden correr dentro del task pane (WebView2).

## Comandos

### helper/
```bash
python -m venv .venv
.venv/Scripts/python -m pip install -r requirements.txt
.venv/Scripts/python -m piper.download_voices es_AR-daniela-high --download-dir voices
.venv/Scripts/python app.py
```
Probar sin Word: `curl -X POST 127.0.0.1:8766/synthesize -H "Content-Type: application/json" -d '{"text":"Hola","offline":false}'` (`offline: true` fuerza Piper).

### addin/
```bash
npm install
npm start          # dev server HTTPS en localhost:3000 + sideload en Word de escritorio
npm run stop
npm run build
npm run lint       # eslint-plugin-office-addins + prettier; lint:fix para corregir
npm run validate   # valida manifest.xml
npx tsc --noEmit -p . --types office-js --skipLibCheck   # chequeo de tipos (webpack usa babel y NO chequea tipos)
```
No hay tests todavía.

## Cómo funciona

**helper/app.py** — `POST /synthesize {text, voice, piper_voice, offline}` → `{engine, mime, audio (base64), words}`.
- Intenta edge-tts (`boundary="WordBoundary"`, timeout 15 s); ante cualquier error cae a Piper. `engine` le dice a la UI cuál se usó.
- `words`: `{text, start, end, char_start}`. edge-tts da offsets en unidades de 100 ns y solo el texto de la palabra; `locate_words` calcula `char_start` buscando en orden dentro del texto enviado. Con Piper `words` es vacía.
- CORS solo para `https://localhost:3000`, con `allow_private_network=True` (WebView2 manda preflight de Private Network Access al llamar a localhost). Si cambia el origen del add-in, actualizar `allow_origins`.

**addin/src/taskpane/** — `synth.ts` es el cliente del helper; `taskpane.ts` tiene UI y lectura.
- Se sintetiza **por párrafo**; mientras suena uno se pide el siguiente (prefetch) para que no haya cortes.
- Cada párrafo se reproduce dentro de su propio `Word.run`, que queda abierto mientras suena el audio: así los `Word.Range` siguen válidos sin `trackedObjects`.
- El texto enviado al helper es exactamente `paragraph.text`, para que `char_start` coincida con los tokens de `paragraph.getTextRanges([" "], false)` (tokens sin recortar, cuya concatenación reproduce el párrafo). Cada palabra se ubica en su token y se refina con `token.search(word)`, todo en una sola `sync`.
- Marcado con `range.select()` (no `highlightColor`: no modifica el documento ni el historial de deshacer). Con Piper se selecciona el párrafo entero.
- Velocidad = `audio.playbackRate`, volumen = `audio.volume`: se aplican en vivo sin re-sintetizar. Los tiempos de palabra se comparan con `audio.currentTime` (tiempo de medio), así que valen a cualquier velocidad. Preferencias en `localStorage`.
- "Detener" recuerda el párrafo (`nextParagraph`) y "Leer" retoma desde ahí.

### Lint
`eslint-plugin-office-addins` da falsos positivos: `load-object-before-read` no reconoce `load("items/text")` y `no-context-sync-in-loop` salta en el loop de marcado (una sync por palabra es inevitable). Están silenciados con `eslint-disable-next-line` y un comentario; no "arreglarlos" con `load("items")`, que carga todas las propiedades.

## Pendiente

- **Distribución para el usuario final**: empaquetar el helper (p. ej. PyInstaller + instalador que lo arranque con Windows), registrar el manifest sin pasos manuales y servir el task pane por HTTPS fuera del dev server.
- Leer desde la posición del cursor en vez de desde el principio.
- Marcado de palabras también con Piper: `piper-tts` 1.8 puede devolver alineaciones por fonema (`include_alignments=True`) si el modelo se parchea con `python -m piper.patch_voice_with_alignment` (requiere el paquete `onnx`).
