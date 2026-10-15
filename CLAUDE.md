# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Objetivo

Complemento para Microsoft Word que lee el documento en voz alta con voces neuronales y **marca en el documento la palabra que se está leyendo**. Controles: leer/pausa/detener, velocidad, volumen y voz.

**Usuario objetivo: persona con conocimientos básicos de computación que solo se maneja dentro de Word.** Todo debe funcionar desde el panel: nada de consolas ni configuración manual, y los mensajes de error en lenguaje simple. La UI está en español rioplatense.

## Estructura

```
addin/    Office Add-in (task pane, TypeScript, generado con yo office, manifest XML)
helper/      Servicio local Python (FastAPI en 127.0.0.1:8766): edge-tts con fallback a Piper
installer/   PyInstaller + Inno Setup → TTSWord-Setup.exe para el usuario final
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
- CORS solo para `https://localhost:3000`, con `allow_private_network=True` (WebView2 manda preflight de Private Network Access al llamar a localhost). Solo hace falta en desarrollo: instalado, el panel y la API comparten origen.
- Si existe `web/` (addin/dist empaquetado) lo sirve como estáticos en `/`; las rutas de datos (`voices/`, `web/`) salen de `sys._MEIPASS` cuando está congelado.

**helper/service.py** — entrada del ejecutable instalado (`app.py` sigue siendo la de desarrollo). Sin argumentos: escribe la clave `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` (valor = `<Id>` del manifest, dato = ruta a `web/manifest.xml`), sale si el 8766 ya está ocupado y si no levanta uvicorn **con HTTPS** usando el certificado de `%PROGRAMDATA%\TTS Word\`. Sin consola, el log va a `%LOCALAPPDATA%\TTS Word\helper.log`. `--generate-cert` crea una CA propia de la máquina (`TTS Word - CA local`), firma el certificado de `localhost`/`127.0.0.1` y descarta la clave de la CA. `--unregister` borra la clave de Word.

`HELPER_URL` en `synth.ts` lo define webpack (`DefinePlugin`): `http://127.0.0.1:8766` en modo development y `""` (mismo origen) en el build de producción, cuyo manifest apunta a `https://localhost:8766/` (`urlProd`).

**addin/src/taskpane/** — `synth.ts` es el cliente del helper; `taskpane.ts` tiene UI y lectura.
- Se sintetiza **por párrafo**; mientras suena uno se pide el siguiente (prefetch) para que no haya cortes.
- Cada párrafo se reproduce dentro de su propio `Word.run`, que queda abierto mientras suena el audio: así los `Word.Range` siguen válidos sin `trackedObjects`.
- El texto enviado al helper es exactamente `paragraph.text`, para que `char_start` coincida con los tokens de `paragraph.getTextRanges([" "], false)` (tokens sin recortar, cuya concatenación reproduce el párrafo). Cada palabra se ubica en su token y se refina con `token.search(word)`, todo en una sola `sync`.
- Marcado con `range.select()` (no `highlightColor`: no modifica el documento ni el historial de deshacer). Con Piper se selecciona el párrafo entero.
- Velocidad = `audio.playbackRate`, volumen = `audio.volume`: se aplican en vivo sin re-sintetizar. Los tiempos de palabra se comparan con `audio.currentTime` (tiempo de medio), así que valen a cualquier velocidad. Preferencias en `localStorage`.
- "Detener" recuerda el párrafo (`nextParagraph`) y "Leer" retoma desde ahí.

### Lint
`eslint-plugin-office-addins` da falsos positivos: `load-object-before-read` no reconoce `load("items/text")` y `no-context-sync-in-loop` salta en el loop de marcado (una sync por palabra es inevitable). Están silenciados con `eslint-disable-next-line` y un comentario; no "arreglarlos" con `load("items")`, que carga todas las propiedades.

## Generar el instalador

Requisitos, una sola vez: el venv de `helper/` con `pip install -r requirements-build.txt`, la voz de Piper en `helper/voices/`, `npm install` en `addin/` e Inno Setup 6 (`winget install JRSoftware.InnoSetup`). Después:

```powershell
.\installer\build.ps1     # npm run build → PyInstaller (tts-helper.spec) → ISCC (tts-word.iss)
```

El logo (`addin/assets/*.png` e `installer/tts-word.ico`) se dibuja con `python installer/make_icons.py` (necesita Pillow); los PNG e ICO generados están en el repo, así que solo hace falta correrlo si cambia el diseño.

Deja `installer/Output/TTSWord-Setup.exe` (~165 MB). Para publicar una versión nueva, subir `<Version>` en `addin/manifest.xml`: es también la versión del instalador.

**Releases en GitHub** (`fcastel2002/tts-word`): `.github/workflows/release.yml` corre `build.ps1` en `windows-latest` al pushear un tag `vX.Y.Z` y adjunta el `.exe` a un release. El tag tiene que coincidir con las tres primeras partes de `<Version>` (1.0.3.0 → `v1.0.3`) o el workflow falla. `ci.yml` corre lint, tipos y build del add-in en cada PR.

Qué hace el instalador (pide permisos de administrador una vez):
1. Copia el helper a `C:\Program Files\TTS Word\` (cierra antes uno que esté corriendo).
2. Corre `tts-helper.exe --generate-cert` y agrega la CA a `LocalMachine\Root` con `certutil` (antes borra una CA anterior con el mismo nombre).
3. Agrega `tts-helper.exe` a `HKLM\...\Run` (arranca al iniciar sesión cualquier usuario) y lo lanza como el usuario original; al arrancar registra el complemento en Word. Word muestra "Leer en voz alta" en Inicio la próxima vez que se abre.

El desinstalador hace `--unregister`, borra la CA, la clave Run, `%PROGRAMDATA%\TTS Word` y los archivos.

Por qué así:
- **Registro**: se usa la clave de desarrollador de WEF (la misma que usa `npm start`) porque es la única opción sin pasos dentro de Word y sin admin. El catálogo de carpeta compartida obliga a agregarlo a mano desde Insertar → Mis complementos, y la implementación centralizada requiere un tenant de Microsoft 365.
- **Hosting**: lo sirve el propio helper (y no un hosting público) para que el panel cargue sin internet, cuando se usa Piper. Al compartir origen no hace falta CORS.
- **Certificado**: la CA se genera en cada PC y su clave no se guarda, así que no hay una CA compartida que se pueda filtrar. Va a `LocalMachine\Root` porque en `CurrentUser\Root` Windows muestra un aviso de seguridad.

Para desarrollar con el helper instalado corriendo hay que cerrarlo antes (`taskkill /IM tts-helper.exe /F`): ocupa el 8766.

## Pendiente

- Firmar `TTSWord-Setup.exe` y `tts-helper.exe`: sin firma, SmartScreen muestra "Windows protegió su PC".
- `office.js` se carga del CDN de Microsoft: sin internet el panel depende de que esté en la caché de WebView2.
- Leer desde la posición del cursor en vez de desde el principio.
- Marcado de palabras también con Piper: `piper-tts` 1.8 puede devolver alineaciones por fonema (`include_alignments=True`) si el modelo se parchea con `python -m piper.patch_voice_with_alignment` (requiere el paquete `onnx`).
