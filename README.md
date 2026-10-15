# TTS Word

A Microsoft Word add-in that reads your document aloud with natural neural voices and **selects each word as it is spoken**, so you can follow along in the document itself.

It replaces Word's built-in Read Aloud with much more fluent voices, and it's built for people who are comfortable in Word but not much beyond it: one installer, then everything happens from a panel inside Word.

## Features

- **Natural voices**: Microsoft neural voices through [edge-tts](https://github.com/rany2/edge-tts). They're free and don't need an account or API key.
- **Word-by-word tracking**: the word being spoken is selected in the document, in sync with the audio.
- **Works offline too**: without internet it falls back to [Piper](https://github.com/OHF-Voice/piper1-gpl), a local voice. In that mode the whole paragraph is selected instead of each word.
- **Controls**: play, pause/resume, stop (resumes from the same paragraph), speed (0.5×–2×) and volume, applied live while reading.
- **Spanish voices** from Argentina, Spain and Mexico. The UI is in Spanish.
- Follows Word's light or dark theme.

## Install

1. Download `TTSWord-Setup.exe` from the [latest release](../../releases/latest).
2. Run it and accept the administrator prompt.
3. Open Word. On the **Home** tab, click **Leer en voz alta** ("Read aloud").

Requires Word for Windows (Microsoft 365 or Office 2016 or later).

> The installer isn't code-signed yet, so Windows SmartScreen may show "Windows protected your PC". Click **More info → Run anyway**.

To uninstall, use **Settings → Apps → Installed apps → TTS Word**.

## How it works

```
Word ── task pane (addin/, TypeScript + Office.js)
           │  POST /synthesize, one paragraph at a time
           ▼
        local helper (helper/, Python + FastAPI on localhost:8766)
           ├─ edge-tts → audio + per-word timings
           └─ Piper    → audio only (offline fallback)
```

- The **add-in** reads the document paragraph by paragraph. While one paragraph plays, it already requests the next one so there are no gaps. It compares the audio position with the word timings and selects the matching word with `range.select()`, which doesn't modify the document or its undo history. Speed and volume are applied to the audio element directly, so changing them doesn't re-synthesize anything.
- The **helper** exists because edge-tts and Piper are Python libraries and can't run inside Word's web-based task pane. Once installed, it starts with Windows and also serves the task pane over HTTPS. It uses a certificate from a local certificate authority that is generated on each machine.
- The **installer** (`installer/`) packages the helper with PyInstaller and builds the setup with Inno Setup. It installs the helper, sets up the certificate, starts it at sign-in and registers the add-in in Word.

## Development

Requirements: Windows, Word for Windows, Node.js 24+ and Python 3.14.

**Helper**

```powershell
cd helper
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python -m piper.download_voices es_AR-daniela-high --download-dir voices
.venv\Scripts\python app.py
```

**Add-in** (in another terminal)

```powershell
cd addin
npm install
npm start      # HTTPS dev server on localhost:3000 + sideloads the add-in into Word
npm run stop   # when you're done
```

Checks:

```powershell
npm run lint
npx tsc --noEmit -p . --types office-js --skipLibCheck   # webpack uses Babel and does not type-check
```

If TTS Word is installed on the same machine, close the installed helper first, because it uses the same port: `taskkill /IM tts-helper.exe /F`.

**Building the installer**

Also requires Inno Setup 6 (`winget install JRSoftware.InnoSetup`) and `pip install -r helper\requirements-build.txt` in the helper venv. Then run:

```powershell
.\installer\build.ps1   # → installer\Output\TTSWord-Setup.exe
```

## Releases

Releases are built by GitHub Actions ([`release.yml`](.github/workflows/release.yml)). To publish one:

1. Bump `<Version>` in `addin/manifest.xml` (for example `1.0.3.0`). It is also the installer version.
2. Commit, then tag and push using the first three parts of the version:

   ```bash
   git tag v1.0.3
   git push origin v1.0.3
   ```

The workflow checks that the tag matches the manifest, builds `TTSWord-Setup.exe` and attaches it to a new GitHub release.

## Contributing

Contributions are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

The source code in this repository is under the [MIT License](LICENSE).

The installer bundles third-party components that have their own licenses: Piper (`piper-tts`, GPL-3.0-or-later), edge-tts (LGPL-3.0) and the Piper voice model. Because the packaged helper includes Piper, **`TTSWord-Setup.exe` as a whole is distributed under GPL-3.0-or-later**. Its complete source is this repository.

edge-tts uses Microsoft's online Read Aloud service unofficially, so it may change or stop working at any time. If that happens, TTS Word falls back to the offline voice.
