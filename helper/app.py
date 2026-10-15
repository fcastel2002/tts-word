"""Servicio local de síntesis para el complemento TTS Word.

Usa edge-tts (voces neuronales, requiere internet) y cae a Piper (local) si falla.
"""

import asyncio
import base64
import io
import logging
import os
import wave
from functools import lru_cache
from pathlib import Path

import edge_tts
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from piper import PiperVoice
from pydantic import BaseModel

log = logging.getLogger("tts-helper")

VOICES_DIR = Path(os.environ.get("TTS_PIPER_VOICES_DIR", Path(__file__).parent / "voices"))
DEFAULT_EDGE_VOICE = "es-AR-ElenaNeural"
DEFAULT_PIPER_VOICE = "es_AR-daniela-high"
EDGE_TIMEOUT_S = 15
TICKS_PER_SECOND = 10_000_000  # edge-tts usa unidades de 100 ns

app = FastAPI(title="TTS Word helper")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://localhost:3000"],
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
    # Chromium/WebView2 pide este permiso (Private Network Access) al llamar a localhost.
    allow_private_network=True,
)


class SynthesisRequest(BaseModel):
    text: str
    voice: str = DEFAULT_EDGE_VOICE
    piper_voice: str = DEFAULT_PIPER_VOICE
    offline: bool = False  # forzar Piper


class Word(BaseModel):
    text: str
    start: float  # segundos dentro del audio
    end: float
    char_start: int  # posición en el texto enviado; -1 si no se encontró


class SynthesisResponse(BaseModel):
    engine: str  # "edge" | "piper"
    mime: str
    audio: str  # base64
    words: list[Word]  # vacía con Piper


def locate_words(text: str, words: list[tuple[str, float, float]]) -> list[Word]:
    """Ubica cada palabra en el texto recorriendo en orden."""
    result, cursor = [], 0
    for word, start, end in words:
        pos = text.find(word, cursor)
        if pos >= 0:
            cursor = pos + len(word)
        result.append(Word(text=word, start=start, end=end, char_start=pos))
    return result


async def synthesize_edge(text: str, voice: str) -> SynthesisResponse:
    communicate = edge_tts.Communicate(text, voice, boundary="WordBoundary")
    audio = bytearray()
    boundaries = []
    async for chunk in communicate.stream():
        if chunk["type"] == "audio":
            audio.extend(chunk["data"])
        elif chunk["type"] == "WordBoundary":
            start = chunk["offset"] / TICKS_PER_SECOND
            boundaries.append((chunk["text"], start, start + chunk["duration"] / TICKS_PER_SECOND))
    return SynthesisResponse(
        engine="edge",
        mime="audio/mpeg",
        audio=base64.b64encode(audio).decode(),
        words=locate_words(text, boundaries),
    )


@lru_cache(maxsize=2)
def load_piper(name: str) -> PiperVoice:
    model = VOICES_DIR / f"{name}.onnx"
    if not model.exists():
        raise FileNotFoundError(f"No está la voz de Piper {model}")
    return PiperVoice.load(str(model))


def synthesize_piper(text: str, voice_name: str) -> SynthesisResponse:
    buffer = io.BytesIO()
    with wave.open(buffer, "wb") as wav:
        load_piper(voice_name).synthesize_wav(text, wav)
    return SynthesisResponse(
        engine="piper",
        mime="audio/wav",
        audio=base64.b64encode(buffer.getvalue()).decode(),
        words=[],
    )


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/synthesize", response_model=SynthesisResponse)
async def synthesize(req: SynthesisRequest):
    if not req.text.strip():
        raise HTTPException(400, "Texto vacío")
    if not req.offline:
        try:
            return await asyncio.wait_for(synthesize_edge(req.text, req.voice), EDGE_TIMEOUT_S)
        except Exception as exc:  # sin internet, timeout, servicio caído
            log.warning("edge-tts falló, uso Piper: %r", exc)
    try:
        return await asyncio.to_thread(synthesize_piper, req.text, req.piper_voice)
    except FileNotFoundError as exc:
        raise HTTPException(503, str(exc))


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=int(os.environ.get("TTS_HELPER_PORT", 8766)))
