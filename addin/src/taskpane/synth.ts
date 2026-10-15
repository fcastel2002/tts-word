/* global fetch, Response */

// Cliente del servicio local de síntesis (helper/app.py).

export const HELPER_URL = "http://127.0.0.1:8766";

export interface SpokenWord {
  text: string;
  start: number; // segundos dentro del audio
  end: number;
  char_start: number; // posición en el texto enviado; -1 si no se encontró
}

export interface Synthesis {
  engine: "edge" | "piper";
  mime: string;
  audio: string; // base64
  words: SpokenWord[]; // vacía con Piper
}

export class HelperUnavailableError extends Error {}

export async function synthesize(text: string, voice: string): Promise<Synthesis> {
  let response: Response;
  try {
    response = await fetch(`${HELPER_URL}/synthesize`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text, voice }),
    });
  } catch {
    throw new HelperUnavailableError("El servicio de voz no responde");
  }
  if (!response.ok) {
    throw new Error(`Error del servicio de voz (${response.status})`);
  }
  return response.json();
}
