/* global console, document, localStorage, setTimeout, Audio, HTMLAudioElement, HTMLButtonElement, HTMLElement, HTMLInputElement, HTMLSelectElement, Office, Word */

import { HelperUnavailableError, synthesize } from "./synth";
import type { SpokenWord, Synthesis } from "./synth";

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

let playButton: HTMLButtonElement;
let pauseButton: HTMLButtonElement;
let stopButton: HTMLButtonElement;
let rateInput: HTMLInputElement;
let volumeInput: HTMLInputElement;
let voiceSelect: HTMLSelectElement;

let audio: HTMLAudioElement | null = null;
let stopped = true;
let nextParagraph = 0; // dónde retomar después de "Detener"

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) return;
  $("sideload-msg").hidden = true;
  $("app-body").hidden = false;

  playButton = $("play");
  pauseButton = $("pause");
  stopButton = $("stop");
  rateInput = $("rate");
  volumeInput = $("volume");
  voiceSelect = $("voice");

  restoreSettings();
  playButton.onclick = play;
  pauseButton.onclick = togglePause;
  stopButton.onclick = stop;
  rateInput.oninput = applySettings;
  volumeInput.oninput = applySettings;
  voiceSelect.onchange = applySettings;
  applySettings();
});

function restoreSettings() {
  try {
    rateInput.value = localStorage.getItem("rate") ?? rateInput.value;
    volumeInput.value = localStorage.getItem("volume") ?? volumeInput.value;
    voiceSelect.value = localStorage.getItem("voice") ?? voiceSelect.value;
  } catch {
    // sin almacenamiento: quedan los valores por defecto
  }
}

function applySettings() {
  const rate = Number(rateInput.value);
  const volume = Number(volumeInput.value);
  $("rate-value").textContent = `${rate.toFixed(1)}×`;
  $("volume-value").textContent = `${Math.round(volume * 100)}%`;
  if (audio) {
    audio.playbackRate = rate;
    audio.volume = volume;
  }
  try {
    localStorage.setItem("rate", rateInput.value);
    localStorage.setItem("volume", volumeInput.value);
    localStorage.setItem("voice", voiceSelect.value);
  } catch {
    // ignorar
  }
}

function setStatus(text: string) {
  $("status").textContent = text;
}

function setPlaying(playing: boolean) {
  playButton.disabled = playing;
  pauseButton.disabled = !playing;
  stopButton.disabled = !playing;
  pauseButton.textContent = "⏸ Pausa";
}

function togglePause() {
  if (!audio) return;
  if (audio.paused) {
    audio.play();
    pauseButton.textContent = "⏸ Pausa";
  } else {
    audio.pause();
    pauseButton.textContent = "▶ Seguir";
  }
}

function stop() {
  stopped = true;
  audio?.pause();
}

async function play() {
  stopped = false;
  setPlaying(true);
  setStatus("Preparando…");
  try {
    const texts = await readParagraphTexts();
    if (nextParagraph >= texts.length) nextParagraph = 0;

    const voice = voiceSelect.value;
    const pending = new Map<number, Promise<Synthesis>>();
    const fetchAudio = (i: number) => {
      if (!pending.has(i)) pending.set(i, synthesize(texts[i], voice));
      return pending.get(i)!;
    };

    for (let i = nextParagraph; i < texts.length && !stopped; i++) {
      nextParagraph = i;
      if (!texts[i].trim()) continue;
      const synthesis = await fetchAudio(i);
      // Pedir el siguiente párrafo mientras suena este, para que no haya cortes.
      const next = texts.findIndex((t, j) => j > i && t.trim() !== "");
      if (next >= 0) fetchAudio(next).catch(() => undefined);
      pending.delete(i);

      setStatus(synthesis.engine === "edge" ? "Leyendo…" : "Leyendo (modo sin conexión)…");
      await playParagraph(i, synthesis);
    }
    if (!stopped) nextParagraph = 0;
    setStatus(stopped ? "Detenido." : "Fin del documento.");
  } catch (error) {
    setStatus(
      error instanceof HelperUnavailableError
        ? "El servicio de voz no está funcionando. Cerrá y volvé a abrir Word."
        : "Ocurrió un error al leer el documento."
    );
    console.error(error);
  } finally {
    audio = null;
    stopped = true;
    setPlaying(false);
  }
}

async function readParagraphTexts(): Promise<string[]> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();
    return paragraphs.items.map((p) => p.text);
  });
}

/** Reproduce un párrafo marcando cada palabra (o el párrafo entero si no hay tiempos). */
async function playParagraph(index: number, synthesis: Synthesis) {
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();
    const paragraph = paragraphs.items[index];

    const ranges = synthesis.words.length
      ? await wordRanges(context, paragraph, synthesis.words)
      : [];
    if (!ranges.length) {
      paragraph.select();
      await context.sync();
    }

    audio = new Audio(`data:${synthesis.mime};base64,${synthesis.audio}`);
    applySettings();
    const ended = new Promise<void>((resolve) => {
      audio!.onended = () => resolve();
      audio!.onerror = () => resolve();
    });
    await audio.play();

    let shown = -1;
    while (!audio.ended && !stopped) {
      const current = currentWord(synthesis.words, audio.currentTime);
      const range = ranges[current];
      if (current !== shown && range) {
        shown = current;
        range.select();
        // Una sync por palabra es inevitable: hay que marcarla en el momento en que suena.
        // eslint-disable-next-line office-addins/no-context-sync-in-loop
        await context.sync();
      }
      await Promise.race([ended, sleep(40)]);
    }
  });
}

/** Busca el Word.Range de cada palabra con una sola sincronización. */
async function wordRanges(
  context: Word.RequestContext,
  paragraph: Word.Paragraph,
  words: SpokenWord[]
): Promise<(Word.Range | undefined)[]> {
  // Los tokens separados por espacio (sin recortar) concatenados reproducen el texto del párrafo,
  // así que su posición acumulada permite ubicar cada palabra por char_start.
  const tokens = paragraph.getTextRanges([" "], false);
  tokens.load("items/text");
  await context.sync();
  // "items/text" sí carga items; la regla no reconoce esa forma.
  // eslint-disable-next-line office-addins/load-object-before-read
  const items = tokens.items;

  const tokenStarts: number[] = [];
  let position = 0;
  for (const token of items) {
    tokenStarts.push(position);
    position += token.text.length;
  }

  const tokenIndexes = words.map((word) => {
    if (word.char_start < 0) return -1;
    let t = tokenStarts.length - 1;
    while (t > 0 && tokenStarts[t] > word.char_start) t--;
    return t;
  });
  const matches = words.map((word, i) =>
    tokenIndexes[i] >= 0
      ? items[tokenIndexes[i]].search(word.text, { matchCase: true }).load("items")
      : null
  );
  await context.sync();

  return words.map((_, i) => {
    const match = matches[i];
    return match ? (match.items[0] ?? items[tokenIndexes[i]]) : undefined;
  });
}

function currentWord(words: SpokenWord[], time: number): number {
  let index = -1;
  for (let i = 0; i < words.length && words[i].start <= time; i++) index = i;
  return index;
}

function sleep(ms: number) {
  return new Promise<void>((resolve) => setTimeout(resolve, ms));
}
