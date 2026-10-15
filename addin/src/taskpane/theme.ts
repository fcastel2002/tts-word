/* global document, Office */

// Adapta los colores del panel al tema de Office (Colorido, Gris oscuro, Negro, Blanco).
// Si Office no informa el tema, quedan los colores de taskpane.css (que siguen al modo de Windows).

function normalize(color: string): string {
  return color.startsWith("#") ? color : `#${color}`;
}

function isDark(color: string): boolean {
  const n = parseInt(normalize(color).slice(1, 7), 16);
  const luminance = 0.299 * (n >> 16) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255);
  return luminance < 128;
}

function apply(theme: Office.OfficeTheme): void {
  const bodyBg = normalize(theme.bodyBackgroundColor);
  const controlBg = normalize(theme.controlBackgroundColor);
  const controlFg = normalize(theme.controlForegroundColor);
  const dark = isDark(bodyBg);
  const style = document.documentElement.style;
  style.setProperty("--bg", bodyBg);
  style.setProperty("--fg", normalize(theme.bodyForegroundColor));
  style.setProperty("--control-bg", controlBg);
  style.setProperty("--control-fg", controlFg);
  style.setProperty("--border", `color-mix(in srgb, ${controlFg} 45%, ${controlBg})`);
  // Azul de Word; en oscuro uno más claro con texto negro para que se lea.
  style.setProperty("--accent", dark ? "#4f8fe6" : "#2b579a");
  style.setProperty("--accent-fg", dark ? "#000000" : "#ffffff");
  style.colorScheme = dark ? "dark" : "light"; // sliders, select y scrollbars nativos
}

export function followOfficeTheme(): void {
  const theme = Office.context.officeTheme;
  if (!theme) return;
  apply(theme);
  try {
    // Cambiar el tema con el panel abierto; no todas las versiones de Word lo avisan.
    Office.context.document.addHandlerAsync(
      Office.EventType.OfficeThemeChanged,
      (event: Office.OfficeThemeChangedEventArgs) => apply(event.officeTheme)
    );
  } catch {
    // sin el evento, el tema se toma de nuevo la próxima vez que se abre el panel
  }
}
