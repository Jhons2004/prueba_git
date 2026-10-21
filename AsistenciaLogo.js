/** Conserva exactamente el logo institucional del sistema central. */
function mdjAsisGetLogoDataUri_() {
  return String(MDJ_LOGO_DATA_URI || '').replace(/\s+/g, '').trim();
}