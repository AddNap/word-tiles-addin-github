# Word Tiles Add-in

Minimalny dodatek Word z "tiles" opartymi o Content Controls.

## Szybki start (DEV)
1. Uruchom lokalny serwer (np. `npx http-server .` albo własny).
2. W Word (Windows/Mac/Web):
   - **Developer** → **Upload Add-in** (lub **My Add-ins** → **Upload**)
   - wskaż `manifest.xml`.
3. Otwórz **Task Pane**, kliknij **Tiles Inspector** → **Sprawdź**.

## Struktura
- `scripts/utils/batch.js` — bezpieczne batchowanie + retry/backoff.
- `scripts/tiles.js` — mapowanie Tile ↔ Content Controls.
- `scripts/inspector.js` — walidator i UI inspektora.

## Wymagania
- `WordApi 1.3+` (zależnie od użytych funkcji). Zob. Requirements dla Office Add-ins. :contentReference[oaicite:5]{index=5}

## Bezpieczeństwo
- Polityka CSP i ładowanie `office.js` wyłącznie po HTTPS. :contentReference[oaicite:6]{index=6}
