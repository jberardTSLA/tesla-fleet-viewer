# Auto-Fetch ZipLabs → Tesla Fleet Viewer

## Setup

```bash
cd auto-fetch
npm install
```

## Primo Login (una volta sola)

```bash
node fetch.js --login-only
```

Si apre Chrome → fai login SSO su ZipLabs → chiudi il browser.
I cookie vengono salvati in `chrome-profile/` e riusati.

## Fetch Dati

```bash
node fetch.js
```

Scarica i CSV dalle 2 dashboard e li salva nella cartella Fleet Viewer.

## Configurazione

Modifica `CONFIG` in `fetch.js`:
- `dashboard1.url` → URL della dashboard Backlog
- `dashboard2.url` → URL della dashboard Enterprise (già configurato: 43372)
- `refreshMinutes` → 0 per singolo fetch, 30 per ogni 30 minuti
- `headless` → false per vedere il browser, true per nasconderlo
