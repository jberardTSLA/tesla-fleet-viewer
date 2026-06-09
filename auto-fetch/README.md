# Auto-Fetch ZipLabs → Tesla Fleet Viewer

Script che scarica automaticamente i dati da ZipLabs e li carica nella Fleet Viewer ogni 5 minuti.

## Requisiti
- **Node.js** installato ([scarica qui](https://nodejs.org/))
- Accesso a **ZipLabs Superset** con SSO aziendale

## Setup (5 minuti, una volta sola)

### 1. Scarica lo script
```bash
git clone https://github.com/jberardTSLA/tesla-fleet-viewer.git
cd tesla-fleet-viewer/auto-fetch
npm install
```

Oppure scarica solo la cartella `auto-fetch/` e mettila dove vuoi.

### 2. Primo login SSO
```bash
node fetch.js --login-only
```
Si apre Chrome → fai login SSO su ZipLabs → **chiudi il browser** quando sei dentro.
I cookie vengono salvati in `chrome-profile/` e riusati automaticamente.

### 3. Avvia il fetch automatico
```bash
node fetch.js
```

Lo script:
1. Apre ZipLabs dashboard 17481 (Backlog) → Force Refresh → Export CSV
2. Apre ZipLabs dashboard 43372 (Enterprise) → Force Refresh → Export CSV
3. Apre Fleet Viewer → carica i 2 file → clicca "Inizializzare Sistema"
4. **Ripete ogni 5 minuti**

## Configurazione

Modifica `CONFIG` in `fetch.js`:

| Parametro | Default | Descrizione |
|---|---|---|
| `refreshMinutes` | 5 | Intervallo refresh (0 = singolo fetch) |
| `headless` | false | true = browser nascosto |
| `dashboard1.url` | 17481 | URL dashboard ordini |
| `dashboard2.url` | 43372 | URL dashboard enterprise |

## Problemi comuni

| Problema | Soluzione |
|---|---|
| "SSO scaduto" | Rifai `node fetch.js --login-only` |
| "Tabella non trovata" | Verifica che il tab sia corretto in CONFIG |
| "Chrome non si apre" | Verifica che Node.js sia installato |
