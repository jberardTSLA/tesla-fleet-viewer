/* ============================================================
   Tesla Fleet Viewer — Auto-Fetch da ZipLabs Superset
   Usa Puppeteer per aprire Chrome, scaricare i dati, salvare CSV.
   
   SETUP:
     1. cd auto-fetch
     2. npm install
     3. node fetch.js --login-only   (primo login SSO)
     4. node fetch.js                (fetch automatico)
   
   Il browser salva i cookie SSO in ./chrome-profile/
   così non devi rifare il login ogni volta.
   ============================================================ */

const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');

// ─── CONFIGURAZIONE ─────────────────────────────────────────
const CONFIG = {
    // URL delle due dashboard ZipLabs
    dashboard1: {
        url: 'https://ziplabs.teslamotors.com/superset/dashboard/17481/',
        tabName: 'Backlog',
        chartTitle: 'Tutti gli ordini con telaio in Italia',
        outputFile: 'file1_ordini.csv',
    },
    dashboard2: {
        url: 'https://ziplabs.teslamotors.com/superset/dashboard/43372/',
        tabName: '272 IT MATCHED CARS',
        chartTitle: '',  // lascia vuoto se vuoi la prima tabella della sezione
        outputFile: 'file2_enterprise.csv',
    },

    // Dove salvare i CSV
    outputDir: path.join(__dirname, '..'),  // cartella TeslaFleetViewer

    // Chrome profile per mantenere il login SSO
    chromeProfile: path.join(__dirname, 'chrome-profile'),

    // Intervallo di refresh (minuti). 0 = singolo fetch
    refreshMinutes: 5,

    // Timeout per caricamento pagina (ms)
    timeout: 60000,

    // Mostra il browser (false = headless, true = vedi cosa fa)
    headless: false,
};

// ─── MAIN ───────────────────────────────────────────────────
async function main() {
    const loginOnly = process.argv.includes('--login-only');

    console.log('');
    console.log('  ╔══════════════════════════════════════════╗');
    console.log('  ║  Tesla Fleet Viewer — Auto Fetch         ║');
    console.log('  ║  ZipLabs Superset → CSV                  ║');
    console.log('  ╚══════════════════════════════════════════╝');
    console.log('');

    // Lancia Chrome con profilo persistente (mantiene cookies SSO)
    const browser = await puppeteer.launch({
        headless: CONFIG.headless ? 'new' : false,
        userDataDir: CONFIG.chromeProfile,
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--window-size=1400,900',
        ],
        defaultViewport: { width: 1400, height: 900 },
    });

    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(CONFIG.timeout);

    if (loginOnly) {
        console.log('[LOGIN] Apro ZipLabs per il login SSO...');
        console.log('[LOGIN] Fai login nel browser, poi chiudi quando sei dentro.');
        await page.goto(CONFIG.dashboard1.url || 'https://ziplabs.teslamotors.com/', { waitUntil: 'networkidle2' });
        console.log('[LOGIN] In attesa che tu faccia login... (chiudi il browser quando finito)');
        // Il browser resta aperto finché l'utente non lo chiude
        await new Promise(() => {}); // wait forever
    }

    // Fetch loop
    const doFetch = async () => {
        try {
            console.log(`[${new Date().toLocaleTimeString()}] Inizio fetch...`);

            // ── Dashboard 1: Ordini ──
            if (CONFIG.dashboard1.url && CONFIG.dashboard1.url !== 'INCOLLA_QUI_URL_DASHBOARD_BACKLOG') {
                await fetchDashboard(page, CONFIG.dashboard1);
            } else {
                console.log('[SKIP] Dashboard 1: URL non configurato');
            }

            // ── Dashboard 2: Enterprise ──
            if (CONFIG.dashboard2.url) {
                await fetchDashboard(page, CONFIG.dashboard2);
            }

            // ── Carica automaticamente nella Fleet Viewer ──
            await loadIntoFleetViewer(page);

            console.log(`[${new Date().toLocaleTimeString()}] Fetch + caricamento completato!\n`);
        } catch (err) {
            console.error('[ERRORE]', err.message);
        }
    };

    await doFetch();

    if (CONFIG.refreshMinutes > 0) {
        console.log(`[AUTO] Refresh ogni ${CONFIG.refreshMinutes} minuti...`);
        setInterval(doFetch, CONFIG.refreshMinutes * 60 * 1000);
    } else {
        await browser.close();
        console.log('[DONE] Browser chiuso.');
    }
}

// ─── Fetch singola dashboard ────────────────────────────────
async function fetchDashboard(page, dashConfig) {
    console.log(`[FETCH] ${dashConfig.outputFile} da ${dashConfig.url}`);

    // Naviga alla dashboard
    await page.goto(dashConfig.url, { waitUntil: 'networkidle2' });
    await sleep(3000);

    // Se c'è un tab specifico, cliccaci
    if (dashConfig.tabName) {
        try {
            const tabs = await page.$$('.ant-tabs-tab, [role="tab"], .dashboard-component-tabs .tab');
            for (const tab of tabs) {
                const text = await page.evaluate(el => el.textContent.trim(), tab);
                if (text.includes(dashConfig.tabName)) {
                    await tab.click();
                    console.log(`  Tab "${dashConfig.tabName}" cliccato`);
                    await sleep(3000);
                    break;
                }
            }
        } catch (e) {
            console.log(`  Tab "${dashConfig.tabName}" non trovato, continuo...`);
        }
    }

    // Force Refresh di tutte le chart nella pagina
    try {
        // Metodo 1: Bottone "Force refresh" nella toolbar Superset
        const refreshBtn = await page.$('[data-test="refresh-dashboard-menu-item"], [aria-label="Refresh dashboard"], .fa-refresh, [data-test="force-refresh"]');
        if (refreshBtn) {
            await refreshBtn.click();
            console.log('  Force Refresh dashboard cliccato');
            await sleep(8000);
        } else {
            // Metodo 2: Shortcut — apri menu "..." della dashboard e clicca refresh
            const moreMenu = await page.$('.header-with-actions [data-test="more-horiz"], .dashboard-header .anticon-more, [aria-label="More actions"]');
            if (moreMenu) {
                await moreMenu.click();
                await sleep(1000);
                const items = await page.$$('.ant-dropdown-menu-item, [role="menuitem"]');
                for (const item of items) {
                    const text = await page.evaluate(el => el.textContent.trim().toLowerCase(), item);
                    if (text.includes('force refresh') || text.includes('refresh')) {
                        await item.click();
                        console.log('  Force Refresh via menu cliccato');
                        await sleep(8000);
                        break;
                    }
                }
            }
        }
        // Metodo 3: Keyboard shortcut (Superset supporta Ctrl+Shift+L per refresh)
        // await page.keyboard.down('Control');
        // await page.keyboard.down('Shift');
        // await page.keyboard.press('KeyL');
        // await page.keyboard.up('Shift');
        // await page.keyboard.up('Control');
    } catch (e) {
        console.log('  Force Refresh non trovato, i dati potrebbero essere cached');
    }

    // Trova la tabella/chart e esporta CSV
    // Metodo 1: Cerca il menu "..." della chart e clicca "Export to CSV"
    const exported = await tryExportCSV(page, dashConfig);

    if (exported) {
        console.log(`  ✓ Salvato: ${dashConfig.outputFile}`);
    } else {
        // Metodo 2: Scrape diretto dalla tabella DOM
        console.log('  Export CSV non trovato, provo scraping DOM...');
        await scrapeTableToCSV(page, dashConfig);
    }
}

// ─── Export CSV via menu Superset ────────────────────────────
async function tryExportCSV(page, dashConfig) {
    try {
        // Trova tutti i menu "..." delle chart
        const menuButtons = await page.$$('.chart-container .dot, .slice_container [data-test="more-horiz"], .chart-header .anticon-more, [aria-label="More Options"]');

        for (const btn of menuButtons) {
            // Se cerchiamo una chart specifica, controlla il titolo
            if (dashConfig.chartTitle) {
                const container = await page.evaluateHandle(el => el.closest('.chart-container, .slice_container, [data-test="chart-container"]'), btn);
                if (container) {
                    const title = await page.evaluate(el => {
                        const h = el.querySelector('.header-title, .slice_header span, h2, h3');
                        return h ? h.textContent.trim() : '';
                    }, container);
                    if (!title.includes(dashConfig.chartTitle)) continue;
                }
            }

            // Clicca il menu
            await btn.click();
            await sleep(1000);

            // Cerca "Export to .CSV" o "Download as CSV"
            const menuItems = await page.$$('.ant-dropdown-menu-item, [role="menuitem"], .dropdown-menu li');
            for (const item of menuItems) {
                const text = await page.evaluate(el => el.textContent.trim().toLowerCase(), item);
                if (text.includes('csv') || text.includes('export')) {
                    // Setup download handler
                    const downloadPath = path.join(CONFIG.outputDir, dashConfig.outputFile);

                    // Listen for download
                    const client = await page.target().createCDPSession();
                    await client.send('Page.setDownloadBehavior', {
                        behavior: 'allow',
                        downloadPath: CONFIG.outputDir,
                    });

                    await item.click();
                    console.log('  Export CSV cliccato, attendo download...');
                    await sleep(5000);

                    // Rinomina il file scaricato (Superset lo chiama con nomi generici)
                    const files = fs.readdirSync(CONFIG.outputDir)
                        .filter(f => f.endsWith('.csv') && f !== dashConfig.outputFile)
                        .map(f => ({ name: f, time: fs.statSync(path.join(CONFIG.outputDir, f)).mtimeMs }))
                        .sort((a, b) => b.time - a.time);

                    if (files.length > 0) {
                        const latest = files[0].name;
                        const src = path.join(CONFIG.outputDir, latest);
                        const dst = path.join(CONFIG.outputDir, dashConfig.outputFile);
                        if (fs.existsSync(dst)) fs.unlinkSync(dst);
                        fs.renameSync(src, dst);
                    }

                    return true;
                }
            }

            // Chiudi menu se non trovato
            await page.keyboard.press('Escape');
        }
    } catch (e) {
        console.log('  Export via menu fallito:', e.message);
    }
    return false;
}

// ─── Scrape tabella DOM → CSV ───────────────────────────────
async function scrapeTableToCSV(page, dashConfig) {
    try {
        const csvData = await page.evaluate(() => {
            // Trova la tabella più grande nella pagina
            const tables = document.querySelectorAll('table');
            let bestTable = null;
            let bestRows = 0;

            tables.forEach(t => {
                const rows = t.querySelectorAll('tbody tr');
                if (rows.length > bestRows) {
                    bestRows = rows.length;
                    bestTable = t;
                }
            });

            if (!bestTable) return null;

            const headers = [];
            bestTable.querySelectorAll('thead th, thead td').forEach(th => {
                headers.push(th.textContent.trim());
            });

            const rows = [];
            bestTable.querySelectorAll('tbody tr').forEach(tr => {
                const cells = [];
                tr.querySelectorAll('td').forEach(td => {
                    cells.push('"' + td.textContent.trim().replace(/"/g, '""') + '"');
                });
                if (cells.length > 0) rows.push(cells.join(','));
            });

            return '\uFEFF' + headers.map(h => '"' + h + '"').join(',') + '\n' + rows.join('\n');
        });

        if (csvData) {
            const filePath = path.join(CONFIG.outputDir, dashConfig.outputFile);
            fs.writeFileSync(filePath, csvData, 'utf8');
            console.log(`  ✓ Scraped ${csvData.split('\n').length - 1} righe → ${dashConfig.outputFile}`);
        } else {
            console.log('  ✗ Nessuna tabella trovata nella pagina');
        }
    } catch (e) {
        console.log('  Scraping fallito:', e.message);
    }
}

// ─── Utils ──────────────────────────────────────────────────
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ─── Carica CSV nella Fleet Viewer ──────────────────────────
async function loadIntoFleetViewer(page) {
    const fleetViewerUrl = 'https://jberardtsla.github.io/tesla-fleet-viewer/';
    const file1Path = path.resolve(CONFIG.outputDir, CONFIG.dashboard1.outputFile);
    const file2Path = path.resolve(CONFIG.outputDir, CONFIG.dashboard2.outputFile);

    const file1Exists = fs.existsSync(file1Path);
    const file2Exists = fs.existsSync(file2Path);

    if (!file1Exists) {
        console.log('[FLEET] File 1 non trovato, skip caricamento');
        return;
    }

    console.log('[FLEET] Apro Fleet Viewer...');
    await page.goto(fleetViewerUrl, { waitUntil: 'networkidle2' });
    await sleep(3000);

    // Carica File 1 (Ordini) nello slot 1
    try {
        const fileInput1 = await page.$('#bootFile1');
        if (fileInput1) {
            await fileInput1.uploadFile(file1Path);
            console.log('[FLEET] ✓ File 1 caricato: ' + CONFIG.dashboard1.outputFile);
            await sleep(2000);
        } else {
            console.log('[FLEET] ✗ Input #bootFile1 non trovato');
            return;
        }
    } catch (e) {
        console.log('[FLEET] Errore caricamento File 1:', e.message);
        return;
    }

    // Carica File 2 (Enterprise) nello slot 2
    if (file2Exists) {
        try {
            const fileInput2 = await page.$('#bootFile2');
            if (fileInput2) {
                await fileInput2.uploadFile(file2Path);
                console.log('[FLEET] ✓ File 2 caricato: ' + CONFIG.dashboard2.outputFile);
                await sleep(2000);
            }
        } catch (e) {
            console.log('[FLEET] Errore caricamento File 2:', e.message);
        }
    }

    // Clicca "INIZIALIZZARE SISTEMA"
    await sleep(1000);
    try {
        const startBtn = await page.$('#bootStartBtn');
        if (startBtn) {
            const isDisabled = await page.evaluate(btn => btn.disabled, startBtn);
            if (!isDisabled) {
                await startBtn.click();
                console.log('[FLEET] ✓ INIZIALIZZARE SISTEMA cliccato');
                await sleep(3000);

                // Skip intro (clicca SALTA INTRO)
                try {
                    const skipBtn = await page.$('.cod-skip');
                    if (skipBtn) {
                        await skipBtn.click();
                        console.log('[FLEET] ✓ Intro saltata');
                    }
                } catch (e) {}

                console.log('[FLEET] ✓ Dashboard caricata con successo!');
            } else {
                console.log('[FLEET] ✗ Bottone ancora disabilitato — file non riconosciuto?');
            }
        }
    } catch (e) {
        console.log('[FLEET] Errore avvio:', e.message);
    }
}

// ─── Start ──────────────────────────────────────────────────
main().catch(console.error);
