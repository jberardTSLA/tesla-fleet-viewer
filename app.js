/* ============================================================
   Tesla Fleet Viewer - Main Application Logic
   ============================================================ */

// ─── Global State ───────────────────────────────────────────
let rawData = [];              // All parsed rows
let filteredData = [];         // After filters
let allHeaders = [];           // Original Excel column names
let columnMap = {};            // Mapped columns
let charts = {};               // Chart.js instances
let currentPage = 1;
const pageSize = 25;
let sortCol = 2;               // default sort by date
let sortAsc = true;
let readyPage = 1;
let readyPageSize = 50;
let paymentPage = 1;
let paymentPageSize = 50;

// Field keys we need
const FIELDS = [
    { key: 'orderId',           label: 'Order ID / VIN' },
    { key: 'reservationNumber', label: 'Reservation Number' },
    { key: 'wdoCheckoutLink',   label: 'WDO Checkout Link' },
    { key: 'location',          label: 'Hub / Location' },
    { key: 'date',              label: 'Data Arrivo' },
    { key: 'deliveryDate',      label: 'Data Consegna' },
    { key: 'model',             label: 'Modello' },
    { key: 'status',            label: 'Stato' },
];

// Colors palette for charts
const CHART_COLORS = [
    '#3b82f6', '#06b6d4', '#22c55e', '#eab308', '#f97316',
    '#a855f7', '#ec4899', '#14b8a6', '#6366f1', '#f43f5e',
    '#84cc16', '#0ea5e9', '#d946ef', '#fb923c', '#2dd4bf',
];

// ─── BOOT SCREEN (PS2 Style) ────────────────────────────────
let bootFile1Loaded = false;
let bootFile2Loaded = false;

// Generate floating particles
(function() {
    const container = document.getElementById('bootParticles');
    if (!container) return;
    for (let i = 0; i < 30; i++) {
        const p = document.createElement('div');
        p.className = 'boot-particle';
        p.style.left = Math.random() * 100 + '%';
        p.style.animationDuration = (4 + Math.random() * 6) + 's';
        p.style.animationDelay = Math.random() * 8 + 's';
        p.style.opacity = (0.2 + Math.random() * 0.5);
        p.style.width = p.style.height = (1 + Math.random() * 3) + 'px';
        const colors = ['#3b82f6', '#06b6d4', '#22c55e', '#a855f7'];
        p.style.background = colors[Math.floor(Math.random() * colors.length)];
        container.appendChild(p);
    }
})();

// Boot file handlers
document.addEventListener('DOMContentLoaded', () => {
    const bf1 = document.getElementById('bootFile1');
    const bf2 = document.getElementById('bootFile2');
    if (bf1) bf1.addEventListener('change', e => { if (e.target.files.length) handleBootFile(e.target.files[0], 1); });
    if (bf2) bf2.addEventListener('change', e => { if (e.target.files.length) handleBootFile(e.target.files[0], 2); });
});

function handleBootFile(file, slot) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const firstSheet = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheet];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

            if (jsonData.length === 0) {
                alert('Il file non contiene dati.');
                return;
            }

            if (slot === 1) {
                // Main file
                allHeaders = Object.keys(jsonData[0]);
                rawData = jsonData;
                const autoMapped = tryAutoMap(allHeaders);
                if (autoMapped) {
                    columnMap = autoMapped;
                } else {
                    columnMap = {};
                    // Will show mapping after boot
                }
                bootFile1Loaded = true;
                const slotEl = document.getElementById('bootSlot1');
                slotEl.classList.add('loaded');
                document.getElementById('bootDesc1').textContent = jsonData.length + ' ordini caricati';
                document.getElementById('bootAction1').innerHTML = '<svg width="20" height="20" viewBox="0 0 20 20" fill="none"><path d="M6 10l3 3 5-5" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg><span>CARICATO</span>';
                if (window.gsapFileLoaded) gsapFileLoaded(slotEl);
            } else {
                // Secondary file — store for later merge
                window._bootFile2Data = jsonData;
                bootFile2Loaded = true;
                const slotEl = document.getElementById('bootSlot2');
                slotEl.classList.add('loaded');
                document.getElementById('bootDesc2').textContent = jsonData.length + ' righe caricate';
                document.getElementById('bootAction2').innerHTML = '<svg width="20" height="20" viewBox="0 0 20 20" fill="none"><path d="M6 10l3 3 5-5" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg><span>CARICATO</span>';
                if (window.gsapFileLoaded) gsapFileLoaded(slotEl);
                // Unlock Bravo
                OPTIMUS_DATA[1].locked = false;
                const lockBadge = document.getElementById('charLock1');
                if (lockBadge) lockBadge.style.display = 'none';
            }

            updateBootButton();
        } catch(err) {
            alert('Errore nel parsing: ' + err.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

function updateBootButton() {
    const btn = document.getElementById('bootStartBtn');
    const text = document.getElementById('bootInsertText');
    if (bootFile1Loaded) {
        btn.disabled = false;
        btn.querySelector('.boot-start-text').textContent = 'INIZIALIZZARE SISTEMA';
        text.textContent = 'Sistema pronto. Premi per continuare.';
        text.style.color = '#22c55e';
    }
}

function bootStart() {
    if (!bootFile1Loaded) return;

    // Process File 1
    processAndRender();

    // Process File 2 if loaded
    if (bootFile2Loaded && window._bootFile2Data) {
        rawData_backup = rawData;
        handleFile2Merge(window._bootFile2Data);
    }

    // Transition: boot → optimus selection (GSAP if available)
    if (window.gsapBootToOptimus) {
        gsapBootToOptimus(() => {
            document.getElementById('bootScreen').style.display = 'none';
            document.getElementById('optimusIntro').style.display = 'flex';
            if (window.gsapOptimus) gsapOptimus();
        });
    } else {
        const boot = document.getElementById('bootScreen');
        boot.classList.add('booting');
        setTimeout(() => {
            boot.style.display = 'none';
            document.getElementById('optimusIntro').style.display = 'flex';
        }, 1200);
    }
}

function handleFile2Merge(jsonData) {
    const file2Headers = Object.keys(jsonData[0]);
    const rnCol = file2Headers.find(h => h.toLowerCase().replace(/[_\s]/g,'') === 'referencenumber') || null;
    if (!rnCol) return;

    const lookup = {};
    jsonData.forEach(row => {
        const rn = String(row[rnCol] || '').trim();
        if (rn) lookup[rn] = row;
    });

    const findCol = (names) => {
        for (const n of names) {
            const found = file2Headers.find(h => h.toLowerCase().replace(/[_\s]/g,'') === n.toLowerCase().replace(/[_\s]/g,''));
            if (found) return found;
        }
        return null;
    };

    const col_eta2sc = findCol(['ETA2SC']);
    const col_enterprise = findCol(['IsEnterpriseOrder']);
    const col_confidentEta = findCol(['IsConfidentETA']);
    const col_scheduled = findCol(['IsScheduled']);
    const col_paymentStatus = findCol(['FinalPaymentStatus']);
    const col_containment = findCol(['IsContainmenthold']);
    const col_lastLocation = findCol(['LastKnownVehicleLocation']);

    rawData.forEach(row => {
        const rn = row.reservationNumber;
        if (!rn || !lookup[rn]) return;
        const f2 = lookup[rn];

        if (col_eta2sc) {
            const eta = parseAnyDate(f2[col_eta2sc]);
            if (eta) {
                row.date = eta;
                row.dateStr = eta.toLocaleDateString('it-IT');
                const refDate = row.deliveryDate || row.date;
                const today = new Date(); today.setHours(0,0,0,0);
                const refClean = new Date(refDate); refClean.setHours(0,0,0,0);
                row.daysUntil = Math.round((refClean - today) / 86400000);
                if (row.daysUntil < 0) row.urgency = 'past';
                else if (row.daysUntil === 0) row.urgency = 'today';
                else if (row.daysUntil <= 2) row.urgency = 'imminent';
                else if (row.daysUntil <= 5) row.urgency = 'soon';
                else row.urgency = 'none';
            }
        }
        if (col_enterprise) { const v = String(f2[col_enterprise]||'').trim().toLowerCase(); row.isEnterprise = (v==='true'||v==='1'||v==='yes'); row.orderChannel = row.isEnterprise ? 'B2B' : 'B2C'; }
        if (col_confidentEta) { const v = String(f2[col_confidentEta]||'').trim().toLowerCase(); row.isConfidentETA = (v==='true'||v==='1'||v==='yes'); }
        if (col_scheduled) { const v = String(f2[col_scheduled]||'').trim().toLowerCase(); row.isScheduled = (v==='true'||v==='1'||v==='yes'); }
        if (col_paymentStatus) { row.finalPaymentStatus = String(f2[col_paymentStatus]||'').trim(); }
        if (col_containment) { const v = String(f2[col_containment]||'').trim().toLowerCase(); row.isContainmentHold = (v==='true'||v==='1'||v==='yes'); }
        if (col_lastLocation) { row.lastKnownLocation = String(f2[col_lastLocation]||'').trim(); }
    });

    applyFilters();
}

// ─── OPTIMUS INTRO (CoD Style) ──────────────────────────────
const OPTIMUS_DATA = [
    { id: 'arrivals', name: 'OPTIMUS ALPHA', classType: 'LOGISTICS COMMAND', color: '#3b82f6',
      skill1: 'MONITOR ARRIVI', desc1: 'Traccia tutti gli arrivi in tempo reale sugli hub Italia',
      skill2: 'TIMELINE VIEW', desc2: 'Visualizza la timeline dei prossimi 30 giorni con volumi',
      requires: 'File 1', locked: false },
    { id: 'ready', name: 'OPTIMUS BRAVO', classType: 'DELIVERY OPS', color: '#22c55e',
      skill1: 'CONSEGNE PRONTE', desc1: 'Identifica veicoli a terra con pagamento completo e KPI 6 giorni',
      skill2: 'GROUND CONTROL', desc2: 'Monitoraggio giorni a terra, scheduling e scadenze',
      requires: 'File 2', locked: true },
    { id: 'alerts', name: 'OPTIMUS CHARLIE', classType: 'THREAT DETECTION', color: '#ef4444',
      skill1: 'CONTAINMENT SCAN', desc1: 'Rileva veicoli con CH attivo e blocchi critici',
      skill2: 'DELAY ALERT', desc2: 'Segnala ordini da posticipare e consegne in conflitto',
      requires: 'File 1', locked: false },
    { id: 'analytics', name: 'OPTIMUS DELTA', classType: 'INTEL ANALYSIS', color: '#a855f7',
      skill1: 'PIVOT MATRIX', desc1: 'Analisi volumi per Location e Modello con totali',
      skill2: 'FLEET INTEL', desc2: 'Grafici B2B/B2C, distribuzione modelli, stato ordini',
      requires: 'File 1', locked: false },
];

let selectedRoster = 0;

function selectRoster(idx) {
    selectedRoster = idx;
    const d = OPTIMUS_DATA[idx];

    // Update character highlights
    document.querySelectorAll('.cod-char').forEach((c, i) => {
        c.classList.toggle('active', i === idx);
    });

    // GSAP switch animation
    if (window.gsapSwitchChar) gsapSwitchChar(idx);

    // Update info bar
    document.getElementById('codInfoName').textContent = d.name;
    document.getElementById('codInfoDesc').textContent = d.classType;
    document.getElementById('codMiniSkill1').textContent = d.skill1;
    document.getElementById('codMiniSkill2').textContent = d.skill2;

    // Deploy button state
    const btn = document.getElementById('codDeployBtn');
    const req = document.getElementById('codReq');
    if (d.locked) {
        btn.classList.add('locked');
        req.textContent = 'RICHIEDE ' + d.requires + ' — Carica i dati per attivare questo operativo';
    } else {
        btn.classList.remove('locked');
        req.textContent = '';
    }
}

function deployOptimus() {
    const d = OPTIMUS_DATA[selectedRoster];
    if (d.locked) return;
    selectOptimus(d.id);
}

function selectOptimus(section) {
    const intro = document.getElementById('optimusIntro');

    const afterClose = () => {
        intro.style.display = 'none';
        // Dashboard reveal with GSAP
        if (window.gsapDashboardReveal) gsapDashboardReveal();
        if (document.getElementById('dashboard').style.display !== 'none') {
            let targetId = null;
            switch(section) {
                case 'arrivals': targetId = 'arrivalsTable'; break;
                case 'ready': targetId = 'readySection'; break;
                case 'alerts': targetId = 'postponeSection'; break;
                case 'analytics': targetId = 'chartByLocation'; break;
            }
            if (targetId) {
                const el = document.getElementById(targetId);
                if (el) setTimeout(() => el.scrollIntoView({ behavior: 'smooth', block: 'start' }), 600);
            }
        }
    };

    if (window.gsapDeploy) {
        gsapDeploy(afterClose);
    } else {
        intro.classList.add('closing');
        setTimeout(afterClose, 800);
    }
}

function skipIntro() {
    const intro = document.getElementById('optimusIntro');
    intro.classList.add('closing');
    setTimeout(() => { intro.style.display = 'none'; }, 800);
}

// Init first selection
setTimeout(() => selectRoster(0), 100);

// ─── Initialize ─────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    // Set current date
    const now = new Date();
    document.getElementById('currentDate').textContent =
        now.toLocaleDateString('it-IT', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });

    // Set default filter dates
    const today = now.toISOString().split('T')[0];
    // Don't set default date filter - show all data
    // document.getElementById('filterDateFrom').value = today;

    // File input listener
    document.getElementById('fileInput').addEventListener('change', handleFileSelect);

    // Second file input listener
    document.getElementById('fileInput2').addEventListener('change', handleFile2Select);

    // Drag & drop
    const uploadCard = document.querySelector('.upload-card');
    uploadCard.addEventListener('dragover', e => {
        e.preventDefault();
        uploadCard.classList.add('drag-over');
    });
    uploadCard.addEventListener('dragleave', () => {
        uploadCard.classList.remove('drag-over');
    });
    uploadCard.addEventListener('drop', e => {
        e.preventDefault();
        uploadCard.classList.remove('drag-over');
        if (e.dataTransfer.files.length) {
            handleFile(e.dataTransfer.files[0]);
        }
    });
});

// ─── File Handling ──────────────────────────────────────────
function handleFileSelect(e) {
    if (e.target.files.length) {
        handleFile(e.target.files[0]);
    }
}

function handleFile(file) {
    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const firstSheet = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheet];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

            if (jsonData.length === 0) {
                alert('Il file non contiene dati.');
                return;
            }

            allHeaders = Object.keys(jsonData[0]);
            rawData = jsonData;

            // Try auto-mapping
            const autoMapped = tryAutoMap(allHeaders);
            if (autoMapped) {
                columnMap = autoMapped;
                processAndRender();
            } else {
                showColumnMapping();
            }
        } catch (err) {
            alert('Errore nel parsing del file: ' + err.message);
            console.error(err);
        }
    };
    reader.readAsArrayBuffer(file);
}

// ─── Second File Handling (Merge) ───────────────────────────
function handleFile2Select(e) {
    if (e.target.files.length) {
        handleFile2(e.target.files[0]);
    }
}

function handleFile2(file) {
    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const firstSheet = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheet];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

            if (jsonData.length === 0) {
                alert('Il secondo file non contiene dati.');
                return;
            }

            // Find ReferenceNumber column in file 2
            const file2Headers = Object.keys(jsonData[0]);
            const rnCol = file2Headers.find(h => h.toLowerCase().replace(/[_\s]/g,'') === 'referencenumber') || null;

            if (!rnCol) {
                alert('Colonna "ReferenceNumber" non trovata nel secondo file. Impossibile incrociare i dati.');
                return;
            }

            // Build lookup by ReferenceNumber
            const lookup = {};
            jsonData.forEach(row => {
                const rn = String(row[rnCol] || '').trim();
                if (rn) lookup[rn] = row;
            });

            // Find columns in file 2
            const findCol = (names) => {
                for (const n of names) {
                    const found = file2Headers.find(h => h.toLowerCase().replace(/[_\s]/g,'') === n.toLowerCase().replace(/[_\s]/g,''));
                    if (found) return found;
                }
                return null;
            };

            const col_eta2sc = findCol(['ETA2SC']);
            const col_enterprise = findCol(['IsEnterpriseOrder']);
            const col_confidentEta = findCol(['IsConfidentETA']);
            const col_scheduled = findCol(['IsScheduled']);
            const col_paymentStatus = findCol(['FinalPaymentStatus']);
            const col_containment = findCol(['IsContainmenthold']);
            const col_lastLocation = findCol(['LastKnownVehicleLocation']);

            // Merge into rawData
            let matched = 0;
            rawData.forEach(row => {
                const rn = row.reservationNumber;
                if (!rn || !lookup[rn]) return;
                const file2Row = lookup[rn];
                matched++;

                // ETA2SC overrides arrival date
                if (col_eta2sc) {
                    const eta2sc = parseAnyDate(file2Row[col_eta2sc]);
                    if (eta2sc) {
                        row.date = eta2sc;
                        row.dateStr = eta2sc.toLocaleDateString('it-IT');
                        // Recalculate urgency
                        const refDate = row.deliveryDate || row.date;
                        const today = new Date(); today.setHours(0,0,0,0);
                        const refClean = new Date(refDate); refClean.setHours(0,0,0,0);
                        row.daysUntil = Math.round((refClean - today) / 86400000);
                        if (row.daysUntil < 0) row.urgency = 'past';
                        else if (row.daysUntil === 0) row.urgency = 'today';
                        else if (row.daysUntil <= 2) row.urgency = 'imminent';
                        else if (row.daysUntil <= 5) row.urgency = 'soon';
                        else row.urgency = 'none';
                    }
                }

                // IsEnterpriseOrder
                if (col_enterprise) {
                    const val = String(file2Row[col_enterprise] || '').trim().toLowerCase();
                    row.isEnterprise = (val === 'true' || val === '1' || val === 'yes');
                    row.orderChannel = row.isEnterprise ? 'B2B' : 'B2C';
                }

                // IsConfidentETA
                if (col_confidentEta) {
                    const val = String(file2Row[col_confidentEta] || '').trim().toLowerCase();
                    row.isConfidentETA = (val === 'true' || val === '1' || val === 'yes');
                }

                // IsScheduled
                if (col_scheduled) {
                    const val = String(file2Row[col_scheduled] || '').trim().toLowerCase();
                    row.isScheduled = (val === 'true' || val === '1' || val === 'yes');
                }

                // FinalPaymentStatus
                if (col_paymentStatus) {
                    row.finalPaymentStatus = String(file2Row[col_paymentStatus] || '').trim();
                }

                // IsContainmenthold
                if (col_containment) {
                    const val = String(file2Row[col_containment] || '').trim().toLowerCase();
                    row.isContainmentHold = (val === 'true' || val === '1' || val === 'yes');
                }

                // LastKnownVehicleLocation
                if (col_lastLocation) {
                    row.lastKnownLocation = String(file2Row[col_lastLocation] || '').trim();
                }
            });

            // Update status
            const statusEl = document.getElementById('dataStatus');
            statusEl.innerHTML = '<span class="status-dot"></span><span>' + rawData.length + ' ordini | File 2 integrato (' + matched + ' match)</span>';

            // Update button
            const btn = document.getElementById('btnFile2');
            btn.style.background = 'rgba(34,197,94,0.1)';
            btn.style.color = '#22c55e';
            btn.style.borderColor = 'rgba(34,197,94,0.2)';
            btn.innerHTML = '<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><path d="M4 8L7 11L12 5" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg> File 2 OK (' + matched + ')';

            // Unlock BRAVO in roster
            OPTIMUS_DATA[1].locked = false;
            const lockBadge = document.getElementById('charLock1');
            if (lockBadge) lockBadge.style.display = 'none';

            // Re-render
            applyFilters();
            alert('File 2 integrato! ' + matched + ' ordini incrociati su ' + jsonData.length + ' righe.');

        } catch (err) {
            alert('Errore nel parsing del secondo file: ' + err.message);
            console.error(err);
        }
    };
    reader.readAsArrayBuffer(file);
}

// ─── Auto-mapping heuristics ────────────────────────────────
function tryAutoMap(headers) {
    const map = {};
    const lowerHeaders = headers.map(h => h.toString().toLowerCase().trim());

    const patterns = {
        orderId:          ['vin', 'order id', 'orderid', 'order_id', 'id ordine', 'ordine', 'vin number', 'vehicle identification', 'telaio', 'chassis'],
        reservationNumber:['referencenumber', 'reference number', 'reference_number', 'reservation number', 'reservation_number', 'reservationnumber', 'reservation #', 'prenotazione', 'numero prenotazione', 'numero reservation'],
        wdoCheckoutLink:  ['wdocheckoutlink', 'wdo checkout link', 'wdo_checkout_link', 'wdo link', 'wdo_link', 'checkout link', 'checkout_link', 'wdo checkout', 'wdo url'],
        location:         ['servicecenterforpickup', 'service center for pick up', 'service_center_for_pickup', 'location', 'hub', 'destinazione', 'sede', 'centro', 'magazzino', 'warehouse', 'site', 'plant', 'stabilimento', 'porto', 'port', 'compound', 'delivery location', 'delivery_location', 'service center', 'sc location', 'storelocation', 'store_location'],
        deliverySpecialist: ['deliveryspecialist', 'delivery_specialist', 'delivery specialist', 'specialist', 'sales advisor', 'salesadvisor', 'advisor'],
        date:             ['vehicleetatoservice', 'vehicle eta to service', 'vehicleeta', 'carrier_eta', 'eta2servicecenter', 'data arrivo', 'arrival date', 'arrival', 'eta', 'data_arrivo', 'arrivo', 'estimated arrival', 'estimated_arrival', 'planned arrival', 'eta date', 'carrier eta', 'scheduled date'],
        deliveryDate:     ['scheduleddeliverydate', 'scheduled delivery date', 'scheduled_delivery_date', 'data consegna', 'delivery date', 'delivery_date', 'consegna', 'data di consegna', 'scheduled delivery', 'data delivery', 'planned delivery', 'customer delivery date', 'handover date', 'handover'],
        model:            ['model', 'modello', 'producttrim', 'product_trim', 'productname', 'veicolo', 'vehicle', 'tipo', 'type', 'prodotto', 'product', 'model name', 'variant', 'trim', 'vehicle model', 'vehicle_model', 'model variant', 'modeltrimname'],
        status:           ['vehiclestatus', 'vehicle_status', 'vehicle status', 'orderstatus', 'order_status', 'order status', 'registrationstatus', 'registration_status', 'stato', 'status', 'state', 'situazione', 'condizione', 'delivery status', 'transport status'],
    };

    // Pass 1: exact match (highest priority)
    for (const [key, terms] of Object.entries(patterns)) {
        if (map[key]) continue;
        for (let i = 0; i < lowerHeaders.length; i++) {
            if (Object.values(map).includes(headers[i])) continue; // already used
            if (terms.some(t => lowerHeaders[i] === t)) {
                map[key] = headers[i];
                break;
            }
        }
    }

    // Pass 2: includes match (fallback for partial matches)
    for (const [key, terms] of Object.entries(patterns)) {
        if (map[key]) continue;
        for (let i = 0; i < lowerHeaders.length; i++) {
            if (Object.values(map).includes(headers[i])) continue; // already used
            if (terms.some(t => lowerHeaders[i].includes(t))) {
                map[key] = headers[i];
                break;
            }
        }
    }

    // Check if we got at least location and date
    console.log('[Fleet Viewer] Auto-map result:', JSON.stringify(map));
    console.log('[Fleet Viewer] Excel headers:', headers.join(', '));
    if (!map.location) console.warn('[Fleet Viewer] MISSING: location column not found!');
    if (!map.date) console.warn('[Fleet Viewer] MISSING: date column not found!');
    return (map.location && map.date) ? map : null;
}

// ─── Column Mapping UI ─────────────────────────────────────
function showColumnMapping() {
    const mappingDiv = document.getElementById('columnMapping');
    const grid = document.getElementById('mappingGrid');
    grid.innerHTML = '';

    FIELDS.forEach(field => {
        const item = document.createElement('div');
        item.className = 'mapping-item';

        const label = document.createElement('label');
        label.textContent = field.label;
        item.appendChild(label);

        const select = document.createElement('select');
        select.id = 'map_' + field.key;

        const optNone = document.createElement('option');
        optNone.value = '';
        optNone.textContent = '-- Non mappato --';
        select.appendChild(optNone);

        allHeaders.forEach(h => {
            const opt = document.createElement('option');
            opt.value = h;
            opt.textContent = h;
            select.appendChild(opt);
        });

        // Pre-select if auto-map partial
        const autoMap = tryAutoMap(allHeaders) || {};
        if (autoMap[field.key]) {
            select.value = autoMap[field.key];
        }

        item.appendChild(select);
        grid.appendChild(item);
    });

    mappingDiv.style.display = 'block';
}

function applyMapping() {
    columnMap = {};
    FIELDS.forEach(field => {
        const val = document.getElementById('map_' + field.key).value;
        if (val) columnMap[field.key] = val;
    });

    if (!columnMap.location || !columnMap.date) {
        alert('Per favore mappa almeno "Hub / Location" e "Data Arrivo".');
        return;
    }

    processAndRender();
}

// ─── Data Processing ────────────────────────────────────────
function processAndRender() {
    // Normalize data
    const processed = rawData.map((row, idx) => {
        const orderId  = columnMap.orderId  ? String(row[columnMap.orderId] || '').trim()  : 'ORD-' + String(idx + 1).padStart(5, '0');
        const location = columnMap.location ? String(row[columnMap.location] || '').trim() : 'N/A';
        const model    = columnMap.model    ? String(row[columnMap.model] || '').trim()     : 'N/A';
        const status   = columnMap.status   ? String(row[columnMap.status] || '').trim()    : 'Programmato';
        const deliverySpecialist = columnMap.deliverySpecialist ? String(row[columnMap.deliverySpecialist] || '').trim() : '';

        // Reservation Number & WDO Link
        // In ZipLabs, WDOCheckoutLink column contains the RN code (e.g. RN127765407), not a URL
        let reservationNumber = columnMap.reservationNumber ? String(row[columnMap.reservationNumber] || '').trim() : '';
        const wdoColValue = columnMap.wdoCheckoutLink ? String(row[columnMap.wdoCheckoutLink] || '').trim() : '';

        // If WDOCheckoutLink value looks like an RN code, use it as reservation number
        if (!reservationNumber && wdoColValue && /^RN\d+$/i.test(wdoColValue)) {
            reservationNumber = wdoColValue;
        }

        // Build DRO Advisor link with RN
        const wdoCheckoutLink = reservationNumber ? 'https://dro.tesla.com/advisor?sidepanel_fullscreen=yes&rn=' + encodeURIComponent(reservationNumber) : '';

        // Parse arrival date
        let dateObj = parseAnyDate(row[columnMap.date]);

        // Parse delivery date
        let deliveryDateObj = columnMap.deliveryDate ? parseAnyDate(row[columnMap.deliveryDate]) : null;

        // Calculate urgency based on delivery date (or arrival date as fallback)
        const refDate = deliveryDateObj || dateObj;
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        let daysUntil = null;
        let urgency = 'none'; // none | imminent | today | soon
        if (refDate) {
            const refClean = new Date(refDate);
            refClean.setHours(0, 0, 0, 0);
            daysUntil = Math.round((refClean - today) / 86400000);
            if (daysUntil < 0)       urgency = 'past';
            else if (daysUntil === 0) urgency = 'today';
            else if (daysUntil <= 2)  urgency = 'imminent';  // 1-2 days -> flashing orange
            else if (daysUntil <= 5)  urgency = 'soon';      // 3-5 days -> soft warning
        }

        return {
            orderId,
            reservationNumber,
            wdoCheckoutLink,
            location,
            date: dateObj,
            dateStr: dateObj ? dateObj.toLocaleDateString('it-IT') : 'N/A',
            deliveryDate: deliveryDateObj,
            deliveryDateStr: deliveryDateObj ? deliveryDateObj.toLocaleDateString('it-IT') : '—',
            model,
            deliverySpecialist,
            status,
            daysUntil,
            urgency,
        };
    }).filter(r => r.date !== null); // Remove rows with no valid date

    // Safety check: if ALL rows got filtered out, keep them anyway with today's date
    if (processed.length === 0 && rawData.length > 0) {
        console.warn('[Fleet Viewer] ATTENZIONE: nessuna data valida trovata! Usando la data odierna come fallback per ' + rawData.length + ' righe.');
        const fallbackDate = new Date();
        const fallbackProcessed = rawData.map((row, idx) => {
            const orderId  = columnMap.orderId  ? String(row[columnMap.orderId] || '').trim()  : 'ORD-' + String(idx + 1).padStart(5, '0');
            const location = columnMap.location ? String(row[columnMap.location] || '').trim() : 'N/A';
            const model    = columnMap.model    ? String(row[columnMap.model] || '').trim()     : 'N/A';
            const status   = columnMap.status   ? String(row[columnMap.status] || '').trim()    : 'N/A';
            const deliverySpecialist = columnMap.deliverySpecialist ? String(row[columnMap.deliverySpecialist] || '').trim() : '';
            let reservationNumber = columnMap.reservationNumber ? String(row[columnMap.reservationNumber] || '').trim() : '';
            const wdoColValue = columnMap.wdoCheckoutLink ? String(row[columnMap.wdoCheckoutLink] || '').trim() : '';
            if (!reservationNumber && wdoColValue && /^RN\d+$/i.test(wdoColValue)) reservationNumber = wdoColValue;
            const wdoCheckoutLink = reservationNumber ? 'https://dro.tesla.com/advisor?sidepanel_fullscreen=yes&rn=' + encodeURIComponent(reservationNumber) : '';
            // Try to get SOME date from the raw value
            const rawDateVal = row[columnMap.date];
            const rawDateStr = rawDateVal ? String(rawDateVal) : '';
            return {
                orderId, reservationNumber, wdoCheckoutLink, location,
                date: fallbackDate,
                dateStr: rawDateStr || fallbackDate.toLocaleDateString('it-IT'),
                deliveryDate: null, deliveryDateStr: '—',
                model, deliverySpecialist, status,
                daysUntil: 0, urgency: 'none',
            };
        });
        rawData = fallbackProcessed;
        alert('Attenzione: le date nel file non sono state riconosciute.\n\nFormato colonna data: "' + (rawData[0]?.dateStr || '?') + '"\n\nI dati sono comunque caricati con data odierna.\nColonne mappate: ' + Object.keys(columnMap).join(', '));
    } else {
        rawData = processed;
    }
    applyFilters();

    // Show dashboard
    document.getElementById('dashboard').style.display = 'block';

    // Update header status
    const statusEl = document.getElementById('dataStatus');
    statusEl.className = 'header-status active';
    statusEl.innerHTML = '<span class="status-dot"></span><span>' + rawData.length + ' ordini caricati</span>';
}

function parseAnyDate(dateVal) {
    if (!dateVal) return null;
    if (dateVal instanceof Date) return isNaN(dateVal.getTime()) ? null : dateVal;
    if (typeof dateVal === 'number') {
        // Excel serial date (number > 40000 is likely a serial date)
        if (dateVal > 40000 && dateVal < 70000) {
            return new Date((dateVal - 25569) * 86400 * 1000);
        }
        // Unix timestamp in seconds
        if (dateVal > 1000000000 && dateVal < 2000000000) {
            return new Date(dateVal * 1000);
        }
        // Unix timestamp in milliseconds
        if (dateVal > 1000000000000) {
            return new Date(dateVal);
        }
        // Fallback: try Excel serial anyway
        return new Date((dateVal - 25569) * 86400 * 1000);
    }
    if (typeof dateVal === 'string' && dateVal.trim()) {
        return parseDateString(dateVal.trim());
    }
    return null;
}

function parseDateString(str) {
    // Remove common wrappers
    str = str.replace(/^\s*["']+|["']+\s*$/g, '').trim();
    if (!str || str === '-' || str === '—' || str.toLowerCase() === 'null' || str.toLowerCase() === 'n/a') return null;

    let m, d;

    // ISO 8601 with time: "2024-06-15T00:00:00.000Z" or "2024-06-15T14:30:00"
    m = str.match(/^(\d{4})-(\d{2})-(\d{2})[T\s](\d{2}):(\d{2})/);
    if (m) {
        d = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]), parseInt(m[4]), parseInt(m[5]));
        if (!isNaN(d.getTime())) return d;
    }

    // yyyy-mm-dd (plain)
    m = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (m) {
        d = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
        if (!isNaN(d.getTime())) return d;
    }

    // dd/mm/yyyy  dd-mm-yyyy  dd.mm.yyyy (Italian/European)
    m = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
    if (m) {
        let day = parseInt(m[1]);
        let month = parseInt(m[2]) - 1;
        let year = parseInt(m[3]);
        if (year < 100) year += 2000;
        d = new Date(year, month, day);
        if (!isNaN(d.getTime())) return d;
    }

    // "15 Jun 2024", "Jun 15, 2024", "June 15 2024", "15 June 2024"
    const MONTHS = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11,
                     gen:0, feb:1, mar:2, apr:3, mag:4, giu:5, lug:6, ago:7, set:8, ott:9, nov:10, dic:11,
                     january:0, february:1, march:2, april:3, june:5, july:6, august:7, september:8, october:9, november:10, december:11,
                     gennaio:0, febbraio:1, marzo:2, aprile:3, maggio:4, giugno:5, luglio:6, agosto:7, settembre:8, ottobre:9, novembre:10, dicembre:11 };
    // "15 Jun 2024" or "15 June 2024"
    m = str.match(/^(\d{1,2})[\s\-]+([a-zA-Z]+)[\s,\-]+(\d{4})$/);
    if (m) {
        const mon = MONTHS[m[2].toLowerCase().substring(0, 3)];
        if (mon !== undefined) { d = new Date(parseInt(m[3]), mon, parseInt(m[1])); if (!isNaN(d.getTime())) return d; }
    }
    // "Jun 15, 2024" or "June 15 2024"
    m = str.match(/^([a-zA-Z]+)[\s]+(\d{1,2})[,\s]+(\d{4})$/);
    if (m) {
        const mon = MONTHS[m[1].toLowerCase().substring(0, 3)];
        if (mon !== undefined) { d = new Date(parseInt(m[3]), mon, parseInt(m[2])); if (!isNaN(d.getTime())) return d; }
    }

    // mm/dd/yyyy (American) — only if first number > 12, swap with Italian
    m = str.match(/^(\d{1,2})[\/](\d{1,2})[\/](\d{4})$/);
    if (m && parseInt(m[1]) > 12) {
        d = new Date(parseInt(m[3]), parseInt(m[1]) - 1, parseInt(m[2]));
        if (!isNaN(d.getTime())) return d;
    }

    // Fallback: let JavaScript try
    d = new Date(str);
    return isNaN(d.getTime()) ? null : d;
}

// ─── Filters ────────────────────────────────────────────────
function applyFilters() {
    const locFilter = document.getElementById('filterLocation').value;
    const statusFilter = document.getElementById('filterStatus').value;
    const modelFilter = document.getElementById('filterModel').value;
    const enterpriseFilter = document.getElementById('filterEnterprise').value;
    const specialistEl = document.getElementById('filterSpecialist');
    const specialistFilter = specialistEl ? specialistEl.value : 'all';
    const dateFrom = document.getElementById('filterDateFrom').value ? new Date(document.getElementById('filterDateFrom').value) : null;
    const dateTo = document.getElementById('filterDateTo').value ? new Date(document.getElementById('filterDateTo').value + 'T23:59:59') : null;

    filteredData = rawData.filter(row => {
        if (locFilter !== 'all' && row.location !== locFilter) return false;
        if (statusFilter !== 'all' && row.status !== statusFilter) return false;
        if (modelFilter !== 'all' && row.model !== modelFilter) return false;
        if (specialistFilter !== 'all' && (row.deliverySpecialist || '') !== specialistFilter) return false;
        if (enterpriseFilter !== 'all') {
            if (enterpriseFilter === 'B2B' && !row.isEnterprise) return false;
            if (enterpriseFilter === 'B2C' && row.isEnterprise) return false;
        }
        if (dateFrom && row.date < dateFrom) return false;
        if (dateTo && row.date > dateTo) return false;
        return true;
    });

    currentPage = 1;
    populateFilterDropdowns();
    updateKPIs();
    renderCharts();
    renderPivotTable();
    renderReadySection();
    renderPostponeTable();
    renderArrivalsTable();
    renderLocationCards();
}

function resetFilters() {
    document.getElementById('filterLocation').value = 'all';
    document.getElementById('filterStatus').value = 'all';
    document.getElementById('filterModel').value = 'all';
    if (document.getElementById('filterSpecialist')) document.getElementById('filterSpecialist').value = 'all';
    document.getElementById('filterEnterprise').value = 'all';
    document.getElementById('filterDateFrom').value = '';
    document.getElementById('filterDateTo').value = '';
    applyFilters();
}

function populateFilterDropdowns() {
    populateSelect('filterLocation', [...new Set(rawData.map(r => r.location))].sort());
    populateSelect('filterStatus', [...new Set(rawData.map(r => r.status))].sort());
    populateSelect('filterModel', [...new Set(rawData.map(r => r.model))].sort());
    const specialists = [...new Set(rawData.map(r => r.deliverySpecialist).filter(s => s && s.length > 0))].sort();
    if (document.getElementById('filterSpecialist')) populateSelect('filterSpecialist', specialists);
}

function populateSelect(id, options) {
    const sel = document.getElementById(id);
    const current = sel.value;
    const firstOpt = sel.options[0];
    sel.innerHTML = '';
    sel.appendChild(firstOpt);
    options.forEach(opt => {
        const o = document.createElement('option');
        o.value = opt;
        o.textContent = opt;
        sel.appendChild(o);
    });
    sel.value = current; // restore selection
}

// ─── KPIs ───────────────────────────────────────────────────
function updateKPIs() {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayEnd = new Date(today);
    todayEnd.setHours(23, 59, 59, 999);
    const next2 = new Date(today);
    next2.setDate(next2.getDate() + 2);
    next2.setHours(23, 59, 59, 999);
    const next7 = new Date(today);
    next7.setDate(next7.getDate() + 7);

    // Use delivery date if available, otherwise arrival date
    const getRefDate = r => r.deliveryDate || r.date;

    const todayDeliveries = filteredData.filter(r => { const d = getRefDate(r); return d >= today && d <= todayEnd; });
    const imminentOrders = filteredData.filter(r => r.urgency === 'imminent' || r.urgency === 'today');
    const upcoming = filteredData.filter(r => { const d = getRefDate(r); return d >= today && d <= next7; });
    const locations = new Set(filteredData.map(r => r.location));

    animateCounter('kpiTotal', filteredData.length);
    animateCounter('kpiToday', todayDeliveries.length);
    animateCounter('kpiImminent', imminentOrders.length);
    animateCounter('kpiUpcoming', upcoming.length);
    animateCounter('kpiLocations', locations.size);

    // Ground > 6 days KPI
    const groundOver6 = filteredData.filter(r => {
        if (!r.date) return false;
        const arrival = new Date(r.date); arrival.setHours(0,0,0,0);
        if (arrival > today) return false; // not arrived yet
        const daysOnGround = Math.round((today - arrival) / 86400000);
        return daysOnGround > 6;
    });
    animateCounter('kpiGround', groundOver6.length);

    // Flash ground KPI if there are overdue vehicles
    const groundCard = document.getElementById('kpiGroundCard');
    if (groundOver6.length > 0) {
        groundCard.classList.add('flashing');
    } else {
        groundCard.classList.remove('flashing');
    }

    // Activate flashing if there are imminent orders
    const imminentCard = document.getElementById('kpiImminentCard');
    if (imminentOrders.length > 0) {
        imminentCard.classList.add('flashing');
    } else {
        imminentCard.classList.remove('flashing');
    }
}

function animateCounter(id, target) {
    const el = document.getElementById(id);
    const start = parseInt(el.textContent) || 0;
    const duration = 600;
    const startTime = performance.now();

    function update(currentTime) {
        const elapsed = currentTime - startTime;
        const progress = Math.min(elapsed / duration, 1);
        const eased = 1 - Math.pow(1 - progress, 3);
        el.textContent = Math.round(start + (target - start) * eased).toLocaleString('it-IT');
        if (progress < 1) requestAnimationFrame(update);
    }
    requestAnimationFrame(update);
}

// ─── Charts ─────────────────────────────────────────────────
function renderCharts() {
    renderPieChart('chartByLocation', groupBy(filteredData, 'location'), 'Volumi per Hub');
    renderPieChart('chartByModel', groupBy(filteredData, 'model'), 'Distribuzione per Modello');
    renderPieChart('chartByStatus', groupBy(filteredData, 'status'), 'Stato Ordini');
    renderTimelineChart();
}

function groupBy(data, key) {
    const groups = {};
    data.forEach(row => {
        const k = row[key] || 'N/A';
        groups[k] = (groups[k] || 0) + 1;
    });
    return groups;
}

function renderPieChart(canvasId, dataObj, title) {
    const ctx = document.getElementById(canvasId).getContext('2d');

    // Destroy existing
    if (charts[canvasId]) {
        charts[canvasId].destroy();
    }

    const labels = Object.keys(dataObj);
    const values = Object.values(dataObj);

    charts[canvasId] = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: values,
                backgroundColor: CHART_COLORS.slice(0, labels.length),
                borderColor: 'rgba(10, 22, 40, 0.8)',
                borderWidth: 2,
                hoverBorderWidth: 3,
                hoverBorderColor: '#fff',
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '55%',
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        color: '#8ba3c7',
                        font: { size: 11, family: 'Inter' },
                        padding: 12,
                        usePointStyle: true,
                        pointStyleWidth: 8,
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(10, 22, 40, 0.95)',
                    titleColor: '#e8edf5',
                    bodyColor: '#8ba3c7',
                    borderColor: 'rgba(59, 130, 246, 0.3)',
                    borderWidth: 1,
                    padding: 12,
                    titleFont: { family: 'Inter', weight: 600 },
                    bodyFont: { family: 'Inter' },
                    callbacks: {
                        label: function(context) {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const pct = ((context.parsed / total) * 100).toFixed(1);
                            return ` ${context.label}: ${context.parsed.toLocaleString('it-IT')} (${pct}%)`;
                        }
                    }
                }
            },
            animation: {
                animateRotate: true,
                animateScale: true,
                duration: 800,
            }
        }
    });
}

function renderTimelineChart() {
    const ctx = document.getElementById('chartTimeline').getContext('2d');
    if (charts['chartTimeline']) charts['chartTimeline'].destroy();

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Group by date for next 30 days
    const dayMap = {};
    for (let i = 0; i < 30; i++) {
        const d = new Date(today);
        d.setDate(d.getDate() + i);
        const key = d.toISOString().split('T')[0];
        dayMap[key] = { date: d, total: 0 };
    }

    // Also include past 7 days
    for (let i = 7; i >= 1; i--) {
        const d = new Date(today);
        d.setDate(d.getDate() - i);
        const key = d.toISOString().split('T')[0];
        dayMap[key] = { date: d, total: 0 };
    }

    filteredData.forEach(row => {
        if (row.date) {
            const key = row.date.toISOString().split('T')[0];
            if (dayMap[key]) {
                dayMap[key].total += 1;
            }
        }
    });

    const sortedKeys = Object.keys(dayMap).sort();
    const labels = sortedKeys.map(k => {
        const d = dayMap[k].date;
        return d.toLocaleDateString('it-IT', { day: '2-digit', month: 'short' });
    });
    const values = sortedKeys.map(k => dayMap[k].total);

    // Find today index for annotation
    const todayKey = today.toISOString().split('T')[0];
    const todayIdx = sortedKeys.indexOf(todayKey);

    charts['chartTimeline'] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Arrivi (qty)',
                data: values,
                backgroundColor: sortedKeys.map((k, i) => {
                    if (k === todayKey) return 'rgba(34, 197, 94, 0.7)';
                    if (k < todayKey) return 'rgba(107, 114, 128, 0.4)';
                    return 'rgba(59, 130, 246, 0.6)';
                }),
                borderColor: sortedKeys.map((k) => {
                    if (k === todayKey) return '#22c55e';
                    if (k < todayKey) return 'rgba(107, 114, 128, 0.6)';
                    return '#3b82f6';
                }),
                borderWidth: 1,
                borderRadius: 4,
                borderSkipped: false,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(10, 22, 40, 0.95)',
                    titleColor: '#e8edf5',
                    bodyColor: '#8ba3c7',
                    borderColor: 'rgba(59, 130, 246, 0.3)',
                    borderWidth: 1,
                    padding: 12,
                    titleFont: { family: 'Inter', weight: 600 },
                    bodyFont: { family: 'Inter' },
                }
            },
            scales: {
                x: {
                    grid: { color: 'rgba(59, 130, 246, 0.06)' },
                    ticks: {
                        color: '#5a7a9e',
                        font: { size: 10, family: 'Inter' },
                        maxRotation: 45,
                    }
                },
                y: {
                    beginAtZero: true,
                    grid: { color: 'rgba(59, 130, 246, 0.06)' },
                    ticks: {
                        color: '#5a7a9e',
                        font: { size: 11, family: 'Inter' },
                        stepSize: 1,
                    }
                }
            },
            animation: { duration: 800 }
        }
    });
}

// ─── Pivot Table ────────────────────────────────────────────
function renderPivotTable() {
    const head = document.getElementById('pivotHead');
    const body = document.getElementById('pivotBody');

    // Get unique locations and models
    const locations = [...new Set(filteredData.map(r => r.location))].sort();
    const models = [...new Set(filteredData.map(r => r.model))].sort();

    // Build matrix
    const matrix = {};
    const locationTotals = {};
    const modelTotals = {};
    let grandTotal = 0;

    locations.forEach(loc => {
        matrix[loc] = {};
        locationTotals[loc] = 0;
        models.forEach(mod => {
            matrix[loc][mod] = 0;
        });
    });
    models.forEach(mod => modelTotals[mod] = 0);

    filteredData.forEach(row => {
        if (matrix[row.location]) {
            matrix[row.location][row.model] = (matrix[row.location][row.model] || 0) + 1;
            locationTotals[row.location] += 1;
            modelTotals[row.model] = (modelTotals[row.model] || 0) + 1;
            grandTotal += 1;
        }
    });

    // Render header
    head.innerHTML = '<tr><th>Location / Hub</th>' +
        models.map(m => '<th>' + escapeHtml(m) + '</th>').join('') +
        '<th>TOTALE</th></tr>';

    // Render rows
    body.innerHTML = locations.map(loc => {
        return '<tr><td><strong>' + escapeHtml(loc) + '</strong></td>' +
            models.map(mod => '<td>' + (matrix[loc][mod] || 0).toLocaleString('it-IT') + '</td>').join('') +
            '<td><strong>' + locationTotals[loc].toLocaleString('it-IT') + '</strong></td></tr>';
    }).join('') +
    // Total row
    '<tr class="pivot-total"><td><strong>TOTALE</strong></td>' +
    models.map(mod => '<td>' + modelTotals[mod].toLocaleString('it-IT') + '</td>').join('') +
    '<td>' + grandTotal.toLocaleString('it-IT') + '</td></tr>';
}

// ─── Postpone Table ─────────────────────────────────────────
function renderPostponeTable() {
    const section = document.getElementById('postponeSection');
    const body = document.getElementById('postponeBody');
    const countEl = document.getElementById('postponeCount');

    // Find orders where delivery date is BEFORE arrival date
    const postponeOrders = filteredData.filter(row => {
        if (!row.deliveryDate || !row.date) return false;
        const delivery = new Date(row.deliveryDate);
        delivery.setHours(0, 0, 0, 0);
        const arrival = new Date(row.date);
        arrival.setHours(0, 0, 0, 0);
        return delivery < arrival;
    }).map(row => {
        const delivery = new Date(row.deliveryDate);
        delivery.setHours(0, 0, 0, 0);
        const arrival = new Date(row.date);
        arrival.setHours(0, 0, 0, 0);
        const delayDays = Math.round((arrival - delivery) / 86400000);
        return { ...row, delayDays };
    }).sort((a, b) => b.delayDays - a.delayDays); // worst first

    if (postponeOrders.length === 0) {
        section.style.display = 'none';
        return;
    }

    section.style.display = 'block';
    countEl.textContent = postponeOrders.length + ' ordini';

    body.innerHTML = postponeOrders.map(row => {
        const statusClass = 'status-' + row.status.toLowerCase().replace(/\s+/g, '-');
        const isCritical = row.delayDays >= 5;

        let rnCell = '';
        if (row.reservationNumber) {
            rnCell = `<td><a href="https://dro.tesla.com/advisor?sidepanel_fullscreen=yes&rn=${encodeURIComponent(row.reservationNumber)}" target="_blank" rel="noopener" class="wdo-link" title="Apri DRO Advisor - ${escapeHtml(row.reservationNumber)}">${escapeHtml(row.reservationNumber)} <svg class="wdo-link-icon" width="12" height="12" viewBox="0 0 12 12" fill="none"><path d="M4 1H2C1.45 1 1 1.45 1 2V10C1 10.55 1.45 11 2 11H10C10.55 11 11 10.55 11 10V8" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"/><path d="M7 1H11V5" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/><path d="M5 7L11 1" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"/></svg></a></td>`;
        } else {
            rnCell = '<td>—</td>';
        }

        return `<tr>
            <td>${escapeHtml(row.orderId)}</td>
            ${rnCell}
            <td>${escapeHtml(row.location)}</td>
            <td style="color:#ef4444;font-weight:600;">${row.deliveryDateStr}</td>
            <td>${row.dateStr}</td>
            <td><span class="delay-days ${isCritical ? 'delay-critical' : ''}">+${row.delayDays}gg</span></td>
            <td>${escapeHtml(row.model)}</td>
            <td><span class="status-badge ${statusClass}">${escapeHtml(row.status)}</span></td>
        </tr>`;
    }).join('');
}

function exportPostponeCSV() {
    const table = document.getElementById('postponeTable');
    downloadTableCSV(table, 'tesla_fleet_da_posticipare.csv');
}

// ─── Ready to Deliver Section ───────────────────────────────
function renderReadySection() {
    const section = document.getElementById('readySection');
    if (!section) return;
    const greenBody = document.getElementById('readyGreenBody');
    const redBody = document.getElementById('readyRedBody');
    const greenCount = document.getElementById('readyGreenCount');
    const redCount = document.getElementById('readyRedCount');
    const redSection = document.getElementById('readyRedSection');

    // Check if File 2 data is available
    const hasFile2 = filteredData.some(r => r.finalPaymentStatus !== undefined);
    if (!hasFile2) {
        section.style.display = 'none';
        return;
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // "A terra in hub" = la Posizione (LastKnownVehicleLocation) combacia con l'Hub assegnato
    // Confronto fuzzy: basta che la posizione CONTENGA il nome dell'hub o viceversa
    // Es: location="Milano Linate SC", lastKnownLocation="Milano Linate SC" → match
    // Es: location="Roma Magliana SC", lastKnownLocation="At SC" → match su "SC"
    // Fallback: se non c'è lastKnownLocation, usa la data arrivo <= oggi
    const isPaymentComplete = (r) => {
        const ps = (r.finalPaymentStatus || '').trim().toUpperCase();
        if (ps.includes('INCOMPLETE') || ps.includes('NOT COMPLETE') || ps.includes('NOT PAID') || ps.includes('PENDING')) return false;
        return ps === 'COMPLETE' || ps === 'PAID' || ps === 'RECEIVED';
    };

    const isAtHub = (r) => {
        const pos = (r.lastKnownLocation || '').trim().toLowerCase();
        const hub = (r.location || '').trim().toLowerCase();

        // Se abbiamo la posizione, verifichiamo che combaci con l'hub
        if (pos && hub) {
            // Match diretto
            if (pos === hub) return true;
            // Match parziale: la posizione contiene il nome dell'hub o viceversa
            if (pos.includes(hub) || hub.includes(pos)) return true;
            // Match per keyword: estraiamo la città dall'hub e cerchiamo nella posizione
            // Es: "Milano Linate SC" → cerchiamo "milano" in posizione
            const hubCity = hub.replace(/\s*(sc|service center|hub|compound|terminal|porto)\s*/gi, '').trim();
            if (hubCity.length >= 3 && pos.includes(hubCity)) return true;
            // Match "At SC" o "at service center" → generico, consideriamo a terra
            if (pos === 'at sc' || pos.includes('at service center')) return true;
            // Se la posizione è diversa dall'hub → NON è a terra
            return false;
        }

        // Fallback se non c'è lastKnownLocation: usa la data arrivo
        if (!r.date) return false;
        const arrival = new Date(r.date);
        arrival.setHours(0, 0, 0, 0);
        return arrival <= today;
    };

    // Filter: Payment complete + at hub = READY
    const readyOrders = filteredData.filter(r => isPaymentComplete(r) && isAtHub(r));

    // Filter: Payment NOT complete + at hub = NEED PAYMENT
    const needPaymentOrders = filteredData.filter(r => !isPaymentComplete(r) && isAtHub(r) && r.finalPaymentStatus);

    // Show/hide payment section
    const paymentSection = document.getElementById('paymentSection');
    const paymentBody = document.getElementById('paymentBody');
    const paymentCount = document.getElementById('paymentCount');

    if (needPaymentOrders.length > 0) {
        paymentSection.style.display = 'block';
        paymentCount.textContent = needPaymentOrders.length + ' da sollecitare';

        // Add days on ground
        needPaymentOrders.forEach(r => {
            const arrival = new Date(r.date); arrival.setHours(0,0,0,0);
            r._daysOnGround = Math.round((today - arrival) / 86400000);
        });
        needPaymentOrders.sort((a, b) => b._daysOnGround - a._daysOnGround);

        // Store for pagination
        window._paymentAll = needPaymentOrders;
        const totalPayPages = Math.ceil(needPaymentOrders.length / paymentPageSize);
        if (paymentPage > totalPayPages) paymentPage = 1;
        const payStart = (paymentPage - 1) * paymentPageSize;
        const payPageData = needPaymentOrders.slice(payStart, payStart + paymentPageSize);

        paymentBody.innerHTML = payPageData.map(row => {
            let rnCell = row.reservationNumber
                ? `<td><a href="https://dro.tesla.com/advisor?sidepanel_fullscreen=yes&rn=${encodeURIComponent(row.reservationNumber)}" target="_blank" class="wdo-link">${escapeHtml(row.reservationNumber)}</a></td>`
                : '<td>—</td>';

            const d = row._daysOnGround;
            let groundClass = 'ground-ok';
            if (d > 6) groundClass = 'ground-critical';
            else if (d >= 4) groundClass = 'ground-warn';
            const groundBadge = `<span class="ground-badge ${groundClass}">${d}gg</span>`;

            const channelBadge = row.orderChannel === 'B2B'
                ? '<span class="channel-badge channel-b2b">B2B</span>'
                : row.orderChannel === 'B2C' ? '<span class="channel-badge channel-b2c">B2C</span>' : '—';

            return `<tr>
                <td>${escapeHtml(row.orderId)}</td>
                ${rnCell}
                <td>${escapeHtml(row.location)}</td>
                <td>${escapeHtml(row.model)}</td>
                <td>${groundBadge}</td>
                <td style="color:#f97316;font-weight:700;">${escapeHtml(row.finalPaymentStatus || '')}</td>
                <td>${escapeHtml(row.deliverySpecialist || '—')}</td>
                <td>${channelBadge}</td>
                <td style="font-size:0.8rem;">${escapeHtml(row.lastKnownLocation || '—')}</td>
            </tr>`;
        }).join('');

        // Payment pagination
        renderTablePagination('paymentPagination', paymentPage, totalPayPages, 'goPaymentPage');
    } else {
        paymentSection.style.display = 'none';
    }

    if (readyOrders.length === 0) {
        section.style.display = 'none';
        return;
    }

    section.style.display = 'block';

    // Split: green (no CH) vs red (CH active)
    const greenOrders = readyOrders.filter(r => !r.isContainmentHold);
    const redOrders = readyOrders.filter(r => r.isContainmentHold);

    // Calculate days on ground for each
    const addDaysOnGround = (r) => {
        const arrival = new Date(r.date); arrival.setHours(0,0,0,0);
        r._daysOnGround = Math.round((today - arrival) / 86400000);
        return r;
    };
    greenOrders.forEach(addDaysOnGround);
    redOrders.forEach(addDaysOnGround);

    // Sort green: worst (most days on ground) first
    greenOrders.sort((a, b) => b._daysOnGround - a._daysOnGround);

    const overdue = greenOrders.filter(r => r._daysOnGround > 6).length;
    greenCount.textContent = greenOrders.length + ' pronti' + (overdue ? ' | ' + overdue + ' SCADUTI' : '');

    // Store globally for pagination
    window._readyGreenAll = greenOrders;

    // Paginate green table
    const totalReadyPages = Math.ceil(greenOrders.length / readyPageSize);
    if (readyPage > totalReadyPages) readyPage = 1;
    const readyStart = (readyPage - 1) * readyPageSize;
    const readyPageData = greenOrders.slice(readyStart, readyStart + readyPageSize);

    // Render green table
    greenBody.innerHTML = readyPageData.map(row => {
        let rnCell = row.reservationNumber
            ? `<td><a href="https://dro.tesla.com/advisor?sidepanel_fullscreen=yes&rn=${encodeURIComponent(row.reservationNumber)}" target="_blank" class="wdo-link">${escapeHtml(row.reservationNumber)}</a></td>`
            : '<td>—</td>';

        const channelBadge = row.orderChannel === 'B2B'
            ? '<span class="channel-badge channel-b2b">B2B</span>'
            : row.orderChannel === 'B2C' ? '<span class="channel-badge channel-b2c">B2C</span>' : '—';

        const etaBadge = row.isConfidentETA === true
            ? '<span class="eta-badge eta-confident">Confident</span>'
            : row.isConfidentETA === false ? '<span class="eta-badge eta-notconfident">Not Confident</span>' : '—';

        const schedBadge = row.isScheduled
            ? '<span style="color:#22c55e;font-weight:700;">Si</span>'
            : '<span style="color:#f97316;font-weight:700;">Da schedulare</span>';

        // Days on ground badge with traffic light
        const d = row._daysOnGround;
        let groundClass = 'ground-ok';
        if (d > 6) groundClass = 'ground-critical';
        else if (d >= 4) groundClass = 'ground-warn';
        const groundBadge = `<span class="ground-badge ${groundClass}">${d}gg${d > 6 ? ' !!!' : ''}</span>`;

        const rowCls = d > 6 ? 'row-containment' : d >= 4 ? 'row-imminent' : '';

        return `<tr class="${rowCls}">
            <td>${escapeHtml(row.orderId)}</td>
            ${rnCell}
            <td>${escapeHtml(row.location)}</td>
            <td>${escapeHtml(row.model)}</td>
            <td>${row.dateStr}</td>
            <td>${groundBadge}</td>
            <td style="color:#22c55e;font-weight:600;">${escapeHtml(row.finalPaymentStatus || '')}</td>
            <td>${escapeHtml(row.deliverySpecialist || '—')}</td>
            <td>${channelBadge}</td>
            <td>${etaBadge}</td>
            <td>${schedBadge}</td>
            <td style="font-size:0.8rem;">${escapeHtml(row.lastKnownLocation || '—')}</td>
        </tr>`;
    }).join('') || '<tr><td colspan="11" style="text-align:center;color:#5a7a9e;">Nessun ordine pronto senza blocchi</td></tr>';

    // Ready pagination
    renderTablePagination('readyPagination', readyPage, totalReadyPages, 'goReadyPage');

    // Render red table (CH blocked)
    if (redOrders.length > 0) {
        redSection.style.display = 'block';
        redCount.textContent = redOrders.length + ' bloccati';

        redBody.innerHTML = redOrders.map(row => {
            let rnCell = row.reservationNumber
                ? `<td><a href="https://dro.tesla.com/advisor?sidepanel_fullscreen=yes&rn=${encodeURIComponent(row.reservationNumber)}" target="_blank" class="wdo-link">${escapeHtml(row.reservationNumber)}</a></td>`
                : '<td>—</td>';

            const statusClass = 'status-' + row.status.toLowerCase().replace(/\s+/g, '-');

            return `<tr class="row-containment">
                <td>${escapeHtml(row.orderId)}</td>
                ${rnCell}
                <td>${escapeHtml(row.location)}</td>
                <td>${escapeHtml(row.model)}</td>
                <td>${row.dateStr}</td>
                <td style="color:#22c55e;font-weight:600;">${escapeHtml(row.finalPaymentStatus || '')}</td>
                <td><span class="ch-badge ch-true">CH</span></td>
                <td style="font-size:0.8rem;">${escapeHtml(row.lastKnownLocation || '—')}</td>
                <td><span class="status-badge ${statusClass}">${escapeHtml(row.status)}</span></td>
            </tr>`;
        }).join('');
    } else {
        redSection.style.display = 'none';
    }
}

function exportReadyCSV(type) {
    const table = document.getElementById(type === 'green' ? 'readyGreenTable' : 'readyRedTable');
    const filename = type === 'green' ? 'tesla_fleet_pronti_consegna.csv' : 'tesla_fleet_bloccati_ch.csv';
    downloadTableCSV(table, filename);
}

function exportPaymentCSV() {
    const table = document.getElementById('paymentTable');
    downloadTableCSV(table, 'tesla_fleet_da_pagare.csv');
}

// ─── Generic Table Pagination ───────────────────────────────
function renderTablePagination(containerId, currentPg, totalPgs, fnName) {
    const pag = document.getElementById(containerId);
    if (!pag || totalPgs <= 1) { if (pag) pag.innerHTML = ''; return; }

    let html = `<button onclick="${fnName}(${currentPg - 1})" ${currentPg === 1 ? 'disabled' : ''}>&laquo;</button>`;
    const maxV = 5;
    let sp = Math.max(1, currentPg - Math.floor(maxV / 2));
    let ep = Math.min(totalPgs, sp + maxV - 1);
    if (ep - sp < maxV - 1) sp = Math.max(1, ep - maxV + 1);

    if (sp > 1) html += `<button onclick="${fnName}(1)">1</button>`;
    if (sp > 2) html += `<button disabled>...</button>`;
    for (let i = sp; i <= ep; i++) {
        html += `<button onclick="${fnName}(${i})" class="${i === currentPg ? 'active' : ''}">${i}</button>`;
    }
    if (ep < totalPgs - 1) html += `<button disabled>...</button>`;
    if (ep < totalPgs) html += `<button onclick="${fnName}(${totalPgs})">${totalPgs}</button>`;
    html += `<button onclick="${fnName}(${currentPg + 1})" ${currentPg === totalPgs ? 'disabled' : ''}>&raquo;</button>`;
    pag.innerHTML = html;
}

function goReadyPage(p) { readyPage = p; renderReadySection(); }
function goPaymentPage(p) { paymentPage = p; renderReadySection(); }
function changeReadyPageSize() { readyPageSize = parseInt(document.getElementById('readyPageSize').value); readyPage = 1; renderReadySection(); }
function changePaymentPageSize() { paymentPageSize = parseInt(document.getElementById('paymentPageSize').value); paymentPage = 1; renderReadySection(); }

// ─── Arrivals Table ─────────────────────────────────────────
function renderArrivalsTable() {
    const body = document.getElementById('arrivalsBody');

    // Check if we have reservation data and delivery date data
    const hasReservation = filteredData.some(r => r.reservationNumber);
    const hasDeliveryDate = filteredData.some(r => r.deliveryDate);

    // Update table header dynamically
    const thead = document.querySelector('#arrivalsTable thead tr');
    let colIdx = 0;
    let colMap = {}; // track which sortCol index maps to which field

    let headerHtml = '';
    headerHtml += `<th onclick="sortTable(${colIdx})">Order ID / VIN <span class="sort-icon">&#8597;</span></th>`;
    colMap[colIdx] = 'orderId'; colIdx++;

    if (hasReservation) {
        headerHtml += `<th onclick="sortTable(${colIdx})">Reservation # <span class="sort-icon">&#8597;</span></th>`;
        colMap[colIdx] = 'reservationNumber'; colIdx++;
    }

    headerHtml += `<th onclick="sortTable(${colIdx})">Hub / Location <span class="sort-icon">&#8597;</span></th>`;
    colMap[colIdx] = 'location'; colIdx++;

    headerHtml += `<th onclick="sortTable(${colIdx})">Data Arrivo <span class="sort-icon">&#8597;</span></th>`;
    colMap[colIdx] = 'date'; colIdx++;

    // Check for File 2 enriched data
    const hasFile2 = filteredData.some(r => r.orderChannel || r.isConfidentETA !== undefined || r.isContainmentHold !== undefined);

    if (hasDeliveryDate) {
        headerHtml += `<th onclick="sortTable(${colIdx})">Data Consegna <span class="sort-icon">&#8597;</span></th>`;
        colMap[colIdx] = 'deliveryDate'; colIdx++;
    }

    headerHtml += `<th onclick="sortTable(${colIdx})">Modello <span class="sort-icon">&#8597;</span></th>`;
    colMap[colIdx] = 'model'; colIdx++;

    headerHtml += `<th onclick="sortTable(${colIdx})">Stato <span class="sort-icon">&#8597;</span></th>`;
    colMap[colIdx] = 'status'; colIdx++;

    if (hasFile2) {
        headerHtml += `<th onclick="sortTable(${colIdx})">B2B/B2C <span class="sort-icon">&#8597;</span></th>`;
        colMap[colIdx] = 'orderChannel'; colIdx++;

        headerHtml += `<th onclick="sortTable(${colIdx})">ETA <span class="sort-icon">&#8597;</span></th>`;
        colMap[colIdx] = 'isConfidentETA'; colIdx++;

        headerHtml += `<th onclick="sortTable(${colIdx})">CH <span class="sort-icon">&#8597;</span></th>`;
        colMap[colIdx] = 'isContainmentHold'; colIdx++;

        headerHtml += `<th onclick="sortTable(${colIdx})">Posizione <span class="sort-icon">&#8597;</span></th>`;
        colMap[colIdx] = 'lastKnownLocation'; colIdx++;

        headerHtml += `<th onclick="sortTable(${colIdx})">Pagamento <span class="sort-icon">&#8597;</span></th>`;
        colMap[colIdx] = 'finalPaymentStatus'; colIdx++;

        headerHtml += `<th onclick="sortTable(${colIdx})">Sched. <span class="sort-icon">&#8597;</span></th>`;
        colMap[colIdx] = 'isScheduled'; colIdx++;
    }

    thead.innerHTML = headerHtml;

    // Sort
    const sortField = colMap[sortCol] || 'date';
    const sorted = [...filteredData].sort((a, b) => {
        let va, vb;
        switch (sortField) {
            case 'orderId':           va = a.orderId; vb = b.orderId; break;
            case 'reservationNumber': va = a.reservationNumber; vb = b.reservationNumber; break;
            case 'location':          va = a.location; vb = b.location; break;
            case 'date':              va = a.date ? a.date.getTime() : 0; vb = b.date ? b.date.getTime() : 0; break;
            case 'deliveryDate':      va = a.deliveryDate ? a.deliveryDate.getTime() : 0; vb = b.deliveryDate ? b.deliveryDate.getTime() : 0; break;
            case 'model':             va = a.model; vb = b.model; break;
            case 'status':            va = a.status; vb = b.status; break;
            case 'orderChannel':      va = a.orderChannel || ''; vb = b.orderChannel || ''; break;
            case 'isConfidentETA':    va = a.isConfidentETA ? 1 : 0; vb = b.isConfidentETA ? 1 : 0; break;
            case 'isContainmentHold': va = a.isContainmentHold ? 1 : 0; vb = b.isContainmentHold ? 1 : 0; break;
            case 'lastKnownLocation': va = a.lastKnownLocation || ''; vb = b.lastKnownLocation || ''; break;
            case 'finalPaymentStatus':va = a.finalPaymentStatus || ''; vb = b.finalPaymentStatus || ''; break;
            case 'isScheduled':       va = a.isScheduled ? 1 : 0; vb = b.isScheduled ? 1 : 0; break;
            default:                  va = 0; vb = 0;
        }
        if (typeof va === 'string') {
            return sortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
        }
        return sortAsc ? va - vb : vb - va;
    });

    // Pagination
    const totalPages = Math.ceil(sorted.length / pageSize);
    const start = (currentPage - 1) * pageSize;
    const pageData = sorted.slice(start, start + pageSize);

    body.innerHTML = pageData.map(row => {
        const statusClass = 'status-' + row.status.toLowerCase().replace(/\s+/g, '-');
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const isToday = row.date && row.date.toDateString() === today.toDateString();
        const isPast = row.date && row.date < today;

        // Urgency CSS class for entire row
        let rowClass = '';
        if (row.urgency === 'imminent') rowClass = 'row-imminent';
        else if (row.urgency === 'today') rowClass = 'row-today';
        else if (row.urgency === 'soon') rowClass = 'row-soon';

        // Build Reservation Number cell
        let reservationCell = '';
        if (hasReservation) {
            if (row.reservationNumber) {
                reservationCell = `<td><a href="https://dro.tesla.com/advisor?sidepanel_fullscreen=yes&rn=${encodeURIComponent(row.reservationNumber)}" target="_blank" rel="noopener" class="wdo-link" title="Apri DRO Advisor - ${escapeHtml(row.reservationNumber)}">${escapeHtml(row.reservationNumber)} <svg class="wdo-link-icon" width="12" height="12" viewBox="0 0 12 12" fill="none"><path d="M4 1H2C1.45 1 1 1.45 1 2V10C1 10.55 1.45 11 2 11H10C10.55 11 11 10.55 11 10V8" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"/><path d="M7 1H11V5" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/><path d="M5 7L11 1" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"/></svg></a></td>`;
            } else {
                reservationCell = `<td>—</td>`;
            }
        }

        // Build delivery date cell with urgency badge
        let deliveryCell = '';
        if (hasDeliveryDate) {
            let urgencyBadge = '';
            if (row.urgency === 'today') {
                urgencyBadge = `<span class="urgency-badge urgency-today">OGGI</span>`;
            } else if (row.urgency === 'imminent') {
                urgencyBadge = `<span class="urgency-badge urgency-imminent">${row.daysUntil}g</span>`;
            } else if (row.urgency === 'soon') {
                urgencyBadge = `<span class="urgency-badge urgency-soon">${row.daysUntil}g</span>`;
            }
            deliveryCell = `<td class="delivery-cell ${row.urgency === 'imminent' || row.urgency === 'today' ? 'delivery-flash' : ''}">${row.deliveryDateStr} ${urgencyBadge}</td>`;
        }

        // Containment Hold override — row flashes red
        if (row.isContainmentHold) rowClass = 'row-containment';

        // Build File 2 extra cells
        let file2Cells = '';
        if (hasFile2) {
            // B2B/B2C
            const channelBadge = row.orderChannel === 'B2B'
                ? '<span class="channel-badge channel-b2b">B2B</span>'
                : row.orderChannel === 'B2C'
                    ? '<span class="channel-badge channel-b2c">B2C</span>'
                    : '—';

            // Confident ETA
            const etaBadge = row.isConfidentETA === true
                ? '<span class="eta-badge eta-confident">Confident</span>'
                : row.isConfidentETA === false
                    ? '<span class="eta-badge eta-notconfident">Not Confident</span>'
                    : '—';

            // Containment Hold
            const chBadge = row.isContainmentHold
                ? '<span class="ch-badge ch-true">CH</span>'
                : row.isContainmentHold === false ? '<span style="color:#5a7a9e;">—</span>' : '—';

            // Last Known Location
            const locCell = row.lastKnownLocation || '—';

            // Payment Status
            const payCell = row.finalPaymentStatus || '—';

            // Scheduled
            const schedBadge = row.isScheduled
                ? '<span style="color:#22c55e;">Si</span>'
                : row.isScheduled === false ? '<span style="color:#5a7a9e;">No</span>' : '—';

            file2Cells = `
                <td>${channelBadge}</td>
                <td>${etaBadge}</td>
                <td>${chBadge}</td>
                <td style="font-size:0.8rem;">${escapeHtml(locCell)}</td>
                <td style="font-size:0.8rem;">${escapeHtml(payCell)}</td>
                <td>${schedBadge}</td>
            `;
        }

        return `<tr class="${rowClass}">
            <td>${escapeHtml(row.orderId)}</td>
            ${reservationCell}
            <td>${escapeHtml(row.location)}</td>
            <td style="${isToday ? 'color:#22c55e;font-weight:600' : isPast ? 'color:#6b7280' : ''}">${row.dateStr}${isToday ? ' (OGGI)' : ''}</td>
            ${deliveryCell}
            <td>${escapeHtml(row.model)}</td>
            <td><span class="status-badge ${statusClass}">${escapeHtml(row.status)}</span></td>
            ${file2Cells}
        </tr>`;
    }).join('');

    // Record count
    document.getElementById('recordCount').textContent = sorted.length + ' record';

    // Pagination controls
    renderPagination(totalPages);
}

function renderPagination(totalPages) {
    const pag = document.getElementById('pagination');
    if (totalPages <= 1) { pag.innerHTML = ''; return; }

    let html = '';
    html += `<button onclick="goToPage(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''}>&laquo; Prec</button>`;

    const maxVisible = 7;
    let startPage = Math.max(1, currentPage - Math.floor(maxVisible / 2));
    let endPage = Math.min(totalPages, startPage + maxVisible - 1);
    if (endPage - startPage < maxVisible - 1) startPage = Math.max(1, endPage - maxVisible + 1);

    if (startPage > 1) {
        html += `<button onclick="goToPage(1)">1</button>`;
        if (startPage > 2) html += `<button disabled>...</button>`;
    }

    for (let i = startPage; i <= endPage; i++) {
        html += `<button onclick="goToPage(${i})" class="${i === currentPage ? 'active' : ''}">${i}</button>`;
    }

    if (endPage < totalPages) {
        if (endPage < totalPages - 1) html += `<button disabled>...</button>`;
        html += `<button onclick="goToPage(${totalPages})">${totalPages}</button>`;
    }

    html += `<button onclick="goToPage(${currentPage + 1})" ${currentPage === totalPages ? 'disabled' : ''}>Succ &raquo;</button>`;

    pag.innerHTML = html;
}

function goToPage(page) {
    currentPage = page;
    renderArrivalsTable();
}

function sortTable(col) {
    if (sortCol === col) {
        sortAsc = !sortAsc;
    } else {
        sortCol = col;
        sortAsc = true;
    }
    renderArrivalsTable();
}

// ─── Location Summary Cards ─────────────────────────────────
function renderLocationCards() {
    const container = document.getElementById('locationCards');
    const locationData = {};

    filteredData.forEach(row => {
        if (!locationData[row.location]) {
            locationData[row.location] = {
                total: 0,
                models: {},
                statuses: {},
                nextArrival: null,
            };
        }
        const loc = locationData[row.location];
        loc.total += 1;
        loc.models[row.model] = (loc.models[row.model] || 0) + 1;
        loc.statuses[row.status] = (loc.statuses[row.status] || 0) + 1;

        const today = new Date();
        today.setHours(0, 0, 0, 0);
        if (row.date >= today && (!loc.nextArrival || row.date < loc.nextArrival)) {
            loc.nextArrival = row.date;
        }
    });

    const maxVolume = Math.max(...Object.values(locationData).map(l => l.total), 1);

    const sortedLocations = Object.entries(locationData).sort((a, b) => b[1].total - a[1].total);

    container.innerHTML = sortedLocations.map(([name, data]) => {
        const barWidth = (data.total / maxVolume * 100).toFixed(1);
        const modelEntries = Object.entries(data.models).sort((a, b) => b[1] - a[1]).slice(0, 4);
        const statusEntries = Object.entries(data.statuses).sort((a, b) => b[1] - a[1]).slice(0, 4);

        return `<div class="location-card">
            <div class="location-card-header">
                <span class="location-card-name">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none"><path d="M12 2 C8.13 2 5 5.13 5 9 C5 14.25 12 22 12 22 C12 22 19 14.25 19 9 C19 5.13 15.87 2 12 2Z" stroke="currentColor" stroke-width="2"/><circle cx="12" cy="9" r="2" stroke="currentColor" stroke-width="2"/></svg>
                    ${escapeHtml(name)}
                </span>
                <span class="location-card-count">${data.total.toLocaleString('it-IT')}</span>
            </div>
            <div class="location-card-details">
                ${modelEntries.map(([m, q]) =>
                    `<div class="location-detail">
                        <span class="location-detail-label">${escapeHtml(m)}</span>
                        <span class="location-detail-value">${q}</span>
                    </div>`
                ).join('')}
            </div>
            ${data.nextArrival ? `<div style="margin-top:8px;font-size:0.75rem;color:#06b6d4;">Prossimo arrivo: ${data.nextArrival.toLocaleDateString('it-IT')}</div>` : ''}
            <div class="location-card-bar">
                <div class="location-card-bar-fill" style="width: ${barWidth}%"></div>
            </div>
        </div>`;
    }).join('');
}

// ─── Export ─────────────────────────────────────────────────
function exportPivotCSV() {
    const table = document.getElementById('pivotTable');
    downloadTableCSV(table, 'tesla_fleet_pivot.csv');
}

function exportTableCSV() {
    const table = document.getElementById('arrivalsTable');
    downloadTableCSV(table, 'tesla_fleet_arrivi.csv');
}

function downloadTableCSV(table, filename) {
    const rows = [];
    const headerCells = table.querySelectorAll('thead th');
    rows.push(Array.from(headerCells).map(c => '"' + c.textContent.replace(/[↕]/g, '').trim() + '"').join(','));

    const bodyRows = table.querySelectorAll('tbody tr');
    bodyRows.forEach(tr => {
        const cells = tr.querySelectorAll('td');
        rows.push(Array.from(cells).map(c => '"' + c.textContent.trim() + '"').join(','));
    });

    const csvContent = '\uFEFF' + rows.join('\n'); // BOM for Excel
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
}

// ─── Demo Data ──────────────────────────────────────────────
function loadDemoData() {
    const locations = ['Milano Hub', 'Roma Compound', 'Napoli Porto', 'Torino Terminal', 'Genova Porto', 'Civitavecchia', 'Livorno Hub', 'Venezia Terminal'];
    const models = ['Model 3', 'Model Y', 'Model S', 'Model X', 'Cybertruck', 'Semi'];
    const statuses = ['In Transito', 'Programmato', 'Arrivato', 'In Ritardo', 'In Attesa'];

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const demoRows = [];
    for (let i = 0; i < 250; i++) {
        const daysOffset = Math.floor(Math.random() * 40) - 7; // -7 to +33
        const arrivalDate = new Date(today);
        arrivalDate.setDate(arrivalDate.getDate() + daysOffset);

        let status;
        if (daysOffset < -1) status = 'Arrivato';
        else if (daysOffset < 0) status = Math.random() > 0.5 ? 'Arrivato' : 'In Ritardo';
        else if (daysOffset === 0) status = Math.random() > 0.3 ? 'In Transito' : 'Arrivato';
        else if (daysOffset < 5) status = Math.random() > 0.3 ? 'In Transito' : 'Programmato';
        else status = Math.random() > 0.2 ? 'Programmato' : 'In Attesa';

        const rn = 'RN' + String(200000000 + Math.floor(Math.random() * 99999999));

        // Delivery date = arrival + 0 to 3 days
        const deliveryDate = new Date(arrivalDate);
        deliveryDate.setDate(deliveryDate.getDate() + Math.floor(Math.random() * 4));

        demoRows.push({
            'Order ID': 'TFV-' + String(10000 + i),
            'Reservation Number': rn,
            'WDOCheckoutLink': 'https://dro.tesla.com/advisor',
            'Location': locations[Math.floor(Math.random() * locations.length)],
            'Data Arrivo': arrivalDate,
            'Data Consegna': deliveryDate,
            'Modello': models[Math.floor(Math.random() * models.length)],
            'Stato': status,
        });
    }

    allHeaders = ['Order ID', 'Reservation Number', 'WDOCheckoutLink', 'Location', 'Data Arrivo', 'Data Consegna', 'Modello', 'Stato'];
    rawData = demoRows;
    columnMap = {
        orderId: 'Order ID',
        reservationNumber: 'Reservation Number',
        wdoCheckoutLink: 'WDOCheckoutLink',
        location: 'Location',
        date: 'Data Arrivo',
        deliveryDate: 'Data Consegna',
        model: 'Modello',
        status: 'Stato',
    };

    processAndRender();
}

// ─── Utilities ──────────────────────────────────────────────
function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

function showUpload() {
    // Go back to boot screen
    document.getElementById('dashboard').style.display = 'none';
    document.getElementById('bootScreen').style.display = 'flex';
    document.getElementById('bootScreen').classList.remove('booting');

    // Reset state
    rawData = [];
    filteredData = [];
    columnMap = {};
    bootFile1Loaded = false;
    bootFile2Loaded = false;

    // Reset boot UI
    document.getElementById('bootSlot1').classList.remove('loaded');
    document.getElementById('bootSlot2').classList.remove('loaded');
    document.getElementById('bootDesc1').textContent = 'Export da chart "Tutti gli ordini con telaio"';
    document.getElementById('bootDesc2').textContent = 'B2B/B2C, ETA2SC, Containment Hold';
    document.getElementById('bootAction1').innerHTML = '<svg width="20" height="20" viewBox="0 0 20 20" fill="none"><path d="M10 3v10M6 7l4-4 4 4" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M3 14v3h14v-3" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg><span>CARICA EXCEL</span>';
    document.getElementById('bootAction2').innerHTML = '<svg width="20" height="20" viewBox="0 0 20 20" fill="none"><path d="M10 3v10M6 7l4-4 4 4" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M3 14v3h14v-3" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg><span>OPZIONALE</span>';
    document.getElementById('bootStartBtn').disabled = true;
    document.getElementById('bootStartBtn').querySelector('.boot-start-text').textContent = 'IN ATTESA DATI...';
    document.getElementById('bootInsertText').textContent = 'Inserisci i file dati per iniziare';
    document.getElementById('bootInsertText').style.color = '';

    const statusEl = document.getElementById('dataStatus');
    statusEl.className = 'header-status';
    statusEl.innerHTML = '<span class="status-dot"></span><span>In attesa dati</span>';
}

// ─── EASTER EGG: Battaglia Navale Tesla ─────────────────────
const GAME = {
    rows: 8,
    cols: 8,
    ships: [],
    board: [],
    revealed: [],
    shots: 0,
    hits: 0,
    sunk: 0,
    totalShipCells: 0,
};

const GAME_ROWS = ['Milano', 'Roma', 'Napoli', 'Torino', 'Genova', 'Livorno', 'Venezia', 'Bari'];
const GAME_COLS = ['M3', 'MY', 'MS', 'MX', 'CT', 'Semi', 'Rd', 'PW'];

const GAME_SHIPS = [
    { name: 'Bisarca Cybertruck', icon: '🚛', size: 4 },
    { name: 'Nave Model Y', icon: '🚢', size: 3 },
    { name: 'Nave Model 3', icon: '⛴️', size: 3 },
    { name: 'Furgone Model S', icon: '🚐', size: 2 },
    { name: 'Auto Model X', icon: '🚗', size: 2 },
];

function openGame() {
    document.getElementById('gameOverlay').style.display = 'flex';
    initGame();
}

function closeGame() {
    document.getElementById('gameOverlay').style.display = 'none';
}

function initGame() {
    GAME.board = Array.from({ length: GAME.rows }, () => Array(GAME.cols).fill(null));
    GAME.revealed = Array.from({ length: GAME.rows }, () => Array(GAME.cols).fill(false));
    GAME.shots = 0;
    GAME.hits = 0;
    GAME.sunk = 0;
    GAME.ships = [];
    GAME.totalShipCells = 0;

    // Place ships
    GAME_SHIPS.forEach(shipDef => {
        let placed = false;
        let attempts = 0;
        while (!placed && attempts < 200) {
            attempts++;
            const horizontal = Math.random() > 0.5;
            const r = Math.floor(Math.random() * (horizontal ? GAME.rows : GAME.rows - shipDef.size + 1));
            const c = Math.floor(Math.random() * (horizontal ? GAME.cols - shipDef.size + 1 : GAME.cols));

            let canPlace = true;
            const cells = [];
            for (let i = 0; i < shipDef.size; i++) {
                const cr = horizontal ? r : r + i;
                const cc = horizontal ? c + i : c;
                if (GAME.board[cr][cc] !== null) { canPlace = false; break; }
                cells.push([cr, cc]);
            }

            if (canPlace) {
                const ship = { ...shipDef, cells: cells, hitsOnShip: 0 };
                cells.forEach(([cr, cc]) => { GAME.board[cr][cc] = ship; });
                GAME.ships.push(ship);
                GAME.totalShipCells += shipDef.size;
                placed = true;
            }
        }
    });

    updateGameUI();
    renderGameGrid();
    setGameMessage('Clicca su una cella per sparare! Trova le 5 Tesla nascoste.', '');
}

function renderGameGrid() {
    const grid = document.getElementById('gameGrid');
    let html = '<tr><th></th>';
    GAME_COLS.forEach(c => { html += '<th>' + c + '</th>'; });
    html += '</tr>';

    for (let r = 0; r < GAME.rows; r++) {
        html += '<tr><th>' + GAME_ROWS[r] + '</th>';
        for (let c = 0; c < GAME.cols; c++) {
            const revealed = GAME.revealed[r][c];
            const ship = GAME.board[r][c];
            let cls = '';
            let content = '';

            if (revealed) {
                cls = 'revealed ';
                if (ship) {
                    if (ship.hitsOnShip >= ship.size) {
                        cls += 'game-sunk';
                        content = ship.icon;
                    } else {
                        cls += 'game-hit';
                        content = '💥';
                    }
                } else {
                    cls += 'game-miss';
                    content = '🌊';
                }
            }

            html += `<td class="${cls}" onclick="fireShot(${r},${c})">${content}</td>`;
        }
        html += '</tr>';
    }
    grid.innerHTML = html;
}

function fireShot(r, c) {
    if (GAME.revealed[r][c]) return;
    if (GAME.sunk >= GAME_SHIPS.length) return;

    GAME.revealed[r][c] = true;
    GAME.shots++;

    const ship = GAME.board[r][c];
    if (ship) {
        GAME.hits++;
        ship.hitsOnShip++;

        if (ship.hitsOnShip >= ship.size) {
            GAME.sunk++;
            // Reveal all cells of sunk ship
            ship.cells.forEach(([sr, sc]) => { GAME.revealed[sr][sc] = true; });

            if (GAME.sunk >= GAME_SHIPS.length) {
                const score = Math.max(0, 1000 - (GAME.shots - GAME.totalShipCells) * 50);
                setGameMessage(`HAI VINTO! Tutte le Tesla trovate in ${GAME.shots} colpi! Punteggio: ${score}`, 'win');
                document.getElementById('gameScore').textContent = score;
            } else {
                setGameMessage(`${ship.icon} ${ship.name} AFFONDATA! (${GAME.sunk}/5)`, 'sunk');
            }
        } else {
            setGameMessage(`💥 COLPITO! Hai beccato qualcosa a ${GAME_ROWS[r]}...`, 'hit');
        }
    } else {
        setGameMessage(`🌊 Acqua a ${GAME_ROWS[r]} / ${GAME_COLS[c]}`, 'miss');
    }

    updateGameUI();
    renderGameGrid();
}

function updateGameUI() {
    document.getElementById('gameShots').textContent = GAME.shots;
    document.getElementById('gameHits').textContent = GAME.hits;
    document.getElementById('gameFound').textContent = GAME.sunk;
}

function setGameMessage(msg, type) {
    const el = document.getElementById('gameMessage');
    el.textContent = msg;
    el.className = 'game-message ' + (type || '');
}

/* ============================================================
   FLEET MAP — Overlay integrata nella dashboard
   Usa direttamente rawData/filteredData, zero localStorage.
   ============================================================ */

const MAP_HUBS = {
    'milano linate sc':[45.4408,9.2773],'milano linate':[45.4408,9.2773],'milano':[45.464,9.19],
    'roma magliana sc':[41.8417,12.4292],'roma magliana':[41.8417,12.4292],
    'roma salaria sc':[41.935,12.51],'roma salaria':[41.935,12.51],'roma':[41.9028,12.4964],
    'napoli afragola sc':[40.9211,14.3117],'napoli afragola':[40.9211,14.3117],'napoli':[40.8518,14.2681],
    'torino moncalieri sc':[44.9931,7.6828],'torino moncalieri':[44.9931,7.6828],'torino':[45.07,7.687],
    'genova sc':[44.4056,8.9463],'genova':[44.4056,8.9463],
    'bologna sc':[44.4949,11.3426],'bologna':[44.4949,11.3426],
    'firenze sc':[43.7696,11.2558],'firenze':[43.7696,11.2558],'florence':[43.7696,11.2558],
    'padova sc':[45.4064,11.8768],'padova':[45.4064,11.8768],
    'brescia sc':[45.5416,10.2118],'brescia':[45.5416,10.2118],
    'catania sc':[37.5079,15.0934],'catania':[37.5079,15.0934],
    'palermo sc':[38.1157,13.3615],'palermo':[38.1157,13.3615],
    'bari sc':[41.1171,16.8719],'bari':[41.1171,16.8719],
    'verona sc':[45.4384,10.9916],'verona':[45.4384,10.9916],
    'livorno hub':[43.5485,10.3106],'livorno':[43.5485,10.3106],
    'venezia terminal':[45.4408,12.3155],'venezia':[45.4408,12.3155],
    'civitavecchia':[42.093,11.7968],
    'campania':[40.85,14.27],'marcianise':[41.038,14.296],
    'lombardia':[45.47,9.19],'vittuone':[45.49,8.95],
    'veneto':[45.44,11.00],'lazio':[41.90,12.50],
    // Transit Europa
    'neuss compound':[51.2,6.683],'neuss':[51.2,6.683],
    'zeebrugge port':[51.333,3.183],'zeebrugge':[51.333,3.183],
    'tilburg hub':[51.556,5.092],'tilburg':[51.556,5.092],
    'ghent terminal':[51.054,3.717],'ghent':[51.054,3.717],
    'gioia tauro porto':[38.424,15.899],'gioia tauro':[38.424,15.899],
    'genova porto':[44.41,8.92],'livorno porto':[43.55,10.3],
    'milano compound':[45.45,9.28],'roma compound':[41.85,12.43],
    'at sc':[42.5,12.5], // generico Italia centro
};

function _findCoords(name) {
    if (!name) return null;
    const low = name.toLowerCase().trim();
    if (MAP_HUBS[low]) return MAP_HUBS[low];
    // Partial match
    for (const [key, coords] of Object.entries(MAP_HUBS)) {
        if (low.includes(key) || key.includes(low)) return coords;
    }
    // Extract city from complex names like "EU-IT-Veneto-Verona-31-Viale delle Nazioni"
    const parts = low.split(/[-_,]/);
    for (const part of parts) {
        const p = part.trim();
        if (p.length >= 4 && MAP_HUBS[p]) return MAP_HUBS[p];
    }
    return null;
}

function _isCarAtHub(r) {
    const pos = (r.lastKnownLocation || '').trim().toLowerCase();
    const hub = (r.location || '').trim().toLowerCase();
    if (pos && hub) {
        if (pos === hub || pos.includes(hub) || hub.includes(pos)) return true;
        const hubCity = hub.replace(/\s*(sc|service center|hub|compound|terminal|porto)\s*/gi, '').trim();
        if (hubCity.length >= 3 && pos.includes(hubCity)) return true;
        if (pos === 'at sc' || pos.includes('at service center')) return true;
        return false;
    }
    if (!r.date) return false;
    const today = new Date(); today.setHours(0,0,0,0);
    const arrival = new Date(r.date); arrival.setHours(0,0,0,0);
    return arrival <= today;
}

let _fleetMap = null;
let _mapMarkerLayer = null;

function openMap() {
    const overlay = document.getElementById('mapOverlay');
    overlay.style.display = 'flex';

    if (!_fleetMap) {
        _fleetMap = L.map('fleetMap', { center: [42.5, 12.5], zoom: 6 });
        L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png', {
            attribution: '&copy; OSM &copy; CARTO',
            subdomains: 'abcd', maxZoom: 19
        }).addTo(_fleetMap);
        _mapMarkerLayer = L.layerGroup().addTo(_fleetMap);
    }

    setTimeout(() => { _fleetMap.invalidateSize(); updateMapMarkers(); }, 100);
}

function closeMap() {
    document.getElementById('mapOverlay').style.display = 'none';
}

function updateMapMarkers() {
    if (!_fleetMap || !_mapMarkerLayer) return;
    _mapMarkerLayer.clearLayers();

    const filter = document.getElementById('mapFilter').value;
    const today = new Date(); today.setHours(0,0,0,0);

    // Classify each order
    const orders = rawData.map(r => {
        const atHub = _isCarAtHub(r);
        const isCH = r.isContainmentHold || false;
        const payIncomplete = r.finalPaymentStatus && !(['COMPLETE','PAID','RECEIVED'].includes((r.finalPaymentStatus||'').toUpperCase()));
        return { ...r, _atHub: atHub, _isCH: isCH, _payIncomplete: payIncomplete };
    });

    // Apply filter
    let filtered;
    switch (filter) {
        case 'ground':  filtered = orders.filter(o => o._atHub); break;
        case 'transit':  filtered = orders.filter(o => !o._atHub); break;
        case 'ch':       filtered = orders.filter(o => o._isCH); break;
        case 'payment':  filtered = orders.filter(o => o._payIncomplete && o._atHub); break;
        default:         filtered = orders;
    }

    // Group by location (for hub markers)
    const hubGroups = {};
    // Group by lastKnownLocation (for transit markers)
    const posGroups = {};
    let totalGround = 0, totalTransit = 0;

    filtered.forEach(o => {
        const hub = o.location || 'N/A';
        if (!hubGroups[hub]) hubGroups[hub] = { orders: [], ground: 0, transit: 0, ch: 0, models: {} };
        hubGroups[hub].orders.push(o);
        hubGroups[hub].models[o.model] = (hubGroups[hub].models[o.model] || 0) + 1;
        if (o._isCH) hubGroups[hub].ch++;

        if (o._atHub) { hubGroups[hub].ground++; totalGround++; }
        else { hubGroups[hub].transit++; totalTransit++; }

        // Transit markers by position
        const pos = (o.lastKnownLocation || '').trim();
        if (pos && !o._atHub && pos.toLowerCase() !== 'in transit') {
            if (!posGroups[pos]) posGroups[pos] = [];
            posGroups[pos].push(o);
        }
    });

    // Stats
    document.getElementById('mapStatTotal').textContent = filtered.length + ' totali';
    document.getElementById('mapStatGround').textContent = totalGround + ' a terra';
    document.getElementById('mapStatTransit').textContent = totalTransit + ' in transito';

    const allCoords = [];

    // Hub markers
    Object.entries(hubGroups).forEach(([hub, data]) => {
        const coords = _findCoords(hub);
        if (!coords) return;
        allCoords.push(coords);

        const total = data.orders.length;
        const size = Math.max(26, Math.min(54, 20 + total * 0.6));
        const color = data.ch > 0 ? '#ef4444' : (data.ground > 0 ? '#22c55e' : '#3b82f6');

        const icon = L.divIcon({
            html: `<div class="map-hub-marker" style="width:${size}px;height:${size}px;background:${color};font-size:${size > 38 ? 13 : 10}px;box-shadow:0 0 12px ${color}80;">${total}</div>`,
            className: '', iconSize: [size, size], iconAnchor: [size/2, size/2],
        });

        const modelHtml = Object.entries(data.models).sort((a,b)=>b[1]-a[1]).slice(0,6)
            .map(([m,q]) => `<span class="popup-model-tag">${escapeHtml(m)}: ${q}</span>`).join('');

        const popup = `<div class="popup-hub">${escapeHtml(hub)}</div>
            <div class="popup-row"><span class="popup-lbl">Totale</span><span class="popup-val">${total}</span></div>
            <div class="popup-row"><span class="popup-lbl">A terra</span><span class="popup-val" style="color:#22c55e;">${data.ground}</span></div>
            <div class="popup-row"><span class="popup-lbl">In arrivo</span><span class="popup-val" style="color:#3b82f6;">${data.transit}</span></div>
            ${data.ch ? `<div class="popup-row"><span class="popup-lbl">CH</span><span class="popup-val" style="color:#ef4444;">${data.ch}</span></div>` : ''}
            <div class="popup-models">${modelHtml}</div>`;

        L.marker(coords, { icon }).addTo(_mapMarkerLayer).bindPopup(popup, { maxWidth: 280 });
    });

    // Transit markers
    Object.entries(posGroups).forEach(([pos, vehicles]) => {
        const coords = _findCoords(pos);
        if (!coords) return;
        allCoords.push(coords);

        const size = Math.max(20, Math.min(38, 16 + vehicles.length));
        const icon = L.divIcon({
            html: `<div class="map-hub-marker" style="width:${size}px;height:${size}px;background:#eab308;font-size:${size > 28 ? 11 : 9}px;opacity:0.85;">${vehicles.length}</div>`,
            className: '', iconSize: [size, size], iconAnchor: [size/2, size/2],
        });

        L.marker(coords, { icon }).addTo(_mapMarkerLayer)
            .bindPopup(`<div class="popup-hub" style="color:#eab308;">${escapeHtml(pos.toUpperCase())}</div>
                <div class="popup-row"><span class="popup-lbl">In transito</span><span class="popup-val">${vehicles.length}</span></div>`);
    });

    // Fit bounds
    if (allCoords.length > 1) _fleetMap.fitBounds(allCoords, { padding: [30, 30] });
    else if (allCoords.length === 1) _fleetMap.setView(allCoords[0], 8);

    // Update VIN panel with non-a-terra vehicles
    _updateVinPanel(orders);
}

// ─── VIN Panel: auto non a terra ────────────────────────────
let _vinPanelCollapsed = false;
let _currentNotAtHub = [];

function _updateVinPanel(orders) {
    const notAtHub = orders.filter(o => !o._atHub).sort((a, b) => {
        // Sort by position then by hub
        const posA = (a.lastKnownLocation || 'zzz').toLowerCase();
        const posB = (b.lastKnownLocation || 'zzz').toLowerCase();
        if (posA !== posB) return posA.localeCompare(posB);
        return (a.location || '').localeCompare(b.location || '');
    });

    _currentNotAtHub = notAtHub;
    document.getElementById('mapVinCount').textContent = notAtHub.length;

    const list = document.getElementById('mapVinList');
    if (notAtHub.length === 0) {
        list.innerHTML = '<div style="padding:10px;color:#5a7a9e;font-size:0.8rem;grid-column:1/-1;">Tutte le auto sono a terra nei rispettivi hub</div>';
        return;
    }

    list.innerHTML = notAtHub.map(o => {
        const pos = o.lastKnownLocation || 'Posizione sconosciuta';
        const hub = o.location || '';
        return `<div class="map-vin-row">
            <span class="vin-code">${escapeHtml(o.orderId)}</span>
            <span class="vin-pos">${escapeHtml(pos)}</span>
            <span class="vin-hub">&rarr; ${escapeHtml(hub)}</span>
        </div>`;
    }).join('');
}

function toggleVinPanel() {
    _vinPanelCollapsed = !_vinPanelCollapsed;
    const body = document.getElementById('mapVinBody');
    const toggle = document.getElementById('mapVinToggle');
    body.classList.toggle('collapsed', _vinPanelCollapsed);
    toggle.classList.toggle('collapsed', _vinPanelCollapsed);
    // Resize map after panel toggle
    setTimeout(() => { if (_fleetMap) _fleetMap.invalidateSize(); }, 350);
}

function copyVinList() {
    if (_currentNotAtHub.length === 0) return;
    const vins = _currentNotAtHub.map(o => o.orderId).join('\n');
    navigator.clipboard.writeText(vins).then(() => {
        const btn = document.querySelector('.map-copy-btn');
        btn.classList.add('copied');
        btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M3 7l3 3 5-5" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg> Copiati!';
        setTimeout(() => {
            btn.classList.remove('copied');
            btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 14 14" fill="none"><rect x="4" y="4" width="9" height="9" rx="1.5" stroke="currentColor" stroke-width="1.2"/><path d="M10 4V2.5A1.5 1.5 0 008.5 1H2.5A1.5 1.5 0 001 2.5v6A1.5 1.5 0 002.5 10H4" stroke="currentColor" stroke-width="1.2"/></svg> Copia VIN';
        }, 2000);
    }).catch(() => {
        // Fallback for older browsers
        const ta = document.createElement('textarea');
        ta.value = vins;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        alert('VIN copiati: ' + _currentNotAtHub.length);
    });
}
