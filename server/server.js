/* ============================================================
   Tesla Fleet Viewer — Mock API Server
   Simula gli endpoint ZipLabs per sviluppo locale
   
   Endpoints:
     GET /api/orders     → File 1 (ordini con VIN, location, date, modello, stato)
     GET /api/enterprise → File 2 (B2B/B2C, ETA2SC, ContainmentHold, PaymentStatus)
     GET /api/health     → Health check
   
   Usage:  node server.js
   Port:   3000 (default) or PORT env var
   ============================================================ */

const http = require('http');
const url = require('url');

const PORT = process.env.PORT || 3000;

// ─── Hub italiani Tesla realistici ──────────────────────────
const HUBS = [
    'Milano Linate SC', 'Roma Magliana SC', 'Roma Salaria SC',
    'Napoli Afragola SC', 'Torino Moncalieri SC', 'Genova SC',
    'Bologna SC', 'Firenze SC', 'Padova SC', 'Brescia SC',
    'Catania SC', 'Palermo SC', 'Bari SC', 'Verona SC'
];

const MODELS = [
    'Model 3 Standard Range Plus', 'Model 3 Long Range AWD', 'Model 3 Performance',
    'Model Y Long Range AWD', 'Model Y Performance', 'Model Y Standard Range',
    'Model S Long Range', 'Model S Plaid',
    'Model X Long Range', 'Model X Plaid'
];

const STATUSES = [
    'Factory Gated', 'In Transit - Rail', 'In Transit - Ship',
    'In Transit - Truck', 'At Service Center', 'Delivered',
    'Ready for Customer', 'Transport Ordered'
];

const SPECIALISTS = [
    'Marco Rossi', 'Luca Bianchi', 'Elena Ferrari',
    'Giulia Conti', 'Andrea Colombo', 'Sara Rizzo',
    'Matteo Greco', 'Francesca Moretti', 'Davide Bruno',
    'Chiara Gallo', 'Alessandro Romano', 'Valentina Costa'
];

const LAST_LOCATIONS = [
    'Neuss Compound', 'Zeebrugge Port', 'Tilburg Hub',
    'Ghent Terminal', 'Livorno Porto', 'Genova Porto',
    'Gioia Tauro Porto', 'Milano Compound', 'Roma Compound',
    'At SC', 'In Transit'
];

// ─── Generazione dati deterministici ────────────────────────
function generateVIN() {
    const chars = 'ABCDEFGHJKLMNPRSTUVWXYZ0123456789';
    let vin = '5YJ';  // Tesla prefix
    const models = ['3', 'Y', 'S', 'X'];
    vin += models[Math.floor(Math.random() * models.length)];
    for (let i = 0; i < 13; i++) {
        vin += chars[Math.floor(Math.random() * chars.length)];
    }
    return vin.substring(0, 17);
}

function generateRN() {
    return 'RN' + String(100000000 + Math.floor(Math.random() * 899999999));
}

function randomDate(daysFromNow, rangeDays) {
    const d = new Date();
    d.setDate(d.getDate() + daysFromNow + Math.floor(Math.random() * rangeDays));
    return d;
}

function formatDate(d) {
    return d.toISOString().split('T')[0];
}

function pick(arr) {
    return arr[Math.floor(Math.random() * arr.length)];
}

// ─── Generate Orders (File 1 structure) ─────────────────────
function generateOrders(count) {
    const orders = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    for (let i = 0; i < count; i++) {
        const vin = generateVIN();
        const rn = generateRN();
        const hub = pick(HUBS);
        const model = pick(MODELS);
        
        // Arrival: -10 to +30 days from now
        const daysOffset = Math.floor(Math.random() * 41) - 10;
        const arrivalDate = new Date(today);
        arrivalDate.setDate(arrivalDate.getDate() + daysOffset);
        
        // Delivery: arrival + 0 to 5 days
        const deliveryDate = new Date(arrivalDate);
        deliveryDate.setDate(deliveryDate.getDate() + Math.floor(Math.random() * 6));
        
        // Status based on timing
        let status;
        if (daysOffset < -3) status = 'At Service Center';
        else if (daysOffset < 0) status = pick(['At Service Center', 'Ready for Customer', 'Delivered']);
        else if (daysOffset === 0) status = pick(['In Transit - Truck', 'At Service Center']);
        else if (daysOffset < 5) status = pick(['In Transit - Ship', 'In Transit - Rail', 'In Transit - Truck']);
        else if (daysOffset < 15) status = pick(['Factory Gated', 'Transport Ordered', 'In Transit - Ship']);
        else status = pick(['Factory Gated', 'Transport Ordered']);

        const specialist = pick(SPECIALISTS);

        orders.push({
            'VIN': vin,
            'ReferenceNumber': rn,
            'WDOCheckoutLink': rn,
            'ServiceCenterForPickup': hub,
            'VehicleETAToService': formatDate(arrivalDate),
            'ScheduledDeliveryDate': formatDate(deliveryDate),
            'ProductTrim': model,
            'ModelTrimName': model,
            'VehicleStatus': status,
            'DeliverySpecialist': specialist
        });
    }
    return orders;
}

// ─── Generate Enterprise Data (File 2 structure) ────────────
function generateEnterprise(orders) {
    return orders.map(order => {
        const isEnterprise = Math.random() < 0.25; // 25% B2B
        const isConfident = Math.random() < 0.7;
        const isScheduled = Math.random() < 0.5;
        const isCH = Math.random() < 0.08; // 8% containment hold
        
        // Payment status
        const payStatuses = ['COMPLETE', 'COMPLETE', 'COMPLETE', 'INCOMPLETE', 'PENDING', 'RECEIVED'];
        const paymentStatus = pick(payStatuses);
        
        // ETA2SC: arrival date + some noise
        const eta2sc = new Date(order['VehicleETAToService']);
        eta2sc.setDate(eta2sc.getDate() + Math.floor(Math.random() * 3) - 1);
        
        return {
            'ReferenceNumber': order['ReferenceNumber'],
            'ETA2SC': formatDate(eta2sc),
            'IsEnterpriseOrder': isEnterprise,
            'IsConfidentETA': isConfident,
            'IsScheduled': isScheduled,
            'FinalPaymentStatus': paymentStatus,
            'IsContainmenthold': isCH,
            'LastKnownVehicleLocation': pick(LAST_LOCATIONS)
        };
    });
}

// ─── Cache per consistenza tra chiamate ─────────────────────
let cachedOrders = null;
let cachedEnterprise = null;
let cacheTimestamp = 0;
const CACHE_TTL = 60 * 1000; // Rigenera dati ogni 60s

function refreshCache() {
    const now = Date.now();
    if (!cachedOrders || (now - cacheTimestamp) > CACHE_TTL) {
        const orderCount = 180 + Math.floor(Math.random() * 70); // 180-250 ordini
        cachedOrders = generateOrders(orderCount);
        cachedEnterprise = generateEnterprise(cachedOrders);
        cacheTimestamp = now;
        console.log(`[${new Date().toISOString()}] Cache refreshed: ${cachedOrders.length} orders generated`);
    }
}

// ─── HTTP Server ────────────────────────────────────────────
const server = http.createServer((req, res) => {
    const parsedUrl = url.parse(req.url, true);
    const path = parsedUrl.pathname;

    // CORS headers
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }

    // Optional auth check (placeholder for Bearer token)
    const authHeader = req.headers['authorization'] || '';
    if (process.env.API_TOKEN && authHeader !== `Bearer ${process.env.API_TOKEN}`) {
        // Only enforce if API_TOKEN env var is set
        if (process.env.API_TOKEN) {
            res.writeHead(401, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: 'Unauthorized', message: 'Bearer token required' }));
            return;
        }
    }

    refreshCache();

    switch (path) {
        case '/api/orders':
            console.log(`[${new Date().toISOString()}] GET /api/orders → ${cachedOrders.length} records`);
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({
                data: cachedOrders,
                meta: {
                    total: cachedOrders.length,
                    generated_at: new Date().toISOString(),
                    source: 'mock-server'
                }
            }));
            break;

        case '/api/enterprise':
            console.log(`[${new Date().toISOString()}] GET /api/enterprise → ${cachedEnterprise.length} records`);
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({
                data: cachedEnterprise,
                meta: {
                    total: cachedEnterprise.length,
                    generated_at: new Date().toISOString(),
                    source: 'mock-server'
                }
            }));
            break;

        case '/api/health':
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({
                status: 'ok',
                uptime: process.uptime(),
                orders_cached: cachedOrders ? cachedOrders.length : 0,
                cache_age_ms: Date.now() - cacheTimestamp
            }));
            break;

        default:
            res.writeHead(404, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({
                error: 'Not Found',
                endpoints: ['/api/orders', '/api/enterprise', '/api/health']
            }));
    }
});

server.listen(PORT, () => {
    console.log('');
    console.log('  ╔══════════════════════════════════════════════╗');
    console.log('  ║  Tesla Fleet Viewer — Mock API Server        ║');
    console.log(`  ║  Running on http://localhost:${PORT}             ║`);
    console.log('  ║                                              ║');
    console.log('  ║  Endpoints:                                  ║');
    console.log('  ║    GET /api/orders     → File 1 (ordini)     ║');
    console.log('  ║    GET /api/enterprise → File 2 (enterprise) ║');
    console.log('  ║    GET /api/health     → Health check        ║');
    console.log('  ║                                              ║');
    console.log('  ║  Set API_TOKEN env var to enable auth        ║');
    console.log('  ╚══════════════════════════════════════════════╝');
    console.log('');
});
