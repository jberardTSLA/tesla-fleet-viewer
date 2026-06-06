/* ============================================================
   Tesla Fleet Viewer — Advanced Graphics Engine
   GSAP Animations + Three.js 3D Background + Canvas Effects
   ============================================================ */

// ─── THREE.JS: 3D Particle Space Background ────────────────
(function initThreeBackground() {
    const canvas = document.getElementById('threeBg');
    if (!canvas || typeof THREE === 'undefined') return;

    const scene = new THREE.Scene();
    const camera = new THREE.PerspectiveCamera(75, window.innerWidth / window.innerHeight, 0.1, 1000);
    const renderer = new THREE.WebGLRenderer({ canvas: canvas, alpha: true, antialias: true });
    renderer.setSize(window.innerWidth, window.innerHeight);
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));

    // Particle system
    const particleCount = 1200;
    const geometry = new THREE.BufferGeometry();
    const positions = new Float32Array(particleCount * 3);
    const colors = new Float32Array(particleCount * 3);
    const sizes = new Float32Array(particleCount);

    const colorPalette = [
        [0.23, 0.51, 0.96],  // blue
        [0.02, 0.71, 0.83],  // cyan
        [0.13, 0.77, 0.37],  // green
        [0.66, 0.33, 0.97],  // purple
        [0.92, 0.70, 0.03],  // gold
        [0.90, 0.90, 0.95],  // white stars
    ];

    for (let i = 0; i < particleCount; i++) {
        positions[i * 3] = (Math.random() - 0.5) * 60;
        positions[i * 3 + 1] = (Math.random() - 0.5) * 60;
        positions[i * 3 + 2] = (Math.random() - 0.5) * 60;

        const c = colorPalette[Math.floor(Math.random() * colorPalette.length)];
        colors[i * 3] = c[0];
        colors[i * 3 + 1] = c[1];
        colors[i * 3 + 2] = c[2];

        sizes[i] = Math.random() * 3 + 0.5;
    }

    geometry.setAttribute('position', new THREE.BufferAttribute(positions, 3));
    geometry.setAttribute('color', new THREE.BufferAttribute(colors, 3));
    geometry.setAttribute('size', new THREE.BufferAttribute(sizes, 1));

    // Create circular particle texture
    const pCanvas = document.createElement('canvas');
    pCanvas.width = 32; pCanvas.height = 32;
    const pCtx = pCanvas.getContext('2d');
    const gradient = pCtx.createRadialGradient(16, 16, 0, 16, 16, 16);
    gradient.addColorStop(0, 'rgba(255,255,255,1)');
    gradient.addColorStop(0.3, 'rgba(255,255,255,0.6)');
    gradient.addColorStop(1, 'rgba(255,255,255,0)');
    pCtx.fillStyle = gradient;
    pCtx.fillRect(0, 0, 32, 32);
    const particleTexture = new THREE.CanvasTexture(pCanvas);

    const material = new THREE.PointsMaterial({
        size: 0.3,
        vertexColors: true,
        transparent: true,
        opacity: 0.7,
        blending: THREE.AdditiveBlending,
        sizeAttenuation: true,
        map: particleTexture,
        depthWrite: false,
    });

    const particles = new THREE.Points(geometry, material);
    scene.add(particles);

    // Nebula glow planes
    const nebulaGeo = new THREE.PlaneGeometry(15, 15);
    const nebulaMat1 = new THREE.MeshBasicMaterial({
        color: 0x3b82f6, transparent: true, opacity: 0.03, side: THREE.DoubleSide
    });
    const nebula1 = new THREE.Mesh(nebulaGeo, nebulaMat1);
    nebula1.position.set(-5, 2, -15);
    nebula1.rotation.z = 0.3;
    scene.add(nebula1);

    const nebulaMat2 = new THREE.MeshBasicMaterial({
        color: 0x06b6d4, transparent: true, opacity: 0.02, side: THREE.DoubleSide
    });
    const nebula2 = new THREE.Mesh(nebulaGeo, nebulaMat2);
    nebula2.position.set(6, -3, -12);
    nebula2.rotation.z = -0.5;
    scene.add(nebula2);

    camera.position.z = 15;

    // Mouse tracking
    let mouseX = 0, mouseY = 0;
    document.addEventListener('mousemove', (e) => {
        mouseX = (e.clientX / window.innerWidth - 0.5) * 2;
        mouseY = (e.clientY / window.innerHeight - 0.5) * 2;
    });

    // Resize handler
    window.addEventListener('resize', () => {
        camera.aspect = window.innerWidth / window.innerHeight;
        camera.updateProjectionMatrix();
        renderer.setSize(window.innerWidth, window.innerHeight);
    });

    // Animation loop
    let running = true;
    function animate() {
        if (!running) return;
        requestAnimationFrame(animate);

        // Rotate particles slowly
        particles.rotation.y += 0.0005;
        particles.rotation.x += 0.0002;

        // Mouse parallax
        camera.position.x += (mouseX * 2 - camera.position.x) * 0.02;
        camera.position.y += (-mouseY * 2 - camera.position.y) * 0.02;
        camera.lookAt(0, 0, 0);

        // Nebula float
        nebula1.rotation.z += 0.001;
        nebula2.rotation.z -= 0.0008;

        // Pulse particle sizes
        const time = Date.now() * 0.001;
        const sizeAttr = geometry.getAttribute('size');
        for (let i = 0; i < particleCount; i++) {
            sizeAttr.array[i] = sizes[i] * (0.8 + 0.4 * Math.sin(time + i * 0.1));
        }
        sizeAttr.needsUpdate = true;

        renderer.render(scene, camera);
    }
    animate();

    // Stop when boot screen is hidden
    const observer = new MutationObserver(() => {
        const boot = document.getElementById('bootScreen');
        if (boot && boot.style.display === 'none') {
            running = false;
            renderer.dispose();
        }
    });
    const bootEl = document.getElementById('bootScreen');
    if (bootEl) observer.observe(bootEl, { attributes: true, attributeFilter: ['style', 'class'] });
})();

// ─── GSAP: Boot Screen Entrance ─────────────────────────────
(function initBootGSAP() {
    if (typeof gsap === 'undefined') return;

    const tl = gsap.timeline({ delay: 0.3 });

    // Logo entrance
    tl.from('.boot-tesla-t', {
        scale: 0,
        rotation: 180,
        opacity: 0,
        duration: 1.2,
        ease: 'elastic.out(1, 0.5)',
    });

    // Title reveal
    tl.from('.boot-title', {
        y: 30,
        opacity: 0,
        duration: 0.8,
        ease: 'power3.out',
    }, '-=0.5');

    // Version
    tl.from('.boot-version', {
        y: 20,
        opacity: 0,
        duration: 0.5,
        ease: 'power2.out',
    }, '-=0.4');

    // Insert text
    tl.from('.boot-insert', {
        opacity: 0,
        duration: 0.6,
    }, '-=0.2');

    // Button
    tl.from('.boot-start-btn', {
        y: 20,
        opacity: 0,
        duration: 0.5,
        ease: 'power2.out',
        clearProps: 'all',
    }, '-=0.2');

    // Hint
    tl.from('.boot-hint', {
        opacity: 0,
        duration: 0.4,
    }, '-=0.2');

    // Glow pulse
    tl.from('.boot-glow', {
        scale: 0.3,
        opacity: 0,
        duration: 1.5,
        ease: 'power1.out',
    }, 0);
})();

// ─── GSAP: Optimus Selection Entrance ───────────────────────
window.gsapOptimus = function() {
    if (typeof gsap === 'undefined') return;

    const tl = gsap.timeline();

    // Header slide down
    tl.from('.cod-header', {
        y: -60,
        opacity: 0,
        duration: 0.6,
        ease: 'power3.out',
    });

    // Characters entrance - staggered from sides
    tl.from('.cod-char', {
        y: 100,
        opacity: 0,
        scale: 0.7,
        duration: 0.8,
        stagger: 0.12,
        ease: 'back.out(1.5)',
    }, '-=0.3');

    // Info bar slide up
    tl.from('.cod-infobar', {
        y: 40,
        opacity: 0,
        duration: 0.5,
        ease: 'power3.out',
    }, '-=0.3');

    // Footer
    tl.from('.cod-footer', {
        opacity: 0,
        duration: 0.4,
    }, '-=0.2');

    // Active character glow pulse
    tl.to('.cod-char.active svg', {
        filter: 'drop-shadow(0 0 20px rgba(59,130,246,0.4))',
        duration: 0.6,
        ease: 'power1.inOut',
    }, '-=0.3');
};

// ─── GSAP: Character Switch Animation (EPIC) ───────────────
window.gsapSwitchChar = function(idx) {
    if (typeof gsap === 'undefined') return;

    const chars = document.querySelectorAll('.cod-char');
    const colors = ['#3b82f6', '#22c55e', '#ef4444', '#a855f7'];
    const glowColors = ['rgba(59,130,246,0.5)', 'rgba(34,197,94,0.5)', 'rgba(239,68,68,0.5)', 'rgba(168,85,247,0.5)'];

    // Camera shake effect on the lineup
    const lineup = document.getElementById('codLineup');
    if (lineup) {
        gsap.fromTo(lineup, {x: -3}, {x: 3, duration: 0.05, repeat: 5, yoyo: true, ease: 'power1.inOut', onComplete: () => gsap.set(lineup, {x: 0})});
    }

    chars.forEach((ch, i) => {
        if (i === idx) {
            // Active: epic entrance
            const tl = gsap.timeline();

            tl.to(ch, {
                scale: 1.06,
                y: -12,
                opacity: 1,
                filter: 'brightness(1.1)',
                duration: 0.5,
                ease: 'back.out(1.7)',
            });

            // LED mega pulse with color flash
            const led = ch.querySelector('.optimus-led');
            if (led) {
                gsap.fromTo(led, { attr: { 'stroke-width': 2.5 } }, {
                    attr: { 'stroke-width': 6 },
                    duration: 0.2,
                    yoyo: true,
                    repeat: 2,
                    ease: 'power2.inOut',
                });
            }

            // SVG neon burst
            const svg = ch.querySelector('svg');
            if (svg) {
                gsap.fromTo(svg, {
                    filter: `drop-shadow(0 0 30px ${glowColors[i]})`
                }, {
                    filter: `drop-shadow(0 0 12px ${glowColors[i]})`,
                    duration: 0.8,
                    ease: 'power2.out',
                });
            }

        } else {
            gsap.to(ch, {
                scale: 0.82,
                y: 12,
                opacity: 0.35,
                filter: 'brightness(0.45)',
                duration: 0.4,
                ease: 'power2.out',
            });
        }
    });

    // Info bar text flash with energy wipe
    const infobar = document.querySelector('.cod-infobar');
    if (infobar) {
        gsap.fromTo(infobar, {
            borderColor: colors[idx],
            boxShadow: `0 0 30px ${glowColors[idx]}`
        }, {
            borderColor: 'rgba(234,179,8,0.25)',
            boxShadow: '0 0 0px transparent',
            duration: 1,
            ease: 'power2.out',
        });
    }

    gsap.fromTo('.cod-class-name', {opacity: 0, x: -20, scale: 0.9}, {
        opacity: 1, x: 0, scale: 1, duration: 0.4, ease: 'back.out(1.5)'
    });
    gsap.fromTo('.cod-class-desc', {opacity: 0, x: -10}, {
        opacity: 1, x: 0, duration: 0.3, delay: 0.1, ease: 'power2.out'
    });

    // Mini skills flash in
    gsap.fromTo('.cod-mini-ability', {opacity: 0, y: 10}, {
        opacity: 1, y: 0, duration: 0.3, stagger: 0.08, delay: 0.15, ease: 'power2.out'
    });
};

// ─── GSAP: Boot to Optimus Transition ───────────────────────
window.gsapBootToOptimus = function(callback) {
    if (typeof gsap === 'undefined') { callback(); return; }

    const tl = gsap.timeline({ onComplete: callback });

    // Flash white
    tl.to('.boot-content', {
        opacity: 0,
        scale: 0.95,
        duration: 0.4,
        ease: 'power2.in',
    });

    tl.to('.boot-glow', {
        scale: 3,
        opacity: 1,
        duration: 0.6,
        ease: 'power2.in',
    }, '-=0.3');

    tl.to('#bootScreen', {
        opacity: 0,
        duration: 0.4,
        ease: 'power1.in',
    }, '-=0.2');
};

// ─── GSAP: Deploy Transition (CINEMATIC) ────────────────────
window.gsapDeploy = function(callback) {
    if (typeof gsap === 'undefined') { callback(); return; }

    const tl = gsap.timeline({ onComplete: callback });
    const colors = ['#3b82f6', '#22c55e', '#ef4444', '#a855f7'];
    const activeIdx = selectedRoster || 0;
    const col = colors[activeIdx];

    // Deploy button flash
    tl.to('.cod-deploy-btn', {
        boxShadow: `0 0 60px ${col}, 0 0 120px ${col}`,
        scale: 1.1,
        duration: 0.3,
        ease: 'power2.in',
    });

    // Active character zooms forward with neon trail
    const active = document.querySelector('.cod-char.active');
    if (active) {
        tl.to(active, {
            scale: 2.0,
            opacity: 0,
            filter: `brightness(3) drop-shadow(0 0 40px ${col})`,
            duration: 0.7,
            ease: 'power3.in',
        }, '-=0.1');
    }

    // Other chars swept away
    tl.to('.cod-char:not(.active)', {
        opacity: 0,
        x: function(i) { return (i % 2 === 0 ? -80 : 80); },
        scale: 0.5,
        duration: 0.4,
        stagger: 0.05,
        ease: 'power3.in',
    }, '-=0.5');

    // UI elements cinematic fade
    tl.to('.cod-header', {
        opacity: 0,
        y: -30,
        duration: 0.3,
        ease: 'power2.in',
    }, '-=0.3');

    tl.to('.cod-infobar', {
        opacity: 0,
        y: 30,
        duration: 0.3,
        ease: 'power2.in',
    }, '-=0.25');

    tl.to('.cod-footer', {
        opacity: 0,
        duration: 0.2,
    }, '-=0.2');

    // Final screen fade with white flash
    tl.to('#optimusIntro', {
        opacity: 0,
        duration: 0.4,
        ease: 'power1.in',
    }, '-=0.1');
};

// ─── GSAP: Dashboard Reveal (EPIC ENTRANCE) ────────────────
window.gsapDashboardReveal = function() {
    if (typeof gsap === 'undefined') return;

    const dashboard = document.getElementById('dashboard');
    if (!dashboard || dashboard.style.display === 'none') return;

    // Ensure everything starts visible (CSS handles initial fadeIn)
    // We add extra GSAP flair on top

    // KPI cards: stagger with scale bounce and glow flash
    const kpiCards = dashboard.querySelectorAll('.kpi-card');
    if (kpiCards.length) {
        gsap.fromTo(kpiCards, {
            opacity: 0,
            y: 40,
            scale: 0.8,
        }, {
            opacity: 1,
            y: 0,
            scale: 1,
            duration: 0.6,
            stagger: 0.08,
            ease: 'back.out(1.4)',
            clearProps: 'transform,opacity',
            onComplete: function() {
                // Add glow pulse to each card after entrance
                kpiCards.forEach((card, i) => {
                    gsap.fromTo(card, {
                        boxShadow: '0 0 0px transparent',
                    }, {
                        boxShadow: '0 0 20px rgba(59,130,246,0.15)',
                        duration: 0.4,
                        delay: i * 0.05,
                        yoyo: true,
                        repeat: 1,
                        ease: 'power1.inOut',
                        clearProps: 'boxShadow',
                    });
                });
            }
        });
    }

    // KPI values: countup effect boost
    const kpiValues = dashboard.querySelectorAll('.kpi-value');
    if (kpiValues.length) {
        gsap.fromTo(kpiValues, {
            scale: 0.5,
            opacity: 0,
        }, {
            scale: 1,
            opacity: 1,
            duration: 0.5,
            stagger: 0.08,
            delay: 0.3,
            ease: 'elastic.out(1, 0.6)',
            clearProps: 'all',
        });
    }

    // Filter bar slide in
    const filterBar = dashboard.querySelector('.filter-bar');
    if (filterBar) {
        gsap.fromTo(filterBar, {
            opacity: 0,
            y: 20,
        }, {
            opacity: 1,
            y: 0,
            duration: 0.5,
            delay: 0.5,
            ease: 'power3.out',
            clearProps: 'all',
        });
    }

    // Chart cards: stagger from bottom with scale
    const chartCards = dashboard.querySelectorAll('.chart-card');
    if (chartCards.length) {
        gsap.fromTo(chartCards, {
            opacity: 0,
            y: 50,
            scale: 0.9,
        }, {
            opacity: 1,
            y: 0,
            scale: 1,
            duration: 0.7,
            stagger: 0.1,
            delay: 0.6,
            ease: 'power3.out',
            clearProps: 'all',
        });
    }

    // Table cards: slide in
    const tableCards = dashboard.querySelectorAll('.table-card');
    if (tableCards.length) {
        gsap.fromTo(tableCards, {
            opacity: 0,
            y: 30,
        }, {
            opacity: 1,
            y: 0,
            duration: 0.5,
            stagger: 0.08,
            delay: 0.9,
            ease: 'power2.out',
            clearProps: 'all',
        });
    }

    // Location cards: pop in
    const locationCards = dashboard.querySelectorAll('.location-card');
    if (locationCards.length) {
        gsap.fromTo(locationCards, {
            opacity: 0,
            scale: 0.85,
            y: 20,
        }, {
            opacity: 1,
            scale: 1,
            y: 0,
            duration: 0.5,
            stagger: 0.06,
            delay: 1.1,
            ease: 'back.out(1.2)',
            clearProps: 'all',
        });
    }
};

// ─── GSAP: File Slot Load Animation ─────────────────────────
window.gsapFileLoaded = function(slotEl) {
    if (typeof gsap === 'undefined') return;

    gsap.fromTo(slotEl, {
        borderColor: 'rgba(34,197,94,0.8)',
        boxShadow: '0 0 30px rgba(34,197,94,0.3)',
    }, {
        borderColor: 'rgba(34,197,94,0.4)',
        boxShadow: '0 0 0px rgba(34,197,94,0)',
        duration: 1,
        ease: 'power2.out',
    });

    const actionEl = slotEl.querySelector('.boot-slot-action');
    if (actionEl) gsap.from(actionEl, { scale: 1.5, duration: 0.4, ease: 'back.out(2)' });
};

// ─── THREE.JS: Optimus Selection Background (EPIC SPACE) ───
window.initThreeOptimus = function() {
    const canvas = document.getElementById('threeOptimus');
    if (!canvas || typeof THREE === 'undefined') return;
    const scene = new THREE.Scene();
    const camera = new THREE.PerspectiveCamera(60, window.innerWidth/window.innerHeight, 0.1, 1000);
    const renderer = new THREE.WebGLRenderer({ canvas, alpha:true, antialias:true });
    renderer.setSize(window.innerWidth, window.innerHeight);
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));

    // Particle texture
    const pc = document.createElement('canvas'); pc.width=64; pc.height=64;
    const pctx = pc.getContext('2d');
    const grad = pctx.createRadialGradient(32,32,0,32,32,32);
    grad.addColorStop(0,'rgba(255,255,255,1)');
    grad.addColorStop(0.2,'rgba(255,255,255,0.8)');
    grad.addColorStop(0.5,'rgba(255,255,255,0.3)');
    grad.addColorStop(1,'rgba(255,255,255,0)');
    pctx.fillStyle=grad; pctx.fillRect(0,0,64,64);
    const pTex = new THREE.CanvasTexture(pc);

    // ═══ STARFIELD ═══
    const starCount = 800;
    const starGeo = new THREE.BufferGeometry();
    const starPos = new Float32Array(starCount*3);
    const starCol = new Float32Array(starCount*3);
    const starSizes = new Float32Array(starCount);
    const pal=[[0.92,0.70,0.03],[0.98,0.45,0.09],[0.23,0.51,0.96],[0.02,0.71,0.83],[0.85,0.85,0.90],[0.95,0.95,1.0]];
    for(let i=0;i<starCount;i++){
        starPos[i*3]=(Math.random()-0.5)*80;
        starPos[i*3+1]=(Math.random()-0.5)*80;
        starPos[i*3+2]=(Math.random()-0.5)*80;
        const c=pal[Math.floor(Math.random()*pal.length)];
        starCol[i*3]=c[0]; starCol[i*3+1]=c[1]; starCol[i*3+2]=c[2];
        starSizes[i] = Math.random() * 2 + 0.3;
    }
    starGeo.setAttribute('position',new THREE.BufferAttribute(starPos,3));
    starGeo.setAttribute('color',new THREE.BufferAttribute(starCol,3));
    starGeo.setAttribute('size',new THREE.BufferAttribute(starSizes,1));
    const starMat=new THREE.PointsMaterial({size:0.3,vertexColors:true,transparent:true,opacity:0.6,blending:THREE.AdditiveBlending,sizeAttenuation:true,map:pTex,depthWrite:false});
    const stars=new THREE.Points(starGeo,starMat);
    scene.add(stars);

    // ═══ GROUND GRID (Tron-style) ═══
    const gridSize = 60;
    const gridDiv = 30;
    const gridGeo = new THREE.BufferGeometry();
    const gridVerts = [];
    const halfGrid = gridSize / 2;
    const step = gridSize / gridDiv;
    for (let i = 0; i <= gridDiv; i++) {
        const pos = -halfGrid + i * step;
        gridVerts.push(-halfGrid, 0, pos, halfGrid, 0, pos);
        gridVerts.push(pos, 0, -halfGrid, pos, 0, halfGrid);
    }
    gridGeo.setAttribute('position', new THREE.BufferAttribute(new Float32Array(gridVerts), 3));
    const gridMat = new THREE.LineBasicMaterial({ color: 0x1a3a6a, transparent: true, opacity: 0.15 });
    const grid = new THREE.LineSegments(gridGeo, gridMat);
    grid.position.y = -12;
    grid.rotation.x = 0;
    scene.add(grid);

    // ═══ ORBITAL RINGS ═══
    const ringColors = [0x3b82f6, 0xeab308, 0x06b6d4];
    const rings = [];
    ringColors.forEach((col, idx) => {
        const ringGeo = new THREE.RingGeometry(14 + idx * 5, 14.15 + idx * 5, 64);
        const ringMat = new THREE.MeshBasicMaterial({ color: col, transparent: true, opacity: 0.08, side: THREE.DoubleSide });
        const ring = new THREE.Mesh(ringGeo, ringMat);
        ring.rotation.x = Math.PI / 2.5 + idx * 0.15;
        ring.rotation.y = idx * 0.5;
        ring.position.y = -2 + idx;
        scene.add(ring);
        rings.push(ring);
    });

    // ═══ VOLUMETRIC NEBULAE ═══
    const nebulaGeo = new THREE.PlaneGeometry(25, 25);
    const nebulae = [];
    const nebulaColors = [0x3b82f6, 0x06b6d4, 0xeab308, 0xa855f7];
    const nebulaPositions = [
        [-10, 5, -25], [12, -4, -20], [0, 8, -18], [-8, -6, -22]
    ];
    nebulaColors.forEach((col, idx) => {
        const mat = new THREE.MeshBasicMaterial({
            color: col, transparent: true, opacity: 0.025, side: THREE.DoubleSide
        });
        const mesh = new THREE.Mesh(nebulaGeo, mat);
        const p = nebulaPositions[idx];
        mesh.position.set(p[0], p[1], p[2]);
        mesh.rotation.z = Math.random() * Math.PI;
        scene.add(mesh);
        nebulae.push(mesh);
    });

    // ═══ ENERGY PARTICLES (floating upward) ═══
    const energyCount = 150;
    const energyGeo = new THREE.BufferGeometry();
    const energyPos = new Float32Array(energyCount * 3);
    const energyCol = new Float32Array(energyCount * 3);
    const energySpeeds = new Float32Array(energyCount);
    for (let i = 0; i < energyCount; i++) {
        energyPos[i*3] = (Math.random()-0.5) * 30;
        energyPos[i*3+1] = (Math.random()-0.5) * 30 - 10;
        energyPos[i*3+2] = (Math.random()-0.5) * 20 - 5;
        const c = pal[Math.floor(Math.random() * 3)]; // gold/orange/blue only
        energyCol[i*3]=c[0]; energyCol[i*3+1]=c[1]; energyCol[i*3+2]=c[2];
        energySpeeds[i] = 0.01 + Math.random() * 0.03;
    }
    energyGeo.setAttribute('position', new THREE.BufferAttribute(energyPos, 3));
    energyGeo.setAttribute('color', new THREE.BufferAttribute(energyCol, 3));
    const energyMat = new THREE.PointsMaterial({ size: 0.15, vertexColors: true, transparent: true, opacity: 0.5, blending: THREE.AdditiveBlending, map: pTex, depthWrite: false });
    const energyParticles = new THREE.Points(energyGeo, energyMat);
    scene.add(energyParticles);

    camera.position.set(0, 2, 22);
    camera.lookAt(0, 0, 0);

    let mx=0, my=0;
    document.addEventListener('mousemove',e=>{mx=(e.clientX/window.innerWidth-0.5)*2;my=(e.clientY/window.innerHeight-0.5)*2;});
    window.addEventListener('resize',()=>{camera.aspect=window.innerWidth/window.innerHeight;camera.updateProjectionMatrix();renderer.setSize(window.innerWidth,window.innerHeight);});

    let running=true;
    const clock = new THREE.Clock();

    function anim(){
        if(!running)return;
        requestAnimationFrame(anim);
        const t = clock.getElapsedTime();

        // Stars rotation
        stars.rotation.y += 0.0002;
        stars.rotation.x += 0.0001;

        // Star size pulsing
        const sizeAttr = starGeo.getAttribute('size');
        for (let i = 0; i < starCount; i++) {
            sizeAttr.array[i] = starSizes[i] * (0.7 + 0.5 * Math.sin(t * 0.8 + i * 0.05));
        }
        sizeAttr.needsUpdate = true;

        // Grid wave effect
        const gridPos = gridGeo.getAttribute('position');
        for (let i = 0; i < gridPos.count; i++) {
            const x = gridPos.getX(i);
            const z = gridPos.getZ(i);
            gridPos.setY(i, Math.sin(x * 0.3 + t * 0.5) * Math.cos(z * 0.3 + t * 0.3) * 0.3 - 12);
        }
        gridPos.needsUpdate = true;

        // Rings orbit
        rings.forEach((ring, idx) => {
            ring.rotation.z = t * (0.05 + idx * 0.02);
            ring.material.opacity = 0.06 + 0.04 * Math.sin(t + idx);
        });

        // Nebulae drift
        nebulae.forEach((neb, idx) => {
            neb.rotation.z += 0.001 * (idx % 2 === 0 ? 1 : -1);
            neb.material.opacity = 0.02 + 0.015 * Math.sin(t * 0.5 + idx * 1.5);
        });

        // Energy particles float up
        const ePos = energyGeo.getAttribute('position');
        for (let i = 0; i < energyCount; i++) {
            let y = ePos.getY(i) + energySpeeds[i];
            if (y > 15) y = -15;
            ePos.setY(i, y);
            ePos.setX(i, ePos.getX(i) + Math.sin(t + i) * 0.002);
        }
        ePos.needsUpdate = true;

        // Camera parallax
        camera.position.x += (mx * 2.5 - camera.position.x) * 0.015;
        camera.position.y += (-my * 1.5 + 2 - camera.position.y) * 0.015;
        camera.lookAt(0, 0, 0);

        renderer.render(scene, camera);
    }
    anim();

    const obs=new MutationObserver(()=>{const el=document.getElementById('optimusIntro');if(el&&(el.style.display==='none'||el.style.opacity==='0')){running=false;renderer.dispose();obs.disconnect();}});
    const optEl=document.getElementById('optimusIntro');if(optEl)obs.observe(optEl,{attributes:true,attributeFilter:['style']});
};

// ─── GSAP: Enhanced Optimus Selection ───────────────────────
window.gsapOptimus = function() {
    if (typeof gsap === 'undefined') return;
    if (window.initThreeOptimus) initThreeOptimus();
    const tl = gsap.timeline();
    tl.fromTo('.cod-header',{y:-40,opacity:0},{y:0,opacity:1,duration:0.5,ease:'power3.out',clearProps:'all'});
    tl.fromTo('.cod-char',{y:80,opacity:0},{y:0,opacity:1,duration:0.6,stagger:0.1,ease:'back.out(1.3)',clearProps:'transform,opacity',onComplete:function(){document.querySelectorAll('.cod-char').forEach(c=>{c.style.opacity='';c.style.transform='';});}},'-=0.2');
    tl.fromTo('.cod-infobar',{y:30,opacity:0},{y:0,opacity:1,duration:0.4,ease:'power3.out',clearProps:'all'},'-=0.2');
    tl.fromTo('.cod-footer',{opacity:0},{opacity:1,duration:0.3,clearProps:'all'},'-=0.1');
    setTimeout(()=>{if(window.gsapSwitchChar)gsapSwitchChar(selectedRoster||0);},800);
};

// ─── Dashboard: Animated header energy system ──────────────
(function(){
    const header=document.querySelector('.header');
    if(!header)return;
    header.style.position='relative';

    // Create energy pulse dots
    for (let i = 0; i < 3; i++) {
        const dot = document.createElement('div');
        dot.style.cssText = `position:absolute;bottom:-3px;width:6px;height:6px;border-radius:50%;background:#06b6d4;box-shadow:0 0 10px #06b6d4,0 0 20px rgba(6,182,212,0.3);animation:hdrDotTravel ${5+i*2}s linear infinite;animation-delay:${i*1.5}s;pointer-events:none;z-index:2;`;
        header.appendChild(dot);
    }

    if(!document.getElementById('hdrGlowCss')){
        const s=document.createElement('style');
        s.id='hdrGlowCss';
        s.textContent=`
            @keyframes hdrDotTravel{0%{left:-6px;opacity:0}5%{opacity:1}95%{opacity:1}100%{left:100%;opacity:0}}
        `;
        document.head.appendChild(s);
    }
})();

// ─── SOUND FX ENGINE: Star Wars Synth (Web Audio API) ───────
// Zero file esterni — tutto sintetizzato con oscillatori + filtri.

const SFX = {
    _ctx: null,
    _enabled: true,
    _volume: 0.6,
    _unlocked: false,

    init() {
        if (this._ctx) {
            if (this._ctx.state === 'suspended') this._ctx.resume();
            return;
        }
        try {
            this._ctx = new (window.AudioContext || window.webkitAudioContext)();
            if (this._ctx.state === 'suspended') this._ctx.resume();
        } catch(e) { this._enabled = false; }
    },

    _gain(vol) {
        const g = this._ctx.createGain();
        g.gain.value = vol * this._volume;
        g.connect(this._ctx.destination);
        return g;
    },

    // R2-D2 BEEP — file caricato
    droidBeep() {
        if (!this._enabled) return; this.init();
        const ctx = this._ctx, now = ctx.currentTime;
        [[800,1800,1200,0,0.15],[1000,2200,1600,0.18,0.35],[1400,2400,null,0.38,0.5]].forEach(([f1,f2,f3,start,end],i) => {
            const o = ctx.createOscillator(); o.type = 'sine';
            const g = this._gain(0.4 - i*0.05);
            o.frequency.setValueAtTime(f1, now+start);
            o.frequency.exponentialRampToValueAtTime(f2, now+start+(end-start)*0.5);
            if (f3) o.frequency.exponentialRampToValueAtTime(f3, now+end);
            o.connect(g); o.start(now+start); o.stop(now+end);
        });
    },

    // LIGHTSABER IGNITE — selezione personaggio
    lightsaber() {
        if (!this._enabled) return; this.init();
        const ctx = this._ctx, now = ctx.currentTime;
        const o = ctx.createOscillator(); o.type = 'sawtooth';
        o.frequency.setValueAtTime(85, now);
        o.frequency.linearRampToValueAtTime(120, now+0.3);
        const f = ctx.createBiquadFilter(); f.type = 'lowpass'; f.frequency.value = 400; f.Q.value = 8;
        const g = this._gain(0);
        g.gain.linearRampToValueAtTime(0.3*this._volume, now+0.05);
        g.gain.linearRampToValueAtTime(0.12*this._volume, now+0.4);
        g.gain.linearRampToValueAtTime(0, now+0.55);
        o.connect(f); f.connect(g); o.start(now); o.stop(now+0.6);
        // Sweep
        const o2 = ctx.createOscillator(); o2.type = 'sine';
        o2.frequency.setValueAtTime(200, now);
        o2.frequency.exponentialRampToValueAtTime(800, now+0.12);
        o2.frequency.exponentialRampToValueAtTime(400, now+0.35);
        const g2 = this._gain(0.15);
        g2.gain.linearRampToValueAtTime(0, now+0.35);
        o2.connect(g2); o2.start(now); o2.stop(now+0.35);
    },

    // HYPERDRIVE — boot start / deploy
    hyperdrive() {
        if (!this._enabled) return; this.init();
        const ctx = this._ctx, now = ctx.currentTime;
        const o = ctx.createOscillator(); o.type = 'sawtooth';
        o.frequency.setValueAtTime(100, now);
        o.frequency.exponentialRampToValueAtTime(2000, now+0.5);
        o.frequency.exponentialRampToValueAtTime(80, now+0.85);
        const f = ctx.createBiquadFilter(); f.type = 'lowpass';
        f.frequency.setValueAtTime(300, now);
        f.frequency.exponentialRampToValueAtTime(4000, now+0.45);
        f.frequency.exponentialRampToValueAtTime(200, now+0.85);
        const g = this._gain(0);
        g.gain.linearRampToValueAtTime(0.25*this._volume, now+0.12);
        g.gain.linearRampToValueAtTime(0.35*this._volume, now+0.45);
        g.gain.linearRampToValueAtTime(0, now+0.85);
        o.connect(f); f.connect(g); o.start(now); o.stop(now+0.9);
        // Rumble
        const buf = ctx.createBuffer(1, ctx.sampleRate*0.7|0, ctx.sampleRate);
        const d = buf.getChannelData(0); for(let i=0;i<d.length;i++) d[i]=(Math.random()*2-1)*0.3;
        const n = ctx.createBufferSource(); n.buffer = buf;
        const nf = ctx.createBiquadFilter(); nf.type='lowpass'; nf.frequency.value=150;
        const gn = this._gain(0); gn.gain.linearRampToValueAtTime(0.15*this._volume,now+0.08);
        gn.gain.linearRampToValueAtTime(0,now+0.7);
        n.connect(nf); nf.connect(gn); n.start(now); n.stop(now+0.7);
    },

    // BLASTER — click
    blaster() {
        if (!this._enabled) return; this.init();
        const ctx = this._ctx, now = ctx.currentTime;
        const o = ctx.createOscillator(); o.type = 'square';
        o.frequency.setValueAtTime(1200, now);
        o.frequency.exponentialRampToValueAtTime(200, now+0.1);
        const g = this._gain(0.15);
        g.gain.exponentialRampToValueAtTime(0.001, now+0.1);
        o.connect(g); o.start(now); o.stop(now+0.11);
    },

    // IMPERIAL MARCH accento (G3 G3 G3 Eb3)
    imperialAccent() {
        if (!this._enabled) return; this.init();
        const ctx = this._ctx, now = ctx.currentTime;
        let t = now;
        [[196,0.12],[196,0.12],[196,0.12],[156,0.35]].forEach(([freq,dur]) => {
            const o = ctx.createOscillator(); o.type = 'square'; o.frequency.value = freq;
            const f = ctx.createBiquadFilter(); f.type='lowpass'; f.frequency.value=800;
            const g = this._gain(0.15);
            g.gain.setValueAtTime(0.15*this._volume, t);
            g.gain.linearRampToValueAtTime(0, t+dur*0.85);
            o.connect(f); f.connect(g); o.start(t); o.stop(t+dur);
            t += dur + 0.02;
        });
    },

    // FORCE THEME — note ascendenti (C4 E4 G4 C5)
    forceTheme() {
        if (!this._enabled) return; this.init();
        const ctx = this._ctx, now = ctx.currentTime;
        let t = now;
        [262,330,392,523].forEach((freq,i) => {
            const o = ctx.createOscillator(); o.type = 'sine'; o.frequency.value = freq;
            const g = this._gain(0.2-i*0.02);
            g.gain.setValueAtTime((0.2-i*0.02)*this._volume, t);
            g.gain.linearRampToValueAtTime(0, t+0.22);
            o.connect(g); o.start(t); o.stop(t+0.25);
            t += 0.11;
        });
    },

    // BLIP — piccolo suono UI
    blip() {
        if (!this._enabled) return; this.init();
        const ctx = this._ctx, now = ctx.currentTime;
        const o = ctx.createOscillator(); o.type = 'sine';
        o.frequency.setValueAtTime(1400, now);
        o.frequency.exponentialRampToValueAtTime(1800, now+0.04);
        const g = this._gain(0.12);
        g.gain.exponentialRampToValueAtTime(0.001, now+0.05);
        o.connect(g); o.start(now); o.stop(now+0.06);
    },
};

// ─── Hook suoni negli eventi ────────────────────────────────

// Sblocca audio al primo click (browser policy) + blip di conferma
document.addEventListener('DOMContentLoaded', () => {
    const unlock = (e) => {
        SFX.init();
        if (SFX._ctx) {
            SFX._ctx.resume().then(() => {
                SFX._unlocked = true;
                if (SFX._enabled) SFX.blip();
            });
        }
        document.removeEventListener('click', unlock);
        document.removeEventListener('touchstart', unlock);
    };
    document.addEventListener('click', unlock);
    document.addEventListener('touchstart', unlock);
});

// Toggle mute
window.toggleMute = function() {
    SFX._enabled = !SFX._enabled;
    const btn = document.getElementById('sfxMuteBtn');
    const iconOn = document.getElementById('sfxIconOn');
    const iconOff = document.getElementById('sfxIconOff');
    if (btn) btn.classList.toggle('muted', !SFX._enabled);
    if (iconOn) iconOn.style.display = SFX._enabled ? '' : 'none';
    if (iconOff) iconOff.style.display = SFX._enabled ? 'none' : '';
    if (SFX._enabled) SFX.blip();
};

// File caricato → R2D2
const _origFileLoaded = window.gsapFileLoaded;
window.gsapFileLoaded = function(s) { SFX.droidBeep(); if(_origFileLoaded)_origFileLoaded(s); };

// Boot start → Hyperdrive
const _origBoot2Opt = window.gsapBootToOptimus;
window.gsapBootToOptimus = function(cb) { SFX.hyperdrive(); if(_origBoot2Opt)_origBoot2Opt(cb); else if(cb)cb(); };

// Character switch → Lightsaber
const _origSwitch = window.gsapSwitchChar;
window.gsapSwitchChar = function(i) { SFX.lightsaber(); if(_origSwitch)_origSwitch(i); };

// Deploy → Imperial March
const _origDeploy = window.gsapDeploy;
window.gsapDeploy = function(cb) { SFX.imperialAccent(); if(_origDeploy)_origDeploy(cb); else if(cb)setTimeout(cb,800); };
