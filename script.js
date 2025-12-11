const EXCEL_FILE = 'DB.xlsx';
 
// --- MAPPATURA COLONNE ---
const COL_MAP_DB = {
    'PN': 'P/N', 
    'Status': 'Status', 
    'LTB': 'LTB', 
    'size': 'Size', 
    'Resolution': 'Resolution', 
    'Type': 'Touch Type', 
    'CPU': 'CPU', 
    'Memory': 'Memory supported', 
    'WorkEnv': 'WORK-ENVIROMENT',
    'Vertical': 'VERTICAL SECTOR', 
    'Comment': 'COMMENT', 
    'Series': 'PRODUCT SERIES',
    'ManualLink': 'DatasheetLink', 
    'DirectLink': 'DirectLink', 
    'ThreeDLink': '3DLINK', 
    'TechLink': 'TECHINCALLINKS'
};

const UI_LABELS = {
    'WorkEnv': 'Work Environment', 
    'Type': 'Touch Technology', 
    'Memory': 'Max Memory',
    'size': 'Screen Size', 
    'OS Support': 'Operating System', 
    'Mounting': 'Mounting Options',
    'Vertical': 'Vertical Sector',
    'Series': 'Product Series'
};

const EXCLUSIONS = ["n/a", "-", "na", "", "configuration submitted successfully!", "undefined", "0", "a"];
     
let INDISPENSABLE_FILTERS = []; 
let OPTIONAL_FILTERS_STATE = {}; 
let GLOBAL_DATA = [];
let DB2_STRUCTURE = {}; 
let ACTIVE_FILTERS = {}; 
let CURRENT_VISIBLE_DATA = [];
const HISTORY_KEY = 'mya_search_history';
let loadingTimeout;

document.addEventListener('DOMContentLoaded', initApp);

async function initApp() {
    injectStyles(); 

    const savedTheme = localStorage.getItem('theme') || 'dark';
    document.documentElement.setAttribute('data-theme', savedTheme);
    updateThemeIcon(savedTheme);

    const search = document.getElementById('pn-search');
    if(search) {
        search.addEventListener('focus', () => { if(search.value.trim() === '') showSearchHistory(); });
        search.addEventListener('input', (e) => {
            if(e.target.value.trim() === '') { showSearchHistory(); applyFilters(); } 
            else { showSuggestions(e.target.value); if(e.target.value.length === 0) applyFilters(); }
        });
        search.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                document.getElementById('search-suggestions-box').style.display = 'none';
                saveToHistory(search.value);
                applyFilters(true);
            }
        });
        document.addEventListener('click', (e) => { if(!e.target.closest('.search-wrapper')) document.getElementById('search-suggestions-box').style.display = 'none'; });
    }
     
    loadingTimeout = setTimeout(() => {
        const loading = document.getElementById('loading');
        if(loading && loading.style.display !== 'none') {
            showError(`
            <span class="error-title">⚠️ Data Loading Issue</span>
            It seems to be taking too long. Possible causes:<br><br>
            1. <b>DB.xlsx</b> is missing from the folder.<br>
            2. You are using <b>file://</b> protocol (Double-clicking HTML).<br>
            3. Please use a local server (like Live Server extension in VS Code).
        `);
        }
    }, 3000); 

    await loadDatabase();
}

function returnToHome() {
    document.getElementById('main-app-container').style.display = 'none';
    document.getElementById('landing-page').style.display = 'flex';
}

function selectCategory(mode) {
    if(mode === 'monitor') {
        INDISPENSABLE_FILTERS = ['size', 'Type', 'Mounting'];
        GLOBAL_DATA.sort((a, b) => {
            const pnA = String(a.PN).toUpperCase();
            const pnB = String(b.PN).toUpperCase();
            const isFPM_A = pnA.startsWith("FPM");
            const isFPM_B = pnB.startsWith("FPM");
            if (isFPM_A && !isFPM_B) return -1;
            if (!isFPM_A && isFPM_B) return 1; 
            return (a._id || 0) - (b._id || 0); 
        });
    } else {
        INDISPENSABLE_FILTERS = ['size', 'Type', 'Mounting', 'CPU', 'Memory'];
        GLOBAL_DATA.sort((a, b) => {
            const pnA = String(a.PN).toUpperCase();
            const pnB = String(b.PN).toUpperCase();
            const isFPM_A = pnA.startsWith("FPM");
            const isFPM_B = pnB.startsWith("FPM");
            if (!isFPM_A && isFPM_B) return -1;
            if (isFPM_A && !isFPM_B) return 1; 
            return (a._id || 0) - (b._id || 0);
        });
    }
    document.getElementById('landing-page').style.display = 'none';
    document.getElementById('main-app-container').style.display = 'block';
    document.getElementById('controls').style.display = 'block';
    buildFilters();
    applyFilters();
}

function injectStyles() {
    const style = document.createElement('style');
    style.innerHTML = `
        @keyframes bounceIcon { 0%, 100% { transform: translateY(0); } 50% { transform: translateY(4px); } }
        .btn-impact { 
            background: linear-gradient(135deg, var(--primary) 0%, var(--accent) 100%); 
            color: white; 
            border: none; 
            padding: 10px 24px; 
            border-radius: 8px; 
            font-weight: 700; 
            text-transform: uppercase; 
            letter-spacing: 0.5px; 
            text-decoration: none; 
            display: inline-flex; 
            align-items: center; 
            gap: 10px; 
            font-size: 13px; 
            transition: all 0.2s ease; 
            box-shadow: 0 4px 15px rgba(0, 112, 224, 0.25); 
        }
        .btn-impact:hover { 
            transform: translateY(-2px); 
            box-shadow: 0 6px 20px rgba(0, 112, 224, 0.4); 
        }
        .btn-impact svg { width: 18px; height: 18px; transition: transform 0.3s; } 
        .btn-impact:hover svg { animation: bounceIcon 0.8s infinite; }
         
        .btn-filter-manage { display: flex; align-items: center; gap: 10px; background: var(--bg-card); color: var(--text-muted); border: 2px solid var(--border); padding: 10px 24px; border-radius: 28px; font-size: 14px; font-weight: 700; cursor: pointer; transition: all 0.2s ease; box-shadow: 0 4px 12px rgba(0,0,0,0.08); white-space: nowrap; height: 44px; }
        .btn-filter-manage:hover { border-color: var(--primary); color: var(--primary); background: var(--hover-row); transform: translateY(-2px); box-shadow: 0 8px 20px rgba(0,0,0,0.15); }
        .filter-dropdown-menu { display: none; position: absolute; top: 120%; right: 0; width: 300px; background: var(--bg-card); border: 2px solid var(--border); border-radius: 14px; box-shadow: 0 20px 50px rgba(0,0,0,0.3); z-index: 1200; padding: 14px; animation: fadeIn 0.2s ease; }
        .filter-dropdown-menu.show { display: block; }
        .filter-dropdown-header { padding: 12px 16px; font-size: 12px; text-transform: uppercase; letter-spacing: 1.5px; color: var(--text-muted); font-weight: 800; border-bottom: 2px solid var(--border); margin-bottom: 10px; }
        .filter-dropdown-item { display: flex; align-items: center; gap: 12px; padding: 12px 16px; border-radius: 8px; cursor: pointer; transition: background 0.2s; color: var(--text-main); font-size: 14px; font-weight: 500; }
        .filter-dropdown-item:hover { background: var(--hover-row); }
        .filter-dropdown-item input { accent-color: var(--primary); width: 18px; height: 18px; cursor: pointer; }
        
        .modal-preview-card { 
            background: transparent !important; 
            border: none !important; 
            box-shadow: none !important; 
            padding: 0 !important; 
            display: flex; 
            flex-direction: column; 
            align-items: center; 
            justify-content: center; 
            height: auto; 
            min-height: 300px; 
            position: relative; 
            cursor: zoom-in; 
            transition: transform 0.4s cubic-bezier(0.25, 0.8, 0.25, 1); 
        }
        .modal-preview-card:hover { transform: translateY(-5px); }
        .modal-product-img { 
            max-width: 100%; 
            max-height: 350px; 
            object-fit: contain; 
            filter: none !important;
            mix-blend-mode: multiply;
            transition: transform 0.4s ease; 
        }
        .modal-preview-card:hover .modal-product-img { transform: scale(1.1); }
         
        .zoom-hint { position: absolute; bottom: 0; background: rgba(0, 112, 224, 0.9); color: white; padding: 6px 16px; border-radius: 18px; font-size: 12px; font-weight: 700; text-transform: uppercase; opacity: 0; transform: translateY(10px); transition: all 0.3s ease; box-shadow: 0 4px 10px rgba(0,0,0,0.2); }
        .modal-preview-card:hover .zoom-hint { opacity: 1; transform: translateY(-20px); }
        #lightbox-overlay { position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: rgba(0, 0, 0, 0.92); backdrop-filter: blur(5px); z-index: 10000; display: flex; align-items: center; justify-content: center; cursor: zoom-out; animation: fadeInLightbox 0.3s ease-out; }
        .lightbox-content img { max-width: 90vw; max-height: 90vh; filter: drop-shadow(0 0 80px rgba(0,0,0,0.5)); animation: scaleInLightbox 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275); }
        @keyframes fadeInLightbox { from { opacity: 0; } to { opacity: 1; } }
        @keyframes scaleInLightbox { from { transform: scale(0.8); opacity: 0; } to { transform: scale(1); opacity: 1; } }
         
        /* --- COMPATIBILITY SECTION STYLES --- */
        .comp-section {
            margin-top: 40px;
            border: 2px solid #38BDF8; 
            border-radius: 8px;
            overflow: hidden;
            font-family: 'Inter', sans-serif;
            background: #FFFFFF;
        }
        .comp-header {
            background: linear-gradient(135deg, #38BDF8 0%, #0070E0 100%); 
            color: #FFFFFF;
            padding: 12px 18px;
            font-weight: 800;
            font-size: 14px;
            text-transform: uppercase;
            letter-spacing: 1px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .comp-body {
            padding: 20px;
            background: #0F172A;
            color: #FFFFFF;
        }
        .comp-bar-container {
            margin-bottom: 20px;
        }
        .comp-score-label {
            display: flex; justify-content: space-between; margin-bottom: 8px; font-size: 13px; color: #94A3B8; font-weight: 600;
        }
        .comp-progress-bg {
            height: 12px; width: 100%; background: #334155; border-radius: 6px; overflow: hidden;
        }
        .comp-progress-fill {
            height: 100%; background: #10B981; border-radius: 6px; transition: width 1s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .comp-grid {
            display: grid;
            grid-template-columns: 1fr 1px 1fr;
            gap: 20px;
            align-items: center;
        }
        .comp-divider {
            width: 1px; height: 100%; background: #334155;
        }
        .comp-col {
            display: flex; flex-direction: column; gap: 4px;
        }
        .comp-label {
            font-size: 11px; 
            color: #94A3B8; 
            font-weight: 800; 
            text-transform: uppercase; 
            letter-spacing: 0.5px;
            margin-bottom: 5px;
            display: block;
        }
        .comp-pn {
            font-size: 20px; font-weight: 900; color: #FFFFFF; margin: 4px 0;
        }
        .internal-link-action {
            cursor: pointer;
            color: #38BDF8;
            border-bottom: 2px dashed #38BDF8;
            transition: all 0.2s;
            display: inline-block;
        }
        .internal-link-action:hover {
            color: #ffffff;
            border-bottom-style: solid;
            border-color: #ffffff;
        }
        .comp-specs {
            font-size: 12px; color: #94A3B8; line-height: 1.4;
        }
        .comp-highlight {
            color: #10B981; font-weight: 700;
        }
        .comp-ltb {
            font-size: 11px; color: #F87171; font-weight: 700; margin-top: 4px;
        }
        .comp-ltb.good { color: #10B981; }
    `;
    document.head.appendChild(style);
}

function openLightbox(url) { if(!url) return; const overlay = document.getElementById('lightbox-overlay'); const img = document.getElementById('lightbox-img'); img.src = url; overlay.style.display = 'flex'; }
function closeLightbox() { const overlay = document.getElementById('lightbox-overlay'); overlay.style.display = 'none'; document.getElementById('lightbox-img').src = ''; }

function parseDate(val) { 
    if(!val) return null; if(val instanceof Date) return val;
    if(typeof val === 'string' && val.includes('/')) {
        const parts = val.split('/');
        if(parts.length === 3) return new Date(parseInt(parts[2]), parseInt(parts[1])-1, parseInt(parts[0]));
    }
    return new Date(val); 
}
function formatDate(val) { 
    if(!val || val === '-') return "-"; 
    let d = parseDate(val); 
    if(d && !isNaN(d)) { 
        let day = String(d.getDate()).padStart(2,'0'); let mon = String(d.getMonth()+1).padStart(2,'0'); 
        return `${day}/${mon}/${d.getFullYear()}`; 
    } 
    return val; 
}

// --- FUNZIONE PER SEMPLIFICARE I VALORI ---
function simplifyValue(val, key) {
    if (!val) return ""; 
    const s = String(val).trim(); 
    const low = s.toLowerCase();

    if (key === 'size') return s.replace(/["']/g, '').trim();
    if (key === 'Type') {
        if (low.includes('combo') || low.includes('pcap') || low.includes('projected')) return 'Pcap'; 
        if (low.includes('res')) return 'Resistive';
        return s;
    }
    // WorkEnv Cleaning (Uniformed for color logic)
    if (key === 'WorkEnv') {
        if (low.includes('class 1') || low.includes('c1d2')) return 'Explosion / Oil & Gas';
        if (low.includes('outdoor') || low.includes('sunlight')) return 'Outdoor / Harsher';
        if (low.includes('general') || low.includes('basic')) return 'General Purpose';
        if (low.includes('embedded')) return 'Embedded';
        return s;
    }
     
    // Clean OS
    if (key.toLowerCase().includes('os')) {
        if (low.includes('win') || low.includes('ws7')) return 'Windows';
        if (low.includes('android')) return 'Android';
        if (low.includes('linux')) return 'Linux';
        return 'Other';
    }
    return s;
}

// --- FILTER LOGIC (Grouped) ---
function simplifyForFilter(val, key) {
    let s = simplifyValue(val, key);
    let low = s.toLowerCase();

    // CUSTOM CPU SORTING (Clean "Core i7" etc.)
    if (key === 'CPU') {
        if (low.includes('i7')) return 'Core i7';
        if (low.includes('i5')) return 'Core i5';
        if (low.includes('i3') || low === 'core i') return 'Core i3'; 
        if (low.includes('xeon')) return 'Xeon Server';
        if (low.includes('pentium')) return 'Pentium';
        if (low.includes('celeron')) return 'Celeron';
        if (low.includes('atom')) return 'Atom';
        // Note: Generic "RISC/ARM" grouping removed
    }
     
    // CUSTOM MEMORY GROUPING (Remove Slot Info for Filter)
    if (key === 'Memory') {
        // Regex to find Capacity (e.g. 16GB) and Type (e.g. DDR4)
        // Ignoring "1 slot", "2 slots", "1x", "2x"
        let capMatch = s.match(/(\d+\s?GB)/i);
        let typeMatch = s.match(/(DDR\w+)/i);
         
        if (capMatch) {
            let cap = capMatch[1].toUpperCase().replace(/\s/g, ''); // 16GB
            let type = typeMatch ? typeMatch[1].toUpperCase() : ''; // DDR4
            return (cap + ' ' + type).trim();
        }
    }
     
    // Raggruppamento Risoluzione
    if (key === 'Resolution') {
        if (low.includes('1920') && low.includes('1080')) return 'FHD (1920x1080)';
        if (low.includes('1024') && low.includes('768')) return 'XGA (1024x768)';
        if (low.includes('1280') && low.includes('800')) return 'WXGA (1280x800)';
        if (low.includes('1366')) return 'HD (1366x768)';
        if (low.includes('800') && low.includes('600')) return 'SVGA (800x600)';
    }

    return s;
}

// --- COLORI SEMANTICI (EXPANDED VARIETY - NO BACKGROUNDS) ---
function getSemanticStyle(val, type) {
    if (!val || val === '-' || val === '0') return 'color: var(--text-muted); opacity: 0.7; font-style: italic;';
    const v = String(val).toLowerCase();

    // 1. WORK ENVIRONMENT / MOUNTING (Detailed Palette)
    if (type === 'WorkEnv') {
        if (v.includes('explosion') || v.includes('oil') || v.includes('class 1')) return 'color: #DC2626;'; // Red
        if (v.includes('outdoor') || v.includes('sunlight') || v.includes('harsher')) return 'color: #EA580C;'; // Orange
        if (v.includes('embedded') || v.includes('open frame')) return 'color: #16A34A;'; // Green
         
        // New Varieties from user hints
        if (v.includes('bench') || v.includes('desktop')) return 'color: #7C3AED;'; // Violet
        if (v.includes('arm') || v.includes('mount')) return 'color: #DB2777;'; // Pink/Magenta
        if (v.includes('cabinet') || v.includes('rack')) return 'color: #D97706;'; // Amber
        if (v.includes('wall')) return 'color: #0891B2;'; // Cyan
        if (v.includes('control') || v.includes('desk')) return 'color: #4F46E5;'; // Indigo
         
        return 'color: #3B82F6;'; // Standard Blue
    }

    // 2. VERTICAL SECTOR (INDUSTRY PALETTE)
    if (type === 'Vertical') {
        if (v.includes('rail') || v.includes('transport') || v.includes('vehicle')) return 'color: #B45309;'; // Amber
        if (v.includes('medic') || v.includes('health')) return 'color: #BE185D;'; // Deep Pink
        if (v.includes('food') || v.includes('beverage')) return 'color: #047857;'; // Emerald
        if (v.includes('automat') || v.includes('factory')) return 'color: #0284C7;'; // Sky Blue
        if (v.includes('iot')) return 'color: #8B5CF6;'; // Purple
        if (v.includes('marine')) return 'color: #1E3A8A;'; // Deep Navy
        if (v.includes('energy') || v.includes('power')) return 'color: #F59E0B;'; // Yellow
        if (v.includes('retail') || v.includes('kiosk')) return 'color: #EC4899;'; // Pink
        return 'color: #475569;'; // Slate
    }

    // 3. PRODUCT SERIES (BRAND PALETTE)
    if (type === 'Series') {
        if (v.startsWith('ppc')) return 'color: #4338CA;'; // Indigo
        if (v.startsWith('fpm')) return 'color: #0F766E;'; // Teal
        if (v.startsWith('tpc')) return 'color: #C2410C;'; // Red-Orange
        if (v.startsWith('uno')) return 'color: #4B5563;'; // Dark Zinc
        if (v.startsWith('utc')) return 'color: #BE123C;'; // Rose
        if (v.startsWith('aim')) return 'color: #701A75;'; // Fuchsia
        if (v.startsWith('hit')) return 'color: #059669;'; // Green
        return 'color: var(--primary);';
    }

    return '';
}

function cleanExtractedUrl(url) {
    if (!url || typeof url !== 'string' || url.length < 3) return "";
    let clean = url.trim();
    const badWords = ['n/a', '-', '0', 'undefined'];
    if (badWords.includes(clean.toLowerCase())) return "";
    if (clean.toLowerCase() === 'datasheet' || clean.toLowerCase() === 'a') return "";
    clean = clean.replace(/^HTTPS\/\//i, "https://"); 
    clean = clean.replace(/^HTTP\/\//i, "http://");
    if (clean.startsWith('/')) return "https://www.advantech.com" + clean;
    if (!clean.match(/^https?:\/\//i)) {
        if(clean.includes('advantech.com') || clean.includes('www.') || clean.includes('.pdf') || clean.includes('.jpg') || clean.includes('.png')) {
            return "https://" + clean;
        }
    }
    return clean;
}

async function loadDatabase() {
    try {
        if (window.location.protocol === 'file:') throw new Error("CORS ERROR: Browsers block reading local files directly. Please use a Local Server.");
        if (typeof XLSX === 'undefined') throw new Error("Library Error: xlsx.bundle.js failed to load.");

        const response = await fetch(EXCEL_FILE);
        if (response.status === 404) throw new Error("File Not Found: DB.xlsx is missing.");
        if (!response.ok) throw new Error("Network Error: Cannot load DB.xlsx.");
        
        const buffer = await response.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
        
        if (!workbook.Sheets['DB'] || !workbook.Sheets['DB2']) throw new Error("Invalid Excel: Missing Sheets 'DB' or 'DB2'.");

        clearTimeout(loadingTimeout);

        const ws = workbook.Sheets['DB'];
        const range = XLSX.utils.decode_range(ws['!ref']);

        let headerRowIndex = 0;
        for (let r = range.s.r; r <= Math.min(range.e.r, 20); r++) {
            let rowHasPN = false;
            for(let c = range.s.c; c <= range.e.c; c++) {
                const cell = ws[XLSX.utils.encode_cell({r: r, c: c})];
                if(cell && cell.v && String(cell.v).toUpperCase().includes('P/N')) {
                    rowHasPN = true; headerRowIndex = r; break;
                }
            }
            if (rowHasPN) break;
        }
        
        const rawDB = XLSX.utils.sheet_to_json(workbook.Sheets['DB'], { range: headerRowIndex, defval: "", blankrows: true }); 

        const sheet2 = workbook.Sheets['DB2'];
        const range2 = XLSX.utils.decode_range(sheet2['!ref']);
        let categories = [], attributes = [];
        
        for (let C = range2.s.c; C <= range2.e.c; C++) {
            const catCell = sheet2[XLSX.utils.encode_cell({r: 0, c: C})];
            let catVal = catCell ? catCell.v : null;
            if (catVal) categories[C] = catVal.toString().trim(); else if (C > 0) categories[C] = categories[C-1];
            const attrCell = sheet2[XLSX.utils.encode_cell({r: 1, c: C})];
            let attrVal = attrCell ? attrCell.v : null;
            attributes[C] = attrVal ? attrVal.toString().trim() : null;
            if(categories[C] && attributes[C] && !categories[C].startsWith('__') && !attributes[C].startsWith('__')) {
                if(!DB2_STRUCTURE[categories[C]]) DB2_STRUCTURE[categories[C]] = [];
                DB2_STRUCTURE[categories[C]].push(attributes[C]);
            }
        }

        const rawDB2 = XLSX.utils.sheet_to_json(sheet2, { range: 1, defval: "" });
        const db2Map = {};
        
        rawDB2.forEach(row => {
            let pnKey = Object.keys(row).find(k => k.toUpperCase().includes('P/N'));
            let rawPn = row[pnKey];
            if(rawPn) {
                // FIXED: REMOVE ALL SPACES FROM P/N
                let pn = String(rawPn).replace(/\s/g, '').trim(); 
                db2Map[pn] = {};
                for (let key in row) {
                    let cellVal = String(row[key]).trim(); let cleanKey = key.trim();
                    if (cleanKey.startsWith('__')) continue; 
                    if(cellVal && !EXCLUSIONS.includes(cellVal.toLowerCase())) {
                        let foundCat = false;
                        for(let cat in DB2_STRUCTURE) {
                            if(DB2_STRUCTURE[cat].includes(cleanKey)) {
                                if(!db2Map[pn][cat]) db2Map[pn][cat] = [];
                                db2Map[pn][cat].push({ label: cleanKey, value: cellVal });
                                foundCat = true;
                            }
                        }
                        if(!foundCat) { if(!db2Map[pn][cleanKey]) db2Map[pn][cleanKey] = []; db2Map[pn][cleanKey].push({ label: cleanKey, value: cellVal }); }
                    } else if (cellVal.toLowerCase() === 'a') {
                          for(let cat in DB2_STRUCTURE) {
                             if(DB2_STRUCTURE[cat].includes(cleanKey)) { if(!db2Map[pn][cat]) db2Map[pn][cat] = []; db2Map[pn][cat].push({ label: cleanKey, value: cellVal }); }
                          }
                    }
                }
            }
        });

        const today = new Date().setHours(0,0,0,0);

        GLOBAL_DATA = rawDB
            .map((row, index) => {
                let item = { _id: index };
                let keys = Object.keys(row);
                let pnKey = keys.find(k => k.toUpperCase().includes('P/N'));
                let statusKey = keys.find(k => k.toUpperCase().includes('STATUS'));
                
                if (!pnKey || !row[pnKey] || String(row[pnKey]).trim().length === 0 || !statusKey) return null;

                // --- CRITICAL FIX: REMOVE SPACES FROM P/N TO ENSURE "ATTACHED" LOOK ---
                item.PN = String(row[pnKey]).replace(/\s/g, '').trim(); 

                let imgKey = keys.find(k => k.trim().toUpperCase() === 'PRIMAGE');
                item.ImageURL = imgKey ? cleanExtractedUrl(String(row[imgKey])) : "";

                for (let k in COL_MAP_DB) {
                    if (k === 'PN') continue; // Skip PN, handled above
                    let excelKey = COL_MAP_DB[k];
                    let foundKey = keys.find(x => x.trim().toUpperCase() === excelKey.trim().toUpperCase());
                    if(!foundKey) foundKey = keys.find(x => x.trim().toUpperCase().includes(excelKey.trim().toUpperCase()));
                    item[k] = normalize(row[foundKey || excelKey], k) || '-';
                }
                
                item._DIRECT_LINK = (!item.DirectLink || item.DirectLink.length < 5) ? "" : cleanExtractedUrl(item.DirectLink);
                item.ManualLink = (item.ManualLink && item.ManualLink.length > 5) ? cleanExtractedUrl(item.ManualLink) : "";

                const pn = item.PN;
                item._DETAILS = db2Map[pn] || {};
                item._SEARCH_STR = [item.PN, item.Status, item.CPU, item.Memory, item.size, item.Resolution, item.Type, item.WorkEnv, item.Vertical, item.Series].join(" || ").toLowerCase(); 
                
                let ltbDate = parseDate(item.LTB);
                if(ltbDate && ltbDate.getTime() < today) item.Status = 'EOL';
                
                return item;
            })
            .filter(item => item !== null);

        document.getElementById('loading').style.display = 'none';
        document.getElementById('landing-page').style.display = 'flex'; 
    } catch (err) { showError(err.message); }
}

function normalize(val, key) { 
    if (val == null) return ""; 
    if (key === 'LTB') return val; 
    let s = String(val).trim(); 
    if (key === 'size') { let n = parseFloat(s); if (!isNaN(n) && !s.includes('"')) return n + '"'; } 
    return s; 
}

function startScrolling(direction) {
    const container = document.getElementById('main-table-container'); const speed = 30; 
    function animate() { if (direction === 'left') container.scrollLeft -= speed; else container.scrollLeft += speed; scrollAnimationId = requestAnimationFrame(animate); }
    stopScrolling(); animate(); 
}
function stopScrolling() { if (scrollAnimationId) { cancelAnimationFrame(scrollAnimationId); scrollAnimationId = null; } }

function showSuggestions(val) {
    const box = document.getElementById('search-suggestions-box');
    if (!val || val.length < 2) { box.style.display = 'none'; return; }
    // Clean search
    const cleanSearch = val.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    
    const matches = GLOBAL_DATA.filter(item => {
        const cleanPN = item.PN.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
        return cleanPN.includes(cleanSearch);
    }).slice(0, 8);
    
    if (matches.length === 0) { box.style.display = 'none'; return; }
    box.innerHTML = matches.map(item => `
        <div class="suggestion-item" onclick="selectSuggestion('${item.PN}')">
            <strong>${item.PN}</strong>
            <span style="margin: 0 5px; color: var(--border);">-</span>
            <span class="text-metallic-blue">${item.CPU || ''}</span>
        </div>
    `).join('');
    box.style.display = 'block';
}

function selectSuggestion(pn) {
    document.getElementById('pn-search').value = pn; document.getElementById('search-suggestions-box').style.display = 'none';
    saveToHistory(pn); resetFilterStateOnly(); applyFilters(true);
}
function saveToHistory(term) {
    if(!term || term.length < 2) return;
    let history = JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]');
    history = history.filter(h => h.toLowerCase() !== term.toLowerCase()); history.unshift(term);
    if(history.length > 3) history = history.slice(0, 3);
    localStorage.setItem(HISTORY_KEY, JSON.stringify(history));
}
function showSearchHistory() {
    const history = JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]'); const box = document.getElementById('search-suggestions-box');
    if(history.length === 0) { box.style.display = 'none'; return; }
    box.innerHTML = `<div style="padding:12px 16px; font-size:11px; font-weight:800; color:var(--accent); background:rgba(0,0,0,0.2); border-bottom:2px solid var(--border); letter-spacing: 0.5px;">RECENT SEARCHES</div>` + 
    history.map(term => `<div class="suggestion-item" onclick="selectSuggestion('${term.replace(/"/g, '&quot;')}')"><strong>${term}</strong></div>`).join('');
    box.style.display = 'block';
}

function buildFilters() {
    const container = document.getElementById('filters-container'); container.innerHTML = '';
    const validData = GLOBAL_DATA.filter(i => ['MP', 'ES', 'NEW'].includes(String(i.Status).trim().toUpperCase()));
    let allFilters = [];
    ['LTB', 'size', 'Resolution', 'Type', 'CPU', 'Memory', 'WorkEnv', 'Vertical', 'Series'].forEach(key => allFilters.push({ key: key, isDB2: false }));
    let db2Keys = new Set();
    validData.forEach(item => { if(item._DETAILS) Object.keys(item._DETAILS).forEach(k => { if(!k.startsWith('__') && k.toLowerCase() !== 'empty') db2Keys.add(k); }); });
    db2Keys.forEach(k => { if(k.toLowerCase() !== 'p/n') allFilters.push({ key: k, isDB2: true }); });
    
    allFilters.forEach(f => {
        const isIndispensable = INDISPENSABLE_FILTERS.some(ind => ind.toLowerCase() === f.key.toLowerCase());
        if (isIndispensable || OPTIONAL_FILTERS_STATE[f.key]) {
           let uniqueVals = getUniqueValues(validData, f.key, f.isDB2);
           if(uniqueVals.length > 0) createFilterUI(f.key, uniqueVals, container, f.isDB2);
        }
    });
    const addBtnContainer = document.createElement('div');
    addBtnContainer.className = 'add-filter-container';
    addBtnContainer.style.alignSelf = "flex-end"; 
    const availableToAdd = allFilters.filter(f => !INDISPENSABLE_FILTERS.some(ind => ind.toLowerCase() === f.key.toLowerCase())).sort((a,b) => a.key.localeCompare(b.key));
    if (availableToAdd.length > 0) {
        let menuHtml = availableToAdd.map(f => {
           const isChecked = OPTIONAL_FILTERS_STATE[f.key] ? 'checked' : '';
           let displayLabel = UI_LABELS[f.key] || (f.key.charAt(0).toUpperCase() + f.key.slice(1));
           // Added data-key for robust selection logic
           return `<label class="filter-dropdown-item"><input type="checkbox" ${isChecked} data-key="${f.key}" onchange="toggleOptionalFilter('${f.key}')">${displayLabel}</label>`;
        }).join('');
        
        // FIXED SELECT ALL LOGIC
        addBtnContainer.innerHTML = `
        <button class="btn-filter-manage" onclick="toggleAddMenu()" title="Customize Filters"><span>✚</span> Customize Filters</button>
        <div id="add-filter-menu" class="filter-dropdown-menu">
            <div class="filter-dropdown-header">Manage View</div>
            <label class="filter-dropdown-item select-all-option">
                <input type="checkbox" onchange="toggleSelectAllOptions(this)"> Select All
            </label>
            ${menuHtml}
        </div>`;
        container.appendChild(addBtnContainer);
    }
}

function toggleAddMenu() { const menu = document.getElementById('add-filter-menu'); document.querySelectorAll('.multiselect-content').forEach(e => e.classList.remove('show')); menu.classList.toggle('show'); }
document.addEventListener('click', (e) => { if (!e.target.closest('.add-filter-container')) { const menu = document.getElementById('add-filter-menu'); if(menu) menu.classList.remove('show'); } });
function toggleOptionalFilter(key) { if (OPTIONAL_FILTERS_STATE[key]) { delete OPTIONAL_FILTERS_STATE[key]; if(ACTIVE_FILTERS[key]) delete ACTIVE_FILTERS[key]; } else { OPTIONAL_FILTERS_STATE[key] = true; } buildFilters(); applyFilters(); }

// --- FIXED: SELECT ALL FILTERS LOGIC ---
function toggleSelectAllOptions(source) {
    const menu = document.getElementById('add-filter-menu');
    // Get inputs that are NOT the "Select All" checkbox
    const inputs = menu.querySelectorAll('.filter-dropdown-item:not(.select-all-option) input');
    
    inputs.forEach(inp => {
        inp.checked = source.checked;
        const key = inp.getAttribute('data-key');
        if(key) {
            if(source.checked) {
                OPTIONAL_FILTERS_STATE[key] = true;
            } else {
                delete OPTIONAL_FILTERS_STATE[key];
                if(ACTIVE_FILTERS[key]) delete ACTIVE_FILTERS[key];
            }
        }
    });
    // Rebuild UI immediately to show changes
    buildFilters(); 
    applyFilters();
}

function getUniqueValues(dataSet, key, isDB2) {
    let vals = new Set();
    dataSet.forEach(i => { 
        if(!isDB2) {
            let rawVal = i[key];
            if(rawVal && rawVal !== '-' && !EXCLUSIONS.includes(String(rawVal).toLowerCase())) { 
                if(key === 'LTB') vals.add(formatDate(rawVal)); 
                else vals.add(simplifyForFilter(rawVal, key)); 
            }
        } else {
            if(i._DETAILS && i._DETAILS[key]) {
                i._DETAILS[key].forEach(attr => {
                    let val = attr.value; let label = attr.label;
                    if (key === 'Mounting' && String(val).toLowerCase().includes('wall')) return;
                    if (val && (val.toLowerCase() === 'a' || val.toLowerCase() === 'y' || val.toLowerCase() === 'yes')) vals.add(label);
                    else if (val && !EXCLUSIONS.includes(String(val).toLowerCase())) vals.add(simplifyForFilter(val, key));
                });
            }
        }
    });
    
    // --- CUSTOM SORTING LOGIC ---
    return Array.from(vals).sort((a, b) => { 
        if (key === 'LTB') { let dA = parseDDMMYYYY(a), dB = parseDDMMYYYY(b); return dB - dA; } 
        if (key === 'Memory') { return parseInt(a) - parseInt(b); } 
        
        // --- CUSTOM CPU SORTING (High End -> Low End -> Others) ---
        if (key === 'CPU') {
            const getPriority = (v) => {
                const s = v.toLowerCase();
                if (s.includes('core i7')) return 1;
                if (s.includes('core i5')) return 2;
                if (s.includes('core i3')) return 3;
                if (s.includes('xeon')) return 4;
                if (s.includes('pentium')) return 5;
                if (s.includes('celeron')) return 6;
                if (s.includes('atom')) return 7;
                return 10; // Rockchip, NXP, etc.
            };
            const pA = getPriority(a);
            const pB = getPriority(b);
            if (pA !== pB) return pA - pB;
        }

        return a.localeCompare(b, undefined, {numeric:true, sensitivity:'base'}); 
    });
}

function parseDDMMYYYY(s) { if(!s) return 0; let parts = s.split('/'); if(parts.length === 3) return new Date(parts[2], parts[1]-1, parts[0]).getTime(); return 0; }

function createFilterUI(key, options, container, isDB2 = false) {
    if(options.length === 0) return;
    let label = UI_LABELS[key] || (key.charAt(0).toUpperCase() + key.slice(1));
    let div = document.createElement('div'); div.className = 'filter-item';
    
    // COLOR SCALE FOR CPU DROPDOWN OPTIONS
    const getOptionStyle = (val) => {
        if (key !== 'CPU') return '';
        const v = val.toLowerCase();
        if(v.includes('i7')) return 'color:#DC2626; font-weight:700;'; // Red
        if(v.includes('i5')) return 'color:#EA580C; font-weight:700;'; // Orange
        if(v.includes('i3')) return 'color:#D97706; font-weight:700;'; // Amber
        if(v.includes('atom')) return 'color:#16A34A;'; // Green
        return '';
    };

    div.innerHTML = `<span class="filter-label">${label}</span><button class="multiselect-btn" onclick="toggleDD('${key}')" id="btn-${key}">All Selected</button><div class="multiselect-content" id="dd-${key}"><label class="select-all-option"><input type="checkbox" checked onchange="toggleSelectAll('${key}', this, ${isDB2})"> Select All</label>${options.map(v => `<label class="multiselect-option" style="${getOptionStyle(v)}"><input type="checkbox" value="${v}" checked onchange="updateLocalState('${key}', ${isDB2})"> ${v}</label>`).join('')}</div>`;
    container.appendChild(div);
}
function toggleDD(key) { let el = document.getElementById(`dd-${key}`); let open = el.classList.contains('show'); document.querySelectorAll('.multiselect-content').forEach(e => e.classList.remove('show')); if(!open) el.classList.add('show'); }
function toggleSelectAll(key, source, isDB2) { let checkboxes = document.querySelectorAll(`#dd-${key} input:not(.select-all-option input)`); checkboxes.forEach(cb => cb.checked = source.checked); updateLocalState(key, isDB2); }
function updateLocalState(key, isDB2) { let chk = document.querySelectorAll(`#dd-${key} input:not(.select-all-option input):checked`); let btn = document.getElementById(`btn-${key}`); let all = document.querySelectorAll(`#dd-${key} input:not(.select-all-option input)`).length; let selectAllBox = document.querySelector(`#dd-${key} .select-all-option input`); if(selectAllBox) selectAllBox.checked = (chk.length === all); if(chk.length===0) btn.innerText="None"; else if(chk.length===all) btn.innerText="All Selected"; else btn.innerText=`${chk.length} Selected`; ACTIVE_FILTERS[key] = { vals: Array.from(chk).map(c => c.value), isDB2: isDB2 }; applyFilters(); }

function resetFilters() { ACTIVE_FILTERS = {}; document.getElementById('pn-search').value = ''; document.getElementById('sort-order').value = 'status'; OPTIONAL_FILTERS_STATE = {}; buildFilters(); document.querySelectorAll('input[type="checkbox"]').forEach(c => c.checked = true); document.querySelectorAll('.multiselect-btn').forEach(b => b.innerText = "All Selected"); applyFilters(); document.getElementById('main-table-container').scrollLeft = 0; }
function resetFilterStateOnly() { document.querySelectorAll('input[type="checkbox"]').forEach(c => c.checked = true); document.querySelectorAll('.multiselect-btn').forEach(b => b.innerText = "All Selected"); ACTIVE_FILTERS = {}; }

function applyFilters(isSearchOverride = false) {
    // --- SMART SEARCH IMPLEMENTATION ---
    // 1. Get value and remove ALL non-alphanumeric chars (spaces, commas, symbols)
    let searchRaw = document.getElementById('pn-search').value.trim();
    let searchClean = searchRaw.replace(/[^a-zA-Z0-9]/g, '').toLowerCase(); 
    
    // Legacy support for space-separated keywords if no exact P/N match is expected
    let searchTerms = searchRaw.toLowerCase().split(/\s+/).filter(t => t.length > 0);

    let res = GLOBAL_DATA.filter(item => {
        if (!isSearchOverride && searchRaw.length === 0) { if (item.Status === 'EOL') return false; }
        
        if (searchRaw.length > 0) {
            // Smart P/N Match: Strip symbols from Item PN and compare with stripped Search
            let itemPNClean = item.PN.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
            
            // Priority 1: Smart Match (e.g. "PPC 318" -> "ppc318" matches "ppc318")
            if(itemPNClean.includes(searchClean)) {
                return true;
            }
            
            // Priority 2: Keyword match in other fields
            let itemStr = item._SEARCH_STR; 
            let match = searchTerms.every(term => { if (/^i[3579]$/.test(term)) return new RegExp(`\\b${term}\\b`, 'i').test(itemStr); if (/^\d+gb$/.test(term)) return new RegExp(`\\b${term}\\b`, 'i').test(itemStr); return itemStr.includes(term); });
            if (!match) return false;
        }
        
        for (let k in ACTIVE_FILTERS) {
           let filter = ACTIVE_FILTERS[k]; let allowed = filter.vals; if(allowed.length === 0) return false; 
           if (!filter.isDB2) { 
               let rawVal = item[k]; 
               let checkVal = (k === 'LTB') ? formatDate(rawVal) : simplifyForFilter(rawVal, k); // Use Simplified for comparison
               if (!allowed.includes(checkVal)) return false; 
           } 
           else { 
               let itemFeatures = []; 
               if(item._DETAILS[k]) { 
                   itemFeatures = item._DETAILS[k].map(d => { 
                       let v = d.value; 
                       if (v && (v.toLowerCase() === 'a' || v.toLowerCase() === 'y' || v.toLowerCase() === 'yes')) return d.label; 
                       return simplifyForFilter(v, k); 
                   }); 
               } 
               let match = allowed.some(val => itemFeatures.includes(val)); 
               if (!match) return false; 
           }
        }
        return true;
    });
    applySort(res);
}

function applySort(data) {
    const sortValue = document.getElementById('sort-order').value;
    data.sort((a, b) => {
        switch(sortValue) {
           case 'status': const rank = (s) => { s=String(s).toUpperCase(); if(s==='NEW') return 1; if(s==='ES') return 2; if(s==='MP') return 3; if(s==='LTB') return 4; return 5; }; return rank(a.Status) - rank(b.Status);
           case 'original': return (a._id || 0) - (b._id || 0);
           case 'pn_asc': return a.PN.localeCompare(b.PN); case 'pn_desc': return b.PN.localeCompare(a.PN);
           case 'ltb_desc': return (parseDate(b.LTB)||new Date(2099,0,1)) - (parseDate(a.LTB)||new Date(2099,0,1));
           case 'ltb_asc': return (parseDate(a.LTB)||new Date(2099,0,1)) - (parseDate(b.LTB)||new Date(2099,0,1));
           default: return 0;
        }
    });
    renderTable(data);
}

function renderTable(data) {
    CURRENT_VISIBLE_DATA = data; const body = document.getElementById('table-body'); body.innerHTML = '';
    if (!data || data.length === 0) { body.innerHTML = '<tr><td colspan="12" style="text-align:center; padding:70px; color:var(--text-muted); font-size:16px;">No results found.</td></tr>'; return; }
    
    // --- EMPTY CELL FORMATTER (Professional N/A) ---
    // If value is null, undefined, "-", "0", or empty string -> Return styled N/A
    const fmt = (val) => {
        if (!val || val === '-' || val === '0' || String(val).trim() === '') {
            return '<span class="empty-val">N/A</span>';
        }
        return val;
    };

    data.forEach(row => {
        try {
           let tr = document.createElement('tr'); let safePN = String(row.PN || "N/A");
           let link = "#"; if (row._DIRECT_LINK && row._DIRECT_LINK.length > 5) link = row._DIRECT_LINK; else if (safePN !== "N/A") link = `https://www.google.com/search?q=site:advantech.com+${safePN}&btnI=1`; 
           
           let statusStyle = getStatusColor(row.Status);
           let envStyle = getSemanticStyle(row.WorkEnv, 'WorkEnv');
           let vertStyle = getSemanticStyle(row.Vertical, 'Vertical');
           let seriesStyle = getSemanticStyle(row.Series, 'Series');
           
           // Clean CPU String
           let cleanCPU = simplifyValue(row.CPU, 'CPU');

           tr.innerHTML = `
               <td><span class="status-tag" style="${statusStyle}">${row.Status || '-'}</span></td>
               <td class="pn-cell">
                   <a href="${link}" target="_blank" class="pn-link" title="Open Product Page">${safePN}</a>
                   <button class="btn-copy-pn" onclick="copyToClipboard('${safePN}')" title="Copy P/N">
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path></svg>
                   </button>
               </td>
               <td>${fmt(formatDate(row.LTB))}</td>
               <td><span class="semantic-tag" style="${envStyle}">${fmt(simplifyValue(row.WorkEnv, 'WorkEnv'))}</span></td>
               <td><span class="semantic-tag" style="${vertStyle}">${fmt(row.Vertical)}</span></td>
               <td><span class="semantic-tag" style="${seriesStyle}">${fmt(row.Series)}</span></td>
               <td>${fmt(row.size)}</td>
               <td>${fmt(row.Resolution)}</td>
               <td class="cpu-cell">${fmt(cleanCPU)}</td>
               <td>${fmt(simplifyValue(row.Type, 'Type'))}</td>
               <td>${fmt(simplifyValue(row.Memory, 'Memory'))}</td>
               <td><button class="action-btn" onclick="openModal('${safePN}')">DETAILS</button></td>
           `;
           body.appendChild(tr);
        } catch (err) { console.warn("Skipping bad row:", row, err); }
    });
}

function findBestAlternative(currentItem) {
    const pn = currentItem.PN;
    if(!pn || pn.length < 5) return null;
    let familyPrefix = pn.split('-')[0]; 
    const candidates = GLOBAL_DATA.filter(cand => {
        if(cand.PN === currentItem.PN) return false;
        if(!cand.PN.startsWith(familyPrefix)) return false;
        if(String(cand.size).trim() !== String(currentItem.size).trim()) return false;
        const currentLTB = parseDate(currentItem.LTB);
        const candLTB = parseDate(cand.LTB);
        if(!currentLTB && candLTB && candLTB < new Date()) return false;
        if(currentLTB) { if(candLTB && candLTB <= currentLTB) return false; }
        return true;
    });
    let bestMatch = null; let maxScore = 0;
    candidates.forEach(cand => {
        let score = 0; score += 50;
        if(cand.Resolution === currentItem.Resolution) score += 20; else return; 
        const curTouch = simplifyValue(currentItem.Type, 'Type').toLowerCase();
        const candTouch = simplifyValue(cand.Type, 'Type').toLowerCase();
        if(curTouch === candTouch) { score += 15; } else if (curTouch.includes('resistive') && candTouch.includes('pcap')) { score += 10; } else { score -= 10; }
        if(cand.CPU !== currentItem.CPU) score += 5; 
        const curMem = parseInt(currentItem.Memory) || 0; const candMem = parseInt(cand.Memory) || 0;
        if(candMem >= curMem) score += 5;
        if(cand.PN.substring(0, 7) === currentItem.PN.substring(0, 7)) score += 5;
        if(score > maxScore) { maxScore = score; bestMatch = cand; }
    });
    if(maxScore >= 75) { return { item: bestMatch, score: maxScore }; }
    return null;
}

// --- INTERNAL SEARCH TRIGGER ---
function triggerInternalSearch(pn) {
    closeModal();
    const searchInput = document.getElementById('pn-search');
    searchInput.value = pn;
    saveToHistory(pn);
    resetFilterStateOnly(); 
    applyFilters(true);
    document.getElementById('main-table-container').scrollLeft = 0;
    window.scrollTo({ top: 100, behavior: 'smooth' });
}

function openModal(pn) {
    let item = GLOBAL_DATA.find(i => i.PN === pn);
    if(!item) return;
    document.getElementById('modal-overlay').style.display = 'flex';

    let productLink = "#";
    if (item._DIRECT_LINK && item._DIRECT_LINK.length > 5) { productLink = item._DIRECT_LINK; } 
    else if (pn !== "N/A") { productLink = `https://www.google.com/search?q=site:advantech.com+${pn}&btnI=1`; }

    const createLinkBtn = (url, label, icon, cssClass) => {
        if (!url || typeof url !== 'string' || url.length < 5 || url.toLowerCase().includes('n/a')) return "";
        return `<a href="${url.trim()}" target="_blank" rel="noopener noreferrer" class="${cssClass}">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">${icon}</svg>
            <span>${label}</span>
        </a>`;
    };

    const iconDS = `<path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"></path><polyline points="7 11 12 16 17 11"></polyline><line x1="12" y1="4" x2="12" y2="16"></line>`;
    let btnsHtml = createLinkBtn(item.ManualLink, "DATASHEET", iconDS, "btn-impact");
    let headerActionsHtml = btnsHtml ? `<div class="header-btn-group">${btnsHtml}</div>` : '';

    let hasImage = (item.ImageURL && typeof item.ImageURL === 'string' && item.ImageURL.length > 10 && !item.ImageURL.toLowerCase().includes('undefined'));
    let layoutClass = hasImage ? 'modal-layout-split' : 'modal-layout-full';
     
    let commentHtml = '';
    if(item.Comment && item.Comment !== '-' && !EXCLUSIONS.includes(item.Comment.toLowerCase())) {
          commentHtml = `<div class="modal-comment-box"><span class="modal-comment-title">Note</span>${item.Comment}</div>`;
    }

    let imgColHtml = '';
    if (hasImage) {
        imgColHtml = `
        <div style="display:flex; flex-direction:column; gap:10px;">
            <div class="modal-preview-card" onclick="openLightbox('${item.ImageURL}')">
                <img src="${item.ImageURL}" class="modal-product-img" alt="${item.PN}" onerror="this.style.display='none'; this.parentElement.style.display='none';">
                <div class="zoom-hint">🔍 Click to Expand</div>
            </div>
            ${commentHtml}
        </div>`;
    } else {
         if(commentHtml) imgColHtml = commentHtml;
    }

    let detailsHtml = "";
    for (let cat in item._DETAILS) {
        if(cat.toUpperCase().includes('P/N') || cat.startsWith('__')) continue; 
        if(item._DETAILS[cat].length > 0) {
            detailsHtml += `<div class="detail-section-title">${cat}</div><div class="detail-grid">`;
            item._DETAILS[cat].forEach(feat => {
                let val = String(feat.value).trim(); 
                let isBoolean = (val.length < 5 && (val.toLowerCase() === 'a' || val.toLowerCase() === 'y' || val.toLowerCase() === 'yes'));
                if (isBoolean) { detailsHtml += `<div class="detail-simple-tag">${feat.label}</div>`; } 
                else { detailsHtml += `<div class="detail-card-value"><div class="detail-label-small">${feat.label}</div><div class="detail-value-text">${feat.value}</div></div>`; }
            });
            detailsHtml += `</div>`;
        }
    }

    const alternative = findBestAlternative(item);
    let alternativeHtml = "";

    if(alternative) {
        const altItem = alternative.item;
        const score = alternative.score;
        let curLTBText = item.LTB ? `LTB: ${formatDate(item.LTB)}` : "Active / MP";
        let altLTBText = altItem.LTB ? `LTB: ${formatDate(altItem.LTB)}` : "Active / MP";
        let altLTBClass = (!altItem.LTB || parseDate(altItem.LTB) > new Date()) ? "good" : "";

        alternativeHtml = `
        <div class="comp-section">
            <div class="comp-header">
                <span>SUGGESTED UPGRADE</span>
                <span style="background:rgba(255,255,255,0.25); color:#fff; padding:2px 8px; border-radius:4px; font-size:10px;">RECOMMENDED</span>
            </div>
            <div class="comp-body">
                <div class="comp-bar-container">
                    <div class="comp-score-label"><span>Compatibility Score</span><span>${score}%</span></div>
                    <div class="comp-progress-bg"><div class="comp-progress-fill" style="width: ${score}%"></div></div>
                </div>
                <div class="comp-grid">
                    <div class="comp-col">
                        <span class="comp-label">CURRENT SELECTION</span>
                        <span class="comp-pn" style="opacity:0.6;">${item.PN}</span>
                        <span class="comp-specs">${simplifyValue(item.CPU, 'CPU')}<br>${item.Memory ? simplifyValue(item.Memory, 'Memory') : 'N/A'}<br><div class="comp-ltb">${curLTBText}</div></span>
                    </div>
                    <div class="comp-divider"></div>
                    <div class="comp-col">
                        <span class="comp-label" style="color:#38BDF8;">NEW MODEL</span>
                        <div class="comp-pn internal-link-action" onclick="triggerInternalSearch('${altItem.PN}')" title="Filter list for this item">${altItem.PN}</div>
                        <span class="comp-specs"><span class="comp-highlight">${simplifyValue(altItem.CPU, 'CPU')}</span><br><span class="comp-highlight">${altItem.Memory ? simplifyValue(altItem.Memory, 'Memory') : 'N/A'}</span><br><div class="comp-ltb ${altLTBClass}">${altLTBText}</div></span>
                    </div>
                </div>
            </div>
        </div>
        `;
    }
    
    document.getElementById('modal-content').innerHTML = `
    <div style="display:flex; justify-content:space-between; align-items:start; border-bottom:2px solid #E2E8F0; padding-bottom:22px; margin-bottom:25px;">
        <div>
            <h2 style="margin:0; font-size:32px; line-height:1.2;">
                <a href="${productLink}" target="_blank" style="color:#0F172A; text-decoration:none; transition:color 0.2s;" onmouseover="this.style.color='#0070E0'" onmouseout="this.style.color='#0F172A'" title="Open Product Page">${pn}</a>
            </h2>
            <div style="color:#64748B; font-size:18px; margin-top:8px; font-weight:500;">${item.Status} <span style="margin:0 10px; color:#E2E8F0;">|</span> ${simplifyValue(item.CPU, 'CPU')}</div>
        </div>
        ${headerActionsHtml}
    </div>

    <div class="modal-container-grid ${layoutClass}">
        ${imgColHtml} 
        <div style="display:grid; grid-template-columns: 1fr; gap:0px; align-content:start;">
           ${detailsHtml || '<div style="padding:30px; color:#64748B; font-size:18px; text-align:center;">No additional technical details found.</div>'}
           ${alternativeHtml}
        </div>
    </div>
    `;
}

function closeModal(e) { if(!e || e.target.id==='modal-overlay' || e.target.className.includes('close')) document.getElementById('modal-overlay').style.display='none'; }
function toggleHelp() { let h=document.getElementById('help-banner'); h.style.display=h.style.display==='block'?'none':'block'; }
function toggleTheme() { const html = document.documentElement; const current = html.getAttribute('data-theme'); const next = current === 'dark' ? 'light' : 'dark'; html.setAttribute('data-theme', next); localStorage.setItem('theme', next); updateThemeIcon(next); }
function updateThemeIcon(theme) { document.querySelector('.theme-toggle').innerHTML = theme === 'dark' ? '☀' : '🌙'; }
function getStatusColor(status) {
    const s = String(status || '').trim().toUpperCase();
    if(s === 'MP') return 'color: #10B981; background: rgba(16, 185, 129, 0.1);'; if(s === 'LTB') return 'color: #F59E0B; background: rgba(245, 158, 11, 0.1);';
    if(s === 'NEW') return 'color: #38BDF8; background: rgba(56, 189, 248, 0.1);'; if(s === 'ES') return 'color: #00C2FF; background: rgba(0, 194, 255, 0.1);'; if(s === 'EOL') return 'color: #EF4444; background: rgba(239, 68, 68, 0.15); border: 1px solid #EF4444;'; return 'color: #EF4444; background: rgba(239, 68, 68, 0.1);'; 
}
function copyToClipboard(text) { navigator.clipboard.writeText(text).then(() => { let toast = document.getElementById("toast"); toast.className = "show"; setTimeout(function(){ toast.className = toast.className.replace("show", ""); }, 3000); }); }
function showError(msg) { document.getElementById('loading').style.display = 'none'; const box = document.getElementById('error-box'); box.style.display = 'block'; box.innerHTML = `<strong>ERROR:</strong> ${msg}`; }
function toggleExportMenu() { const menu = document.getElementById('export-menu'); const isShown = menu.classList.contains('show'); document.querySelectorAll('.multiselect-content').forEach(e => e.classList.remove('show')); if (isShown) menu.classList.remove('show'); else menu.classList.add('show'); }
document.addEventListener('click', function(e) { if (!e.target.closest('.export-wrapper')) { const menu = document.getElementById('export-menu'); if(menu) menu.classList.remove('show'); } });

function exportData(type) {
    if (!CURRENT_VISIBLE_DATA || CURRENT_VISIBLE_DATA.length === 0) { alert("No data to export!"); return; }
    try {
        const dateStr = new Date().toISOString().split('T')[0]; const fileName = `Advantech_List_${dateStr}`;
        if (type === 'pdf') {
            if (!window.jspdf || !window.jspdf.jsPDF) { alert("PDF Library missing!"); return; }
            const { jsPDF } = window.jspdf; const doc = new jsPDF({ orientation: 'landscape' });
            const logoImg = document.querySelector('.advantech-logo'); let logoData = null; if(logoImg && logoImg.complete && logoImg.naturalWidth > 0) { try { const canvas = document.createElement("canvas"); canvas.width = logoImg.naturalWidth; canvas.height = logoImg.naturalHeight; const ctx = canvas.getContext("2d"); ctx.drawImage(logoImg, 0, 0); logoData = canvas.toDataURL("image/png"); } catch(e) {} }
            const headerY = 15; if (logoData) doc.addImage(logoData, 'PNG', 14, 8, 35, 10); doc.setFontSize(16); doc.setTextColor(0, 86, 210); doc.text(`Product List Export`, 14, headerY + 15); doc.setFontSize(10); doc.setTextColor(100); doc.text(`Generated: ${dateStr}`, 14, headerY + 22);
            const tableColumn = ["Status", "P/N", "LTB", "Env", "Size", "Res", "CPU", "Touch", "Mem"]; const tableRows = CURRENT_VISIBLE_DATA.map(row => [ row.Status, row.PN, formatDate(row.LTB), row.WorkEnv, row.size, row.Resolution, row.CPU, row.Type, row.Memory ]);
            doc.autoTable({ head: [tableColumn], body: tableRows, startY: headerY + 28, theme: 'grid', styles: { fontSize: 8, cellPadding: 2, valign: 'middle' }, headStyles: { fillColor: [0, 112, 224], textColor: 255, fontStyle: 'bold' }, alternateRowStyles: { fillColor: [241, 245, 249] }, margin: { top: 20 } }); doc.save(`${fileName}.pdf`);
        } else {
            const cleanData = CURRENT_VISIBLE_DATA.map(row => { let detailStr = ""; if (row._DETAILS) { detailStr = Object.entries(row._DETAILS).map(([k, v]) => { let vals = v.map(x => { let s = String(x.value).trim(); if (s.length < 3 && s.toUpperCase().includes('A')) return x.label; return x.value; }).join(', '); return `${k.toUpperCase()}: ${vals}`; }).join(' | \n'); } return { "Status": row.Status, "P/N": row.PN, "LTB Date": row.LTB instanceof Date ? row.LTB.toLocaleDateString() : row.LTB, "Work Environment": row.WorkEnv, "Size": row.size, "Resolution": row.Resolution, "CPU": row.CPU, "Touch Type": row.Type, "Memory": row.Memory, "Vertical": row.Vertical, "Series": row.Series, "Comment": row.Comment, "Features Summary": detailStr }; });
            const ws = XLSX.utils.json_to_sheet(cleanData); if (type === 'csv') { const csvOutput = XLSX.utils.sheet_to_csv(ws); const blob = new Blob(["\ufeff" + csvOutput], { type: 'text/csv;charset=utf-8;' }); const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = `${fileName}.csv`; document.body.appendChild(a); a.click(); document.body.removeChild(a); } else { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "ProductList"); XLSX.writeFile(wb, `${fileName}.xlsx`); }
        }
    } catch (err) { console.error("Export Error:", err); alert("Export Error: " + err.message); } const menu = document.getElementById('export-menu'); if (menu) menu.classList.remove('show');
}

if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('./sw.js').then(reg => console.log('Service Worker Registered!')).catch(err => console.log('Service Worker Error:', err));
    });
}
