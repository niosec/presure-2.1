/**
 * PresuRE v2.1 - Lógica Principal de la Aplicación
 */
// --- CONSTANTES ---
const STORAGE_KEY = 'costos_bolivia_data';
const initialDB = {
    materiales: [
        { code: "1", desc: "Cemento portland", unit: "kg", price: 1.26 },
        { code: "2", desc: "Arena fina", unit: "m3", price: 150 },
        { code: "3", desc: "Grava", unit: "m3", price: 160 }
    ],
    mano_obra: [
        { code: "1", desc: "Albañil", unit: "hr", price: 20 },
        { code: "2", desc: "Ayudante", unit: "hr", price: 15 }
    ],
    equipos: [
        { code: "1", desc: "Mezcladora", unit: "hr", price: 28.67 },
        { code: "2", desc: "Vibradora", unit: "hr", price: 21 }
    ]
};

// --- ESTADO GLOBAL ---
let appData = {
    settings: {
        social: 55.00, iva_mo: 14.94, tools: 5.00, gg: 10.00, util: 10.00, it: 3.09,
        decimals_yield: 5,
        decimals_price: 3,
        decimals_partial: 4,
        decimals_total: 2,
        decimals_qty: 2,
        numberFormat: 'intl'
    },
    database: JSON.parse(JSON.stringify(initialDB)),
    itemBank: [],
    projectItems: [],
    modules: [],
    activeModuleId: null
};

let editorContext = null;
let currentEditId = null;
let editorBackup = null;
let isCreatingNew = false;

let dbSortState = {
    materiales: { field: 'desc', dir: 1 },
    mano_obra: { field: 'desc', dir: 1 },
    equipos: { field: 'desc', dir: 1 }
};

let currentInsumosType = '';
let insumosList = [];
let conflictData = null;

let currentPage = {
    'b1': 1,
    'bank': 1,
    'db-mat': 1,
    'db-mo': 1,
    'db-eq': 1
};
const itemsPerPage = 50;
let saveTimeout;

let currentEditField = null;
let currentEditRow = null;
let currentEditType = null;
let selectedBankIds = new Set();

// Variables para PWA (v2.0)
let deferredPrompt;

// --- FUNCIONES AUXILIARES ---
// --- FUNCIONES DE DEBOUNCE (Búsquedas Optimizadas) ---
function debounce(func, wait) {
    let timeout;
    return function (...args) {
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(this, args), wait);
    };
}

// Envoltorios debounced para nuestras funciones de búsqueda (300ms de espera)
const handleBankSearch = debounce(() => { currentPage.bank = 1; renderBankList(); }, 300);
const handleModalSearch = debounce(filterModalList, 300);
const handleBankModalSearch = debounce(filterBankModal, 300);
const handleCalcSearch = debounce(searchCalcItem, 300);

function escapeHtml(text) {
    if (text == null) return "";
    return text.toString()
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function solveMathExpression(value) {
    if (!value) return 0;
    let clean = value.toString().trim();
    if (clean.startsWith('=')) clean = clean.substring(1);

    const format = appData.settings.numberFormat || 'intl';

    let tempCheck = clean;
    if (format === 'euro' || format === 'latam') {
        tempCheck = clean.replace(/,/g, '.');
    }

    if (/^[\d\.\,\+\-\*\/\(\)\s]+$/.test(clean)) {
        try {
            let jsExpression = clean;
            if (format === 'euro') {
                jsExpression = jsExpression.replace(/\./g, '').replace(/,/g, '.');
            } else if (format === 'latam') {
                jsExpression = jsExpression.replace(/,/g, '');
            } else {
                jsExpression = jsExpression.replace(/\s/g, '').replace(/,/g, '.');
            }

            const result = new Function('return ' + jsExpression)();
            if (!isFinite(result) || isNaN(result)) return 0;
            return result;
        } catch (e) {
            return 0;
        }
    }
    return parseFloat(clean) || 0;
}

function parseNumber(value, allowExpressions = false) {
    if (typeof value === 'number') return isNaN(value) ? 0 : value;
    if (value === null || value === undefined) return 0;

    let clean = value.toString().trim();
    if (clean === '' || clean === '-') return 0;

    if (allowExpressions) return solveMathExpression(value);

    const format = appData.settings.numberFormat || 'intl';

    switch (format) {
        case 'euro':
            clean = clean.replace(/\./g, '');
            clean = clean.replace(/,/g, '.');
            break;
        case 'latam':
            clean = clean.replace(/,/g, '');
            break;
        case 'intl':
        default:
            clean = clean.replace(/\s/g, '');
            clean = clean.replace(/,/g, '.');
            break;
    }

    const result = parseFloat(clean);
    if (!isFinite(result) || isNaN(result)) return 0;
    if (Math.abs(result) > 1e15) return 0;
    return result;
}

function roundToConfig(value, type = 'total') {
    if (value === null || value === undefined || isNaN(value)) return 0;
    let decimals;
    const settings = appData.settings;
    switch (type) {
        case 'yield': decimals = settings.decimals_yield; break;
        case 'price': decimals = settings.decimals_price; break;
        case 'partial': decimals = settings.decimals_partial; break;
        case 'total': decimals = settings.decimals_total; break;
        case 'qty': decimals = settings.decimals_qty; break;
        default: decimals = 2;
    }
    return Number(Math.round(value + 'e' + decimals) + 'e-' + decimals);
}

function formatNumber(value, type = 'total') {
    if (value === null || value === undefined || isNaN(value)) return '0';
    const settings = appData.settings;
    let decimals;
    switch (type) {
        case 'yield': decimals = settings.decimals_yield; break;
        case 'price': decimals = settings.decimals_price; break;
        case 'partial': decimals = settings.decimals_partial; break;
        case 'total': decimals = settings.decimals_total; break;
        case 'qty': decimals = settings.decimals_qty; break;
        case 'code': return Math.round(value).toString();
        default: decimals = settings.decimals_total;
    }
    let numStr = parseFloat(value).toFixed(decimals);
    let parts = numStr.split('.');
    let integerPart = parts[0];
    let decimalPart = parts[1] || '';
    const format = settings.numberFormat || 'intl';
    switch (format) {
        case 'intl':
            integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, " ");
            return decimalPart ? `${integerPart}.${decimalPart}` : integerPart;
        case 'latam':
            integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
            return decimalPart ? `${integerPart}.${decimalPart}` : integerPart;
        case 'euro':
            integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
            return decimalPart ? `${integerPart},${decimalPart}` : integerPart;
        case 'raw':
        default:
            return decimalPart ? `${integerPart}.${decimalPart}` : integerPart;
    }
}

function fmt(number, type = 'total') {
    return formatNumber(number, type);
}

function autoformatInput(inputElement, type = 'total') {
    if (!inputElement.value || inputElement.value.trim() === '') return;
    const parsed = parseNumber(inputElement.value);
    const rounded = roundToConfig(parsed, type);
    inputElement.value = formatNumber(rounded, type);
}

function handleMathInput(inputElement, type = 'total') {
    const rawResult = solveMathExpression(inputElement.value);
    const roundedResult = roundToConfig(rawResult, type);
    inputElement.value = formatNumber(roundedResult, type);
    return roundedResult;
}

function detectUserLocale() {
    const locale = navigator.language || navigator.userLanguage || 'en-US';
    const localeMap = {
        'es-ES': 'euro', 'de-DE': 'euro', 'fr-FR': 'euro', 'it-IT': 'euro',
        'es-MX': 'latam', 'es-AR': 'latam', 'en-US': 'latam', 'pt-BR': 'latam'
    };
    if (localeMap[locale]) return localeMap[locale];
    const lang = locale.split('-')[0];
    const langDefaults = { 'es': 'latam', 'en': 'latam', 'de': 'euro', 'fr': 'euro', 'it': 'euro', 'pt': 'latam' };
    return langDefaults[lang] || 'intl';
}

function showToast(message) {
    const x = document.getElementById("toast");
    x.textContent = message;
    x.className = "show";
    setTimeout(function () { x.className = x.className.replace("show", ""); }, 3000);
}

// --- PERSISTENCIA (Migrada a IndexedDB con localForage) ---
function saveData() {
    const indicator = document.getElementById('save-indicator');
    
    // 1. Mostrar que está guardando
    if (indicator) {
        indicator.classList.remove('saved', 'fade');
        indicator.classList.add('saving');
    }
    
    clearTimeout(saveTimeout);

    saveTimeout = setTimeout(() => {
        localforage.setItem(STORAGE_KEY, appData)
            .then(() => {
                // 2. Mostrar que terminó
                if (indicator) {
                    indicator.classList.remove('saving');
                    indicator.classList.add('saved');
                    
                    // 3. Empezar a desvanecer después de 1 segundo
                    setTimeout(() => {
                        indicator.classList.add('fade');
                    }, 1000);
                    
                    // 4. Resetear el ancho a 0 cuando ya no se ve
                    setTimeout(() => {
                        indicator.classList.remove('saved', 'fade');
                    }, 2000);
                }
            })
            .catch(err => {
                console.error("Error al guardar en IndexedDB", err);
                showToast('Error crítico: No se pudo guardar el proyecto.');
                // En caso de error, podríamos poner la línea en rojo
                if(indicator) {
                    indicator.classList.remove('saving');
                    indicator.style.backgroundColor = 'var(--danger)';
                }
            });
    }, 1500);
}

async function loadFromStorage() {
    try {
        // Obtenemos el objeto directamente, sin JSON.parse
        const stored = await localforage.getItem(STORAGE_KEY);

        if (stored) {
            appData = { ...appData, ...stored };

            // --- Tu lógica original de inicialización se mantiene intacta ---
            if (!appData.database) appData.database = JSON.parse(JSON.stringify(initialDB));
            if (!appData.projectItems) appData.projectItems = [];
            if (!appData.itemBank) appData.itemBank = [];
            if (!appData.modules || appData.modules.length === 0) {
                const defaultId = 'mod_' + Date.now();
                appData.modules = [{ id: defaultId, name: "General" }];
                appData.activeModuleId = defaultId;
            }
            ['materiales', 'mano_obra', 'equipos'].forEach(type => {
                if (appData.database[type]) {
                    appData.database[type].forEach((item, idx) => {
                        if (!item.id) item.id = Date.now() + idx + Math.floor(Math.random() * 10000);
                    });
                }
            });
            if (!appData.settings) appData.settings = {};
            const defaults = { social: 55, iva_mo: 14.94, tools: 5, gg: 10, util: 10, it: 3.09, decimals_yield: 5, decimals_price: 3, decimals_total: 2, numberFormat: 'intl', hiddenColumnsB1: [] };
            appData.settings = { ...defaults, ...appData.settings };
        } else {
            // Valores por defecto si la base de datos está vacía
            const defaultId = 'mod_' + Date.now();
            appData.modules = [{ id: defaultId, name: "General" }];
            appData.activeModuleId = defaultId;
            ['materiales', 'mano_obra', 'equipos'].forEach(type => {
                appData.database[type].forEach((item, idx) => {
                    item.id = Date.now() + idx + Math.floor(Math.random() * 10000);
                });
            });
        }
    } catch (e) {
        console.error("Error cargando IndexedDB", e);
    }
}

// --- FUNCIONES DE NAVEGACIÓN Y RENDER (v1.3) ---
function switchTab(tabId) {
    // 1. Lógica visual de cambio de pestaña (Código original)
    localStorage.setItem('presure_active_tab', tabId);
    document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));

    let pId = '';
    if (tabId === 'b1') pId = 'b1';
    else if (tabId === 'items') pId = 'items';
    else if (tabId === 'db') pId = 'db';
    else if (tabId === 'calc') pId = 'calc';
    else pId = 'config';

    const targetPanel = document.getElementById('panel-' + pId);
    if (targetPanel) targetPanel.classList.add('active');

    const tabs = document.querySelectorAll('.tab');
    if (tabId === 'b1') tabs[0].classList.add('active');
    if (tabId === 'items') tabs[1].classList.add('active');
    if (tabId === 'db') tabs[2].classList.add('active');
    if (tabId === 'calc') tabs[3].classList.add('active');

    window.scrollTo(0, 0);

    // --- Auto-foco Inteligente ---
    setTimeout(() => {
        if (tabId === 'items') {
            // Foco en buscador del Banco
            const searchInput = document.getElementById('main-bank-search');
        } else if (tabId === 'calc') {
            // Foco en buscador de Calculadora
            const searchInput = document.getElementById('calc-search-input');
            if (searchInput) { searchInput.focus(); searchInput.select(); }
        } else if (tabId === 'db') {
            // Opcional: Si quisieras foco en insumos, pero como es tabla, mejor no molestar.
        }
    }, 150); // Pequeño retardo para dar tiempo a la animación CSS
}

function changePage(section, direction) {
    const totalItems = getTotalItemsForSection(section);
    const totalPages = Math.ceil(totalItems / itemsPerPage);
    currentPage[section] += direction;
    if (currentPage[section] < 1) currentPage[section] = 1;
    if (currentPage[section] > totalPages) currentPage[section] = totalPages;
    switch (section) {
        case 'b1': renderB1(); break;
        case 'bank': renderBankList(); break;
        case 'db-mat': case 'db-mo': case 'db-eq': renderDBTables(); break;
    }
}

function getTotalItemsForSection(section) {
    switch (section) {
        case 'b1': return appData.projectItems.filter(item => item.moduleId === appData.activeModuleId).length;
        case 'bank': return appData.itemBank.length;
        case 'db-mat': return appData.database.materiales.length;
        case 'db-mo': return appData.database.mano_obra.length;
        case 'db-eq': return appData.database.equipos.length;
        default: return 0;
    }
}

function updatePaginationInfo(section, totalItems) {
    const totalPages = Math.ceil(totalItems / itemsPerPage);
    const paginationDiv = document.getElementById(`${section}-pagination`);
    const pageInfo = document.getElementById(`${section}-page-info`);
    if (totalPages > 1) {
        paginationDiv.style.display = 'flex';
        pageInfo.textContent = `Página ${currentPage[section]} de ${totalPages} (${totalItems} items)`;
    } else {
        paginationDiv.style.display = 'none';
    }
}

// --- MÓDULOS ---

function renderModuleTabs() {
    const container = document.getElementById('module-tabs');
    container.innerHTML = '';
    appData.modules.forEach(mod => {
        const tab = document.createElement('div');
        tab.className = `module-tab ${mod.id === appData.activeModuleId ? 'active' : ''}`;
        tab.textContent = mod.name;
        tab.onclick = () => switchModule(mod.id);
        tab.ondblclick = () => renameCurrentModule();
        container.appendChild(tab);
    });
    const currentMod = appData.modules.find(m => m.id === appData.activeModuleId);
    if (currentMod) document.getElementById('current-module-name-display').innerText = currentMod.name;
}

function switchModule(id) {
    appData.activeModuleId = id;
    currentPage.b1 = 1;
    saveData();
    renderB1();
}

function addModule() {
    const name = prompt("Nombre del nuevo módulo:", "Nuevo Módulo");
    if (name) {
        const newId = 'mod_' + Date.now() + Math.floor(Math.random() * 1000);
        appData.modules.push({ id: newId, name: name });
        switchModule(newId);
        showToast("Módulo creado");
    }
}

function renameCurrentModule() {
    const current = appData.modules.find(m => m.id === appData.activeModuleId);
    if (!current) return;
    const newName = prompt("Renombrar módulo:", current.name);
    if (newName && newName.trim() !== "") {
        current.name = newName;
        saveData();
        renderModuleTabs();
    }
}

function cloneCurrentModule() {
    const currentMod = appData.modules.find(m => m.id === appData.activeModuleId);
    if (!currentMod) return;
    const newName = prompt("Nombre para el módulo clonado:", "Copia de " + currentMod.name);
    if (newName === null) return;
    const finalName = newName.trim() === "" ? "Copia de " + currentMod.name : newName;
    const newModuleId = 'mod_' + Date.now() + Math.floor(Math.random() * 1000);
    appData.modules.push({ id: newModuleId, name: finalName });
    const itemsToClone = appData.projectItems.filter(i => i.moduleId === appData.activeModuleId);
    let baseTime = Date.now();
    itemsToClone.forEach((item, index) => {
        const newItem = JSON.parse(JSON.stringify(item));
        newItem.id = baseTime + index + Math.floor(Math.random() * 10000);
        newItem.moduleId = newModuleId;
        appData.projectItems.push(newItem);
    });
    saveData();
    switchModule(newModuleId);
    showToast("Módulo clonado exitosamente");
}

function deleteCurrentModule() {
    if (appData.modules.length <= 1) {
        showToast("Debe existir al menos un módulo.");
        return;
    }
    const itemsInModule = appData.projectItems.filter(i => i.moduleId === appData.activeModuleId);
    if (itemsInModule.length > 0) {
        if (!confirm(`El módulo tiene ${itemsInModule.length} ítems. ¿Estás seguro de eliminarlos todos?`)) return;
        appData.projectItems = appData.projectItems.filter(i => i.moduleId !== appData.activeModuleId);
    }
    appData.modules = appData.modules.filter(m => m.id !== appData.activeModuleId);
    appData.activeModuleId = appData.modules[appData.modules.length - 1].id;
    saveData();
    renderB1();
    showToast("Módulo eliminado");
}

// --- PRESUPUESTO B1 ---

function renderB1() {
    renderModuleTabs();
    const tbody = document.getElementById('b1-body');
    tbody.innerHTML = '';
    const moduleItems = appData.projectItems.filter(item => item.moduleId === appData.activeModuleId);
    const totalItems = moduleItems.length;
    const start = (currentPage.b1 - 1) * itemsPerPage;
    const end = Math.min(start + itemsPerPage, totalItems);
    const itemsToRender = moduleItems.slice(start, end);
    const fragment = document.createDocumentFragment();
    let moduleTotal = 0;
    itemsToRender.forEach((item, index) => {
        const globalIndex = start + index;
        const rawUnitPrice = calculateUnitPrice(item);
        const roundedUnitPrice = roundToConfig(rawUnitPrice, 'total');
        const roundedQty = roundToConfig(item.quantity, 'qty');
        const total = roundedUnitPrice * roundedQty;
        moduleTotal += total;
        const tr = document.createElement('tr');
        tr.setAttribute('data-id', item.id);
        tr.innerHTML = `
            <td class="text-center">
                <button class="btn btn-primary btn-sm" onclick="editItem('project', ${item.id})" title="Editar">
                    <i class="fas fa-edit"></i>
                </button>
            </td>
            <td class="text-center" style="font-weight:600;">${globalIndex + 1}</td>
            <td><input type="text" value="${item.projectCode || ''}" onchange="updateProjectItem(${item.id}, 'projectCode', this.value)" placeholder="-" style="text-align:center;"></td>
            <td><textarea class="table-input" onchange="updateProjectItem(${item.id}, 'description', this.value)">${escapeHtml(item.description)}</textarea></td>
            <td><input type="text" value="${escapeHtml(item.unit)}" onchange="updateProjectItem(${item.id}, 'unit', this.value)"></td>
            <td><input type="text" inputmode="decimal" value="${fmt(item.quantity, 'qty')}" onchange="updateProjectItem(${item.id}, 'quantity', handleMathInput(this, 'qty'))" onfocus="this.select()" style="text-align:center;"></td>
            <td class="text-right" style="font-weight:600;">${fmt(rawUnitPrice, 'total')}</td>
            <td class="text-right font-bold">${fmt(total, 'total')}</td>
            <td class="text-center">
                <div class="actions-row">
                    <button class="btn btn-info btn-sm" onclick="duplicateProjectItem(${item.id})" title="Duplicar">
                        <i class="fas fa-copy"></i>
                    </button>
                    <button class="btn btn-warning btn-sm" onclick="saveProjectItemToBank(${item.id})" title="Guardar en Banco">
                        <i class="fas fa-save"></i>
                    </button>
                    <button class="btn btn-danger btn-sm" onclick="deleteProjectItem(${item.id})" title="Eliminar">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
            </td>
        `;
        fragment.appendChild(tr);
    });
    tbody.appendChild(fragment);
    updatePaginationInfo('b1', totalItems);
    recalculateTotalsDisplay();
    applyColumnVisibility();
    initSortableB1();
}

function updateProjectItem(id, field, value) {
    const item = appData.projectItems.find(i => i.id === id);
    if (item) {
        item[field] = (field === 'quantity') ? parseNumber(value) : value;
        saveData();
        if (field === 'quantity') updateRowCalculations(id);
    }
}

function updateRowCalculations(itemId) {
    const item = appData.projectItems.find(i => i.id === itemId);
    if (!item) return;
    const row = document.querySelector(`tr[data-id="${itemId}"]`);
    if (!row) return;
    const rawUnitPrice = calculateUnitPrice(item);
    const roundedUnitPrice = roundToConfig(rawUnitPrice, 'total');
    const roundedQty = roundToConfig(item.quantity, 'qty');
    const total = roundedUnitPrice * roundedQty;
    const totalCell = row.cells[6];
    if (totalCell) totalCell.innerText = fmt(total, 'total');
    recalculateTotalsDisplay();
}

function recalculateTotalsDisplay() {
    const getLineTotal = (item) => {
        const rawUnitPrice = calculateUnitPrice(item);
        const roundedUnitPrice = roundToConfig(rawUnitPrice, 'total');
        const roundedQty = roundToConfig(item.quantity, 'qty');
        return roundedUnitPrice * roundedQty;
    };
    const moduleItems = appData.projectItems.filter(item => item.moduleId === appData.activeModuleId);
    const moduleTotal = moduleItems.reduce((sum, item) => sum + getLineTotal(item), 0);
    const grandTotal = appData.projectItems.reduce((sum, item) => sum + getLineTotal(item), 0);
    document.getElementById('module-total-display').innerText = fmt(moduleTotal, 'total');
    document.getElementById('grand-total-display').innerText = fmt(grandTotal, 'total');
}

function createEmptyItemB1() {
    const newItem = {
        id: Date.now() + Math.floor(Math.random() * 1000),
        projectCode: "",
        description: "Nuevo Ítem",
        unit: "glb",
        quantity: 1,
        materiales: [],
        mano_obra: [],
        equipos: [],
        moduleId: appData.activeModuleId
    };
    appData.projectItems.push(newItem);
    saveData();
    renderB1();
    setTimeout(() => {
        const rows = document.querySelectorAll('#b1-body tr');
        if (rows.length > 0) {
            const lastRow = rows[rows.length - 1];
            const descInput = lastRow.querySelector('textarea');
            if (descInput) { descInput.focus(); descInput.select(); }
        }
    }, 100);
}

function duplicateProjectItem(id) {
    const item = appData.projectItems.find(i => i.id === id);
    if (!item) return;
    const newItem = JSON.parse(JSON.stringify(item));
    newItem.id = Date.now() + Math.floor(Math.random() * 1000);
    const index = appData.projectItems.indexOf(item);
    appData.projectItems.splice(index + 1, 0, newItem);
    saveData();
    renderB1();
    showToast('Ítem duplicado correctamente');
}

function deleteProjectItem(id) {
    if (confirm("¿Eliminar este ítem del presupuesto?")) {
        appData.projectItems = appData.projectItems.filter(i => i.id !== id);
        saveData();
        renderB1();
    }
}

function saveProjectItemToBank(id) {
    const item = appData.projectItems.find(i => i.id === id);
    if (!item) return;
    if (confirm(`¿Guardar "${item.description}" en el Banco de Ítems?`)) {
        const newItem = JSON.parse(JSON.stringify(item));
        newItem.id = Date.now() + Math.floor(Math.random() * 1000);
        newItem.code = getNextBankCode();
        newItem.quantity = 1;
        delete newItem.moduleId;
        appData.itemBank.push(newItem);
        saveData();
        renderBankList();
        showToast(`Guardado en banco con código: ${newItem.code}`);
    }
}

function getNextBankCode() {
    let max = 0;
    appData.itemBank.forEach(i => {
        const num = parseInt(i.code);
        if (!isNaN(num) && num > max) max = num;
    });
    return max + 1;
}

function toggleB1Column(colIndex) {
    if (!appData.settings.hiddenColumnsB1) appData.settings.hiddenColumnsB1 = [];
    
    const idx = appData.settings.hiddenColumnsB1.indexOf(colIndex);
    if (idx > -1) {
        // Si ya está colapsada, la removemos del array (la expandimos)
        appData.settings.hiddenColumnsB1.splice(idx, 1);
    } else {
        // Si está expandida, la agregamos al array (la colapsamos)
        appData.settings.hiddenColumnsB1.push(colIndex);
    }
    
    saveData();
    applyColumnVisibility();
}

function applyColumnVisibility() {
    const table = document.getElementById('table-b1');
    if (!table) return;
    
    const hiddenCols = appData.settings.hiddenColumnsB1 || [];

    // 1. Aplicar a los encabezados (<th>)
    const headers = table.querySelectorAll('thead tr th');
    headers.forEach((th, index) => {
        if (hiddenCols.includes(index)) {
            th.classList.add('col-collapsed');
            th.title = "Clic para expandir";
        } else {
            th.classList.remove('col-collapsed');
            if (th.classList.contains('toggleable-col')) {
                th.title = "Clic para ocultar";
            }
        }
    });

    // 2. Aplicar a las celdas de datos (<td>)
    const rows = table.querySelectorAll('tbody tr');
    rows.forEach(row => {
        const cells = row.children;
        for (let i = 0; i < cells.length; i++) {
            if (hiddenCols.includes(i)) {
                cells[i].classList.add('col-collapsed');
            } else {
                cells[i].classList.remove('col-collapsed');
            }
        }
    });
}

// --- BANCO DE ÍTEMS ---

function renderBankList() {
    const tbody = document.getElementById('bank-body');
    const searchTerm = document.getElementById('main-bank-search')?.value.toLowerCase() || "";
    tbody.innerHTML = '';
    let filteredItems = appData.itemBank.filter(item =>
        item.description.toLowerCase().includes(searchTerm) ||
        item.code.toString().includes(searchTerm)
    );
    filteredItems.sort((a, b) => a.description.localeCompare(b.description, undefined, { numeric: true, sensitivity: 'base' }));
    const totalItems = filteredItems.length;
    const start = (currentPage.bank - 1) * itemsPerPage;
    const end = start + itemsPerPage;
    const pagedItems = filteredItems.slice(start, end);
    pagedItems.forEach((item, index) => {
        const unitPrice = calculateUnitPrice(item);
        const visualIndex = start + index + 1;
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="text-center">
                <button class="btn btn-primary btn-sm" onclick="addSingleItemToProject(${item.id})">
                    <i class="fas fa-plus"></i>
                </button>
            </td>
            <td class="text-center" style="font-weight:600;">${visualIndex}</td>
            <td>${item.description}</td>
            <td class="text-center">${item.unit}</td>
            <td class="text-right">${fmt(unitPrice, 'total')}</td>
            <td>
                <div class="actions-row">
                    <button class="btn btn-primary btn-sm" onclick="editItem('bank', ${item.id})" title="Editar">
                        <i class="fas fa-edit"></i>
                    </button>
                    <button class="btn btn-danger btn-sm" onclick="deleteBankItem(${item.id})" title="Eliminar">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });
    updatePaginationInfo('bank', totalItems);
}

function createBankItem() {
    const newItem = {
        id: Date.now() + Math.floor(Math.random() * 1000),
        code: getNextBankCode(),
        description: "Nuevo APU",
        unit: "m3",
        quantity: 1,
        materiales: [],
        mano_obra: [],
        equipos: []
    };
    appData.itemBank.push(newItem);
    saveData();
    renderBankList();
    const newItemId = appData.itemBank[appData.itemBank.length - 1].id;
    editItem('bank', newItemId, true);
    setTimeout(() => {
        const descInput = document.getElementById('apu-desc');
        if (descInput) { descInput.focus(); descInput.select(); }
    }, 300);
}

function deleteBankItem(id) {
    if (confirm("¿Eliminar este ítem del Banco?")) {
        appData.itemBank = appData.itemBank.filter(i => i.id !== id);
        saveData();
        renderBankList();
    }
}

function addSingleItemToProject(id) {
    const item = appData.itemBank.find(i => i.id === id);
    if (item) {
        let newItem = JSON.parse(JSON.stringify(item));
        newItem.id = Date.now() + Math.floor(Math.random() * 1000);
        newItem.quantity = 1;
        newItem.moduleId = appData.activeModuleId;
        appData.projectItems.push(newItem);
        saveData();
        renderB1();
        showToast('Ítem agregado');
    }
}

function calculateUnitPrice(item) {
    const s = appData.settings;
    const mats = item.materiales || [];
    const mo = item.mano_obra || [];
    const eqs = item.equipos || [];

    const sumMat = mats.reduce((a, b) => a + (b.qty * b.price), 0);
    const sumMo = mo.reduce((a, b) => a + (b.qty * b.price), 0);

    const totalMo = sumMo + (sumMo * (s.social / 100)) + ((sumMo + (sumMo * (s.social / 100))) * (s.iva_mo / 100));

    const sumEq = eqs.reduce((a, b) => a + (b.qty * b.price), 0);
    const totalEq = sumEq + (totalMo * (s.tools / 100));

    const direct = sumMat + totalMo + totalEq;
    const gg = direct * (s.gg / 100);
    const util = (direct + gg) * (s.util / 100);
    const it = (direct + gg + util) * (s.it / 100);

    return direct + gg + util + it;
}

// --- EDITOR APU ---

function editItem(context, id, isNew = false) {
    editorContext = context;
    currentEditId = id;
    isCreatingNew = isNew;
    let item = (context === 'project') ? appData.projectItems.find(i => i.id === id) : appData.itemBank.find(i => i.id === id);
    if (!item) return;
    editorBackup = JSON.parse(JSON.stringify(item));
    document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
    document.getElementById('panel-editor').classList.add('active');
    document.getElementById('editor-title').innerText = (context === 'project' ? "Editar Ítem Presup." : "Editar Ítem Banco");
    document.getElementById('apu-desc').value = item.description;
    document.getElementById('apu-unit').value = item.unit;
    document.getElementById('apu-qty').value = formatNumber(item.quantity, 'qty');
    document.getElementById('apu-qty').disabled = (context === 'bank');
    renderAPUTables(item);
}

function renderAPUTables(item) {
    const setText = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.innerText = val;
    };
    const createRows = (arr, type) => {
        let html = '';
        let sum = 0;
        if (!arr) arr = [];
        arr.forEach((r, idx) => {
            const qty = parseFloat(r.qty) || 0;
            const price = parseFloat(r.price) || 0;
            let t = qty * price;
            sum += t;
            html += `<tr>
                <td><textarea class="table-input" onchange="updateRes('${type}', ${idx}, 'desc', this.value)">${escapeHtml(r.desc)}</textarea></td>
                <td><input type="text" value="${escapeHtml(r.unit)}" onchange="updateRes('${type}', ${idx}, 'unit', this.value)"></td>
                <td><input type="text" inputmode="decimal" value="${fmt(qty, 'yield')}" onchange="updateRes('${type}', ${idx}, 'qty', handleMathInput(this, 'yield'))" onfocus="this.select()" style="text-align:center;"></td>
                <td><input type="text" inputmode="decimal" value="${fmt(price, 'price')}" onchange="updateRes('${type}', ${idx}, 'price', handleMathInput(this, 'price'))" onfocus="this.select()" style="text-align:right;"></td>
                <td class="text-right" style="font-weight:600;">${fmt(t, 'partial')}</td>
                <td class="text-center"><button class="btn btn-danger btn-sm" onclick="removeRes('${type}', ${idx})" title="Eliminar"><i class="fas fa-times"></i></button></td>
            </tr>`;
        });
        return { html, sum };
    };
    const s = appData.settings;
    const mat = createRows(item.materiales, 'materiales');
    const tableMat = document.querySelector('#table-mat tbody');
    if (tableMat) tableMat.innerHTML = mat.html;
    setText('sub-mat', fmt(mat.sum, 'partial'));

    const mo = createRows(item.mano_obra, 'mano_obra');
    const tableMo = document.querySelector('#table-mo tbody');
    if (tableMo) tableMo.innerHTML = mo.html;
    const socVal = mo.sum * (s.social / 100);
    const ivaVal = (mo.sum + socVal) * (s.iva_mo / 100);
    const totalMo = mo.sum + socVal + ivaVal;
    setText('mo-basic', fmt(mo.sum, 'partial'));
    setText('lbl-soc', s.social);
    setText('mo-soc', fmt(socVal, 'partial'));
    setText('lbl-iva', s.iva_mo);
    setText('mo-iva', fmt(ivaVal, 'partial'));
    setText('sub-mo', fmt(totalMo, 'partial'));

    const eq = createRows(item.equipos, 'equipos');
    const tableEq = document.querySelector('#table-eq tbody');
    if (tableEq) tableEq.innerHTML = eq.html;
    const toolsVal = totalMo * (s.tools / 100);
    const totalEq = eq.sum + toolsVal;
    setText('eq-basic', fmt(eq.sum, 'partial'));
    setText('lbl-tools', s.tools);
    setText('eq-tools', fmt(toolsVal, 'partial'));
    setText('sub-eq', fmt(totalEq, 'partial'));

    const costoDirect = mat.sum + totalMo + totalEq;
    setText('sum-direct', fmt(costoDirect, 'partial'));
    const ggVal = costoDirect * (s.gg / 100);
    const utilVal = (costoDirect + ggVal) * (s.util / 100);
    const itVal = (costoDirect + ggVal + utilVal) * (s.it / 100);
    setText('lbl-gg', s.gg);
    setText('sum-gg', fmt(ggVal, 'partial'));
    setText('lbl-util', s.util);
    setText('sum-util', fmt(utilVal, 'partial'));
    setText('lbl-it', s.it);
    setText('sum-it', fmt(itVal, 'partial'));

    const precioUnit = costoDirect + ggVal + utilVal + itVal;
    setText('final-price', fmt(precioUnit, 'total'));

    initSortableEditor('materiales');
    initSortableEditor('mano_obra');
    initSortableEditor('equipos');
}

function updateEditorMeta() {
    let item = (editorContext === 'project') ? appData.projectItems.find(i => i.id === currentEditId) : appData.itemBank.find(i => i.id === currentEditId);
    if (!item) return;
    item.description = document.getElementById('apu-desc').value;
    item.unit = document.getElementById('apu-unit').value;
    if (editorContext === 'project') item.quantity = parseNumber(document.getElementById('apu-qty').value);
    saveData();
}

function updateRes(type, idx, field, val) {
    let item = (editorContext === 'project') ? appData.projectItems.find(i => i.id === currentEditId) : appData.itemBank.find(i => i.id === currentEditId);
    if (!item) return;
    if (field === 'qty' || field === 'price') {
        item[type][idx][field] = parseNumber(val);
    } else {
        item[type][idx][field] = val;
    }
    renderAPUTables(item);
    saveData();
}

function removeRes(type, idx) {
    let item = (editorContext === 'project') ? appData.projectItems.find(i => i.id === currentEditId) : appData.itemBank.find(i => i.id === currentEditId);
    if (!item) return;
    if (confirm("¿Eliminar este recurso?")) {
        item[type].splice(idx, 1);
        renderAPUTables(item);
        saveData();
    }
}

function navigateEditor(direction) {
    if (!editorContext || !currentEditId) return;
    let navigationList = [];
    if (editorContext === 'project') {
        appData.modules.forEach(mod => {
            const itemsInModule = appData.projectItems.filter(i => i.moduleId === mod.id);
            navigationList = navigationList.concat(itemsInModule);
        });
    } else {
        navigationList = [...appData.itemBank].sort((a, b) => a.description.localeCompare(b.description, undefined, { numeric: true, sensitivity: 'base' }));
    }
    const currentIndex = navigationList.findIndex(item => item.id === currentEditId);
    if (currentIndex === -1) return;
    const newIndex = currentIndex + direction;
    if (newIndex < 0 || newIndex >= navigationList.length) {
        showToast(direction > 0 ? "Has llegado al final" : "Es el primer ítem");
        return;
    }
    if (hasUnsavedChanges()) {
        if (!confirm("Tienes cambios sin guardar. ¿Quieres guardar antes de navegar?")) return;
    }
    const sourceArray = editorContext === 'project' ? appData.projectItems : appData.itemBank;
    const currentRealItem = sourceArray.find(i => i.id === currentEditId);
    if (currentRealItem) {
        currentRealItem.description = document.getElementById('apu-desc').value;
        currentRealItem.unit = document.getElementById('apu-unit').value;
        if (editorContext === 'project') currentRealItem.quantity = parseNumber(document.getElementById('apu-qty').value);
    }
    const nextItem = navigationList[newIndex];
    if (editorContext === 'project' && nextItem.moduleId !== appData.activeModuleId) {
        appData.activeModuleId = nextItem.moduleId;
        showToast(`Cambiando al módulo: ${appData.modules.find(m => m.id === nextItem.moduleId)?.name}`);
        renderModuleTabs();
    }
    editItem(editorContext, nextItem.id);
}

function hasUnsavedChanges() {
    if (!editorContext || !currentEditId) return false;
    const source = editorContext === 'project' ? appData.projectItems : appData.itemBank;
    const currentItem = source.find(item => item.id === currentEditId);
    if (!currentItem || !editorBackup) return false;
    if (currentItem.description !== editorBackup.description ||
        currentItem.unit !== editorBackup.unit ||
        currentItem.quantity !== editorBackup.quantity) return true;
    const resourceTypes = ['materiales', 'mano_obra', 'equipos'];
    for (const type of resourceTypes) {
        if (currentItem[type].length !== editorBackup[type].length) return true;
        for (let i = 0; i < currentItem[type].length; i++) {
            const curr = currentItem[type][i];
            const backup = editorBackup[type][i];
            if (!curr || !backup) return true;
            if (curr.desc !== backup.desc ||
                curr.unit !== backup.unit ||
                curr.qty !== backup.qty ||
                curr.price !== backup.price) return true;
        }
    }
    return false;
}

function discardEditor() {
    if (!editorBackup) { closeEditor(); return; }
    if (confirm("¿Descartar cambios y salir?")) {
        if (isCreatingNew) {
            if (editorContext === 'project') appData.projectItems = appData.projectItems.filter(i => i.id !== currentEditId);
            else appData.itemBank = appData.itemBank.filter(i => i.id !== currentEditId);
        } else {
            if (editorContext === 'project') {
                const idx = appData.projectItems.findIndex(i => i.id === currentEditId);
                if (idx !== -1) appData.projectItems[idx] = editorBackup;
            } else {
                const idx = appData.itemBank.findIndex(i => i.id === currentEditId);
                if (idx !== -1) appData.itemBank[idx] = editorBackup;
            }
        }
        saveData();
        closeEditor();
    }
}

function closeEditor() {
    if (editorContext === 'project') {
        renderB1();
        switchTab('b1');
    } else {
        renderBankList();
        switchTab('items');
    }
    editorContext = null;
    currentEditId = null;
    editorBackup = null;
    isCreatingNew = false;
}

function quickCreateResource(type) {
    const desc = prompt("Descripción del nuevo insumo:", "");
    if (desc === null || desc.trim() === "") return;
    const unit = prompt("Unidad (ej: m3, pza, hr):", "u");
    if (unit === null) return;
    const priceStr = prompt("Precio Unitario referencial:", "0");
    if (priceStr === null) return;
    const price = parseNumber(priceStr);
    const newId = Date.now() + Math.floor(Math.random() * 100000);
    const dbItem = {
        id: newId,
        code: (appData.database[type].length + 1).toString(),
        desc: desc.trim(),
        unit: unit.trim(),
        price: price
    };
    appData.database[type].push(dbItem);
    const apuItem = {
        id: Date.now() + Math.floor(Math.random() * 100000) + 1,
        desc: desc.trim(),
        unit: unit.trim(),
        price: price,
        qty: 1
    };
    let targetItem = (editorContext === 'project') ?
        appData.projectItems.find(i => i.id === currentEditId) :
        appData.itemBank.find(i => i.id === currentEditId);
    if (targetItem) {
        let targetArray = targetItem[type];
        if (!targetArray) targetItem[type] = [];
        targetItem[type].push(apuItem);
        renderAPUTables(targetItem);
        saveData();
        showToast("Insumo creado y agregado");
        setTimeout(() => {
            let tableId = (type === 'materiales') ? 'table-mat' : (type === 'mano_obra') ? 'table-mo' : 'table-eq';
            let rows = document.querySelectorAll(`#${tableId} tbody tr`);
            if (rows.length > 0) {
                let lastRow = rows[rows.length - 1];
                let qtyInput = lastRow.querySelector('td:nth-child(3) input');
                if (qtyInput) { qtyInput.focus(); qtyInput.select(); }
            }
        }, 100);
    }
}

function openResourceModal(type) {
    modalTargetType = type;
    document.getElementById('resourceModal').style.display = 'block';
    document.getElementById('modal-search').value = '';
    filterModalList();
    document.getElementById('modal-search').focus();
}

function filterModalList() {
    const txt = document.getElementById('modal-search').value.toLowerCase();
    const tbody = document.getElementById('modal-list-body');
    tbody.innerHTML = '';
    let list = [...appData.database[modalTargetType]].sort((a, b) => a.desc.localeCompare(b.desc));
    let count = 0;
    const maxResults = 50;
    for (let i = 0; i < list.length; i++) {
        if (count >= maxResults) break;
        const item = list[i];
        if (item.desc.toLowerCase().includes(txt) || (item.code && item.code.toString().toLowerCase().includes(txt))) {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="text-center" style="font-weight:600;">${item.code || ''}</td>
                <td>${item.desc}</td>
                <td class="text-center">${item.unit}</td>
                <td class="text-right" style="font-weight:600;">${fmt(item.price, 'price')}</td>
                <td class="text-center"><button class="btn btn-primary btn-sm" title="Añadir"><i class="fas fa-plus"></i></button></td>`;
            tr.onclick = function () {
                const newItem = { desc: item.desc, unit: item.unit, price: item.price, qty: 1 };
                tryAddResource(newItem);
            };
            tbody.appendChild(tr);
            count++;
        }
    }
    if (count === 0 && txt) {
        tbody.innerHTML = `<tr><td colspan="5" class="text-center" style="padding: 20px; color: #718096;"><i class="fas fa-search"></i> No se encontraron resultados para "${txt}"</td></tr>`;
    }
}

function closeResourceModal() { document.getElementById('resourceModal').style.display = 'none'; }

function tryAddResource(newItem) {
    let tItem = (editorContext === 'project') ? appData.projectItems.find(i => i.id === currentEditId) : appData.itemBank.find(i => i.id === currentEditId);
    if (!tItem) return;
    let targetArray = (modalTargetType === 'materiales') ? tItem.materiales :
        (modalTargetType === 'mano_obra') ? tItem.mano_obra :
            tItem.equipos;
    const duplicateIndex = targetArray.findIndex(r => r.desc.toLowerCase() === newItem.desc.toLowerCase());
    if (duplicateIndex >= 0) {
        conflictData = { tItem, targetArray, newItem, existingIndex: duplicateIndex };
        document.getElementById('conflict-msg').innerHTML = `El insumo <b>"${newItem.desc}"</b> ya existe en este APU.<br><br>¿Qué deseas hacer?`;
        document.getElementById('resourceModal').style.display = 'none';
        document.getElementById('conflictModal').style.display = 'block';
    } else {
        targetArray.push(newItem);
        finishAddResource(tItem);
    }
}

function resolveConflict(action) {
    if (!conflictData) return;
    const { tItem, targetArray, newItem, existingIndex } = conflictData;
    if (action === 'replace') {
        targetArray[existingIndex].price = newItem.price;
        targetArray[existingIndex].unit = newItem.unit;
        showToast('Insumo actualizado');
        finishAddResource(tItem);
    } else {
        showToast('Operación cancelada');
        document.getElementById('conflictModal').style.display = 'none';
        document.getElementById('resourceModal').style.display = 'block';
        return;
    }
    document.getElementById('conflictModal').style.display = 'none';
}

function finishAddResource(tItem) {
    renderAPUTables(tItem);
    saveData();
    document.getElementById('conflictModal').style.display = 'none';
    document.getElementById('resourceModal').style.display = 'none';
}

// --- MODAL BANCO CON MULTI-SELECCIÓN ---

function openItemBankModalForSelect() {
    selectedBankIds.clear();
    document.getElementById('bankModal').style.display = 'block';
    document.getElementById('bank-search').value = '';
    filterBankModal();
}

function filterBankModal() {
    const txt = document.getElementById('bank-search').value.toLowerCase();
    const tbody = document.getElementById('bank-modal-body');
    const bulkBtn = document.getElementById('btn-bulk-add');
    tbody.innerHTML = '';
    document.getElementById('bank-select-all').checked = false;
    let count = 0;
    const maxResults = 50;
    for (let i = 0; i < appData.itemBank.length; i++) {
        if (count >= maxResults) break;
        const item = appData.itemBank[i];
        if (item.description.toLowerCase().includes(txt) || item.code.toString().includes(txt)) {
            const isChecked = selectedBankIds.has(item.id) ? 'checked' : '';
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="text-center">
                    <input type="checkbox" class="bank-item-cb" ${isChecked} onchange="toggleBankSelection(${item.id}, this.checked)">
                </td>
                <td class="text-center" style="font-weight:600;">${item.code}</td>
                <td>${item.description}</td>
                <td class="text-center">${item.unit}</td>
                <td class="text-center">
                    <button class="btn btn-primary btn-sm" onclick="addSingleItemToProject(${item.id})">
                        <i class="fas fa-plus"></i>
                    </button>
                </td>`;
            tbody.appendChild(tr);
            count++;
        }
    }
    updateBulkCounter();
}

function toggleBankSelection(id, isChecked) {
    if (isChecked) selectedBankIds.add(id);
    else selectedBankIds.delete(id);
    updateBulkCounter();
}

function toggleSelectAllBank(master) {
    const visibleCheckboxes = document.querySelectorAll('.bank-item-cb');
    const isChecked = master.checked;
    visibleCheckboxes.forEach(cb => {
        cb.checked = isChecked;
        const id = parseInt(cb.getAttribute('onchange').match(/\d+/)[0]);
        if (isChecked) selectedBankIds.add(id);
        else selectedBankIds.delete(id);
    });
    updateBulkCounter();
}

function updateBulkCounter() {
    const selectedCount = selectedBankIds.size;
    const btn = document.getElementById('btn-bulk-add');
    const countSpan = document.getElementById('selected-count');
    if (countSpan) countSpan.innerText = selectedCount;
    if (btn) btn.style.display = selectedCount > 0 ? 'inline-flex' : 'none';
}

function addSelectedBankItems() {
    if (selectedBankIds.size === 0) return;
    let addedCount = 0;
    selectedBankIds.forEach(id => {
        const item = appData.itemBank.find(i => i.id === id);
        if (item) {
            let newItem = JSON.parse(JSON.stringify(item));
            newItem.id = Date.now() + Math.floor(Math.random() * 10000) + addedCount;
            newItem.quantity = 1;
            newItem.moduleId = appData.activeModuleId;
            appData.projectItems.push(newItem);
            addedCount++;
        }
    });
    saveData();
    renderB1();
    closeBankModal();
    showToast(`${addedCount} ítems agregados con éxito`);
}

function closeBankModal() { document.getElementById('bankModal').style.display = 'none'; }

// --- BASE DE DATOS DE INSUMOS ---

function renderDBTables() {
    const dbSections = [
        { type: 'materiales', pageKey: 'db-mat', tableId: 'db-table-materiales', paginationId: 'db-mat-pagination', pageInfoId: 'db-mat-page-info' },
        { type: 'mano_obra', pageKey: 'db-mo', tableId: 'db-table-mano_obra', paginationId: 'db-mo-pagination', pageInfoId: 'db-mo-page-info' },
        { type: 'equipos', pageKey: 'db-eq', tableId: 'db-table-equipos', paginationId: 'db-eq-pagination', pageInfoId: 'db-eq-page-info' }
    ];
    dbSections.forEach(section => {
        const { type, pageKey, tableId, paginationId, pageInfoId } = section;
        const sortState = dbSortState[type];
        appData.database[type].sort((a, b) => {
            let valA = a[sortState.field];
            let valB = b[sortState.field];
            if (sortState.field === 'code') {
                const numA = parseFloat(valA);
                const numB = parseFloat(valB);
                if (!isNaN(numA) && !isNaN(numB)) return (numA - numB) * sortState.dir;
            }
            valA = valA ? valA.toString().toLowerCase() : "";
            valB = valB ? valB.toString().toLowerCase() : "";
            if (sortState.field === 'desc') {
                if (valA === "" && valB !== "") return -1;
                if (valA !== "" && valB === "") return 1;
            }
            if (valA < valB) return -1 * sortState.dir;
            if (valA > valB) return 1 * sortState.dir;
            return 0;
        });
        appData.database[type].forEach((item, index) => { item.code = (index + 1).toString(); });
        const totalItems = appData.database[type].length;
        const start = (currentPage[pageKey] - 1) * itemsPerPage;
        const end = Math.min(start + itemsPerPage, totalItems);
        const itemsToRender = appData.database[type].slice(start, end);
        const tbody = document.querySelector(`#${tableId} tbody`);
        tbody.innerHTML = '';
        const fragment = document.createDocumentFragment();
        itemsToRender.forEach((item) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="text-center"><input type="text" value="${item.code}" readonly title="Número de Fila (Auto)"></td>
                <td><textarea class="db-desc-input" onkeydown="handleDescEnter(event, this)" onchange="updateDbItem('${type}', ${item.id}, 'desc', this.value)" placeholder="Descripción">${escapeHtml(item.desc)}</textarea></td>
                <td><input type="text" value="${escapeHtml(item.unit)}" onchange="updateDbItem('${type}', ${item.id}, 'unit', this.value)" placeholder="u"></td>
                <td><input type="text" inputmode="decimal" value="${fmt(item.price, 'price')}" onchange="updateDbItem('${type}', ${item.id}, 'price', handleMathInput(this, 'price'))" onfocus="this.select()" style="text-align:right;"></td>
                <td class="text-center"><button class="btn btn-danger btn-sm" onclick="deleteDbItem('${type}', ${item.id})" title="Eliminar"><i class="fas fa-trash"></i></button></td>`;
            fragment.appendChild(tr);
        });
        tbody.appendChild(fragment);
        const totalPages = Math.ceil(totalItems / itemsPerPage);
        const paginationDiv = document.getElementById(paginationId);
        const pageInfo = document.getElementById(pageInfoId);
        if (totalPages > 1) {
            paginationDiv.style.display = 'flex';
            pageInfo.textContent = `Página ${currentPage[pageKey]} de ${totalPages} (${totalItems} items)`;
        } else {
            paginationDiv.style.display = 'none';
        }
        const table = document.getElementById(tableId);
        table.querySelectorAll('.sort-icon').forEach(i => i.className = 'fas fa-sort sort-icon');
        const activeHeader = table.querySelector(`.col-${sortState.field} .sort-icon`);
        if (activeHeader) {
            activeHeader.className = sortState.dir === 1 ? 'fas fa-sort-down sort-icon' : 'fas fa-sort-up sort-icon';
        }
    });
}

function toggleDbSort(type, field) {
    if (dbSortState[type].field === field) dbSortState[type].dir *= -1;
    else {
        dbSortState[type].field = field;
        dbSortState[type].dir = 1;
    }
    renderDBTables();
}

function handleDescEnter(event, textarea) {
    if (event.key === 'Enter') {
        event.preventDefault();
        const row = textarea.closest('tr');
        const unitInput = row.querySelector('td:nth-child(3) input');
        if (unitInput) { unitInput.focus(); unitInput.select(); }
    }
}

function updateDbItem(type, id, field, value) {
    const item = appData.database[type].find(i => i.id === id);
    if (item) {
        if (field === 'price') value = parseNumber(value);
        item[field] = value;
        saveData();
    }
}

function deleteDbItem(type, id) {
    if (confirm("¿Borrar este insumo de la base de datos?")) {
        appData.database[type] = appData.database[type].filter(i => i.id !== id);
        saveData();
        renderDBTables();
    }
}

function addDbRow(type) {
    const newId = Date.now() + Math.floor(Math.random() * 100000);
    appData.database[type].push({ id: newId, code: "0", desc: "", unit: "u", price: 0 });
    saveData();
    renderDBTables();
    setTimeout(() => {
        const inputs = document.querySelectorAll(`#db-table-${type} textarea`);
        for (let i = inputs.length - 1; i >= 0; i--) {
            if (inputs[i].value === "" && inputs[i].classList.contains('db-desc-input')) {
                inputs[i].focus();
                break;
            }
        }
    }, 100);
}

function cleanAndMergeDuplicates() {
    if (!confirm("Esta acción buscará insumos con exactamente la misma Descripción y Unidad, y los fusionará manteniendo el PRECIO MÁS ALTO encontrado. ¿Continuar?")) return;
    const types = ['materiales', 'mano_obra', 'equipos'];
    let totalRemoved = 0;
    types.forEach(type => {
        const originalList = appData.database[type];
        const map = new Map();
        let duplicatesInType = 0;
        originalList.forEach(item => {
            const key = (item.desc || "").trim().toLowerCase() + "||" + (item.unit || "").trim().toLowerCase();
            if (!key || key === "||") return;
            if (map.has(key)) {
                const existing = map.get(key);
                if (parseNumber(item.price) > parseNumber(existing.price)) existing.price = parseNumber(item.price);
                duplicatesInType++;
            } else {
                map.set(key, { ...item });
            }
        });
        if (duplicatesInType > 0) {
            appData.database[type] = Array.from(map.values());
            totalRemoved += duplicatesInType;
        }
    });
    if (totalRemoved > 0) {
        saveData();
        renderDBTables();
        showToast(`Se fusionaron ${totalRemoved} duplicados correctamente.`);
    } else showToast("No se encontraron duplicados.");
}

function syncDbToBank() {
    if (!confirm("Esta acción actualizará los PRECIOS de todos los ítems en el BANCO basándose en la lista actual de INSUMOS.\n\nLa coincidencia se hace por Descripción y Unidad.\n\n¿Deseas continuar?")) return;
    let updateCount = 0;
    let itemsAffected = 0;
    const priceMap = new Map();
    const types = ['materiales', 'mano_obra', 'equipos'];
    types.forEach(type => {
        appData.database[type].forEach(dbItem => {
            const key = `${type}|${dbItem.desc.trim().toLowerCase()}|${dbItem.unit.trim().toLowerCase()}`;
            priceMap.set(key, parseFloat(dbItem.price) || 0);
        });
    });
    const propMap = { 'materiales': 'materiales', 'mano_obra': 'mano_obra', 'equipos': 'equipos' };
    appData.itemBank.forEach(apu => {
        let itemUpdated = false;
        Object.keys(propMap).forEach(apuProp => {
            const dbType = propMap[apuProp];
            if (apu[apuProp] && Array.isArray(apu[apuProp])) {
                apu[apuProp].forEach(resource => {
                    const key = `${dbType}|${resource.desc.trim().toLowerCase()}|${resource.unit.trim().toLowerCase()}`;
                    if (priceMap.has(key)) {
                        const newPrice = priceMap.get(key);
                        if (Math.abs(resource.price - newPrice) > 0.001) {
                            resource.price = newPrice;
                            updateCount++;
                            itemUpdated = true;
                        }
                    }
                });
            }
        });
        if (itemUpdated) itemsAffected++;
    });
    let projectUpdatedCount = 0;
    if (updateCount > 0 && appData.projectItems.length > 0) {
        if (confirm(`Se actualizaron ${updateCount} precios en ${itemsAffected} ítems del BANCO.\n\n¿Deseas aplicar estos precios también al PRESUPUESTO (Proyecto actual)?`)) {
            appData.projectItems.forEach(apu => {
                Object.keys(propMap).forEach(apuProp => {
                    const dbType = propMap[apuProp];
                    if (apu[apuProp] && Array.isArray(apu[apuProp])) {
                        apu[apuProp].forEach(resource => {
                            const key = `${dbType}|${resource.desc.trim().toLowerCase()}|${resource.unit.trim().toLowerCase()}`;
                            if (priceMap.has(key)) {
                                const newPrice = priceMap.get(key);
                                if (Math.abs(resource.price - newPrice) > 0.001) {
                                    resource.price = newPrice;
                                    projectUpdatedCount++;
                                }
                            }
                        });
                    }
                });
            });
        }
    }
    if (updateCount > 0 || projectUpdatedCount > 0) {
        saveData();
        if (document.getElementById('panel-items').classList.contains('active')) renderBankList();
        if (document.getElementById('panel-b1').classList.contains('active')) renderB1();
        let msg = `Sincronización completa.\nBanco: ${updateCount} cambios.`;
        if (projectUpdatedCount > 0) msg += `\nPresupuesto: ${projectUpdatedCount} cambios.`;
        showToast(msg);
    } else showToast("No se encontraron diferencias de precios para actualizar.");
}

// --- IMPORTADOR EXCEL APU ---

function importComplexAPUReport(input) {
    const file = input.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async function (e) {
        try {
            const buffer = e.target.result;
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);
            const worksheet = workbook.getWorksheet(1);
            const jsonData = sheetToDataArray(worksheet);
            let importedCount = 0;
            let stats = { new: 0, updated: 0 };
            let currentItem = null;
            let currentSection = null;
            let baseCode = getNextBankCode();
            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                const rowStr = row.map(c => c ? c.toString().toUpperCase() : "").join(" ");
                let itemFound = false;
                for (let c = 0; c < row.length; c++) {
                    if (row[c] && row[c].toString().toUpperCase().includes("ITEM:")) {
                        if (currentItem && currentItem.description) { appData.itemBank.push(currentItem); importedCount++; }
                        let parts = row[c].toString().split(":");
                        let name = parts.slice(1).join(":");
                        if (name.trim() === "" && row[c + 1]) name = row[c + 1];
                        let unit = "glb";
                        for (let u = c; u < row.length; u++) {
                            if (row[u] && row[u].toString().toUpperCase().includes("UNIDAD:")) {
                                let uParts = row[u].toString().split(":");
                                let unitVal = uParts.slice(1).join(":");
                                if (unitVal.trim() === "" && row[u + 1]) unitVal = row[u + 1];
                                unit = unitVal.trim();
                            }
                        }
                        currentItem = {
                            id: Date.now() + Math.random(),
                            code: baseCode + importedCount,
                            description: name ? name.trim() : "Item Importado",
                            unit: unit,
                            quantity: 1,
                            materiales: [], mano_obra: [], equipos: [],
                            moduleId: appData.activeModuleId
                        };
                        currentSection = null; itemFound = true; break;
                    }
                }
                if (itemFound) continue;
                if (!currentItem) continue;
                if (rowStr.includes("MATERIALES") && !rowStr.includes("TOTAL")) { currentSection = 'MAT'; continue; }
                if (rowStr.includes("MANO DE OBRA") && !rowStr.includes("TOTAL")) { currentSection = 'MO'; continue; }
                if ((rowStr.includes("EQUIPO") || rowStr.includes("HERRAMIENTAS")) && !rowStr.includes("TOTAL")) { currentSection = 'EQ'; continue; }
                if (rowStr.includes("TOTAL") || rowStr.includes("SUBTOTAL")) { currentSection = null; continue; }
                if (currentSection) {
                    let desc = row[2];
                    let unit = row[3];
                    let qty = parseNumber(row[4]);
                    let price = parseNumber(row[5]);
                    if (desc && desc.toString().trim() !== "" && !isNaN(price)) {
                        let res = { desc: desc.toString().trim(), unit: unit || "u", qty: qty || 0, price: price || 0 };
                        let dbType = '';
                        if (currentSection === 'MAT') { currentItem.materiales.push(res); dbType = 'materiales'; }
                        else if (currentSection === 'MO') { currentItem.mano_obra.push(res); dbType = 'mano_obra'; }
                        else if (currentSection === 'EQ') { currentItem.equipos.push(res); dbType = 'equipos'; }
                        if (dbType) {
                            const result = syncImportedResourceToDB(dbType, res);
                            if (result === 'new') stats.new++;
                            if (result === 'updated') stats.updated++;
                        }
                    }
                }
            }
            if (currentItem && currentItem.description) { appData.itemBank.push(currentItem); importedCount++; }
            saveData();
            renderBankList();
            renderDBTables();
            if (importedCount > 0) showToast(`Importados ${importedCount} APUs.`);
            else showToast('No se encontraron ítems');
            input.value = '';
        } catch (e) {
            console.error(e);
            showToast("Error leyendo archivo");
        }
    };
    reader.readAsArrayBuffer(file);
}

function syncImportedResourceToDB(type, resource) {
    if (!appData.database[type]) return 'error';
    const list = appData.database[type];
    const resDesc = resource.desc.trim().toLowerCase();
    const resUnit = resource.unit.trim().toLowerCase();
    const resPrice = parseNumber(resource.price);
    const existingItem = list.find(item => (item.desc || "").trim().toLowerCase() === resDesc && (item.unit || "").trim().toLowerCase() === resUnit);
    if (existingItem) {
        if (resPrice > 0 && Math.abs(existingItem.price - resPrice) > 0.001) {
            existingItem.price = resPrice;
            return 'updated';
        }
        return 'exists';
    } else {
        const newCode = (list.length + 1).toString();
        list.push({ id: Date.now() + Math.floor(Math.random() * 100000), code: newCode, desc: resource.desc.trim(), unit: resource.unit.trim(), price: resPrice });
        return 'new';
    }
}

function sheetToDataArray(worksheet) {
    const data = [];
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const rowValues = Array.isArray(row.values) ? row.values.slice(1) : [];
        data.push(rowValues);
    });
    return data;
}

// --- IMPORTADOR DE VOLÚMENES DESDE EXCEL ---

function importBudgetFromExcel(input) {
    const file = input.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async function (e) {
        try {
            const buffer = e.target.result;
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);
            const worksheet = workbook.getWorksheet(1);
            const jsonData = sheetToDataArray(worksheet);
            if (jsonData.length === 0) { showToast("El archivo está vacío."); return; }
            let headerRowIndex = -1;
            let colDesc = -1, colUnit = -1, colQty = -1, colCode = -1;
            for (let i = 0; i < Math.min(jsonData.length, 20); i++) {
                if (!Array.isArray(jsonData[i])) continue;
                const row = jsonData[i].map(cell => cell ? cell.toString().toUpperCase().trim() : "");
                if (row.includes("DESCRIPCIÓN") || row.includes("DESCRIPCION")) {
                    headerRowIndex = i;
                    row.forEach((cell, idx) => {
                        if (cell.includes("CÓDIGO") || cell.includes("CODIGO")) colCode = idx;
                        if (cell.includes("DESCRIPCIÓN") || cell.includes("DESCRIPCION")) colDesc = idx;
                        if (cell.includes("UNIDAD") || cell.includes("UND")) colUnit = idx;
                        if (cell.includes("CANTIDAD") || cell.includes("TOTAL")) colQty = idx;
                    });
                    break;
                }
            }
            if (colDesc === -1 || colQty === -1) {
                showToast("Error: No se encontraron columnas 'Descripción' y 'Cantidad'.");
                return;
            }
            let itemsAdded = 0;
            let itemsMatched = 0;
            for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (!row || row.length === 0) continue;
                const descRaw = (row[colDesc] !== undefined && row[colDesc] !== null) ? row[colDesc].toString().trim() : "";
                if (descRaw === "") continue;
                const rawCode = (colCode !== -1 && row[colCode]) ? row[colCode].toString().trim() : "";
                const rawUnit = (colUnit !== -1 && row[colUnit]) ? row[colUnit].toString().trim() : "glb";
                let rawQty = 0;
                if (colQty !== -1 && row[colQty] !== undefined && row[colQty] !== null) {
                    let val = row[colQty];
                    if (typeof val === 'object' && val.result !== undefined) val = val.result;
                    rawQty = parseNumber(val);
                }
                if (isNaN(rawQty)) rawQty = 0;
                const match = findBestMatchInBank(descRaw, rawUnit);
                let newItem = match ? JSON.parse(JSON.stringify(match)) : { materiales: [], mano_obra: [], equipos: [] };
                if (match) itemsMatched++;
                newItem.id = Date.now() + Math.floor(Math.random() * 100000) + i;
                newItem.moduleId = appData.activeModuleId;
                newItem.projectCode = rawCode;
                newItem.description = descRaw;
                newItem.unit = rawUnit;
                newItem.quantity = rawQty;
                appData.projectItems.push(newItem);
                itemsAdded++;
            }
            saveData();
            renderB1();
            showToast(`Importación completa: ${itemsAdded} ítems (${itemsMatched} del Banco).`);
        } catch (err) {
            console.error(err);
            showToast("Error al procesar el archivo Excel.");
        }
        input.value = '';
    };
    reader.readAsArrayBuffer(file);
}

function normalizeStr(str) {
    return str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ').trim();
}

function findBestMatchInBank(targetDesc, targetUnit) {
    if (!appData.itemBank || appData.itemBank.length === 0) return null;
    const nTargetDesc = normalizeStr(targetDesc);
    const nTargetUnit = normalizeStr(targetUnit);
    let bestMatch = null;
    let bestScore = 0;
    const minScoreThreshold = 0.4;
    appData.itemBank.forEach(bankItem => {
        let score = 0;
        const nBankDesc = normalizeStr(bankItem.description);
        const nBankUnit = normalizeStr(bankItem.unit);
        if (nBankUnit === nTargetUnit) score += 0.3;
        else if ((nBankUnit.includes(nTargetUnit) || nTargetUnit.includes(nBankUnit)) && nTargetUnit.length > 1) score += 0.15;
        if (nBankDesc === nTargetDesc) score += 0.7;
        else {
            const targetWords = nTargetDesc.split(' ').filter(w => w.length > 2);
            if (targetWords.length > 0) {
                let matches = 0;
                targetWords.forEach(word => { if (nBankDesc.includes(word)) matches++; });
                score += (matches / targetWords.length) * 0.7;
            }
        }
        if (score > bestScore && score >= minScoreThreshold) {
            bestScore = score;
            bestMatch = bankItem;
        }
    });
    return bestMatch;
}

// --- INSUMOS DEL PRESUPUESTO ---

function collectInsumosFromProject(type) {
    insumosList = [];
    const resourceKey = type;
    const insumosMap = new Map();
    appData.projectItems.forEach(item => {
        if (item[resourceKey]) {
            item[resourceKey].forEach(resource => {
                const cantidadTotal = resource.qty * item.quantity;
                const precioTotal = cantidadTotal * resource.price;
                const key = `${resource.desc.toLowerCase()}|${resource.unit}`;
                if (insumosMap.has(key)) {
                    const entry = insumosMap.get(key);
                    entry.cantidadTotal += cantidadTotal;
                    entry.precioTotal += precioTotal;
                    if (Math.abs(entry.precioUnitario - resource.price) > 0.001) {
                        entry.tieneConflictosPrecio = true;
                        if (!entry.preciosEncontrados) entry.preciosEncontrados = [entry.precioUnitario];
                        if (!entry.preciosEncontrados.includes(resource.price)) entry.preciosEncontrados.push(resource.price);
                    }
                    entry.occurrences.push({ itemId: item.id, itemDesc: item.description, rendimiento: resource.qty, cantidadItem: item.quantity, cantidadTotal, precioUnitario: resource.price, precioTotal });
                } else {
                    insumosMap.set(key, {
                        desc: resource.desc,
                        unit: resource.unit,
                        cantidadTotal,
                        precioUnitario: resource.price,
                        precioTotal,
                        selected: false,
                        tieneConflictosPrecio: false,
                        preciosEncontrados: [resource.price],
                        occurrences: [{ itemId: item.id, itemDesc: item.description, rendimiento: resource.qty, cantidadItem: item.quantity, cantidadTotal, precioUnitario: resource.price, precioTotal }]
                    });
                }
            });
        }
    });
    insumosList = Array.from(insumosMap.values());
}

function openInsumosModal(type) {
    currentInsumosType = type;
    const tipoTexto = { 'materiales': 'Materiales', 'mano_obra': 'Mano de Obra', 'equipos': 'Equipos' };
    document.getElementById('modalTipoInsumo').textContent = tipoTexto[type];
    document.getElementById('insumosModal').style.display = 'block';
    collectInsumosFromProject(type);
    renderInsumosTable();
}

function renderInsumosTable() {
    const tbody = document.getElementById('insumosTableBody');
    tbody.innerHTML = '';
    if (insumosList.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center" style="padding:20px; color:var(--text-muted);"><i class="fas fa-info-circle"></i> No se encontraron insumos de este tipo en el presupuesto.</td></tr>';
        return;
    }
    insumosList.sort((a, b) => a.desc.localeCompare(b.desc));
    const fragment = document.createDocumentFragment();
    insumosList.forEach((insumo, index) => {
        const tr = document.createElement('tr');
        if (insumo.tieneConflictosPrecio) tr.classList.add('row-danger');
        tr.innerHTML = `
            <td class="text-center">
                <input type="checkbox" class="insumo-checkbox" ${insumo.selected ? 'checked' : ''} onchange="toggleInsumoSelection(${index}, this.checked)">
            </td>
            <td>${insumo.desc} ${insumo.tieneConflictosPrecio ? '<span style="color: #e53e3e; font-size: 0.8em;" title="Este insumo tiene diferentes precios en el presupuesto">⚠</span>' : ''}</td>
            <td class="text-center">${insumo.unit}</td>
            <td class="text-right">${fmt(insumo.cantidadTotal, 'qty')}</td>
            <td class="text-right">
                <input type="text" inputmode="decimal" class="precio-editable" value="${insumo.precioUnitario}" onfocus="this.select()" onchange="updatePrecioInsumo(${index}, handleMathInput(this, 'price'))">
            </td>
            <td class="text-right font-bold">${fmt(insumo.precioTotal, 'total')}</td>
        `;
        fragment.appendChild(tr);
    });
    const totalGeneral = insumosList.reduce((sum, item) => sum + item.precioTotal, 0);
    const trTotal = document.createElement('tr');
    trTotal.className = 'row-total'; // Usa la clase CSS
    trTotal.style.fontWeight = '700';
    trTotal.innerHTML = `
        <td colspan="5" class="text-right" style="padding:12px 10px; font-size:1.1em;">TOTAL PRESUPUESTO:</td>
        <td class="text-right" style="padding:12px 6px; font-size:1.2em;">${fmt(totalGeneral, 'total')}</td>
    `;
    fragment.appendChild(trTotal);
    tbody.appendChild(fragment);
    updateSelectAllCheckbox();
}

function updatePrecioInsumo(index, value) {
    const precio = parseNumber(value);
    insumosList[index].precioUnitario = precio;
    insumosList[index].precioTotal = insumosList[index].cantidadTotal * precio;
    const tbody = document.getElementById('insumosTableBody');
    const row = tbody.children[index];
    if (row) {
        const totalCell = row.querySelector('td:nth-child(6)');
        if (totalCell) totalCell.textContent = fmt(insumosList[index].precioTotal, 'total');
        const precioInput = row.querySelector('td:nth-child(5) input');
        if (precioInput) precioInput.value = precio;
    }
}

function toggleInsumoSelection(index, checked) {
    insumosList[index].selected = checked;
    updateSelectAllCheckbox();
}

function toggleSelectAllInsumos(checkbox) {
    const isChecked = checkbox.checked;
    insumosList.forEach((insumo, index) => { insumo.selected = isChecked; });
    renderInsumosTable();
}

function updateSelectAllCheckbox() {
    const selectAllCheckbox = document.getElementById('selectAllInsumos');
    if (!selectAllCheckbox) return;
    const total = insumosList.length;
    const selected = insumosList.filter(i => i.selected).length;
    if (selected === 0) { selectAllCheckbox.checked = false; selectAllCheckbox.indeterminate = false; }
    else if (selected === total) { selectAllCheckbox.checked = true; selectAllCheckbox.indeterminate = false; }
    else { selectAllCheckbox.checked = false; selectAllCheckbox.indeterminate = true; }
}

function aplicarPreciosEditados() {
    let cambiosTotales = 0;
    const insumosActualizados = [];
    insumosList.forEach((insumo, index) => {
        const precio = insumo.precioUnitario;
        const desc = insumo.desc;
        const unit = insumo.unit;
        let cambiosInsumo = 0;
        insumo.occurrences.forEach(ocurrencia => {
            const item = appData.projectItems.find(i => i.id === ocurrencia.itemId);
            if (item) {
                const resourceKey = currentInsumosType;
                if (item[resourceKey]) {
                    item[resourceKey].forEach(resource => {
                        if (resource.desc.toLowerCase() === desc.toLowerCase() && resource.unit === unit) {
                            resource.price = precio;
                            cambiosInsumo++;
                        }
                    });
                }
            }
        });
        if (cambiosInsumo > 0) {
            cambiosTotales += cambiosInsumo;
            insumosActualizados.push(desc);
        }
    });
    if (cambiosTotales > 0) {
        saveData();
        renderB1();
        showToast(`Precios actualizados para ${insumosActualizados.length} insumo(s) en todo el presupuesto`);
    } else showToast('No se realizaron cambios en el presupuesto');
}

function closeInsumosModal() {
    document.getElementById('insumosModal').style.display = 'none';
    insumosList = [];
}

async function exportInsumosToExcel() {
    if (insumosList.length === 0) { showToast('No hay insumos para exportar'); return; }
    const tipoTexto = { 'materiales': 'Materiales', 'mano_obra': 'Mano_de_Obra', 'equipos': 'Equipos' };
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Insumos');
    worksheet.mergeCells('A1:E1');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = `INSUMOS ${tipoTexto[currentInsumosType].toUpperCase()} - PRESUPUESTO`;
    titleCell.font = { bold: true, size: 14 };
    titleCell.alignment = { horizontal: 'center' };
    worksheet.getRow(2).values = ['Descripción', 'Unidad', 'Cantidad Total', 'Precio Unitario', 'Total'];
    const headerRow = worksheet.getRow(2);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D3748' } };
    worksheet.getColumn(1).width = 50;
    worksheet.getColumn(2).width = 10;
    worksheet.getColumn(3).width = 15;
    worksheet.getColumn(4).width = 15;
    worksheet.getColumn(5).width = 15;
    const insumosOrdenados = [...insumosList].sort((a, b) => a.desc.localeCompare(b.desc, 'es', { sensitivity: 'base' }));
    insumosOrdenados.forEach(insumo => {
        const row = worksheet.addRow([insumo.desc, insumo.unit, insumo.cantidadTotal, insumo.precioUnitario, insumo.precioTotal]);
        row.getCell(3).numFmt = '#,##0.00';
        row.getCell(4).numFmt = '#,##0.00';
        row.getCell(5).numFmt = '#,##0.00';
    });
    const totalGeneral = insumosList.reduce((sum, insumo) => sum + insumo.precioTotal, 0);
    const totalRow = worksheet.addRow(['', '', '', 'TOTAL PRESUPUESTO:', totalGeneral]);
    totalRow.font = { bold: true };
    totalRow.getCell(5).numFmt = '#,##0.00 "Bs"';
    totalRow.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBEE3F8' } };
    const buffer = await workbook.xlsx.writeBuffer();
    saveExcelBuffer(buffer, `Insumos_${tipoTexto[currentInsumosType]}_${new Date().toISOString().slice(0, 10)}.xlsx`);
    showToast('Excel exportado correctamente');
}

function saveExcelBuffer(buffer, filename) {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, filename);
}

// --- CONFIGURACIÓN ---

function openConfigModal() {
    loadSettingsToUI();
    document.getElementById('configModal').style.display = 'block';
}

function closeConfigModal() {
    document.getElementById('configModal').style.display = 'none';
}

function loadSettingsToUI() {
    document.getElementById('conf-dec-yield').value = appData.settings.decimals_yield || 5;
    document.getElementById('conf-dec-price').value = appData.settings.decimals_price || 3;
    document.getElementById('conf-dec-partial').value = appData.settings.decimals_partial || 4;
    document.getElementById('conf-dec-total').value = appData.settings.decimals_total || 2;
    document.getElementById('conf-dec-qty').value = appData.settings.decimals_qty || 2;
    document.getElementById('conf-social').value = appData.settings.social;
    document.getElementById('conf-iva-mo').value = appData.settings.iva_mo;
    document.getElementById('conf-tools').value = appData.settings.tools;
    document.getElementById('conf-gg').value = appData.settings.gg;
    document.getElementById('conf-util').value = appData.settings.util;
    document.getElementById('conf-it').value = appData.settings.it;
    if (appData.settings.numberFormat) document.getElementById('conf-format').value = appData.settings.numberFormat;
    updatePrecisionIndicator();
}

function saveSettings() {
    appData.settings.decimals_yield = parseInt(document.getElementById('conf-dec-yield').value) || 5;
    appData.settings.decimals_price = parseInt(document.getElementById('conf-dec-price').value) || 3;
    appData.settings.decimals_partial = parseInt(document.getElementById('conf-dec-partial').value) || 4;
    appData.settings.decimals_total = parseInt(document.getElementById('conf-dec-total').value) || 2;
    appData.settings.decimals_qty = parseInt(document.getElementById('conf-dec-qty').value) || 2;
    appData.settings.social = parseFloat(document.getElementById('conf-social').value) || 0;
    appData.settings.iva_mo = parseFloat(document.getElementById('conf-iva-mo').value) || 0;
    appData.settings.tools = parseFloat(document.getElementById('conf-tools').value) || 0;
    appData.settings.gg = parseFloat(document.getElementById('conf-gg').value) || 0;
    appData.settings.util = parseFloat(document.getElementById('conf-util').value) || 0;
    appData.settings.it = parseFloat(document.getElementById('conf-it').value) || 0;
    appData.settings.numberFormat = document.getElementById('conf-format').value;
    saveData();
    updatePrecisionIndicator();
    renderB1();
    renderBankList();
    renderDBTables();
    renderInsumosTable();
    recalcCalcTable();
    if (currentEditId) {
        let item = (editorContext === 'project') ? appData.projectItems.find(i => i.id === currentEditId) : appData.itemBank.find(i => i.id === currentEditId);
        if (item) renderAPUTables(item);
    }
    showToast('Formato numérico actualizado en todas las pestañas');
}

function updatePrecisionIndicator() {
    const indicator = document.getElementById('precision-indicator');
    if (indicator) {
        const s = appData.settings;
        indicator.innerHTML = `<i class="fas fa-sliders-h"></i> Rend:${s.decimals_yield} | Prec:${s.decimals_price} | Total:${s.decimals_total}`;
        indicator.title = `Configuración de decimales:\nRendimientos: ${s.decimals_yield} decimales\nPrecios: ${s.decimals_price} decimales\nParciales: ${s.decimals_partial} decimales\nTotales: ${s.decimals_total} decimales\nCantidades: ${s.decimals_qty} decimales`;
    }
}

// --- EXPORTACIÓN A EXCEL B1 ---

async function exportB1ToExcel() {
    if (typeof JSZip === 'undefined') { showToast("Error: Librería JSZip no cargada."); return; }
    const pName = document.getElementById('reportProjectName')?.value || "PROYECTO";
    const zip = new JSZip();
    showToast("Generando archivos con 5 decimales...");
    const n = (num) => {
        if (num === null || num === undefined || isNaN(num)) return 0;
        return parseFloat(Number(num).toFixed(5));
    };
    const borderStyle = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    const styleHeader = { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D3748' } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: borderStyle };
    const styleCell = { border: borderStyle, alignment: { vertical: 'middle' } };
    const styleItemFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F4F8' } };
    const getProjectInsumos = (type) => {
        const map = new Map();
        appData.projectItems.forEach(item => {
            let resources = [];
            if (type === 'mat') resources = item.materiales || [];
            else if (type === 'mo') resources = item.mano_obra || [];
            else resources = item.equipos || [];
            resources.forEach(r => {
                const key = r.desc.trim().toUpperCase();
                if (map.has(key)) {
                    const existing = map.get(key);
                    existing.qty += (r.qty * item.quantity);
                    existing.totalPrice += (r.qty * r.price * item.quantity);
                } else {
                    map.set(key, { desc: r.desc, unit: r.unit, price: r.price, qty: r.qty * item.quantity, totalPrice: r.qty * r.price * item.quantity });
                }
            });
        });
        return Array.from(map.values());
    };
    const specs = [
        { type: 'mat', name: 'B3-MATERIALES.xlsx', sheet: 'B3-MATERIALES', title: 'MATERIALES' },
        { type: 'mo', name: 'B3-MANO_DE_OBRA.xlsx', sheet: 'Hoja1', title: 'MANO DE OBRA' },
        { type: 'eq', name: 'B3-MAQUINARIA_Y_EQUIPO.xlsx', sheet: 'Hoja1', title: 'EQUIPO, MAQUINARIA Y HERRAMIENTAS' }
    ];
    for (const spec of specs) {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet(spec.sheet);
        ws.addRow(["", spec.title, "", ""]);
        ws.mergeCells('B1:D1');
        ws.getCell('B1').font = { bold: true, size: 12 };
        const hRow = ws.addRow(["Nº", "DESCRIPCIÓN", "UNIDAD", "PRECIO UNITARIO"]);
        for (let i = 1; i <= 4; i++) hRow.getCell(i).style = styleHeader;
        ws.getColumn(1).width = 5; ws.getColumn(2).width = 50; ws.getColumn(3).width = 15; ws.getColumn(4).width = 20;
        const data = getProjectInsumos(spec.type);
        const dataOrdenada = [...data].sort((a, b) => a.desc.localeCompare(b.desc, 'es', { sensitivity: 'base' }));
        dataOrdenada.forEach((ins, i) => {
            const row = ws.addRow([i + 1, ins.desc, ins.unit, n(ins.price)]);
            row.eachCell(c => Object.assign(c, styleCell));
            row.getCell(4).numFmt = '0.00000';
        });
        const starRow = ws.addRow(["*", "", "", ""]);
        for (let i = 1; i <= 4; i++) starRow.getCell(i).border = borderStyle;
        const buffer = await wb.xlsx.writeBuffer();
        zip.file(spec.name, buffer);
    }
    const wbForm = new ExcelJS.Workbook();
    const forms = [
        { sheet: 'PPTO. MATERIALES', title: 'MATERIALES-INSUMOS', key: 'materiales' },
        { sheet: 'PPTO. MANO DE OBRA', title: 'MANO DE OBRA-INSUMOS', key: 'mano_obra' },
        { sheet: 'PPTO. MAQUINARIA Y EQUIPO', title: 'EQUIPO, MAQUINARIA Y HERRAMIENTAS-INSUMOS', key: 'equipos' }
    ];
    forms.forEach(f => {
        const ws = wbForm.addWorksheet(f.sheet);
        ws.addRow(["", f.title]);
        const hRow = ws.addRow(["Código", "Descripción", "Unidad", "Cantidad", "Precio Unitario", "Precio Total (Bs)"]);
        for (let i = 1; i <= 6; i++) hRow.getCell(i).style = styleHeader;
        ws.getColumn(1).width = 10; ws.getColumn(2).width = 45; ws.getColumn(3).width = 10; ws.getColumn(4).width = 15; ws.getColumn(5).width = 15; ws.getColumn(6).width = 18;
        appData.projectItems.forEach(item => {
            const code = item.projectCode ? `>${item.projectCode}` : `>${item.code || ''}`;
            const itemRow = ws.addRow([code, item.description, "", "", "", ""]);
            for (let c = 1; c <= 6; c++) { const cell = itemRow.getCell(c); cell.fill = styleItemFill; cell.font = { bold: true }; cell.border = borderStyle; }
            const itemQty = n(item.quantity);
            (item[f.key] || []).forEach((r, idx) => {
                const totalQty = n(r.qty * itemQty);
                const unitPrice = n(r.price);
                const totalPrice = n(totalQty * unitPrice);
                const row = ws.addRow([idx + 1, r.desc, r.unit, totalQty, unitPrice, totalPrice]);
                row.eachCell(c => Object.assign(c, styleCell));
                row.getCell(4).numFmt = '0.00000';
                row.getCell(5).numFmt = '0.00000';
                row.getCell(6).numFmt = '0.00000';
            });
        });
        const starRow = ws.addRow(["*", "", "", "", "", ""]);
        for (let c = 1; c <= 6; c++) starRow.getCell(c).border = borderStyle;
    });
    const bufferForm = await wbForm.xlsx.writeBuffer();
    zip.file("FORM_CANTIDADES.xlsx", bufferForm);
    showToast("Comprimiendo ZIP...");
    zip.generateAsync({ type: "blob" }).then(function (content) {
        const dateStr = new Date().toISOString().slice(0, 10);
        saveAs(content, `Reportes_SICOES_5DEC_${pName}_${dateStr}.zip`);
        showToast("Descarga lista");
    });
}

// --- EXPORTACIÓN JSON ---

function exportProjectJSON() {
    const s = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(appData));
    const dl = document.createElement('a'); dl.href = s;
    const ahora = new Date(); const dia = String(ahora.getDate()).padStart(2, '0'); const mes = String(ahora.getMonth() + 1).padStart(2, '0'); const año = ahora.getFullYear();
    dl.download = `Proy_Save_${dia}${mes}${año}.sure`;
    dl.click();
    showToast('Proyecto guardado');
}

function openSaveOptionsModal() {
    document.getElementById('saveOptionsModal').style.display = 'block';
}

function exportBudgetOnlyJSON() {
    const budgetData = { settings: appData.settings, modules: appData.modules, activeModuleId: appData.activeModuleId, projectItems: appData.projectItems, database: { materiales: [], mano_obra: [], equipos: [] }, itemBank: [] };
    const s = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(budgetData));
    const dl = document.createElement('a'); dl.href = s;
    const ahora = new Date(); const dia = String(ahora.getDate()).padStart(2, '0'); const mes = String(ahora.getMonth() + 1).padStart(2, '0'); const año = ahora.getFullYear();
    dl.download = `PG_${dia}${mes}${año}.sure`;
    dl.click();
    showToast('Presupuesto exportado (sin BD)');
}

function loadProjectJSON(input) {
    const file = input.files[0];
    if (!file) return;
    const fileName = file.name.toLowerCase();
    const isValidExtension = fileName.endsWith('.sure') || fileName.endsWith('.json');
    const isValidMimeType = file.type === 'application/json' || file.type === 'text/json' || file.type === 'text/plain' || file.type === '';
    if (!isValidExtension && !isValidMimeType) { showToast('Por favor, seleccione un archivo .sure o .json válido.'); input.value = ''; return; }
    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const loadedData = JSON.parse(e.target.result);
            appData = { ...appData, ...loadedData };
            if (!appData.activeModuleId && appData.modules.length > 0) appData.activeModuleId = appData.modules[0].id;
            if (!appData.database) appData.database = { materiales: [], mano_obra: [], equipos: [] };
            saveData();
            renderB1(); renderBankList(); renderDBTables(); loadSettingsToUI();
            switchTab('b1');
            showToast("Proyecto cargado correctamente");
        } catch (err) {
            console.error(err); showToast("Error al leer el archivo JSON. Verifique que el archivo sea válido.");
        }
        input.value = '';
    };
    reader.readAsText(file);
}

function exportBankJSON() {
    const s = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(appData.itemBank));
    const dl = document.createElement('a'); dl.href = s;
    const ahora = new Date(); const dia = String(ahora.getDate()).padStart(2, '0'); const mes = String(ahora.getMonth() + 1).padStart(2, '0'); const año = ahora.getFullYear();
    dl.download = `BD_${dia}${mes}${año}.sure`;
    dl.click();
    showToast('Banco exportado');
}

function importBankJSON(input) {
    const r = new FileReader();
    r.onload = function (e) {
        try {
            let loaded = JSON.parse(e.target.result);
            if (Array.isArray(loaded)) {
                let apuCount = 0; let insumosCount = 0;
                loaded.forEach((i, index) => {
                    i.code = getNextBankCode();
                    i.id = Date.now() + Math.floor(Math.random() * 1000) + index;
                    appData.itemBank.push(i);
                    apuCount++;
                    if (i.materiales && Array.isArray(i.materiales)) i.materiales.forEach(res => { const result = syncImportedResourceToDB('materiales', res); if (result === 'new') insumosCount++; });
                    if (i.mano_obra && Array.isArray(i.mano_obra)) i.mano_obra.forEach(res => { const result = syncImportedResourceToDB('mano_obra', res); if (result === 'new') insumosCount++; });
                    if (i.equipos && Array.isArray(i.equipos)) i.equipos.forEach(res => { const result = syncImportedResourceToDB('equipos', res); if (result === 'new') insumosCount++; });
                });
                saveData();
                renderBankList(); renderDBTables();
                showToast(`Importado: ${apuCount} APUs y ${insumosCount} insumos nuevos`);
            }
        } catch (e) {
            console.error(e); showToast("Error al leer el archivo JSON o formato incorrecto");
        }
        input.value = '';
    };
    r.readAsText(input.files[0]);
}

function importDBFromExcel() {
    const f = document.getElementById('dbFile').files[0];
    const type = document.getElementById('dbImportType').value;
    if (!f) { showToast('Seleccione un archivo primero'); return; }
    const r = new FileReader();
    r.onload = async function (e) {
        try {
            const buffer = e.target.result;
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);
            const worksheet = workbook.getWorksheet(1);
            const json = sheetToDataArray(worksheet);
            let c = 0;
            for (let i = 1; i < json.length; i++) {
                let row = json[i];
                if (row && row.length >= 2 && row[1]) {
                    const desc = row[1] ? row[1].toString().trim() : "";
                    const unit = row[2] ? row[2].toString().trim() : "u";
                    let price = parseFloat(row[3]);
                    if (isNaN(price)) price = parseFloat(row[4]);
                    if (isNaN(price)) price = 0;
                    appData.database[type].push({ id: Date.now() + Math.floor(Math.random() * 100000) + c, code: "0", desc, unit, price });
                    c++;
                }
            }
            if (c > 0) { saveData(); renderDBTables(); showToast(`Importados ${c} insumos`); }
            else showToast('No se encontraron datos válidos');
            document.getElementById('dbFile').value = '';
        } catch (err) { console.error(err); showToast('Error al procesar el archivo'); }
    };
    r.readAsArrayBuffer(f);
}

// --- BORRADO SELECTIVO ---

function openDeleteModal() { document.getElementById('deleteOptionsModal').style.display = 'block'; }
function closeDeleteModal() { document.getElementById('deleteOptionsModal').style.display = 'none'; }

function clearBudgetOnly() {
    if (confirm("¿Estás seguro de vaciar SOLAMENTE los ítems del presupuesto actual?\n\nTu Banco de Ítems y Base de Insumos NO se borrarán.")) {
        appData.projectItems = [];
        const defaultId = 'mod_' + Date.now();
        appData.modules = [{ id: defaultId, name: "General" }];
        appData.activeModuleId = defaultId;
        saveData();
        renderB1();
        recalculateTotalsDisplay();
        closeDeleteModal();
        showToast("Presupuesto vaciado correctamente");
    }
}

function confirmFullReset() {
    closeDeleteModal();
    setTimeout(() => { resetData(); }, 100);
}

async function resetData() {
    if (!confirm("⚠️ ATENCIÓN ⚠️\nEsta acción realizará un BORRADO DE FÁBRICA:\n\n1. Eliminará todos los proyectos y bancos de ítems.\n2. Borrará la configuración personal.\n3. Limpiará la Memoria Caché y Cookies de la aplicación.\n\n¿Estás absolutamente seguro de continuar?")) return;
    showToast("Iniciando limpieza profunda...");
    try { localStorage.clear(); sessionStorage.clear(); } catch (e) { console.error("Error limpiando Storage", e); }
    try { await localforage.clear(); localStorage.clear(); sessionStorage.clear(); } catch (e) { console.error("Error limpiando Storage", e); }
    try { document.cookie.split(";").forEach((c) => { document.cookie = c.replace(/^ +/, "").replace(/=.*/, "=;expires=" + new Date().toUTCString() + ";path=/"); }); } catch (e) { console.error("Error limpiando Cookies", e); }
    if ('caches' in window) { try { const keys = await caches.keys(); await Promise.all(keys.map(key => caches.delete(key))); } catch (e) { console.error("Error limpiando Cache Storage", e); } }
    if ('serviceWorker' in navigator) { try { const registrations = await navigator.serviceWorker.getRegistrations(); for (const registration of registrations) await registration.unregister(); } catch (e) { console.error("Error desregistrando SW", e); } }
    showToast("Limpieza completada. Reiniciando sistema...");
    setTimeout(() => { window.location.reload(true); }, 1500);
}

// --- CALCULADORA --
let currentCalcItemData = null;

function searchCalcItem() {
    const txt = document.getElementById('calc-search-input').value.toLowerCase();
    const resultsDiv = document.getElementById('calc-search-results');
    if (txt.length < 2) { resultsDiv.style.display = 'none'; return; }
    resultsDiv.innerHTML = '';
    const matches = appData.itemBank.filter(i => i.description.toLowerCase().includes(txt) || i.code.toString().includes(txt));
    if (matches.length === 0) resultsDiv.innerHTML = '<div style="padding:10px; color:var(--text-muted);">No hay coincidencias en el Banco.</div>';
    else {
        matches.forEach(item => {
            const div = document.createElement('div');
            div.style.cssText = 'padding:8px 10px; cursor:pointer; border-bottom:1px solid #f7fafc; hover:background:#f7fafc;';
            div.innerHTML = `<strong>${item.code}</strong> - ${item.description} <small>(${item.unit})</small>`;
            div.className = 'search-result-item'; // Agregaremos este CSS abajo
            div.onclick = function () { selectItemForCalc(item); resultsDiv.style.display = 'none'; document.getElementById('calc-search-input').value = ''; };
            resultsDiv.appendChild(div);
        });
    }
    resultsDiv.style.display = 'block';
}

function selectItemForCalc(item) {
    currentCalcItemData = JSON.parse(JSON.stringify(item));

    // Forzamos la actualización visual
    const nameEl = document.getElementById('calc-selected-name');
    const unitEl = document.getElementById('calc-selected-unit');

    if (nameEl) nameEl.textContent = item.description;
    if (unitEl) unitEl.textContent = item.unit;

    // Opcional: Poner también el nombre en el buscador para que el usuario sepa qué buscó
    document.getElementById('calc-search-input').value = item.description;

    recalcCalcTable();
}

function recalcCalcTable() {
    const tbody = document.getElementById('calc-body');
    tbody.innerHTML = '';
    if (!currentCalcItemData) return;
    const inputQty = parseFloat(document.getElementById('calc-quantity').value) || 0;
    const createCalcRow = (res, type, index) => {
        const totalBase = res.qty * inputQty;
        const descLower = res.desc.toLowerCase();
        if (res.calcFactor === undefined) {
            if (type === 'materiales') { if (descLower.includes('cemento') || descLower.includes('yeso')) res.calcFactor = 50; else res.calcFactor = 1; }
            else if (type === 'mano_obra' || type === 'equipos') { res.calcFactor = 8; }
            else res.calcFactor = 1;
        }
        let finalValue = totalBase;
        let labelSuffix = "";
        let textClass = "text-calc-material"; // Por defecto (verde/azul según tema)

        if (res.calcFactor && parseFloat(res.calcFactor) !== 0) {
            finalValue = totalBase / res.calcFactor;
            if (parseFloat(res.calcFactor) !== 1) {
                if (type === 'materiales') {
                    if (descLower.includes('cemento') || descLower.includes('yeso')) {
                        labelSuffix = " Bolsas";
                        textClass = "text-calc-special"; // Púrpura
                    }
                }
                else if (type === 'mano_obra' || type === 'equipos') {
                    labelSuffix = " Días";
                    textClass = "text-calc-work"; // Naranja/Amarillo
                }
            }
        }
        return `<tr><td>${res.desc}</td><td class="text-center">${res.unit}</td><td><input type="text" inputmode="decimal" value="${fmt(res.qty, 'yield')}" onchange="updateCalcValue('${type}', ${index}, 'qty', handleMathInput(this, 'yield'))" style="width:100%; text-align:center; border:1px dashed #cbd5e0; background:transparent;"></td>
        <td><input type="text" inputmode="decimal" value="${fmt(res.calcFactor, 'yield')}" onchange="updateCalcValue('${type}', ${index}, 'calcFactor', handleMathInput(this, 'yield'))" onfocus="this.select()" style="width:100%; text-align:center; border:1px solid #e2e8f0; font-weight:bold; color:var(--primary);"></td>
        <td class="text-right font-bold ${textClass}" style="font-size:1.05rem;">${fmt(finalValue, 'qty')}${labelSuffix}</td>
        </tr>`;

    };
    let html = '';
    const sections = [{ id: 'materiales', l: 'Materiales', i: 'fa-box' }, { id: 'mano_obra', l: 'Mano de Obra', i: 'fa-users' }, { id: 'equipos', l: 'Equipos', i: 'fa-truck-monster' }];
    sections.forEach(sec => {

        if (currentCalcItemData[sec.id] && currentCalcItemData[sec.id].length > 0) {
            html += `<tr class="calc-category-row">
                <td colspan="5" class="calc-category-cell">
                    <i class="fas ${sec.i}"></i> ${sec.l}
                </td>
            </tr>`;
            currentCalcItemData[sec.id].forEach((r, i) => html += createCalcRow(r, sec.id, i));
        }
    });
    tbody.innerHTML = html || `<tr><td colspan="5" class="text-center">No hay insumos.</td></tr>`;
}

function updateCalcValue(type, index, field, newVal) {
    if (!currentCalcItemData) return;
    currentCalcItemData[type][index][field] = parseFloat(newVal) || 0;
    recalcCalcTable();
}

function printCalcReport() {
    if (!currentCalcItemData) return alert("Selecciona un ítem primero.");
    const printWindow = window.open('', '_blank');
    const qty = document.getElementById('calc-quantity').value;
    const date = new Date().toLocaleDateString();
    let tableRows = '';
    const rows = document.querySelectorAll('#calc-body tr');
    rows.forEach(row => {
        if (row.style.background.includes('rgb(237, 242, 247)')) tableRows += `<tr><td colspan="4" style="background:#f1f5f9; font-weight:bold; padding:10px; border:1px solid #e2e8f0;">${row.innerText.toUpperCase()}</td></tr>`;
        else {
            const cells = row.querySelectorAll('td');
            if (cells.length >= 5) {
                const desc = cells[0].innerText;
                const unit = cells[1].innerText;
                const yieldVal = cells[2].querySelector('input').value;
                const totalText = cells[4].innerText;
                tableRows += `<tr><td style="padding:10px; border:1px solid #e2e8f0;">${desc}</td><td style="padding:10px; border:1px solid #e2e8f0; text-align:center;">${unit}</td><td style="padding:10px; border:1px solid #e2e8f0; text-align:center;">${yieldVal}</td><td style="padding:10px; border:1px solid #e2e8f0; text-align:right; font-weight:bold;">${totalText}</td></tr>`;
            }
        }
    });
    printWindow.document.write(`
        <html><head><title>Reporte de Insumos - PresuRE</title><style>body{font-family:'Segoe UI',sans-serif;color:#1a202c;padding:40px;}.header{border-bottom:3px solid #1a365d;margin-bottom:30px;padding-bottom:10px;}table{width:100%;border-collapse:collapse;margin-bottom:30px;}th{background:#1a365d;color:white;padding:12px;text-align:left;text-transform:uppercase;font-size:12px;}.footer{text-align:center;font-size:10px;color:#a0aec0;margin-top:50px;border-top:1px solid #e2e8f0;padding-top:10px;}</style></head><body><div class="header"><h2 style="margin:0;color:#1a365d;">LISTA DE REQUERIMIENTO DE INSUMOS</h2><p style="margin:5px 0;font-size:14px;">Generado por PresuRE</p></div><div style="margin-bottom:20px;background:#f8fafc;padding:15px;border-radius:8px;"><strong>ACTIVIDAD/ÍTEM:</strong> ${currentCalcItemData.description}<br><strong>CANTIDAD PROGRAMADA:</strong> ${qty} ${currentCalcItemData.unit}<br><strong>FECHA DE REPORTE:</strong> ${date}</div><table><thead><tr><th>Descripción del Insumo</th><th style="text-align:center;">Unid. Base</th><th style="text-align:center;">Rend.</th><th style="text-align:right;">Total Calculado</th></tr></thead><tbody>${tableRows}</tbody></table><div class="footer">Este documento es una guía de cálculo para despacho de materiales en obra.</div></body></html>
    `);
    printWindow.document.close();
    setTimeout(() => { printWindow.print(); }, 250);
}

document.addEventListener('click', function (e) {
    const searchContainer = document.getElementById('calc-search-input').parentElement.parentElement;
    const results = document.getElementById('calc-search-results');
    if (!searchContainer.contains(e.target) && results) results.style.display = 'none';
});

// --- REPORTES ---

let currentReportFormat = 'pdf';

function openReportModal() {
    const input = document.getElementById('reportProjectName');

    // Valor por defecto si está vacío
    if (input && !input.value) input.value = "CONSTRUCCIÓN...";

    document.getElementById('reportModal').style.display = 'block';

    // Lógica de Foco: Pequeño retraso para asegurar que el modal sea visible
    setTimeout(() => {
        if (input) {
            input.select(); // Selecciona todo el texto para facilitar el reemplazo rápido
        }
    }, 100);
}

function closeReportModal() { document.getElementById('reportModal').style.display = 'none'; }

function setReportFormat(format) {
    currentReportFormat = format;
    document.querySelectorAll('.format-option').forEach(el => el.classList.remove('active'));
    document.querySelector(`.format-option input[value="${format}"]`).parentElement.classList.add('active');
}

function generateSelectedReport(type) {
    const pName = document.getElementById('reportProjectName').value || "PROYECTO SIN NOMBRE";
    showToast("Generando reporte...");
    setTimeout(() => {
        if (currentReportFormat === 'pdf') generatePDFReportSICOES(type, pName);
        else generateExcelReportSICOES(type, pName);
        closeReportModal();
    }, 200);
}

function generatePDFReportSICOES(type, projectName) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    const dateStr = new Date().toLocaleDateString();
    const f = (num, t) => formatNumber(num, t);
    const s = appData.settings;
    const tableStyles = { cellPadding: 1.5, minCellHeight: 6, fontSize: 8, valign: 'middle' };
    const headerStyles = { fillColor: [230, 230, 230], textColor: 0, fontStyle: 'bold', lineWidth: 0.1, lineColor: 150, halign: 'center' };
    const addPagePG = () => {
        doc.setFontSize(14); doc.setFont("helvetica", "bold");
        doc.text("FORMULARIO B-1", 105, 15, { align: "center" });
        doc.text("PRESUPUESTO GENERAL DE OBRA", 105, 22, { align: "center" });
        doc.setFontSize(9); doc.setFont("helvetica", "normal");
        doc.text(`PROYECTO: ${projectName}`, 14, 32);
        doc.text(`FECHA: ${dateStr}`, 195, 32, { align: "right" });
        const tableBody = [];
        let grandTotal = 0;
        let itemGlobalCounter = 1;
        appData.modules.forEach(mod => {
            tableBody.push([{ content: mod.name.toUpperCase(), colSpan: 6, styles: { fillColor: [245, 245, 245], fontStyle: 'bold' } }]);
            const modItems = appData.projectItems.filter(i => i.moduleId === mod.id);
            modItems.forEach(item => {
                const rawPu = calculateUnitPrice(item);
                const pu = roundToConfig(rawPu, 'total');
                const qty = roundToConfig(item.quantity, 'qty');
                const total = pu * qty;
                grandTotal += total;
                tableBody.push([itemGlobalCounter++, item.description, item.unit, f(qty, 'qty'), f(pu, 'total'), f(total, 'total')]);
            });
        });
        tableBody.push([{ content: "TOTAL PRESUPUESTO", colSpan: 5, styles: { fontStyle: 'bold', halign: 'right' } }, { content: f(grandTotal, 'total'), styles: { fontStyle: 'bold' } }]);
        doc.autoTable({ startY: 35, head: [['Nº', 'DESCRIPCIÓN', 'UND.', 'CANTIDAD', 'P. UNITARIO', 'TOTAL']], body: tableBody, theme: 'grid', headStyles: headerStyles, styles: tableStyles, columnStyles: { 0: { halign: 'left', cellWidth: 10 }, 2: { halign: 'center' }, 3: { halign: 'right' }, 4: { halign: 'right' }, 5: { halign: 'right' } } });
    };
    const addPageAPUs = () => {
        const apuColumnStyles = { 0: { cellWidth: 95 }, 1: { cellWidth: 15, halign: 'center' }, 2: { cellWidth: 22, halign: 'right' }, 3: { cellWidth: 25, halign: 'right' }, 4: { cellWidth: 25, halign: 'right' } };
        let itemCounter = 1;
        appData.modules.forEach(mod => {
            const modItems = appData.projectItems.filter(i => i.moduleId === mod.id);
            modItems.forEach(item => {
                doc.addPage();
                let y = 15;
                doc.setFontSize(12); doc.setFont("helvetica", "bold");
                doc.text("FORMULARIO B-2", 105, y, { align: "center" }); y += 6;
                doc.text("ANÁLISIS DE PRECIOS UNITARIOS", 105, y, { align: "center" }); y += 8;
                doc.setDrawColor(0); doc.setFillColor(255, 255, 255); doc.rect(14, y, 182, 16);
                doc.setFontSize(9); doc.setFont("helvetica", "normal");
                doc.text(`PROYECTO: ${projectName}`, 16, y + 5);
                doc.text(`ÍTEM ${itemCounter++}: ${item.description}`, 16, y + 11);
                doc.text(`UNIDAD: ${item.unit}`, 140, y + 5);
                doc.text(`CANTIDAD: ${f(item.quantity, 'qty')}`, 140, y + 11);
                y += 20;
                const matArr = item.materiales || [];
                const moArr = item.mano_obra || [];
                const eqArr = item.equipos || [];
                const sumMat = matArr.reduce((a, b) => a + (b.qty * b.price), 0);
                const sumMo = moArr.reduce((a, b) => a + (b.qty * b.price), 0);
                const valSoc = sumMo * (s.social / 100);
                const valIvaMo = (sumMo + valSoc) * (s.iva_mo / 100);
                const totalMo = sumMo + valSoc + valIvaMo;
                const sumEq = eqArr.reduce((a, b) => a + (b.qty * b.price), 0);
                const valTools = totalMo * (s.tools / 100);
                const totalEq = sumEq + valTools;
                const costoDirecto = sumMat + totalMo + totalEq;
                const valGG = costoDirecto * (s.gg / 100);
                const valUtil = (costoDirecto + valGG) * (s.util / 100);
                const valIT = (costoDirecto + valGG + valUtil) * (s.it / 100);
                const precioFinal = costoDirecto + valGG + valUtil + valIT;
                const bodyMat = matArr.map(r => [r.desc, r.unit, f(r.qty, 'yield'), f(r.price, 'price'), f(r.qty * r.price, 'partial')]);
                bodyMat.push([{ content: 'SUBTOTAL MATERIALES', colSpan: 4, styles: { halign: 'right', fontStyle: 'bold' } }, f(sumMat, 'partial')]);
                doc.autoTable({ startY: y, head: [['MATERIALES', 'UND.', 'REND.', 'PRECIO', 'PARCIAL']], body: bodyMat, theme: 'grid', headStyles: headerStyles, styles: tableStyles, columnStyles: apuColumnStyles, margin: { left: 14, right: 14 } });
                y = doc.lastAutoTable.finalY + 4;
                const bodyMo = moArr.map(r => [r.desc, r.unit, f(r.qty, 'yield'), f(r.price, 'price'), f(r.qty * r.price, 'partial')]);
                bodyMo.push([{ content: 'SUBTOTAL MANO DE OBRA', colSpan: 4, styles: { halign: 'right', fontStyle: 'bold' } }, f(sumMo, 'partial')]);
                bodyMo.push([{ content: `Cargas Sociales (${s.social}%)`, colSpan: 4, styles: { halign: 'right', textColor: 100 } }, f(valSoc, 'partial')]);
                bodyMo.push([{ content: `IVA MO (${s.iva_mo}%)`, colSpan: 4, styles: { halign: 'right', textColor: 100 } }, f(valIvaMo, 'partial')]);
                bodyMo.push([{ content: 'TOTAL MANO DE OBRA', colSpan: 4, styles: { halign: 'right', fontStyle: 'bold', fillColor: [245, 245, 245] } }, { content: f(totalMo, 'partial'), styles: { fontStyle: 'bold' } }]);
                doc.autoTable({ startY: y, head: [['MANO DE OBRA', 'UND.', 'REND.', 'PRECIO', 'PARCIAL']], body: bodyMo, theme: 'grid', headStyles: headerStyles, styles: tableStyles, columnStyles: apuColumnStyles, margin: { left: 14, right: 14 } });
                y = doc.lastAutoTable.finalY + 4;
                const bodyEq = eqArr.map(r => [r.desc, r.unit, f(r.qty, 'yield'), f(r.price, 'price'), f(r.qty * r.price, 'partial')]);
                bodyEq.push([{ content: 'SUBTOTAL EQUIPO', colSpan: 4, styles: { halign: 'right', fontStyle: 'bold' } }, f(sumEq, 'partial')]);
                bodyEq.push([{ content: `Herramientas Menores (${s.tools}% de MO)`, colSpan: 4, styles: { halign: 'right', textColor: 100 } }, f(valTools, 'partial')]);
                bodyEq.push([{ content: 'TOTAL EQUIPO Y MAQUINARIA', colSpan: 4, styles: { halign: 'right', fontStyle: 'bold', fillColor: [245, 245, 245] } }, { content: f(totalEq, 'partial'), styles: { fontStyle: 'bold' } }]);
                if (y > 220) { doc.addPage(); y = 20; }
                doc.autoTable({ startY: y, head: [['EQUIPO Y MAQUINARIA', 'UND.', 'REND.', 'PRECIO', 'PARCIAL']], body: bodyEq, theme: 'grid', headStyles: headerStyles, styles: tableStyles, columnStyles: apuColumnStyles, margin: { left: 14, right: 14 } });
                y = doc.lastAutoTable.finalY + 4;
                if (y > 230) { doc.addPage(); y = 20; }
                doc.autoTable({ startY: y, body: [['COSTO DIRECTO', f(costoDirecto, 'total')], [`Gastos Generales (${s.gg}%)`, f(valGG, 'total')], [`Utilidad (${s.util}%)`, f(valUtil, 'total')], [`Impuesto IT (${s.it}%)`, f(valIT, 'total')], ['PRECIO UNITARIO TOTAL', f(precioFinal, 'total')]], theme: 'plain', styles: { ...tableStyles, cellPadding: 1 }, columnStyles: { 0: { halign: 'right', fontStyle: 'bold', cellWidth: 157 }, 1: { halign: 'right', fontStyle: 'bold', cellWidth: 25, fillColor: [240, 240, 240] } }, margin: { left: 14, right: 14 } });
            });
        });
    };
    const addPageInsumos = (cat) => {
        collectInsumosFromProjectGlobal(cat);
        doc.addPage();
        doc.setFontSize(14); doc.setFont("helvetica", "bold");
        doc.text(`LISTADO DE ${cat.toUpperCase()}`, 105, 20, { align: 'center' });
        doc.setFontSize(10); doc.setFont("helvetica", "normal");
        doc.text(`PROYECTO: ${projectName}`, 105, 26, { align: "center" });
        doc.text(`FECHA: ${dateStr}`, 195, 26, { align: "right" });
        const body = insumosGlobalList.map(i => [i.desc, i.unit, f(i.cantidadTotal, 'qty'), f(i.precioUnitario, 'price'), f(i.precioTotal, 'total')]);
        const total = insumosGlobalList.reduce((a, b) => a + b.precioTotal, 0);
        body.push([{ content: 'TOTAL', colSpan: 4, styles: { halign: 'right', fontStyle: 'bold' } }, { content: f(total, 'total'), styles: { fontStyle: 'bold' } }]);
        doc.autoTable({ startY: 32, head: [['DESCRIPCIÓN', 'UND.', 'CANT.', 'P. UNIT', 'TOTAL']], body: body, theme: 'grid', headStyles: headerStyles, styles: tableStyles, columnStyles: { 1: { halign: 'center' }, 2: { halign: 'right' }, 3: { halign: 'right' }, 4: { halign: 'right' } } });
    };
    if (type === 'masivo') { addPagePG(); addPageAPUs(); addPageInsumos('materiales'); addPageInsumos('mano_obra'); addPageInsumos('equipos'); doc.save(`PROYECTO_${projectName}_COMPLETO.pdf`); }
    else if (type === 'pg') { addPagePG(); doc.save(`B1_${projectName}.pdf`); }
    else if (type === 'apu') { doc.deletePage(1); addPageAPUs(); doc.save(`B2_${projectName}.pdf`); }
    else { doc.deletePage(1); addPageInsumos(type); doc.save(`Insumos_${type}.pdf`); }
}

async function generateExcelReportSICOES(type, projectName) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "PresuRE App";
    const s = appData.settings;
    const n = (num, t) => {
        let d = 2;
        if (t === 'qty') d = s.decimals_qty;
        if (t === 'total') d = s.decimals_total;
        if (t === 'price') d = s.decimals_price;
        if (t === 'yield') d = s.decimals_yield;
        if (t === 'partial') d = s.decimals_partial;
        return parseFloat(parseFloat(num).toFixed(d));
    };
    const buildB1 = () => {
        const ws = workbook.addWorksheet("B-1");
        ws.columns = [{ width: 5 }, { width: 45 }, { width: 10 }, { width: 12 }, { width: 12 }, { width: 15 }];
        ws.addRow(["FORMULARIO B-1"]).font = { bold: true, size: 14 };
        ws.addRow(["PRESUPUESTO GENERAL DE OBRA"]).font = { bold: true };
        ws.addRow([`PROYECTO: ${projectName}`]);
        ws.addRow([]);
        const header = ws.addRow(["Nº", "DESCRIPCIÓN", "UNIDAD", "CANTIDAD", "P. UNITARIO", "TOTAL"]);
        for(let i = 1; i <= 6; i++) {
            header.getCell(i).font = { bold: true, color: { argb: 'FFFFFFFF' } };
            header.getCell(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
        }
        let c = 1, gt = 0;
        appData.modules.forEach(mod => {
            const rMod = ws.addRow([mod.name.toUpperCase(), "", "", "", "", ""]);
            ws.mergeCells(rMod.number, 1, rMod.number, 6); // Opcional: fusiona las celdas para que se vea mejor
            for(let i = 1; i <= 6; i++) {
                rMod.getCell(i).font = { bold: true, italic: true };
                rMod.getCell(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
            }
            appData.projectItems.filter(i => i.moduleId === mod.id).forEach(item => {
                const pu = calculateUnitPrice(item);
                const puR = roundToConfig(pu, 'total');
                const qtyR = roundToConfig(item.quantity, 'qty');
                const tot = puR * qtyR;
                gt += tot;
                ws.addRow([c++, item.description, item.unit, n(qtyR, 'qty'), n(puR, 'total'), n(tot, 'total')]);
            });
        });
        const rTot = ws.addRow(["", "TOTAL PRESUPUESTO", "", "", "", n(gt, 'total')]);
        rTot.font = { bold: true };
        rTot.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
    };
    const buildAPU = () => {
        const ws = workbook.addWorksheet("B-2");
        ws.columns = [{ width: 40 }, { width: 10 }, { width: 12 }, { width: 12 }, { width: 15 }];
        ws.addRow(["FORMULARIO B-2"]).font = { bold: true, size: 14 };
        ws.addRow(["ANÁLISIS DE PRECIOS UNITARIOS"]);
        ws.addRow([`PROYECTO: ${projectName}`]);
        ws.addRow([]);
        let c = 1;
        appData.modules.forEach(mod => {
            appData.projectItems.filter(i => i.moduleId === mod.id).forEach(item => {
                const matArr = item.materiales || [];
                const moArr = item.mano_obra || [];
                const eqArr = item.equipos || [];
                const sumMat = matArr.reduce((a, b) => a + (b.qty * b.price), 0);
                const sumMo = moArr.reduce((a, b) => a + (b.qty * b.price), 0);
                const valSoc = sumMo * (s.social / 100);
                const valIvaMo = (sumMo + valSoc) * (s.iva_mo / 100);
                const totalMo = sumMo + valSoc + valIvaMo;
                const sumEq = eqArr.reduce((a, b) => a + (b.qty * b.price), 0);
                const valTools = totalMo * (s.tools / 100);
                const totalEq = sumEq + valTools;
                const cd = sumMat + totalMo + totalEq;
                const valGG = cd * (s.gg / 100);
                const valUtil = (cd + valGG) * (s.util / 100);
                const valIT = (cd + valGG + valUtil) * (s.it / 100);
                let finalP = cd + valGG + valUtil + valIT;
                finalP = roundToConfig(finalP, 'total');
                // Agregamos un texto vacío al final para la columna 5
                const rHead = ws.addRow(["ÍTEM " + c++, item.description, "UNIDAD: " + item.unit, "CANT: " + n(item.quantity, 'qty'), ""]);
                for(let i = 1; i <= 5; i++) {
                    rHead.getCell(i).font = { bold: true };
                    rHead.getCell(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDBEAFE' } };
                }
                ws.addRow(["DESCRIPCIÓN", "UNIDAD", "RENDIMIENTO", "PRECIO", "PARCIAL"]).font = { italic: true };
                ws.addRow(["MATERIALES"]).font = { bold: true, underline: true };
                matArr.forEach(r => ws.addRow(["  " + r.desc, r.unit, n(r.qty, 'yield'), n(r.price, 'price'), n(r.qty * r.price, 'partial')]));
                ws.addRow(["SUBTOTAL MATERIALES", "", "", "", n(sumMat, 'partial')]).font = { bold: true };
                ws.addRow(["MANO DE OBRA"]).font = { bold: true, underline: true };
                moArr.forEach(r => ws.addRow(["  " + r.desc, r.unit, n(r.qty, 'yield'), n(r.price, 'price'), n(r.qty * r.price, 'partial')]));
                ws.addRow(["SUBTOTAL MANO DE OBRA", "", "", "", n(sumMo, 'partial')]);
                ws.addRow([`CARGAS SOCIALES (${s.social}%)`, "", "", "", n(valSoc, 'partial')]).font = { color: { argb: 'FF718096' } };
                ws.addRow([`IVA MO (${s.iva_mo}%)`, "", "", "", n(valIvaMo, 'partial')]).font = { color: { argb: 'FF718096' } };
                ws.addRow(["TOTAL MANO DE OBRA", "", "", "", n(totalMo, 'partial')]).font = { bold: true };
                ws.addRow(["EQUIPO Y MAQUINARIA"]).font = { bold: true, underline: true };
                eqArr.forEach(r => ws.addRow(["  " + r.desc, r.unit, n(r.qty, 'yield'), n(r.price, 'price'), n(r.qty * r.price, 'partial')]));
                ws.addRow(["SUBTOTAL EQUIPO", "", "", "", n(sumEq, 'partial')]);
                ws.addRow([`HERRAMIENTAS MENORES (${s.tools}% de MO)`, "", "", "", n(valTools, 'partial')]).font = { color: { argb: 'FF718096' } };
                ws.addRow(["TOTAL EQUIPO", "", "", "", n(totalEq, 'partial')]).font = { bold: true };
                ws.addRow(["COSTO DIRECTO", "", "", "", n(cd, 'total')]);
                ws.addRow([`GASTOS GENERALES (${s.gg}%)`, "", "", "", n(valGG, 'total')]);
                ws.addRow([`UTILIDAD (${s.util}%)`, "", "", "", n(valUtil, 'total')]);
                ws.addRow([`IMPUESTO IT (${s.it}%)`, "", "", "", n(valIT, 'total')]);
                const rFin = ws.addRow(["TOTAL PRECIO UNITARIO", "", "", "", n(finalP, 'total')]);
                for(let i = 1; i <= 5; i++) {
                    rFin.getCell(i).font = { bold: true, size: 11 };
                    rFin.getCell(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
                }
                ws.addRow([]);
            });
        });
    };
    const buildInsumos = (cat, sheetName) => {
        collectInsumosFromProjectGlobal(cat);
        const ws = workbook.addWorksheet(sheetName);
        ws.columns = [{ width: 40 }, { width: 10 }, { width: 15 }, { width: 15 }, { width: 15 }];
        ws.addRow([`LISTADO DE ${cat.toUpperCase()}`]).font = { bold: true, size: 14 };
        ws.addRow([`PROYECTO: ${projectName}`]);
        ws.addRow([]);
        const head = ws.addRow(["DESCRIPCIÓN", "UNIDAD", "CANT. TOTAL", "PRECIO UNIT.", "TOTAL"]);
        for(let i = 1; i <= 5; i++) {
            head.getCell(i).font = { bold: true, color: { argb: 'FFFFFFFF' } };
            head.getCell(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
        }
        let t = 0;
        insumosGlobalList.forEach(i => { t += i.precioTotal; ws.addRow([i.desc, i.unit, n(i.cantidadTotal, 'qty'), n(i.precioUnitario, 'price'), n(i.precioTotal, 'total')]); });
        const rTot = ws.addRow(["COSTO TOTAL", "", "", "", n(t, 'total')]);
        rTot.font = { bold: true };
        rTot.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBAE6FD' } };
    };
    if (type === 'masivo') { buildB1(); buildAPU(); buildInsumos('materiales', "Materiales"); buildInsumos('mano_obra', "Mano de Obra"); buildInsumos('equipos', "Equipos"); const buffer = await workbook.xlsx.writeBuffer(); saveExcelBuffer(buffer, `PROYECTO_${projectName}_COMPLETO.xlsx`); }
    else if (type === 'pg') { buildB1(); const buffer = await workbook.xlsx.writeBuffer(); saveExcelBuffer(buffer, "B1.xlsx"); }
    else if (type === 'apu') { buildAPU(); const buffer = await workbook.xlsx.writeBuffer(); saveExcelBuffer(buffer, "B2.xlsx"); }
    else { buildInsumos(type, type); const buffer = await workbook.xlsx.writeBuffer(); saveExcelBuffer(buffer, `Insumos_${type}.xlsx`); }
}

let insumosGlobalList = [];
function collectInsumosFromProjectGlobal(type) {
    insumosGlobalList = [];
    const resourceKey = type;
    const insumosMap = new Map();
    appData.projectItems.forEach(item => {
        if (item[resourceKey]) {
            item[resourceKey].forEach(resource => {
                const cantidadTotal = resource.qty * item.quantity;
                const precioTotal = cantidadTotal * resource.price;
                const key = `${resource.desc.toLowerCase()}|${resource.unit}`;
                if (insumosMap.has(key)) {
                    const entry = insumosMap.get(key);
                    entry.cantidadTotal += cantidadTotal;
                    entry.precioTotal += precioTotal;
                } else {
                    insumosMap.set(key, { desc: resource.desc, unit: resource.unit, cantidadTotal, precioUnitario: resource.price, precioTotal });
                }
            });
        }
    });
    insumosGlobalList = Array.from(insumosMap.values()).sort((a, b) => a.desc.localeCompare(b.desc));
}

// --- SORTABLE (DRAG & DROP) ---

function initSortableB1() {
    const el = document.getElementById('b1-body');
    if (!el) return;
    if (el._sortable) el._sortable.destroy();
    el._sortable = Sortable.create(el, {
        animation: 150, handle: 'tr', filter: 'input, textarea, button, select, i', preventOnFilter: false,
        ghostClass: 'sortable-ghost', delay: 200, delayOnTouchOnly: true,
        onEnd: function (evt) {
            if (evt.oldIndex === evt.newIndex) return;
            const rows = Array.from(document.querySelectorAll('#b1-body tr'));
            const visibleIds = rows.map(r => parseFloat(r.getAttribute('data-id'))).filter(id => !isNaN(id));
            const newOrderMap = new Map();
            visibleIds.forEach((id, index) => newOrderMap.set(id, index));
            const activeModuleId = appData.activeModuleId;
            const otherModuleItems = appData.projectItems.filter(i => i.moduleId !== activeModuleId);
            let activeModuleItems = appData.projectItems.filter(i => i.moduleId === activeModuleId);
            activeModuleItems.sort((a, b) => {
                const idxA = newOrderMap.has(a.id) ? newOrderMap.get(a.id) : -1;
                const idxB = newOrderMap.has(b.id) ? newOrderMap.get(b.id) : -1;
                if (idxA === -1) return 1;
                if (idxB === -1) return -1;
                return idxA - idxB;
            });
            appData.projectItems = [...otherModuleItems, ...activeModuleItems];
            saveData();
            rows.forEach((row, index) => {
                const offset = (currentPage.b1 - 1) * itemsPerPage;
                row.cells[1].innerText = offset + index + 1;
            });
        }
    });
}

function initSortableEditor(type) {
    const config = { 'materiales': { id: 'table-mat', prop: 'materiales' }, 'mano_obra': { id: 'table-mo', prop: 'mano_obra' }, 'equipos': { id: 'table-eq', prop: 'equipos' } };
    const conf = config[type];
    const table = document.getElementById(conf.id);
    if (!table) return;
    const el = table.querySelector('tbody');
    if (el._sortable) el._sortable.destroy();
    el._sortable = Sortable.create(el, {
        animation: 150, filter: 'input, textarea, button, select', preventOnFilter: false,
        ghostClass: 'sortable-ghost', delay: 200, delayOnTouchOnly: true,
        onEnd: function (evt) {
            let item = (editorContext === 'project') ? appData.projectItems.find(i => i.id === currentEditId) : appData.itemBank.find(i => i.id === currentEditId);
            if (!item) return;
            const movedRes = item[conf.prop].splice(evt.oldIndex, 1)[0];
            item[conf.prop].splice(evt.newIndex, 0, movedRes);
            saveData();
            renderAPUTables(item);
        }
    });
}

// --- UNIFICACIÓN DE INSUMOS ---

let pendingUnificationList = [];

function initiateUnification() {
    const selected = insumosList.filter(i => i.selected);
    if (selected.length < 2) { showToast("Selecciona al menos 2 insumos para fusionar."); return; }
    pendingUnificationList = selected;
    const container = document.getElementById('unify-list-container');
    document.getElementById('unify-count').textContent = selected.length;
    container.innerHTML = '';
    selected.forEach((item, index) => {
        let apuListHTML = '';
        if (item.occurrences && item.occurrences.length > 0) {
            apuListHTML = item.occurrences.map(occ => `<li><i class="fas fa-caret-right" style="color:var(--accent)"></i> ${occ.itemDesc} <span style="color:var(--text-muted)">(${fmt(occ.rendimiento, 'yield')})</span></li>`).join('');
        } else apuListHTML = '<li>No se encontraron usos registrados.</li>';
        const div = document.createElement('div');
        div.style.marginBottom = "10px";
        div.innerHTML = `
            <div class="unify-option" style="margin-bottom:0; border-bottom-left-radius:0; border-bottom-right-radius:0;">
                <div class="unify-details" style="flex-grow:1;">
                    <span class="unify-name">${item.desc}</span>
                    <div class="unify-meta">
                        <span><i class="fas fa-ruler-combined"></i> ${item.unit}</span>
                        <span><i class="fas fa-tag"></i> ${fmt(item.precioUnitario, 'price')}</span>
                    </div>
                    <div class="usage-badge" onclick="toggleUnifyDetails('usage-${index}', event)">
                        <i class="fas fa-list-ul"></i> Ver en qué ítems se usa (${item.occurrences.length}) <i class="fas fa-chevron-down"></i>
                    </div>
                </div>
                <button class="btn btn-success btn-sm" onclick="executeUnification(${index})" style="margin-left:10px;">Mantener este <i class="fas fa-check"></i></button>
            </div>
            <ul id="usage-${index}" class="apu-usage-list" style="margin-top:0; border-top:none; border-top-left-radius:0; border-top-right-radius:0;">${apuListHTML}</ul>
        `;
        container.appendChild(div);
    });
    document.getElementById('unifyModal').style.display = 'block';
}

function toggleUnifyDetails(elementId, event) {
    if (event) event.stopPropagation();
    const list = document.getElementById(elementId);
    if (list.style.display === 'block') list.style.display = 'none';
    else {
        document.querySelectorAll('.apu-usage-list').forEach(el => el.style.display = 'none');
        list.style.display = 'block';
    }
}

function closeUnifyModal() {
    document.getElementById('unifyModal').style.display = 'none';
    pendingUnificationList = [];
}

function executeUnification(masterIndex) {
    if (!confirm("¿Estás seguro? Esta acción no se puede deshacer.")) return;
    const master = pendingUnificationList[masterIndex];
    const victims = pendingUnificationList.filter((_, idx) => idx !== masterIndex);
    let changesCount = 0;
    const getKey = (desc, unit) => `${desc.trim().toLowerCase()}|${unit.trim().toLowerCase()}`;
    const victimKeys = new Set(victims.map(v => getKey(v.desc, v.unit)));
    appData.projectItems.forEach(apu => {
        const type = currentInsumosType;
        const resources = apu[type];
        if (resources && Array.isArray(resources)) {
            resources.forEach(res => {
                const currentKey = getKey(res.desc, res.unit);
                if (victimKeys.has(currentKey)) {
                    res.desc = master.desc;
                    res.unit = master.unit;
                    res.price = master.precioUnitario;
                    changesCount++;
                }
            });
        }
    });
    saveData();
    closeUnifyModal();
    collectInsumosFromProject(currentInsumosType);
    renderInsumosTable();
    renderB1();
    showToast(`¡Fusión completada! Se actualizaron ${changesCount} ocurrencias.`);
}

// --- ATAJOS DE TECLADO ---

function setupEditModeShortcuts() {
    document.addEventListener('keydown', function (e) {
        if (!document.getElementById('panel-editor').classList.contains('active')) return;
        if (e.ctrlKey) {
            switch (e.key) {
                case 'ArrowLeft': e.preventDefault(); navigateEditor(-1); break;
                case 'ArrowRight': e.preventDefault(); navigateEditor(1); break;
                case 's': e.preventDefault(); closeEditor(); break;
                case 'Escape': e.preventDefault(); discardEditor(); break;
            }
        }
        if (!currentEditField) return;
        switch (e.key) {
            case 'Enter': e.preventDefault(); saveCurrentField(); moveToNextField(); break;
            case 'Escape': e.preventDefault(); cancelEdit(); break;
            case 'Tab': e.preventDefault(); moveToNextField(); break;
            case 'ArrowUp': e.preventDefault(); navigateRows(-1); break;
            case 'ArrowDown': e.preventDefault(); navigateRows(1); break;
        }
    });
    document.addEventListener('focusin', function (e) {
        if (e.target.tagName === 'TEXTAREA' || e.target.tagName === 'INPUT') {
            const field = e.target;
            const row = field.closest('tr');
            const table = field.closest('table');
            if (table && table.id) {
                currentEditField = field;
                currentEditRow = row;
                if (table.id === 'table-mat') currentEditType = 'materiales';
                else if (table.id === 'table-mo') currentEditType = 'mano_obra';
                else if (table.id === 'table-eq') currentEditType = 'equipos';
                else currentEditType = null;
                setTimeout(() => { if (field.value && field.tagName === 'INPUT') field.select(); }, 10);
            }
        }
    });
    document.addEventListener('focusout', function (e) {
        currentEditField = null;
        currentEditRow = null;
        currentEditType = null;
    });
}

function setupGlobalShortcuts() {
    document.addEventListener('keydown', function (e) {
        // Ignoramos si el usuario está escribiendo en un input o textarea
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;

        // Atajos SOLO con la tecla Alt (Alt + 1, 2, 3, 4)
        // Nos aseguramos de que Ctrl y Shift NO estén presionadas
        if (e.altKey && !e.ctrlKey && !e.shiftKey) {
            let tabToSwitch = null;

            switch (e.key) {
                case '1': tabToSwitch = 'b1'; break;     // Presupuesto
                case '2': tabToSwitch = 'items'; break;  // Ítems
                case '3': tabToSwitch = 'db'; break;     // Base de Datos
                case '4': tabToSwitch = 'calc'; break;   // Calculadora
            }

            if (tabToSwitch) {
                e.preventDefault(); // Evitamos acciones secundarias del navegador
                switchTab(tabToSwitch);
            }
        }
    });
}

function saveCurrentField() {
    if (!currentEditField || !currentEditType) return;
    const field = currentEditField;
    const value = field.value;
    const rows = Array.from(currentEditRow.parentElement.children);
    const rowIndex = rows.indexOf(currentEditRow);
    const item = (editorContext === 'project') ? appData.projectItems.find(i => i.id === currentEditId) : appData.itemBank.find(i => i.id === currentEditId);
    if (!item) return;
    const cellIndex = Array.from(currentEditRow.children).indexOf(field.parentElement);
    const fieldMap = { 0: 'desc', 1: 'unit', 2: 'qty', 3: 'price' };
    const fieldName = fieldMap[cellIndex];
    if (fieldName) {
        if (fieldName === 'qty' || fieldName === 'price') item[currentEditType][rowIndex][fieldName] = solveMathExpression(value);
        else item[currentEditType][rowIndex][fieldName] = value;
        renderAPUTables(item);
        saveData();
    }
}

function moveToNextField() {
    if (!currentEditField || !currentEditRow) return;
    const cells = Array.from(currentEditRow.children);
    const currentCell = currentEditField.parentElement;
    const currentIndex = cells.indexOf(currentCell);
    let nextCell;
    if (currentIndex < cells.length - 2) nextCell = cells[currentIndex + 1];
    else {
        const rows = Array.from(currentEditRow.parentElement.children);
        const rowIndex = rows.indexOf(currentEditRow);
        if (rowIndex < rows.length - 1) nextCell = rows[rowIndex + 1].children[0];
        else {
            addResourceToCurrentTable();
            setTimeout(() => {
                const newRows = Array.from(currentEditRow.parentElement.children);
                if (newRows.length > rows.length) {
                    const lastRow = newRows[newRows.length - 1];
                    const firstCell = lastRow.children[0];
                    const textarea = firstCell.querySelector('textarea');
                    if (textarea) textarea.focus();
                }
            }, 50);
            return;
        }
    }
    if (nextCell) {
        const input = nextCell.querySelector('input, textarea');
        if (input) { input.focus(); if (input.tagName === 'INPUT') input.select(); }
    }
}

function cancelEdit() {
    if (!currentEditField) return;
    currentEditField.blur();
    currentEditField = null;
    currentEditRow = null;
    currentEditType = null;
}

function navigateRows(direction) {
    if (!currentEditRow || !currentEditType) return;
    const rows = Array.from(currentEditRow.parentElement.children);
    const rowIndex = rows.indexOf(currentEditRow);
    const newIndex = rowIndex + direction;
    if (newIndex >= 0 && newIndex < rows.length) {
        const newRow = rows[newIndex];
        const currentCellIndex = Array.from(currentEditRow.children).indexOf(currentEditField.parentElement);
        const newCell = newRow.children[currentCellIndex];
        const input = newCell.querySelector('input, textarea');
        if (input) { input.focus(); if (input.tagName === 'INPUT') input.select(); }
    }
}

function addResourceToCurrentTable() {
    if (!currentEditType) return;
    const item = (editorContext === 'project') ? appData.projectItems.find(i => i.id === currentEditId) : appData.itemBank.find(i => i.id === currentEditId);
    if (!item) return;
    item[currentEditType].push({ desc: "", unit: "u", qty: 1, price: 0 });
    renderAPUTables(item);
    saveData();
}

// --- MODO OSCURO (v2.0) ---

function loadTheme() {
    const savedTheme = localStorage.getItem('theme');
    const systemPrefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    if (savedTheme === 'dark' || (!savedTheme && systemPrefersDark)) {
        document.documentElement.setAttribute('data-theme', 'dark');
        updateThemeIcon(true);
    } else {
        document.documentElement.removeAttribute('data-theme');
        updateThemeIcon(false);
    }
}

function toggleTheme() {
    const currentTheme = document.documentElement.getAttribute('data-theme');
    if (currentTheme === 'dark') {
        document.documentElement.removeAttribute('data-theme');
        localStorage.setItem('theme', 'light');
        updateThemeIcon(false);
    } else {
        document.documentElement.setAttribute('data-theme', 'dark');
        localStorage.setItem('theme', 'dark');
        updateThemeIcon(true);
    }
}

function updateThemeIcon(isDark) {
    const icon = document.querySelector('#theme-toggle i');
    if (icon) {
        if (isDark) { icon.classList.remove('fa-moon'); icon.classList.add('fa-sun'); }
        else { icon.classList.remove('fa-sun'); icon.classList.add('fa-moon'); }
    }
}

// --- PWA INSTALL ---
function checkPwaInstall() {
    window.addEventListener('beforeinstallprompt', (e) => {
        e.preventDefault();
        deferredPrompt = e;
        document.getElementById('install-btn').classList.remove('hidden');
    });
    document.getElementById('install-btn').addEventListener('click', async () => {
        if (!deferredPrompt) return;
        deferredPrompt.prompt();
        const { outcome } = await deferredPrompt.userChoice;
        if (outcome === 'accepted') document.getElementById('install-btn').classList.add('hidden');
        deferredPrompt = null;
    });
    const isIos = /iPhone|iPad|iPod/.test(navigator.userAgent) && !window.MSStream;
    const isStandalone = window.navigator.standalone || window.matchMedia('(display-mode: standalone)').matches;
    if (isIos && !isStandalone) showIosInstallToast();
}

function showIosInstallToast() {
    setTimeout(() => {
        const toast = document.getElementById('pwa-toast');
        toast.innerHTML = `
            <div class="pwa-content">
                <i class="fab fa-apple" style="font-size: 24px;"></i>
                <div style="display:flex; flex-direction:column; align-items:flex-start;">
                    <span>Instalar en <strong>iPhone</strong>:</span>
                    <span style="font-size:12px; opacity:0.9;">Toca <i class="fas fa-share-square"></i> y elige "Agregar a Inicio" <i class="fas fa-plus-square"></i></span>
                </div>
                <button class="pwa-close" aria-label="Cerrar">&times;</button>
            </div>
        `;
        toast.classList.add('show');
        toast.querySelector('.pwa-close').addEventListener('click', () => toast.classList.remove('show'));
    }, 3000);
}

// --- INICIALIZACIÓN ---

window.onload = async function () {
    await loadFromStorage();

    initializeDefaultSettings();
    renderB1();
    renderBankList();
    renderDBTables();
    loadSettingsToUI();
    setupEditModeShortcuts();
    setupGlobalShortcuts();
    updatePrecisionIndicator();
    loadTheme();
    checkPwaInstall();
    document.getElementById('theme-toggle').addEventListener('click', toggleTheme);
    // --- NUEVO: Restaurar pestaña activa tras recargar ---
    const savedTab = localStorage.getItem('presure_active_tab');
    if (savedTab) {
        switchTab(savedTab);
    }
    // --- Lógica Avanzada Header Scroll & Nav Sticky (Definitiva) ---
    let lastScrollTop = 0;
    const headerElement = document.querySelector('.app-header');
    const navElement = document.querySelector('.app-nav');

    // Umbral para evitar rebotes (hysteresis)
    const scrollThreshold = 5;

    function handleScroll() {
        // 1. Si es móvil (< 769px), NO HACEMOS NADA. 
        // El CSS con !important se encarga de forzarlo abajo.
        if (window.innerWidth < 769) return;

        const currentScroll = window.pageYOffset || document.documentElement.scrollTop;

        // Protección contra rebote en iOS (scroll negativo)
        if (currentScroll <= 0) {
            headerElement.classList.remove('header-hidden');
            if (navElement) navElement.classList.remove('sticky-mode');
            lastScrollTop = 0;
            return;
        }

        // Si estamos muy arriba (cerca del tope), reseteamos todo a la posición original
        // 60px es aprox la altura del header
        if (currentScroll < 60) {
            headerElement.classList.remove('header-hidden');
            if (navElement) navElement.classList.remove('sticky-mode');
            lastScrollTop = currentScroll;
            return;
        }

        // Detectar dirección del scroll solo si supera el umbral
        if (Math.abs(currentScroll - lastScrollTop) > scrollThreshold) {
            if (currentScroll > lastScrollTop) {
                // SCROLL HACIA ABAJO ->
                // 1. Ocultar Header
                headerElement.classList.add('header-hidden');
                // 2. Subir Nav (activar clase sticky-mode que pone top: 15px)
                if (navElement) navElement.classList.add('sticky-mode');
            } else {
                // SCROLL HACIA ARRIBA ->
                // 1. Mostrar Header
                headerElement.classList.remove('header-hidden');
                // 2. Bajar Nav (quitar clase sticky-mode para volver a top: 75px)
                if (navElement) navElement.classList.remove('sticky-mode');
            }
            lastScrollTop = currentScroll;
        }
    }

    // Usar 'passive: true' mejora el rendimiento del scroll
    window.addEventListener('scroll', handleScroll, { passive: true });

    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('./sw.js');
    }

    // Escuchar cambios de tema del sistema operativo en tiempo real
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', e => {
        // Solo cambiamos si el usuario no ha fijado un tema manualmente en localStorage
        if (!localStorage.getItem('theme')) {
            if (e.matches) {
                document.documentElement.setAttribute('data-theme', 'dark');
                updateThemeIcon(true);
            } else {
                document.documentElement.removeAttribute('data-theme');
                updateThemeIcon(false);
            }
        }
    });

    // --- LÓGICA DE APERTURA DESDE EL SISTEMA OPERATIVO (File Handling API) ---
    if ('launchQueue' in window) {
        window.launchQueue.setConsumer(async (launchParams) => {
            if (!launchParams.files || launchParams.files.length === 0) return;
            
            // Tomamos el primer archivo que el usuario haya intentado abrir
            const fileHandle = launchParams.files[0];
            const file = await fileHandle.getFile();
            
            const reader = new FileReader();
            reader.onload = function (e) {
                try {
                    const loadedData = JSON.parse(e.target.result);
                    appData = { ...appData, ...loadedData };
                    if (!appData.activeModuleId && appData.modules.length > 0) appData.activeModuleId = appData.modules[0].id;
                    if (!appData.database) appData.database = { materiales: [], mano_obra: [], equipos: [] };
                    
                    saveData();
                    renderB1(); 
                    renderBankList(); 
                    renderDBTables(); 
                    loadSettingsToUI();
                    switchTab('b1');
                    
                    showToast("Proyecto cargado automáticamente");
                } catch (err) {
                    console.error(err); 
                    showToast("Error al leer el archivo. Verifique que sea válido.");
                }
            };
            reader.readAsText(file);
        });
    }

};

function initializeDefaultSettings() {
    if (!appData.settings.numberFormat) appData.settings.numberFormat = detectUserLocale();
    saveData();
}