let products = [];
let filteredProducts = [];
let cart = [];
let currentTotal = 0;

// DOM Elements
const fileInput = document.getElementById('file-input');
const uploadOverlay = document.getElementById('upload-overlay');
const dashboard = document.getElementById('dashboard');
const searchInput = document.getElementById('search-input');
const categoryFilter = document.getElementById('category-filter');
const sortFilter = document.getElementById('sort-filter');
const productsBody = document.getElementById('products-body');
const cartItemsContainer = document.getElementById('cart-items');
const subtotalVal = document.getElementById('subtotal-val');
const totalVal = document.getElementById('total-val');
const clientNameInput = document.getElementById('client-name-input');
const markupInput = document.getElementById('markup-input');
const customTotalInput = document.getElementById('custom-total-input');
const discountInfo = document.getElementById('discount-info');
const discountVal = document.getElementById('discount-val');
const clearCartBtn = document.getElementById('clear-cart-btn');
const exportBtn = document.getElementById('export-btn');
const pdfBtn = document.getElementById('pdf-btn');
const copyBtn = document.getElementById('copy-btn');
const fetchOnlineBtn = document.getElementById('fetch-online-btn');
const versionBadge = document.getElementById('version-badge');
const changeListBtn = document.getElementById('change-list-btn');
const mobileWarningBanner = document.getElementById('mobile-warning-banner');
const closeWarningBtn = document.getElementById('close-warning-btn');
const clientViewToggle = document.getElementById('client-view-toggle');
const lightModeToggle = document.getElementById('light-mode-toggle');
const googleThemeToggle = document.getElementById('google-theme-toggle');
const closeUploadBtn = document.getElementById('close-upload-btn');
const toastContainer = document.getElementById('toast-container');
const saveBudgetBtn = document.getElementById('save-budget-btn');
const loadBudgetBtn = document.getElementById('load-budget-btn');
const savedBudgetsOverlay = document.getElementById('saved-budgets-overlay');
const closeSavedBudgetsBtn = document.getElementById('close-saved-budgets-btn');
const savedBudgetsList = document.getElementById('saved-budgets-list');

let currentCurrency = 'ARS';
let dollarRate = 1;
const dollarRateInput = document.getElementById('dollar-rate-input');
const currencyArsBtn = document.getElementById('currency-ars');
const currencyUsdBtn = document.getElementById('currency-usd');
const priceCurrencyLabel = document.getElementById('price-currency-label');

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    loadState();
    // attemptAutoLoad(); // Deshabilitado para que el usuario siempre vea la pantalla de inicio

    fileInput.addEventListener('change', handleFileUpload);
    searchInput.addEventListener('input', applyFilters);
    categoryFilter.addEventListener('change', applyFilters);
    sortFilter.addEventListener('change', applyFilters);
    markupInput.addEventListener('input', () => { saveState(); renderCart(); renderProducts(); });
    customTotalInput.addEventListener('input', () => { saveState(); calculateDiscount(); });
    clientNameInput.addEventListener('input', saveState);
    if (clientViewToggle) clientViewToggle.addEventListener('change', () => { saveState(); renderProducts(); });
    if (lightModeToggle) lightModeToggle.addEventListener('change', () => { 
        document.body.classList.toggle('light-mode', lightModeToggle.checked);
        saveState(); 
    });
    if (googleThemeToggle) googleThemeToggle.addEventListener('change', () => { 
        document.body.classList.toggle('google-theme', googleThemeToggle.checked);
        saveState(); 
    });
    clearCartBtn.addEventListener('click', clearCart);
    exportBtn.addEventListener('click', exportCart);
    if (pdfBtn) pdfBtn.addEventListener('click', exportPDF);
    copyBtn.addEventListener('click', copyCart);
    if (fetchOnlineBtn) fetchOnlineBtn.addEventListener('click', loadOnlineExcel);
    
    if (saveBudgetBtn) saveBudgetBtn.addEventListener('click', saveCurrentBudget);
    if (loadBudgetBtn) loadBudgetBtn.addEventListener('click', showSavedBudgets);
    if (closeSavedBudgetsBtn) closeSavedBudgetsBtn.addEventListener('click', () => {
        savedBudgetsOverlay.style.display = 'none';
    });

    if (dollarRateInput) {
        dollarRateInput.addEventListener('input', () => {
            const val = parseFloat(dollarRateInput.value);
            if (!isNaN(val) && val > 0) {
                dollarRate = val;
                renderProducts();
                renderCart();
            }
        });
    }
    
    if (currencyArsBtn && currencyUsdBtn) {
        currencyArsBtn.addEventListener('click', () => setCurrency('ARS'));
        currencyUsdBtn.addEventListener('click', () => setCurrency('USD'));
    }

    fetchDollarRate();

    if (closeUploadBtn) {
        closeUploadBtn.addEventListener('click', () => {
            uploadOverlay.style.display = 'none';
        });
    }
    if (changeListBtn) changeListBtn.addEventListener('click', () => {
        uploadOverlay.style.display = 'flex';
        if (closeUploadBtn) {
            closeUploadBtn.style.display = products.length > 0 ? 'block' : 'none';
        }
    });
    if (closeWarningBtn && mobileWarningBanner) {
        closeWarningBtn.addEventListener('click', () => {
            mobileWarningBanner.style.setProperty('display', 'none', 'important');
        });
    }
});

// LocalStorage Persistence
function saveState() {
    const state = {
        cart: cart,
        markup: markupInput.value,
        customTotal: customTotalInput.value,
        clientName: clientNameInput.value,
        clientView: clientViewToggle ? clientViewToggle.checked : false,
        lightMode: lightModeToggle ? lightModeToggle.checked : false,
        googleTheme: googleThemeToggle ? googleThemeToggle.checked : false
    };
    localStorage.setItem('decsatech_state', JSON.stringify(state));
}

function loadState() {
    const saved = localStorage.getItem('decsatech_state');
    if (saved) {
        try {
            const state = JSON.parse(saved);
            if (state.cart) cart = state.cart;
            if (state.markup !== undefined) markupInput.value = state.markup;
            if (state.customTotal !== undefined) customTotalInput.value = state.customTotal;
            if (state.clientName !== undefined) clientNameInput.value = state.clientName;
            if (state.clientView !== undefined && clientViewToggle) clientViewToggle.checked = state.clientView;
            if (state.lightMode !== undefined && lightModeToggle) {
                lightModeToggle.checked = state.lightMode;
                document.body.classList.toggle('light-mode', state.lightMode);
            }
            if (state.googleTheme !== undefined && googleThemeToggle) {
                googleThemeToggle.checked = state.googleTheme;
                document.body.classList.toggle('google-theme', state.googleTheme);
            }
            renderCart();
        } catch (e) {
            console.error("Error loading state", e);
        }
    }
}

// Try to auto-load the file if served via HTTP
async function attemptAutoLoad() {
    try {
        const res = await fetch('Lista_de_Precios_Decsatech.xlsx');
        if (!res.ok) throw new Error("File not available via fetch");
        const arrayBuffer = await res.arrayBuffer();
        processExcel(arrayBuffer, "Carga Local", "badge-local");
    } catch (e) {
        console.log("Could not auto-load file. Waiting for manual upload.");
        uploadOverlay.style.display = 'flex';
    }
}

// Load online file from GitHub
async function loadOnlineExcel() {
    const rawUrl = 'https://raw.githubusercontent.com/juone0450/DECSATECH/main/Lista_de_Precios_Decsatech.xlsx';
    const originalText = fetchOnlineBtn.textContent;
    fetchOnlineBtn.textContent = 'Cargando...';
    fetchOnlineBtn.disabled = true;
    
    try {
        const res = await fetch(rawUrl);
        if (!res.ok) throw new Error("No se pudo descargar el archivo online");
        const arrayBuffer = await res.arrayBuffer();
        processExcel(arrayBuffer, "Online (GitHub)", "badge-online");
    } catch (e) {
        console.error(e);
        showToast("Hubo un error al cargar la lista desde GitHub. Verifica tu conexión.", "error");
        fetchOnlineBtn.textContent = originalText;
        fetchOnlineBtn.disabled = false;
    }
}

// Handle manual file upload
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = e.target.result;
        const time = new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
        processExcel(data, `Local - ${time}`, "badge-local");
    };
    reader.readAsArrayBuffer(file);
}

// Process the Excel buffer
function processExcel(buffer, sourceName = "Local", sourceClass = "badge-local") {
    try {
        const workbook = XLSX.read(buffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        products = [];
        const categories = new Set();

        // Data starts roughly at row 7.
        // Columns: 0: SKU, 1: Category, 2: Subcategory, 3: Description, 4: Price
        for (let i = 6; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length < 5) continue;
            
            const sku = row[0];
            const category = row[1];
            const description = row[3];
            const price = parseFloat(row[4]);

            if (sku && description && !isNaN(price)) {
                let catStr = category ? String(category).trim() : '';
                if (!catStr) {
                    catStr = 'sin nombrar';
                }
                categories.add(catStr);

                products.push({
                    id: String(sku).trim() + '-' + i, // unique ID just in case
                    sku: String(sku).trim(),
                    category: catStr,
                    description: String(description).trim(),
                    price: price
                });
            }
        }

        populateCategoryFilter(Array.from(categories).sort());
        filteredProducts = [...products];
        
        // Hide overlay, show dashboard
        uploadOverlay.style.display = 'none';
        dashboard.style.display = 'flex';
        
        if (versionBadge) {
            versionBadge.textContent = sourceName;
            versionBadge.className = `badge ${sourceClass}`;
            versionBadge.style.display = 'inline-block';
        }
        
        renderProducts();
    } catch (error) {
        showToast("Error al leer el archivo Excel. Asegúrate de que sea el formato correcto.", "error");
        console.error(error);
    }
}

// Populate Category Dropdown
function populateCategoryFilter(categories) {
    categoryFilter.innerHTML = '<option value="">Todas las Categorías</option>';
    categories.forEach(cat => {
        const option = document.createElement('option');
        option.value = cat;
        option.textContent = cat;
        categoryFilter.appendChild(option);
    });
}

// Apply Search and Category Filters
function applyFilters() {
    const searchTerm = searchInput.value.toLowerCase().trim();
    const searchTerms = searchTerm ? searchTerm.split(/[\s+]+/) : [];
    const selectedCategory = categoryFilter.value;
    const sortBy = sortFilter.value;

    filteredProducts = products.filter(p => {
        let matchesSearch = true;
        if (searchTerms.length > 0) {
            const desc = p.description.toLowerCase();
            const sku = p.sku.toLowerCase();
            matchesSearch = searchTerms.every(term => desc.includes(term) || sku.includes(term));
        }
        const matchesCategory = selectedCategory === "" || p.category === selectedCategory;
        return matchesSearch && matchesCategory;
    });

    if (sortBy === 'price-asc') {
        filteredProducts.sort((a, b) => a.price - b.price);
    } else if (sortBy === 'price-desc') {
        filteredProducts.sort((a, b) => b.price - a.price);
    }

    renderProducts();
}

// Format Currency
function formatMoney(amount) {
    if (currentCurrency === 'USD') {
        const usdAmount = amount / dollarRate;
        const formatted = new Intl.NumberFormat('en-US', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(usdAmount);
        return 'USD ' + formatted;
    }
    return new Intl.NumberFormat('es-AR', {
        style: 'currency',
        currency: 'ARS',
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(amount);
}

// Render Products Table
function renderProducts() {
    productsBody.innerHTML = '';
    
    // Limit to 100 to prevent lagging the UI on massive lists
    const displayList = filteredProducts.slice(0, 100);

    displayList.forEach(p => {
        const tr = document.createElement('tr');
        
        const descTd = document.createElement('td');
        descTd.textContent = p.description;
        
        const catTd = document.createElement('td');
        catTd.textContent = p.category;
        
        const priceTd = document.createElement('td');
        const isClientView = clientViewToggle && clientViewToggle.checked;
        const markupVal = parseFloat(markupInput.value) || 0;
        const markupMultiplier = 1 + (markupVal / 100);
        const displayPrice = isClientView ? p.price * markupMultiplier : p.price;
        priceTd.textContent = formatMoney(displayPrice);
        
        const actionTd = document.createElement('td');
        const addBtn = document.createElement('button');
        addBtn.className = 'add-btn';
        addBtn.textContent = 'Agregar';
        addBtn.onclick = () => addToCart(p);
        actionTd.appendChild(addBtn);

        tr.appendChild(descTd);
        tr.appendChild(catTd);
        tr.appendChild(priceTd);
        tr.appendChild(actionTd);
        
        productsBody.appendChild(tr);
    });

    if (filteredProducts.length === 0) {
        productsBody.innerHTML = '<tr><td colspan="4" style="text-align:center; color: var(--text-muted)">No se encontraron productos.</td></tr>';
    } else if (filteredProducts.length > 100) {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td colspan="4" style="text-align:center; color: var(--text-muted); font-size: 0.9em;">Mostrando 100 de ${filteredProducts.length} resultados. Usa el buscador para encontrar componentes específicos.</td>`;
        productsBody.appendChild(tr);
    }
}

// Add Item to Cart
function addToCart(product) {
    const existing = cart.find(item => item.id === product.id);
    if (existing) {
        existing.quantity = (existing.quantity || 1) + 1;
        saveState();
        renderCart();
    } else {
        const cartItem = { ...product, cartId: Date.now() + Math.random(), quantity: 1 };
        cart.push(cartItem);
        saveState();
        renderCart();
        fetchProductImage(cartItem);
    }
}

// Adjust Quantity
function updateQuantity(cartId, delta) {
    const item = cart.find(i => i.cartId === cartId);
    if (item) {
        item.quantity = (item.quantity || 1) + delta;
        if (item.quantity <= 0) {
            removeFromCart(cartId);
        } else {
            saveState();
            renderCart();
        }
    }
}

// Remove Item from Cart
function removeFromCart(cartId) {
    cart = cart.filter(item => item.cartId !== cartId);
    saveState();
    renderCart();
}

// Clear Entire Cart
function clearCart() {
    cart = [];
    clientNameInput.value = '';
    customTotalInput.value = '';
    markupInput.value = '0';
    saveState();
    renderCart();
    renderProducts(); // Refresh client view price if markup is reset
}

// Export Cart to PDF / Print
function exportPDF() {
    if (cart.length === 0) {
        showToast("El presupuesto está vacío.", "warning");
        return;
    }

    const clientName = clientNameInput.value.trim();
    const dateStr = new Date().toLocaleDateString();

    const formatARS = (amount) => new Intl.NumberFormat('es-AR', { style: 'currency', currency: 'ARS', minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(amount);
    const formatUSD = (amount) => 'USD ' + new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(amount / dollarRate);

    let html = `
    <html>
    <head>
        <title>Presupuesto_Decsatech_${clientName || 'Cliente'}_${dateStr}</title>
        <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&display=swap" rel="stylesheet">
        <style>
            body {
                font-family: 'Outfit', sans-serif;
                color: #202124;
                margin: 0;
                padding: 40px;
                background: #fff;
            }
            .header {
                border-bottom: 2px solid #0b57d0;
                padding-bottom: 20px;
                margin-bottom: 30px;
                display: flex;
                justify-content: space-between;
                align-items: flex-end;
            }
            h1 { margin: 0; color: #1f1f1f; font-size: 28px; font-weight: 800; }
            h1 span { color: #0b57d0; }
            table {
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 30px;
                font-size: 13px;
            }
            th { padding: 12px; text-align: left; border-bottom: 2px solid #dadce0; color: #444746; font-weight: 600; }
            td { padding: 12px; border-bottom: 1px solid #f1f3f4; color: #1f1f1f; }
            .right { text-align: right; }
            .total-box {
                width: 320px;
                background: #f8fafd;
                padding: 20px;
                border-radius: 12px;
                float: right;
                border: 1px solid #e1e3e1;
            }
            .flex-row { display: flex; justify-content: space-between; margin-bottom: 10px; }
            .usd-text { font-size: 10px; color: #9aa0a6; font-weight: 400; }
            
            @media print {
                @page { margin: 15mm; size: A4; }
                body { padding: 0; }
                .total-box { border: 1px solid #dadce0; }
            }
        </style>
    </head>
    <body>
        <div class="header">
            <div>
                <h1>DECSATECH <span>PRO</span></h1>
                <p style="margin: 5px 0 0 0; color: #5f6368; font-size: 14px;">Presupuesto de Componentes</p>
            </div>
            <div class="right">
                <p style="margin: 0; color: #444746; font-size: 14px;"><strong>Fecha:</strong> ${dateStr}</p>
                ${clientName ? `<p style="margin: 5px 0 0 0; color: #444746; font-size: 14px;"><strong>Cliente:</strong> ${clientName}</p>` : ''}
                <p style="margin: 5px 0 0 0; color: #5f6368; font-size: 12px;"><strong>Cotización Dólar:</strong> $${dollarRate}</p>
            </div>
        </div>
        
        <table>
            <thead>
                <tr style="background: #f8fafd;">
                    <th style="width: 5%;">Cant.</th>
                    <th style="width: 43%;">Descripción del Producto</th>
                    <th class="right" style="width: 13%;">P.U. (ARS)</th>
                    <th class="right" style="width: 13%;">P.U. (USD)</th>
                    <th class="right" style="width: 13%;">Subtotal (ARS)</th>
                    <th class="right" style="width: 13%;">Subtotal (USD)</th>
                </tr>
            </thead>
            <tbody>
    `;

    let total = 0;
    const markupVal = parseFloat(markupInput.value) || 0;
    const markupMultiplier = 1 + (markupVal / 100);

    cart.forEach(item => {
        const qty = item.quantity || 1;
        const itemPrice = item.price * markupMultiplier;
        const rowTotal = itemPrice * qty;
        total += rowTotal;

        html += `
            <tr>
                <td>${qty}</td>
                <td>
                    <div style="font-weight: 600;">${item.description}</div>
                    <div class="usd-text">SKU: ${item.sku}</div>
                </td>
                <td class="right" style="font-weight: 600;">${formatARS(itemPrice)}</td>
                <td class="right" style="color: #0b57d0; font-weight: 600;">${formatUSD(itemPrice)}</td>
                <td class="right" style="font-weight: 600;">${formatARS(rowTotal)}</td>
                <td class="right" style="color: #0b57d0; font-weight: 600;">${formatUSD(rowTotal)}</td>
            </tr>
        `;
    });

    html += `
            </tbody>
        </table>
        
        <div style="display: flow-root;">
            <div class="total-box">
                <div class="flex-row" style="font-size: 14px;">
                    <span style="color: #444746;">Subtotal:</span>
                    <span style="font-weight: 600;" class="right">${formatARS(total)} <br><span class="usd-text">${formatUSD(total)}</span></span>
                </div>
    `;

    const customValInput = parseFloat(customTotalInput.value);
    const customValARS = currentCurrency === 'USD' ? customValInput * dollarRate : customValInput;
    let finalTotal = total;
    
    if (!isNaN(customValInput) && customValInput > 0 && total > 0) {
        finalTotal = customValARS;
        const discountAmount = total - customValARS;
        
        if (discountAmount > 0) {
            html += `
                <div class="flex-row" style="color: #146c2e; font-size: 14px;">
                    <span>Descuento:</span>
                    <span>-${formatARS(discountAmount)}</span>
                </div>
            `;
        } else if (discountAmount < 0) {
            html += `
                <div class="flex-row" style="color: #b3261e; font-size: 14px;">
                    <span>Recargo:</span>
                    <span>+${formatARS(Math.abs(discountAmount))}</span>
                </div>
            `;
        }
    }

    html += `
                <div class="flex-row" style="margin-top: 15px; padding-top: 15px; border-top: 2px solid #dadce0; align-items: flex-end;">
                    <span style="font-size: 18px; font-weight: 800; color: #1f1f1f;">TOTAL:</span>
                    <div class="right">
                        <span style="font-size: 18px; font-weight: 800; color: #0b57d0;">${formatARS(finalTotal)}</span>
                        <div style="font-size: 11px; font-weight: 400; color: #9aa0a6; margin-top: 2px;">${formatUSD(finalTotal)}</div>
                    </div>
                </div>
            </div>
        </div>
        
        <div style="margin-top: 60px; text-align: center; color: #5f6368; font-size: 12px; border-top: 1px solid #e1e3e1; padding-top: 20px;">
            Los precios pueden estar sujetos a modificaciones. Presupuesto generado por DECSATECH PRO.
        </div>
    </body>
    </html>
    `;

    const printWindow = window.open('', '_blank');
    if (!printWindow) {
        showToast("Por favor habilita las ventanas emergentes (popups) para imprimir el PDF.", "error");
        return;
    }
    printWindow.document.open();
    printWindow.document.write(html);
    printWindow.document.close();

    // Give time for fonts to load before triggering print dialog
    setTimeout(() => {
        printWindow.focus();
        printWindow.print();
        // Optional: close the window automatically after print dialog is closed
        // printWindow.close(); 
    }, 500);
}

function exportCart() {
    if (cart.length === 0) {
        showToast("El presupuesto está vacío.", "warning");
        return;
    }

    const clientName = clientNameInput.value.trim();
    let text = "PRESUPUESTO - DECSATECH PRO\n";
    if (clientName) {
        text += `CLIENTE: ${clientName}\n`;
    }
    text += "========================================\n\n";
    
    let total = 0;
    const markupVal = parseFloat(markupInput.value) || 0;
    const markupMultiplier = 1 + (markupVal / 100);

    cart.forEach(item => {
        const qty = item.quantity || 1;
        const itemPrice = item.price * markupMultiplier;
        const rowTotal = itemPrice * qty;
        text += `- ${qty}x ${item.description}\n`;
        text += `  SKU: ${item.sku} | Precio U.: ${formatMoney(itemPrice)} | Subtotal: ${formatMoney(rowTotal)}\n\n`;
        total += rowTotal;
    });

    text += "========================================\n";
    text += `TOTAL: ${formatMoney(total)}\n`;
    
    const customValInput = parseFloat(customTotalInput.value);
    const customValARS = currentCurrency === 'USD' ? customValInput * dollarRate : customValInput;
    if (!isNaN(customValInput) && customValInput > 0 && total > 0) {
        const discountAmount = total - customValARS;
        const discountPercent = (discountAmount / total) * 100;
        text += `TOTAL ACORDADO: ${formatMoney(customValARS)}\n`;
        if (discountAmount > 0) {
            text += `DESCUENTO APLICADO: ${formatMoney(discountAmount)} (${discountPercent.toFixed(1)}%)\n`;
        } else if (discountAmount < 0) {
            text += `RECARGO APLICADO: ${formatMoney(Math.abs(discountAmount))} (${Math.abs(discountPercent).toFixed(1)}%)\n`;
        }
    }

    const blob = new Blob([text], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const safeClientName = clientName ? `_${clientName.replace(/[^a-z0-9]/gi, '_')}` : '';
    a.download = `Presupuesto_Decsatech${safeClientName}_${new Date().toISOString().split('T')[0]}.txt`;
    a.click();
    URL.revokeObjectURL(url);
}

// Copy Cart to Clipboard
function copyCart() {
    if (cart.length === 0) {
        showToast("El presupuesto está vacío.", "warning");
        return;
    }

    const clientName = clientNameInput.value.trim();
    let text = "PRESUPUESTO - DECSATECH PRO\n";
    if (clientName) {
        text += `CLIENTE: ${clientName}\n`;
    }
    text += "========================================\n\n";
    
    let total = 0;
    const markupVal = parseFloat(markupInput.value) || 0;
    const markupMultiplier = 1 + (markupVal / 100);

    cart.forEach(item => {
        const qty = item.quantity || 1;
        const itemPrice = item.price * markupMultiplier;
        const rowTotal = itemPrice * qty;
        text += `- ${qty}x ${item.description}\n`;
        text += `  SKU: ${item.sku} | Precio U.: ${formatMoney(itemPrice)} | Subtotal: ${formatMoney(rowTotal)}\n\n`;
        total += rowTotal;
    });

    text += "========================================\n";
    text += `TOTAL: ${formatMoney(total)}\n`;
    
    const customValInput = parseFloat(customTotalInput.value);
    const customValARS = currentCurrency === 'USD' ? customValInput * dollarRate : customValInput;
    if (!isNaN(customValInput) && customValInput > 0 && total > 0) {
        const discountAmount = total - customValARS;
        const discountPercent = (discountAmount / total) * 100;
        text += `TOTAL ACORDADO: ${formatMoney(customValARS)}\n`;
        if (discountAmount > 0) {
            text += `DESCUENTO APLICADO: ${formatMoney(discountAmount)} (${discountPercent.toFixed(1)}%)\n`;
        } else if (discountAmount < 0) {
            text += `RECARGO APLICADO: ${formatMoney(Math.abs(discountAmount))} (${Math.abs(discountPercent).toFixed(1)}%)\n`;
        }
    }

    navigator.clipboard.writeText(text).then(() => {
        showToast("¡Presupuesto copiado al portapapeles!", "success");
    }).catch(err => {
        console.error("Error al copiar: ", err);
        showToast("Error al copiar el presupuesto. Es posible que el navegador no lo permita sin HTTPS.", "error");
    });
}

// Render Cart Panel
function renderCart() {
    cartItemsContainer.innerHTML = '';
    
    if (cart.length === 0) {
        cartItemsContainer.innerHTML = '<div class="empty-cart">No hay componentes en el presupuesto. Selecciona componentes del catálogo.</div>';
        updateTotals(0);
        return;
    }

    let total = 0;
    const markupVal = parseFloat(markupInput.value) || 0;
    const markupMultiplier = 1 + (markupVal / 100);

    cart.forEach(item => {
        const qty = item.quantity || 1;
        const itemPrice = item.price * markupMultiplier;
        const rowTotal = itemPrice * qty;
        total += rowTotal;

        const div = document.createElement('div');
        div.className = 'cart-item';
        
        const leftDiv = document.createElement('div');
        leftDiv.className = 'cart-item-left';

        const img = document.createElement('img');
        img.className = 'cart-item-image' + (item.imageUrl && item.imageUrl !== 'placeholder' ? '' : ' loading');
        img.src = item.imageUrl && item.imageUrl !== 'placeholder' ? item.imageUrl : 'data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="50" height="50" viewBox="0 0 24 24" fill="none" stroke="%23666" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect><circle cx="8.5" cy="8.5" r="1.5"></circle><polyline points="21 15 16 10 5 21"></polyline></svg>';

        const infoDiv = document.createElement('div');
        infoDiv.className = 'cart-item-info';
        
        const title = document.createElement('h4');
        title.textContent = item.category;
        
        const desc = document.createElement('p');
        desc.textContent = item.description;
        
        infoDiv.appendChild(title);
        infoDiv.appendChild(desc);

        leftDiv.appendChild(img);
        leftDiv.appendChild(infoDiv);

        const rightDiv = document.createElement('div');
        rightDiv.style.display = 'flex';
        rightDiv.style.alignItems = 'center';
        rightDiv.style.gap = '10px';

        const qtyControls = document.createElement('div');
        qtyControls.style.display = 'flex';
        qtyControls.style.alignItems = 'center';
        qtyControls.style.gap = '5px';
        qtyControls.style.background = 'rgba(0,0,0,0.2)';
        qtyControls.style.borderRadius = '4px';
        qtyControls.style.padding = '2px 5px';

        const minusBtn = document.createElement('button');
        minusBtn.textContent = '-';
        minusBtn.style.background = 'none';
        minusBtn.style.border = 'none';
        minusBtn.style.color = 'var(--text-main)';
        minusBtn.style.cursor = 'pointer';
        minusBtn.style.fontSize = '1.1rem';
        minusBtn.onclick = () => updateQuantity(item.cartId, -1);

        const qtySpan = document.createElement('span');
        qtySpan.textContent = qty;
        qtySpan.style.fontSize = '0.9rem';
        qtySpan.style.minWidth = '20px';
        qtySpan.style.textAlign = 'center';

        const plusBtn = document.createElement('button');
        plusBtn.textContent = '+';
        plusBtn.style.background = 'none';
        plusBtn.style.border = 'none';
        plusBtn.style.color = 'var(--text-main)';
        plusBtn.style.cursor = 'pointer';
        plusBtn.style.fontSize = '1.1rem';
        plusBtn.onclick = () => updateQuantity(item.cartId, 1);

        qtyControls.appendChild(minusBtn);
        qtyControls.appendChild(qtySpan);
        qtyControls.appendChild(plusBtn);

        const priceSpan = document.createElement('span');
        priceSpan.className = 'cart-item-price';
        priceSpan.textContent = formatMoney(rowTotal);

        const rmBtn = document.createElement('button');
        rmBtn.className = 'remove-btn';
        rmBtn.innerHTML = '×';
        rmBtn.title = 'Eliminar componente';
        rmBtn.onclick = () => removeFromCart(item.cartId);

        rightDiv.appendChild(qtyControls);
        rightDiv.appendChild(priceSpan);
        rightDiv.appendChild(rmBtn);

        div.appendChild(leftDiv);
        div.appendChild(rightDiv);
        
        cartItemsContainer.appendChild(div);
    });

    updateTotals(total);
}

function updateTotals(total) {
    currentTotal = total;
    const formatted = formatMoney(total);
    subtotalVal.textContent = formatted;
    totalVal.textContent = formatted;
    calculateDiscount();
}

function calculateDiscount() {
    const customValInput = parseFloat(customTotalInput.value);
    const customValARS = currentCurrency === 'USD' ? customValInput * dollarRate : customValInput;
    
    if (isNaN(customValInput) || customValInput <= 0 || currentTotal <= 0) {
        discountInfo.style.display = 'none';
        return;
    }

    const discountAmount = currentTotal - customValARS;
    const discountPercent = (discountAmount / currentTotal) * 100;
    
    discountInfo.style.display = 'flex';
    
    if (discountAmount > 0) {
        discountVal.textContent = `${formatMoney(discountAmount)} (${discountPercent.toFixed(1)}%)`;
        discountVal.className = 'success-text';
    } else if (discountAmount < 0) {
        discountVal.textContent = `Aumento: ${formatMoney(Math.abs(discountAmount))} (${Math.abs(discountPercent).toFixed(1)}%)`;
        discountVal.className = 'warning-text';
    } else {
        discountVal.textContent = `$0.00 (0%)`;
        discountVal.className = '';
    }
}

// TOAST NOTIFICATIONS
function showToast(message, type = 'success') {
    if (!toastContainer) return;
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;
    
    toast.onclick = () => {
        toast.style.animation = 'toastOut 0.3s ease-in forwards';
        setTimeout(() => toast.remove(), 300);
    };

    toastContainer.appendChild(toast);
    
    setTimeout(() => {
        if (toast.parentElement) {
            toast.style.animation = 'toastOut 0.3s ease-in forwards';
            setTimeout(() => toast.remove(), 300);
        }
    }, 4000);
}

// SAVED BUDGETS
function getSavedBudgets() {
    const saved = localStorage.getItem('decsatech_saved_budgets');
    return saved ? JSON.parse(saved) : [];
}

function saveCurrentBudget() {
    if (cart.length === 0) {
        showToast("No puedes guardar un presupuesto vacío.", "warning");
        return;
    }

    let name = clientNameInput.value.trim();
    if (!name) {
        name = prompt("Ingresa un nombre para guardar este presupuesto:");
        if (!name) return; // Cancelled
        clientNameInput.value = name;
        saveState(); // Update current state client name
    }

    const budgets = getSavedBudgets();
    
    // Check if we already have one with this name
    const existingIndex = budgets.findIndex(b => b.name.toLowerCase() === name.toLowerCase());
    
    const budgetData = {
        id: Date.now().toString(),
        name: name,
        date: new Date().toLocaleDateString() + ' ' + new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}),
        total: currentTotal,
        state: {
            cart: cart,
            markup: markupInput.value,
            customTotal: customTotalInput.value,
            clientName: clientNameInput.value,
            clientView: clientViewToggle ? clientViewToggle.checked : false
        }
    };

    if (existingIndex >= 0) {
        if (confirm(`Ya existe un presupuesto llamado "${name}". ¿Deseas sobreescribirlo?`)) {
            budgets[existingIndex] = budgetData;
        } else {
            return;
        }
    } else {
        budgets.push(budgetData);
    }

    localStorage.setItem('decsatech_saved_budgets', JSON.stringify(budgets));
    showToast(`Presupuesto "${name}" guardado exitosamente.`, "success");
}

function showSavedBudgets() {
    savedBudgetsList.innerHTML = '';
    const budgets = getSavedBudgets();

    if (budgets.length === 0) {
        savedBudgetsList.innerHTML = '<p style="color: var(--text-muted);">No tienes presupuestos guardados.</p>';
    } else {
        // Sort descending by id (timestamp)
        budgets.sort((a, b) => b.id - a.id).forEach(b => {
            const item = document.createElement('div');
            item.className = 'saved-budget-item';
            
            const info = document.createElement('div');
            info.className = 'saved-budget-info';
            info.innerHTML = `
                <h3>${b.name}</h3>
                <p>Fecha: ${b.date} | Total: ${formatMoney(b.total)}</p>
            `;
            
            const actions = document.createElement('div');
            actions.className = 'saved-budget-actions';
            
            const loadBtn = document.createElement('button');
            loadBtn.className = 'action-btn primary';
            loadBtn.style.padding = '0.5rem 1rem';
            loadBtn.style.marginBottom = '0';
            loadBtn.textContent = 'Cargar';
            loadBtn.onclick = () => loadSavedBudget(b);
            
            const delBtn = document.createElement('button');
            delBtn.className = 'action-btn outline';
            delBtn.style.padding = '0.5rem 1rem';
            delBtn.style.marginBottom = '0';
            delBtn.style.color = 'var(--danger)';
            delBtn.style.borderColor = 'var(--danger)';
            delBtn.textContent = 'Eliminar';
            delBtn.onclick = () => deleteSavedBudget(b.id);
            
            actions.appendChild(loadBtn);
            actions.appendChild(delBtn);
            
            item.appendChild(info);
            item.appendChild(actions);
            savedBudgetsList.appendChild(item);
        });
    }

    savedBudgetsOverlay.style.display = 'flex';
}

function loadSavedBudget(budget) {
    if (cart.length > 0 && !confirm("Esto reemplazará el presupuesto actual. ¿Estás seguro?")) {
        return;
    }
    
    let loadedCart = JSON.parse(JSON.stringify(budget.state.cart || []));
    let priceChanged = false;
    let oldTotal = 0;
    let newTotal = 0;
    
    loadedCart.forEach(item => {
        const qty = item.quantity || 1;
        oldTotal += item.price * qty;
        
        const currentProduct = products.find(p => p.sku === item.sku);
        if (currentProduct) {
            newTotal += currentProduct.price * qty;
            if (currentProduct.price !== item.price) {
                priceChanged = true;
            }
        } else {
            newTotal += item.price * qty; // Product missing, use old price
        }
    });

    if (priceChanged) {
        const msg = `Atención: Hay diferencias de precio en la lista actual.\n\nPrecio Original: ${formatMoney(oldTotal)}\nPrecio Actual: ${formatMoney(newTotal)}\n\n¿Deseas actualizar el presupuesto a los PRECIOS NUEVOS?\n(Pulsa "Cancelar" para mantener los precios viejos de cuando lo guardaste).`;
        
        if (confirm(msg)) {
            loadedCart.forEach(item => {
                const currentProduct = products.find(p => p.sku === item.sku);
                if (currentProduct) item.price = currentProduct.price;
            });
            showToast("Precios actualizados a la lista de hoy.", "info");
        } else {
            showToast("Se cargaron los precios viejos originales.", "info");
        }
    } else {
        showToast(`Presupuesto "${budget.name}" cargado.`, "success");
    }
    
    cart = loadedCart;
    markupInput.value = budget.state.markup || '0';
    customTotalInput.value = budget.state.customTotal || '';
    clientNameInput.value = budget.state.clientName || '';
    if (clientViewToggle) clientViewToggle.checked = budget.state.clientView || false;
    
    saveState();
    renderCart();
    renderProducts();
    
    savedBudgetsOverlay.style.display = 'none';
}

function deleteSavedBudget(id) {
    if (confirm("¿Seguro que deseas eliminar este presupuesto guardado?")) {
        let budgets = getSavedBudgets();
        budgets = budgets.filter(b => b.id !== id);
        localStorage.setItem('decsatech_saved_budgets', JSON.stringify(budgets));
        showSavedBudgets(); // Refresh list
        showToast("Presupuesto eliminado.", "info");
    }
}

// DOLLAR API & CURRENCY
async function fetchDollarRate() {
    try {
        const res = await fetch('https://dolarapi.com/v1/dolares/blue');
        if (!res.ok) throw new Error("Network response was not ok");
        const data = await res.json();
        if (data && data.venta) {
            dollarRate = parseFloat(data.venta);
            if (dollarRateInput) dollarRateInput.value = dollarRate;
            renderProducts();
            renderCart();
            calculateDiscount();
        }
    } catch (e) {
        console.error("No se pudo obtener el dólar", e);
        showToast("No se pudo actualizar el dólar automáticamente.", "warning");
    }
}

function setCurrency(currency) {
    currentCurrency = currency;
    if (currency === 'ARS') {
        currencyArsBtn.classList.add('active');
        currencyUsdBtn.classList.remove('active');
        priceCurrencyLabel.textContent = 'ARS';
    } else {
        currencyUsdBtn.classList.add('active');
        currencyArsBtn.classList.remove('active');
        priceCurrencyLabel.textContent = 'USD';
    }
    renderProducts();
    renderCart();
    calculateDiscount();
}

// MERCADOLIBRE IMAGES
const imageCache = {};

async function fetchProductImage(cartItem) {
    if (imageCache[cartItem.sku]) {
        cartItem.imageUrl = imageCache[cartItem.sku];
        renderCart();
        saveState();
        return;
    }
    
    try {
        const query = encodeURIComponent(cartItem.description.substring(0, 50));
        const res = await fetch(`https://api.mercadolibre.com/sites/MLA/search?q=${query}&limit=1`);
        const data = await res.json();
        if (data.results && data.results.length > 0 && data.results[0].thumbnail) {
            let imgUrl = data.results[0].thumbnail.replace("http://", "https://");
            // Improve image quality slightly if possible
            imgUrl = imgUrl.replace("-I.jpg", "-O.jpg");
            cartItem.imageUrl = imgUrl;
            imageCache[cartItem.sku] = imgUrl;
        } else {
            cartItem.imageUrl = "placeholder"; // no image found
        }
    } catch (e) {
        console.error("Error fetching image", e);
        cartItem.imageUrl = "placeholder";
    }
    renderCart();
    saveState();
}
