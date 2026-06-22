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
const copyBtn = document.getElementById('copy-btn');
const fetchOnlineBtn = document.getElementById('fetch-online-btn');

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    attemptAutoLoad();

    fileInput.addEventListener('change', handleFileUpload);
    searchInput.addEventListener('input', applyFilters);
    categoryFilter.addEventListener('change', applyFilters);
    sortFilter.addEventListener('change', applyFilters);
    markupInput.addEventListener('input', renderCart);
    customTotalInput.addEventListener('input', calculateDiscount);
    clearCartBtn.addEventListener('click', clearCart);
    exportBtn.addEventListener('click', exportCart);
    copyBtn.addEventListener('click', copyCart);
    if (fetchOnlineBtn) fetchOnlineBtn.addEventListener('click', loadOnlineExcel);
});

// Try to auto-load the file if served via HTTP
async function attemptAutoLoad() {
    try {
        const res = await fetch('Lista_de_Precios_Decsatech.xlsx');
        if (!res.ok) throw new Error("File not available via fetch");
        const arrayBuffer = await res.arrayBuffer();
        processExcel(arrayBuffer);
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
        processExcel(arrayBuffer);
    } catch (e) {
        console.error(e);
        alert("Hubo un error al cargar la lista desde GitHub. Verifica tu conexión.");
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
        processExcel(data);
    };
    reader.readAsArrayBuffer(file);
}

// Process the Excel buffer
function processExcel(buffer) {
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
        
        renderProducts();
    } catch (error) {
        alert("Error al leer el archivo Excel. Asegúrate de que sea el formato correcto.");
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
        priceTd.textContent = formatMoney(p.price);
        
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
    // Generate a unique instance ID so identical items can be removed individually
    const cartItem = { ...product, cartId: Date.now() + Math.random() };
    cart.push(cartItem);
    renderCart();
}

// Remove Item from Cart
function removeFromCart(cartId) {
    cart = cart.filter(item => item.cartId !== cartId);
    renderCart();
}

// Clear Entire Cart
function clearCart() {
    cart = [];
    clientNameInput.value = '';
    customTotalInput.value = '';
    markupInput.value = '0';
    renderCart();
}

// Export Cart to Text
function exportCart() {
    if (cart.length === 0) {
        alert("El presupuesto está vacío.");
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
        const itemPrice = item.price * markupMultiplier;
        text += `- ${item.description}\n`;
        text += `  SKU: ${item.sku} | Precio: ${formatMoney(itemPrice)}\n\n`;
        total += itemPrice;
    });

    text += "========================================\n";
    text += `TOTAL: ${formatMoney(total)}\n`;
    
    const customVal = parseFloat(customTotalInput.value);
    if (!isNaN(customVal) && customVal > 0 && total > 0) {
        const discountAmount = total - customVal;
        const discountPercent = (discountAmount / total) * 100;
        text += `TOTAL ACORDADO: ${formatMoney(customVal)}\n`;
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
        alert("El presupuesto está vacío.");
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
        const itemPrice = item.price * markupMultiplier;
        text += `- ${item.description}\n`;
        text += `  SKU: ${item.sku} | Precio: ${formatMoney(itemPrice)}\n\n`;
        total += itemPrice;
    });

    text += "========================================\n";
    text += `TOTAL: ${formatMoney(total)}\n`;
    
    const customVal = parseFloat(customTotalInput.value);
    if (!isNaN(customVal) && customVal > 0 && total > 0) {
        const discountAmount = total - customVal;
        const discountPercent = (discountAmount / total) * 100;
        text += `TOTAL ACORDADO: ${formatMoney(customVal)}\n`;
        if (discountAmount > 0) {
            text += `DESCUENTO APLICADO: ${formatMoney(discountAmount)} (${discountPercent.toFixed(1)}%)\n`;
        } else if (discountAmount < 0) {
            text += `RECARGO APLICADO: ${formatMoney(Math.abs(discountAmount))} (${Math.abs(discountPercent).toFixed(1)}%)\n`;
        }
    }

    navigator.clipboard.writeText(text).then(() => {
        alert("¡Presupuesto copiado al portapapeles!");
    }).catch(err => {
        console.error("Error al copiar: ", err);
        alert("Error al copiar el presupuesto. Es posible que el navegador no lo permita sin HTTPS.");
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
        const itemPrice = item.price * markupMultiplier;
        total += itemPrice;

        const div = document.createElement('div');
        div.className = 'cart-item';
        
        const infoDiv = document.createElement('div');
        infoDiv.className = 'cart-item-info';
        
        const title = document.createElement('h4');
        title.textContent = item.category;
        
        const desc = document.createElement('p');
        desc.textContent = item.description;
        
        infoDiv.appendChild(title);
        infoDiv.appendChild(desc);

        const rightDiv = document.createElement('div');
        rightDiv.style.display = 'flex';
        rightDiv.style.alignItems = 'center';

        const priceSpan = document.createElement('span');
        priceSpan.className = 'cart-item-price';
        priceSpan.textContent = formatMoney(itemPrice);

        const rmBtn = document.createElement('button');
        rmBtn.className = 'remove-btn';
        rmBtn.innerHTML = '×';
        rmBtn.title = 'Eliminar componente';
        rmBtn.onclick = () => removeFromCart(item.cartId);

        rightDiv.appendChild(priceSpan);
        rightDiv.appendChild(rmBtn);

        div.appendChild(infoDiv);
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
    const customVal = parseFloat(customTotalInput.value);
    
    if (isNaN(customVal) || customVal <= 0 || currentTotal <= 0) {
        discountInfo.style.display = 'none';
        return;
    }

    const discountAmount = currentTotal - customVal;
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
