const fs = require('fs');

let content = fs.readFileSync('app.js', 'utf8');

const replacement1 = `let componentsData = {
    cpu: [],
    motherboard: [],
    ram: [],
    storageType: [
        { id: 'sata', name: 'SSD SATA' },
        { id: 'nvme', name: 'SSD M.2 NVMe' },
        { id: 'hdd', name: 'Disco Duro HDD' }
    ],
    storage: {
        sata: [],
        nvme: [],
        hdd: []
    },
    case: [],
    psu: [],
    gpu: [],
    cooling: []
};

async function loadExcelData() {
    try {
        const response = await fetch('Lista para Armador pc.xlsx');
        if (!response.ok) throw new Error("No se pudo cargar el archivo Excel");
        
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        let currentCategory = '';
        
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            
            if (!row || row.length === 0 || (row.length === 1 && !row[0])) continue;
            
            if (row.length === 1 || (row.length >= 2 && !row[1])) {
                currentCategory = String(row[0]).trim().toLowerCase();
                continue;
            }
            
            const name = String(row[0]).trim();
            const price = parseFloat(row[1]) || 0;
            const id = 'item_' + i;
            const product = { id, name, price };
            
            if (currentCategory.includes('procesador')) {
                let socket = 'AM4';
                if (currentCategory.includes('am5') || name.includes('AM5') || name.includes('7600') || name.includes('7700') || name.includes('7900') || name.includes('8500') || name.includes('8700')) socket = 'AM5';
                else if (currentCategory.includes('am4') || name.includes('AM4')) socket = 'AM4';
                else if (currentCategory.includes('1700') || name.includes('1700') || name.includes('12100') || name.includes('12400') || name.includes('13400') || name.includes('14700') || name.includes('12700')) socket = 'LGA1700';
                else if (currentCategory.includes('1200') || name.includes('1200') || name.includes('10500')) socket = 'LGA1200';
                else if (currentCategory.includes('1851') || name.includes('1851')) socket = 'LGA1851';
                
                product.socket = socket;
                componentsData.cpu.push(product);
            }
            else if (currentCategory.includes('mother')) {
                let socket = 'AM4';
                if (currentCategory.includes('am5') || name.toUpperCase().includes('B650') || name.toUpperCase().includes('X670') || name.toUpperCase().includes('A620') || name.toUpperCase().includes('AM5')) socket = 'AM5';
                else if (currentCategory.includes('am4') || name.toUpperCase().includes('B550') || name.toUpperCase().includes('A520') || name.toUpperCase().includes('X570') || name.toUpperCase().includes('AM4')) socket = 'AM4';
                else if (currentCategory.includes('1700') || name.toUpperCase().includes('H610') || name.toUpperCase().includes('B760') || name.toUpperCase().includes('Z790') || name.toUpperCase().includes('1700')) socket = 'LGA1700';
                else if (currentCategory.includes('1200') || name.toUpperCase().includes('H510') || name.toUpperCase().includes('B560') || name.toUpperCase().includes('1200')) socket = 'LGA1200';
                else if (currentCategory.includes('1851')) socket = 'LGA1851';
                
                product.socket = socket;
                componentsData.motherboard.push(product);
            }
            else if (currentCategory.includes('memoria') && !currentCategory.includes('video')) {
                componentsData.ram.push(product);
            }
            else if (currentCategory.includes('disco') || currentCategory.includes('ssd') || currentCategory.includes('almacenamiento')) {
                if (currentCategory.includes('nvme') || name.toLowerCase().includes('nvme') || name.toLowerCase().includes('m.2')) {
                    componentsData.storage.nvme.push(product);
                } else if (currentCategory.includes('hdd') || name.toLowerCase().includes('hdd') || currentCategory.includes('duro')) {
                    componentsData.storage.hdd.push(product);
                } else {
                    componentsData.storage.sata.push(product);
                }
            }
            else if (currentCategory.includes('gabinete')) {
                componentsData.case.push(product);
            }
            else if (currentCategory.includes('fuente') || currentCategory.includes('psu')) {
                componentsData.psu.push(product);
            }
            else if (currentCategory.includes('video') || currentCategory.includes('vga') || currentCategory.includes('gpu') || currentCategory.includes('placa')) {
                componentsData.gpu.push(product);
            }
            else if (currentCategory.includes('cooler') || currentCategory.includes('refrigeraci') || currentCategory.includes('water') || currentCategory.includes('coler')) {
                componentsData.cooling.push(product);
            }
        }
        
        if (componentsData.cooling.length === 0) {
            componentsData.cooling.push({ id: 'co_stock', name: 'Cooler Stock (Incluido con CPU)', price: 0 });
        }
        if (componentsData.gpu.length === 0) {
            componentsData.gpu.push({ id: 'gpu_int', name: 'Sin placa (Integrados)', price: 0 });
        }
        if (componentsData.psu.length === 0) {
            componentsData.psu.push({ id: 'psu_gen', name: 'Fuente', price: 0 });
        }
        
    } catch (error) {
        console.error("Error loading Excel:", error);
        alert("Hubo un error cargando la base de datos de precios desde Excel.");
    }
}`;

const replacement2 = `async function init() {
    await loadExcelData();
    
    initTheme();
    populateSelect(caseSelect, componentsData.case);
    populateSelect(psuSelect, componentsData.psu);
    populateSelect(gpuSelect, componentsData.gpu);
    populateSelect(coolingSelect, componentsData.cooling);
    addStorageRow(); // Initial storage row
    
    // Listeners
    platformBtns.forEach(btn => btn.addEventListener('click', handlePlatformClick));
    mbSelect.addEventListener('change', handleMotherboardChange);
    addStorageBtn.addEventListener('click', addStorageRow);
    copyBtn.addEventListener('click', copyConfig);
    if (whatsappBtn) {
        whatsappBtn.addEventListener('click', handleWhatsAppClick);
    }
    
    buyMultipleCheckbox.addEventListener('change', () => {
        if (buyMultipleCheckbox.checked) {
            whatsappInfo.classList.add('show');
        } else {
            whatsappInfo.classList.remove('show');
        }
        updatePrice();
    });
    
    // Price update listeners
    document.querySelector('.configurator-main').addEventListener('change', updatePrice);
}`;

// Find indices for componentsData block
const startIdx = content.indexOf('// DUMMY DATA');
const endIdx = content.indexOf('};\r\n', startIdx) + 4; // Or search for '};\n' if CRLF vs LF

let newContent = content;
if (startIdx !== -1) {
    let actualEndIdx = content.indexOf('};\r\n', startIdx);
    if (actualEndIdx === -1) actualEndIdx = content.indexOf('};\n', startIdx);
    if (actualEndIdx !== -1) {
        actualEndIdx += (content.substring(actualEndIdx, actualEndIdx + 4) === '};\r\n' ? 4 : 3);
        newContent = newContent.substring(0, startIdx) + replacement1 + newContent.substring(actualEndIdx);
    }
}

// Find indices for init()
const initStartIdx = newContent.indexOf('// Initialize\r\nfunction init() {');
let actualInitStartIdx = initStartIdx;
if (initStartIdx === -1) actualInitStartIdx = newContent.indexOf('// Initialize\nfunction init() {');

if (actualInitStartIdx !== -1) {
    const initEndStr = '    document.querySelector(\'.configurator-main\').addEventListener(\'change\', updatePrice);\r\n}\r\n';
    let initEndIdx = newContent.indexOf(initEndStr, actualInitStartIdx);
    let offset = initEndStr.length;
    
    if (initEndIdx === -1) {
        const initEndStrLF = '    document.querySelector(\'.configurator-main\').addEventListener(\'change\', updatePrice);\n}\n';
        initEndIdx = newContent.indexOf(initEndStrLF, actualInitStartIdx);
        offset = initEndStrLF.length;
    }
    
    if (initEndIdx !== -1) {
        newContent = newContent.substring(0, actualInitStartIdx) + '// Initialize\n' + replacement2 + '\n' + newContent.substring(initEndIdx + offset);
    }
}

fs.writeFileSync('app.js', newContent, 'utf8');
