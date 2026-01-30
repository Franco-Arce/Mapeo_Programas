/**
 * Mapeador de Programas Universitarios
 * Motor de mapeo con fuzzy matching
 */

// ========================================
// State Management
// ========================================
const state = {
    officialPrograms: [],
    fileData: null,
    headers: [],
    detectedColumns: {
        program: null,
        email: null,
        phone: null,
        contactId: null,
        database: null
    },
    uniqueDatabases: [],
    mappings: new Map(), // original -> { mapped, score, status }
    programCounts: new Map(), // original -> count
    currentEditProgram: null
};

// ========================================
// DOM Elements
// ========================================
const elements = {
    // Inputs
    officialPrograms: document.getElementById('officialPrograms'),
    programCount: document.getElementById('programCount'),
    loadPrograms: document.getElementById('loadPrograms'),
    dropZone: document.getElementById('dropZone'),
    fileInput: document.getElementById('fileInput'),
    fileInfo: document.getElementById('fileInfo'),
    fileName: document.getElementById('fileName'),
    fileStats: document.getElementById('fileStats'),
    removeFile: document.getElementById('removeFile'),

    // Sections
    columnSection: document.getElementById('columnSection'),
    columnGrid: document.getElementById('columnGrid'),
    actionSection: document.getElementById('actionSection'),
    startMapping: document.getElementById('startMapping'),
    resultsSection: document.getElementById('resultsSection'),

    // Results
    autoMapped: document.getElementById('autoMapped'),
    needsReview: document.getElementById('needsReview'),
    unmapped: document.getElementById('unmapped'),
    mappingTableBody: document.getElementById('mappingTableBody'),
    exportBtn: document.getElementById('exportBtn'),

    // Modal
    modal: document.getElementById('mappingModal'),
    modalOriginal: document.getElementById('modalOriginal'),
    modalSelect: document.getElementById('modalSelect'),
    leaveUnmapped: document.getElementById('leaveUnmapped'),
    closeModal: document.getElementById('closeModal'),
    cancelMapping: document.getElementById('cancelMapping'),
    confirmMapping: document.getElementById('confirmMapping'),

    // Tabs
    tabBtns: document.querySelectorAll('.tab-btn')
};

// ========================================
// Fuzzy Matching Utils
// ========================================

/**
 * Normalize a string for comparison
 */
function normalize(str) {
    if (!str) return '';
    return str
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '') // Remove accents
        .replace(/_/g, ' ')              // Replace underscores with spaces
        .replace(/\s+/g, ' ')            // Normalize whitespace
        .trim();
}

/**
 * Calculate Levenshtein distance between two strings
 */
function levenshteinDistance(a, b) {
    const matrix = [];

    for (let i = 0; i <= b.length; i++) {
        matrix[i] = [i];
    }

    for (let j = 0; j <= a.length; j++) {
        matrix[0][j] = j;
    }

    for (let i = 1; i <= b.length; i++) {
        for (let j = 1; j <= a.length; j++) {
            if (b.charAt(i - 1) === a.charAt(j - 1)) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(
                    matrix[i - 1][j - 1] + 1,
                    matrix[i][j - 1] + 1,
                    matrix[i - 1][j] + 1
                );
            }
        }
    }

    return matrix[b.length][a.length];
}

/**
 * Calculate similarity score (0-100)
 */
function calculateSimilarity(str1, str2) {
    const norm1 = normalize(str1);
    const norm2 = normalize(str2);

    // Exact match after normalization
    if (norm1 === norm2) return 100;

    // Check if one contains the other (without year)
    const withoutYear1 = norm1.replace(/\s*\d{4}\s*$/, '').trim();
    const withoutYear2 = norm2.replace(/\s*\d{4}\s*$/, '').trim();

    if (withoutYear1 === withoutYear2) return 95;

    // Levenshtein-based similarity
    const maxLen = Math.max(norm1.length, norm2.length);
    if (maxLen === 0) return 100;

    const distance = levenshteinDistance(norm1, norm2);
    const similarity = ((maxLen - distance) / maxLen) * 100;

    return Math.round(similarity);
}

/**
 * Find best match for a program
 */
function findBestMatch(program, officialList) {
    let bestMatch = null;
    let bestScore = 0;

    for (const official of officialList) {
        const score = calculateSimilarity(program, official);
        if (score > bestScore) {
            bestScore = score;
            bestMatch = official;
        }
    }

    return { match: bestMatch, score: bestScore };
}

// ========================================
// Column Detection
// ========================================

const COLUMN_PATTERNS = {
    program: ['program', 'programa', 'interes', 'carrera'],
    email: ['email', 'mail', 'correo', 'eml'],
    phone: ['tel', 'phone', 'celular', 'telefono', 'whatsapp'],
    contactId: ['idinterno', 'id', 'contacto', 'codigo', 'identificador'],
    database: ['iddatabase', 'database', 'base']
};

function detectColumns(headers) {
    const detected = {
        program: null,
        email: null,
        phone: null,
        contactId: null,
        database: null
    };

    for (const header of headers) {
        const normalizedHeader = normalize(header);

        // Program detection - highest priority for "Program aInteres"
        if (normalizedHeader.includes('program') && normalizedHeader.includes('interes')) {
            detected.program = header;
        } else if (!detected.program && COLUMN_PATTERNS.program.some(p => normalizedHeader.includes(p))) {
            detected.program = header;
        }

        // Email detection
        if (COLUMN_PATTERNS.email.some(p => normalizedHeader.includes(p))) {
            detected.email = header;
        }

        // Phone detection
        if (COLUMN_PATTERNS.phone.some(p => normalizedHeader.includes(p))) {
            detected.phone = header;
        }

        // Contact ID detection
        if (normalizedHeader.includes('id') && normalizedHeader.includes('contacto')) {
            detected.contactId = header;
        }

        // Database detection
        if (COLUMN_PATTERNS.database.some(p => normalizedHeader.includes(p))) {
            detected.database = header;
        }
    }

    return detected;
}

// ========================================
// File Handling
// ========================================

function handleFile(file) {
    const reader = new FileReader();

    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Get first sheet
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            // Convert to JSON
            const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: '' });

            if (jsonData.length === 0) {
                alert('El archivo está vacío o no tiene datos válidos.');
                return;
            }

            // Store data
            state.fileData = jsonData;
            state.headers = Object.keys(jsonData[0]);

            // Detect columns
            state.detectedColumns = detectColumns(state.headers);

            // Extract unique databases if database column exists
            state.uniqueDatabases = [];
            if (state.detectedColumns.database) {
                const databaseSet = new Set();
                for (const row of jsonData) {
                    const dbValue = row[state.detectedColumns.database];
                    if (dbValue !== null && dbValue !== undefined && dbValue !== '') {
                        databaseSet.add(dbValue.toString());
                    }
                }
                state.uniqueDatabases = [...databaseSet].sort((a, b) => Number(a) - Number(b));
            }

            // Update UI
            elements.dropZone.classList.add('hidden');
            elements.fileInfo.classList.remove('hidden');
            elements.fileName.textContent = file.name;
            elements.fileStats.textContent = `${jsonData.length} registros • ${state.headers.length} columnas`;

            // Show column detection
            renderColumnDetection();
            elements.columnSection.classList.remove('hidden');

            // Show action button if ready
            checkReady();

        } catch (error) {
            console.error('Error reading file:', error);
            alert('Error al leer el archivo. Asegúrate de que sea un archivo Excel o CSV válido.');
        }
    };

    reader.readAsArrayBuffer(file);
}

function renderColumnDetection() {
    const { detectedColumns, headers, uniqueDatabases } = state;

    const keyColumns = [
        { key: 'program', label: 'Programa', icon: '🎓' },
        { key: 'email', label: 'Email', icon: '📧' },
        { key: 'phone', label: 'Teléfono', icon: '📱' },
        { key: 'contactId', label: 'ID Contacto', icon: '🆔' },
        { key: 'database', label: 'Base de Datos', icon: '🗄️' }
    ];

    let html = keyColumns.map(col => {
        const detected = detectedColumns[col.key];
        return `
            <div class="column-tag ${detected ? 'detected' : ''}">
                <span class="icon">${col.icon}</span>
                <span>${col.label}:</span>
                <strong>${detected || 'No detectado'}</strong>
            </div>
        `;
    }).join('');

    // Add unique databases info if detected
    if (uniqueDatabases.length > 0) {
        html += `
            <div class="column-tag detected database-list">
                <span class="icon">📊</span>
                <span>Bases encontradas:</span>
                <strong>${uniqueDatabases.join(', ')}</strong>
                <span class="badge">${uniqueDatabases.length} base${uniqueDatabases.length > 1 ? 's' : ''}</span>
            </div>
        `;
    }

    elements.columnGrid.innerHTML = html;
}

// ========================================
// Program Loading
// ========================================

/**
 * Parse DAX DATATABLE format and extract program names
 * Handles format like: {"Program Name", "Type"},
 */
function parseDaxDatatable(text) {
    // Check if this looks like a DAX DATATABLE
    if (!text.includes('DATATABLE') && !text.includes('{') && !text.includes('"')) {
        return null; // Not DAX format
    }

    const programs = [];

    // Match rows like: {"Program Name", "Type"},
    // The program name is the first quoted value in each row
    const rowPattern = /\{\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\}/g;
    let match;

    while ((match = rowPattern.exec(text)) !== null) {
        const programName = match[1].trim();
        if (programName && programName.length > 0) {
            programs.push(programName);
        }
    }

    return programs.length > 0 ? programs : null;
}

function loadOfficialPrograms() {
    const text = elements.officialPrograms.value.trim();

    // First try to parse as DAX DATATABLE format
    let programs = parseDaxDatatable(text);

    if (programs) {
        console.log('📊 Formato DAX DATATABLE detectado, extrayendo nombres de programas...');
    } else {
        // Standard line-by-line parsing
        programs = text
            .split('\n')
            .map(p => p.trim())
            .filter(p => p.length > 0);
    }

    state.officialPrograms = programs;
    elements.programCount.textContent = `${programs.length} programas`;

    checkReady();
}

function checkReady() {
    const ready = state.officialPrograms.length > 0 &&
        state.fileData !== null &&
        state.detectedColumns.program !== null;

    if (ready) {
        elements.actionSection.classList.remove('hidden');
    } else {
        elements.actionSection.classList.add('hidden');
    }
}

// ========================================
// Mapping Logic
// ========================================

function startMapping() {
    const programColumn = state.detectedColumns.program;

    if (!programColumn) {
        alert('No se detectó la columna de programas. Revisa el archivo.');
        return;
    }

    // Extract unique programs and count occurrences
    state.programCounts.clear();
    state.mappings.clear();

    for (const row of state.fileData) {
        const program = row[programColumn];
        if (program && program.toString().trim()) {
            const trimmed = program.toString().trim();
            state.programCounts.set(trimmed, (state.programCounts.get(trimmed) || 0) + 1);
        }
    }

    // Map each unique program
    for (const [program] of state.programCounts) {
        const { match, score } = findBestMatch(program, state.officialPrograms);

        let status;
        let mapped;

        if (score >= 90) {
            status = 'mapped';
            mapped = match;
        } else if (score >= 70) {
            status = 'pending';
            mapped = match;
        } else {
            status = 'unmapped';
            mapped = null;
        }

        state.mappings.set(program, { mapped, score, status });
    }

    // Render results
    renderResults();
    elements.resultsSection.classList.remove('hidden');

    // Scroll to results
    elements.resultsSection.scrollIntoView({ behavior: 'smooth' });
}

function renderResults(filter = 'all') {
    // Update stats
    let autoMapped = 0;
    let pending = 0;
    let unmapped = 0;

    for (const [, mapping] of state.mappings) {
        if (mapping.status === 'mapped') autoMapped++;
        else if (mapping.status === 'pending') pending++;
        else unmapped++;
    }

    elements.autoMapped.textContent = `${autoMapped} auto-mapeados`;
    elements.needsReview.textContent = `${pending} pendientes`;
    elements.unmapped.textContent = `${unmapped} sin mapear`;

    // Render table
    let rows = [];

    for (const [original, mapping] of state.mappings) {
        if (filter !== 'all') {
            if (filter === 'mapped' && mapping.status !== 'mapped') continue;
            if (filter === 'pending' && mapping.status !== 'pending') continue;
            if (filter === 'unmapped' && mapping.status !== 'unmapped') continue;
        }

        const count = state.programCounts.get(original) || 0;
        const scoreClass = mapping.score >= 90 ? 'high' : mapping.score >= 70 ? 'medium' : 'low';

        let mappedCell;
        if (mapping.status === 'mapped') {
            mappedCell = `<span class="mapped">${mapping.mapped}</span>`;
        } else if (mapping.status === 'pending') {
            mappedCell = `<span class="pending">${mapping.mapped || 'Pendiente'}</span>`;
        } else {
            mappedCell = `<span class="unmapped">Sin mapear</span>`;
        }

        rows.push(`
            <tr data-original="${escapeHtml(original)}">
                <td class="original">${escapeHtml(original)}</td>
                <td>${mappedCell}</td>
                <td>
                    <div class="match-score">
                        <div class="score-bar">
                            <div class="score-fill ${scoreClass}" style="width: ${mapping.score}%"></div>
                        </div>
                        <span>${mapping.score}%</span>
                    </div>
                </td>
                <td>${count}</td>
                <td>
                    <button class="btn btn-secondary btn-small edit-btn">✏️ Editar</button>
                </td>
            </tr>
        `);
    }

    elements.mappingTableBody.innerHTML = rows.join('');

    // Add edit button handlers
    document.querySelectorAll('.edit-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const row = e.target.closest('tr');
            const original = row.dataset.original;
            openMappingModal(original);
        });
    });
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// ========================================
// Modal
// ========================================

function openMappingModal(original) {
    state.currentEditProgram = original;
    const mapping = state.mappings.get(original);

    // Set original program
    elements.modalOriginal.textContent = original;

    // Populate select
    elements.modalSelect.innerHTML = '<option value="">-- Seleccionar --</option>' +
        state.officialPrograms.map(p =>
            `<option value="${escapeHtml(p)}" ${mapping.mapped === p ? 'selected' : ''}>${escapeHtml(p)}</option>`
        ).join('');

    // Set checkbox
    elements.leaveUnmapped.checked = mapping.status === 'unmapped' && !mapping.mapped;

    // Show modal
    elements.modal.classList.remove('hidden');
}

function closeMappingModal() {
    elements.modal.classList.add('hidden');
    state.currentEditProgram = null;
}

function confirmMappingChange() {
    if (!state.currentEditProgram) return;

    const original = state.currentEditProgram;
    const selectedProgram = elements.modalSelect.value;
    const leaveUnmapped = elements.leaveUnmapped.checked;

    if (leaveUnmapped) {
        state.mappings.set(original, {
            mapped: null,
            score: 0,
            status: 'unmapped'
        });
    } else if (selectedProgram) {
        state.mappings.set(original, {
            mapped: selectedProgram,
            score: 100, // Manual mapping = 100%
            status: 'mapped'
        });
    }

    closeMappingModal();
    renderResults(getCurrentFilter());
}

function getCurrentFilter() {
    const activeTab = document.querySelector('.tab-btn.active');
    return activeTab ? activeTab.dataset.filter : 'all';
}

// ========================================
// Export
// ========================================

function exportMappedData() {
    const { program, email, contactId } = state.detectedColumns;

    // Create clean data with only 3 columns
    const cleanData = state.fileData.map(row => {
        const originalProgram = row[program]?.toString().trim();
        let mappedProgram = originalProgram || '';

        if (originalProgram) {
            const mapping = state.mappings.get(originalProgram);
            if (mapping && mapping.mapped) {
                mappedProgram = mapping.mapped;
            }
        }

        return {
            'ID Contacto': row[contactId] || '',
            'Email': row[email] || '',
            'Programa de Interés': mappedProgram
        };
    });

    // Create worksheet
    const worksheet = XLSX.utils.json_to_sheet(cleanData);

    // Set column widths
    worksheet['!cols'] = [
        { wch: 15 },  // ID Contacto
        { wch: 40 },  // Email
        { wch: 50 }   // Programa de Interés
    ];

    // Apply styles to cells (SheetJS community edition has limited styling)
    // We'll add some basic formatting
    const range = XLSX.utils.decode_range(worksheet['!ref']);

    // Style header cells
    for (let C = range.s.c; C <= range.e.c; C++) {
        const headerAddr = XLSX.utils.encode_cell({ r: 0, c: C });
        if (worksheet[headerAddr]) {
            worksheet[headerAddr].s = {
                font: { bold: true, color: { rgb: "FFFFFF" } },
                fill: { fgColor: { rgb: "6366F1" } },
                alignment: { horizontal: "center", vertical: "center" },
                border: {
                    top: { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } },
                    left: { style: "thin", color: { rgb: "000000" } },
                    right: { style: "thin", color: { rgb: "000000" } }
                }
            };
        }
    }

    // Style data cells
    for (let R = 1; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
            const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
            if (worksheet[cellAddr]) {
                worksheet[cellAddr].s = {
                    alignment: { horizontal: C === 0 ? "center" : "left", vertical: "center" },
                    border: {
                        top: { style: "thin", color: { rgb: "E5E5E5" } },
                        bottom: { style: "thin", color: { rgb: "E5E5E5" } },
                        left: { style: "thin", color: { rgb: "E5E5E5" } },
                        right: { style: "thin", color: { rgb: "E5E5E5" } }
                    }
                };
            }
        }
    }

    // Create workbook
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Programas Mapeados');

    // Generate filename with date
    const date = new Date().toISOString().slice(0, 10);
    const filename = `programas_mapeados_${date}.xlsx`;

    // Download
    XLSX.writeFile(workbook, filename);

    // Show success message
    const totalMapped = [...state.mappings.values()].filter(m => m.mapped).length;
    alert(`✅ Excel exportado exitosamente!\n\n📊 ${cleanData.length} registros\n🎯 ${totalMapped} programas mapeados`);
}

// ========================================
// Event Listeners
// ========================================

// Load programs button
elements.loadPrograms.addEventListener('click', loadOfficialPrograms);

// Textarea change
elements.officialPrograms.addEventListener('input', () => {
    const lines = elements.officialPrograms.value.split('\n').filter(l => l.trim()).length;
    elements.programCount.textContent = `${lines} programas`;
});

// Drop zone
elements.dropZone.addEventListener('click', () => elements.fileInput.click());

elements.dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    elements.dropZone.classList.add('dragover');
});

elements.dropZone.addEventListener('dragleave', () => {
    elements.dropZone.classList.remove('dragover');
});

elements.dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    elements.dropZone.classList.remove('dragover');

    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
});

elements.fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handleFile(file);
});

// Remove file
elements.removeFile.addEventListener('click', () => {
    state.fileData = null;
    state.headers = [];
    state.detectedColumns = { program: null, email: null, phone: null, contactId: null, database: null };
    state.uniqueDatabases = [];

    elements.fileInfo.classList.add('hidden');
    elements.dropZone.classList.remove('hidden');
    elements.columnSection.classList.add('hidden');
    elements.actionSection.classList.add('hidden');
    elements.resultsSection.classList.add('hidden');
    elements.fileInput.value = '';
});

// Start mapping
elements.startMapping.addEventListener('click', startMapping);

// Filter tabs
elements.tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        elements.tabBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        renderResults(btn.dataset.filter);
    });
});

// Modal
elements.closeModal.addEventListener('click', closeMappingModal);
elements.cancelMapping.addEventListener('click', closeMappingModal);
elements.confirmMapping.addEventListener('click', confirmMappingChange);

elements.modal.querySelector('.modal-overlay').addEventListener('click', closeMappingModal);

// Checkbox behavior
elements.leaveUnmapped.addEventListener('change', () => {
    if (elements.leaveUnmapped.checked) {
        elements.modalSelect.value = '';
        elements.modalSelect.disabled = true;
    } else {
        elements.modalSelect.disabled = false;
    }
});

// Export
elements.exportBtn.addEventListener('click', exportMappedData);

// Keyboard shortcuts
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !elements.modal.classList.contains('hidden')) {
        closeMappingModal();
    }
});

// ========================================
// Initialize
// ========================================
console.log('🎯 Mapeador de Programas Universitarios v1.0 iniciado');
