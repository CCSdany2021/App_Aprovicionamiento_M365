let currentAction = '';
let selectedFile = null;

// DOM Elements
const uploadSection = document.getElementById('upload-section');
const resultsSection = document.getElementById('results-section');
const cardsGrid = document.querySelector('.cards-grid');
const actionTitle = document.getElementById('action-title');
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const fileNameDisplay = document.getElementById('file-name');
const processBtn = document.getElementById('process-btn');
const loadingOverlay = document.getElementById('loading');

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    // Drag and Drop events
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
    });

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, unhighlight, false);
    });

    dropZone.addEventListener('drop', handleDrop, false);
    fileInput.addEventListener('change', handleFileSelect, false);
});

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

function highlight(e) {
    dropZone.classList.add('dragover');
}

function unhighlight(e) {
    dropZone.classList.remove('dragover');
}

function selectAction(action) {
    currentAction = action;
    cardsGrid.classList.add('hidden');
    uploadSection.classList.remove('hidden');
    
    const titles = {
        'crear': 'Crear Estudiantes Nuevos',
        'actualizar': 'Actualizar Datos de Estudiantes',
        'eliminar': 'Eliminar Usuarios'
    };
    
    actionTitle.innerText = titles[action];
}

function goBack() {
    uploadSection.classList.add('hidden');
    resultsSection.classList.add('hidden');
    cardsGrid.classList.remove('hidden');
    resetFile();
}

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    handleFiles(files);
}

function handleFileSelect(e) {
    const files = e.target.files;
    handleFiles(files);
}

function handleFiles(files) {
    if (files.length > 0) {
        selectedFile = files[0];
        if (selectedFile.name.endsWith('.xlsx') || selectedFile.name.endsWith('.csv')) {
            fileNameDisplay.innerText = `Archivo seleccionado: ${selectedFile.name}`;
            processBtn.disabled = false;
        } else {
            alert('Por favor selecciona un archivo Excel (.xlsx) o CSV válida.');
            resetFile();
        }
    }
}

function resetFile() {
    selectedFile = null;
    fileInput.value = '';
    fileNameDisplay.innerText = '';
    processBtn.disabled = true;
}

function resetAll() {
    goBack();
}

async function processFile() {
    if (!selectedFile || !currentAction) return;

    showLoading(true);
    
    const formData = new FormData();
    formData.append('file', selectedFile);

    try {
        const response = await fetch(`/upload/${currentAction}`, {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        if (data.success) {
            showResults(data.resultados);
        } else {
            alert('Error: ' + data.error);
        }

    } catch (error) {
        alert('Error de conexión: ' + error);
    } finally {
        showLoading(false);
    }
}

function showLoading(show) {
    if (show) {
        loadingOverlay.classList.remove('hidden');
    } else {
        loadingOverlay.classList.add('hidden');
    }
}

function showResults(resultados) {
    uploadSection.classList.add('hidden');
    resultsSection.classList.remove('hidden');

    // Update Stats
    document.getElementById('stat-total').innerText = resultados.summary.total;
    
    let successCount = 0;
    if (currentAction === 'crear') successCount = resultados.summary.creados;
    if (currentAction === 'actualizar') successCount = resultados.summary.actualizados;
    if (currentAction === 'eliminar') successCount = resultados.summary.eliminados;
    
    document.getElementById('stat-success').innerText = successCount;
    document.getElementById('stat-error').innerText = resultados.summary.errores;

    // Populate Table
    const tbody = document.getElementById('log-body');
    tbody.innerHTML = '';

    resultados.details.forEach(item => {
        const row = document.createElement('tr');
        const statusClass = item.estado.toLowerCase().includes('error') ? 'text-danger' : 'text-success';
        
        row.innerHTML = `
            <td>${item.codigo}</td>
            <td>${item.nombre}</td>
            <td style="color: ${item.estado.toLowerCase().includes('error') ? '#ef4444' : '#10b981'}">${item.estado}</td>
            <td>${item.mensaje}</td>
        `;
        tbody.appendChild(row);
    });
}
