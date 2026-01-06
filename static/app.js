/**
 * PDF to PowerPoint Converter - JSON Mode Frontend
 * Allows uploading both PDF and JSON files for direct conversion
 */

const API_BASE = window.location.origin;

// State
let currentJobId = null;
let selectedPdfFile = null;
let selectedJsonFile = null;
let pollInterval = null;
let selectedMode = 'precision';

// DOM Elements
const uploadSection = document.getElementById('upload-section');
const processingSection = document.getElementById('processing-section');
const resultSection = document.getElementById('result-section');

const pdfUploadZone = document.getElementById('pdf-upload-zone');
const pdfFileInput = document.getElementById('pdf-file-input');
const pdfFileInfo = document.getElementById('pdf-file-info');
const pdfFileName = document.getElementById('pdf-file-name');
const removePdfFileBtn = document.getElementById('remove-pdf-file');

const jsonUploadZone = document.getElementById('json-upload-zone');
const jsonFileInput = document.getElementById('json-file-input');
const jsonFileInfo = document.getElementById('json-file-info');
const jsonFileName = document.getElementById('json-file-name');
const removeJsonFileBtn = document.getElementById('remove-json-file');

const modeSelector = document.getElementById('mode-selector');
const convertBtn = document.getElementById('convert-btn');

const progressFill = document.getElementById('progress-fill');
const progressText = document.getElementById('progress-text');
const processingTitle = document.getElementById('processing-title');
const processingMessage = document.getElementById('processing-message');

const resultFilename = document.getElementById('result-filename');
const downloadBtn = document.getElementById('download-btn');
const newConvertBtn = document.getElementById('new-convert-btn');

// Steps (simplified for JSON mode - no analyze step)
const steps = {
    upload: document.getElementById('step-upload'),
    generate: document.getElementById('step-generate'),
    complete: document.getElementById('step-complete')
};

// ===== Event Listeners =====

// PDF Upload Zone
pdfUploadZone.addEventListener('click', () => pdfFileInput.click());
pdfFileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handlePdfSelect(e.target.files[0]);
});

// JSON Upload Zone
jsonUploadZone.addEventListener('click', () => jsonFileInput.click());
jsonFileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleJsonSelect(e.target.files[0]);
});

// Drag & Drop for PDF
pdfUploadZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    pdfUploadZone.classList.add('drag-over');
});
pdfUploadZone.addEventListener('dragleave', () => pdfUploadZone.classList.remove('drag-over'));
pdfUploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    pdfUploadZone.classList.remove('drag-over');
    const files = e.dataTransfer.files;
    if (files.length > 0 && files[0].type === 'application/pdf') {
        handlePdfSelect(files[0]);
    } else {
        showError('Please drop a PDF file');
    }
});

// Drag & Drop for JSON
jsonUploadZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    jsonUploadZone.classList.add('drag-over');
});
jsonUploadZone.addEventListener('dragleave', () => jsonUploadZone.classList.remove('drag-over'));
jsonUploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    jsonUploadZone.classList.remove('drag-over');
    const files = e.dataTransfer.files;
    if (files.length > 0 && files[0].name.endsWith('.json')) {
        handleJsonSelect(files[0]);
    } else {
        showError('Please drop a JSON file');
    }
});

// Remove buttons
removePdfFileBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    resetPdfUpload();
});
removeJsonFileBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    resetJsonUpload();
});

// Mode Selection
document.querySelectorAll('input[name="conversion-mode"]').forEach(radio => {
    radio.addEventListener('change', (e) => {
        selectedMode = e.target.value;
    });
});

// Convert Button
convertBtn.addEventListener('click', startConversion);

// Download Button
downloadBtn.addEventListener('click', downloadResult);

// New Convert Button
newConvertBtn.addEventListener('click', resetAll);

// ===== Functions =====

function handlePdfSelect(file) {
    if (file.type !== 'application/pdf') {
        showError('Please select a PDF file');
        return;
    }

    selectedPdfFile = file;
    pdfFileName.textContent = file.name;
    pdfUploadZone.classList.add('hidden');
    pdfFileInfo.classList.remove('hidden');

    // Show JSON upload zone
    jsonUploadZone.classList.remove('hidden');

    updateConvertButtonVisibility();
}

function handleJsonSelect(file) {
    if (!file.name.endsWith('.json')) {
        showError('Please select a JSON file');
        return;
    }

    selectedJsonFile = file;
    jsonFileName.textContent = file.name;
    jsonUploadZone.classList.add('hidden');
    jsonFileInfo.classList.remove('hidden');

    updateConvertButtonVisibility();
}

function updateConvertButtonVisibility() {
    if (selectedPdfFile && selectedJsonFile) {
        modeSelector.classList.remove('hidden');
        convertBtn.classList.remove('hidden');
    } else {
        modeSelector.classList.add('hidden');
        convertBtn.classList.add('hidden');
    }
}

function resetPdfUpload() {
    selectedPdfFile = null;
    pdfFileInput.value = '';
    pdfFileName.textContent = '';
    pdfUploadZone.classList.remove('hidden');
    pdfFileInfo.classList.add('hidden');

    // Hide JSON zone if PDF is removed
    jsonUploadZone.classList.add('hidden');
    resetJsonUpload();

    updateConvertButtonVisibility();
}

function resetJsonUpload() {
    selectedJsonFile = null;
    jsonFileInput.value = '';
    jsonFileName.textContent = '';

    if (selectedPdfFile) {
        jsonUploadZone.classList.remove('hidden');
    }
    jsonFileInfo.classList.add('hidden');

    updateConvertButtonVisibility();
}

async function startConversion() {
    if (!selectedPdfFile || !selectedJsonFile) return;

    try {
        // Show processing section
        uploadSection.classList.add('hidden');
        processingSection.classList.remove('hidden');

        updateStep('upload');
        updateProgress(5, 'Uploading files...');

        // Upload both files
        const formData = new FormData();
        formData.append('pdf_file', selectedPdfFile);
        formData.append('json_file', selectedJsonFile);
        formData.append('mode', selectedMode);

        const uploadResponse = await fetch(`${API_BASE}/api/upload`, {
            method: 'POST',
            body: formData
        });

        if (!uploadResponse.ok) {
            const err = await uploadResponse.json();
            throw new Error(err.detail || 'Upload failed');
        }

        const uploadData = await uploadResponse.json();
        currentJobId = uploadData.job_id;

        updateProgress(15, 'Upload complete. Starting conversion...');
        completeStep('upload');
        updateStep('generate');

        // Start processing
        const processResponse = await fetch(`${API_BASE}/api/process/${currentJobId}`, {
            method: 'POST'
        });

        if (!processResponse.ok) {
            throw new Error('Failed to start processing');
        }

        // Start polling for status
        startPolling();

    } catch (error) {
        showError(error.message);
        resetAll();
    }
}

function startPolling() {
    pollInterval = setInterval(async () => {
        try {
            const response = await fetch(`${API_BASE}/api/status/${currentJobId}`);

            if (!response.ok) {
                throw new Error('Failed to get status');
            }

            const status = await response.json();
            handleStatusUpdate(status);

        } catch (error) {
            console.error('Polling error:', error);
        }
    }, 1000);
}

function stopPolling() {
    if (pollInterval) {
        clearInterval(pollInterval);
        pollInterval = null;
    }
}

function handleStatusUpdate(status) {
    updateProgress(status.progress, status.message);

    switch (status.status) {
        case 'processing':
        case 'generating':
            updateStep('generate');
            break;
        case 'completed':
            stopPolling();
            completeStep('generate');
            completeStep('complete');
            showResult(status);
            break;
        case 'error':
            stopPolling();
            showError(status.message);
            break;
    }
}

function updateProgress(percent, message) {
    progressFill.style.width = `${percent}%`;
    progressText.textContent = `${Math.round(percent)}%`;
    if (message) {
        processingMessage.textContent = message;
    }
}

function updateStep(stepId) {
    Object.values(steps).forEach(step => step.classList.remove('active'));
    if (steps[stepId]) {
        steps[stepId].classList.add('active');
    }
}

function completeStep(stepId) {
    if (steps[stepId]) {
        steps[stepId].classList.add('completed');
    }
}

function showResult(status) {
    processingSection.classList.add('hidden');
    resultSection.classList.remove('hidden');
    resultFilename.textContent = status.output_filename || 'presentation.pptx';
}

async function downloadResult() {
    if (!currentJobId) return;

    // Simply open the download URL directly - browser handles the download
    window.location.href = `${API_BASE}/api/download/${currentJobId}`;
}

function resetAll() {
    stopPolling();
    currentJobId = null;
    selectedPdfFile = null;
    selectedJsonFile = null;

    // Reset UI
    uploadSection.classList.remove('hidden');
    processingSection.classList.add('hidden');
    resultSection.classList.add('hidden');

    // Reset file inputs
    pdfUploadZone.classList.remove('hidden');
    pdfFileInfo.classList.add('hidden');
    pdfFileInput.value = '';
    pdfFileName.textContent = '';

    jsonUploadZone.classList.add('hidden');
    jsonFileInfo.classList.add('hidden');
    jsonFileInput.value = '';
    jsonFileName.textContent = '';

    modeSelector.classList.add('hidden');
    convertBtn.classList.add('hidden');

    // Reset mode
    selectedMode = 'precision';
    document.querySelector('input[name="conversion-mode"][value="precision"]').checked = true;

    // Reset progress
    updateProgress(0, 'Initializing...');

    // Reset steps
    Object.values(steps).forEach(step => {
        step.classList.remove('active', 'completed');
    });
}

function showError(message) {
    alert(`Error: ${message}`);
    console.error(message);
}

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    console.log('PDF to PowerPoint Converter (JSON Mode) initialized');
});
