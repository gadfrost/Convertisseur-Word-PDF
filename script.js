// Configuration de PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// Éléments DOM
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const tabBtns = document.querySelectorAll('.tab-btn');
const processSection = document.getElementById('process-section');
const downloadSection = document.getElementById('download-section');
const convertBtn = document.getElementById('convert-btn');
const downloadBtn = document.getElementById('download-btn');
const resetBtn = document.getElementById('reset-btn');
const progressBar = document.getElementById('progress-bar');
const statusText = document.getElementById('status-text');
const fileNameDisplay = document.getElementById('file-name');
const fileIcon = document.getElementById('file-icon');
const fileInfoText = document.getElementById('file-info-text');

// État de l'application
let currentMode = 'word-to-pdf'; // 'word-to-pdf' ou 'pdf-to-word'
let selectedFile = null;
let convertedBlob = null;
let convertedFileName = "";

// --- Gestion des Onglets ---
tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        tabBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentMode = btn.dataset.mode;
        updateUIForMode();
        resetUI();
    });
});

function updateUIForMode() {
    if (currentMode === 'word-to-pdf') {
        fileInfoText.textContent = "Fichiers acceptés : .docx";
        fileIcon.className = "fas fa-file-word";
    } else {
        fileInfoText.textContent = "Fichiers acceptés : .pdf";
        fileIcon.className = "fas fa-file-pdf";
    }
}

// --- Gestion du Drag & Drop et Sélection de Fichier ---
dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    handleFiles(e.dataTransfer.files);
});

fileInput.addEventListener('change', (e) => {
    handleFiles(e.target.files);
});

function handleFiles(files) {
    if (files.length === 0) return;
    const file = files[0];
    const extension = file.name.split('.').pop().toLowerCase();

    // Validation simple
    if (currentMode === 'word-to-pdf' && extension !== 'docx') {
        alert("Veuillez sélectionner un fichier .docx pour ce mode.");
        return;
    }
    if (currentMode === 'pdf-to-word' && extension !== 'pdf') {
        alert("Veuillez sélectionner un fichier .pdf pour ce mode.");
        return;
    }

    selectedFile = file;
    fileNameDisplay.textContent = file.name;
    dropZone.classList.add('hidden');
    processSection.classList.remove('hidden');
    convertBtn.classList.remove('hidden');
    downloadSection.classList.add('hidden');
    progressBar.style.width = '0%';
    statusText.textContent = "Fichier chargé. Prêt pour la conversion.";
}

// --- Logique de Conversion ---
convertBtn.addEventListener('click', async () => {
    if (!selectedFile) return;

    convertBtn.classList.add('hidden');
    statusText.textContent = "Conversion en cours...";
    updateProgress(10);

    try {
        if (currentMode === 'word-to-pdf') {
            await convertWordToPdf(selectedFile);
        } else {
            await convertPdfToWord(selectedFile);
        }
    } catch (error) {
        console.error(error);
        statusText.textContent = "Erreur : " + error.message;
        convertBtn.classList.remove('hidden');
        progressBar.style.backgroundColor = 'var(--error-color)';
    }
});

// Conversion Word (docx) -> PDF
async function convertWordToPdf(file) {
    updateProgress(30);
    const reader = new FileReader();
    
    reader.onload = async function(e) {
        try {
            const arrayBuffer = e.target.result;
            updateProgress(50);
            
            // 1. Convertir docx en HTML via Mammoth
            const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
            const html = result.value;
            
            // Créer un élément temporaire pour le rendu PDF
            const element = document.createElement('div');
            element.innerHTML = `<div style="padding: 40px; font-family: Arial, sans-serif;">${html}</div>`;
            document.body.appendChild(element);
            element.style.position = 'absolute';
            element.style.left = '-9999px';

            updateProgress(70);

            // 2. Convertir HTML en PDF via html2pdf
            const opt = {
                margin: 10,
                filename: file.name.replace('.docx', '.pdf'),
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2 },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
            };

            const pdfWorker = html2pdf().from(element).set(opt);
            convertedBlob = await pdfWorker.output('blob');
            convertedFileName = file.name.replace('.docx', '.pdf');

            document.body.removeChild(element);
            showDownload();
        } catch (err) {
            throw new Error("Échec de la conversion Word vers PDF.");
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// Conversion PDF -> Word (docx)
async function convertPdfToWord(file) {
    updateProgress(20);
    const reader = new FileReader();
    
    reader.onload = async function(e) {
        try {
            const typedarray = new Uint8Array(e.target.result);
            const pdf = await pdfjsLib.getDocument(typedarray).promise;
            let fullText = "";

            updateProgress(40);
            
            // Extraire le texte de chaque page
            for (let i = 1; i <= pdf.numPages; i++) {
                const page = await pdf.getPage(i);
                const textContent = await page.getTextContent();
                const pageText = textContent.items.map(item => item.str).join(' ');
                fullText += pageText + "\n\n";
                updateProgress(40 + (i / pdf.numPages) * 40);
            }

            // Générer le fichier .docx via docx.js
            const doc = new docx.Document({
                sections: [{
                    properties: {},
                    children: fullText.split('\n').map(line => 
                        new docx.Paragraph({
                            children: [new docx.TextRun(line)],
                        })
                    ),
                }],
            });

            convertedBlob = await docx.Packer.toBlob(doc);
            convertedFileName = file.name.replace('.pdf', '.docx');
            
            showDownload();
        } catch (err) {
            throw new Error("Échec de l'extraction du texte du PDF.");
        }
    };
    
    reader.readAsArrayBuffer(file);
}

function updateProgress(percent) {
    progressBar.style.width = percent + '%';
}

function showDownload() {
    updateProgress(100);
    statusText.textContent = "Terminé !";
    downloadSection.classList.remove('hidden');
    
    // Libération de la mémoire du fichier original
    selectedFile = null;
}

downloadBtn.addEventListener('click', () => {
    if (convertedBlob) {
        saveAs(convertedBlob, convertedFileName);
    }
});

resetBtn.addEventListener('click', () => {
    resetUI();
});

function resetUI() {
    selectedFile = null;
    convertedBlob = null;
    convertedFileName = "";
    dropZone.classList.remove('hidden');
    processSection.classList.add('hidden');
    downloadSection.classList.add('hidden');
    progressBar.style.width = '0%';
    progressBar.style.backgroundColor = '';
    fileInput.value = "";
    updateUIForMode();
}
