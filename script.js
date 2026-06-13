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
const installBtn = document.getElementById('install-btn');

// État de l'application
let currentMode = 'word-to-pdf'; 
let selectedFile = null;
let convertedBlob = null;
let convertedFileName = "";
let deferredPrompt;

// --- PWA Service Worker Registration ---
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('./sw.js')
            .then(reg => console.log('Service Worker enregistré !'))
            .catch(err => console.log('Erreur Service Worker:', err));
    });
}

// --- PWA Installation Logic ---
window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    installBtn.classList.remove('hidden');
});

installBtn.addEventListener('click', async () => {
    if (deferredPrompt) {
        deferredPrompt.prompt();
        const { outcome } = await deferredPrompt.userChoice;
        if (outcome === 'accepted') {
            installBtn.classList.add('hidden');
        }
        deferredPrompt = null;
    }
});

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

// --- Gestion du Drag & Drop ---
dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    handleFiles(e.dataTransfer.files);
});

fileInput.addEventListener('change', (e) => handleFiles(e.target.files));

function handleFiles(files) {
    if (files.length === 0) return;
    const file = files[0];
    const extension = file.name.split('.').pop().toLowerCase();

    if (currentMode === 'word-to-pdf' && extension !== 'docx') {
        alert("Veuillez sélectionner un fichier .docx");
        return;
    }
    if (currentMode === 'pdf-to-word' && extension !== 'pdf') {
        alert("Veuillez sélectionner un fichier .pdf");
        return;
    }

    selectedFile = file;
    fileNameDisplay.textContent = file.name;
    dropZone.classList.add('hidden');
    processSection.classList.remove('hidden');
    convertBtn.classList.remove('hidden');
    downloadSection.classList.add('hidden');
    updateProgress(0);
    statusText.textContent = "Fichier prêt.";
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
        console.error("Erreur de conversion:", error);
        statusText.textContent = "Erreur : " + error.message;
        convertBtn.classList.remove('hidden');
        progressBar.style.backgroundColor = 'var(--error-color)';
    }
});

async function convertWordToPdf(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async function(e) {
            try {
                const arrayBuffer = e.target.result;
                updateProgress(40);
                
                const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
                const html = result.value;
                
                const element = document.createElement('div');
                element.innerHTML = `<div style="padding: 40px; font-family: Arial, sans-serif;">${html}</div>`;
                document.body.appendChild(element);
                element.style.position = 'absolute';
                element.style.left = '-9999px';

                updateProgress(70);

                const opt = {
                    margin: 10,
                    filename: file.name.replace('.docx', '.pdf'),
                    image: { type: 'jpeg', quality: 0.98 },
                    html2canvas: { scale: 2, useCORS: true, logging: false },
                    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
                };

                html2pdf().from(element).set(opt).toPdf().get('pdf').then((pdf) => {
                    convertedBlob = pdf.output('blob');
                    convertedFileName = file.name.replace('.docx', '.pdf');
                    document.body.removeChild(element);
                    showDownload();
                    resolve();
                }).catch(reject);
            } catch (err) {
                reject(new Error("Impossible de lire le fichier Word."));
            }
        };
        reader.onerror = () => reject(new Error("Erreur de lecture du fichier."));
        reader.readAsArrayBuffer(file);
    });
}

async function convertPdfToWord(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async function(e) {
            try {
                const typedarray = new Uint8Array(e.target.result);
                const loadingTask = pdfjsLib.getDocument(typedarray);
                
                const pdf = await loadingTask.promise;
                let fullText = "";

                updateProgress(30);
                
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    const pageText = textContent.items.map(item => item.str).join(' ');
                    fullText += pageText + "\n\n";
                    updateProgress(30 + (i / pdf.numPages) * 50);
                }

                if (!fullText.trim()) {
                    throw new Error("Aucun texte extractible trouvé dans ce PDF.");
                }

                const doc = new docx.Document({
                    sections: [{
                        children: fullText.split('\n').map(line => 
                            new docx.Paragraph({
                                children: [new docx.TextRun(line.trim() || " ")],
                            })
                        ),
                    }],
                });

                convertedBlob = await docx.Packer.toBlob(doc);
                convertedFileName = file.name.replace('.pdf', '.docx');
                
                showDownload();
                resolve();
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = () => reject(new Error("Erreur de lecture du PDF."));
        reader.readAsArrayBuffer(file);
    });
}

function updateProgress(percent) {
    progressBar.style.width = percent + '%';
}

function showDownload() {
    updateProgress(100);
    statusText.textContent = "Terminé !";
    downloadSection.classList.remove('hidden');
    selectedFile = null;
}

downloadBtn.addEventListener('click', () => {
    if (convertedBlob) {
        saveAs(convertedBlob, convertedFileName);
    }
});

resetBtn.addEventListener('click', () => resetUI());

function resetUI() {
    selectedFile = null;
    convertedBlob = null;
    convertedFileName = "";
    dropZone.classList.remove('hidden');
    processSection.classList.add('hidden');
    downloadSection.classList.add('hidden');
    updateProgress(0);
    progressBar.style.backgroundColor = '';
    fileInput.value = "";
    updateUIForMode();
}
