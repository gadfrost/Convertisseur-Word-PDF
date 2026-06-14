// --- 1. CONFIGURATION INITIALE ---
const pdfjsLib = window['pdfjs-dist/build/pdf'];
if (pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
}

// --- 2. ÉLÉMENTS DU DOM ---
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

// --- 3. ÉTAT ---
let currentMode = 'word-to-pdf'; 
let selectedFile = null;
let convertedBlob = null;
let convertedFileName = "";
let deferredPrompt;

// --- 4. PWA ---
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('./sw.js').catch(err => console.error('SW Error:', err));
    });
}

window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
});

if (installBtn) {
    installBtn.addEventListener('click', async () => {
        if (deferredPrompt) {
            deferredPrompt.prompt();
            deferredPrompt = null;
        } else {
            alert("Utilisez le menu de votre navigateur pour installer l'application.");
        }
    });
}

// --- 5. NAVIGATION ---
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
        fileInput.accept = ".docx";
    } else {
        fileInfoText.textContent = "Fichiers acceptés : .pdf";
        fileIcon.className = "fas fa-file-pdf";
        fileInput.accept = ".pdf";
    }
}

// --- 6. GESTION DES FICHIERS ---
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
        alert("Veuillez choisir un fichier .docx");
        return;
    }
    if (currentMode === 'pdf-to-word' && extension !== 'pdf') {
        alert("Veuillez choisir un fichier .pdf");
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

// --- 7. CONVERSION ---
convertBtn.addEventListener('click', async () => {
    if (!selectedFile) return;
    convertBtn.classList.add('hidden');
    statusText.textContent = "Traitement en cours...";
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
        progressBar.style.backgroundColor = '#ef4444';
    }
});

async function convertWordToPdf(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async function(e) {
            try {
                const arrayBuffer = e.target.result;
                updateProgress(30);
                
                const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
                const html = result.value;
                
                // Nettoyage du HTML pour éviter les répétitions ou désordres
                const cleanHtml = html.replace(/<p><\/p>/g, '<br/>');

                const container = document.createElement('div');
                container.innerHTML = `<div style="padding: 30px; font-family: Arial, sans-serif;">${cleanHtml}</div>`;
                
                // Important : ne pas ajouter au body pour éviter les interférences visuelles
                updateProgress(60);

                const opt = {
                    margin: 10,
                    filename: file.name.replace('.docx', '.pdf'),
                    image: { type: 'jpeg', quality: 0.98 },
                    html2canvas: { scale: 2, useCORS: true },
                    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
                };

                html2pdf().from(container).set(opt).outputPdf('blob').then(blob => {
                    convertedBlob = blob;
                    convertedFileName = file.name.replace('.docx', '.pdf');
                    showDownload();
                    resolve();
                }).catch(reject);
            } catch (err) {
                reject(err);
            }
        };
        reader.readAsArrayBuffer(file);
    });
}

async function convertPdfToWord(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async function(e) {
            try {
                const docxLib = window.docx || window.DOCX || docx;
                if (!docxLib) throw new Error("Bibliothèque DOCX manquante.");

                const typedarray = new Uint8Array(e.target.result);
                const pdf = await pdfjsLib.getDocument(typedarray).promise;
                let fullText = "";

                updateProgress(20);

                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    
                    // Reconstruction du texte par lignes
                    let lastY = -1;
                    let pageText = "";
                    for (const item of textContent.items) {
                        if (lastY !== -1 && Math.abs(item.transform[5] - lastY) > 5) {
                            pageText += "\n";
                        }
                        pageText += item.str + " ";
                        lastY = item.transform[5];
                    }
                    fullText += pageText + "\n\n";
                    updateProgress(20 + (i / pdf.numPages) * 50);
                }

                // Utilisation de docx.js pour créer un VRAI fichier .docx
                const doc = new docxLib.Document({
                    sections: [{
                        children: fullText.split('\n').map(line => 
                            new docxLib.Paragraph({
                                children: [new docxLib.TextRun(line.trim() || " ")],
                            })
                        ),
                    }],
                });

                const blob = await docxLib.Packer.toBlob(doc);
                convertedBlob = blob;
                convertedFileName = file.name.replace('.pdf', '.docx');
                showDownload();
                resolve();
            } catch (err) {
                reject(err);
            }
        };
        reader.readAsArrayBuffer(file);
    });
}

function updateProgress(percent) {
    progressBar.style.width = percent + '%';
}

function showDownload() {
    updateProgress(100);
    statusText.textContent = "Conversion réussie !";
    downloadSection.classList.remove('hidden');
}

downloadBtn.addEventListener('click', () => {
    if (convertedBlob) saveAs(convertedBlob, convertedFileName);
});

resetBtn.addEventListener('click', () => resetUI());

function resetUI() {
    selectedFile = null;
    convertedBlob = null;
    fileInput.value = "";
    dropZone.classList.remove('hidden');
    processSection.classList.add('hidden');
    downloadSection.classList.add('hidden');
    updateProgress(0);
    progressBar.style.backgroundColor = '';
}
