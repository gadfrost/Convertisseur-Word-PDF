// --- Configuration et injection automatique des CDN requis ---
(function injectRequiredLibraries() {
    const libraries = [
        { id: 'mammoth-cdn', src: "https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js" },
        { id: 'html2pdf-cdn', src: "https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js" },
        { id: 'pdfjs-cdn', src: "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js" },
        { id: 'filesaver-cdn', src: "https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js" }
    ];

    libraries.forEach(lib => {
        if (!document.getElementById(lib.id)) {
            const script = document.createElement('script');
            script.id = lib.id;
            script.src = lib.src;
            script.async = false;
            document.head.appendChild(script);
        }
    });
})();

// Configuration de PDF.js worker (sécurisé après chargement du script)
window.addEventListener('load', () => {
    if (typeof pdfjsLib !== 'undefined') {
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    }
});

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
            .then(reg => console.log('Service Worker enregistré avec succès !'))
            .catch(err => console.error('Erreur d\'enregistrement du Service Worker:', err));
    });
}

// --- Logique d'installation de la PWA ---
window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    if (installBtn) installBtn.classList.remove('hidden');
});

if (installBtn) {
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
}

// --- Gestion des Onglets (Modes) ---
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
        if (fileInfoText) fileInfoText.textContent = "Fichiers acceptés : .docx";
        if (fileIcon) fileIcon.className = "fas fa-file-word main-icon";
        if (fileInput) fileInput.accept = ".docx";
    } else {
        if (fileInfoText) fileInfoText.textContent = "Fichiers acceptés : .pdf";
        if (fileIcon) fileIcon.className = "fas fa-file-pdf main-icon";
        if (fileInput) fileInput.accept = ".pdf";
    }
}

// --- Gestion du Drag & Drop ---
if (dropZone) {
    dropZone.addEventListener('click', () => { if (fileInput) fileInput.click(); });
    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        handleFiles(e.dataTransfer.files);
    });
}

if (fileInput) {
    fileInput.addEventListener('change', (e) => handleFiles(e.target.files));
}

function handleFiles(files) {
    if (files.length === 0) return;
    const file = files[0];
    const extension = file.name.split('.').pop().toLowerCase();

    if (currentMode === 'word-to-pdf' && extension !== 'docx') {
        alert("Veuillez sélectionner un fichier valide avec l'extension .docx");
        return;
    }
    if (currentMode === 'pdf-to-word' && extension !== 'pdf') {
        alert("Veuillez sélectionner un fichier valide avec l'extension .pdf");
        return;
    }

    selectedFile = file;
    if (fileNameDisplay) fileNameDisplay.textContent = file.name;
    
    if (dropZone) dropZone.classList.add('hidden');
    if (processSection) processSection.classList.remove('hidden');
    if (convertBtn) convertBtn.classList.remove('hidden');
    if (downloadSection) downloadSection.classList.add('hidden');
    
    updateProgress(0);
    if (statusText) statusText.textContent = "Fichier chargé avec succès. Prêt pour l'action.";
}

// --- Contrôleur de Conversion ---
if (convertBtn) {
    convertBtn.addEventListener('click', async () => {
        if (!selectedFile) return;

        convertBtn.classList.add('hidden');
        if (statusText) statusText.textContent = "Analyse et traitement du document...";
        updateProgress(10);

        try {
            if (currentMode === 'word-to-pdf') {
                await convertWordToPdf(selectedFile);
            } else {
                await convertPdfToWord(selectedFile);
            }
        } catch (error) {
            console.error("Erreur détectée pendant la conversion :", error);
            if (statusText) statusText.textContent = "Erreur : " + error.message;
            convertBtn.classList.remove('hidden');
            if (progressBar) progressBar.style.backgroundColor = '#ef4444';
        }
    });
}

// LOGIQUE CORRIGÉE : Word -> PDF (Capture asynchrone sécurisée)
async function convertWordToPdf(file) {
    return new Promise((resolve, reject) => {
        if (typeof mammoth === 'undefined' || typeof html2pdf === 'undefined') {
            reject(new Error("Les moteurs de décodage ne sont pas encore prêts. Veuillez patienter une seconde."));
            return;
        }

        const reader = new FileReader();
        reader.onload = function(e) {
            const arrayBuffer = e.target.result;
            updateProgress(40);
            
            mammoth.convertToHtml({ arrayBuffer: arrayBuffer })
            .then(function(result) {
                const htmlContent = result.value;
                
                if (!htmlContent || !htmlContent.trim()) {
                    reject(new Error("Le fichier inséré ne contient aucun texte exploitable."));
                    return;
                }

                // Injection dynamique forcée dans une boîte isolée (Résout l'anomalie de la page blanche)
                const sandbox = document.createElement('div');
                sandbox.innerHTML = htmlContent;
                sandbox.style.padding = "40px";
                sandbox.style.width = "650px";
                sandbox.style.color = "#000000";
                sandbox.style.backgroundColor = "#ffffff";
                sandbox.style.fontFamily = "Arial, sans-serif";
                sandbox.style.position = "absolute";
                sandbox.style.left = "-9999px"; 
                document.body.appendChild(sandbox);

                updateProgress(70);

                const opt = {
                    margin:       15,
                    filename:     file.name.replace('.docx', '.pdf'),
                    image:        { type: 'jpeg', quality: 0.98 },
                    html2canvas:  { scale: 2, useCORS: true, logging: false },
                    jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
                };

                // Pipeline html2pdf direct sur le flux
                html2pdf().set(opt).from(sandbox).outputPdf('blob').then((blob) => {
                    convertedBlob = blob;
                    convertedFileName = file.name.replace('.docx', '.pdf');
                    document.body.removeChild(sandbox); // Libération immédiate de la mémoire dom
                    showDownload();
                    resolve();
                }).catch(err => {
                    if (document.body.contains(sandbox)) document.body.removeChild(sandbox);
                    reject(err);
                });
            })
            .catch(err => reject(new Error("Mammoth n'a pas pu traiter la structure : " + err.message)));
        };
        reader.onerror = () => reject(new Error("Erreur physique lors de l'accès au fichier."));
        reader.readAsArrayBuffer(file);
    });
}

// LOGIQUE CORRIGÉE : PDF -> Word (Extraction unifiée et encapsulation Blob MS-Word)
async function convertPdfToWord(file) {
    return new Promise((resolve, reject) => {
        if (typeof pdfjsLib === 'undefined') {
            reject(new Error("La bibliothèque PDF.js requise n'a pas pu être initialisée."));
            return;
        }

        const reader = new FileReader();
        reader.onload = async function(e) {
            try {
                const typedarray = new Uint8Array(e.target.result);
                const loadingTask = pdfjsLib.getDocument(typedarray);
                const pdf = await loadingTask.promise;
                let textAccumulator = "";

                updateProgress(30);
                
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    const pageText = textContent.items.map(item => item.str).join(' ');
                    textAccumulator += `<p style="margin-bottom: 14px; line-height: 1.5; font-size: 11pt;">${pageText}</p>\n`;
                    updateProgress(30 + Math.floor((i / pdf.numPages) * 50));
                }

                if (!textAccumulator.replace(/<[^>]*>/g, '').trim()) {
                    throw new Error("Aucune chaîne textuelle n'a pu être extraite (Le PDF provient probablement d'un scan ou d'une image).");
                }

                // Génération du conteneur HTML interprétable nativement par Microsoft Word
                const docContent = `
                    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
                    <head><meta charset="utf-8"></head>
                    <body style="font-family: Arial, sans-serif; padding: 40px;">
                        ${textAccumulator}
                    </body>
                    </html>
                `;

                convertedBlob = new Blob(['\ufeff' + docContent], { type: 'application/msword' });
                convertedFileName = file.name.replace('.pdf', '.doc'); // .doc assure une rétrocompatibilité complète sans plantage
                
                showDownload();
                resolve();
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = () => reject(new Error("Erreur lors de la lecture du flux PDF."));
        reader.readAsArrayBuffer(file);
    });
}

function updateProgress(percent) {
    if (progressBar) progressBar.style.width = percent + '%';
}

function showDownload() {
    updateProgress(100);
    if (statusText) statusText.textContent = "Opération terminée avec succès !";
    if (downloadSection) downloadSection.classList.remove('hidden');
}

if (downloadBtn) {
    downloadBtn.addEventListener('click', () => {
        if (convertedBlob) {
            if (typeof saveAs !== 'undefined') {
                saveAs(convertedBlob, convertedFileName);
            } else {
                // Solution Fallback natif standard
                const url = URL.createObjectURL(convertedBlob);
                const a = document.createElement('a');
                a.href = url;
                a.download = convertedFileName;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }
        }
    });
}

if (resetBtn) {
    resetBtn.addEventListener('click', () => resetUI());
}

function resetUI() {
    selectedFile = null;
    convertedBlob = null;
    convertedFileName = "";
    if (fileInput) fileInput.value = "";
    if (dropZone) dropZone.classList.remove('hidden');
    if (processSection) processSection.classList.add('hidden');
    if (downloadSection) downloadSection.classList.add('hidden');
    updateProgress(0);
    if (progressBar) progressBar.style.backgroundColor = '';
    updateUIForMode();
}
