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
            .then(reg => console.log('Service Worker enregistré !'))
            .catch(err => console.log('Erreur Service Worker:', err));
    });
}

// --- PWA Installation Logic ---
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
        if (fileInfoText) fileInfoText.textContent = "Fichiers acceptés : .docx";
        if (fileIcon) fileIcon.className = "fas fa-file-word";
    } else {
        if (fileInfoText) fileInfoText.textContent = "Fichiers acceptés : .pdf";
        if (fileIcon) fileIcon.className = "fas fa-file-pdf";
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
        progressBar.style.backgroundColor = 'var(--error-color, #ef4444)';
    }
});

// REPARATION : Word -> PDF (Plus de page blanche)
async function convertWordToPdf(file) {
    return new Promise((resolve, reject) => {
        if (typeof mammoth === 'undefined' || typeof html2pdf === 'undefined') {
            reject(new Error("Les bibliothèques de conversion ne sont pas encore prêtes. Réessayez dans une seconde."));
            return;
        }

        const reader = new FileReader();
        reader.onload = function(e) {
            const arrayBuffer = e.target.result;
            updateProgress(40);
            
            mammoth.convertToHtml({ arrayBuffer: arrayBuffer })
            .then(function(result) {
                const htmlContent = result.value;
                
                if (!htmlContent.trim()) {
                    reject(new Error("Le fichier Word semble vide ou illisible."));
                    return;
                }

                // Création d'un conteneur visible temporairement pour forcer html2pdf à voir le texte
                const element = document.createElement('div');
                element.innerHTML = htmlContent;
                element.style.padding = "30px";
                element.style.width = "600px";
                element.style.color = "#000000";
                element.style.backgroundColor = "#ffffff";
                element.style.fontFamily = "Arial, sans-serif";
                element.style.position = "absolute";
                element.style.left = "-9999px"; // Caché hors de l'écran mais lisible par le moteur de rendu
                document.body.appendChild(element);

                updateProgress(70);

                const opt = {
                    margin:       15,
                    filename:     file.name.replace('.docx', '.pdf'),
                    image:        { type: 'jpeg', quality: 0.98 },
                    html2canvas:  { scale: 2, useCORS: true, logging: false },
                    jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
                };

                // html2pdf direct sur le flux pour éviter la rupture de promesse
                html2pdf().set(opt).from(element).outputPdf('blob').then((blob) => {
                    convertedBlob = blob;
                    convertedFileName = file.name.replace('.docx', '.pdf');
                    document.body.removeChild(element);
                    showDownload();
                    resolve();
                }).catch(err => {
                    if(document.body.contains(element)) document.body.removeChild(element);
                    reject(err);
                });
            })
            .catch(err => reject(new Error("Erreur lors du décodage du Word : " + err.message)));
        };
        reader.onerror = () => reject(new Error("Erreur physique de lecture du fichier."));
        reader.readAsArrayBuffer(file);
    });
}

// REPARATION : PDF -> Word (Extraction universelle en Blob MS-Word sans plantage de bibliothèque)
async function convertPdfToWord(file) {
    return new Promise((resolve, reject) => {
        if (typeof pdfjsLib === 'undefined') {
            reject(new Error("La bibliothèque PDF.js n'est pas chargée."));
            return;
        }

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
                    fullText += `<p style="margin-bottom: 15px; line-height: 1.6;">${pageText}</p>\n`;
                    updateProgress(30 + (i / pdf.numPages) * 50);
                }

                if (!fullText.replace(/<[^>]*>/g, '').trim()) {
                    throw new Error("Aucun texte extractible trouvé (Le PDF est peut-être une image scannée).");
                }

                // Génération d'un fichier Word structuré au format HTML-DOC, lisible nativement par Microsoft Word
                const docContent = `
                    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
                    <head><meta charset="utf-8"><title>Converti</title></head>
                    <body style="font-family: Arial, sans-serif; padding: 40px;">
                        ${fullText}
                    </body>
                    </html>
                `;

                convertedBlob = new Blob(['\ufeff' + docContent], { type: 'application/msword' });
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
    if (progressBar) progressBar.style.width = percent + '%';
}

function showDownload() {
    updateProgress(100);
    statusText.textContent = "Terminé !";
    downloadSection.classList.remove('hidden');
}

downloadBtn.addEventListener('click', () => {
    if (convertedBlob) {
        if (typeof saveAs !== 'undefined') {
            saveAs(convertedBlob, convertedFileName);
        } else {
            // Solution de secours native si FileSaver échoue
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

resetBtn.addEventListener('click', () => resetUI());

function resetUI() {
    selectedFile = null;
    convertedBlob = null;
    convertedFileName = "";
    dropZone.classList.remove('hidden');
    processSection.classList.add('hidden');
    downloadSection.classList.add('hidden');
    updateProgress(0);
    if (progressBar) progressBar.style.backgroundColor = '';
    fileInput.value = "";
    updateUIForMode();
}
