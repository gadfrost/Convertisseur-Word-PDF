// --- 1. CONFIGURATION INITIALE & DEPLOYEMENT DES WORKERS ---
const pdfjsLib = window['pdfjs-dist/build/pdf'];
if (pdfjsLib) {
    // Force le chemin du worker pour éviter les échecs sur mobile et GitHub Pages
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
}

// --- 2. CAPTURE DES ÉLÉMENTS DU DOM ---
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

// --- 3. ÉTAT DE L'APPLICATION ---
let currentMode = 'word-to-pdf'; 
let selectedFile = null;
let convertedBlob = null;
let convertedFileName = "";
let deferredPrompt;

// --- 4. ENREGISTREMENT DU SERVICE WORKER (PWA) ---
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('./sw.js')
            .then(reg => console.log('Service Worker enregistré !'))
            .catch(err => console.error('Erreur SW:', err));
    });
}

// --- 5. LOGIQUE D'INSTALLATION PWA (BOUTON TOUJOURS VISIBLE) ---
if (installBtn) {
    installBtn.classList.remove('hidden');
}

window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
});

if (installBtn) {
    installBtn.addEventListener('click', async () => {
        if (deferredPrompt) {
            deferredPrompt.prompt();
            const { outcome } = await deferredPrompt.userChoice;
            console.log(`Statut d'installation : ${outcome}`);
            deferredPrompt = null;
        } else {
            alert("L'installation est directement gérée par votre navigateur. Cliquez sur les 3 petits points en haut à droite de Chrome, puis sur 'Ajouter à l'écran d'accueil'.");
        }
    });
}

// --- 6. GESTION DES ONGLETS (MODES) ---
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
        if (fileInput) fileInput.accept = ".docx";
    } else {
        if (fileInfoText) fileInfoText.textContent = "Fichiers acceptés : .pdf";
        if (fileIcon) fileIcon.className = "fas fa-file-pdf";
        if (fileInput) fileInput.accept = ".pdf";
    }
}

// --- 7. GESTION DU DRAG & DROP & CLIC DE ZONE ---
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

// --- 8. ÉCOUTEUR DU BOUTON DE CONVERSION ---
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

// --- 9. FONCTION DE CONVERSION : WORD -> PDF (RENDU VISUEL AMÉLIORÉ ET SÉCURISÉ) ---
async function convertWordToPdf(file) {
    return new Promise((resolve, reject) => {
        if (typeof mammoth === 'undefined' || typeof html2pdf === 'undefined') {
            reject(new Error("Les moteurs de décodage ne sont pas prêts. Veuillez actualiser la page."));
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

                const sandbox = document.createElement('div');
                sandbox.id = "pdf-sandbox";
                sandbox.innerHTML = `<div class="word-content">${htmlContent}</div>`;
                
                // Style simulant une feuille A4 visible pour forcer Chrome Mobile à calculer les géométries textuelles
                sandbox.style.position = "relative";
                sandbox.style.width = "100%";
                sandbox.style.maxWidth = "800px";
                sandbox.style.margin = "40px auto";
                sandbox.style.padding = "50px";
                sandbox.style.backgroundColor = "#ffffff";
                sandbox.style.color = "#000000";
                sandbox.style.fontFamily = "'Times New Roman', Times, serif";
                sandbox.style.fontSize = "14px";
                sandbox.style.lineHeight = "1.6";
                sandbox.style.textAlign = "justify";
                
                // Styles injectés pour redimensionner proprement les logos et restructurer les blocs de la page de garde
                const styleTag = document.createElement('style');
                styleTag.innerHTML = `
                    #pdf-sandbox img { max-width: 180px !important; height: auto !important; display: block; margin: 15px auto !important; }
                    #pdf-sandbox p { margin-bottom: 12px; }
                    #pdf-sandbox h1, #pdf-sandbox h2, #pdf-sandbox h3 { text-align: center; margin-top: 15px; margin-bottom: 15px; font-weight: bold; }
                `;
                sandbox.appendChild(styleTag);
                
                document.body.appendChild(sandbox);
                updateProgress(60);
                if (statusText) statusText.textContent = "Génération de la mise en page graphique...";

                // Temporisation d'attente de rendu
                setTimeout(() => {
                    const opt = {
                        margin:       15,
                        filename:     file.name.replace('.docx', '.pdf'),
                        image:        { type: 'jpeg', quality: 0.98 },
                        html2canvas:  { scale: 2, useCORS: true, logging: false },
                        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
                    };

                    html2pdf().set(opt).from(sandbox).outputPdf('blob').then((blob) => {
                        convertedBlob = blob;
                        convertedFileName = file.name.replace('.docx', '.pdf');
                        document.body.removeChild(sandbox); 
                        showDownload();
                        resolve();
                    }).catch(err => {
                        if (document.body.contains(sandbox)) document.body.removeChild(sandbox);
                        reject(err);
                    });
                }, 400);
            })
            .catch(err => reject(new Error("Erreur de décodage du Word : " + err.message)));
        };
        reader.onerror = () => reject(new Error("Erreur physique de lecture."));
        reader.readAsArrayBuffer(file);
    });
}

// --- 10. FONCTION DE CONVERSION : PDF -> WORD (RECONSTRUCTION STRATEGIQUE MULTI-PAGE ET COMPATIBLE WORD MOBILE) ---
async function convertPdfToWord(file) {
    return new Promise((resolve, reject) => {
        if (typeof pdfjsLib === 'undefined') {
            reject(new Error("La bibliothèque fondamentale PDF.js n'est pas chargée."));
            return;
        }

        const reader = new FileReader();
        reader.onload = async function(e) {
            try {
                const typedarray = new Uint8Array(e.target.result);
                const loadingTask = pdfjsLib.getDocument(typedarray);
                const pdf = await loadingTask.promise;
                let htmlAccumulator = "";
                let isScannedPdf = false;

                updateProgress(20);
                if (statusText) statusText.textContent = "Analyse de l'alignement des paragraphes...";
                
                // Parcours pour extraction du texte natif structuré
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    
                    let lastY = null;
                    let pageText = "";
                    
                    // Reconstruction fine avec détection des coupures de lignes d'origine
                    for (let item of textContent.items) {
                        if (lastY !== null && Math.abs(item.transform[5] - lastY) > 5) {
                            pageText += "<br/>";
                        }
                        pageText += item.str + " ";
                        lastY = item.transform[5];
                    }

                    if (pageText.trim()) {
                        htmlAccumulator += `<div class="page" style="page-break-after:always; font-family:Arial, sans-serif; font-size:11pt; line-height:1.5; margin-bottom:30px;">${pageText}</div>`;
                    }
                    updateProgress(20 + Math.floor((i / pdf.numPages) * 20));
                }

                // Mode OCR Tesseract v5 si le PDF ne contient aucune couche texte native
                if (!htmlAccumulator.replace(/<[^>]*>/g, '').trim()) {
                    if (typeof Tesseract === 'undefined') {
                        throw new Error("Le moteur OCR (Tesseract) n'est pas accessible.");
                    }
                    
                    isScannedPdf = true;
                    if (statusText) statusText.textContent = "Document image identifié. Lancement de la reconnaissance optique...";
                    updateProgress(45);
                    
                    const worker = await Tesseract.createWorker('fra');
                    
                    for (let i = 1; i <= pdf.numPages; i++) {
                        if (statusText) statusText.textContent = `Traitement de la page ${i}/${pdf.numPages}...`;
                        
                        const page = await pdf.getPage(i);
                        const viewport = page.getViewport({ scale: 1.5 }); 
                        const canvas = document.createElement('canvas');
                        const context = canvas.getContext('2d');
                        canvas.height = viewport.height;
                        canvas.width = viewport.width;

                        await page.render({ canvasContext: context, viewport: viewport }).promise;
                        
                        const { data: { text } } = await worker.recognize(canvas);
                        const formattedText = text.split('\n')
                                                  .map(line => line.trim() ? `<p style="margin:0 0 6px 0;">${line}</p>` : '<br/>')
                                                  .join('');
                                                  
                        htmlAccumulator += `<div class="page" style="page-break-after:always;">${formattedText}</div>`;
                        updateProgress(45 + Math.floor((i / pdf.numPages) * 45));
                    }
                    await worker.terminate();
                }

                if (!htmlAccumulator.replace(/<[^>]*>/g, '').trim()) {
                    throw new Error("Aucune chaîne textuelle n'a pu être interceptée.");
                }

                if (statusText) statusText.textContent = "Génération de l'en-tête natif Microsoft Office...";

                // Utilisation de l'espace de nommage Office XML complet pour forcer l'application Word Mobile à décoder et accepter le document
                const docContent = `
                    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
                    <head>
                        <meta charset="utf-8">
                        <style>
                            body { font-family: 'Arial', sans-serif; padding: 50px; color: #000000; background-color: #ffffff; }
                            p { margin: 0 0 10px 0; line-height: 1.4; }
                            .page { page-break-after: always; }
                        </style>
                    </head>
                    <body>
                        ${htmlAccumulator}
                    </body>
                    </html>
                `;

                // Création du flux binaire avec le BOM UTF-8 (\ufeff) pour éviter les caractères corrompus
                convertedBlob = new Blob(['\ufeff' + docContent], { type: 'application/msword;charset=utf-8' });
                convertedFileName = file.name.replace('.pdf', isScannedPdf ? '_scan.doc' : '.doc');
                
                showDownload();
                resolve();
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = () => reject(new Error("Erreur physique de lecture."));
        reader.readAsArrayBuffer(file);
    });
}

// --- 11. INTERFACES GRAPHIQUES DE RETOUR (PROGRESSION / RESET) ---
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
