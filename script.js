/**
 * CONVERTISSEUR UNIVERSEL - SCRIPT OPTIMISÉ
 * Architecture Client-Side pour GitHub Pages
 * Optimisations : Web Workers (via Tesseract), Blob Management, Proper DOCX Generation
 */

// --- 1. CONFIGURATION ET INITIALISATION ---
const pdfjsLib = window['pdfjs-dist/build/pdf'];
if (pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
}

// --- 2. ÉLÉMENTS DU DOM ---
const elements = {
    dropZone: document.getElementById('drop-zone'),
    fileInput: document.getElementById('file-input'),
    tabBtns: document.querySelectorAll('.tab-btn'),
    processSection: document.getElementById('process-section'),
    qrSection: document.getElementById('qr-section'),
    downloadSection: document.getElementById('download-section'),
    convertBtn: document.getElementById('convert-btn'),
    downloadBtn: document.getElementById('download-btn'),
    resetBtn: document.getElementById('reset-btn'),
    progressBar: document.getElementById('progress-bar'),
    statusText: document.getElementById('status-text'),
    fileNameDisplay: document.getElementById('file-name'),
    fileIcon: document.getElementById('file-icon'),
    fileInfoText: document.getElementById('file-info-text'),
    qrInput: document.getElementById('qr-input'),
    qrDisplay: document.getElementById('qr-display'),
    generateQrBtn: document.getElementById('generate-qr-btn')
};

// --- 3. ÉTAT DE L'APPLICATION ---
let state = {
    currentMode: 'word-to-pdf',
    selectedFile: null,
    convertedBlob: null,
    convertedFileName: "",
    qrCodeInstance: null
};

// --- 4. NAVIGATION ET ONGLETS ---
elements.tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        elements.tabBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.currentMode = btn.dataset.mode;
        updateUIForMode();
        resetUI();
    });
});

function updateUIForMode() {
    // Cache toutes les sections par défaut
    elements.dropZone.classList.add('hidden');
    elements.qrSection.classList.add('hidden');
    elements.processSection.classList.add('hidden');

    if (state.currentMode === 'qr-generator') {
        elements.qrSection.classList.remove('hidden');
    } else {
        elements.dropZone.classList.remove('hidden');
        if (state.currentMode === 'word-to-pdf') {
            elements.fileInfoText.textContent = "Fichiers acceptés : .docx";
            elements.fileIcon.className = "fas fa-file-word";
            elements.fileInput.accept = ".docx";
        } else {
            elements.fileInfoText.textContent = "Fichiers acceptés : .pdf";
            elements.fileIcon.className = "fas fa-file-pdf";
            elements.fileInput.accept = ".pdf";
        }
    }
}

// --- 5. MODULE QR CODE ---
elements.generateQrBtn.addEventListener('click', () => {
    const text = elements.qrInput.value.trim();
    if (!text) {
        alert("Veuillez entrer un texte ou une URL.");
        return;
    }

    // Nettoyage de l'affichage précédent
    elements.qrDisplay.innerHTML = "";
    
    // Création du QR Code
    state.qrCodeInstance = new QRCode(elements.qrDisplay, {
        text: text,
        width: 256,
        height: 256,
        colorDark: "#1e293b",
        colorLight: "#ffffff",
        correctLevel: QRCode.CorrectLevel.H
    });

    // Transformation du bouton en bouton de téléchargement après génération
    elements.generateQrBtn.textContent = "Télécharger le QR Code (.png)";
    elements.generateQrBtn.onclick = downloadQRCode;
});

function downloadQRCode() {
    const img = elements.qrDisplay.querySelector('img');
    if (img) {
        const link = document.createElement('a');
        link.href = img.src;
        link.download = "qrcode-universel.png";
        link.click();
    }
}

// --- 6. GESTION DES FICHIERS (DRAG & DROP) ---
elements.dropZone.addEventListener('click', () => elements.fileInput.click());
elements.dropZone.addEventListener('dragover', (e) => { e.preventDefault(); elements.dropZone.classList.add('drag-over'); });
elements.dropZone.addEventListener('dragleave', () => elements.dropZone.classList.remove('drag-over'));
elements.dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    elements.dropZone.classList.remove('drag-over');
    handleFiles(e.dataTransfer.files);
});

elements.fileInput.addEventListener('change', (e) => handleFiles(e.target.files));

function handleFiles(files) {
    if (files.length === 0) return;
    const file = files[0];
    const extension = file.name.split('.').pop().toLowerCase();

    if (state.currentMode === 'word-to-pdf' && extension !== 'docx') {
        alert("Format invalide. Utilisez .docx");
        return;
    }
    if (state.currentMode === 'pdf-to-word' && extension !== 'pdf') {
        alert("Format invalide. Utilisez .pdf");
        return;
    }

    state.selectedFile = file;
    elements.fileNameDisplay.textContent = file.name;
    elements.dropZone.classList.add('hidden');
    elements.processSection.classList.remove('hidden');
    elements.convertBtn.classList.remove('hidden');
    elements.downloadSection.classList.add('hidden');
    updateProgress(0);
    elements.statusText.textContent = "Fichier prêt pour la conversion.";
}

// --- 7. MOTEUR DE CONVERSION : WORD ➡️ PDF ---
// Stratégie : Injection de styles spécifiques pour préserver la structure des images et logos
async function convertWordToPdf(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async function(e) {
            try {
                const arrayBuffer = e.target.result;
                updateProgress(20);
                
                // Mammoth convertit le Word en HTML
                const options = {
                    styleMap: [
                        "p[style-name='Title'] => h1:fresh",
                        "p[style-name='Subtitle'] => h2:fresh"
                    ]
                };
                
                const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer }, options);
                let html = result.value;
                
                // --- OPTIMISATION STRUCTURELLE ---
                // On enveloppe le HTML dans un conteneur avec des styles de préservation
                const container = document.createElement('div');
                container.className = "pdf-render-container";
                container.style.cssText = `
                    padding: 40px;
                    font-family: 'Arial', sans-serif;
                    background: white;
                    color: black;
                    line-height: 1.5;
                `;
                
                // Injection de styles CSS correctifs pour les images et l'alignement
                const styleInjection = `
                    <style>
                        img { max-width: 100% !important; height: auto !important; display: block; margin: 10px auto; }
                        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
                        td, th { border: 1px solid #ddd; padding: 8px; }
                        h1, h2, h3 { color: #2c3e50; margin-top: 20px; }
                        p { margin-bottom: 12px; text-align: justify; }
                    </style>
                `;
                
                container.innerHTML = styleInjection + html;
                updateProgress(50);

                // Configuration de html2pdf pour une capture haute fidélité
                const opt = {
                    margin: 10,
                    filename: file.name.replace('.docx', '.pdf'),
                    image: { type: 'jpeg', quality: 0.98 },
                    html2canvas: { 
                        scale: 2, 
                        useCORS: true,
                        logging: false,
                        letterRendering: true
                    },
                    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
                };

                html2pdf().from(container).set(opt).outputPdf('blob').then(blob => {
                    state.convertedBlob = blob;
                    state.convertedFileName = file.name.replace('.docx', '.pdf');
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

// --- 8. MOTEUR DE CONVERSION : PDF ➡️ WORD ---
// Stratégie : Reconstruction de la structure via docx.js (True DOCX)
async function convertPdfToWord(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async function(e) {
            try {
                const docxLib = window.docx || window.DOCX;
                if (!docxLib) throw new Error("Bibliothèque de génération Word non chargée.");

                const typedarray = new Uint8Array(e.target.result);
                const pdf = await pdfjsLib.getDocument(typedarray).promise;
                let sections = [];

                updateProgress(15);

                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    
                    // --- ANALYSE DE LA STRUCTURE SPATIALE ---
                    // On trie les éléments par position Y (ligne) puis X (colonne)
                    const items = textContent.items.sort((a, b) => {
                        if (Math.abs(a.transform[5] - b.transform[5]) < 5) {
                            return a.transform[4] - b.transform[4];
                        }
                        return b.transform[5] - a.transform[5];
                    });

                    let lastY = -1;
                    let paragraphs = [];
                    let currentLine = "";

                    items.forEach(item => {
                        if (lastY !== -1 && Math.abs(item.transform[5] - lastY) > 5) {
                            // Nouvelle ligne détectée
                            paragraphs.push(new docxLib.Paragraph({
                                children: [new docxLib.TextRun({ text: currentLine.trim(), size: 24 })],
                                spacing: { after: 200 }
                            }));
                            currentLine = "";
                        }
                        currentLine += item.str + " ";
                        lastY = item.transform[5];
                    });

                    // Ajout de la dernière ligne
                    if (currentLine) {
                        paragraphs.push(new docxLib.Paragraph({
                            children: [new docxLib.TextRun({ text: currentLine.trim(), size: 24 })]
                        }));
                    }

                    sections.push({
                        properties: {},
                        children: paragraphs
                    });

                    updateProgress(15 + (i / pdf.numPages) * 70);
                }

                const doc = new docxLib.Document({ sections: sections });
                const blob = await docxLib.Packer.toBlob(doc);
                
                state.convertedBlob = blob;
                state.convertedFileName = file.name.replace('.pdf', '.docx');
                showDownload();
                resolve();
            } catch (err) {
                reject(err);
            }
        };
        reader.readAsArrayBuffer(file);
    });
}

// --- 9. ACTIONS ET UI ---
elements.convertBtn.addEventListener('click', async () => {
    if (!state.selectedFile) return;
    elements.convertBtn.classList.add('hidden');
    elements.statusText.textContent = "Conversion en cours (Traitement local)...";
    
    try {
        if (state.currentMode === 'word-to-pdf') {
            await convertWordToPdf(state.selectedFile);
        } else {
            await convertPdfToWord(state.selectedFile);
        }
    } catch (error) {
        console.error(error);
        elements.statusText.textContent = "Erreur : " + error.message;
        elements.convertBtn.classList.remove('hidden');
        elements.progressBar.style.backgroundColor = '#f43f5e';
    }
});

function updateProgress(percent) {
    elements.progressBar.style.width = percent + '%';
}

function showDownload() {
    updateProgress(100);
    elements.statusText.textContent = "Terminé ! Votre fichier est prêt.";
    elements.downloadSection.classList.remove('hidden');
}

elements.downloadBtn.addEventListener('click', () => {
    if (state.convertedBlob) {
        saveAs(state.convertedBlob, state.convertedFileName);
    }
});

elements.resetBtn.addEventListener('click', () => resetUI());

function resetUI() {
    state.selectedFile = null;
    state.convertedBlob = null;
    elements.fileInput.value = "";
    elements.qrInput.value = "";
    elements.qrDisplay.innerHTML = "";
    elements.generateQrBtn.textContent = "Générer le QR Code";
    elements.generateQrBtn.onclick = null; // Reset click handler
    
    updateUIForMode();
    elements.downloadSection.classList.add('hidden');
    updateProgress(0);
    elements.progressBar.style.backgroundColor = '';
}

// Initialisation au chargement
updateUIForMode();
