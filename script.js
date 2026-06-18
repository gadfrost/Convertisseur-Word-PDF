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
    btnQr: document.getElementById('btn-qr-generator'),
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
// Gestion robuste des onglets
elements.tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        console.log("Changement de mode vers :", btn.dataset.mode);
        elements.tabBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        state.currentMode = btn.dataset.mode;
        updateUIForMode();
        // Ne pas appeler resetUI ici pour éviter de vider le QR Code généré si on change d'onglet
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
    const canvas = elements.qrDisplay.querySelector('canvas');
    
    let dataUrl;
    if (img && img.src) {
        dataUrl = img.src;
    } else if (canvas) {
        dataUrl = canvas.toDataURL("image/png");
    }

    if (dataUrl) {
        const link = document.createElement('a');
        link.href = dataUrl;
        link.download = "qrcode-universel.png";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    } else {
        alert("Erreur lors de la préparation du téléchargement.");
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
                
                // Conversion Word -> HTML
                const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
                let html = result.value;
                
                // Création d'un conteneur de rendu temporaire
                const renderDiv = document.createElement('div');
                renderDiv.id = 'pdf-render-temp';
                renderDiv.style.position = 'absolute';
                renderDiv.style.left = '-9999px';
                renderDiv.style.top = '0';
                renderDiv.style.width = '794px'; // Largeur A4 standard (210mm @ 96dpi)
                renderDiv.style.backgroundColor = 'white';
                
                // Styles CSS pour garantir la fidélité et le multi-pages
                const styles = `
                    <style>
                        #pdf-render-temp {
                            padding: 50px;
                            color: #000;
                            font-family: Arial, sans-serif;
                            line-height: 1.5;
                        }
                        #pdf-render-temp img {
                            max-width: 100%;
                            height: auto;
                            display: block;
                            margin: 15px auto;
                        }
                        #pdf-render-temp table {
                            width: 100%;
                            border-collapse: collapse;
                            margin-bottom: 20px;
                        }
                        #pdf-render-temp td, #pdf-render-temp th {
                            border: 1px solid #ccc;
                            padding: 8px;
                        }
                        #pdf-render-temp h1, #pdf-render-temp h2, #pdf-render-temp h3 {
                            color: #333;
                            page-break-after: avoid;
                        }
                        #pdf-render-temp p {
                            margin-bottom: 12px;
                            text-align: justify;
                            orphans: 3;
                            widows: 3;
                        }
                    </style>
                `;
                
                renderDiv.innerHTML = styles + html;
                document.body.appendChild(renderDiv);
                
                updateProgress(50);

                // Attendre que les images soient chargées pour éviter les trous blancs
                const images = renderDiv.getElementsByTagName('img');
                const imagePromises = Array.from(images).map(img => {
                    if (img.complete) return Promise.resolve();
                    return new Promise(res => { img.onload = res; img.onerror = res; });
                });
                await Promise.all(imagePromises);

                // Configuration html2pdf
                const opt = {
                    margin: 10,
                    filename: file.name.replace('.docx', '.pdf'),
                    image: { type: 'jpeg', quality: 0.98 },
                    html2canvas: { 
                        scale: 2, 
                        useCORS: true,
                        letterRendering: true,
                        scrollY: 0
                    },
                    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
                    pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
                };

                // Lancer la capture
                console.log("Démarrage de la capture PDF...");
                html2pdf().from(renderDiv).set(opt).toPdf().get('pdf').then(function (pdf) {
                    console.log("Capture réussie, génération du blob...");
                    const blob = pdf.output('blob');
                    
                    if (blob.size < 1000) {
                        console.warn("Attention : Le PDF généré semble très petit ou vide.");
                    }
                    
                    state.convertedBlob = blob;
                    state.convertedFileName = file.name.replace('.docx', '.pdf');
                    
                    // Nettoyage
                    document.body.removeChild(renderDiv);
                    
                    showDownload();
                    resolve();
                }).catch(err => {
                    console.error("Erreur html2pdf :", err);
                    if (document.getElementById('pdf-render-temp')) {
                        document.body.removeChild(renderDiv);
                    }
                    reject(err);
                });
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
