// --- 9. FONCTION DE CONVERSION : WORD -> PDF (CORRIGÉE POUR MOBILE) ---
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

                // Correctif Page Blanche Mobile : Le conteneur doit avoir une taille réelle hors de l'écran
                const sandbox = document.createElement('div');
                sandbox.innerHTML = htmlContent;
                
                sandbox.style.position = "fixed";
                sandbox.style.left = "0";
                sandbox.style.top = "-9999px"; 
                sandbox.style.width = "794px"; /* Largeur standard A4 en pixels à 96 DPI */
                sandbox.style.padding = "40px";
                sandbox.style.backgroundColor = "#ffffff";
                sandbox.style.color = "#000000";
                sandbox.style.fontFamily = "Arial, sans-serif";
                
                document.body.appendChild(sandbox);
                updateProgress(70);

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
            })
            .catch(err => reject(new Error("Erreur de structure Mammoth : " + err.message)));
        };
        reader.onerror = () => reject(new Error("Erreur physique de lecture."));
        reader.readAsArrayBuffer(file);
    });
}

// --- 10. FONCTION DE CONVERSION : PDF -> WORD (CORRIGÉE POUR TESSERACT v5) ---
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
                let textAccumulator = "";
                let isScannedPdf = false;

                updateProgress(20);
                if (statusText) statusText.textContent = "Analyse de la structure textuelle...";
                
                // Tentative 1 : Lecture du texte natif
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    const pageText = textContent.items.map(item => item.str).join(' ');
                    if (pageText.trim()) {
                        textAccumulator += `<p style="margin-bottom: 14px; line-height: 1.5; font-family: Arial, sans-serif;">${pageText}</p>\n`;
                    }
                    updateProgress(20 + Math.floor((i / pdf.numPages) * 20));
                }

                // Tentative 2 : Si vide, utilisation stricte de l'API de Tesseract v5 Web Worker
                if (!textAccumulator.replace(/<[^>]*>/g, '').trim()) {
                    if (typeof Tesseract === 'undefined') {
                        throw new Error("Le moteur OCR (Tesseract) n'est pas accessible.");
                    }
                    
                    isScannedPdf = true;
                    if (statusText) statusText.textContent = "PDF Image détecté. Initialisation du scanner optique...";
                    updateProgress(45);
                    
                    // Syntaxe d'initialisation obligatoire pour la v5 en environnement isolé
                    const worker = await Tesseract.createWorker('fra');
                    
                    for (let i = 1; i <= pdf.numPages; i++) {
                        if (statusText) statusText.textContent = `Numérisation de la page ${i} sur ${pdf.numPages}...`;
                        
                        const page = await pdf.getPage(i);
                        const viewport = page.getViewport({ scale: 2.0 }); // Résolution augmentée à 2.0 pour plus de précision
                        
                        const canvas = document.createElement('canvas');
                        const context = canvas.getContext('2d');
                        canvas.height = viewport.height;
                        canvas.width = viewport.width;

                        await page.render({ canvasContext: context, viewport: viewport }).promise;
                        
                        // Exécution via le worker instancié
                        const { data: { text } } = await worker.recognize(canvas);
                        
                        const formattedText = text.split('\n')
                                                  .map(line => line.trim() ? `<p style="margin-bottom: 12px; line-height: 1.5; font-family: Arial, sans-serif;">${line}</p>` : '')
                                                  .join('\n');
                                                  
                        textAccumulator += `\n${formattedText}\n`;
                        
                        updateProgress(45 + Math.floor((i / pdf.numPages) * 45));
                    }
                    // Libération immédiate de la mémoire système du téléphone
                    await worker.terminate();
                }

                // Double vérification de sécurité du tampon de texte
                if (!textAccumulator.replace(/<[^>]*>/g, '').trim()) {
                    throw new Error("L'analyse optique n'a extrait aucune donnée intelligible.");
                }

                if (statusText) statusText.textContent = "Création du document Word final...";

                // Encodage strict au format Microsoft Word HTML standard (Document complet et lourd)
                const docContent = `
                    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
                    <head>
                        <meta charset="utf-8">
                    </head>
                    <body style="padding: 90px; color: #000000; background-color: #ffffff;">
                        ${textAccumulator}
                    </body>
                    </html>
                `;

                // Ajout du BOM UTF-8 (\ufeff) pour éviter les corruptions de caractères à l'ouverture
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
