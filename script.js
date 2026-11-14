document.addEventListener('DOMContentLoaded', () => {
    const uploadZone = document.getElementById('upload-zone');
    const fileInput = document.getElementById('file-input');
    const previewSection = document.getElementById('preview-section');
    const imageGrid = document.getElementById('image-grid');
    const imageCount = document.getElementById('image-count');
    const clearAllBtn = document.getElementById('clear-all-btn');
    const addMoreImagesBtn = document.getElementById('add-more-images-btn');
    const generatePdfBtn = document.getElementById('generate-pdf-btn');
    const compressDownloadBtn = document.getElementById('compress-download-btn');
    const loadingOverlay = document.getElementById('loading-overlay');

    // Tab Elements
    const tabBtns = document.querySelectorAll('.tab-btn');
    const imageToolsPanel = document.getElementById('image-tools-panel');
    const pdfToolsPanel = document.getElementById('pdf-tools-panel');
    const compressorSettings = document.getElementById('compressor-settings');
    const pdfActions = document.getElementById('pdf-actions');
    const compressActions = document.getElementById('compress-actions');

    // PDF Tools Elements
    const pdfUploadZone = document.getElementById('pdf-upload-zone');
    const pdfFileInput = document.getElementById('pdf-file-input');
    const pdfPreviewSection = document.getElementById('pdf-preview-section');
    const pdfFilename = document.getElementById('pdf-filename');
    const pdfOriginalSize = document.getElementById('pdf-original-size');
    const pdfClearBtn = document.getElementById('pdf-clear-btn');
    const compressPdfBtn = document.getElementById('compress-pdf-btn');
    const pdfTargetKbInput = document.getElementById('pdf-target-kb-input');

    // Word Tools Elements
    const wordToolsPanel = document.getElementById('word-tools-panel');
    const wordUploadZone = document.getElementById('word-upload-zone');
    const wordFileInput = document.getElementById('word-file-input');
    const wordPreviewSection = document.getElementById('word-preview-section');
    const wordFilename = document.getElementById('word-filename');
    const wordOriginalSize = document.getElementById('word-original-size');
    const wordClearBtn = document.getElementById('word-clear-btn');
    const convertWordBtn = document.getElementById('convert-word-btn');
    const wordRenderContainer = document.getElementById('word-render-container');

    // PDF to Word Tools Elements
    const pdfWordToolsPanel = document.getElementById('pdf-word-tools-panel');
    const pdfWordUploadZone = document.getElementById('pdf-word-upload-zone');
    const pdfWordFileInput = document.getElementById('pdf-word-file-input');
    const pdfWordPreviewSection = document.getElementById('pdf-word-preview-section');
    const pdfWordFilename = document.getElementById('pdf-word-filename');
    const pdfWordOriginalSize = document.getElementById('pdf-word-original-size');
    const pdfWordClearBtn = document.getElementById('pdf-word-clear-btn');
    const convertPdfWordBtn = document.getElementById('convert-pdf-word-btn');

    let uploadedWordFile = null;
    let uploadedPdfWordFile = null;
    let uploadedPdfWordDataUrl = null;

    // PDF Translator Elements
    const pdfTranslateToolsPanel = document.getElementById('pdf-translate-tools-panel');
    const pdfTranslateUploadZone = document.getElementById('pdf-translate-upload-zone');
    const pdfTranslateFileInput = document.getElementById('pdf-translate-file-input');
    const pdfTranslatePreviewSection = document.getElementById('pdf-translate-preview-section');
    const pdfTranslateFilename = document.getElementById('pdf-translate-filename');
    const pdfTranslateOriginalSize = document.getElementById('pdf-translate-original-size');
    const pdfTranslateClearBtn = document.getElementById('pdf-translate-clear-btn');
    const convertPdfTranslateBtn = document.getElementById('convert-pdf-translate-btn');
    const pdfTranslateLang = document.getElementById('pdf-translate-lang');

    let uploadedPdfTranslateFile = null;
    let uploadedPdfTranslateDataUrl = null;

    // Text to PDF Elements
    const textPdfToolsPanel = document.getElementById('text-pdf-tools-panel');
    const convertTextPdfBtn = document.getElementById('convert-text-pdf-btn');

    let uploadedPdfFile = null;
    let uploadedPdfDataUrl = null;

    // Merge PDF Elements
    const mergePdfToolsPanel = document.getElementById('merge-pdf-tools-panel');
    const mergePdfUploadZone = document.getElementById('merge-pdf-upload-zone');
    const mergePdfFileInput = document.getElementById('merge-pdf-file-input');
    const mergePdfPreviewSection = document.getElementById('merge-pdf-preview-section');
    const mergePdfList = document.getElementById('merge-pdf-list');
    const mergePdfCount = document.getElementById('merge-pdf-count');
    const mergePdfClearBtn = document.getElementById('merge-pdf-clear-btn');
    const mergePdfAddMoreBtn = document.getElementById('merge-pdf-add-more-btn');
    const mergePdfBtn = document.getElementById('merge-pdf-btn');
    let mergePdfFiles = []; // Array of files

    // Split PDF Elements
    const splitPdfToolsPanel = document.getElementById('split-pdf-tools-panel');
    const splitPdfUploadZone = document.getElementById('split-pdf-upload-zone');
    const splitPdfFileInput = document.getElementById('split-pdf-file-input');
    const splitPdfPreviewSection = document.getElementById('split-pdf-preview-section');
    const splitPdfFilename = document.getElementById('split-pdf-filename');
    const splitPdfPagesCount = document.getElementById('split-pdf-pages-count');
    const splitPdfRange = document.getElementById('split-pdf-range');
    const splitPdfClearBtn = document.getElementById('split-pdf-clear-btn');
    const splitPdfBtn = document.getElementById('split-pdf-btn');
    let splitPdfFile = null;
    let splitPdfArrayBuffer = null;

    // Watermark PDF Elements
    const watermarkPdfToolsPanel = document.getElementById('watermark-pdf-tools-panel');
    const watermarkPdfUploadZone = document.getElementById('watermark-pdf-upload-zone');
    const watermarkPdfFileInput = document.getElementById('watermark-pdf-file-input');
    const watermarkPdfPreviewSection = document.getElementById('watermark-pdf-preview-section');
    const watermarkPdfFilename = document.getElementById('watermark-pdf-filename');
    const watermarkPdfClearBtn = document.getElementById('watermark-pdf-clear-btn');
    const watermarkPdfBtn = document.getElementById('watermark-pdf-btn');
    let watermarkPdfFile = null;
    let watermarkPdfArrayBuffer = null;

    let selectedFiles = []; // { id, file, dataUrl }
    let currentMode = 'pdf-mode'; // 'pdf-mode' or 'compress-mode'

    // Initialize SortableJS
    let sortable = new Sortable(imageGrid, {
        animation: 150,
        ghostClass: 'sortable-ghost',
        onEnd: () => {
            updateNumbers();
            reorderFilesArray();
        }
    });

    // --- Tab Logic ---
    const allPanels = [
        'image-tools-panel', 'pdf-tools-panel', 'word-tools-panel',
        'pdf-word-tools-panel', 'pdf-translate-tools-panel', 'text-pdf-tools-panel',
        'image-text-tools-panel', 'text-image-tools-panel', 'invert-image-tools-panel',
        'pdf-text-tools-panel', 'word-text-tools-panel', 'merge-pdf-tools-panel',
        'split-pdf-tools-panel', 'watermark-pdf-tools-panel', 'pdf-jpg-tools-panel',
        'rotate-pdf-tools-panel', 'page-numbers-tools-panel', 'html-pdf-tools-panel'
    ];

    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            tabBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentMode = btn.dataset.tab;

            // Clear workspace on switch
            selectedFiles = [];
            updateUI();

            // Hide all panels first
            allPanels.forEach(id => {
                const el = document.getElementById(id);
                if (el) el.classList.add('hidden');
            });

            // Specific UI logic per mode
            if (currentMode === 'pdf-mode' || currentMode === 'compress-mode') {
                const panel = document.getElementById('image-tools-panel');
                if (panel) panel.classList.remove('hidden');
                
                if (currentMode === 'pdf-mode') {
                    compressorSettings.classList.add('hidden');
                    pdfActions.classList.remove('hidden');
                    compressActions.classList.add('hidden');
                    document.querySelector('.drag-hint').classList.remove('hidden');
                } else {
                    compressorSettings.classList.remove('hidden');
                    pdfActions.classList.add('hidden');
                    compressActions.classList.remove('hidden');
                    document.querySelector('.drag-hint').classList.add('hidden');
                }
            } else if (currentMode === 'pdf-compress-mode') {
                document.getElementById('pdf-tools-panel').classList.remove('hidden');
                uploadedPdfFile = null;
                uploadedPdfDataUrl = null;
                if(pdfPreviewSection) pdfPreviewSection.classList.add('hidden');
                if(pdfUploadZone) pdfUploadZone.classList.remove('hidden');
            } else if (currentMode === 'word-mode') {
                document.getElementById('word-tools-panel').classList.remove('hidden');
                uploadedWordFile = null;
                if(wordPreviewSection) wordPreviewSection.classList.add('hidden');
                if(wordUploadZone) wordUploadZone.classList.remove('hidden');
            } else if (currentMode === 'pdf-word-mode') {
                document.getElementById('pdf-word-tools-panel').classList.remove('hidden');
                uploadedPdfWordFile = null;
                uploadedPdfWordDataUrl = null;
                if(pdfWordPreviewSection) pdfWordPreviewSection.classList.add('hidden');
                if(pdfWordUploadZone) pdfWordUploadZone.classList.remove('hidden');
            } else if (currentMode === 'pdf-translate-mode') {
                document.getElementById('pdf-translate-tools-panel').classList.remove('hidden');
                uploadedPdfTranslateFile = null;
                uploadedPdfTranslateDataUrl = null;
                if(pdfTranslatePreviewSection) pdfTranslatePreviewSection.classList.add('hidden');
                if(pdfTranslateUploadZone) pdfTranslateUploadZone.classList.remove('hidden');
            } else if (currentMode === 'text-pdf-mode') {
                document.getElementById('text-pdf-tools-panel').classList.remove('hidden');
            } else if (currentMode === 'image-text-mode') {
                document.getElementById('image-text-tools-panel').classList.remove('hidden');
                document.getElementById('image-text-preview-section').classList.add('hidden');
                document.getElementById('image-text-upload-zone').classList.remove('hidden');
            } else if (currentMode === 'text-image-mode') {
                document.getElementById('text-image-tools-panel').classList.remove('hidden');
            } else if (currentMode === 'invert-image-mode') {
                document.getElementById('invert-image-tools-panel').classList.remove('hidden');
                document.getElementById('invert-image-preview-section').classList.add('hidden');
                document.getElementById('invert-image-upload-zone').classList.remove('hidden');
            } else if (currentMode === 'pdf-text-mode') {
                document.getElementById('pdf-text-tools-panel').classList.remove('hidden');
                document.getElementById('pdf-text-preview-section').classList.add('hidden');
                document.getElementById('pdf-text-upload-zone').classList.remove('hidden');
            } else if (currentMode === 'word-text-mode') {
                document.getElementById('word-text-tools-panel').classList.remove('hidden');
                document.getElementById('word-text-preview-section').classList.add('hidden');
                document.getElementById('word-text-upload-zone').classList.remove('hidden');
            } else if (currentMode === 'merge-pdf-mode') {
                document.getElementById('merge-pdf-tools-panel').classList.remove('hidden');
                mergePdfFiles = [];
                if(mergePdfList) mergePdfList.innerHTML = '';
                document.getElementById('merge-pdf-preview-section').classList.add('hidden');
                document.getElementById('merge-pdf-upload-zone').classList.remove('hidden');
            } else if (currentMode === 'split-pdf-mode') {
                document.getElementById('split-pdf-tools-panel').classList.remove('hidden');
                splitPdfFile = null;
                splitPdfArrayBuffer = null;
                document.getElementById('split-pdf-preview-section').classList.add('hidden');
                document.getElementById('split-pdf-upload-zone').classList.remove('hidden');
            } else if (currentMode === 'watermark-pdf-mode') {
                document.getElementById('watermark-pdf-tools-panel').classList.remove('hidden');
                watermarkPdfFile = null;
                watermarkPdfArrayBuffer = null;
                document.getElementById('watermark-pdf-preview-section').classList.add('hidden');
                document.getElementById('watermark-pdf-upload-zone').classList.remove('hidden');
            } else if (currentMode === 'pdf-jpg-mode') {
                document.getElementById('pdf-jpg-tools-panel').classList.remove('hidden');
                pdfJpgFile = null;
                pdfJpgArrayBuffer = null;
                document.getElementById('pdf-jpg-preview-section').classList.add('hidden');
                document.getElementById('pdf-jpg-upload-zone').classList.remove('hidden');
            } else if (currentMode === 'rotate-pdf-mode') {
                document.getElementById('rotate-pdf-tools-panel').classList.remove('hidden');
                rotatePdfFile = null;
                rotatePdfArrayBuffer = null;
                document.getElementById('rotate-pdf-preview-section').classList.add('hidden');
                document.getElementById('rotate-pdf-upload-zone').classList.remove('hidden');
            } else if (currentMode === 'page-numbers-mode') {
                document.getElementById('page-numbers-tools-panel').classList.remove('hidden');
                pageNumbersFile = null;
                pageNumbersArrayBuffer = null;
                document.getElementById('page-numbers-preview-section').classList.add('hidden');
                document.getElementById('page-numbers-upload-zone').classList.remove('hidden');
            } else if (currentMode === 'html-pdf-mode') {
                document.getElementById('html-pdf-tools-panel').classList.remove('hidden');
                htmlPdfFile = null;
                document.getElementById('html-pdf-preview-section').classList.add('hidden');
                document.getElementById('html-pdf-upload-zone').classList.remove('hidden');
            }
        });
    });

    // Handle toggle for custom KB input
    window.toggleCustomKB = function () {
        const select = document.getElementById('compression-level');
        const container = document.getElementById('custom-kb-container');
        if (select.value === 'custom') {
            container.classList.remove('hidden');
        } else {
            container.classList.add('hidden');
        }
    };

    const compressionSelect = document.getElementById('compression-level');
    if (compressionSelect) {
        compressionSelect.addEventListener('change', window.toggleCustomKB);
        window.toggleCustomKB();
    }

    // --- Event Listeners ---

    uploadZone.addEventListener('click', (e) => {
        if (e.target !== fileInput && !e.target.closest('label')) {
            fileInput.click();
        }
    });

    fileInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
        fileInput.value = '';
    });

    uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadZone.classList.add('drag-active');
    });

    ['dragleave', 'dragend'].forEach(type => {
        uploadZone.addEventListener(type, () => {
            uploadZone.classList.remove('drag-active');
        });
    });

    uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadZone.classList.remove('drag-active');
        if (e.dataTransfer.files.length) {
            handleFiles(e.dataTransfer.files);
        }
    });

    if (addMoreImagesBtn) {
        addMoreImagesBtn.addEventListener('click', () => {
            fileInput.click();
        });
    }

    clearAllBtn.addEventListener('click', () => {
        selectedFiles = [];
        updateUI();
    });

    generatePdfBtn.addEventListener('click', generatePDF);
    compressDownloadBtn.addEventListener('click', compressAndDownloadImages);

    // --- PDF Compressor Event Listeners ---
    pdfUploadZone.addEventListener('click', (e) => {
        if (e.target !== pdfFileInput && !e.target.closest('label')) {
            pdfFileInput.click();
        }
    });

    pdfFileInput.addEventListener('change', (e) => {
        handlePdfFile(e.target.files[0]);
        pdfFileInput.value = '';
    });

    pdfUploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        pdfUploadZone.classList.add('drag-active');
    });

    ['dragleave', 'dragend'].forEach(type => {
        pdfUploadZone.addEventListener(type, () => {
            pdfUploadZone.classList.remove('drag-active');
        });
    });

    pdfUploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        pdfUploadZone.classList.remove('drag-active');
        if (e.dataTransfer.files.length) {
            handlePdfFile(e.dataTransfer.files[0]);
        }
    });

    pdfClearBtn.addEventListener('click', () => {
        uploadedPdfFile = null;
        uploadedPdfDataUrl = null;
        pdfPreviewSection.classList.add('hidden');
        pdfUploadZone.classList.remove('hidden');
    });

    compressPdfBtn.addEventListener('click', compressUploadedPDF);

    // --- Word to PDF Event Listeners ---
    if (wordUploadZone) {
        wordUploadZone.addEventListener('click', (e) => {
            if (e.target !== wordFileInput && !e.target.closest('label')) {
                wordFileInput.click();
            }
        });

        wordFileInput.addEventListener('change', (e) => {
            handleWordFile(e.target.files[0]);
            wordFileInput.value = '';
        });

        wordUploadZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            wordUploadZone.classList.add('drag-active');
        });

        ['dragleave', 'dragend'].forEach(type => {
            wordUploadZone.addEventListener(type, () => {
                wordUploadZone.classList.remove('drag-active');
            });
        });

        wordUploadZone.addEventListener('drop', (e) => {
            e.preventDefault();
            wordUploadZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) {
                handleWordFile(e.dataTransfer.files[0]);
            }
        });

        wordClearBtn.addEventListener('click', () => {
            uploadedWordFile = null;
            wordPreviewSection.classList.add('hidden');
            wordUploadZone.classList.remove('hidden');
        });

        convertWordBtn.addEventListener('click', convertWordToPDF);
    }

    // --- PDF to Word Event Listeners ---
    if (pdfWordUploadZone) {
        pdfWordUploadZone.addEventListener('click', (e) => {
            if (e.target !== pdfWordFileInput && !e.target.closest('label')) {
                pdfWordFileInput.click();
            }
        });

        pdfWordFileInput.addEventListener('change', (e) => {
            handlePdfWordFile(e.target.files[0]);
            pdfWordFileInput.value = '';
        });

        pdfWordUploadZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            pdfWordUploadZone.classList.add('drag-active');
        });

        ['dragleave', 'dragend'].forEach(type => {
            pdfWordUploadZone.addEventListener(type, () => {
                pdfWordUploadZone.classList.remove('drag-active');
            });
        });

        pdfWordUploadZone.addEventListener('drop', (e) => {
            e.preventDefault();
            pdfWordUploadZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) {
                handlePdfWordFile(e.dataTransfer.files[0]);
            }
        });

        pdfWordClearBtn.addEventListener('click', () => {
            uploadedPdfWordFile = null;
            uploadedPdfWordDataUrl = null;
            pdfWordPreviewSection.classList.add('hidden');
            pdfWordUploadZone.classList.remove('hidden');
        });

        convertPdfWordBtn.addEventListener('click', convertPdfToWord);
    }

    // --- PDF Translator Event Listeners ---
    if (pdfTranslateUploadZone) {
        pdfTranslateUploadZone.addEventListener('click', (e) => {
            if (e.target !== pdfTranslateFileInput && !e.target.closest('label')) {
                pdfTranslateFileInput.click();
            }
        });

        pdfTranslateFileInput.addEventListener('change', (e) => {
            handlePdfTranslateFile(e.target.files[0]);
            pdfTranslateFileInput.value = '';
        });

        pdfTranslateUploadZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            pdfTranslateUploadZone.classList.add('drag-active');
        });

        ['dragleave', 'dragend'].forEach(type => {
            pdfTranslateUploadZone.addEventListener(type, () => {
                pdfTranslateUploadZone.classList.remove('drag-active');
            });
        });

        pdfTranslateUploadZone.addEventListener('drop', (e) => {
            e.preventDefault();
            pdfTranslateUploadZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) {
                handlePdfTranslateFile(e.dataTransfer.files[0]);
            }
        });

        pdfTranslateClearBtn.addEventListener('click', () => {
            uploadedPdfTranslateFile = null;
            uploadedPdfTranslateDataUrl = null;
            pdfTranslatePreviewSection.classList.add('hidden');
            pdfTranslateUploadZone.classList.remove('hidden');
        });

        convertPdfTranslateBtn.addEventListener('click', convertPdfTranslate);
    }

    // --- Text to PDF Initialization ---
    let quill;
    if (document.getElementById('editor-container')) {
        quill = new Quill('#editor-container', {
            theme: 'snow',
            placeholder: 'Enter or Paste text here...',
            modules: {
                toolbar: [
                    [{ 'header': [1, 2, 3, false] }],
                    [{ 'font': [] }],
                    [{ 'size': ['small', false, 'large', 'huge'] }],
                    ['bold', 'italic', 'underline', 'strike'],
                    [{ 'color': [] }, { 'background': [] }],
                    [{ 'list': 'ordered'}, { 'list': 'bullet' }],
                    [{ 'align': [] }],
                    ['link', 'image'],
                    ['clean']
                ]
            }
        });
    }

    const textPdfClearBtn = document.getElementById('text-pdf-clear-btn');
    if (textPdfClearBtn) {
        textPdfClearBtn.addEventListener('click', () => {
            if (quill) quill.setContents([]);
            document.getElementById('text-pdf-output-filename-input').value = '';
        });
    }

    if (convertTextPdfBtn) {
        convertTextPdfBtn.addEventListener('click', async () => {
            if (!quill) return;
            const content = quill.root.innerHTML;
            if (!quill.getText().trim() && content === '<p><br></p>') {
                alert('Please enter some text before converting.');
                return;
            }

            loadingOverlay.classList.remove('hidden');
            loadingOverlay.querySelector('p').textContent = 'Converting Text to PDF...';

            try {
                let finalName = document.getElementById('text-pdf-output-filename-input').value.trim();
                if (!finalName) finalName = 'Document';
                if (!finalName.toLowerCase().endsWith('.pdf')) finalName += '.pdf';

                const htmlContent = `
                    <div class="ql-editor" style="font-family: Arial, sans-serif; color: black; background: white; text-align: left; line-height: 1.5; font-size: 16px; height: auto !important; overflow: visible !important;">
                        ${content}
                    </div>
                `;

                const opt = {
                    margin: 15,
                    filename: finalName,
                    image: { type: 'jpeg', quality: 0.98 },
                    html2canvas: { scale: 2, useCORS: true, scrollY: 0 },
                    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
                };

                await html2pdf().set(opt).from(htmlContent).save();
            } catch (error) {
                console.error("Error converting text to PDF:", error);
                alert("An error occurred during conversion.");
            } finally {
                loadingOverlay.classList.add('hidden');
            }
        });
    }

    // --- Core Functions ---

    function handleFiles(files) {
        const validFiles = Array.from(files).filter(file => file.type.startsWith('image/'));
        if (validFiles.length === 0) return;

        let processedCount = 0;

        validFiles.forEach(file => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const dataUrl = e.target.result;
                selectedFiles.push({
                    id: 'img_' + Math.random().toString(36).substr(2, 9),
                    file: file,
                    dataUrl: dataUrl
                });

                processedCount++;
                if (processedCount === validFiles.length) {
                    updateUI();
                }
            };
            reader.readAsDataURL(file);
        });
    }

    function updateUI() {
        if (selectedFiles.length > 0) {
            previewSection.classList.remove('hidden');
            renderImages();
        } else {
            previewSection.classList.add('hidden');
            imageGrid.innerHTML = '';
        }
        imageCount.textContent = `${selectedFiles.length} image${selectedFiles.length !== 1 ? 's' : ''}`;
    }

    function renderImages() {
        imageGrid.innerHTML = '';
        selectedFiles.forEach((item, index) => {
            const card = document.createElement('div');
            card.className = 'image-card';
            card.dataset.id = item.id;

            card.innerHTML = `
                <img src="${item.dataUrl}" alt="Selected image">
                <button class="remove-btn" onclick="removeImage('${item.id}', event)">
                    <i class="fa-solid fa-xmark"></i>
                </button>
                <div class="image-number">${index + 1}</div>
            `;
            imageGrid.appendChild(card);
        });
    }

    window.removeImage = function (id, event) {
        event.stopPropagation();
        selectedFiles = selectedFiles.filter(item => item.id !== id);
        updateUI();
    };

    function updateNumbers() {
        const cards = imageGrid.querySelectorAll('.image-card');
        cards.forEach((card, index) => {
            card.querySelector('.image-number').textContent = index + 1;
        });
    }

    function reorderFilesArray() {
        const newOrder = Array.from(imageGrid.children).map(card => card.dataset.id);
        const reordered = [];
        newOrder.forEach(id => {
            const file = selectedFiles.find(item => item.id === id);
            if (file) reordered.push(file);
        });
        selectedFiles = reordered;
    }

    async function generatePDF() {
        if (selectedFiles.length === 0) return;
        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Generating PDF...';

        try {
            const { jsPDF } = window.jspdf;
            const pdf = new jsPDF('p', 'mm', 'a4');
            const pdfWidth = pdf.internal.pageSize.getWidth();
            const pdfHeight = pdf.internal.pageSize.getHeight();

            for (let i = 0; i < selectedFiles.length; i++) {
                if (i > 0) pdf.addPage();

                const item = selectedFiles[i];
                const img = new Image();
                img.src = item.dataUrl;
                await new Promise(resolve => img.onload = resolve);

                const margin = 10;
                const availableWidth = pdfWidth - (margin * 2);
                const availableHeight = pdfHeight - (margin * 2);

                const ratio = Math.min(availableWidth / img.width, availableHeight / img.height);
                const drawWidth = img.width * ratio;
                const drawHeight = img.height * ratio;

                const x = (pdfWidth - drawWidth) / 2;
                const y = (pdfHeight - drawHeight) / 2;

                const format = item.file.type === 'image/png' ? 'PNG' : 'JPEG';
                pdf.addImage(item.dataUrl, format, x, y, drawWidth, drawHeight);
            }

            let filename = document.getElementById('output-filename-input').value.trim();
            if (!filename) filename = 'Document';
            if (!filename.toLowerCase().endsWith('.pdf')) filename += '.pdf';

            pdf.save(filename);
        } catch (error) {
            console.error("Error:", error);
            alert("An error occurred while generating the PDF.");
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }

    async function compressAndDownloadImages() {
        if (selectedFiles.length === 0) return;
        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Compressing Images...';

        const compressionLevel = compressionSelect ? compressionSelect.value : 'none';

        let maxWidth = 2000;
        if (compressionLevel === 'low') maxWidth = 1500;
        if (compressionLevel === 'medium') maxWidth = 1000;
        if (compressionLevel === 'high') maxWidth = 800;
        if (compressionLevel === 'extreme') maxWidth = 600;

        try {
            for (let i = 0; i < selectedFiles.length; i++) {
                const item = selectedFiles[i];
                const img = new Image();
                img.src = item.dataUrl;
                await new Promise(resolve => img.onload = resolve);

                let finalDataUrl = item.dataUrl;

                if (compressionLevel !== 'none') {
                    let finalWidth = img.width;
                    let finalHeight = img.height;

                    if (compressionLevel === 'custom') {
                        const targetKB = parseFloat(document.getElementById('target-kb-input').value) || 50;
                        const targetBytes = targetKB * 1024;
                        let currentQ = 0.8;
                        let iterationWidth = img.width;
                        let iterationHeight = img.height;

                        while (true) {
                            const canvas = document.createElement('canvas');
                            canvas.width = iterationWidth;
                            canvas.height = iterationHeight;
                            const ctx = canvas.getContext('2d');
                            ctx.fillStyle = '#ffffff';
                            ctx.fillRect(0, 0, iterationWidth, iterationHeight);
                            ctx.drawImage(img, 0, 0, iterationWidth, iterationHeight);

                            let testDataUrl = canvas.toDataURL('image/jpeg', currentQ);
                            let estimatedBytes = Math.round(testDataUrl.length * 3 / 4);

                            if (estimatedBytes <= targetBytes || (currentQ <= 0.1 && iterationWidth <= 300)) {
                                finalDataUrl = testDataUrl;
                                break;
                            }

                            if (currentQ > 0.3) {
                                currentQ -= 0.2;
                            } else {
                                iterationWidth = Math.max(300, iterationWidth * 0.7);
                                iterationHeight = Math.max(300 * (img.height / img.width), iterationHeight * 0.7);
                                currentQ = 0.5;
                            }
                        }
                    } else {
                        if (finalWidth > maxWidth) {
                            finalHeight = Math.round(finalHeight * maxWidth / finalWidth);
                            finalWidth = maxWidth;
                        }
                        const canvas = document.createElement('canvas');
                        canvas.width = finalWidth;
                        canvas.height = finalHeight;
                        const ctx = canvas.getContext('2d');
                        ctx.fillStyle = '#ffffff';
                        ctx.fillRect(0, 0, finalWidth, finalHeight);
                        ctx.drawImage(img, 0, 0, finalWidth, finalHeight);

                        let qualityNum = 0.8;
                        switch (compressionLevel) {
                            case 'low': qualityNum = 0.8; break;
                            case 'medium': qualityNum = 0.6; break;
                            case 'high': qualityNum = 0.4; break;
                            case 'extreme': qualityNum = 0.15; break;
                        }
                        finalDataUrl = canvas.toDataURL('image/jpeg', qualityNum);
                    }
                }

                // Trigger individual download
                const link = document.createElement('a');
                link.href = finalDataUrl;

                let prefix = document.getElementById('output-filename-input').value.trim();

                // Preserve original name but change extension if converted to jpeg
                const originalName = item.file.name;
                const baseName = originalName.substring(0, originalName.lastIndexOf('.')) || originalName;
                const ext = compressionLevel === 'none' ? originalName.substring(originalName.lastIndexOf('.')) : '.jpg';

                let finalName = prefix ? `${prefix}_${i + 1}${ext}` : `${baseName}_compressed${ext}`;
                if (selectedFiles.length === 1 && prefix) {
                    finalName = prefix.toLowerCase().endsWith(ext) ? prefix : `${prefix}${ext}`;
                }

                link.download = finalName;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);

                // Small delay to allow browser to process multiple downloads
                await new Promise(r => setTimeout(r, 300));
            }
        } catch (error) {
            console.error("Error compressing:", error);
            alert("An error occurred during compression.");
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }

    // --- PDF Compressor Logic ---
    function handlePdfFile(file) {
        if (!file || file.type !== 'application/pdf') {
            alert('Please select a valid PDF file.');
            return;
        }

        uploadedPdfFile = file;

        const reader = new FileReader();
        reader.onload = (e) => {
            uploadedPdfDataUrl = e.target.result;

            // Format size
            const sizeKB = (file.size / 1024).toFixed(2);
            let sizeStr = sizeKB + ' KB';
            if (sizeKB > 1024) {
                sizeStr = (sizeKB / 1024).toFixed(2) + ' MB';
            }

            pdfFilename.textContent = file.name;
            pdfOriginalSize.textContent = sizeStr;

            const baseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
            document.getElementById('pdf-output-filename-input').value = baseName + '_compressed';

            pdfUploadZone.classList.add('hidden');
            pdfPreviewSection.classList.remove('hidden');
        };
        reader.readAsDataURL(file);
    }

    async function compressUploadedPDF() {
        if (!uploadedPdfFile || !uploadedPdfDataUrl) return;

        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Rasterizing & Compressing PDF... (This may take a minute)';

        try {
            // Target KB
            const totalTargetKB = parseFloat(pdfTargetKbInput.value) || 500;
            const totalTargetBytes = totalTargetKB * 1024;

            // Load PDF using pdf.js
            const pdfjsLib = window['pdfjs-dist/build/pdf'];
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

            // Convert base64 to Uint8Array for pdf.js
            const base64Data = uploadedPdfDataUrl.split(',')[1];
            const binaryString = window.atob(base64Data);
            const bytes = new Uint8Array(binaryString.length);
            for (let i = 0; i < binaryString.length; i++) {
                bytes[i] = binaryString.charCodeAt(i);
            }

            const loadingTask = pdfjsLib.getDocument({ data: bytes });
            const pdfDocument = await loadingTask.promise;
            const numPages = pdfDocument.numPages;

            // Target bytes per page roughly
            const targetBytesPerPage = totalTargetBytes / numPages;

            const { jsPDF } = window.jspdf;
            const newPdf = new jsPDF('p', 'mm', 'a4');
            const pdfWidth = newPdf.internal.pageSize.getWidth();
            const pdfHeight = newPdf.internal.pageSize.getHeight();

            for (let pageNum = 1; pageNum <= numPages; pageNum++) {
                if (pageNum > 1) {
                    newPdf.addPage();
                }

                loadingOverlay.querySelector('p').textContent = `Compressing Page ${pageNum} of ${numPages}...`;

                const page = await pdfDocument.getPage(pageNum);

                // Render at a decent scale first (e.g. 2.0 for 144 DPI equivalent)
                let viewport = page.getViewport({ scale: 2.0 });
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                canvas.width = viewport.width;
                canvas.height = viewport.height;

                const renderContext = {
                    canvasContext: ctx,
                    viewport: viewport
                };
                await page.render(renderContext).promise;

                // Iterative compression logic per page
                let currentQ = 0.8;
                let iterationWidth = canvas.width;
                let iterationHeight = canvas.height;
                let finalDataUrl = '';

                while (true) {
                    const tempCanvas = document.createElement('canvas');
                    tempCanvas.width = iterationWidth;
                    tempCanvas.height = iterationHeight;
                    const tempCtx = tempCanvas.getContext('2d');
                    tempCtx.fillStyle = '#ffffff';
                    tempCtx.fillRect(0, 0, iterationWidth, iterationHeight);
                    tempCtx.drawImage(canvas, 0, 0, iterationWidth, iterationHeight);

                    let testDataUrl = tempCanvas.toDataURL('image/jpeg', currentQ);
                    let estimatedBytes = Math.round(testDataUrl.length * 3 / 4);

                    if (estimatedBytes <= targetBytesPerPage || (currentQ <= 0.1 && iterationWidth <= 600)) {
                        finalDataUrl = testDataUrl;
                        break;
                    }

                    if (currentQ > 0.3) {
                        currentQ -= 0.2;
                    } else {
                        iterationWidth = Math.max(400, iterationWidth * 0.7);
                        iterationHeight = Math.max(400 * (canvas.height / canvas.width), iterationHeight * 0.7);
                        currentQ = 0.5;
                    }
                }

                const ratio = Math.min(pdfWidth / viewport.width, pdfHeight / viewport.height);
                const drawWidth = viewport.width * ratio;
                const drawHeight = viewport.height * ratio;

                const x = (pdfWidth - drawWidth) / 2;
                const y = (pdfHeight - drawHeight) / 2;

                newPdf.addImage(finalDataUrl, 'JPEG', x, y, drawWidth, drawHeight);
            }

            // Save
            let finalName = document.getElementById('pdf-output-filename-input').value.trim();
            if (!finalName) {
                const originalName = uploadedPdfFile.name;
                finalName = originalName.substring(0, originalName.lastIndexOf('.')) || originalName;
                finalName += '_compressed';
            }
            if (!finalName.toLowerCase().endsWith('.pdf')) finalName += '.pdf';
            newPdf.save(finalName);

        } catch (error) {
            console.error("Error compressing PDF:", error);
            alert("An error occurred while compressing the PDF.");
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }

    // --- Word to PDF Logic ---
    function handleWordFile(file) {
        if (!file || (!file.name.toLowerCase().endsWith('.doc') && !file.name.toLowerCase().endsWith('.docx'))) {
            alert('Please select a valid Word document (.doc or .docx).');
            return;
        }

        uploadedWordFile = file;

        // Format size
        const sizeKB = (file.size / 1024).toFixed(2);
        let sizeStr = sizeKB + ' KB';
        if (sizeKB > 1024) {
            sizeStr = (sizeKB / 1024).toFixed(2) + ' MB';
        }

        wordFilename.textContent = file.name;
        wordOriginalSize.textContent = sizeStr;

        const baseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
        document.getElementById('word-output-filename-input').value = baseName;

        wordUploadZone.classList.add('hidden');
        wordPreviewSection.classList.remove('hidden');
    }

    async function convertWordToPDF() {
        if (!uploadedWordFile) return;

        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Converting Word to PDF...';

        try {
            const reader = new FileReader();
            reader.onload = async (e) => {
                const arrayBuffer = e.target.result;

                try {
                    // Use Mammoth to extract HTML
                    const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });

                    const htmlContent = `
                        <div style="font-family: 'Times New Roman', serif; color: black; background: white; text-align: left; line-height: 1.5;">
                            ${result.value || '<p>No readable text found in this document.</p>'}
                        </div>
                    `;

                    // Setup html2pdf options
                    let finalName = document.getElementById('word-output-filename-input').value.trim();
                    if (!finalName) {
                        const originalName = uploadedWordFile.name;
                        finalName = originalName.substring(0, originalName.lastIndexOf('.')) || originalName;
                    }
                    if (!finalName.toLowerCase().endsWith('.pdf')) finalName += '.pdf';

                    const opt = {
                        margin: 15,
                        filename: finalName,
                        image: { type: 'jpeg', quality: 0.98 },
                        html2canvas: { scale: 2 },
                        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
                    };

                    // Generate PDF from the HTML string directly
                    await html2pdf().set(opt).from(htmlContent).save();

                } catch (err) {
                    console.error("Mammoth/html2pdf error:", err);
                    alert("Failed to convert the document. It might be corrupted or unsupported.");
                } finally {
                    loadingOverlay.classList.add('hidden');
                }
            };

            reader.onerror = () => {
                console.error("FileReader error");
                loadingOverlay.classList.add('hidden');
                alert("Failed to read the file.");
            };

            reader.readAsArrayBuffer(uploadedWordFile);

        } catch (error) {
            console.error("Error starting conversion:", error);
            loadingOverlay.classList.add('hidden');
        }
    }

    // --- PDF to Word Logic ---
    function handlePdfWordFile(file) {
        if (!file || file.type !== 'application/pdf') {
            alert('Please select a valid PDF file.');
            return;
        }

        uploadedPdfWordFile = file;

        const reader = new FileReader();
        reader.onload = (e) => {
            uploadedPdfWordDataUrl = e.target.result;

            // Format size
            const sizeKB = (file.size / 1024).toFixed(2);
            let sizeStr = sizeKB + ' KB';
            if (sizeKB > 1024) {
                sizeStr = (sizeKB / 1024).toFixed(2) + ' MB';
            }

            pdfWordFilename.textContent = file.name;
            pdfWordOriginalSize.textContent = sizeStr;

            const baseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
            document.getElementById('pdf-word-output-filename-input').value = baseName;

            pdfWordUploadZone.classList.add('hidden');
            pdfWordPreviewSection.classList.remove('hidden');
        };
        reader.readAsDataURL(file);
    }

    async function convertPdfToWord() {
        if (!uploadedPdfWordFile || !uploadedPdfWordDataUrl) return;

        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Extracting Text from PDF...';

        try {
            const pdfjsLib = window['pdfjs-dist/build/pdf'];
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

            const base64Data = uploadedPdfWordDataUrl.split(',')[1];
            const binaryString = window.atob(base64Data);
            const bytes = new Uint8Array(binaryString.length);
            for (let i = 0; i < binaryString.length; i++) {
                bytes[i] = binaryString.charCodeAt(i);
            }

            const loadingTask = pdfjsLib.getDocument({ data: bytes });
            const pdfDocument = await loadingTask.promise;
            const numPages = pdfDocument.numPages;

            let extractedHtml = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
            <head><meta charset='utf-8'><title>Export HTML To Doc</title></head><body>`;

            let hasText = false;

            for (let pageNum = 1; pageNum <= numPages; pageNum++) {
                loadingOverlay.querySelector('p').textContent = `Processing Page ${pageNum} of ${numPages}...`;
                const page = await pdfDocument.getPage(pageNum);

                let pageHtml = "";
                let hasExtractableText = false;

                const textContent = await page.getTextContent();
                if (textContent.items.length > 0) {
                    hasExtractableText = true;
                    hasText = true;
                    // Group items by approximate Y position to form lines/paragraphs
                    let lastY = -1;
                    let paragraph = "<p style='margin-bottom: 10px; font-family: \"Times New Roman\", serif; font-size: 12pt; color: black;'>";

                    textContent.items.forEach(item => {
                        if (lastY !== -1 && Math.abs(item.transform[5] - lastY) > 5) {
                            // New line
                            paragraph += "</p><p style='margin-bottom: 10px; font-family: \"Times New Roman\", serif; font-size: 12pt; color: black;'>";
                        }
                        paragraph += escapeHtml(item.str) + " ";
                        lastY = item.transform[5];
                    });
                    paragraph += "</p>";
                    pageHtml = paragraph;
                }

                if (!hasExtractableText) {
                    // Fallback to OCR or forced OCR
                    loadingOverlay.querySelector('p').textContent = `Starting OCR on Page ${pageNum}... (First time may take a moment)`;

                    // Render page to canvas
                    const viewport = page.getViewport({ scale: 2.0 }); // High scale for better OCR accuracy
                    const canvas = document.createElement('canvas');
                    const ctx = canvas.getContext('2d');
                    canvas.width = viewport.width;
                    canvas.height = viewport.height;

                    await page.render({
                        canvasContext: ctx,
                        viewport: viewport
                    }).promise;

                    const imageData = canvas.toDataURL('image/png');

                    // Run Tesseract
                    const result = await Tesseract.recognize(
                        imageData,
                        'eng',
                        {
                            logger: m => {
                                if (m.status === 'recognizing text') {
                                    const pct = Math.round(m.progress * 100);
                                    loadingOverlay.querySelector('p').textContent = `OCR Page ${pageNum}: ${pct}% completed...`;
                                } else if (m.status === 'loading tesseract core' || m.status === 'loading language traineddata') {
                                    loadingOverlay.querySelector('p').textContent = `Downloading OCR Engines...`;
                                }
                            }
                        }
                    );

                    if (result.data.text.trim().length > 0) {
                        hasText = true;
                        // Split by newlines and create paragraphs
                        const lines = result.data.text.split('\n').filter(line => line.trim().length > 0);
                        const paragraphs = lines.map(p => `<p style='margin-bottom: 10px; font-family: "Times New Roman", serif; font-size: 12pt; color: black;'>${escapeHtml(p)}</p>`).join('');
                        pageHtml = paragraphs;
                    }
                }

                if (pageHtml) {
                    extractedHtml += pageHtml;
                    extractedHtml += "<br style='page-break-before:always; clear:both' />"; // Page break
                }
            }

            extractedHtml += "</body></html>";

            if (!hasText) {
                alert("No text could be extracted, even with OCR. The document might be blank or corrupted.");
                loadingOverlay.classList.add('hidden');
                return;
            }

            // Create Word Doc Blob
            const blob = new Blob(['\ufeff', extractedHtml], {
                type: 'application/msword'
            });

            // Save File
            let finalName = document.getElementById('pdf-word-output-filename-input').value.trim();
            if (!finalName) {
                const originalName = uploadedPdfWordFile.name;
                finalName = originalName.substring(0, originalName.lastIndexOf('.')) || originalName;
            }
            if (!finalName.toLowerCase().endsWith('.doc')) finalName += '.doc';

            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = finalName;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

        } catch (error) {
            console.error("Error extracting PDF to Word:", error);
            alert("An error occurred while converting the PDF.");
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }

    // --- PDF Translator Logic ---
    function handlePdfTranslateFile(file) {
        if (!file) return;

        const validExtensions = ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.pptx', '.txt'];
        const extension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

        if (!validExtensions.includes(extension)) {
            alert('Please select a supported document type (PDF, Word, Excel, PPT, TXT).');
            return;
        }

        uploadedPdfTranslateFile = file;

        const reader = new FileReader();
        reader.onload = (e) => {
            uploadedPdfTranslateDataUrl = e.target.result; // Stores ArrayBuffer

            // Format size
            const sizeKB = (file.size / 1024).toFixed(2);
            let sizeStr = sizeKB + ' KB';
            if (sizeKB > 1024) {
                sizeStr = (sizeKB / 1024).toFixed(2) + ' MB';
            }

            pdfTranslateFilename.textContent = file.name;
            pdfTranslateOriginalSize.textContent = sizeStr;

            const baseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
            document.getElementById('pdf-translate-output-filename-input').value = baseName + '_translated';

            pdfTranslateUploadZone.classList.add('hidden');
            pdfTranslatePreviewSection.classList.remove('hidden');
        };
        reader.readAsArrayBuffer(file);
    }

    async function translateTextChunk(text, targetLang) {
        if (!text.trim()) return "";
        try {
            const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=${targetLang}&dt=t&q=${encodeURIComponent(text)}`;
            const response = await fetch(url);
            if (!response.ok) throw new Error("Translation API Error");
            const data = await response.json();
            let translated = "";
            if (data && data[0]) {
                data[0].forEach(item => {
                    if (item[0]) translated += item[0];
                });
            }
            return translated;
        } catch (e) {
            console.error("Translation failed for chunk:", e);
            return text; // fallback to original
        }
    }

    async function extractTextFromFile(file, arrayBuffer) {
        const extension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
        let extractedSections = []; // Array of { sectionName: string, paragraphs: [string] }

        loadingOverlay.querySelector('p').textContent = 'Extracting text from document...';

        if (extension === '.pdf') {
            const pdfjsLib = window['pdfjs-dist/build/pdf'];
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
            const bytes = new Uint8Array(arrayBuffer);
            const loadingTask = pdfjsLib.getDocument({ data: bytes });
            const pdfDocument = await loadingTask.promise;
            const numPages = pdfDocument.numPages;

            for (let pageNum = 1; pageNum <= numPages; pageNum++) {
                loadingOverlay.querySelector('p').textContent = `Reading PDF Page ${pageNum}...`;
                const page = await pdfDocument.getPage(pageNum);
                const textContent = await page.getTextContent();
                let paragraphs = [];
                let hasExtractableText = false;

                if (textContent.items.length > 0) {
                    hasExtractableText = true;
                    let lastY = -1;
                    let currentParagraph = "";
                    textContent.items.forEach(item => {
                        if (lastY !== -1 && Math.abs(item.transform[5] - lastY) > 5) {
                            if (currentParagraph.trim()) paragraphs.push(currentParagraph.trim());
                            currentParagraph = "";
                        }
                        currentParagraph += item.str + " ";
                        lastY = item.transform[5];
                    });
                    if (currentParagraph.trim()) paragraphs.push(currentParagraph.trim());
                }

                if (!hasExtractableText) {
                    loadingOverlay.querySelector('p').textContent = `Running OCR on Page ${pageNum}...`;
                    const viewport = page.getViewport({ scale: 2.0 });
                    const canvas = document.createElement('canvas');
                    const ctx = canvas.getContext('2d');
                    canvas.width = viewport.width;
                    canvas.height = viewport.height;
                    await page.render({ canvasContext: ctx, viewport: viewport }).promise;
                    const imageData = canvas.toDataURL('image/png');
                    const result = await Tesseract.recognize(imageData, 'eng', {
                        logger: m => {
                            if (m.status === 'recognizing text') {
                                loadingOverlay.querySelector('p').textContent = `OCR Page ${pageNum}: ${Math.round(m.progress * 100)}%...`;
                            }
                        }
                    });
                    if (result.data.text.trim().length > 0) {
                        paragraphs = result.data.text.split('\n').filter(line => line.trim().length > 0);
                    }
                }

                extractedSections.push({ sectionName: `Page ${pageNum}`, paragraphs });
            }
        } else if (extension === '.doc' || extension === '.docx') {
            loadingOverlay.querySelector('p').textContent = 'Extracting Word Document...';
            const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
            const text = result.value;
            const paragraphs = text.split('\n').filter(line => line.trim().length > 0);
            extractedSections.push({ sectionName: 'Word Document', paragraphs });
        } else if (extension === '.xls' || extension === '.xlsx') {
            loadingOverlay.querySelector('p').textContent = 'Extracting Excel Sheet...';
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            workbook.SheetNames.forEach((sheetName) => {
                const worksheet = workbook.Sheets[sheetName];
                const csv = XLSX.utils.sheet_to_csv(worksheet);
                const paragraphs = csv.split('\n').filter(line => line.trim().length > 0 && line.replace(/,/g, '').trim().length > 0);
                if (paragraphs.length > 0) {
                    extractedSections.push({ sectionName: `Sheet: ${sheetName}`, paragraphs: [`--- ${sheetName} ---`, ...paragraphs] });
                }
            });
        } else if (extension === '.pptx') {
            loadingOverlay.querySelector('p').textContent = 'Extracting PowerPoint...';
            const zip = new JSZip();
            const zipFiles = await zip.loadAsync(arrayBuffer);
            const slideFiles = [];
            zipFiles.folder("ppt/slides").forEach((relativePath, file) => {
                if (relativePath.match(/^slide\d+\.xml$/)) {
                    slideFiles.push({ path: relativePath, file: file });
                }
            });

            slideFiles.sort((a, b) => {
                const numA = parseInt(a.path.match(/\d+/)[0]);
                const numB = parseInt(b.path.match(/\d+/)[0]);
                return numA - numB;
            });

            for (let i = 0; i < slideFiles.length; i++) {
                loadingOverlay.querySelector('p').textContent = `Reading Slide ${i + 1}...`;
                const xmlString = await slideFiles[i].file.async("string");
                const textMatches = xmlString.match(/<a:t>.*?<\/a:t>/g);
                if (textMatches) {
                    const paragraphs = textMatches.map(t => t.replace(/<[^>]+>/g, '')).filter(t => t.trim().length > 0);
                    if (paragraphs.length > 0) {
                        extractedSections.push({ sectionName: `Slide ${i + 1}`, paragraphs: [`--- Slide ${i + 1} ---`, ...paragraphs] });
                    }
                }
            }
        } else if (extension === '.txt') {
            loadingOverlay.querySelector('p').textContent = 'Reading Text File...';
            const decoder = new TextDecoder('utf-8');
            const text = decoder.decode(arrayBuffer);
            const paragraphs = text.split('\n').filter(line => line.trim().length > 0);
            extractedSections.push({ sectionName: 'Text Document', paragraphs });
        }

        return extractedSections;
    }

    async function convertPdfTranslate() {
        if (!uploadedPdfTranslateFile || !uploadedPdfTranslateDataUrl) return;

        const targetLang = pdfTranslateLang.value;

        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Preparing Document...';

        try {
            const extractedSections = await extractTextFromFile(uploadedPdfTranslateFile, uploadedPdfTranslateDataUrl);

            let htmlString = `<div style="padding: 15mm; font-family: 'Outfit', sans-serif; color: black; background: white; text-align: left; line-height: 1.6;">`;

            if (extractedSections.length === 0) {
                htmlString += `<p style="margin-bottom: 12px; font-size: 14pt;"><i>No text found in document.</i></p>`;
            }

            for (const section of extractedSections) {
                const { sectionName, paragraphs } = section;

                if (paragraphs.length > 0) {
                    // Group into larger chunks for API
                    let chunks = [];
                    let currentChunk = "";
                    paragraphs.forEach(p => {
                        if ((currentChunk.length + p.length) > 1000) {
                            chunks.push(currentChunk);
                            currentChunk = p;
                        } else {
                            currentChunk += (currentChunk ? "\n\n" : "") + p;
                        }
                    });
                    if (currentChunk) chunks.push(currentChunk);

                    // Translate chunks
                    for (let i = 0; i < chunks.length; i++) {
                        loadingOverlay.querySelector('p').textContent = `Translating ${sectionName}...`;
                        const translatedText = await translateTextChunk(chunks[i], targetLang);

                        // Add to HTML
                        const transLines = translatedText.split('\n\n');
                        transLines.forEach(line => {
                            if (line.trim()) {
                                htmlString += `<p style="margin-bottom: 12px; font-size: 14pt;">${escapeHtml(line.trim())}</p>`;
                            }
                        });

                        // Add a small delay to avoid rate limiting
                        await new Promise(r => setTimeout(r, 500));
                    }
                    htmlString += `<div style="page-break-after: always;"></div>`;
                } else {
                    htmlString += `<p style="margin-bottom: 12px; font-size: 14pt;"><i>[${sectionName}: No text available to translate.]</i></p><div style="page-break-after: always;"></div>`;
                }
            }

            htmlString += "</div>";

            loadingOverlay.querySelector('p').textContent = `Generating PDF...`;

            let finalName = document.getElementById('pdf-translate-output-filename-input').value.trim();
            if (!finalName) {
                const originalName = uploadedPdfTranslateFile.name;
                finalName = originalName.substring(0, originalName.lastIndexOf('.')) || originalName;
                finalName += '_translated';
            }
            if (!finalName.toLowerCase().endsWith('.pdf')) finalName += '.pdf';

            const opt = {
                margin: 0,
                filename: finalName,
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2, useCORS: true },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
            };

            await html2pdf().set(opt).from(htmlString).save();

        } catch (error) {
            console.error("Error translating PDF:", error);
            alert("An error occurred while translating the PDF. Check console for details.");
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }

    // Top-level HTML Escaping function
    function escapeHtml(unsafe) {
        if (!unsafe) return "";
        return unsafe
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    // --- Image to Text ---
    const imageTextUploadZone = document.getElementById('image-text-upload-zone');
    const imageTextFileInput = document.getElementById('image-text-file-input');
    const imageTextPreviewSection = document.getElementById('image-text-preview-section');
    const imageTextOutput = document.getElementById('image-text-output');
    const imageTextClearBtn = document.getElementById('image-text-clear-btn');
    const imageTextCopyBtn = document.getElementById('image-text-copy-btn');

    if (imageTextUploadZone) {
        imageTextUploadZone.addEventListener('click', (e) => {
            if (e.target !== imageTextFileInput && !e.target.closest('label')) {
                imageTextFileInput.click();
            }
        });

        imageTextFileInput.addEventListener('change', async (e) => {
            if (e.target.files.length) await extractTextFromImage(e.target.files[0]);
            imageTextFileInput.value = '';
        });

        imageTextUploadZone.addEventListener('dragover', (e) => { e.preventDefault(); imageTextUploadZone.classList.add('drag-active'); });
        ['dragleave', 'dragend'].forEach(type => { imageTextUploadZone.addEventListener(type, () => imageTextUploadZone.classList.remove('drag-active')); });
        imageTextUploadZone.addEventListener('drop', async (e) => {
            e.preventDefault();
            imageTextUploadZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) await extractTextFromImage(e.dataTransfer.files[0]);
        });

        imageTextClearBtn.addEventListener('click', () => {
            imageTextOutput.value = '';
            imageTextPreviewSection.classList.add('hidden');
            imageTextUploadZone.classList.remove('hidden');
        });

        imageTextCopyBtn.addEventListener('click', () => {
            imageTextOutput.select();
            document.execCommand('copy');
            alert('Text copied to clipboard!');
        });
    }

    async function extractTextFromImage(file) {
        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Extracting Text (OCR)...';
        try {
            const { data: { text } } = await Tesseract.recognize(file, 'eng');
            imageTextOutput.value = text || 'No readable text found.';
            imageTextUploadZone.classList.add('hidden');
            imageTextPreviewSection.classList.remove('hidden');
        } catch (err) {
            console.error('OCR Error:', err);
            alert('Failed to extract text from image.');
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }

    // --- Text to Image ---
    const textImageInput = document.getElementById('text-image-input');
    const textColorPicker = document.getElementById('text-color-picker');
    const bgColorPicker = document.getElementById('bg-color-picker');
    const convertTextImageBtn = document.getElementById('convert-text-image-btn');
    const textImageClearBtn = document.getElementById('text-image-clear-btn');

    if (textImageClearBtn) {
        textImageClearBtn.addEventListener('click', () => {
            textImageInput.value = '';
            textColorPicker.value = '#ffffff';
            bgColorPicker.value = '#3b82f6';
        });
    }

    if (convertTextImageBtn) {
        convertTextImageBtn.addEventListener('click', () => {
            const text = textImageInput.value.trim();
            if (!text) return alert('Please enter some text.');

            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            
            ctx.font = '24px Outfit, sans-serif';
            const lines = text.split('\n');
            let maxWidth = 0;
            lines.forEach(line => {
                const width = ctx.measureText(line).width;
                if(width > maxWidth) maxWidth = width;
            });

            canvas.width = Math.max(800, maxWidth + 100);
            canvas.height = Math.max(400, (lines.length * 35) + 100);

            const ctxFinal = canvas.getContext('2d');
            ctxFinal.fillStyle = bgColorPicker.value;
            ctxFinal.fillRect(0, 0, canvas.width, canvas.height);

            ctxFinal.font = '24px Outfit, sans-serif';
            ctxFinal.fillStyle = textColorPicker.value;
            ctxFinal.textAlign = 'center';
            ctxFinal.textBaseline = 'middle';

            const startY = (canvas.height - (lines.length * 35)) / 2 + 15;
            lines.forEach((line, index) => {
                ctxFinal.fillText(line, canvas.width / 2, startY + (index * 35));
            });

            const link = document.createElement('a');
            link.download = 'text_image.png';
            link.href = canvas.toDataURL('image/png');
            link.click();
        });
    }

    // --- Invert Image ---
    const invertImageUploadZone = document.getElementById('invert-image-upload-zone');
    const invertImageFileInput = document.getElementById('invert-image-file-input');
    const invertImagePreviewSection = document.getElementById('invert-image-preview-section');
    const invertImageCanvas = document.getElementById('invert-image-canvas');
    const invertImageClearBtn = document.getElementById('invert-image-clear-btn');
    const invertImageDownloadBtn = document.getElementById('invert-image-download-btn');

    if (invertImageUploadZone) {
        invertImageUploadZone.addEventListener('click', (e) => {
            if (e.target !== invertImageFileInput && !e.target.closest('label')) {
                invertImageFileInput.click();
            }
        });

        invertImageFileInput.addEventListener('change', (e) => {
            if (e.target.files.length) processInvertImage(e.target.files[0]);
            invertImageFileInput.value = '';
        });

        invertImageUploadZone.addEventListener('dragover', (e) => { e.preventDefault(); invertImageUploadZone.classList.add('drag-active'); });
        ['dragleave', 'dragend'].forEach(type => { invertImageUploadZone.addEventListener(type, () => invertImageUploadZone.classList.remove('drag-active')); });
        invertImageUploadZone.addEventListener('drop', (e) => {
            e.preventDefault();
            invertImageUploadZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) processInvertImage(e.dataTransfer.files[0]);
        });

        invertImageClearBtn.addEventListener('click', () => {
            invertImagePreviewSection.classList.add('hidden');
            invertImageUploadZone.classList.remove('hidden');
        });

        invertImageDownloadBtn.addEventListener('click', () => {
            const link = document.createElement('a');
            link.download = 'inverted_image.png';
            link.href = invertImageCanvas.toDataURL('image/png');
            link.click();
        });
    }

    function processInvertImage(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            const img = new Image();
            img.onload = () => {
                invertImageCanvas.width = img.width;
                invertImageCanvas.height = img.height;
                const ctx = invertImageCanvas.getContext('2d');
                ctx.drawImage(img, 0, 0);
                
                const imageData = ctx.getImageData(0, 0, invertImageCanvas.width, invertImageCanvas.height);
                const data = imageData.data;
                for (let i = 0; i < data.length; i += 4) {
                    data[i] = 255 - data[i];       // red
                    data[i + 1] = 255 - data[i + 1]; // green
                    data[i + 2] = 255 - data[i + 2]; // blue
                }
                ctx.putImageData(imageData, 0, 0);
                
                invertImageUploadZone.classList.add('hidden');
                invertImagePreviewSection.classList.remove('hidden');
            };
            img.src = e.target.result;
        };
        reader.readAsDataURL(file);
    }

    // --- PDF to Text ---
    const pdfTextUploadZone = document.getElementById('pdf-text-upload-zone');
    const pdfTextFileInput = document.getElementById('pdf-text-file-input');
    const pdfTextPreviewSection = document.getElementById('pdf-text-preview-section');
    const pdfTextOutput = document.getElementById('pdf-text-output');
    const pdfTextClearBtn = document.getElementById('pdf-text-clear-btn');
    const pdfTextCopyBtn = document.getElementById('pdf-text-copy-btn');

    if (pdfTextUploadZone) {
        pdfTextUploadZone.addEventListener('click', (e) => {
            if (e.target !== pdfTextFileInput && !e.target.closest('label')) {
                pdfTextFileInput.click();
            }
        });

        pdfTextFileInput.addEventListener('change', async (e) => {
            if (e.target.files.length) await extractTextFromPDFFile(e.target.files[0]);
            pdfTextFileInput.value = '';
        });

        pdfTextUploadZone.addEventListener('dragover', (e) => { e.preventDefault(); pdfTextUploadZone.classList.add('drag-active'); });
        ['dragleave', 'dragend'].forEach(type => { pdfTextUploadZone.addEventListener(type, () => pdfTextUploadZone.classList.remove('drag-active')); });
        pdfTextUploadZone.addEventListener('drop', async (e) => {
            e.preventDefault();
            pdfTextUploadZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) await extractTextFromPDFFile(e.dataTransfer.files[0]);
        });

        pdfTextClearBtn.addEventListener('click', () => {
            pdfTextOutput.value = '';
            pdfTextPreviewSection.classList.add('hidden');
            pdfTextUploadZone.classList.remove('hidden');
        });

        pdfTextCopyBtn.addEventListener('click', () => {
            pdfTextOutput.select();
            document.execCommand('copy');
            alert('Text copied to clipboard!');
        });
    }

    async function extractTextFromPDFFile(file) {
        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Extracting Text from PDF...';
        try {
            const arrayBuffer = await file.arrayBuffer();
            const pdfjsLib = window['pdfjs-dist/build/pdf'];
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
            
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            let fullText = '';
            for (let i = 1; i <= pdf.numPages; i++) {
                const page = await pdf.getPage(i);
                const textContent = await page.getTextContent();
                const pageText = textContent.items.map(item => item.str).join(' ');
                fullText += pageText + '\n\n';
            }
            pdfTextOutput.value = fullText.trim() || 'No readable text found in PDF.';
            pdfTextUploadZone.classList.add('hidden');
            pdfTextPreviewSection.classList.remove('hidden');
        } catch (err) {
            console.error('PDF Extract Error:', err);
            alert('Failed to extract text from PDF.');
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }

    // --- Word to Text ---
    const wordTextUploadZone = document.getElementById('word-text-upload-zone');
    const wordTextFileInput = document.getElementById('word-text-file-input');
    const wordTextPreviewSection = document.getElementById('word-text-preview-section');
    const wordTextOutput = document.getElementById('word-text-output');
    const wordTextClearBtn = document.getElementById('word-text-clear-btn');
    const wordTextCopyBtn = document.getElementById('word-text-copy-btn');

    if (wordTextUploadZone) {
        wordTextUploadZone.addEventListener('click', (e) => {
            if (e.target !== wordTextFileInput && !e.target.closest('label')) {
                wordTextFileInput.click();
            }
        });

        wordTextFileInput.addEventListener('change', async (e) => {
            if (e.target.files.length) await extractTextFromWordFile(e.target.files[0]);
            wordTextFileInput.value = '';
        });

        wordTextUploadZone.addEventListener('dragover', (e) => { e.preventDefault(); wordTextUploadZone.classList.add('drag-active'); });
        ['dragleave', 'dragend'].forEach(type => { wordTextUploadZone.addEventListener(type, () => wordTextUploadZone.classList.remove('drag-active')); });
        wordTextUploadZone.addEventListener('drop', async (e) => {
            e.preventDefault();
            wordTextUploadZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) await extractTextFromWordFile(e.dataTransfer.files[0]);
        });

        wordTextClearBtn.addEventListener('click', () => {
            wordTextOutput.value = '';
            wordTextPreviewSection.classList.add('hidden');
            wordTextUploadZone.classList.remove('hidden');
        });

        wordTextCopyBtn.addEventListener('click', () => {
            wordTextOutput.select();
            document.execCommand('copy');
            alert('Text copied to clipboard!');
        });
    }

    async function extractTextFromWordFile(file) {
        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Extracting Text from Word...';
        try {
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
            wordTextOutput.value = result.value || 'No readable text found in Word document.';
            wordTextUploadZone.classList.add('hidden');
            wordTextPreviewSection.classList.remove('hidden');
        } catch (err) {
            console.error('Word Extract Error:', err);
            alert('Failed to extract text from Word document.');
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }
    // ==========================================
    // ADVANCED PDF TOOLS (MERGE, SPLIT, WATERMARK)
    // ==========================================

    // --- Merge PDF Logic ---
    if (mergePdfUploadZone) {
        mergePdfUploadZone.addEventListener('click', (e) => {
            if (e.target !== mergePdfFileInput && !e.target.closest('label')) {
                mergePdfFileInput.click();
            }
        });

        mergePdfFileInput.addEventListener('change', (e) => {
            handleMergePdfFiles(e.target.files);
            mergePdfFileInput.value = '';
        });

        mergePdfUploadZone.addEventListener('dragover', (e) => { e.preventDefault(); mergePdfUploadZone.classList.add('drag-active'); });
        ['dragleave', 'dragend'].forEach(type => { mergePdfUploadZone.addEventListener(type, () => mergePdfUploadZone.classList.remove('drag-active')); });
        mergePdfUploadZone.addEventListener('drop', (e) => {
            e.preventDefault();
            mergePdfUploadZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) handleMergePdfFiles(e.dataTransfer.files);
        });

        if (mergePdfAddMoreBtn) {
            mergePdfAddMoreBtn.addEventListener('click', () => {
                mergePdfFileInput.click();
            });
        }

        mergePdfClearBtn.addEventListener('click', () => {
            mergePdfFiles = [];
            mergePdfList.innerHTML = '';
            mergePdfPreviewSection.classList.add('hidden');
            mergePdfUploadZone.classList.remove('hidden');
        });

        mergePdfBtn.addEventListener('click', processMergePdf);

        // Sortable for Merge List
        new Sortable(mergePdfList, {
            animation: 150,
            onEnd: () => {
                const newOrder = Array.from(mergePdfList.children).map(card => card.dataset.id);
                const reordered = [];
                newOrder.forEach(id => {
                    const file = mergePdfFiles.find(item => item.id === id);
                    if (file) reordered.push(file);
                });
                mergePdfFiles = reordered;
            }
        });
    }

    async function handleMergePdfFiles(files) {
        const validFiles = Array.from(files).filter(file => file.type === 'application/pdf');
        if (validFiles.length === 0) return;

        for (const file of validFiles) {
            const arrayBuffer = await file.arrayBuffer();
            const id = 'pdf_' + Math.random().toString(36).substr(2, 9);
            mergePdfFiles.push({ id, file, arrayBuffer });

            const card = document.createElement('div');
            card.className = 'pdf-list-item';
            card.dataset.id = id;
            card.style.display = 'flex';
            card.style.alignItems = 'center';
            card.style.justifyContent = 'space-between';
            card.style.padding = '10px 15px';
            card.style.background = '#f8fafc';
            card.style.border = '1px solid #e2e8f0';
            card.style.borderRadius = '8px';
            card.style.cursor = 'grab';

            card.innerHTML = `
                <div style="display: flex; align-items: center; gap: 10px;">
                    <i class="fa-solid fa-grip-vertical" style="color: #94a3b8;"></i>
                    <i class="fa-solid fa-file-pdf" style="color: #ef4444; font-size: 1.2rem;"></i>
                    <span style="font-weight: 500; font-size: 0.9rem; color: #1e293b;">${file.name}</span>
                </div>
                <button class="remove-btn" style="position: relative; top: 0; right: 0; transform: none; width: 28px; height: 28px; background: #fee2e2; color: #ef4444;" onclick="removeMergePdf('${id}', event)">
                    <i class="fa-solid fa-xmark"></i>
                </button>
            `;
            mergePdfList.appendChild(card);
        }

        mergePdfCount.textContent = `${mergePdfFiles.length} file${mergePdfFiles.length !== 1 ? 's' : ''}`;
        mergePdfUploadZone.classList.add('hidden');
        mergePdfPreviewSection.classList.remove('hidden');
    }

    window.removeMergePdf = function (id, event) {
        event.stopPropagation();
        mergePdfFiles = mergePdfFiles.filter(item => item.id !== id);
        const card = mergePdfList.querySelector(`[data-id="${id}"]`);
        if (card) card.remove();
        mergePdfCount.textContent = `${mergePdfFiles.length} file${mergePdfFiles.length !== 1 ? 's' : ''}`;
        if (mergePdfFiles.length === 0) {
            mergePdfPreviewSection.classList.add('hidden');
            mergePdfUploadZone.classList.remove('hidden');
        }
    };

    async function processMergePdf() {
        if (mergePdfFiles.length < 2) {
            alert("Please select at least 2 PDFs to merge.");
            return;
        }

        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Merging PDFs...';

        try {
            const { PDFDocument } = PDFLib;
            const mergedPdf = await PDFDocument.create();

            for (const item of mergePdfFiles) {
                const pdf = await PDFDocument.load(item.arrayBuffer);
                const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
                copiedPages.forEach((page) => mergedPdf.addPage(page));
            }

            const pdfBytes = await mergedPdf.save();
            const blob = new Blob([pdfBytes], { type: 'application/pdf' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            let filename = document.getElementById('merge-pdf-output-filename').value.trim() || 'Merged_Document';
            if (!filename.toLowerCase().endsWith('.pdf')) filename += '.pdf';
            link.download = filename;
            link.click();
        } catch (error) {
            console.error("Merge PDF Error:", error);
            alert("Failed to merge PDFs. Please check if files are valid.");
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }

    // --- Split PDF Logic ---
    if (splitPdfUploadZone) {
        splitPdfUploadZone.addEventListener('click', (e) => {
            if (e.target !== splitPdfFileInput && !e.target.closest('label')) splitPdfFileInput.click();
        });

        splitPdfFileInput.addEventListener('change', async (e) => {
            if (e.target.files.length) await handleSplitPdfFile(e.target.files[0]);
            splitPdfFileInput.value = '';
        });

        splitPdfUploadZone.addEventListener('dragover', (e) => { e.preventDefault(); splitPdfUploadZone.classList.add('drag-active'); });
        ['dragleave', 'dragend'].forEach(type => { splitPdfUploadZone.addEventListener(type, () => splitPdfUploadZone.classList.remove('drag-active')); });
        splitPdfUploadZone.addEventListener('drop', async (e) => {
            e.preventDefault();
            splitPdfUploadZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) await handleSplitPdfFile(e.dataTransfer.files[0]);
        });

        splitPdfClearBtn.addEventListener('click', () => {
            splitPdfFile = null;
            splitPdfArrayBuffer = null;
            splitPdfRange.value = '';
            splitPdfPreviewSection.classList.add('hidden');
            splitPdfUploadZone.classList.remove('hidden');
        });

        splitPdfBtn.addEventListener('click', processSplitPdf);
    }

    async function handleSplitPdfFile(file) {
        if (file.type !== 'application/pdf') {
            alert('Please select a valid PDF file.');
            return;
        }

        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Loading PDF...';

        try {
            splitPdfFile = file;
            splitPdfArrayBuffer = await file.arrayBuffer();
            
            const { PDFDocument } = PDFLib;
            const pdfDoc = await PDFDocument.load(splitPdfArrayBuffer);
            const numPages = pdfDoc.getPageCount();

            splitPdfFilename.textContent = file.name;
            splitPdfPagesCount.textContent = numPages;

            splitPdfUploadZone.classList.add('hidden');
            splitPdfPreviewSection.classList.remove('hidden');
        } catch (error) {
            console.error("Load Split PDF Error:", error);
            alert("Failed to read PDF.");
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }

    async function processSplitPdf() {
        if (!splitPdfArrayBuffer) return;

        const rangeText = splitPdfRange.value.trim();
        
        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Extracting Pages...';

        try {
            const { PDFDocument } = PDFLib;
            const pdfDoc = await PDFDocument.load(splitPdfArrayBuffer);
            const totalPages = pdfDoc.getPageCount();

            let pagesToExtract = [];

            if (!rangeText) {
                const zip = new JSZip();
                for (let i = 0; i < totalPages; i++) {
                    const newPdf = await PDFDocument.create();
                    const [copiedPage] = await newPdf.copyPages(pdfDoc, [i]);
                    newPdf.addPage(copiedPage);
                    const pdfBytes = await newPdf.save();
                    zip.file(`Page_${i + 1}.pdf`, pdfBytes);
                }
                const zipContent = await zip.generateAsync({ type: 'blob' });
                const link = document.createElement('a');
                link.href = URL.createObjectURL(zipContent);
                const baseName = splitPdfFile.name.substring(0, splitPdfFile.name.lastIndexOf('.')) || splitPdfFile.name;
                link.download = `${baseName}_split.zip`;
                link.click();
                loadingOverlay.classList.add('hidden');
                return;
            }

            const parts = rangeText.split(',');
            for (let part of parts) {
                part = part.trim();
                if (part.includes('-')) {
                    const bounds = part.split('-');
                    let start = parseInt(bounds[0], 10);
                    let end = parseInt(bounds[1], 10);
                    if (isNaN(start) || isNaN(end) || start < 1 || end > totalPages || start > end) {
                        throw new Error("Invalid range");
                    }
                    for (let i = start; i <= end; i++) pagesToExtract.push(i - 1);
                } else {
                    let pageNum = parseInt(part, 10);
                    if (isNaN(pageNum) || pageNum < 1 || pageNum > totalPages) {
                        throw new Error("Invalid page number");
                    }
                    pagesToExtract.push(pageNum - 1);
                }
            }

            pagesToExtract = [...new Set(pagesToExtract)].sort((a, b) => a - b);

            if (pagesToExtract.length === 0) throw new Error("No valid pages selected");

            const newPdf = await PDFDocument.create();
            const copiedPages = await newPdf.copyPages(pdfDoc, pagesToExtract);
            copiedPages.forEach((page) => newPdf.addPage(page));

            const pdfBytes = await newPdf.save();
            const blob = new Blob([pdfBytes], { type: 'application/pdf' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            const baseName = splitPdfFile.name.substring(0, splitPdfFile.name.lastIndexOf('.')) || splitPdfFile.name;
            link.download = `${baseName}_extracted.pdf`;
            link.click();

        } catch (error) {
            console.error("Split PDF Error:", error);
            alert("Failed to extract pages. Please check your range format (e.g. 1, 3-5).");
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }

    // --- Watermark PDF Logic ---
    if (watermarkPdfUploadZone) {
        watermarkPdfUploadZone.addEventListener('click', (e) => {
            if (e.target !== watermarkPdfFileInput && !e.target.closest('label')) watermarkPdfFileInput.click();
        });

        watermarkPdfFileInput.addEventListener('change', async (e) => {
            if (e.target.files.length) await handleWatermarkPdfFile(e.target.files[0]);
            watermarkPdfFileInput.value = '';
        });

        watermarkPdfUploadZone.addEventListener('dragover', (e) => { e.preventDefault(); watermarkPdfUploadZone.classList.add('drag-active'); });
        ['dragleave', 'dragend'].forEach(type => { watermarkPdfUploadZone.addEventListener(type, () => watermarkPdfUploadZone.classList.remove('drag-active')); });
        watermarkPdfUploadZone.addEventListener('drop', async (e) => {
            e.preventDefault();
            watermarkPdfUploadZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) await handleWatermarkPdfFile(e.dataTransfer.files[0]);
        });

        watermarkPdfClearBtn.addEventListener('click', () => {
            watermarkPdfFile = null;
            watermarkPdfArrayBuffer = null;
            watermarkPdfPreviewSection.classList.add('hidden');
            watermarkPdfUploadZone.classList.remove('hidden');
        });

        watermarkPdfBtn.addEventListener('click', processWatermarkPdf);
    }

    async function handleWatermarkPdfFile(file) {
        if (file.type !== 'application/pdf') {
            alert('Please select a valid PDF file.');
            return;
        }

        watermarkPdfFile = file;
        watermarkPdfArrayBuffer = await file.arrayBuffer();
        watermarkPdfFilename.textContent = file.name;
        
        const baseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
        document.getElementById('watermark-pdf-output-filename').value = baseName + '_watermarked';

        watermarkPdfUploadZone.classList.add('hidden');
        watermarkPdfPreviewSection.classList.remove('hidden');
    }

    async function processWatermarkPdf() {
        if (!watermarkPdfArrayBuffer) return;

        const text = document.getElementById('watermark-text').value.trim() || 'CONFIDENTIAL';
        const colorHex = document.getElementById('watermark-color').value;
        const opacity = parseFloat(document.getElementById('watermark-opacity').value);

        const r = parseInt(colorHex.substr(1, 2), 16) / 255;
        const g = parseInt(colorHex.substr(3, 2), 16) / 255;
        const b = parseInt(colorHex.substr(5, 2), 16) / 255;

        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Adding Watermark...';

        try {
            const { PDFDocument, rgb, degrees } = PDFLib;
            const pdfDoc = await PDFDocument.load(watermarkPdfArrayBuffer);
            const pages = pdfDoc.getPages();

            for (const page of pages) {
                const { width, height } = page.getSize();
                page.drawText(text, {
                    x: width / 4,
                    y: height / 4,
                    size: 60,
                    color: rgb(r, g, b),
                    opacity: opacity,
                    rotate: degrees(45),
                });
            }

            const pdfBytes = await pdfDoc.save();
            const blob = new Blob([pdfBytes], { type: 'application/pdf' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            let filename = document.getElementById('watermark-pdf-output-filename').value.trim() || 'Watermarked_PDF';
            if (!filename.toLowerCase().endsWith('.pdf')) filename += '.pdf';
            link.download = filename;
            link.click();
        } catch (error) {
            console.error("Watermark PDF Error:", error);
            alert("Failed to add watermark.");
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }
    // --- PDF to JPG Logic ---
    const pdfJpgUploadZone = document.getElementById('pdf-jpg-upload-zone');
    const pdfJpgFileInput = document.getElementById('pdf-jpg-file-input');
    const pdfJpgPreviewSection = document.getElementById('pdf-jpg-preview-section');
    const pdfJpgClearBtn = document.getElementById('pdf-jpg-clear-btn');
    const pdfJpgBtn = document.getElementById('pdf-jpg-btn');
    let pdfJpgFile = null;
    let pdfJpgArrayBuffer = null;

    if (pdfJpgUploadZone) {
        pdfJpgUploadZone.addEventListener('click', (e) => {
            if (e.target !== pdfJpgFileInput && !e.target.closest('label')) pdfJpgFileInput.click();
        });

        pdfJpgFileInput.addEventListener('change', async (e) => {
            if (e.target.files.length) await handlePdfJpgFile(e.target.files[0]);
            pdfJpgFileInput.value = '';
        });

        pdfJpgUploadZone.addEventListener('dragover', (e) => { e.preventDefault(); pdfJpgUploadZone.classList.add('drag-active'); });
        ['dragleave', 'dragend'].forEach(type => { pdfJpgUploadZone.addEventListener(type, () => pdfJpgUploadZone.classList.remove('drag-active')); });
        pdfJpgUploadZone.addEventListener('drop', async (e) => {
            e.preventDefault();
            pdfJpgUploadZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) await handlePdfJpgFile(e.dataTransfer.files[0]);
        });

        pdfJpgClearBtn.addEventListener('click', () => {
            pdfJpgFile = null;
            pdfJpgArrayBuffer = null;
            pdfJpgPreviewSection.classList.add('hidden');
            pdfJpgUploadZone.classList.remove('hidden');
        });

        pdfJpgBtn.addEventListener('click', processPdfToJpg);
    }

    async function handlePdfJpgFile(file) {
        if (file.type !== 'application/pdf') return alert('Please select a valid PDF file.');
        pdfJpgFile = file;
        pdfJpgArrayBuffer = await file.arrayBuffer();
        
        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Loading PDF...';
        try {
            const pdfjsLib = window['pdfjs-dist/build/pdf'];
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
            const pdf = await pdfjsLib.getDocument({ data: pdfJpgArrayBuffer.slice(0) }).promise;
            document.getElementById('pdf-jpg-filename').textContent = file.name;
            document.getElementById('pdf-jpg-pages-count').textContent = pdf.numPages;
            pdfJpgUploadZone.classList.add('hidden');
            pdfJpgPreviewSection.classList.remove('hidden');
        } catch (err) {
            console.error(err);
            alert('Failed to read PDF.');
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }

    async function processPdfToJpg() {
        if (!pdfJpgArrayBuffer) return;
        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Converting to JPG...';
        try {
            const pdfjsLib = window['pdfjs-dist/build/pdf'];
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
            const pdf = await pdfjsLib.getDocument({ data: pdfJpgArrayBuffer.slice(0) }).promise;
            
            for (let i = 1; i <= pdf.numPages; i++) {
                loadingOverlay.querySelector('p').textContent = `Processing Page ${i}/${pdf.numPages}...`;
                const page = await pdf.getPage(i);
                const viewport = page.getViewport({ scale: 2.0 });
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                canvas.width = viewport.width;
                canvas.height = viewport.height;
                
                await page.render({ canvasContext: ctx, viewport: viewport }).promise;
                
                const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', 0.95));
                const link = document.createElement('a');
                link.href = URL.createObjectURL(blob);
                const baseName = pdfJpgFile.name.substring(0, pdfJpgFile.name.lastIndexOf('.')) || pdfJpgFile.name;
                link.download = `${baseName}_Page_${i}.jpg`;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                
                // Small delay to help browser handle multiple downloads
                if (i < pdf.numPages) await new Promise(r => setTimeout(r, 400));
            }
        } catch (err) {
            console.error(err);
            alert('Failed to convert PDF to JPG.');
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }

    // --- Rotate PDF Logic ---
    const rotatePdfUploadZone = document.getElementById('rotate-pdf-upload-zone');
    const rotatePdfFileInput = document.getElementById('rotate-pdf-file-input');
    const rotatePdfPreviewSection = document.getElementById('rotate-pdf-preview-section');
    const rotatePdfClearBtn = document.getElementById('rotate-pdf-clear-btn');
    const rotatePdfBtn = document.getElementById('rotate-pdf-btn');
    let rotatePdfFile = null;
    let rotatePdfArrayBuffer = null;

    if (rotatePdfUploadZone) {
        rotatePdfUploadZone.addEventListener('click', (e) => {
            if (e.target !== rotatePdfFileInput && !e.target.closest('label')) rotatePdfFileInput.click();
        });

        rotatePdfFileInput.addEventListener('change', async (e) => {
            if (e.target.files.length) await handleRotatePdfFile(e.target.files[0]);
            rotatePdfFileInput.value = '';
        });

        rotatePdfUploadZone.addEventListener('dragover', (e) => { e.preventDefault(); rotatePdfUploadZone.classList.add('drag-active'); });
        ['dragleave', 'dragend'].forEach(type => { rotatePdfUploadZone.addEventListener(type, () => rotatePdfUploadZone.classList.remove('drag-active')); });
        rotatePdfUploadZone.addEventListener('drop', async (e) => {
            e.preventDefault();
            rotatePdfUploadZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) await handleRotatePdfFile(e.dataTransfer.files[0]);
        });

        rotatePdfClearBtn.addEventListener('click', () => {
            rotatePdfFile = null;
            rotatePdfArrayBuffer = null;
            rotatePdfPreviewSection.classList.add('hidden');
            rotatePdfUploadZone.classList.remove('hidden');
        });

        rotatePdfBtn.addEventListener('click', processRotatePdf);
    }

    async function handleRotatePdfFile(file) {
        if (file.type !== 'application/pdf') return alert('Please select a valid PDF file.');
        rotatePdfFile = file;
        rotatePdfArrayBuffer = await file.arrayBuffer();
        document.getElementById('rotate-pdf-filename').textContent = file.name;
        const baseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
        document.getElementById('rotate-pdf-output-filename').value = baseName + '_rotated';
        rotatePdfUploadZone.classList.add('hidden');
        rotatePdfPreviewSection.classList.remove('hidden');
    }

    async function processRotatePdf() {
        if (!rotatePdfArrayBuffer) return;
        const angle = parseInt(document.getElementById('rotate-angle').value);
        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Rotating PDF...';
        try {
            const { PDFDocument, degrees } = PDFLib;
            const pdfDoc = await PDFDocument.load(rotatePdfArrayBuffer);
            const pages = pdfDoc.getPages();
            for (const page of pages) {
                const currentRotation = page.getRotation().angle;
                page.setRotation(degrees(currentRotation + angle));
            }
            const pdfBytes = await pdfDoc.save();
            const blob = new Blob([pdfBytes], { type: 'application/pdf' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            let filename = document.getElementById('rotate-pdf-output-filename').value.trim() || 'Rotated_PDF';
            if (!filename.toLowerCase().endsWith('.pdf')) filename += '.pdf';
            link.download = filename;
            link.click();
        } catch (error) {
            console.error(error);
            alert("Failed to rotate PDF.");
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }

    // --- Page Numbers Logic ---
    const pageNumbersUploadZone = document.getElementById('page-numbers-upload-zone');
    const pageNumbersFileInput = document.getElementById('page-numbers-file-input');
    const pageNumbersPreviewSection = document.getElementById('page-numbers-preview-section');
    const pageNumbersClearBtn = document.getElementById('page-numbers-clear-btn');
    const pageNumbersBtn = document.getElementById('page-numbers-btn');
    let pageNumbersFile = null;
    let pageNumbersArrayBuffer = null;

    if (pageNumbersUploadZone) {
        pageNumbersUploadZone.addEventListener('click', (e) => {
            if (e.target !== pageNumbersFileInput && !e.target.closest('label')) pageNumbersFileInput.click();
        });

        pageNumbersFileInput.addEventListener('change', async (e) => {
            if (e.target.files.length) await handlePageNumbersFile(e.target.files[0]);
            pageNumbersFileInput.value = '';
        });

        pageNumbersUploadZone.addEventListener('dragover', (e) => { e.preventDefault(); pageNumbersUploadZone.classList.add('drag-active'); });
        ['dragleave', 'dragend'].forEach(type => { pageNumbersUploadZone.addEventListener(type, () => pageNumbersUploadZone.classList.remove('drag-active')); });
        pageNumbersUploadZone.addEventListener('drop', async (e) => {
            e.preventDefault();
            pageNumbersUploadZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) await handlePageNumbersFile(e.dataTransfer.files[0]);
        });

        pageNumbersClearBtn.addEventListener('click', () => {
            pageNumbersFile = null;
            pageNumbersArrayBuffer = null;
            pageNumbersPreviewSection.classList.add('hidden');
            pageNumbersUploadZone.classList.remove('hidden');
        });

        pageNumbersBtn.addEventListener('click', processPageNumbers);
    }

    async function handlePageNumbersFile(file) {
        if (file.type !== 'application/pdf') return alert('Please select a valid PDF file.');
        pageNumbersFile = file;
        pageNumbersArrayBuffer = await file.arrayBuffer();
        document.getElementById('page-numbers-filename').textContent = file.name;
        const baseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
        document.getElementById('page-numbers-output-filename').value = baseName + '_numbered';
        pageNumbersUploadZone.classList.add('hidden');
        pageNumbersPreviewSection.classList.remove('hidden');
    }

    async function processPageNumbers() {
        if (!pageNumbersArrayBuffer) return;
        const position = document.getElementById('page-number-position').value;
        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Adding Page Numbers...';
        try {
            const { PDFDocument, rgb, StandardFonts } = PDFLib;
            const pdfDoc = await PDFDocument.load(pageNumbersArrayBuffer);
            const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
            const pages = pdfDoc.getPages();
            
            for (let i = 0; i < pages.length; i++) {
                const page = pages[i];
                const { width, height } = page.getSize();
                const text = `${i + 1}`;
                const textSize = 12;
                const textWidth = font.widthOfTextAtSize(text, textSize);
                const margin = 30;
                
                let x = margin;
                let y = margin;
                
                if (position.includes('center')) x = (width - textWidth) / 2;
                if (position.includes('right')) x = width - margin - textWidth;
                
                if (position.includes('top')) y = height - margin - textSize;
                
                page.drawText(text, {
                    x, y, size: textSize, font, color: rgb(0, 0, 0)
                });
            }
            const pdfBytes = await pdfDoc.save();
            const blob = new Blob([pdfBytes], { type: 'application/pdf' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            let filename = document.getElementById('page-numbers-output-filename').value.trim() || 'Numbered_PDF';
            if (!filename.toLowerCase().endsWith('.pdf')) filename += '.pdf';
            link.download = filename;
            link.click();
        } catch (error) {
            console.error(error);
            alert("Failed to add page numbers.");
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }

    // --- HTML to PDF Logic ---
    const htmlPdfUploadZone = document.getElementById('html-pdf-upload-zone');
    const htmlPdfFileInput = document.getElementById('html-pdf-file-input');
    const htmlPdfPreviewSection = document.getElementById('html-pdf-preview-section');
    const htmlPdfClearBtn = document.getElementById('html-pdf-clear-btn');
    const htmlPdfBtn = document.getElementById('html-pdf-btn');
    let htmlPdfFile = null;

    if (htmlPdfUploadZone) {
        htmlPdfUploadZone.addEventListener('click', (e) => {
            if (e.target !== htmlPdfFileInput && !e.target.closest('label')) htmlPdfFileInput.click();
        });

        htmlPdfFileInput.addEventListener('change', (e) => {
            if (e.target.files.length) handleHtmlPdfFile(e.target.files[0]);
            htmlPdfFileInput.value = '';
        });

        htmlPdfUploadZone.addEventListener('dragover', (e) => { e.preventDefault(); htmlPdfUploadZone.classList.add('drag-active'); });
        ['dragleave', 'dragend'].forEach(type => { htmlPdfUploadZone.addEventListener(type, () => htmlPdfUploadZone.classList.remove('drag-active')); });
        htmlPdfUploadZone.addEventListener('drop', (e) => {
            e.preventDefault();
            htmlPdfUploadZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) handleHtmlPdfFile(e.dataTransfer.files[0]);
        });

        htmlPdfClearBtn.addEventListener('click', () => {
            htmlPdfFile = null;
            htmlPdfPreviewSection.classList.add('hidden');
            htmlPdfUploadZone.classList.remove('hidden');
        });

        htmlPdfBtn.addEventListener('click', processHtmlPdf);
    }

    function handleHtmlPdfFile(file) {
        if (!file.name.endsWith('.html') && !file.name.endsWith('.htm')) return alert('Please select a valid HTML file.');
        htmlPdfFile = file;
        document.getElementById('html-pdf-filename').textContent = file.name;
        const baseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
        document.getElementById('html-pdf-output-filename').value = baseName + '_converted';
        htmlPdfUploadZone.classList.add('hidden');
        htmlPdfPreviewSection.classList.remove('hidden');
    }

    async function processHtmlPdf() {
        if (!htmlPdfFile) return;
        loadingOverlay.classList.remove('hidden');
        loadingOverlay.querySelector('p').textContent = 'Converting HTML to PDF...';
        try {
            const text = await htmlPdfFile.text();
            
            let finalName = document.getElementById('html-pdf-output-filename').value.trim() || 'Converted_HTML';
            if (!finalName.toLowerCase().endsWith('.pdf')) finalName += '.pdf';

            const opt = {
                margin: 10,
                filename: finalName,
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2, useCORS: true },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
            };

            const container = document.createElement('div');
            container.innerHTML = text;
            
            await html2pdf().set(opt).from(container).save();
        } catch (error) {
            console.error(error);
            alert("Failed to convert HTML to PDF.");
        } finally {
            loadingOverlay.classList.add('hidden');
        }
    }
});
