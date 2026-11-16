window.mnshor = window.mnshor || {};

if (window.pdfjsLib && !window.pdfjsLib.GlobalWorkerOptions.workerSrc) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = '/js/pdf.worker.min.mjs';
}

window.mnshor._pdfDoc = null;
window.mnshor._pdfLoadingTask = null;
window.mnshor._thumbCache = new Map();
window.mnshor._renderingPages = new Set();
window.mnshor._intersectionObserver = null;
window.mnshor._mutationObserver = null;

window.mnshor.downloadFileFromByteArray = function (fileName, contentType, byteArray) {
    const blob = new Blob([byteArray], { type: contentType });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = fileName ?? 'download';
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    setTimeout(() => URL.revokeObjectURL(url), 15000);
};

window.mnshor.setDocumentDirection = function (direction) {
    document.documentElement.setAttribute('dir', direction);
    document.documentElement.setAttribute('lang', direction === 'rtl' ? 'ar' : 'en');
};

window.mnshor.toggleFullscreen = function (elementId) {
    const el = document.getElementById(elementId);
    if (!el) return;
    if (!document.fullscreenElement) {
        if (el.requestFullscreen) el.requestFullscreen();
    } else {
        if (document.exitFullscreen) document.exitFullscreen();
    }
};

window.mnshor.disposePdfPreview = function () {
    if (window.mnshor._intersectionObserver) {
        window.mnshor._intersectionObserver.disconnect();
        window.mnshor._intersectionObserver = null;
    }
    if (window.mnshor._mutationObserver) {
        window.mnshor._mutationObserver.disconnect();
        window.mnshor._mutationObserver = null;
    }
    const loadingTask = window.mnshor._pdfLoadingTask
        || (window.mnshor._pdfDoc && window.mnshor._pdfDoc.loadingTask)
        || null;
    window.mnshor._pdfLoadingTask = null;
    window.mnshor._pdfDoc = null;
    if (loadingTask && typeof loadingTask.destroy === 'function') {
        loadingTask.destroy().catch(() => { });
    }
    window.mnshor._thumbCache.clear();
    window.mnshor._renderingPages.clear();
};

window.mnshor._applyCachedThumbnail = function (holder, pageNum) {
    if (!window.mnshor._thumbCache.has(pageNum)) return false;
    const dataUrl = window.mnshor._thumbCache.get(pageNum);
    const img = holder.querySelector('img.page-thumbnail-canvas');
    if (img && img.getAttribute('src') !== dataUrl) {
        img.setAttribute('src', dataUrl);
    }
    holder.setAttribute('data-loaded', 'true');
    return true;
};

window.mnshor._renderSinglePage = async function (pageNum) {
    if (!window.mnshor._pdfDoc) return;
    if (window.mnshor._thumbCache.has(pageNum) || window.mnshor._renderingPages.has(pageNum)) {
        return;
    }

    window.mnshor._renderingPages.add(pageNum);
    try {
        const page = await window.mnshor._pdfDoc.getPage(pageNum);
        const baseViewport = page.getViewport({ scale: 1.0 });
        const targetWidth = 185;
        const scale = targetWidth / baseViewport.width;
        const viewport = page.getViewport({ scale: scale });

        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;

        const context = canvas.getContext('2d');
        await page.render({
            canvas: canvas,
            canvasContext: context,
            viewport: viewport
        }).promise;

        const dataUrl = canvas.toDataURL('image/jpeg', 0.8);
        window.mnshor._thumbCache.set(pageNum, dataUrl);

        // Update all matching holders currently in the DOM
        const holders = document.querySelectorAll(`.preview-canvas-holder[data-pdf-page="${pageNum}"]`);
        holders.forEach(h => window.mnshor._applyCachedThumbnail(h, pageNum));
    } catch (err) {
        console.warn(`Failed rendering page ${pageNum}:`, err);
    } finally {
        window.mnshor._renderingPages.delete(pageNum);
    }
};

window.mnshor.syncLazyThumbnails = function (containerId) {
    const container = document.getElementById(containerId || 'pdf-thumbnails-grid');
    if (!container) return;

    const holders = container.querySelectorAll('.preview-canvas-holder[data-pdf-page]');
    holders.forEach(holder => {
        const pageNum = parseInt(holder.getAttribute('data-pdf-page'), 10);
        if (!pageNum || isNaN(pageNum)) return;

        if (!window.mnshor._applyCachedThumbnail(holder, pageNum)) {
            holder.removeAttribute('data-loaded');
            if (window.mnshor._intersectionObserver) {
                window.mnshor._intersectionObserver.observe(holder);
            }
        }
    });
};

window.mnshor.initLazyPdfPreview = async function (containerId, byteArray) {
    if (!window.pdfjsLib) {
        const ready = await new Promise((resolve) => {
            const timeout = setTimeout(() => resolve(false), 10000);
            window.addEventListener('mnshor-pdfjs-ready', () => {
                clearTimeout(timeout);
                resolve(!!window.pdfjsLib);
            }, { once: true });
        });
        if (!ready) {
            console.error('PDF.js library is not loaded');
            return;
        }
    }

    if (!window.pdfjsLib.GlobalWorkerOptions.workerSrc) {
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = '/js/pdf.worker.min.mjs';
    }

    window.mnshor.disposePdfPreview();

    try {
        const uint8 = new Uint8Array(byteArray);
        const loadingTask = window.pdfjsLib.getDocument({
            data: uint8,
            wasmUrl: window.mnshor.pdfWasmUrl || '/js/pdfjs/wasm/'
        });
        window.mnshor._pdfLoadingTask = loadingTask;
        window.mnshor._pdfDoc = await loadingTask.promise;

        window.mnshor._intersectionObserver = new IntersectionObserver((entries) => {
            entries.forEach(entry => {
                if (entry.isIntersecting) {
                    const holder = entry.target;
                    const pageNum = parseInt(holder.getAttribute('data-pdf-page'), 10);
                    if (pageNum && !isNaN(pageNum)) {
                        if (!window.mnshor._applyCachedThumbnail(holder, pageNum)) {
                            window.mnshor._renderSinglePage(pageNum);
                        }
                    }
                    window.mnshor._intersectionObserver.unobserve(holder);
                }
            });
        }, {
            root: null,
            rootMargin: '400px 0px',
            threshold: 0.01
        });

        const container = document.getElementById(containerId || 'pdf-thumbnails-grid');
        if (container) {
            window.mnshor._mutationObserver = new MutationObserver(() => {
                window.mnshor.syncLazyThumbnails(containerId);
            });
            window.mnshor._mutationObserver.observe(container, {
                childList: true,
                subtree: true,
                attributes: true,
                attributeFilter: ['data-pdf-page']
            });
        }

        window.mnshor.syncLazyThumbnails(containerId);
    } catch (err) {
        console.error('Error initializing lazy PDF preview:', err);
    }
};

window.mnshor.scrollToCard = function (cardElementId) {
    setTimeout(() => {
        const el = document.getElementById(cardElementId);
        if (el) {
            el.scrollIntoView({ behavior: 'smooth', block: 'center' });
            el.classList.add('card-jump-highlight');
            setTimeout(() => el.classList.remove('card-jump-highlight'), 1800);
        }
    }, 60);
};
