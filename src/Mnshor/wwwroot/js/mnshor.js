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

window.mnshor._promoPdfCache = window.mnshor._promoPdfCache || {};
window.mnshor.renderPromoPdfPage = async function (pdfUrl, canvasId) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    if (!window.pdfjsLib) {
        await new Promise((resolve) => {
            const onReady = () => { window.removeEventListener('mnshor-pdfjs-ready', onReady); resolve(); };
            window.addEventListener('mnshor-pdfjs-ready', onReady);
            setTimeout(() => { window.removeEventListener('mnshor-pdfjs-ready', onReady); resolve(); }, 4000);
        });
    }
    if (!window.pdfjsLib) return;
    try {
        if (!window.pdfjsLib.GlobalWorkerOptions.workerSrc || window.pdfjsLib.GlobalWorkerOptions.workerSrc.endsWith('.js')) {
            window.pdfjsLib.GlobalWorkerOptions.workerSrc = '/js/pdf.worker.min.mjs';
        }
        let pdf = window.mnshor._promoPdfCache[pdfUrl];
        if (!pdf) {
            const resp = await fetch(pdfUrl);
            if (!resp.ok) throw new Error('HTTP ' + resp.status);
            const buf = await resp.arrayBuffer();
            pdf = await window.pdfjsLib.getDocument({
                data: new Uint8Array(buf),
                wasmUrl: window.mnshor.pdfWasmUrl || '/js/pdfjs/wasm/'
            }).promise;
            window.mnshor._promoPdfCache[pdfUrl] = pdf;
        }
        const page = await pdf.getPage(1);
        const baseViewport = page.getViewport({ scale: 1 });
        const targetWidth = canvas.parentElement ? Math.max(240, canvas.parentElement.clientWidth) : 340;
        const scale = (targetWidth / baseViewport.width) * (window.devicePixelRatio || 1.5);
        const viewport = page.getViewport({ scale });
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        const ctx = canvas.getContext('2d');
        await page.render({ canvasContext: ctx, viewport }).promise;
    } catch (e) {
        console.warn('Promo PDF render error:', e);
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

// =========================================================================
// Sign PDF: High-Res Active Page Stage, Freehand Pad & 60fps Drag Placement
// =========================================================================

window.mnshor.renderHighResSignPage = async function (canvasId, pageNum) {
    if (!window.mnshor._pdfDoc) return null;
    const canvas = document.getElementById(canvasId);
    if (!canvas) return null;

    try {
        const page = await window.mnshor._pdfDoc.getPage(pageNum);
        const baseViewport = page.getViewport({ scale: 1.0 });
        const targetWidth = 740;
        const scale = targetWidth / baseViewport.width;
        const viewport = page.getViewport({ scale: scale });

        canvas.width = viewport.width;
        canvas.height = viewport.height;

        const ctx = canvas.getContext('2d');
        await page.render({
            canvas: canvas,
            canvasContext: ctx,
            viewport: viewport
        }).promise;

        return {
            width: baseViewport.width,
            height: baseViewport.height,
            aspect: baseViewport.width / baseViewport.height
        };
    } catch (err) {
        console.warn('Error rendering high-res sign page:', err);
        return null;
    }
};

window.mnshor._sigPadState = {
    canvas: null,
    ctx: null,
    drawing: false,
    points: [],
    hasInk: false,
    inkColor: '#0F172A',
    strokeWidth: 3.2
};

window.mnshor.initSignaturePad = function (canvasId, inkColor, strokeWidth) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    const state = window.mnshor._sigPadState;
    state.canvas = canvas;
    state.ctx = canvas.getContext('2d');
    state.inkColor = inkColor || '#0F172A';
    state.strokeWidth = strokeWidth || 3.2;
    state.drawing = false;
    state.points = [];
    state.hasInk = false;

    const ctx = state.ctx;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';

    const getPos = (e) => {
        const rect = canvas.getBoundingClientRect();
        const scaleX = canvas.width / rect.width;
        const scaleY = canvas.height / rect.height;
        return {
            x: (e.clientX - rect.left) * scaleX,
            y: (e.clientY - rect.top) * scaleY
        };
    };

    canvas.onpointerdown = (e) => {
        e.preventDefault();
        canvas.setPointerCapture(e.pointerId);
        state.drawing = true;
        state.hasInk = true;
        const pt = getPos(e);
        state.points = [pt];

        ctx.strokeStyle = state.inkColor;
        ctx.fillStyle = state.inkColor;
        ctx.lineWidth = state.strokeWidth;

        ctx.beginPath();
        ctx.arc(pt.x, pt.y, state.strokeWidth * 0.45, 0, Math.PI * 2);
        ctx.fill();
    };

    canvas.onpointermove = (e) => {
        if (!state.drawing) return;
        e.preventDefault();
        const pt = getPos(e);
        state.points.push(pt);

        ctx.strokeStyle = state.inkColor;
        ctx.lineWidth = state.strokeWidth;
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';

        const pts = state.points;
        if (pts.length < 3) {
            ctx.beginPath();
            ctx.moveTo(pts[0].x, pts[0].y);
            ctx.lineTo(pt.x, pt.y);
            ctx.stroke();
            return;
        }

        const len = pts.length;
        const p0 = pts[len - 3];
        const p1 = pts[len - 2];
        const p2 = pts[len - 1];
        const mid1 = { x: (p0.x + p1.x) / 2, y: (p0.y + p1.y) / 2 };
        const mid2 = { x: (p1.x + p2.x) / 2, y: (p1.y + p2.y) / 2 };

        ctx.beginPath();
        ctx.moveTo(mid1.x, mid1.y);
        ctx.quadraticCurveTo(p1.x, p1.y, mid2.x, mid2.y);
        ctx.stroke();
    };

    const stopDraw = (e) => {
        if (!state.drawing) return;
        state.drawing = false;
        state.points = [];
    };

    canvas.onpointerup = stopDraw;
    canvas.onpointercancel = stopDraw;
};

window.mnshor.updateSignaturePadStyle = function (inkColor, strokeWidth) {
    const state = window.mnshor._sigPadState;
    if (inkColor) state.inkColor = inkColor;
    if (strokeWidth) state.strokeWidth = strokeWidth;
};

window.mnshor.clearSignaturePad = function (canvasId) {
    const canvas = document.getElementById(canvasId) || window.mnshor._sigPadState.canvas;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    window.mnshor._sigPadState.hasInk = false;
    window.mnshor._sigPadState.points = [];
};

window.mnshor.exportSignaturePadPng = function (canvasId) {
    const canvas = document.getElementById(canvasId) || window.mnshor._sigPadState.canvas;
    if (!canvas) return null;

    const ctx = canvas.getContext('2d');
    const w = canvas.width;
    const h = canvas.height;
    const imgData = ctx.getImageData(0, 0, w, h).data;

    let minX = w, minY = h, maxX = -1, maxY = -1;
    for (let y = 0; y < h; y++) {
        for (let x = 0; x < w; x++) {
            const alpha = imgData[(y * w + x) * 4 + 3];
            if (alpha > 10) {
                if (x < minX) minX = x;
                if (x > maxX) maxX = x;
                if (y < minY) minY = y;
                if (y > maxY) maxY = y;
            }
        }
    }

    if (maxX < minX || maxY < minY) {
        return null; // Pad is empty
    }

    const pad = 16;
    const cropX = Math.max(0, minX - pad);
    const cropY = Math.max(0, minY - pad);
    const cropW = Math.min(w - cropX, (maxX - minX + 1) + pad * 2);
    const cropH = Math.min(h - cropY, (maxY - minY + 1) + pad * 2);

    const outCanvas = document.createElement('canvas');
    outCanvas.width = Math.max(48, cropW);
    outCanvas.height = Math.max(32, cropH);
    const outCtx = outCanvas.getContext('2d');
    outCtx.drawImage(canvas, cropX, cropY, cropW, cropH, 0, 0, cropW, cropH);

    const dataUrl = outCanvas.toDataURL('image/png');
    const base64 = dataUrl.replace(/^data:image\/png;base64,/, '');
    return {
        base64: base64,
        aspectRatio: outCanvas.width / outCanvas.height
    };
};

window.mnshor.startSignFieldDrag = function (clientX, clientY, stageId, fieldDomId, fieldId, initXPct, initYPct, dotNetRef) {
    const stage = document.getElementById(stageId);
    const fieldEl = document.getElementById(fieldDomId);
    if (!stage || !fieldEl) return;

    const stageRect = stage.getBoundingClientRect();
    if (stageRect.width <= 0 || stageRect.height <= 0) return;

    let curXPct = initXPct;
    let curYPct = initYPct;

    const onMove = (e) => {
        const dx = (e.clientX - clientX) / stageRect.width;
        const dy = (e.clientY - clientY) / stageRect.height;

        curXPct = Math.max(0.0, Math.min(0.92, initXPct + dx));
        curYPct = Math.max(0.0, Math.min(0.94, initYPct + dy));

        fieldEl.style.left = (curXPct * 100).toFixed(2) + '%';
        fieldEl.style.top = (curYPct * 100).toFixed(2) + '%';
    };

    const onUp = () => {
        window.removeEventListener('pointermove', onMove);
        window.removeEventListener('pointerup', onUp);
        if (dotNetRef) {
            dotNetRef.invokeMethodAsync('OnFieldDraggedFromJs', fieldId, curXPct, curYPct);
        }
    };

    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
};
