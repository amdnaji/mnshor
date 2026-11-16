import * as pdfjsLib from './pdf.min.mjs';

// Classic scripts (mnshor.js) read this global. pdf.js 4+ ships ESM only.
window.pdfjsLib = pdfjsLib;
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('pdf.worker.min.mjs', import.meta.url).href;
window.mnshor = window.mnshor || {};
window.mnshor.pdfWasmUrl = new URL('pdfjs/wasm/', import.meta.url).href;
window.dispatchEvent(new Event('mnshor-pdfjs-ready'));
