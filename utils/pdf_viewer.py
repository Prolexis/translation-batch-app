"""
Módulo de visualización interactiva de PDF usando Mozilla PDF.js y Canvas HTML5.
Evita el bloqueo de seguridad de Google Chrome (ERR_BLOCKED_BY_CLIENT / data:application/pdf en iframe).
"""

import json

def get_pdf_js_viewer_html(b64_pdf: str, filename: str = "documento.pdf") -> str:
    """
    Genera el HTML autocontenido con PDF.js para renderizar el documento PDF
    directamente en elementos Canvas de HTML5.
    
    Compatible con todos los navegadores (Chrome, Edge, Firefox, Safari, Móviles)
    sin riesgo de bloqueo por plugins o políticas de iframe.
    """
    # Escapamos el nombre del archivo para JS
    safe_filename = json.dumps(filename)
    
    html_code = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Visor PDF Interactivo</title>
  <!-- PDF.js CDN con fallback -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
  <style>
    :root {{
      --bg-color: #0f172a;
      --bar-bg: #1e293b;
      --bar-border: #334155;
      --text-main: #f8fafc;
      --text-muted: #94a3b8;
      --btn-bg: #334155;
      --btn-hover: #475569;
      --btn-active: #2563eb;
      --accent: #3b82f6;
      --paper-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.5), 0 8px 10px -6px rgba(0, 0, 0, 0.5);
    }}

    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    }}

    body {{
      background: var(--bg-color);
      color: var(--text-main);
      overflow: hidden;
      height: 100vh;
      display: flex;
      flex-direction: column;
    }}

    /* Barra de herramientas superior */
    .toolbar {{
      background: var(--bar-bg);
      border-bottom: 1px solid var(--bar-border);
      padding: 8px 14px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 8px;
      z-index: 10;
      user-select: none;
    }}

    .toolbar-group {{
      display: flex;
      align-items: center;
      gap: 6px;
    }}

    .btn {{
      background: var(--btn-bg);
      color: var(--text-main);
      border: 1px solid var(--bar-border);
      padding: 6px 10px;
      border-radius: 6px;
      cursor: pointer;
      font-size: 13px;
      font-weight: 500;
      display: inline-flex;
      align-items: center;
      gap: 4px;
      transition: all 0.15s ease;
    }}

    .btn:hover:not(:disabled) {{
      background: var(--btn-hover);
      border-color: #64748b;
    }}

    .btn:disabled {{
      opacity: 0.4;
      cursor: not-allowed;
    }}

    .btn.active {{
      background: var(--btn-active);
      border-color: var(--btn-active);
    }}

    .page-input {{
      width: 44px;
      background: #0f172a;
      border: 1px solid var(--bar-border);
      color: var(--text-main);
      padding: 4px 6px;
      border-radius: 4px;
      text-align: center;
      font-size: 13px;
    }}

    .page-count {{
      font-size: 13px;
      color: var(--text-muted);
    }}

    .badge-zoom {{
      font-size: 12px;
      color: var(--text-muted);
      min-width: 42px;
      text-align: center;
    }}

    /* Contenedor del PDF */
    #viewer-container {{
      flex: 1;
      overflow-y: auto;
      overflow-x: auto;
      padding: 20px;
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 20px;
      background: #0b0f19;
      position: relative;
    }}

    .page-wrapper {{
      position: relative;
      background: white;
      box-shadow: var(--paper-shadow);
      border-radius: 4px;
      transition: transform 0.1s ease-out;
    }}

    canvas {{
      display: block;
      border-radius: 4px;
    }}

    .page-label {{
      position: absolute;
      bottom: -22px;
      left: 50%;
      transform: translateX(-50%);
      font-size: 11px;
      color: var(--text-muted);
      white-space: nowrap;
    }}

    /* Spinner de carga */
    .loading-overlay {{
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(15, 23, 42, 0.85);
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 12px;
      z-index: 50;
      color: var(--text-main);
    }}

    .spinner {{
      width: 36px;
      height: 36px;
      border: 3px solid rgba(255, 255, 255, 0.15);
      border-top-color: var(--accent);
      border-radius: 50%;
      animation: spin 0.8s linear infinite;
    }}

    @keyframes spin {{
      to {{ transform: rotate(360deg); }}
    }}

    /* Mensaje de error / Fallback */
    .error-card {{
      background: #1e293b;
      border: 1px solid #ef4444;
      border-radius: 8px;
      padding: 20px;
      max-width: 500px;
      text-align: center;
      margin-top: 40px;
    }}
    .error-card h4 {{ color: #f87171; margin-bottom: 8px; }}
    .error-card p {{ color: #cbd5e1; font-size: 13px; line-height: 1.5; margin-bottom: 14px; }}
  </style>
</head>
<body>

  <!-- Barra de control interactiva -->
  <div class="toolbar">
    <div class="toolbar-group">
      <button id="prev-btn" class="btn" title="Página anterior">◀ Anterior</button>
      <input type="number" id="page-num" class="page-input" value="1" min="1">
      <span class="page-count">/ <span id="page-count">1</span></span>
      <button id="next-btn" class="btn" title="Página siguiente">Siguiente ▶</button>
    </div>

    <div class="toolbar-group">
      <button id="zoom-out" class="btn" title="Reducir zoom">🔍 -</button>
      <span id="zoom-val" class="badge-zoom">100%</span>
      <button id="zoom-in" class="btn" title="Aumentar zoom">🔍 +</button>
      <button id="zoom-fit" class="btn" title="Ajustar al ancho">Ajustar ancho</button>
    </div>

    <div class="toolbar-group">
      <button id="mode-toggle" class="btn active" title="Alternar entre ver todo o página a página">📜 Todas las Págs</button>
      <button id="btn-open-tab" class="btn" title="Abrir en pestaña nueva">↗ Ventana</button>
      <button id="btn-download" class="btn" style="background:#2563eb; border-color:#3b82f6;" title="Descargar archivo PDF">⬇ Descargar</button>
    </div>
  </div>

  <!-- Contenedor del documento -->
  <div id="viewer-container">
    <div id="loading" class="loading-overlay">
      <div class="spinner"></div>
      <div id="loading-text" style="font-size:14px; font-weight:500;">Cargando documento PDF...</div>
    </div>
  </div>

  <script>
    const PDF_BASE64 = "{b64_pdf}";
    const FILE_NAME = {safe_filename};

    // Fallback script loader si CDN principal no responde
    function ensurePdfJsLoaded(callback) {{
      if (typeof pdfjsLib !== 'undefined') {{
        callback();
        return;
      }}
      const fallbackScript = document.createElement('script');
      fallbackScript.src = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js';
      fallbackScript.onload = callback;
      fallbackScript.onerror = function() {{
        showFallbackError("No se pudo conectar con el motor de renderizado PDF.js");
      }};
      document.head.appendChild(fallbackScript);
    }}

    let pdfDoc = null;
    let currentPage = 1;
    let currentScale = 1.15;
    let isContinuous = true;
    let renderedCanvases = [];
    let pdfBlobUrl = null;

    function b64ToUint8Array(b64) {{
      const bin = atob(b64);
      const len = bin.length;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i++) {{
        bytes[i] = bin.charCodeAt(i);
      }}
      return bytes;
    }}

    function showFallbackError(msg) {{
      const container = document.getElementById('viewer-container');
      const loading = document.getElementById('loading');
      if (loading) loading.style.display = 'none';

      container.innerHTML = `
        <div class="error-card">
          <h4>⚠️ Visor en Línea No Disponible</h4>
          <p>${{msg}}.<br>Puedes descargar el PDF maquetado directamente o consultar la pestaña <strong>Hoja de Paper Científico</strong>.</p>
          <button class="btn" style="background:#2563eb; padding:8px 16px;" onclick="triggerDownload()">⬇️ Descargar ${{FILE_NAME}}</button>
        </div>
      `;
    }}

    function triggerDownload() {{
      if (!pdfBlobUrl) {{
        const bytes = b64ToUint8Array(PDF_BASE64);
        const blob = new Blob([bytes], {{ type: 'application/pdf' }});
        pdfBlobUrl = URL.createObjectURL(blob);
      }}
      const a = document.createElement('a');
      a.href = pdfBlobUrl;
      a.download = FILE_NAME;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }}

    function openInNewTab() {{
      if (!pdfBlobUrl) {{
        const bytes = b64ToUint8Array(PDF_BASE64);
        const blob = new Blob([bytes], {{ type: 'application/pdf' }});
        pdfBlobUrl = URL.createObjectURL(blob);
      }}
      window.open(pdfBlobUrl, '_blank');
    }}

    async function initPdf() {{
      try {{
        if (typeof pdfjsLib === 'undefined') {{
          showFallbackError("Biblioteca PDF.js no encontrada");
          return;
        }}
        
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        
        const pdfBytes = b64ToUint8Array(PDF_BASE64);
        const loadingTask = pdfjsLib.getDocument({{ data: pdfBytes }});
        
        pdfDoc = await loadingTask.promise;
        
        document.getElementById('page-count').textContent = pdfDoc.numPages;
        document.getElementById('page-num').max = pdfDoc.numPages;
        
        // Ocultar overlay de carga
        const loading = document.getElementById('loading');
        if (loading) loading.style.display = 'none';

        renderView();
      }} catch (err) {{
        console.error("Error al procesar PDF:", err);
        showFallbackError("Error al interpretar el archivo PDF: " + err.message);
      }}
    }}

    async function renderPage(pageNum, container) {{
      const page = await pdfDoc.getPage(pageNum);
      const viewport = page.getViewport({{ scale: currentScale }});
      const pixelRatio = window.devicePixelRatio || 1;

      const pageWrapper = document.createElement('div');
      pageWrapper.className = 'page-wrapper';
      pageWrapper.id = `page-wrap-${{pageNum}}`;

      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');

      canvas.width = Math.floor(viewport.width * pixelRatio);
      canvas.height = Math.floor(viewport.height * pixelRatio);
      canvas.style.width = Math.floor(viewport.width) + 'px';
      canvas.style.height = Math.floor(viewport.height) + 'px';

      const renderContext = {{
        canvasContext: ctx,
        transform: pixelRatio !== 1 ? [pixelRatio, 0, 0, pixelRatio, 0, 0] : null,
        viewport: viewport
      }};

      pageWrapper.appendChild(canvas);

      if (isContinuous && pdfDoc.numPages > 1) {{
        const label = document.createElement('span');
        label.className = 'page-label';
        label.textContent = `Página ${{pageNum}} de ${{pdfDoc.numPages}}`;
        pageWrapper.appendChild(label);
      }}

      container.appendChild(pageWrapper);
      await page.render(renderContext).promise;
    }}

    async function renderView() {{
      const container = document.getElementById('viewer-container');
      container.querySelectorAll('.page-wrapper').forEach(e => e.remove());

      if (isContinuous) {{
        for (let i = 1; i <= pdfDoc.numPages; i++) {{
          await renderPage(i, container);
        }}
      }} else {{
        await renderPage(currentPage, container);
      }}

      updateControls();
    }}

    function updateControls() {{
      document.getElementById('page-num').value = currentPage;
      document.getElementById('zoom-val').textContent = Math.round(currentScale * 100) + '%';
      
      const prevBtn = document.getElementById('prev-btn');
      const nextBtn = document.getElementById('next-btn');

      if (isContinuous) {{
        prevBtn.disabled = true;
        nextBtn.disabled = true;
      }} else {{
        prevBtn.disabled = (currentPage <= 1);
        nextBtn.disabled = (currentPage >= pdfDoc.numPages);
      }}
    }}

    // Event Listeners
    document.getElementById('prev-btn').addEventListener('click', () => {{
      if (!isContinuous && currentPage > 1) {{
        currentPage--;
        renderView();
      }}
    }});

    document.getElementById('next-btn').addEventListener('click', () => {{
      if (!isContinuous && currentPage < pdfDoc.numPages) {{
        currentPage++;
        renderView();
      }}
    }});

    document.getElementById('page-num').addEventListener('change', (e) => {{
      let val = parseInt(e.target.value, 10);
      if (isNaN(val)) val = 1;
      if (val < 1) val = 1;
      if (val > pdfDoc.numPages) val = pdfDoc.numPages;
      currentPage = val;

      if (isContinuous) {{
        const target = document.getElementById(`page-wrap-${{currentPage}}`);
        if (target) target.scrollIntoView({{ behavior: 'smooth' }});
      }} else {{
        renderView();
      }}
    }});

    document.getElementById('zoom-in').addEventListener('click', () => {{
      if (currentScale < 2.5) {{
        currentScale += 0.15;
        renderView();
      }}
    }});

    document.getElementById('zoom-out').addEventListener('click', () => {{
      if (currentScale > 0.5) {{
        currentScale -= 0.15;
        renderView();
      }}
    }});

    document.getElementById('zoom-fit').addEventListener('click', () => {{
      const containerWidth = document.getElementById('viewer-container').clientWidth - 60;
      if (containerWidth > 200 && pdfDoc) {{
        pdfDoc.getPage(1).then(p => {{
          const unscaledVp = p.getViewport({{ scale: 1.0 }});
          currentScale = Math.max(0.6, Math.min(2.0, containerWidth / unscaledVp.width));
          renderView();
        }});
      }}
    }});

    document.getElementById('mode-toggle').addEventListener('click', (e) => {{
      isContinuous = !isContinuous;
      if (isContinuous) {{
        e.target.textContent = '📜 Todas las Págs';
        e.target.classList.add('active');
      }} else {{
        e.target.textContent = '📄 Pág individual';
        e.target.classList.remove('active');
      }}
      renderView();
    }});

    document.getElementById('btn-download').addEventListener('click', triggerDownload);
    document.getElementById('btn-open-tab').addEventListener('click', openInNewTab);

    // Inicialización al cargar la ventana
    window.addEventListener('DOMContentLoaded', () => {{
      ensurePdfJsLoaded(initPdf);
    }});
  </script>
</body>
</html>"""
    return html_code
