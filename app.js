(function() {
// Config
      const PAPER_CONFIG = {
        letter: { width: 8.5, height: 11, unit: 'in' },
        legal: { width: 8.5, height: 14, unit: 'in' },
        ledger: { width: 11, height: 17, unit: 'in' },
        a4: { width: 210, height: 297, unit: 'mm' },
        b5: { width: 176, height: 250, unit: 'mm' }
      };
      const PPI = 72; const PMM = 2.83465;

      // State
      let state = {
        paper: 'letter', gridSize: 18, unit: 'in', tool: 'select', 
        pages: [{ elements: [] }], currentPageIndex: 0,
        history: [], zoom: 1.0, globalOpacity: 100,
        isDrawing: false, isDragging: false, isModifying: false, isFilling: false, isSelecting: false,
        activeKeys: {}, clipboard: [], editingIndex: null,
        startPoint: null, dragElementIndex: null, selectedIndices: [], modifyHandleType: null,
        gridVisible: true, rulerVisible: true, gridOffset: { x: 0, y: 0 },
        viewMode: 'canvas', bleedUnits: 1, orientation: 'portrait',
        bleedVisible: true,
        notebookLayout: { cols: 1, pageWidth: 280 },
        lastCanvasPageIndex: null,
        notebookClickTimer: null,
        pendingQualityChoice: null
      };
      window.state = state;
      const textMeasureCanvas = document.createElement('canvas');
      const textMeasureCtx = textMeasureCanvas.getContext('2d');

      // --- Internal Helpers ---

      function updateUndoButton() {
        const undoBtn = document.getElementById('undoBtn');
        if (undoBtn) undoBtn.disabled = state.history.length === 0;
      }

      function saveHistory() {
        state.history.push(JSON.stringify(state.pages));
        if (state.history.length > 50) state.history.shift();
        updateUndoButton();
      }

      function expandTextBoxToFit(element, contentNode) {
        if (!element || element.type !== 'text' || !contentNode) return false;
        const computed = window.getComputedStyle(contentNode);
        const paddingTop = parseFloat(computed.paddingTop || '0');
        const paddingBottom = parseFloat(computed.paddingBottom || '0');
        const neededHeight = contentNode.scrollHeight + paddingTop + paddingBottom;
        const snappedHeight = Math.max(state.gridSize, Math.ceil(neededHeight / state.gridSize) * state.gridSize);
        if (snappedHeight > element.h) {
          element.h = snappedHeight;
          return true;
        }
        return false;
      }

      function getPaperMetrics() {
        const config = PAPER_CONFIG[state.paper];
        const baseWidth = state.unit === 'in' ? config.width * PPI : config.width * PMM;
        const baseHeight = state.unit === 'in' ? config.height * PPI : config.height * PMM;
        return state.orientation === 'landscape'
          ? { width: baseHeight, height: baseWidth, unit: config.unit }
          : { width: baseWidth, height: baseHeight, unit: config.unit };
      }

      function getWorkspaceMetrics() {
        const { width, height } = getPaperMetrics();
        const bleedPx = getVisibleBleedPx();
        const rulerOffset = getCanvasRulerOffset();
        return {
          pageWidth: width,
          pageHeight: height,
          bleedPx,
          rulerOffset,
          totalWidth: width + 2 * bleedPx,
          totalHeight: height + 2 * bleedPx,
          framedWidth: width + 2 * bleedPx + rulerOffset + getCanvasFrameExtra(),
          framedHeight: height + 2 * bleedPx + rulerOffset + getCanvasFrameExtra()
        };
      }

      function getTextFontSize(el) {
        return Number(el?.style?.fontSize || el?.fontSize || Math.max(12, Math.round(state.gridSize * 0.7)));
      }

      function getTextFontString(el) {
        const fontStyle = el?.style?.fontStyle || 'normal';
        const fontWeight = el?.style?.fontWeight || '400';
        const fontSize = getTextFontSize(el);
        return `${fontStyle} ${fontWeight} ${fontSize}px Helvetica, Arial, sans-serif`;
      }

      function wrapTextToWidth(text, maxWidth, fontString) {
        if (!textMeasureCtx) return (text || '').split('\n');
        textMeasureCtx.font = fontString;
        const paragraphs = String(text || '').split('\n');
        const lines = [];
        paragraphs.forEach(paragraph => {
          if (paragraph === '') {
            lines.push('');
            return;
          }
          const words = paragraph.split(/\s+/);
          let current = '';
          words.forEach(word => {
            const candidate = current ? `${current} ${word}` : word;
            if (textMeasureCtx.measureText(candidate).width <= maxWidth || !current) {
              current = candidate;
            } else {
              lines.push(current);
              current = word;
            }
          });
          if (current) lines.push(current);
        });
        return lines.length ? lines : [''];
      }

      function updatePageIndicators() { 
        const cp = document.getElementById('currentPageDisplay');
        const tp = document.getElementById('totalPagesDisplay');
        const db = document.getElementById('deletePageBtn');
        if(cp) cp.innerText = state.currentPageIndex + 1; 
        if(tp) tp.innerText = state.pages.length; 
        if(db) db.disabled = state.pages.length <= 1; 
      }

      function updateOrientationButton() {
        const label = document.getElementById('orientationLabel');
        if (label) label.innerText = state.orientation === 'landscape' ? 'Landscape' : 'Portrait';
      }

      function formatPxValue(value) {
        const numeric = Number(value || 0);
        return `${numeric % 1 === 0 ? numeric.toFixed(0) : numeric.toFixed(1)}px`;
      }

      function updatePropsPanel() {
        const panel = document.getElementById('selectionToolbar');
        if (!panel) return;
        const selectedElements = state.pages[state.currentPageIndex].elements.filter((_, i) => state.selectedIndices.includes(i));
        if (selectedElements.length === 0) { panel.classList.add('hidden'); return; }
        panel.classList.remove('hidden');
        panel.classList.add('flex');
        const el = selectedElements[0];
        const allText = selectedElements.every(sel => sel.type === 'text');
        const strokeStyleControls = document.getElementById('strokeStyleControls');
        const weightControl = document.getElementById('weightControl');
        const textAlignControls = document.getElementById('textAlignControls');
        const textVerticalControls = document.getElementById('textVerticalControls');
        const textFormatControls = document.getElementById('textFormatControls');
        const textSizeControl = document.getElementById('textSizeControl');
        document.getElementById('selectionToolbarTitle').innerText = state.selectedIndices.length > 1 ? `Batch (${state.selectedIndices.length})` : `${el.type} Selected`;
        document.getElementById('weightVal').innerText = formatPxValue(el.style.weight || 1);
        document.getElementById('weightSlider').value = el.style.weight || 1;
        document.getElementById('opacityVal').innerText = `${el.style.opacity || 100}%`;
        document.getElementById('opacitySlider').value = el.style.opacity || 100;
        if (strokeStyleControls) strokeStyleControls.classList.toggle('hidden', allText);
        if (weightControl) weightControl.classList.toggle('hidden', allText);
        if (textAlignControls) textAlignControls.classList.toggle('hidden', !allText);
        if (textVerticalControls) textVerticalControls.classList.toggle('hidden', !allText);
        if (textFormatControls) textFormatControls.classList.toggle('hidden', !allText);
        if (textSizeControl) textSizeControl.classList.toggle('hidden', !allText);
        if (allText) {
          const fontSize = el.style.fontSize || el.fontSize || Math.max(12, Math.round(state.gridSize * 0.7));
          document.getElementById('textSizeVal').innerText = formatPxValue(fontSize);
          document.getElementById('textSizeSlider').value = fontSize;
        }
        document.querySelectorAll('.style-btn').forEach(btn => {
          const dash = el.style.dash || 'none';
          const btnType = btn.id.replace('style-', '');
          const activeType = dash === 'none' ? 'none' : (dash === '5,5' ? 'dashed' : (dash === '0,4' ? 'dotted' : 'dashdot'));
          btn.classList.toggle('bg-white', btnType === activeType);
          btn.classList.toggle('ring-1', btnType === activeType);
          btn.classList.toggle('ring-indigo-300', btnType === activeType);
        });
        document.querySelectorAll('.align-btn').forEach(btn => {
          const textAlign = el.style.textAlign || 'left';
          const active = btn.id === `align-${textAlign}`;
          btn.classList.toggle('bg-white', active);
          btn.classList.toggle('ring-1', active);
          btn.classList.toggle('ring-indigo-300', active);
        });
        document.querySelectorAll('.valign-btn').forEach(btn => {
          const verticalAlign = el.style.verticalAlign || 'top';
          const active = btn.id === `valign-${verticalAlign}`;
          btn.classList.toggle('bg-white', active);
          btn.classList.toggle('ring-1', active);
          btn.classList.toggle('ring-indigo-300', active);
        });
        document.querySelectorAll('.text-format-btn').forEach(btn => {
          const active =
            (btn.id === 'font-bold' && (el.style.fontWeight || '400') === '700') ||
            (btn.id === 'font-italic' && (el.style.fontStyle || 'normal') === 'italic');
          btn.classList.toggle('bg-white', active);
          btn.classList.toggle('ring-1', active);
          btn.classList.toggle('ring-indigo-300', active);
        });
      }

      const PDF_QUALITY_PRESETS = [
        { id: 'minimal', label: 'Minimal', scale: 1.25, mime: 'image/jpeg', quality: 0.6, compression: 'FAST' },
        { id: 'small', label: 'Small', scale: 1.5, mime: 'image/jpeg', quality: 0.72, compression: 'MEDIUM' },
        { id: 'standard', label: 'Standard', scale: 2, mime: 'image/jpeg', quality: 0.82, compression: 'MEDIUM' },
        { id: 'high', label: 'High', scale: 2.5, mime: 'image/jpeg', quality: 0.9, compression: 'SLOW' }
      ];

      function estimatePdfSizeLabel(preset) {
        const pageCount = Math.max(state.pages.length, 1);
        const { width: pageWidth, height: pageHeight } = getPaperMetrics();
        const megapixels = (pageWidth * preset.scale * pageHeight * preset.scale) / 1000000;
        const qualityFactor = preset.mime === 'image/png' ? 1.4 : Math.max(0.45, preset.quality);
        const estimatedMb = Math.max(0.2, pageCount * megapixels * qualityFactor * 0.42);
        const rounded = estimatedMb < 1 ? estimatedMb.toFixed(1) : Math.round(estimatedMb);
        return `~${rounded} MB`;
      }

      function renderQualityOptions(selectedId) {
        const container = document.getElementById('qualityOptions');
        if (!container) return;
        container.innerHTML = '';
        PDF_QUALITY_PRESETS.forEach((preset) => {
          const button = document.createElement('button');
          button.type = 'button';
          button.className = `quality-option w-full text-left px-4 py-3 border rounded-xl transition-all hover:border-indigo-300 ${preset.id === selectedId ? 'active' : 'border-slate-200'}`;
          button.innerHTML = `<div class="flex items-center justify-between gap-3"><span class="text-sm font-bold text-slate-800">${preset.label}</span><span class="text-xs font-medium text-slate-400 uppercase tracking-widest">${estimatePdfSizeLabel(preset)}</span></div>`;
          button.onclick = () => {
            state.pendingQualityChoice = preset.id;
            renderQualityOptions(preset.id);
          };
          container.appendChild(button);
        });
      }

      function openQualityModal(defaultId = 'standard') {
        state.pendingQualityChoice = defaultId;
        renderQualityOptions(defaultId);
        const modal = document.getElementById('qualityModal');
        if (modal) modal.classList.remove('hidden');
      }

      window.closeQualityModal = function() {
        const modal = document.getElementById('qualityModal');
        if (modal) modal.classList.add('hidden');
        if (state.qualityModalReject) {
          state.qualityModalReject(new Error('cancelled'));
          state.qualityModalResolve = null;
          state.qualityModalReject = null;
        }
      };

      window.confirmQualityModal = function() {
        const choice = PDF_QUALITY_PRESETS.find((preset) => preset.id === state.pendingQualityChoice) || PDF_QUALITY_PRESETS[2];
        const modal = document.getElementById('qualityModal');
        if (modal) modal.classList.add('hidden');
        if (state.qualityModalResolve) {
          state.qualityModalResolve(choice);
          state.qualityModalResolve = null;
          state.qualityModalReject = null;
        }
      };

      function promptForPdfQuality() {
        return new Promise((resolve, reject) => {
          state.qualityModalResolve = resolve;
          state.qualityModalReject = reject;
          openQualityModal('standard');
        });
      }

      function renderRulerContent(svg, size, orientation) {
        svg.innerHTML = '';
        const unitScale = state.unit === 'in' ? PPI : PMM;
        const bleedPx = getVisibleBleedPx();
        const rulerOffset = getCanvasRulerOffset();
        const step = state.unit === 'in' ? 0.125 * unitScale : 1 * unitScale; 

        for (let i = 0; i <= size + 0.5; i += step) {
          const rel = (i - bleedPx) / unitScale;
          const isInch = state.unit === 'in';
          const isMajor = isInch ? (Math.abs(rel % 1) < 0.01) : (Math.abs(rel % 10) < 0.01);
          const isHalf = isInch ? (Math.abs(rel % 0.5) < 0.01) : (Math.abs(rel % 5) < 0.01);
          const isQuarter = isInch ? (Math.abs(rel % 0.25) < 0.01) : false;

          const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
          let tickSize = isMajor ? 20 : (isHalf ? 12 : (isQuarter ? 8 : 5));
          
          if (orientation === 'horizontal') {
            line.setAttribute("x1", i); line.setAttribute("x2", i);
            line.setAttribute("y1", rulerOffset - tickSize); line.setAttribute("y2", rulerOffset);
          } else {
            line.setAttribute("y1", i); line.setAttribute("y2", i);
            line.setAttribute("x1", rulerOffset - tickSize); line.setAttribute("x2", rulerOffset);
          }
          
          line.setAttribute("class", Math.abs(rel) < 0.001 ? "ruler-center-marker" : "ruler-tick");
          svg.appendChild(line);

          if (isMajor) {
            const text = document.createElementNS("http://www.w3.org/2000/svg", "text");
            text.setAttribute("class", "ruler-label");
            text.textContent = Math.round(rel);
            if (orientation === 'horizontal') {
              text.setAttribute("x", i + 2); text.setAttribute("y", 10);
            } else {
              text.setAttribute("x", 2); text.setAttribute("y", i + 8);
            }
            svg.appendChild(text);
          }
        }
      }

      function renderGridLines(elG, w, h) {
        const step = state.gridSize;
        const bleedPx = state.bleedVisible ? state.bleedUnits * state.gridSize : 0;
        const sheetW = w - 2 * bleedPx;
        const sheetH = h - 2 * bleedPx;
        state.gridOffset = { x: (sheetW / 2 + bleedPx) % step, y: (sheetH / 2 + bleedPx) % step };
        for (let x = state.gridOffset.x; x <= w + 0.1; x += step) {
          const l = document.createElementNS("http://www.w3.org/2000/svg", "line");
          l.setAttribute("x1", x); l.setAttribute("y1", 0); l.setAttribute("x2", x); l.setAttribute("y2", h);
          l.setAttribute("stroke", Math.abs(x - (sheetW / 2 + bleedPx)) < 0.1 ? "#cbd5e1" : "#f1f5f9");
          elG.appendChild(l);
        }
        for (let y = state.gridOffset.y; y <= h + 0.1; y += step) {
          const l = document.createElementNS("http://www.w3.org/2000/svg", "line");
          l.setAttribute("x1", 0); l.setAttribute("y1", y); l.setAttribute("x2", w); l.setAttribute("y2", y);
          l.setAttribute("stroke", Math.abs(y - (sheetH / 2 + bleedPx)) < 0.1 ? "#cbd5e1" : "#f1f5f9");
          elG.appendChild(l);
        }
      }

      function getElementHandles(el, bleedPx) {
        if (el.type === 'line') {
          const x1 = el.x1 + bleedPx, y1 = el.y1 + bleedPx, x2 = el.x2 + bleedPx, y2 = el.y2 + bleedPx;
          return {
            center: { x: (x1 + x2) / 2, y: (y1 + y2) / 2 },
            edges: [{ x: x1, y: y1 }, { x: x2, y: y2 }]
          };
        }
        if (el.type === 'rect' || el.type === 'text') {
          const x = el.x + bleedPx, y = el.y + bleedPx, w = el.w, h = el.h;
          return {
            center: { x: x + w / 2, y: y + h / 2 },
            edges: [
              { x: x + w / 2, y },
              { x: x + w, y: y + h / 2 },
              { x: x + w / 2, y: y + h },
              { x, y: y + h / 2 }
            ]
          };
        }
        const x = el.x + bleedPx, y = el.y + bleedPx;
        return {
          center: { x, y },
          edges: [
            { x, y: y - 12 },
            { x: x + 12, y },
            { x, y: y + 12 },
            { x: x - 12, y }
          ]
        };
      }

      function applyHandleModification(el, handle, point) {
        if (!el || !handle) return;

        if (el.type === 'line') {
          if (handle.kind === 'edge' && handle.index === 0) {
            el.x1 = point.x;
            el.y1 = point.y;
          } else if (handle.kind === 'edge' && handle.index === 1) {
            el.x2 = point.x;
            el.y2 = point.y;
          }
          return;
        }

        if ((el.type === 'rect' || el.type === 'text') && handle.kind === 'edge') {
          const left = el.x;
          const right = el.x + el.w;
          const top = el.y;
          const bottom = el.y + el.h;

          if (handle.index === 0) {
            const nextTop = Math.min(point.y, bottom - state.gridSize / 2);
            el.y = nextTop;
            el.h = bottom - nextTop;
          } else if (handle.index === 1) {
            const nextRight = Math.max(point.x, left + state.gridSize / 2);
            el.w = nextRight - left;
          } else if (handle.index === 2) {
            const nextBottom = Math.max(point.y, top + state.gridSize / 2);
            el.h = nextBottom - top;
          } else if (handle.index === 3) {
            const nextLeft = Math.min(point.x, right - state.gridSize / 2);
            el.x = nextLeft;
            el.w = right - nextLeft;
          }
        }
      }

      function getElementBounds(el) {
        if (el.type === 'line') {
          return {
            minX: Math.min(el.x1, el.x2),
            minY: Math.min(el.y1, el.y2),
            maxX: Math.max(el.x1, el.x2),
            maxY: Math.max(el.y1, el.y2)
          };
        }
        if (el.type === 'rect' || el.type === 'text') {
          return {
            minX: el.x,
            minY: el.y,
            maxX: el.x + el.w,
            maxY: el.y + el.h
          };
        }
        if (el.type === 'cross') {
          return {
            minX: el.x - 6,
            minY: el.y - 6,
            maxX: el.x + 6,
            maxY: el.y + 6
          };
        }
        return {
          minX: el.x - 6,
          minY: el.y - 6,
          maxX: el.x + 6,
          maxY: el.y + 6
        };
      }

      function getMarqueeSelection(rect) {
        return state.pages[state.currentPageIndex].elements
          .map((el, idx) => ({ bounds: getElementBounds(el), idx }))
          .filter(({ bounds }) =>
            bounds.minX >= rect.minX &&
            bounds.maxX <= rect.maxX &&
            bounds.minY >= rect.minY &&
            bounds.maxY <= rect.maxY
          )
          .map(({ idx }) => idx);
      }

      function getSelectionGroupCenter(page, indices, bleedPx) {
        if (!indices.length) return null;
        const bounds = indices
          .map((idx) => getElementBounds(page.elements[idx]))
          .reduce((acc, box) => ({
            minX: Math.min(acc.minX, box.minX),
            minY: Math.min(acc.minY, box.minY),
            maxX: Math.max(acc.maxX, box.maxX),
            maxY: Math.max(acc.maxY, box.maxY)
          }), { minX: Infinity, minY: Infinity, maxX: -Infinity, maxY: -Infinity });

        return {
          x: (bounds.minX + bounds.maxX) / 2 + bleedPx,
          y: (bounds.minY + bounds.maxY) / 2 + bleedPx
        };
      }

      function createFillElements(tool, start, end, style) {
        const minX = Math.min(start.x, end.x);
        const minY = Math.min(start.y, end.y);
        const maxX = Math.max(start.x, end.x);
        const maxY = Math.max(start.y, end.y);
        const step = state.gridSize;

        if (tool === 'line') {
          if (Math.abs(maxX - minX) < 1 || Math.abs(maxY - minY) < 1) {
            return [{ type: 'line', x1: start.x, y1: start.y, x2: end.x, y2: end.y, style }];
          }

          const elements = [];
          for (let y = minY + step; y < maxY; y += step) {
            elements.push({
              type: 'line',
              x1: minX,
              y1: y,
              x2: maxX,
              y2: y,
              style: { ...style }
            });
          }
          return elements.length ? elements : [{ type: 'line', x1: start.x, y1: start.y, x2: end.x, y2: end.y, style }];
        }

        if (Math.abs(maxX - minX) < 1 || Math.abs(maxY - minY) < 1) {
          return [{ type: tool, x: end.x, y: end.y, style }];
        }

        const elements = [];
        for (let x = minX + step; x < maxX; x += step) {
          for (let y = minY + step; y < maxY; y += step) {
            elements.push({ type: tool, x, y, style: { ...style } });
          }
        }

        return elements.length ? elements : [{ type: tool, x: end.x, y: end.y, style }];
      }

      function isElementLargeEnough(tool, start, end) {
        const minSize = state.gridSize / 2;
        const dx = Math.abs(end.x - start.x);
        const dy = Math.abs(end.y - start.y);

        if (tool === 'line') return Math.hypot(dx, dy) >= minSize;
        if (tool === 'rect') return dx >= minSize && dy >= minSize;
        return true;
      }

      function updateNotebookLayout(totalW, totalH) {
        if (state.viewMode !== 'notebook') return;
        const viewport = document.getElementById('viewport');
        const workspace = document.getElementById('workspace');
        if (!viewport || !workspace) return;

        const pageCount = Math.max(state.pages.length, 1);
        const gap = 24;
        const paddingX = 64;
        const paddingY = 64;
        const availableWidth = Math.max(viewport.clientWidth - paddingX, 240);
        const availableHeight = Math.max(viewport.clientHeight - paddingY, 240);
        const pageAspect = totalW / totalH;
        const viewportAspect = availableWidth / availableHeight;
        const estimatedCols = Math.sqrt(pageCount * (viewportAspect / pageAspect));
        const cols = Math.max(1, Math.min(pageCount, Math.ceil(estimatedCols)));
        const rows = Math.ceil(pageCount / cols);
        const widthFromSpace = (availableWidth - gap * (cols - 1)) / cols;
        const heightFromSpace = (availableHeight - gap * (rows - 1)) / rows;
        const safeWidth = Math.max(120, Math.floor(Math.min(widthFromSpace, heightFromSpace * pageAspect)));

        state.notebookLayout = { cols, pageWidth: safeWidth };
        workspace.style.gridTemplateColumns = `repeat(${cols}, ${safeWidth}px)`;
      }

      function getVisibleBleedPx() {
        return state.bleedVisible ? state.bleedUnits * state.gridSize : 0;
      }

      function getCanvasRulerOffset() {
        return state.viewMode === 'canvas' && state.rulerVisible ? 24 : 0;
      }

      function getCanvasFrameExtra() {
        return 20;
      }

      function centerActivePage() {
        if (state.viewMode !== 'canvas') return;
        const viewport = document.getElementById('viewport');
        const activeWrapper = document.querySelector('.page-assembly-wrapper.active-page');
        if (!viewport || !activeWrapper) return;

        const viewportRect = viewport.getBoundingClientRect();
        const wrapperRect = activeWrapper.getBoundingClientRect();
        const targetLeft = viewport.scrollLeft + (wrapperRect.left - viewportRect.left) - (viewport.clientWidth - wrapperRect.width) / 2;
        const targetTop = viewport.scrollTop + (wrapperRect.top - viewportRect.top) - (viewport.clientHeight - wrapperRect.height) / 2;

        viewport.scrollTo({
          left: Math.max(0, targetLeft),
          top: Math.max(0, targetTop),
          behavior: 'smooth'
        });
      }

      // --- Exposed Global Actions ---

      window.init = function() {
        try { lucide.createIcons(); } catch(e) {}
        window.setPaper('letter');
        updateOrientationButton();
        window.setupEventListeners();
        window.updateGridSize('0.25');
        setTimeout(window.fitAuto, 300);
      };

      window.setPaper = function(type) {
        saveHistory(); state.paper = type; state.unit = PAPER_CONFIG[type].unit;
        updateOrientationButton();
        window.renderWorkspace();
      };

      window.toggleOrientation = function() {
        state.orientation = state.orientation === 'portrait' ? 'landscape' : 'portrait';
        updateOrientationButton();
        const viewport = document.getElementById('viewport');
        if (viewport) {
          viewport.scrollTop = 0;
          viewport.scrollLeft = 0;
        }
        window.renderWorkspace();
      };

      window.updateGridSize = function(val) {
        state.gridSize = val.endsWith('mm') ? parseFloat(val) * PMM : parseFloat(val) * PPI;
        window.renderWorkspace();
      };

      window.toggleGrid = function() {
        state.gridVisible = !state.gridVisible;
        const btn = document.getElementById('gridToggle');
        if(btn) { btn.classList.toggle('bg-indigo-600', state.gridVisible); btn.classList.toggle('text-white', state.gridVisible); }
        window.renderWorkspace();
      };

      window.toggleRulers = function() {
        state.rulerVisible = !state.rulerVisible;
        const btn = document.getElementById('rulerToggle');
        if(btn) { btn.classList.toggle('bg-indigo-600', state.rulerVisible); btn.classList.toggle('text-white', state.rulerVisible); }
        window.renderWorkspace();
      };

      window.toggleBleed = function() {
        state.bleedVisible = !state.bleedVisible;
        const btn = document.getElementById('bleedToggle');
        if(btn) { btn.classList.toggle('bg-indigo-600', state.bleedVisible); btn.classList.toggle('text-white', state.bleedVisible); }
        window.renderWorkspace();
      };

      window.toggleNotebookView = function() {
        state.viewMode = state.viewMode === 'canvas' ? 'notebook' : 'canvas';
        document.body.classList.toggle('notebook-view', state.viewMode === 'notebook');
        const btn = document.getElementById('viewToggle');
        if(btn) { btn.classList.toggle('bg-slate-800', state.viewMode === 'notebook'); btn.classList.toggle('text-white', state.viewMode === 'notebook'); }
        window.renderWorkspace();
      };

      window.undo = () => {
        if (state.history.length) {
          state.pages = JSON.parse(state.history.pop());
          state.currentPageIndex = Math.min(state.currentPageIndex, state.pages.length - 1);
          state.selectedIndices = [];
          updatePropsPanel();
          updateUndoButton();
          window.renderWorkspace();
        }
      };
      window.prevPage = () => { if (state.currentPageIndex > 0) { state.currentPageIndex--; window.renderWorkspace(); } };
      window.nextPage = () => { if (state.currentPageIndex < state.pages.length - 1) { state.currentPageIndex++; window.renderWorkspace(); } };
      window.addPage = () => { saveHistory(); state.pages.push({ elements: [] }); state.currentPageIndex = state.pages.length - 1; window.renderWorkspace(); };
      window.copyPage = () => {
        saveHistory();
        const clonedPage = JSON.parse(JSON.stringify(state.pages[state.currentPageIndex]));
        state.pages.splice(state.currentPageIndex + 1, 0, clonedPage);
        state.currentPageIndex += 1;
        window.renderWorkspace();
      };
      window.deletePage = () => { if (state.pages.length > 1) { saveHistory(); state.pages.splice(state.currentPageIndex, 1); state.currentPageIndex = Math.max(0, state.currentPageIndex - 1); window.renderWorkspace(); } };
      
      window.setTool = (t) => {
        state.tool = t;
        document.querySelectorAll('.tool-btn').forEach(b => {
          const isActive = b.id === `tool-${t}`;
          b.classList.toggle('bg-indigo-600', isActive); b.classList.toggle('text-white', isActive);
          b.classList.toggle('text-slate-400', !isActive); b.classList.toggle('shadow-md', isActive);
        });
        state.selectedIndices = []; state.editingIndex = null;
        updatePropsPanel();
        window.renderWorkspace();
      };

      window.updateSelectedStyle = (p, v) => {
        if (state.selectedIndices.length === 0) return;
        saveHistory();
        state.selectedIndices.forEach(idx => state.pages[state.currentPageIndex].elements[idx].style[p] = v);
        updatePropsPanel();
        window.renderWorkspace();
      };
      window.toggleSelectedTextStyle = (prop, activeValue, inactiveValue) => {
        if (state.selectedIndices.length === 0) return;
        const elements = state.selectedIndices.map(idx => state.pages[state.currentPageIndex].elements[idx]).filter(el => el.type === 'text');
        if (!elements.length) return;
        saveHistory();
        const shouldEnable = elements.some(el => (el.style?.[prop] || inactiveValue) !== activeValue);
        state.selectedIndices.forEach(idx => {
          const el = state.pages[state.currentPageIndex].elements[idx];
          if (el.type === 'text') el.style[prop] = shouldEnable ? activeValue : inactiveValue;
        });
        updatePropsPanel();
        window.renderWorkspace();
      };
      window.updateSelectedTextFontSize = (value) => {
        if (state.selectedIndices.length === 0) return;
        const fontSize = Number(value);
        if (!fontSize) return;
        saveHistory();
        state.selectedIndices.forEach(idx => {
          const el = state.pages[state.currentPageIndex].elements[idx];
          if (el.type === 'text') el.style.fontSize = fontSize;
        });
        document.getElementById('textSizeVal').innerText = formatPxValue(fontSize);
        updatePropsPanel();
        window.renderWorkspace();
      };
      window.deleteSelected = () => { state.pages[state.currentPageIndex].elements = state.pages[state.currentPageIndex].elements.filter((_, i) => !state.selectedIndices.includes(i)); state.selectedIndices = []; updatePropsPanel(); window.renderWorkspace(); };
      window.clearCanvas = () => { if(confirm("Reset entire page?")) { saveHistory(); state.pages[state.currentPageIndex].elements = []; state.selectedIndices = []; updatePropsPanel(); window.renderWorkspace(); } };
      
      window.updateZoom = function(newZoom) {
        state.zoom = Math.max(0.1, Math.min(3, newZoom));
        const ws = document.getElementById('workspace');
        if (ws) ws.style.transform = `scale(${state.zoom})`;
      };

      window.fitAuto = function() {
        const viewport = document.getElementById('viewport');
        if (!viewport) return;
        if (state.viewMode === 'notebook') {
          window.renderWorkspace();
          return;
        }
        const metrics = getWorkspaceMetrics();
        const fitHeight = Math.max(0.1, (viewport.clientHeight - 14) / metrics.framedHeight);
        const fitWidth = Math.max(0.1, (viewport.clientWidth - 32) / metrics.framedWidth);
        window.updateZoom(Math.min(fitHeight, fitWidth));
      };

      window.downloadPDF = async function() {
        const jsPDFCtor = window.jspdf?.jsPDF;
        if (!jsPDFCtor) {
          alert('PDF export is unavailable because jsPDF did not load.');
          return;
        }
        try {
          const exportQuality = await promptForPdfQuality();
          const now = new Date();
          const pad2 = (value) => String(value).padStart(2, '0');
          const stamp = `${pad2(now.getMonth() + 1)}${pad2(now.getDate())}${String(now.getFullYear()).slice(-2)}-${pad2(now.getHours())}${pad2(now.getMinutes())}`;
          const { width: pageWidth, height: pageHeight } = getPaperMetrics();
          const pdf = new jsPDFCtor({
            orientation: pageWidth >= pageHeight ? 'landscape' : 'portrait',
            unit: 'pt',
            format: [pageWidth, pageHeight],
            compress: true
          });

          const svgNs = "http://www.w3.org/2000/svg";
          const renderPageToDataUrl = async (page) => {
            const svg = document.createElementNS(svgNs, "svg");
            svg.setAttribute("xmlns", svgNs);
            svg.setAttribute("width", pageWidth);
            svg.setAttribute("height", pageHeight);
            svg.setAttribute("viewBox", `0 0 ${pageWidth} ${pageHeight}`);

            const bg = document.createElementNS(svgNs, "rect");
            bg.setAttribute("x", "0");
            bg.setAttribute("y", "0");
            bg.setAttribute("width", `${pageWidth}`);
            bg.setAttribute("height", `${pageHeight}`);
            bg.setAttribute("fill", "#ffffff");
            svg.appendChild(bg);

            page.elements.forEach((el) => {
              const weight = Number(el.style?.weight || 1);
              const dash = el.style?.dash || 'none';
              const opacity = Number(el.style?.opacity || 100) / 100;
              let node;

              if (el.type === 'line') {
                node = document.createElementNS(svgNs, "line");
                node.setAttribute("x1", `${el.x1}`);
                node.setAttribute("y1", `${el.y1}`);
                node.setAttribute("x2", `${el.x2}`);
                node.setAttribute("y2", `${el.y2}`);
                node.setAttribute("fill", "none");
              } else if (el.type === 'rect') {
                node = document.createElementNS(svgNs, "rect");
                node.setAttribute("x", `${el.x}`);
                node.setAttribute("y", `${el.y}`);
                node.setAttribute("width", `${el.w}`);
                node.setAttribute("height", `${el.h}`);
                node.setAttribute("fill", "none");
              } else if (el.type === 'dot') {
                node = document.createElementNS(svgNs, "circle");
                node.setAttribute("cx", `${el.x}`);
                node.setAttribute("cy", `${el.y}`);
                node.setAttribute("r", `${Math.max(0.5, Math.min(5, weight))}`);
                node.setAttribute("fill", "#334155");
              } else if (el.type === 'cross') {
                node = document.createElementNS(svgNs, "path");
                const size = 4;
                node.setAttribute("d", `M ${el.x-size} ${el.y} H ${el.x+size} M ${el.x} ${el.y-size} V ${el.y+size}`);
                node.setAttribute("fill", "none");
                node.setAttribute("stroke-linecap", "round");
              } else if (el.type === 'text') {
                node = document.createElementNS(svgNs, "g");
                const paddingX = 4;
                const paddingY = 2;
                const fontSize = getTextFontSize(el);
                const lineHeight = fontSize * 1.2;
                const lines = wrapTextToWidth(el.text || '', Math.max(0, el.w - paddingX * 2), getTextFontString(el));
                const textAlign = el.style?.textAlign || 'left';
                const verticalAlign = el.style?.verticalAlign || 'top';
                const contentHeight = Math.max(lineHeight, lines.length * lineHeight);
                const startY = verticalAlign === 'center'
                  ? el.y + paddingY + Math.max(0, (el.h - paddingY * 2 - contentHeight) / 2)
                  : verticalAlign === 'bottom'
                    ? el.y + el.h - paddingY - contentHeight
                    : el.y + paddingY;
                const textNode = document.createElementNS(svgNs, "text");
                textNode.setAttribute("font-family", "Helvetica, Arial, sans-serif");
                textNode.setAttribute("font-size", `${fontSize}`);
                textNode.setAttribute("font-weight", el.style?.fontWeight || '400');
                textNode.setAttribute("font-style", el.style?.fontStyle || 'normal');
                textNode.setAttribute("fill", "#334155");
                const xPos = textAlign === 'center'
                  ? el.x + (el.w / 2)
                  : textAlign === 'right'
                    ? el.x + el.w - paddingX
                    : el.x + paddingX;
                textNode.setAttribute("x", `${xPos}`);
                textNode.setAttribute("text-anchor", textAlign === 'center' ? 'middle' : textAlign === 'right' ? 'end' : 'start');
                lines.forEach((line, lineIndex) => {
                  const tspan = document.createElementNS(svgNs, "tspan");
                  tspan.setAttribute("x", `${xPos}`);
                  tspan.setAttribute("y", `${startY + fontSize + (lineIndex * lineHeight)}`);
                  tspan.textContent = line;
                  textNode.appendChild(tspan);
                });
                node.appendChild(textNode);
              }

              if (!node) return;
              node.setAttribute("opacity", `${opacity}`);
              if (el.type !== 'text') {
                node.setAttribute("stroke", "#334155");
                node.setAttribute("stroke-width", `${weight}`);
                if (dash !== 'none') node.setAttribute("stroke-dasharray", dash);
              }
              svg.appendChild(node);
            });

            const svgString = new XMLSerializer().serializeToString(svg);
            const blob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
            const url = URL.createObjectURL(blob);

            try {
              const img = await new Promise((resolve, reject) => {
                const image = new Image();
                image.onload = () => resolve(image);
                image.onerror = reject;
                image.src = url;
              });

              const scale = exportQuality.scale;
              const canvas = document.createElement('canvas');
              canvas.width = Math.round(pageWidth * scale);
              canvas.height = Math.round(pageHeight * scale);
              const ctx = canvas.getContext('2d');
              ctx.fillStyle = '#ffffff';
              ctx.fillRect(0, 0, canvas.width, canvas.height);
              ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
              return canvas.toDataURL(exportQuality.mime, exportQuality.quality);
            } finally {
              URL.revokeObjectURL(url);
            }
          };

          for (let index = 0; index < state.pages.length; index++) {
            if (index > 0) pdf.addPage([pageWidth, pageHeight], pageWidth >= pageHeight ? 'landscape' : 'portrait');
            const pageImage = await renderPageToDataUrl(state.pages[index]);
            pdf.addImage(pageImage, exportQuality.mime === 'image/png' ? 'PNG' : 'JPEG', 0, 0, pageWidth, pageHeight, undefined, exportQuality.compression);
          }

          pdf.save(`PDFbaker-${stamp}.pdf`);
        } catch (error) {
          if (error?.message === 'cancelled') return;
          console.error('PDF export failed:', error);
          alert('PDF export failed. Please refresh and try again.');
        }
      };

      window.getSnappedCoords = function(e) {
        const stage = document.getElementById(`stage-${state.currentPageIndex}`);
        if (!stage) return {x:0, y:0};
        const rect = stage.getBoundingClientRect();
        const rawX = (e.clientX - rect.left) / state.zoom;
        const rawY = (e.clientY - rect.top) / state.zoom;
        const { bleedPx, pageWidth: sheetW, pageHeight: sheetH } = getWorkspaceMetrics();
        const step = state.gridSize / 2;
        const ox = (sheetW / 2 + bleedPx) % state.gridSize;
        const oy = (sheetH / 2 + bleedPx) % state.gridSize;
        const fx = Math.round((rawX - ox) / step) * step + ox;
        const fy = Math.round((rawY - oy) / step) * step + oy;

        const dot = document.getElementById(`snapDot-${state.currentPageIndex}`);
        if (dot) {
          dot.setAttribute('cx', fx); dot.setAttribute('cy', fy);
          if (state.tool !== 'select' || state.isDrawing || state.isDragging) dot.classList.remove('opacity-0');
          else dot.classList.add('opacity-0');
        }
        const cd = document.getElementById('coordDisplay');
        if(cd) cd.innerText = `X: ${((fx - bleedPx - sheetW/2) / (state.unit === 'in' ? PPI : PMM)).toFixed(2)} Y: ${((fy - bleedPx - sheetH/2) / (state.unit === 'in' ? PPI : PMM)).toFixed(2)}`;
        return { x: fx - bleedPx, y: fy - bleedPx };
      };

      window.renderWorkspace = function() {
        const workspace = document.getElementById('workspace');
        if (!workspace) return;
        workspace.innerHTML = '';
        workspace.style.gridTemplateColumns = '';
        workspace.style.alignSelf = state.viewMode === 'canvas' && state.pages.length === 1 ? 'center' : (state.viewMode === 'canvas' ? 'flex-start' : 'center');
        const pageChangeDirection = state.viewMode === 'canvas' && state.lastCanvasPageIndex !== null
          ? (state.currentPageIndex > state.lastCanvasPageIndex ? 'next' : (state.currentPageIndex < state.lastCanvasPageIndex ? 'prev' : null))
          : null;
        const { pageWidth: w, pageHeight: h, bleedPx, rulerOffset, totalWidth: totalW, totalHeight: totalH } = getWorkspaceMetrics();

        updateNotebookLayout(totalW, totalH);

        state.pages.forEach((page, idx) => {
          const wrapper = document.createElement('div');
          wrapper.className = `page-assembly-wrapper ${idx === state.currentPageIndex ? 'active-page' : ''}`;
          if (state.viewMode === 'canvas' && idx === state.currentPageIndex && pageChangeDirection) {
            wrapper.classList.add(pageChangeDirection === 'next' ? 'canvas-page-enter-next' : 'canvas-page-enter-prev');
          }
          if (state.viewMode === 'notebook') wrapper.style.width = `${state.notebookLayout.pageWidth}px`;
          if (state.viewMode === 'notebook') {
            wrapper.onclick = () => {
              clearTimeout(state.notebookClickTimer);
              state.notebookClickTimer = setTimeout(() => {
                state.currentPageIndex = idx;
                window.renderWorkspace();
                updatePageIndicators();
                state.notebookClickTimer = null;
              }, 180);
            };
            wrapper.ondblclick = () => {
              clearTimeout(state.notebookClickTimer);
              state.notebookClickTimer = null;
              state.currentPageIndex = idx;
              state.viewMode = 'canvas';
              document.body.classList.remove('notebook-view');
              const btn = document.getElementById('viewToggle');
              if (btn) {
                btn.classList.remove('bg-slate-800');
                btn.classList.remove('text-white');
              }
              window.renderWorkspace();
              updatePageIndicators();
            };
          }
          const assembly = document.createElement('div');
          assembly.className = 'page-assembly';
          
          if (state.viewMode === 'canvas') {
              assembly.style.width = `${totalW + rulerOffset}px`; assembly.style.height = `${totalH + rulerOffset}px`;
              if (state.rulerVisible) {
                const rt = document.createElementNS("http://www.w3.org/2000/svg", "svg");
                rt.setAttribute("class", "ruler absolute top-0 border-b");
                rt.style.left = `${rulerOffset}px`; rt.style.width = `${totalW}px`; rt.style.height = `${rulerOffset}px`;
                renderRulerContent(rt, totalW, 'horizontal');
                const rl = document.createElementNS("http://www.w3.org/2000/svg", "svg");
                rl.setAttribute("class", "ruler absolute left-0 border-r");
                rl.style.top = `${rulerOffset}px`; rl.style.width = `${rulerOffset}px`; rl.style.height = `${totalH}px`;
                renderRulerContent(rl, totalH, 'vertical');
                const cr = document.createElement('div'); cr.id = "rulerCorner"; cr.className = "w-8 h-8 bg-white border-b border-r border-slate-200 absolute left-0 top-0";
                cr.style.width = `${rulerOffset}px`;
                cr.style.height = `${rulerOffset}px`;
                assembly.appendChild(cr); assembly.appendChild(rt); assembly.appendChild(rl);
              }
          } else {
              assembly.style.width = '100%'; assembly.style.height = 'auto'; assembly.style.aspectRatio = `${totalW} / ${totalH}`;
          }

          const container = document.createElement('div');
          container.className = "page-container";
          if (state.viewMode === 'canvas') { container.classList.add('absolute'); container.style.left = `${rulerOffset}px`; container.style.top = `${rulerOffset}px`; }
          container.style.width = `${totalW}px`; container.style.height = `${totalH}px`;
          container.onclick = () => {
            state.currentPageIndex = idx;
            window.renderWorkspace();
            updatePageIndicators();
          };

          const stage = document.createElementNS("http://www.w3.org/2000/svg", "svg");
          stage.setAttribute("class", "absolute inset-0 w-full h-full");
          stage.style.cursor = state.tool === 'select' ? 'default' : 'crosshair';
          stage.id = `stage-${idx}`;
          stage.setAttribute("viewBox", `0 0 ${totalW} ${totalH}`);
          
          // Layer 0: Physical Sheet
          const sheet = document.createElementNS("http://www.w3.org/2000/svg", "rect");
          sheet.setAttribute("x", bleedPx); sheet.setAttribute("y", bleedPx);
          sheet.setAttribute("width", w); sheet.setAttribute("height", h);
          sheet.setAttribute("class", "visible-sheet");
          
          // Layer 1: Grid
          const gridLayer = document.createElementNS("http://www.w3.org/2000/svg", "g");
          if (state.gridVisible) renderGridLines(gridLayer, totalW, totalH);

          // Layer 2: Design Elements
          const gElements = document.createElementNS("http://www.w3.org/2000/svg", "g");
          
          // Layer 3: Helpers
          const marquee = document.createElementNS("http://www.w3.org/2000/svg", "rect");
          marquee.id = `marquee-${idx}`; marquee.setAttribute("class", "hidden");
          marquee.setAttribute("fill", "rgba(99, 102, 241, 0.1)"); marquee.setAttribute("stroke", "#6366f1");
          marquee.setAttribute("stroke-width", "1"); marquee.setAttribute("stroke-dasharray", "4");

          const snapDot = document.createElementNS("http://www.w3.org/2000/svg", "circle");
          snapDot.id = `snapDot-${idx}`; snapDot.setAttribute("r", "3.5"); snapDot.setAttribute("fill", "#6366f1");
          snapDot.setAttribute("class", "opacity-0 pointer-events-none transition-opacity duration-75");

          const tLine = document.createElementNS("http://www.w3.org/2000/svg", "line");
          tLine.id = `tempLine-${idx}`; tLine.setAttribute("class", "hidden");
          tLine.setAttribute("stroke", "#6366f1"); tLine.setAttribute("stroke-width", "1.5"); tLine.setAttribute("stroke-dasharray", "4");

          const tRect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
          tRect.id = `tempRect-${idx}`; tRect.setAttribute("class", "hidden");
          tRect.setAttribute("fill", "rgba(99, 102, 241, 0.1)"); tRect.setAttribute("stroke", "#6366f1");
          tRect.setAttribute("stroke-width", "1"); tRect.setAttribute("stroke-dasharray", "4");

          // Layer 4: Bleed Frame (Top)
          const bleed = document.createElementNS("http://www.w3.org/2000/svg", "path");
          bleed.setAttribute("d", `M 0 0 H ${totalW} V ${totalH} H 0 Z M ${bleedPx} ${bleedPx} V ${bleedPx+h} H ${bleedPx+w} V ${bleedPx} Z`);
          bleed.setAttribute("class", `bleed-overlay${state.bleedVisible ? '' : ' hidden'}`);

          stage.appendChild(sheet); stage.appendChild(gridLayer); stage.appendChild(gElements);
          stage.appendChild(marquee); stage.appendChild(tLine); stage.appendChild(tRect); stage.appendChild(snapDot); stage.appendChild(bleed);
          container.appendChild(stage); assembly.appendChild(container); wrapper.appendChild(assembly); workspace.appendChild(wrapper);

          // Element Rendering
          const mOp = 1;
          const singleSelection = idx === state.currentPageIndex && state.selectedIndices.length === 1 ? state.selectedIndices[0] : null;
          page.elements.forEach((el, i) => {
            const isS = (idx === state.currentPageIndex && state.selectedIndices.includes(i));
            const op = ((el.style?.opacity || 100)/100)*mOp, wt = el.style?.weight || 1;
            let node;
            if (el.type==='line') { node = document.createElementNS("http://www.w3.org/2000/svg", "line"); node.setAttribute("x1", el.x1+bleedPx); node.setAttribute("y1", el.y1+bleedPx); node.setAttribute("x2", el.x2+bleedPx); node.setAttribute("y2", el.y2+bleedPx); node.setAttribute("stroke", "#334155"); node.setAttribute("stroke-width", wt); if(el.style?.dash !== 'none') node.setAttribute("stroke-dasharray", el.style.dash); }
            else if (el.type==='rect') { node = document.createElementNS("http://www.w3.org/2000/svg", "rect"); node.setAttribute("x", el.x+bleedPx); node.setAttribute("y", el.y+bleedPx); node.setAttribute("width", el.w); node.setAttribute("height", el.h); node.setAttribute("fill", "none"); node.setAttribute("stroke", "#334155"); node.setAttribute("stroke-width", wt); if(el.style?.dash !== 'none') node.setAttribute("stroke-dasharray", el.style.dash); }
            else if (el.type==='dot') { node = document.createElementNS("http://www.w3.org/2000/svg", "circle"); node.setAttribute("cx", el.x+bleedPx); node.setAttribute("cy", el.y+bleedPx); node.setAttribute("r", Math.max(0.5, Math.min(5, wt))); node.setAttribute("fill", "#334155"); }
            else if (el.type==='cross') {
              node = document.createElementNS("http://www.w3.org/2000/svg", "path");
              const cx = el.x + bleedPx, cy = el.y + bleedPx, size = 4;
              node.setAttribute("d", `M ${cx-size} ${cy} H ${cx+size} M ${cx} ${cy-size} V ${cy+size}`);
              node.setAttribute("fill", "none");
              node.setAttribute("stroke", "#334155");
              node.setAttribute("stroke-width", wt || 1);
              node.setAttribute("stroke-linecap", "round");
            }
            else if (el.type==='text') {
              node = document.createElementNS("http://www.w3.org/2000/svg", "g");
              const textBoxRect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
              textBoxRect.setAttribute("x", el.x + bleedPx);
              textBoxRect.setAttribute("y", el.y + bleedPx);
              textBoxRect.setAttribute("width", el.w);
              textBoxRect.setAttribute("height", el.h);
              textBoxRect.setAttribute("fill", "#ffffff");
              textBoxRect.setAttribute("fill-opacity", "0.72");
              textBoxRect.setAttribute("stroke", "#6366f1");
              textBoxRect.setAttribute("stroke-opacity", "0.18");
              textBoxRect.setAttribute("stroke-width", "1");
              textBoxRect.setAttribute("stroke-dasharray", "4 3");
              node.appendChild(textBoxRect);
              const textLayer = document.createElementNS("http://www.w3.org/2000/svg", "foreignObject");
              textLayer.setAttribute("x", el.x + bleedPx);
              textLayer.setAttribute("y", el.y + bleedPx);
              textLayer.setAttribute("width", el.w);
              textLayer.setAttribute("height", el.h);
              const wrapperDiv = document.createElement('div');
              wrapperDiv.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml');
              wrapperDiv.className = 'text-box-container';
              wrapperDiv.style.justifyContent =
                (el.style.verticalAlign || 'top') === 'center' ? 'center'
                : (el.style.verticalAlign || 'top') === 'bottom' ? 'flex-end'
                : 'flex-start';
              const contentDiv = document.createElement('div');
              contentDiv.className = 'text-box-content';
              contentDiv.textContent = el.text || '';
              contentDiv.style.fontSize = `${el.style?.fontSize || el.fontSize || Math.max(12, state.gridSize * 0.7)}px`;
              contentDiv.style.fontWeight = el.style?.fontWeight || '400';
              contentDiv.style.fontStyle = el.style?.fontStyle || 'normal';
              contentDiv.style.textAlign = el.style.textAlign || 'left';
              contentDiv.contentEditable = String(isS && state.editingIndex === i);
              if (isS && state.editingIndex === i) {
                contentDiv.onkeydown = (e) => {
                  if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    contentDiv.blur();
                  }
                };
                contentDiv.oninput = () => {
                  const textElement = state.pages[state.currentPageIndex].elements[i];
                  textElement.text = contentDiv.textContent || '';
                  const expanded = expandTextBoxToFit(textElement, contentDiv);
                  if (expanded) {
                    textBoxRect.setAttribute("height", textElement.h);
                    textLayer.setAttribute("height", textElement.h);
                  }
                };
                contentDiv.onblur = () => {
                  state.editingIndex = null;
                  updatePropsPanel();
                  window.renderWorkspace();
                };
              } else {
                contentDiv.onmousedown = (e) => e.stopPropagation();
                contentDiv.ondblclick = (e) => {
                  e.stopPropagation();
                  state.selectedIndices = [i];
                  state.editingIndex = i;
                  updatePropsPanel();
                  window.renderWorkspace();
                };
              }
              wrapperDiv.appendChild(contentDiv);
              textLayer.appendChild(wrapperDiv);
              node.appendChild(textLayer);
            }
            if (node) {
              node.setAttribute("opacity", op);
              if (el.type === 'text') {
                const textBoxRect = node.querySelector('rect');
                if (textBoxRect) {
                  textBoxRect.classList.add('element-hover');
                  if (isS) textBoxRect.classList.add('element-selected');
                }
              } else {
                node.classList.add('element-hover');
                if (isS) node.classList.add('element-selected');
              }
              gElements.appendChild(node);
            }
            const hit = document.createElementNS("http://www.w3.org/2000/svg", (el.type==='line'?'line':((el.type==='rect' || el.type==='text')?'rect':'circle')));
            hit.setAttribute("class", "hit-target");
            if (isS) hit.classList.add('hit-target-selected');
            if (state.tool !== 'select' && !isS) hit.style.pointerEvents = 'none';
            if (el.type==='line') { hit.setAttribute("x1", el.x1+bleedPx); hit.setAttribute("y1", el.y1+bleedPx); hit.setAttribute("x2", el.x2+bleedPx); hit.setAttribute("y2", el.y2+bleedPx); }
            else if (el.type==='rect' || el.type==='text') { hit.setAttribute("x", el.x+bleedPx); hit.setAttribute("y", el.y+bleedPx); hit.setAttribute("width", el.w); hit.setAttribute("height", el.h); }
            else { hit.setAttribute("cx", el.x+bleedPx); hit.setAttribute("cy", el.y+bleedPx); hit.setAttribute("r", 10); }
            hit.onmousedown = (e) => {
              e.stopPropagation();
              if (state.selectedIndices.includes(i) && state.tool === 'select') {
                saveHistory();
                state.isDragging = true;
                state.dragElementIndex = i;
                state.startPoint = window.getSnappedCoords(e);
                document.body.classList.add('is-grabbing');
                return;
              }
              if (!state.selectedIndices.includes(i)) {
                if (!e.shiftKey) state.selectedIndices = [i];
                else state.selectedIndices.push(i);
              } else if (e.shiftKey) {
                state.selectedIndices = state.selectedIndices.filter(idx => idx !== i);
              }
              updatePropsPanel();
              window.renderWorkspace();
            };
            if (el.type === 'text') {
              hit.ondblclick = (e) => {
                e.stopPropagation();
                if (state.tool !== 'select') return;
                state.selectedIndices = [i];
                state.editingIndex = i;
                updatePropsPanel();
                window.renderWorkspace();
              };
            }
            gElements.appendChild(hit);

            if (isS && idx === state.currentPageIndex && singleSelection === i) {
              const handles = getElementHandles(el, bleedPx);
              handles.edges.forEach((pt) => {
                const edgeHandle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
                edgeHandle.setAttribute("cx", pt.x);
                edgeHandle.setAttribute("cy", pt.y);
                edgeHandle.setAttribute("r", "5");
                edgeHandle.setAttribute("class", "handle-edge");
                edgeHandle.onmousedown = (e) => {
                  e.stopPropagation();
                  saveHistory();
                  state.isModifying = true;
                  state.dragElementIndex = i;
                  state.modifyHandleType = { kind: 'edge', index: handles.edges.indexOf(pt) };
                  state.startPoint = window.getSnappedCoords(e);
                };
                gElements.appendChild(edgeHandle);
              });

              const centerHandle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
              centerHandle.setAttribute("cx", handles.center.x);
              centerHandle.setAttribute("cy", handles.center.y);
              centerHandle.setAttribute("r", "6");
              centerHandle.setAttribute("class", "handle-center");
              centerHandle.onmousedown = (e) => {
                e.stopPropagation();
                saveHistory();
                state.isDragging = true;
                state.dragElementIndex = i;
                state.startPoint = window.getSnappedCoords(e);
                document.body.classList.add('is-grabbing');
              };
              gElements.appendChild(centerHandle);
            }
          });

          if (idx === state.currentPageIndex && state.selectedIndices.length > 1) {
            const groupCenter = getSelectionGroupCenter(page, state.selectedIndices, bleedPx);
            if (groupCenter) {
              const centerHandle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
              centerHandle.setAttribute("cx", groupCenter.x);
              centerHandle.setAttribute("cy", groupCenter.y);
              centerHandle.setAttribute("r", "7");
              centerHandle.setAttribute("class", "handle-center");
              centerHandle.onmousedown = (e) => {
                e.stopPropagation();
                saveHistory();
                state.isDragging = true;
                state.dragElementIndex = null;
                state.startPoint = window.getSnappedCoords(e);
                document.body.classList.add('is-grabbing');
              };
              gElements.appendChild(centerHandle);
            }
          }

          if (idx === state.currentPageIndex) {
              stage.onmousedown = (e) => {
                const c = window.getSnappedCoords(e);
                if (e.target === stage || e.target.classList.contains('visible-sheet')) {
                  if (state.tool === 'select') {
                    state.isSelecting = true;
                    state.startPoint = c;
                    const startBleedPx = getVisibleBleedPx();
                    const marqueeEl = document.getElementById(`marquee-${state.currentPageIndex}`);
                    if (marqueeEl) {
                      marqueeEl.classList.remove('hidden');
                      marqueeEl.setAttribute('x', c.x + startBleedPx);
                      marqueeEl.setAttribute('y', c.y + startBleedPx);
                      marqueeEl.setAttribute('width', 0);
                      marqueeEl.setAttribute('height', 0);
                    }
                    if (!e.shiftKey) state.selectedIndices = [];
                    updatePropsPanel();
                    return;
                  }
                  if (state.selectedIndices.length > 0) {
                    state.selectedIndices = [];
                    updatePropsPanel();
                    window.renderWorkspace();
                    return;
                  }
                }
                state.isDrawing = true;
                state.isFilling = (state.tool === 'line' && e.shiftKey) || state.tool === 'dot' || state.tool === 'cross';
                state.startPoint = c;
              };
          }
        });
        updatePageIndicators();
        requestAnimationFrame(() => {
          if (state.viewMode === 'canvas') {
            window.fitAuto();
            centerActivePage();
            state.lastCanvasPageIndex = state.currentPageIndex;
          }
          if (state.editingIndex !== null) {
            const activeEditor = document.querySelector('.text-box-content[contenteditable="true"]');
            if (activeEditor) {
              const textElement = state.pages[state.currentPageIndex].elements[state.editingIndex];
              expandTextBoxToFit(textElement, activeEditor);
              activeEditor.focus();
              const selection = window.getSelection();
              const range = document.createRange();
              range.selectNodeContents(activeEditor);
              range.collapse(false);
              selection.removeAllRanges();
              selection.addRange(range);
            }
          }
        });
      };

      // --- Setup Handlers ---

      window.setupEventListeners = function() {
        window.onresize = () => {
          if (state.viewMode === 'notebook') window.renderWorkspace();
          else {
            window.fitAuto();
            centerActivePage();
          }
        };
        window.onmousemove = (e) => {
          const c = window.getSnappedCoords(e);
          const bleedPx = getVisibleBleedPx();
          if (state.isSelecting) {
            const marqueeEl = document.getElementById(`marquee-${state.currentPageIndex}`);
            if (marqueeEl) {
              marqueeEl.classList.remove('hidden');
              marqueeEl.setAttribute('x', Math.min(state.startPoint.x, c.x) + bleedPx);
              marqueeEl.setAttribute('y', Math.min(state.startPoint.y, c.y) + bleedPx);
              marqueeEl.setAttribute('width', Math.abs(c.x - state.startPoint.x));
              marqueeEl.setAttribute('height', Math.abs(c.y - state.startPoint.y));
            }
            return;
          }
          if (state.isDrawing) {
            const tl = document.getElementById(`tempLine-${state.currentPageIndex}`);
            const tr = document.getElementById(`tempRect-${state.currentPageIndex}`);
            if (state.tool === 'line' && tl && tr) {
              tl.classList.remove('hidden');
              tl.setAttribute('x1', state.startPoint.x + bleedPx);
              tl.setAttribute('y1', state.startPoint.y + bleedPx);
              tl.setAttribute('x2', c.x + bleedPx);
              tl.setAttribute('y2', c.y + bleedPx);
              if (state.isFilling) {
                tr.classList.remove('hidden');
                tr.setAttribute('x', Math.min(state.startPoint.x, c.x) + bleedPx);
                tr.setAttribute('y', Math.min(state.startPoint.y, c.y) + bleedPx);
                tr.setAttribute('width', Math.abs(c.x - state.startPoint.x));
                tr.setAttribute('height', Math.abs(c.y - state.startPoint.y));
              } else {
                tr.classList.add('hidden');
              }
            } 
            else if ((state.tool === 'rect' || state.tool === 'dot' || state.tool === 'cross' || state.tool === 'label') && tr) { tr.classList.remove('hidden'); tr.setAttribute('x', Math.min(state.startPoint.x, c.x) + bleedPx); tr.setAttribute('y', Math.min(state.startPoint.y, c.y) + bleedPx); tr.setAttribute('width', Math.abs(c.x - state.startPoint.x)); tr.setAttribute('height', Math.abs(c.y - state.startPoint.y)); }
          }
          if (state.isModifying && state.dragElementIndex !== null) {
            const el = state.pages[state.currentPageIndex].elements[state.dragElementIndex];
            applyHandleModification(el, state.modifyHandleType, c);
            window.renderWorkspace();
            return;
          }
          if (state.isDragging && state.selectedIndices.length > 0) {
            const dx = c.x - state.startPoint.x, dy = c.y - state.startPoint.y;
            if (dx === 0 && dy === 0) return;
            state.selectedIndices.forEach(idx => { const el = state.pages[state.currentPageIndex].elements[idx]; if (el.type === 'line') { el.x1 += dx; el.y1 += dy; el.x2 += dx; el.y2 += dy; } else { el.x += dx; el.y += dy; } });
            state.startPoint = c; window.renderWorkspace();
          }
        };
        window.onmouseup = (e) => {
          if (state.isSelecting) {
            state.isSelecting = false;
            const c = window.getSnappedCoords(e);
            const marqueeRect = {
              minX: Math.min(state.startPoint.x, c.x),
              minY: Math.min(state.startPoint.y, c.y),
              maxX: Math.max(state.startPoint.x, c.x),
              maxY: Math.max(state.startPoint.y, c.y)
            };
            const marqueeEl = document.getElementById(`marquee-${state.currentPageIndex}`);
            if (marqueeEl) marqueeEl.classList.add('hidden');
            const hasArea = Math.abs(c.x - state.startPoint.x) > 2 || Math.abs(c.y - state.startPoint.y) > 2;
            if (hasArea) state.selectedIndices = getMarqueeSelection(marqueeRect);
            updatePropsPanel();
            window.renderWorkspace();
            return;
          }
          if (state.isDrawing) {
              state.isDrawing = false;
              const c = window.getSnappedCoords(e), s = { opacity: 100, weight: 1, dash: 'none', textAlign: 'left', verticalAlign: 'top', fontWeight: '400', fontStyle: 'normal', fontSize: Math.max(12, Math.round(state.gridSize * 0.7)) };
              let el = state.tool === 'line' ? { type:'line', x1: state.startPoint.x, y1: state.startPoint.y, x2: c.x, y2: c.y, style:s } : { type:'rect', x: Math.min(state.startPoint.x, c.x), y: Math.min(state.startPoint.y, c.y), w: Math.abs(c.x-state.startPoint.x), h: Math.abs(c.y-state.startPoint.y), style:s };
              if (state.tool === 'label') {
                const boxW = Math.abs(c.x - state.startPoint.x);
                const boxH = Math.abs(c.y - state.startPoint.y);
                if (boxW >= state.gridSize / 2 && boxH >= state.gridSize / 2) {
                  saveHistory();
                  const textEl = {
                    type: 'text',
                    x: Math.min(state.startPoint.x, c.x),
                    y: Math.min(state.startPoint.y, c.y),
                    w: boxW,
                    h: boxH,
                    text: '',
                    fontSize: Math.max(12, Math.round(state.gridSize * 0.7)),
                    style: s
                  };
                  state.pages[state.currentPageIndex].elements.push(textEl);
                  state.selectedIndices = [state.pages[state.currentPageIndex].elements.length - 1];
                  state.editingIndex = state.selectedIndices[0];
                  updatePropsPanel();
                }
              } else if ((state.tool === 'dot' || state.tool === 'cross') || isElementLargeEnough(state.tool, state.startPoint, c)) {
                saveHistory();
                if (state.tool === 'dot' || state.tool === 'cross' || (state.tool === 'line' && state.isFilling)) {
                  state.pages[state.currentPageIndex].elements.push(...createFillElements(state.tool, state.startPoint, c, s));
                } else {
                  state.pages[state.currentPageIndex].elements.push(el);
                }
              }
              state.isFilling = false;
              window.renderWorkspace();
          }
          state.isModifying = false;
          state.modifyHandleType = null;
          state.isDragging = false;
          state.dragElementIndex = null;
          document.body.classList.remove('is-grabbing');
        };
        window.onkeydown = (e) => {
          const cmdCtrl = (navigator.platform.toUpperCase().indexOf('MAC') >= 0) ? e.metaKey : e.ctrlKey;
          if (e.key === 'z' && cmdCtrl) { e.preventDefault(); window.undo(); return; }
          if (e.key.toLowerCase() === 'c' && cmdCtrl && state.selectedIndices.length > 0) { e.preventDefault(); state.clipboard = state.selectedIndices.map(idx => JSON.parse(JSON.stringify(state.pages[state.currentPageIndex].elements[idx]))); return; }
          if (e.key.toLowerCase() === 'v' && cmdCtrl && state.clipboard.length > 0) {
              e.preventDefault(); saveHistory();
              const step = state.gridSize; const newIdx = [];
              state.clipboard.forEach(el => {
                  const pastedEl = JSON.parse(JSON.stringify(el));
                  if (pastedEl.type === 'line') { pastedEl.x1 += step; pastedEl.y1 += step; pastedEl.x2 += step; pastedEl.y2 += step; } 
                  else { pastedEl.x += step; pastedEl.y += step; }
                  state.pages[state.currentPageIndex].elements.push(pastedEl);
                  newIdx.push(state.pages[state.currentPageIndex].elements.length - 1);
              });
              state.selectedIndices = newIdx; window.renderWorkspace(); return;
          }
          if ((e.key === 'Delete' || e.key === 'Backspace') && state.selectedIndices.length > 0) { e.preventDefault(); window.deleteSelected(); return; }
          if (state.selectedIndices.length > 0 && e.key.startsWith('Arrow')) {
              e.preventDefault(); saveHistory();
              const step = state.gridSize / 2; let dx = 0, dy = 0;
              if (e.key === 'ArrowUp') dy = -step; if (e.key === 'ArrowDown') dy = step; if (e.key === 'ArrowLeft') dx = -step; if (e.key === 'ArrowRight') dx = step;
              state.selectedIndices.forEach(idx => {
                  const el = state.pages[state.currentPageIndex].elements[idx];
                  if (el.type === 'line') { el.x1 += dx; el.y1 += dy; el.x2 += dx; el.y2 += dy; } else { el.x += dx; el.y += dy; }
              });
              window.renderWorkspace();
          }
        };
      };

      window.onload = window.init;
})();

