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
      const ACCENT = '#C8A46A';
      const ACCENT_SOFT = 'rgba(200, 164, 106, 0.12)';
      const ACCENT_SOFTER = 'rgba(200, 164, 106, 0.08)';

      // State
      let state = {
        paper: 'letter', gridSize: 18, unit: 'in', tool: 'select', 
        pages: [{ elements: [], orientation: 'portrait' }], currentPageIndex: 0,
        history: [], zoom: 1.0, globalOpacity: 100,
        isDrawing: false, isDragging: false, isModifying: false, isFilling: false, isSelecting: false,
        activeKeys: {}, clipboard: [], editingIndex: null,
        startPoint: null, dragElementIndex: null, selectedIndices: [], modifyHandleType: null,
        gridVisible: true, rulerVisible: true, gridOffset: { x: 0, y: 0 },
        viewMode: 'canvas', bleedUnits: 1, orientation: 'portrait',
        bleedVisible: true,
        notebookLayout: { cols: 1, pageWidth: 280 },
        lastCanvasPageIndex: null,
        lastSnappedCoords: null,
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

      function normalizePage(page) {
        if (!page || typeof page !== 'object') return { elements: [], orientation: 'portrait' };
        if (!Array.isArray(page.elements)) page.elements = [];
        if (page.orientation !== 'portrait' && page.orientation !== 'landscape') page.orientation = 'portrait';
        return page;
      }

      function normalizePages() {
        state.pages = state.pages.map((page) => normalizePage(page));
      }

      function getActivePage() {
        normalizePages();
        const page = state.pages[state.currentPageIndex] || state.pages[0];
        return normalizePage(page);
      }

      function getPageOrientation(page = getActivePage()) {
        return normalizePage(page).orientation;
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

      function getPaperMetrics(orientation = state.orientation) {
        const config = PAPER_CONFIG[state.paper];
        const baseWidth = state.unit === 'in' ? config.width * PPI : config.width * PMM;
        const baseHeight = state.unit === 'in' ? config.height * PPI : config.height * PMM;
        return orientation === 'landscape'
          ? { width: baseHeight, height: baseWidth, unit: config.unit }
          : { width: baseWidth, height: baseHeight, unit: config.unit };
      }

      function getWorkspaceMetrics(page = getActivePage()) {
        const orientation = getPageOrientation(page);
        const { width, height } = getPaperMetrics(orientation);
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

      function getWorkspaceMetricsForPaper(orientation) {
        const { width, height } = getPaperMetrics(orientation);
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

      function rotatePointForOrientation(point, sourceMetrics, direction) {
        if (direction === 'cw') {
          return { x: sourceMetrics.height - point.y, y: point.x };
        }
        return { x: point.y, y: sourceMetrics.width - point.x };
      }

      function roundCoord(value) {
        return Math.round(Number(value) * 1000) / 1000;
      }

      function rotateElementForOrientation(el, fromOrientation, toOrientation) {
        if (!el || fromOrientation === toOrientation) return el;
        const sourceMetrics = getPaperMetrics(fromOrientation);
        const direction = fromOrientation === 'portrait' && toOrientation === 'landscape' ? 'cw' : 'ccw';

        const rotatePoint = (x, y) => rotatePointForOrientation({ x, y }, sourceMetrics, direction);
        const applyBox = (x, y, w, h) => {
          const corners = [
            rotatePoint(x, y),
            rotatePoint(x + w, y),
            rotatePoint(x + w, y + h),
            rotatePoint(x, y + h)
          ];
          const xs = corners.map((pt) => pt.x);
          const ys = corners.map((pt) => pt.y);
          return {
            x: roundCoord(Math.min(...xs)),
            y: roundCoord(Math.min(...ys)),
            w: roundCoord(Math.max(...xs) - Math.min(...xs)),
            h: roundCoord(Math.max(...ys) - Math.min(...ys))
          };
        };

        if (el.type === 'line') {
          const start = rotatePoint(el.x1, el.y1);
          const end = rotatePoint(el.x2, el.y2);
          el.x1 = roundCoord(start.x);
          el.y1 = roundCoord(start.y);
          el.x2 = roundCoord(end.x);
          el.y2 = roundCoord(end.y);
          return el;
        }

        if (el.type === 'rect' || el.type === 'text') {
          const box = applyBox(el.x, el.y, el.w, el.h);
          el.x = box.x;
          el.y = box.y;
          el.w = box.w;
          el.h = box.h;
          return el;
        }

        if (el.type === 'dot' || el.type === 'cross') {
          const pt = rotatePoint(el.x, el.y);
          el.x = roundCoord(pt.x);
          el.y = roundCoord(pt.y);
        }
        return el;
      }

      function rotatePageElements(page, fromOrientation, toOrientation) {
        if (!page || fromOrientation === toOrientation) return page;
        page.elements = page.elements.map((el) => rotateElementForOrientation(JSON.parse(JSON.stringify(el)), fromOrientation, toOrientation));
        page.orientation = toOrientation;
        return page;
      }

      function syncActivePageOrientation() {
        const page = getActivePage();
        state.orientation = getPageOrientation(page);
        updateOrientationButton();
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
          btn.classList.toggle('ring-teal-300', btnType === activeType);
        });
        document.querySelectorAll('.align-btn').forEach(btn => {
          const textAlign = el.style.textAlign || 'left';
          const active = btn.id === `align-${textAlign}`;
          btn.classList.toggle('bg-white', active);
          btn.classList.toggle('ring-1', active);
          btn.classList.toggle('ring-teal-300', active);
        });
        document.querySelectorAll('.valign-btn').forEach(btn => {
          const verticalAlign = el.style.verticalAlign || 'top';
          const active = btn.id === `valign-${verticalAlign}`;
          btn.classList.toggle('bg-white', active);
          btn.classList.toggle('ring-1', active);
          btn.classList.toggle('ring-teal-300', active);
        });
        document.querySelectorAll('.text-format-btn').forEach(btn => {
          const active =
            (btn.id === 'font-bold' && (el.style.fontWeight || '400') === '700') ||
            (btn.id === 'font-italic' && (el.style.fontStyle || 'normal') === 'italic');
          btn.classList.toggle('bg-white', active);
          btn.classList.toggle('ring-1', active);
          btn.classList.toggle('ring-teal-300', active);
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
          button.className = `quality-option w-full text-left px-4 py-3 border rounded-xl transition-all hover:border-teal-300 ${preset.id === selectedId ? 'active' : 'border-slate-200'}`;
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
        const isInch = state.unit === 'in';
        const originMode = getGridOriginMode();
        const sheetSize = size - (2 * bleedPx);
        const axisOrigin = originMode === 'center'
          ? bleedPx + (sheetSize / 2)
          : bleedPx;
        const imperialInches = isInch ? (state.gridSize / PPI) : 0;
        const isDecimalImperial = isInch && Math.abs(imperialInches - 0.2) < 0.01;

        // Normalize the ruler origin once, then let tick cadence vary by preset family.
        const tickStepPx = isMetricMode()
          ? state.gridSize / 2
          : (isDecimalImperial ? state.gridSize : (PPI / 4));
        const tickPositions = buildAxisTickPositions(size, axisOrigin, tickStepPx, originMode)
          .filter((i) => i >= -0.1 && i <= size + 0.1);

        if (isMetricMode()) {
          const majorStepUnits = state.gridSize / unitScale;
          tickPositions.forEach((i) => {
            const relMm = (i - axisOrigin) / unitScale;
            const absMm = Math.abs(relMm);
            const isZero = Math.abs(relMm) < 0.01;
            const isMajor = isZero || Math.abs((absMm / majorStepUnits) - Math.round(absMm / majorStepUnits)) < 0.01;
            const isHalf = !isMajor && Math.abs((absMm / (majorStepUnits / 2)) - Math.round(absMm / (majorStepUnits / 2))) < 0.01;
            const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
            const tickSize = isZero ? 20 : (isMajor ? 16 : 8);

            if (orientation === 'horizontal') {
              line.setAttribute("x1", i); line.setAttribute("x2", i);
              line.setAttribute("y1", rulerOffset - tickSize); line.setAttribute("y2", rulerOffset);
            } else {
              line.setAttribute("y1", i); line.setAttribute("y2", i);
              line.setAttribute("x1", rulerOffset - tickSize); line.setAttribute("x2", rulerOffset);
            }

            line.setAttribute("class", isZero ? "ruler-center-marker" : (isHalf ? "ruler-mini-tick" : "ruler-tick"));
            svg.appendChild(line);

            if (isMajor) {
              const text = document.createElementNS("http://www.w3.org/2000/svg", "text");
              text.setAttribute("class", "ruler-label");
              text.textContent = formatMetricValue(relMm);
              if (orientation === 'horizontal') {
                text.setAttribute("x", i + 2); text.setAttribute("y", 10);
              } else {
                text.setAttribute("x", 2); text.setAttribute("y", i + 8);
              }
              svg.appendChild(text);
            }
          });
          return;
        }

        if (isDecimalImperial) {
          const labelPositions = buildAxisTickPositions(size, axisOrigin, PPI / 2, originMode)
            .filter((i) => i >= -0.1 && i <= size + 0.1);
          tickPositions.forEach((i) => {
            const rel = (i - axisOrigin) / unitScale;
            const absRel = Math.abs(rel);
            const isZero = absRel < 0.01;
            const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
            const tickSize = isZero ? 20 : 7;

            if (orientation === 'horizontal') {
              line.setAttribute("x1", i); line.setAttribute("x2", i);
              line.setAttribute("y1", rulerOffset - tickSize); line.setAttribute("y2", rulerOffset);
            } else {
              line.setAttribute("y1", i); line.setAttribute("y2", i);
              line.setAttribute("x1", rulerOffset - tickSize); line.setAttribute("x2", rulerOffset);
            }

            line.setAttribute("class", isZero ? "ruler-center-marker" : "ruler-tick");
            svg.appendChild(line);
          });

          labelPositions.forEach((i) => {
            const rel = (i - axisOrigin) / unitScale;
            const absRel = Math.abs(rel);
            const isWhole = Math.abs(absRel - Math.round(absRel)) < 0.01;
            const isHalfLabel = !isWhole && Math.abs((absRel * 2) - Math.round(absRel * 2)) < 0.01;
            if (!isWhole && !isHalfLabel) return;
            const text = document.createElementNS("http://www.w3.org/2000/svg", "text");
            text.setAttribute("class", "ruler-label");
            text.textContent = isWhole ? Math.round(rel) : rel.toFixed(1);
            if (orientation === 'horizontal') {
              text.setAttribute("x", i + 2); text.setAttribute("y", 10);
            } else {
              text.setAttribute("x", 2); text.setAttribute("y", i + 8);
            }
            svg.appendChild(text);
          });
          return;
        }

        tickPositions.forEach((i) => {
          const rel = (i - axisOrigin) / unitScale;
          const absRel = Math.abs(rel);
          const isZero = absRel < 0.01;
          const isMajor = isZero || Math.abs(absRel - Math.round(absRel)) < 0.01;
          const isHalf = !isMajor && Math.abs((absRel * 2) - Math.round(absRel * 2)) < 0.01;
          const isQuarter = !isMajor && !isHalf && Math.abs((absRel * 4) - Math.round(absRel * 4)) < 0.01;
          const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
          const tickSize = isZero ? 20 : (isMajor ? 20 : (isHalf ? 16 : (isQuarter ? 8 : 7)));

          if (orientation === 'horizontal') {
            line.setAttribute("x1", i); line.setAttribute("x2", i);
            line.setAttribute("y1", rulerOffset - tickSize); line.setAttribute("y2", rulerOffset);
          } else {
            line.setAttribute("y1", i); line.setAttribute("y2", i);
            line.setAttribute("x1", rulerOffset - tickSize); line.setAttribute("x2", rulerOffset);
          }

          line.setAttribute("class", isZero ? "ruler-center-marker" : (isHalf ? "ruler-mini-tick" : "ruler-tick"));
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
        });
      }

      function renderGridLines(elG, w, h) {
        const step = state.gridSize;
        const bleedPx = state.bleedVisible ? state.bleedUnits * state.gridSize : 0;
        const sheetW = w - 2 * bleedPx;
        const sheetH = h - 2 * bleedPx;
        const originMode = getGridOriginMode();
        const origin = getCanvasPageOrigin(sheetW, sheetH, bleedPx, originMode);
        const xPositions = buildAxisTickPositions(w, origin.x, step, originMode);
        const yPositions = buildAxisTickPositions(h, origin.y, step, originMode);
        state.gridOffset = { x: origin.x % step, y: origin.y % step };

        xPositions.forEach((x) => {
          const l = document.createElementNS("http://www.w3.org/2000/svg", "line");
          l.setAttribute("x1", x); l.setAttribute("y1", 0); l.setAttribute("x2", x); l.setAttribute("y2", h);
          l.setAttribute("stroke", Math.abs(x - origin.x) < 0.1 ? "#cbd5e1" : "#f1f5f9");
          elG.appendChild(l);
        });

        yPositions.forEach((y) => {
          const l = document.createElementNS("http://www.w3.org/2000/svg", "line");
          l.setAttribute("x1", 0); l.setAttribute("y1", y); l.setAttribute("x2", w); l.setAttribute("y2", y);
          l.setAttribute("stroke", Math.abs(y - origin.y) < 0.1 ? "#cbd5e1" : "#f1f5f9");
          elG.appendChild(l);
        });

        // Emphasize the true page center independently of grid spacing.
        const centerX = bleedPx + (sheetW / 2);
        const centerY = bleedPx + (sheetH / 2);
        const centerV = document.createElementNS("http://www.w3.org/2000/svg", "line");
        centerV.setAttribute("x1", centerX); centerV.setAttribute("y1", 0);
        centerV.setAttribute("x2", centerX); centerV.setAttribute("y2", h);
        centerV.setAttribute("class", "grid-center-guide");
        elG.appendChild(centerV);

        const centerH = document.createElementNS("http://www.w3.org/2000/svg", "line");
        centerH.setAttribute("x1", 0); centerH.setAttribute("y1", centerY);
        centerH.setAttribute("x2", w); centerH.setAttribute("y2", centerY);
        centerH.setAttribute("class", "grid-center-guide");
        elG.appendChild(centerH);
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
        const pageAspects = state.pages.length ? state.pages.map((page) => {
          const metrics = getWorkspaceMetricsForPaper(getPageOrientation(page));
          return metrics.totalWidth / metrics.totalHeight;
        }) : [1];
        const pageAspect = Math.min(...pageAspects);
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

      function isMetricMode() {
        return state.unit === 'mm';
      }

      function getGridOriginMode() {
        return isMetricMode() ? 'center' : 'top-left';
      }

      // Normalize the page reference once so ruler, grid, and snapping all share the same origin.
      function getCanvasPageOrigin(pageWidth, pageHeight, bleedPx, originMode = getGridOriginMode()) {
        return originMode === 'center'
          ? { x: bleedPx + (pageWidth / 2), y: bleedPx + (pageHeight / 2) }
          : { x: bleedPx, y: bleedPx };
      }

      function toCanvasPoint(point, pageWidth, pageHeight, bleedPx) {
        const origin = getCanvasPageOrigin(pageWidth, pageHeight, bleedPx);
        return { x: origin.x + point.x, y: origin.y + point.y };
      }

      function getGridStepPx() {
        return isMetricMode() ? state.gridSize / 2 : state.gridSize / 2;
      }

      function formatMetricValue(value) {
        const rounded = Math.round(value);
        return Math.abs(value - rounded) < 0.001 ? `${rounded}` : `${Math.round(value * 10) / 10}`;
      }

      function updateCoordinateHud(point) {
        const hud = document.getElementById('coordDisplay');
        if (!hud) return;
        const nextPoint = point && Number.isFinite(point.x) && Number.isFinite(point.y) ? point : { x: 0, y: 0 };
        if (point && Number.isFinite(point.x) && Number.isFinite(point.y)) {
          state.lastSnappedCoords = { x: nextPoint.x, y: nextPoint.y };
        }
        if (isMetricMode()) {
          hud.innerText = `X: ${formatMetricValue(nextPoint.x / PMM)} mm Y: ${formatMetricValue(nextPoint.y / PMM)} mm`;
        } else {
          hud.innerText = `X: ${(nextPoint.x / PPI).toFixed(2)} in Y: ${(nextPoint.y / PPI).toFixed(2)} in`;
        }
      }

      // Normalize spacing once so grid lines and ruler ticks share the same step math.
      function buildAxisTickPositions(sizePx, originPx, stepPx, originMode = getGridOriginMode()) {
        const ticks = [];
        const addTick = (value) => {
          const rounded = Math.round(value * 100) / 100;
          if (!ticks.some((existing) => Math.abs(existing - rounded) < 0.001)) ticks.push(rounded);
        };

        if (originMode === 'center') {
          const maxDistance = Math.max(originPx, sizePx - originPx);
          for (let offset = 0; offset <= maxDistance + (stepPx / 2); offset += stepPx) {
            addTick(originPx - offset);
            if (offset > 0) addTick(originPx + offset);
          }
        } else {
          for (let pos = originPx; pos <= sizePx + (stepPx / 2); pos += stepPx) {
            addTick(pos);
          }
        }

        return ticks.sort((a, b) => a - b);
      }

      function setGridSelectValue(value) {
        const select = document.getElementById('gridSelect');
        if (select) select.value = value;
      }

      function syncGridToPaperUnit(nextUnit) {
        const currentValue = nextUnit === 'mm' ? state.gridSize / PMM : state.gridSize / PPI;
        if (nextUnit === 'mm') {
          if (![4, 5, 6].some((allowed) => Math.abs(currentValue - allowed) < 0.01)) {
            state.gridSize = 5 * PMM;
            setGridSelectValue('5mm');
          }
        } else if (![0.25, 0.2, 0.28125, 0.34375].some((allowed) => Math.abs(currentValue - allowed) < 0.001)) {
          state.gridSize = 0.25 * PPI;
          setGridSelectValue('0.25');
        }
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
        normalizePages();
        window.setPaper('letter');
        syncActivePageOrientation();
        window.setupEventListeners();
        window.updateGridSize('0.25');
        setTimeout(window.fitAuto, 300);
      };

      window.setPaper = function(type) {
        saveHistory(); state.paper = type; state.unit = PAPER_CONFIG[type].unit;
        syncGridToPaperUnit(state.unit);
        state.lastSnappedCoords = null;
        updateCoordinateHud();
        syncActivePageOrientation();
        window.renderWorkspace();
      };

      window.toggleOrientation = function() {
        const page = getActivePage();
        const fromOrientation = getPageOrientation(page);
        const toOrientation = fromOrientation === 'portrait' ? 'landscape' : 'portrait';
        saveHistory();
        rotatePageElements(page, fromOrientation, toOrientation);
        state.orientation = toOrientation;
        updateOrientationButton();
        const viewport = document.getElementById('viewport');
        if (viewport) {
          viewport.scrollTop = 0;
          viewport.scrollLeft = 0;
        }
        state.lastSnappedCoords = null;
        updateCoordinateHud();
        window.renderWorkspace();
      };

      window.updateGridSize = function(val) {
        if (val.endsWith('mm')) {
          const metricMm = [4, 5, 6].includes(parseFloat(val)) ? parseFloat(val) : 5;
          state.gridSize = metricMm * PMM;
        } else {
          state.gridSize = parseFloat(val) * PPI;
        }
        window.renderWorkspace();
      };

      window.toggleGrid = function() {
        state.gridVisible = !state.gridVisible;
        const btn = document.getElementById('gridToggle');
        if(btn) { btn.classList.toggle('bg-[#C8A46A]', state.gridVisible); btn.classList.toggle('text-white', state.gridVisible); }
        window.renderWorkspace();
      };

      window.toggleRulers = function() {
        state.rulerVisible = !state.rulerVisible;
        const btn = document.getElementById('rulerToggle');
        if(btn) { btn.classList.toggle('bg-[#C8A46A]', state.rulerVisible); btn.classList.toggle('text-white', state.rulerVisible); }
        window.renderWorkspace();
      };

      window.toggleBleed = function() {
        state.bleedVisible = !state.bleedVisible;
        const btn = document.getElementById('bleedToggle');
        if(btn) { btn.classList.toggle('bg-[#C8A46A]', state.bleedVisible); btn.classList.toggle('text-white', state.bleedVisible); }
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
          normalizePages();
          state.currentPageIndex = Math.min(state.currentPageIndex, state.pages.length - 1);
          state.selectedIndices = [];
          state.lastSnappedCoords = null;
          updateCoordinateHud();
          updatePropsPanel();
          updateUndoButton();
          syncActivePageOrientation();
          window.renderWorkspace();
        }
      };
      window.prevPage = () => { if (state.currentPageIndex > 0) { state.currentPageIndex--; state.lastSnappedCoords = null; updateCoordinateHud(); syncActivePageOrientation(); window.renderWorkspace(); } };
      window.nextPage = () => { if (state.currentPageIndex < state.pages.length - 1) { state.currentPageIndex++; state.lastSnappedCoords = null; updateCoordinateHud(); syncActivePageOrientation(); window.renderWorkspace(); } };
      window.addPage = () => { saveHistory(); state.pages.push({ elements: [], orientation: state.orientation }); state.currentPageIndex = state.pages.length - 1; state.lastSnappedCoords = null; updateCoordinateHud(); syncActivePageOrientation(); window.renderWorkspace(); };
      window.copyPage = () => {
        saveHistory();
        const clonedPage = JSON.parse(JSON.stringify(state.pages[state.currentPageIndex]));
        normalizePage(clonedPage);
        state.pages.splice(state.currentPageIndex + 1, 0, clonedPage);
        state.currentPageIndex += 1;
        state.lastSnappedCoords = null;
        updateCoordinateHud();
        syncActivePageOrientation();
        window.renderWorkspace();
      };
      window.deletePage = () => { if (state.pages.length > 1) { saveHistory(); state.pages.splice(state.currentPageIndex, 1); state.currentPageIndex = Math.max(0, state.currentPageIndex - 1); state.lastSnappedCoords = null; updateCoordinateHud(); syncActivePageOrientation(); window.renderWorkspace(); } };
      
      window.setTool = (t) => {
        state.tool = t;
        document.querySelectorAll('.tool-btn').forEach(b => {
          const isActive = b.id === `tool-${t}`;
        b.classList.toggle('bg-[#C8A46A]', isActive); b.classList.toggle('text-white', isActive);
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
          const firstPage = normalizePage(JSON.parse(JSON.stringify(state.pages[0] || { elements: [], orientation: 'portrait' })));
          const firstOrientation = getPageOrientation(firstPage);
          const { width: pageWidth, height: pageHeight } = getPaperMetrics(firstOrientation);
          const pdf = new jsPDFCtor({
            orientation: pageWidth >= pageHeight ? 'landscape' : 'portrait',
            unit: 'pt',
            format: [pageWidth, pageHeight],
            compress: true
          });

          const svgNs = "http://www.w3.org/2000/svg";
          const renderPageToDataUrl = async (page) => {
            const pageOrientation = getPageOrientation(page);
            const { width: pageWidth, height: pageHeight } = getPaperMetrics(pageOrientation);
            const toPdf = (point) => toCanvasPoint(point, pageWidth, pageHeight, 0);
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
                const p1 = toPdf({ x: el.x1, y: el.y1 });
                const p2 = toPdf({ x: el.x2, y: el.y2 });
                node = document.createElementNS(svgNs, "line");
                node.setAttribute("x1", `${p1.x}`);
                node.setAttribute("y1", `${p1.y}`);
                node.setAttribute("x2", `${p2.x}`);
                node.setAttribute("y2", `${p2.y}`);
                node.setAttribute("fill", "none");
              } else if (el.type === 'rect') {
                const p = toPdf({ x: el.x, y: el.y });
                node = document.createElementNS(svgNs, "rect");
                node.setAttribute("x", `${p.x}`);
                node.setAttribute("y", `${p.y}`);
                node.setAttribute("width", `${el.w}`);
                node.setAttribute("height", `${el.h}`);
                node.setAttribute("fill", "none");
              } else if (el.type === 'dot') {
                const p = toPdf({ x: el.x, y: el.y });
                node = document.createElementNS(svgNs, "circle");
                node.setAttribute("cx", `${p.x}`);
                node.setAttribute("cy", `${p.y}`);
                node.setAttribute("r", `${Math.max(0.5, Math.min(5, weight))}`);
                node.setAttribute("fill", "#334155");
              } else if (el.type === 'cross') {
                node = document.createElementNS(svgNs, "path");
                const size = 4;
                const p = toPdf({ x: el.x, y: el.y });
                node.setAttribute("d", `M ${p.x-size} ${p.y} H ${p.x+size} M ${p.x} ${p.y-size} V ${p.y+size}`);
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
                const p = toPdf({ x: el.x, y: el.y });
                const startY = verticalAlign === 'center'
                  ? p.y + paddingY + Math.max(0, (el.h - paddingY * 2 - contentHeight) / 2)
                  : verticalAlign === 'bottom'
                    ? p.y + el.h - paddingY - contentHeight
                    : p.y + paddingY;
                const textNode = document.createElementNS(svgNs, "text");
                textNode.setAttribute("font-family", "Helvetica, Arial, sans-serif");
                textNode.setAttribute("font-size", `${fontSize}`);
                textNode.setAttribute("font-weight", el.style?.fontWeight || '400');
                textNode.setAttribute("font-style", el.style?.fontStyle || 'normal');
                textNode.setAttribute("fill", "#334155");
                const xPos = textAlign === 'center'
                  ? p.x + (el.w / 2)
                  : textAlign === 'right'
                    ? p.x + el.w - paddingX
                    : p.x + paddingX;
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
            const page = normalizePage(state.pages[index]);
            const pageOrientation = getPageOrientation(page);
            const { width: nextPageWidth, height: nextPageHeight } = getPaperMetrics(pageOrientation);
            if (index > 0) pdf.addPage([nextPageWidth, nextPageHeight], pageOrientation);
            const pageImage = await renderPageToDataUrl(page);
            pdf.addImage(pageImage, exportQuality.mime === 'image/png' ? 'PNG' : 'JPEG', 0, 0, nextPageWidth, nextPageHeight, undefined, exportQuality.compression);
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
        const originMode = getGridOriginMode();
        const origin = getCanvasPageOrigin(sheetW, sheetH, bleedPx, originMode);
        const fx = Math.round((rawX - origin.x) / step) * step;
        const fy = Math.round((rawY - origin.y) / step) * step;

        const dot = document.getElementById(`snapDot-${state.currentPageIndex}`);
        if (dot) {
          const dotPoint = isMetricMode() ? toCanvasPoint({ x: fx, y: fy }, sheetW, sheetH, bleedPx) : { x: fx + origin.x, y: fy + origin.y };
          dot.setAttribute('cx', dotPoint.x); dot.setAttribute('cy', dotPoint.y);
          if (state.tool !== 'select' || state.isDrawing || state.isDragging) dot.classList.remove('opacity-0');
          else dot.classList.add('opacity-0');
        }
        const snapped = { x: fx, y: fy };
        updateCoordinateHud(snapped);
        return snapped;
      };

      window.renderWorkspace = function() {
        const workspace = document.getElementById('workspace');
        if (!workspace) return;
        workspace.innerHTML = '';
        workspace.style.gridTemplateColumns = '';
        const activePage = getActivePage();
        state.orientation = getPageOrientation(activePage);
        updateOrientationButton();
        workspace.style.alignSelf = state.viewMode === 'canvas' && state.pages.length === 1 ? 'center' : (state.viewMode === 'canvas' ? 'flex-start' : 'center');
        const pageChangeDirection = state.viewMode === 'canvas' && state.lastCanvasPageIndex !== null
          ? (state.currentPageIndex > state.lastCanvasPageIndex ? 'next' : (state.currentPageIndex < state.lastCanvasPageIndex ? 'prev' : null))
          : null;
        const activeMetrics = getWorkspaceMetrics(activePage);

        updateNotebookLayout(activeMetrics.totalWidth, activeMetrics.totalHeight);

        state.pages.forEach((page, idx) => {
          normalizePage(page);
          const pageOrientation = getPageOrientation(page);
          const { pageWidth: w, pageHeight: h, bleedPx, rulerOffset, totalWidth: totalW, totalHeight: totalH } = getWorkspaceMetricsForPaper(pageOrientation);
          const toCanvas = (point) => toCanvasPoint(point, w, h, bleedPx);
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
            marquee.setAttribute("fill", ACCENT_SOFT); marquee.setAttribute("stroke", ACCENT);
          marquee.setAttribute("stroke-width", "1"); marquee.setAttribute("stroke-dasharray", "4");

          const snapDot = document.createElementNS("http://www.w3.org/2000/svg", "circle");
          snapDot.id = `snapDot-${idx}`; snapDot.setAttribute("r", "3.5"); snapDot.setAttribute("fill", ACCENT);
          snapDot.setAttribute("class", "opacity-0 pointer-events-none transition-opacity duration-75");

          const tLine = document.createElementNS("http://www.w3.org/2000/svg", "line");
          tLine.id = `tempLine-${idx}`; tLine.setAttribute("class", "hidden");
          tLine.setAttribute("stroke", ACCENT); tLine.setAttribute("stroke-width", "1.5"); tLine.setAttribute("stroke-dasharray", "4");

          const tRect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
          tRect.id = `tempRect-${idx}`; tRect.setAttribute("class", "hidden");
          tRect.setAttribute("fill", ACCENT_SOFT); tRect.setAttribute("stroke", ACCENT);
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
            if (el.type==='line') { const p1 = toCanvas({ x: el.x1, y: el.y1 }); const p2 = toCanvas({ x: el.x2, y: el.y2 }); node = document.createElementNS("http://www.w3.org/2000/svg", "line"); node.setAttribute("x1", p1.x); node.setAttribute("y1", p1.y); node.setAttribute("x2", p2.x); node.setAttribute("y2", p2.y); node.setAttribute("stroke", "#334155"); node.setAttribute("stroke-width", wt); if(el.style?.dash !== 'none') node.setAttribute("stroke-dasharray", el.style.dash); }
            else if (el.type==='rect') { const p = toCanvas({ x: el.x, y: el.y }); node = document.createElementNS("http://www.w3.org/2000/svg", "rect"); node.setAttribute("x", p.x); node.setAttribute("y", p.y); node.setAttribute("width", el.w); node.setAttribute("height", el.h); node.setAttribute("fill", "none"); node.setAttribute("stroke", "#334155"); node.setAttribute("stroke-width", wt); if(el.style?.dash !== 'none') node.setAttribute("stroke-dasharray", el.style.dash); }
            else if (el.type==='dot') { const p = toCanvas({ x: el.x, y: el.y }); node = document.createElementNS("http://www.w3.org/2000/svg", "circle"); node.setAttribute("cx", p.x); node.setAttribute("cy", p.y); node.setAttribute("r", Math.max(0.5, Math.min(5, wt))); node.setAttribute("fill", "#334155"); }
            else if (el.type==='cross') {
              node = document.createElementNS("http://www.w3.org/2000/svg", "path");
              const { x: cx, y: cy } = toCanvas({ x: el.x, y: el.y });
              const size = 4;
              node.setAttribute("d", `M ${cx-size} ${cy} H ${cx+size} M ${cx} ${cy-size} V ${cy+size}`);
              node.setAttribute("fill", "none");
              node.setAttribute("stroke", "#334155");
              node.setAttribute("stroke-width", wt || 1);
              node.setAttribute("stroke-linecap", "round");
            }
            else if (el.type==='text') {
              node = document.createElementNS("http://www.w3.org/2000/svg", "g");
              const textBoxRect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
              const p = toCanvas({ x: el.x, y: el.y });
              textBoxRect.setAttribute("x", p.x);
              textBoxRect.setAttribute("y", p.y);
              textBoxRect.setAttribute("width", el.w);
              textBoxRect.setAttribute("height", el.h);
              textBoxRect.setAttribute("fill", "#ffffff");
              textBoxRect.setAttribute("fill-opacity", "0.72");
              textBoxRect.setAttribute("stroke", ACCENT);
              textBoxRect.setAttribute("stroke-opacity", "0.18");
              textBoxRect.setAttribute("stroke-width", "1");
              textBoxRect.setAttribute("stroke-dasharray", "4 3");
              node.appendChild(textBoxRect);
              const textLayer = document.createElementNS("http://www.w3.org/2000/svg", "foreignObject");
              textLayer.setAttribute("x", p.x);
              textLayer.setAttribute("y", p.y);
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
            if (el.type==='line') { const p1 = toCanvas({ x: el.x1, y: el.y1 }); const p2 = toCanvas({ x: el.x2, y: el.y2 }); hit.setAttribute("x1", p1.x); hit.setAttribute("y1", p1.y); hit.setAttribute("x2", p2.x); hit.setAttribute("y2", p2.y); }
            else if (el.type==='rect' || el.type==='text') { const p = toCanvas({ x: el.x, y: el.y }); hit.setAttribute("x", p.x); hit.setAttribute("y", p.y); hit.setAttribute("width", el.w); hit.setAttribute("height", el.h); }
            else { const p = toCanvas({ x: el.x, y: el.y }); hit.setAttribute("cx", p.x); hit.setAttribute("cy", p.y); hit.setAttribute("r", 10); }
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
              const canvasHandles = {
                center: toCanvas({ x: handles.center.x - bleedPx, y: handles.center.y - bleedPx }),
                edges: handles.edges.map((pt) => toCanvas({ x: pt.x - bleedPx, y: pt.y - bleedPx }))
              };
              canvasHandles.edges.forEach((pt, edgeIndex) => {
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
                  state.modifyHandleType = { kind: 'edge', index: edgeIndex };
                  state.startPoint = window.getSnappedCoords(e);
                };
                gElements.appendChild(edgeHandle);
              });

              const centerHandle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
              centerHandle.setAttribute("cx", canvasHandles.center.x);
              centerHandle.setAttribute("cy", canvasHandles.center.y);
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
              const canvasGroupCenter = toCanvas({ x: groupCenter.x - bleedPx, y: groupCenter.y - bleedPx });
              const centerHandle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
              centerHandle.setAttribute("cx", canvasGroupCenter.x);
              centerHandle.setAttribute("cy", canvasGroupCenter.y);
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
                    const marqueeEl = document.getElementById(`marquee-${state.currentPageIndex}`);
                    if (marqueeEl) {
                      marqueeEl.classList.remove('hidden');
                      const startCanvas = toCanvas(c);
                      marqueeEl.setAttribute('x', startCanvas.x);
                      marqueeEl.setAttribute('y', startCanvas.y);
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
        updateCoordinateHud(state.lastSnappedCoords);
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
          if (state.isSelecting) {
            const marqueeEl = document.getElementById(`marquee-${state.currentPageIndex}`);
            if (marqueeEl) {
              marqueeEl.classList.remove('hidden');
              const page = getActivePage();
              const { pageWidth: w, pageHeight: h, bleedPx } = getWorkspaceMetricsForPaper(getPageOrientation(page));
              const left = Math.min(state.startPoint.x, c.x);
              const top = Math.min(state.startPoint.y, c.y);
              const right = Math.max(state.startPoint.x, c.x);
              const bottom = Math.max(state.startPoint.y, c.y);
              const startCanvas = toCanvasPoint({ x: left, y: top }, w, h, bleedPx);
              const endCanvas = toCanvasPoint({ x: right, y: bottom }, w, h, bleedPx);
              marqueeEl.setAttribute('x', startCanvas.x);
              marqueeEl.setAttribute('y', startCanvas.y);
              marqueeEl.setAttribute('width', endCanvas.x - startCanvas.x);
              marqueeEl.setAttribute('height', endCanvas.y - startCanvas.y);
            }
            return;
          }
          if (state.isDrawing) {
            const tl = document.getElementById(`tempLine-${state.currentPageIndex}`);
            const tr = document.getElementById(`tempRect-${state.currentPageIndex}`);
            if (state.tool === 'line' && tl && tr) {
              const page = getActivePage();
              const { pageWidth: w, pageHeight: h, bleedPx } = getWorkspaceMetricsForPaper(getPageOrientation(page));
              const startCanvas = toCanvasPoint(state.startPoint, w, h, bleedPx);
              const currentCanvas = toCanvasPoint(c, w, h, bleedPx);
              tl.classList.remove('hidden');
              tl.setAttribute('x1', startCanvas.x);
              tl.setAttribute('y1', startCanvas.y);
              tl.setAttribute('x2', currentCanvas.x);
              tl.setAttribute('y2', currentCanvas.y);
              if (state.isFilling) {
                tr.classList.remove('hidden');
                const left = Math.min(state.startPoint.x, c.x);
                const top = Math.min(state.startPoint.y, c.y);
                const right = Math.max(state.startPoint.x, c.x);
                const bottom = Math.max(state.startPoint.y, c.y);
                const start = toCanvasPoint({ x: left, y: top }, w, h, bleedPx);
                const end = toCanvasPoint({ x: right, y: bottom }, w, h, bleedPx);
                tr.setAttribute('x', start.x);
                tr.setAttribute('y', start.y);
                tr.setAttribute('width', end.x - start.x);
                tr.setAttribute('height', end.y - start.y);
              } else {
                tr.classList.add('hidden');
              }
            } 
            else if ((state.tool === 'rect' || state.tool === 'dot' || state.tool === 'cross' || state.tool === 'label') && tr) {
              const page = getActivePage();
              const { pageWidth: w, pageHeight: h, bleedPx } = getWorkspaceMetricsForPaper(getPageOrientation(page));
              tr.classList.remove('hidden');
              const left = Math.min(state.startPoint.x, c.x);
              const top = Math.min(state.startPoint.y, c.y);
              const right = Math.max(state.startPoint.x, c.x);
              const bottom = Math.max(state.startPoint.y, c.y);
              const start = toCanvasPoint({ x: left, y: top }, w, h, bleedPx);
              const end = toCanvasPoint({ x: right, y: bottom }, w, h, bleedPx);
              tr.setAttribute('x', start.x);
              tr.setAttribute('y', start.y);
              tr.setAttribute('width', end.x - start.x);
              tr.setAttribute('height', end.y - start.y);
            }
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

