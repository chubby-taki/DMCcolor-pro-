// DMC Color Matcher Pro - Legend-based Design
// Professional color matching with intelligent label placement

// State
let dmcColors = [];
let dmcPositions = {}; // DMC色番号の位置マッピング
let currentImage = null;
let imageDataCache = null;
let pickedColor = null;
let pickedX = 0;
let pickedY = 0;
let colorHistory = [];

// Viewport state
let viewportWidth = 0;
let viewportHeight = 0;
let imageScale = 1;
let zoom = 1;
let offsetX = 0;
let offsetY = 0;

// Interaction state
let isPanning = false;
let panStartX = 0;
let panStartY = 0;
let startOffsetX = 0;
let startOffsetY = 0;
let mouseDownX = 0;
let mouseDownY = 0;
let isDragging = false;
const DRAG_THRESHOLD = 5;

// Legend interaction
let hoveredLegendIndex = null;
let selectedLegendIndex = null;

// DOM Elements
const uploadArea = document.getElementById('uploadArea');
const imageInput = document.getElementById('imageInput');
const workspace = document.getElementById('workspace');
const imageCanvas = document.getElementById('imageCanvas');
const ctx = imageCanvas.getContext('2d');
const cursor = document.getElementById('cursor');
const sizeButtons = document.getElementById('sizeButtons');
let currentEyedropperSize = 1;
const colorSwatch = document.getElementById('colorSwatch');
const rgbValue = document.getElementById('rgbValue');
const hexValue = document.getElementById('hexValue');
const resultsSection = document.getElementById('resultsSection');
const matchList = document.getElementById('matchList');
const legendList = document.getElementById('legendList');
const clearHistoryBtn = document.getElementById('clearHistoryBtn');
const exportCsvBtn = document.getElementById('exportCsvBtn');
const exportPdfBtn = document.getElementById('exportPdfBtn');
const zoomLevel = document.getElementById('zoomLevel');
const minimapContainer = document.getElementById('minimapContainer');
const minimapCanvas = document.getElementById('minimapCanvas');
const minimapCtx = minimapCanvas.getContext('2d');
const minimapViewport = document.getElementById('minimapViewport');
const headerFull = document.getElementById('headerFull');
const headerCompact = document.getElementById('headerCompact');

async function init() {
    try {
        const response = await fetch('./dmc_master_data.json');
        const data = await response.json();
        dmcColors = data.colors;
        console.log(`Loaded ${dmcColors.length} DMC colors`);
    } catch (e) {
        console.error('Failed to load DMC data:', e);
    }

    // DMC位置データを読み込み
    try {
        const posResponse = await fetch('./dmc_position_data.json');
        const posData = await posResponse.json();
        dmcPositions = posData.positions;
        console.log(`Loaded ${Object.keys(dmcPositions).length} DMC position mappings`);
    } catch (e) {
        console.error('Failed to load DMC position data:', e);
    }

    workspace.style.display = 'none';

    colorHistory = [];
    localStorage.removeItem('dmcHistory');
    renderLegend();

    setupEventListeners();
}

function setupEventListeners() {
    // Upload
    uploadArea.addEventListener('dragover', (e) => { e.preventDefault(); uploadArea.classList.add('dragover'); });
    uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        if (e.dataTransfer.files.length) handleImageFile(e.dataTransfer.files[0]);
    });
    imageInput.addEventListener('change', (e) => {
        if (e.target.files.length) handleImageFile(e.target.files[0]);
    });

    const zoomSlider = document.getElementById('zoomSlider');

    zoomSlider.addEventListener('input', (e) => {
        const newZoom = parseFloat(e.target.value);
        setZoom(newZoom);
    });

    // Compact header - change image button
    const changeImageBtn = document.getElementById('changeImageBtn');
    const imageInputCompact = document.getElementById('imageInputCompact');
    if (changeImageBtn && imageInputCompact) {
        changeImageBtn.addEventListener('click', () => imageInputCompact.click());
        imageInputCompact.addEventListener('change', (e) => {
            if (e.target.files.length) handleImageFile(e.target.files[0]);
        });
    }

    sizeButtons.addEventListener('click', (e) => {
        if (e.target.classList.contains('size-btn')) {
            sizeButtons.querySelectorAll('.size-btn').forEach(btn => btn.classList.remove('active'));
            e.target.classList.add('active');
            currentEyedropperSize = parseInt(e.target.dataset.size);
        }
    });

    // Canvas events
    imageCanvas.addEventListener('mousedown', onMouseDown);
    imageCanvas.addEventListener('mousemove', onMouseMove);
    imageCanvas.addEventListener('mouseup', onMouseUp);
    imageCanvas.addEventListener('mouseleave', onMouseLeave);
    imageCanvas.addEventListener('wheel', onWheel, { passive: false });

    // Touch
    imageCanvas.addEventListener('touchstart', onTouchStart, { passive: false });
    imageCanvas.addEventListener('touchmove', onTouchMove, { passive: false });
    imageCanvas.addEventListener('touchend', onTouchEnd);

    // Buttons
    clearHistoryBtn.addEventListener('click', clearHistory);
    exportCsvBtn.addEventListener('click', exportToCSV);
    if (exportPdfBtn) exportPdfBtn.addEventListener('click', exportToPDF);

    // Brightness slider
    const brightnessSlider = document.getElementById('brightnessSlider');
    const brightnessValue = document.getElementById('brightnessValue');
    const brightnessResetBtn = document.getElementById('brightnessResetBtn');
    if (brightnessSlider) {
        brightnessSlider.addEventListener('input', () => {
            brightnessValue.textContent = brightnessSlider.value;
        });
        brightnessSlider.addEventListener('change', () => {
            if (pickedColor && dmcColors.length > 0) findMatches();
        });
        brightnessResetBtn.addEventListener('click', () => {
            brightnessSlider.value = 0;
            brightnessValue.textContent = '0';
            if (pickedColor && dmcColors.length > 0) findMatches();
        });
    }

    // Minimap
    minimapContainer.addEventListener('click', onMinimapClick);

    // Legend interaction
    legendList.addEventListener('mousemove', onLegendMouseMove);
    legendList.addEventListener('mouseleave', onLegendMouseLeave);
    legendList.addEventListener('click', onLegendClick);
}

function handleImageFile(file) {
    if (!file.type.startsWith('image/')) return;

    if (currentImage && colorHistory.length > 0) {
        if (confirm('新しい画像を読み込みます。現在の抽出履歴をクリアしますか？')) {
            colorHistory = [];
            localStorage.setItem('dmcHistory', JSON.stringify(colorHistory));
            renderLegend();
        }
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        const img = new Image();
        img.onload = () => {
            currentImage = img;
            const tempCanvas = document.createElement('canvas');
            tempCanvas.width = img.width;
            tempCanvas.height = img.height;
            const tempCtx = tempCanvas.getContext('2d');
            tempCtx.drawImage(img, 0, 0);
            imageDataCache = tempCtx.getImageData(0, 0, img.width, img.height);

            // Switch to compact header
            headerFull.style.display = 'none';
            uploadArea.style.display = 'none';
            headerCompact.style.display = 'flex';
            workspace.style.display = 'flex';
            setupViewport();
            pickedColor = null;
        };
        img.onerror = () => {
            alert('画像の読み込みに失敗しました。');
            console.error('Failed to load image');
        };
        img.src = e.target.result;
    };
    reader.onerror = () => {
        alert('ファイルの読み込みに失敗しました。');
        console.error('Failed to read file');
    };
    reader.readAsDataURL(file);
}

function setupViewport() {
    const container = imageCanvas.parentElement;

    // Get actual container dimensions
    const containerRect = container.getBoundingClientRect();
    viewportWidth = containerRect.width;

    // Calculate optimal viewport height based on image aspect ratio
    const imageAspect = currentImage.width / currentImage.height;

    let maxHeight;
    if (window.innerWidth >= 1024) {
        maxHeight = 600;
    } else if (window.innerWidth >= 768) {
        maxHeight = 500;
    } else {
        maxHeight = 350;
    }

    // Calculate height to maintain image aspect ratio
    if (imageAspect >= 1) {
        // Landscape or square image
        viewportHeight = Math.min(viewportWidth / imageAspect, maxHeight);
    } else {
        // Portrait image
        viewportHeight = Math.min(maxHeight, viewportWidth / imageAspect);
        // If calculated height is too large, recalculate based on max height
        if (viewportHeight > maxHeight) {
            viewportHeight = maxHeight;
        }
    }

    // Set canvas size to viewport size
    imageCanvas.width = viewportWidth;
    imageCanvas.height = viewportHeight;

    // Calculate scale to fit image in viewport
    const scaleX = viewportWidth / currentImage.width;
    const scaleY = viewportHeight / currentImage.height;
    imageScale = Math.min(scaleX, scaleY, 1);

    initMinimap();
    resetView();
}

function resetView() {
    zoom = 1;
    const scaledW = currentImage.width * imageScale;
    const scaledH = currentImage.height * imageScale;
    offsetX = (viewportWidth - scaledW) / 2;
    offsetY = (viewportHeight - scaledH) / 2;
    clampOffsets();
    render();
    updateZoomDisplay();
}

function applyZoom(factor, centerX, centerY) {
    const oldZoom = zoom;
    zoom = Math.max(0.5, Math.min(16, zoom * factor));
    if (centerX !== undefined && centerY !== undefined) {
        offsetX = centerX - (centerX - offsetX) * (zoom / oldZoom);
        offsetY = centerY - (centerY - offsetY) * (zoom / oldZoom);
    }
    clampOffsets();
    render();
    updateZoomDisplay();

    const slider = document.getElementById('zoomSlider');
    if (slider) slider.value = zoom;
}

function updateZoomDisplay() {
    if (zoomLevel) zoomLevel.textContent = `${Math.round(zoom * 100)}%`;
}

function setZoom(newZoom) {
    zoom = newZoom;
    zoomLevel.textContent = `${Math.round(zoom * 100)}%`;

    const slider = document.getElementById('zoomSlider');
    if (slider && Math.abs(parseFloat(slider.value) - zoom) > 0.01) {
        slider.value = zoom;
    }

    render();
}

function render() {
    if (!currentImage) return;
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, viewportWidth, viewportHeight);
    ctx.save();
    ctx.translate(offsetX, offsetY);
    ctx.scale(zoom * imageScale, zoom * imageScale);
    ctx.drawImage(currentImage, 0, 0);
    ctx.restore();

    // Draw pins with enhanced design
    const dmcToNumber = new Map();
    let numberCounter = 1;
    [...colorHistory].reverse().forEach(entry => {
        if (!dmcToNumber.has(entry.dmc_id)) {
            dmcToNumber.set(entry.dmc_id, numberCounter++);
        }
    });

    [...colorHistory].reverse().forEach((entry) => {
        if (entry.x === undefined || entry.y === undefined) return;

        const px = (entry.x * zoom * imageScale) + offsetX;
        const py = (entry.y * zoom * imageScale) + offsetY;

        if (px < -10 || px > viewportWidth + 10 || py < -10 || py > viewportHeight + 10) return;

        const pinNumber = dmcToNumber.get(entry.dmc_id);
        const isHovered = hoveredLegendIndex !== null && pinNumber === hoveredLegendIndex + 1;

        ctx.save();

        // Enhanced pin marker with color
        const dmcColor = entry.hex;

        // Outer glow
        ctx.shadowColor = isHovered ? 'rgba(255, 255, 255, 1)' : 'rgba(255, 255, 255, 0.8)';
        ctx.shadowBlur = isHovered ? 12 : 6;

        // White outer circle
        ctx.beginPath();
        ctx.arc(px, py, isHovered ? 11 : 9, 0, Math.PI * 2);
        ctx.fillStyle = 'white';
        ctx.fill();

        // DMC color inner circle
        ctx.shadowColor = 'rgba(0, 0, 0, 0.3)';
        ctx.shadowBlur = 2;
        ctx.beginPath();
        ctx.arc(px, py, isHovered ? 9 : 7, 0, Math.PI * 2);
        ctx.fillStyle = dmcColor;
        ctx.fill();

        // White border
        ctx.strokeStyle = 'white';
        ctx.lineWidth = 2;
        ctx.stroke();

        // Number (white text)
        ctx.shadowColor = 'rgba(0, 0, 0, 0.8)';
        ctx.shadowBlur = 3;
        ctx.fillStyle = 'white';
        ctx.font = `bold ${isHovered ? 16 : 14}px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif`;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(pinNumber, px, py);

        ctx.restore();
    });

    clampOffsets();
    imageCanvas.style.cursor = isPanning ? 'grabbing' : 'crosshair';

    updateMinimapViewport();
}

function clampOffsets() {
    if (!currentImage) return;
    const scaledW = currentImage.width * imageScale * zoom;
    const scaledH = currentImage.height * imageScale * zoom;

    if (scaledW <= viewportWidth) {
        offsetX = (viewportWidth - scaledW) / 2;
    } else {
        const minX = viewportWidth - scaledW;
        if (offsetX > 0) offsetX = 0;
        if (offsetX < minX) offsetX = minX;
    }

    if (scaledH <= viewportHeight) {
        offsetY = (viewportHeight - scaledH) / 2;
    } else {
        const minY = viewportHeight - scaledH;
        if (offsetY > 0) offsetY = 0;
        if (offsetY < minY) offsetY = minY;
    }
}

function canvasToImage(canvasX, canvasY) {
    const imgX = (canvasX - offsetX) / (zoom * imageScale);
    const imgY = (canvasY - offsetY) / (zoom * imageScale);
    return { x: imgX, y: imgY };
}

// Minimap functions
function initMinimap() {
    if (!currentImage) return;

    const containerWidth = 120;
    const containerHeight = 90;
    const imgAspect = currentImage.width / currentImage.height;
    const containerAspect = containerWidth / containerHeight;

    let mapW, mapH;
    if (imgAspect > containerAspect) {
        mapW = containerWidth;
        mapH = containerWidth / imgAspect;
    } else {
        mapH = containerHeight;
        mapW = containerHeight * imgAspect;
    }

    minimapCanvas.width = mapW;
    minimapCanvas.height = mapH;
    minimapCanvas.style.width = mapW + 'px';
    minimapCanvas.style.height = mapH + 'px';

    minimapCanvas.style.position = 'absolute';
    minimapCanvas.style.left = ((containerWidth - mapW) / 2) + 'px';
    minimapCanvas.style.top = ((containerHeight - mapH) / 2) + 'px';

    minimapCtx.drawImage(currentImage, 0, 0, mapW, mapH);

    minimapContainer.classList.add('visible');
}

function updateMinimapViewport() {
    if (!currentImage || !minimapCanvas.width) return;

    const mapW = minimapCanvas.width;
    const mapH = minimapCanvas.height;

    const visibleLeft = -offsetX / (imageScale * zoom);
    const visibleTop = -offsetY / (imageScale * zoom);
    const visibleWidth = viewportWidth / (imageScale * zoom);
    const visibleHeight = viewportHeight / (imageScale * zoom);

    const scaleToMinimap = mapW / currentImage.width;

    let vpLeft = visibleLeft * scaleToMinimap;
    let vpTop = visibleTop * scaleToMinimap;
    let vpWidth = visibleWidth * scaleToMinimap;
    let vpHeight = visibleHeight * scaleToMinimap;

    vpLeft = Math.max(0, vpLeft);
    vpTop = Math.max(0, vpTop);
    vpWidth = Math.min(mapW - vpLeft, vpWidth);
    vpHeight = Math.min(mapH - vpTop, vpHeight);

    const containerWidth = 120;
    const containerHeight = 90;
    const canvasOffsetLeft = (containerWidth - mapW) / 2;
    const canvasOffsetTop = (containerHeight - mapH) / 2;

    minimapViewport.style.left = (canvasOffsetLeft + vpLeft) + 'px';
    minimapViewport.style.top = (canvasOffsetTop + vpTop) + 'px';
    minimapViewport.style.width = vpWidth + 'px';
    minimapViewport.style.height = vpHeight + 'px';

    if (vpWidth >= mapW - 2 && vpHeight >= mapH - 2) {
        minimapViewport.style.display = 'none';
    } else {
        minimapViewport.style.display = 'block';
    }
}

function onMinimapClick(e) {
    if (!currentImage) return;

    const rect = minimapCanvas.getBoundingClientRect();
    const clickX = e.clientX - rect.left;
    const clickY = e.clientY - rect.top;

    const imgX = (clickX / minimapCanvas.width) * currentImage.width;
    const imgY = (clickY / minimapCanvas.height) * currentImage.height;

    offsetX = viewportWidth / 2 - imgX * imageScale * zoom;
    offsetY = viewportHeight / 2 - imgY * imageScale * zoom;

    clampOffsets();
    render();
    updateMinimapViewport();
}

// Mouse Events
function onMouseDown(e) {
    const rect = imageCanvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    mouseDownX = e.clientX;
    mouseDownY = e.clientY;
    isDragging = false;

    panStartX = e.clientX;
    panStartY = e.clientY;
    startOffsetX = offsetX;
    startOffsetY = offsetY;
}

function onMouseMove(e) {
    const rect = imageCanvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    if (e.buttons === 1) {
        e.preventDefault();
        const dx = e.clientX - mouseDownX;
        const dy = e.clientY - mouseDownY;
        const distance = Math.sqrt(dx * dx + dy * dy);

        if (distance > DRAG_THRESHOLD) {
            if (!isDragging) {
                isDragging = true;
                isPanning = true;
                panStartX = e.clientX;
                panStartY = e.clientY;
                startOffsetX = offsetX;
                startOffsetY = offsetY;
                imageCanvas.style.cursor = 'grabbing';
            }
            offsetX = startOffsetX + (e.clientX - panStartX);
            offsetY = startOffsetY + (e.clientY - panStartY);
            clampOffsets();
            render();
        }
    } else {
        showCursor(x, y);
    }
}

function onMouseUp(e) {
    const rect = imageCanvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    if (!isDragging) {
        pickColor(x, y);
    }

    isPanning = false;
    isDragging = false;
    imageCanvas.style.cursor = 'crosshair';
}

function onMouseLeave() {
    isPanning = false;
    cursor.style.display = 'none';
}

function onWheel(e) {
    e.preventDefault();
    const rect = imageCanvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    const factor = e.deltaY > 0 ? 0.9 : 1.1;
    applyZoom(factor, x, y);
}

// Touch Events
let touchStartDist = 0;
let touchStartZoom = 1;
let touchStartX = 0;
let touchStartY = 0;
let isTouchDragging = false;

function onTouchStart(e) {
    if (e.touches.length === 2) {
        e.preventDefault();
        touchStartDist = getTouchDist(e.touches);
        touchStartZoom = zoom;
        isTouchDragging = true;
    } else if (e.touches.length === 1) {
        touchStartX = e.touches[0].clientX;
        touchStartY = e.touches[0].clientY;
        isTouchDragging = false;

        panStartX = e.touches[0].clientX;
        panStartY = e.touches[0].clientY;
        startOffsetX = offsetX;
        startOffsetY = offsetY;
    }
}

function onTouchMove(e) {
    if (e.touches.length === 2) {
        e.preventDefault();
        const dist = getTouchDist(e.touches);
        const newZoom = Math.max(0.5, Math.min(16, touchStartZoom * (dist / touchStartDist)));

        const rect = imageCanvas.getBoundingClientRect();
        const centerX = ((e.touches[0].clientX + e.touches[1].clientX) / 2) - rect.left;
        const centerY = ((e.touches[0].clientY + e.touches[1].clientY) / 2) - rect.top;

        const oldZoom = zoom;
        zoom = newZoom;

        offsetX = centerX - (centerX - offsetX) * (zoom / oldZoom);
        offsetY = centerY - (centerY - offsetY) * (zoom / oldZoom);

        clampOffsets();
        render();
        updateZoomDisplay();

        const slider = document.getElementById('zoomSlider');
        if (slider) slider.value = zoom;
    } else if (e.touches.length === 1) {
        const dx = e.touches[0].clientX - touchStartX;
        const dy = e.touches[0].clientY - touchStartY;
        const distance = Math.sqrt(dx * dx + dy * dy);

        if (distance > DRAG_THRESHOLD) {
            e.preventDefault();
            if (!isTouchDragging) {
                isTouchDragging = true;
                isPanning = true;
                panStartX = e.touches[0].clientX;
                panStartY = e.touches[0].clientY;
                startOffsetX = offsetX;
                startOffsetY = offsetY;
            }
            offsetX = startOffsetX + (e.touches[0].clientX - panStartX);
            offsetY = startOffsetY + (e.touches[0].clientY - panStartY);
            clampOffsets();
            render();
        }
    }
}

function onTouchEnd(e) {
    if (!isTouchDragging && e.changedTouches.length === 1) {
        const rect = imageCanvas.getBoundingClientRect();
        const x = e.changedTouches[0].clientX - rect.left;
        const y = e.changedTouches[0].clientY - rect.top;
        pickColor(x, y);
    }
    isPanning = false;
    isTouchDragging = false;
}

function getTouchDist(touches) {
    const dx = touches[0].clientX - touches[1].clientX;
    const dy = touches[0].clientY - touches[1].clientY;
    return Math.sqrt(dx * dx + dy * dy);
}

// Cursor preview
function showCursor(x, y) {
    const size = currentEyedropperSize;
    const actualSize = size * zoom * imageScale;
    const minSize = 8 + (size * 4);
    const displaySize = Math.max(actualSize, minSize);
    cursor.style.display = 'block';
    cursor.style.width = `${displaySize}px`;
    cursor.style.height = `${displaySize}px`;
    cursor.style.left = `${x - displaySize / 2}px`;
    cursor.style.top = `${y - displaySize / 2}px`;
    const color = getColorAt(x, y, size);
    if (color) {
        cursor.style.backgroundColor = `rgb(${color.r}, ${color.g}, ${color.b})`;
    }
}

// Pick color at canvas position
function pickColor(canvasX, canvasY) {
    const size = currentEyedropperSize;
    const color = getColorAt(canvasX, canvasY, size);
    if (color) {
        const imgCoords = canvasToImage(canvasX, canvasY);
        pickedX = Math.round(imgCoords.x);
        pickedY = Math.round(imgCoords.y);
        pickedColor = color;
        colorSwatch.style.backgroundColor = `rgb(${color.r}, ${color.g}, ${color.b})`;
        rgbValue.textContent = `${color.r}, ${color.g}, ${color.b}`;
        hexValue.textContent = `${color.hex}`;
        // Auto-search DMC matches
        if (dmcColors.length > 0) findMatches();
    }
}

// Get color from cached image data
function getColorAt(canvasX, canvasY, size) {
    if (!currentImage || !imageDataCache) return null;
    const imgCoords = canvasToImage(canvasX, canvasY);
    const imgX = Math.round(imgCoords.x);
    const imgY = Math.round(imgCoords.y);
    if (imgX < 0 || imgX >= currentImage.width || imgY < 0 || imgY >= currentImage.height) {
        return null;
    }

    const halfSize = Math.floor(size / 2);
    const startX = Math.max(0, imgX - halfSize);
    const startY = Math.max(0, imgY - halfSize);
    const endX = Math.min(currentImage.width, imgX + halfSize + 1);
    const endY = Math.min(currentImage.height, imgY + halfSize + 1);

    const data = imageDataCache.data;
    const width = currentImage.width;
    let r = 0, g = 0, b = 0, count = 0;

    for (let y = startY; y < endY; y++) {
        for (let x = startX; x < endX; x++) {
            const idx = (y * width + x) * 4;
            r += data[idx];
            g += data[idx + 1];
            b += data[idx + 2];
            count++;
        }
    }

    if (count === 0) return null;
    r = Math.round(r / count);
    g = Math.round(g / count);
    b = Math.round(b / count);
    const hex = '#' + [r, g, b].map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
    return { r, g, b, hex };
}

// Find DMC matches
function findMatches() {
    if (!pickedColor || dmcColors.length === 0) return;
    const targetLab = ColorMatcher.rgbToLab(pickedColor.r, pickedColor.g, pickedColor.b);

    // 明度補正を適用
    const brightnessSlider = document.getElementById('brightnessSlider');
    const offset = brightnessSlider ? parseFloat(brightnessSlider.value) : 0;
    const adjustedLab = [targetLab[0] + offset, targetLab[1], targetLab[2]];
    // L*は0-100の範囲にクランプ
    adjustedLab[0] = Math.max(0, Math.min(100, adjustedLab[0]));

    const matches = ColorMatcher.findClosestDMC(adjustedLab, dmcColors, 5);
    window.currentMatches = matches;

    matchList.innerHTML = matches.map((match, i) => `
      <div class="match-item${i === 0 ? ' selected' : ''}" onclick="selectMatch(${i})">
        <div class="match-rank">${i + 1}</div>
        <div class="match-swatch" style="background-color: ${match.hex}"></div>
        <div class="match-info">
          <div class="dmc-number">DMC ${match.dmc_id}</div>
          <div class="dmc-name">${match.name_en}</div>
        </div>
        <div class="match-score">
          <div class="delta-e">ΔE: ${match.deltaE.toFixed(2)}</div>
        </div>
      </div>
    `).join('');

    // Results are always visible in the hero area
}

function selectMatch(index) {
    if (!window.currentMatches || !window.currentMatches[index]) {
        console.error('Invalid match index:', index);
        return;
    }
    const match = window.currentMatches[index];
    document.querySelectorAll('.match-item').forEach((el, i) => {
        el.classList.toggle('selected', i === index);
    });
    addToHistory(match);
}

window.selectMatch = selectMatch;

function addToHistory(match) {
    const entry = {
        id: Date.now(),
        dmc_id: match.dmc_id,
        name_en: match.name_en,
        hex: match.hex,
        rgb: match.rgb,
        picked_hex: pickedColor.hex,
        picked_rgb: [pickedColor.r, pickedColor.g, pickedColor.b],
        deltaE: match.deltaE,
        x: pickedX,
        y: pickedY,
        timestamp: new Date().toISOString()
    };
    colorHistory.unshift(entry);
    localStorage.setItem('dmcHistory', JSON.stringify(colorHistory));
    render();
    renderLegend();
}

// Legend rendering
function renderLegend() {
    if (colorHistory.length === 0) {
        legendList.innerHTML = '<div class="legend-empty"><p>色を抽出すると、ここに凡例が表示されます</p></div>';
        return;
    }

    const dmcToNumber = new Map();
    let numberCounter = 1;
    [...colorHistory].reverse().forEach(entry => {
        if (!dmcToNumber.has(entry.dmc_id)) {
            dmcToNumber.set(entry.dmc_id, numberCounter++);
        }
    });

    // Get unique entries
    const uniqueEntries = [];
    const seenDmc = new Set();
    [...colorHistory].reverse().forEach(entry => {
        if (!seenDmc.has(entry.dmc_id)) {
            seenDmc.add(entry.dmc_id);
            uniqueEntries.push(entry);
        }
    });

    legendList.innerHTML = uniqueEntries.map((entry, index) => {
        const pinNumber = index + 1;
        const position = dmcPositions[entry.dmc_id];
        const positionText = position ? ` (${position.col}-${position.row})` : '';
        return `
            <div class="legend-entry" data-index="${index}" data-dmc-id="${entry.dmc_id}">
                <div class="legend-number">${pinNumber}</div>
                <div class="legend-color-swatch" style="background-color: ${entry.hex}"></div>
                <div class="legend-info">
                    <div class="legend-dmc">DMC ${entry.dmc_id}${positionText}</div>
                    <div class="legend-name">${entry.name_en}</div>
                </div>
                <div class="legend-delta">ΔE ${entry.deltaE.toFixed(1)}</div>
                <button class="legend-delete-btn" onclick="deleteLegendColor('${entry.dmc_id}'); event.stopPropagation();" title="この色を削除">×</button>
            </div>
        `;
    }).join('');
}

// Legend interaction
function onLegendMouseMove(e) {
    const entry = e.target.closest('.legend-entry');
    if (entry) {
        const index = parseInt(entry.dataset.index);
        if (hoveredLegendIndex !== index) {
            hoveredLegendIndex = index;

            // Update legend UI
            document.querySelectorAll('.legend-entry').forEach((el, i) => {
                el.classList.toggle('highlight', i === index);
            });

            // Update canvas
            render();
        }
    }
}

function onLegendMouseLeave() {
    hoveredLegendIndex = null;
    document.querySelectorAll('.legend-entry').forEach(el => {
        el.classList.remove('highlight');
    });
    render();
}

function onLegendClick(e) {
    const entry = e.target.closest('.legend-entry');
    if (entry) {
        const index = parseInt(entry.dataset.index);
        selectedLegendIndex = selectedLegendIndex === index ? null : index;

        // Center view on this pin
        const pinEntry = [...colorHistory].reverse().find(e => {
            const dmcToNumber = new Map();
            let numberCounter = 1;
            [...colorHistory].reverse().forEach(entry => {
                if (!dmcToNumber.has(entry.dmc_id)) {
                    dmcToNumber.set(entry.dmc_id, numberCounter++);
                }
            });
            return dmcToNumber.get(e.dmc_id) === index + 1;
        });

        if (pinEntry && pinEntry.x !== undefined && pinEntry.y !== undefined) {
            // Center on pin
            const targetX = viewportWidth / 2 - pinEntry.x * imageScale * zoom;
            const targetY = viewportHeight / 2 - pinEntry.y * imageScale * zoom;

            // Smooth animation
            const startX = offsetX;
            const startY = offsetY;
            const duration = 300;
            const startTime = Date.now();

            function animate() {
                const elapsed = Date.now() - startTime;
                const progress = Math.min(elapsed / duration, 1);
                const eased = 1 - Math.pow(1 - progress, 3); // Ease out cubic

                offsetX = startX + (targetX - startX) * eased;
                offsetY = startY + (targetY - startY) * eased;

                clampOffsets();
                render();

                if (progress < 1) {
                    requestAnimationFrame(animate);
                }
            }

            animate();
        }
    }
}

// Delete all entries of a specific DMC color
function deleteLegendColor(dmcId) {
    if (confirm(`DMC ${dmcId} の色をすべて削除しますか？`)) {
        colorHistory = colorHistory.filter(e => e.dmc_id !== dmcId);
        localStorage.setItem('dmcHistory', JSON.stringify(colorHistory));

        if (colorHistory.length === 0) {
            pickedColor = null;
            colorSwatch.style.backgroundColor = '#444';
            rgbValue.textContent = '-';
            hexValue.textContent = '-';
            matchList.innerHTML = '<div class="match-empty">画像をクリックして色を拾ってください</div>';
        }

        render();
        renderLegend();
    }
}

// Make deleteLegendColor available globally
window.deleteLegendColor = deleteLegendColor;

function clearHistory() {
    if (colorHistory.length === 0) {
        alert('クリアする履歴がありません');
        return;
    }
    if (confirm('履歴をすべて削除しますか？')) {
        colorHistory = [];
        localStorage.setItem('dmcHistory', JSON.stringify(colorHistory));

        pickedColor = null;
        // reset pick state
        colorSwatch.style.backgroundColor = '#444';
        rgbValue.textContent = '-';
        hexValue.textContent = '-';
        matchList.innerHTML = '<div class="match-empty">画像をクリックして色を拾ってください</div>';

        render();
        renderLegend();
    }
}

function exportToCSV() {
    try {
        console.log('Starting CSV Export...');
        const headers = ['DMC番号', '色名', 'DMC HEX', 'DMC RGB', '抽出色 HEX', '抽出色 RGB', 'ΔE', '日時'];
        const rows = colorHistory.map(e => [
            e.dmc_id, e.name_en, e.hex, e.rgb.join(' '),
            e.picked_hex, e.picked_rgb.join(' '), e.deltaE.toFixed(2),
            new Date(e.timestamp).toLocaleString('ja-JP')
        ]);
        const csvContent = [headers, ...rows].map(row => row.join(',')).join('\n');
        console.log('CSV Data Generated:', csvContent.length, 'bytes');

        const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        const filename = `dmc_colors_${dateStr}.csv`;

        const BOM = '\uFEFF';
        const blob = new Blob([BOM + csvContent], { type: 'text/csv;charset=utf-8;' });

        if (window.saveAs) {
            window.saveAs(blob, filename);
        } else {
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.setAttribute('download', filename);
            document.body.appendChild(link);
            link.click();
            setTimeout(() => document.body.removeChild(link), 100);
        }

        console.log('Download initiated');

    } catch (e) {
        console.error('CSV Export Failed:', e);
        alert('CSV出力に失敗しました:\n' + e.message);
    }
}

// Helper function to convert hex to RGB
function hexToRgb(hex) {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
    } : { r: 0, g: 0, b: 0 };
}

// PDF Export with Legend System
function exportToPDF() {
    try {
        console.log('Starting PDF Export with Legend System...');
        if (!window.jspdf) {
            throw new Error('PDFライブラリが読み込まれていません。');
        }
        if (!currentImage) {
            throw new Error('画像が読み込まれていません。');
        }
        if (colorHistory.length === 0) {
            throw new Error('抽出履歴がありません。');
        }

        const { jsPDF } = window.jspdf;

        // Build DMC to number mapping
        const dmcToNumber = new Map();
        let numberCounter = 1;
        [...colorHistory].reverse().forEach(entry => {
            if (!dmcToNumber.has(entry.dmc_id)) {
                dmcToNumber.set(entry.dmc_id, numberCounter++);
            }
        });

        // Get unique entries
        const uniqueDmcEntries = [];
        const seenDmc = new Set();
        [...colorHistory].reverse().forEach(item => {
            if (!seenDmc.has(item.dmc_id)) {
                seenDmc.add(item.dmc_id);
                uniqueDmcEntries.push(item);
            }
        });

        // Create landscape PDF
        const doc = new jsPDF('l');
        const pageWidth = 297;
        const pageHeight = 210;

        // ===== PAGE 1: Image with Pins + Legend =====

        // Page header
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(16);
        doc.setTextColor(44, 36, 22);
        doc.text('DMC Color Matcher Pro', 14, 14);

        doc.setFont('helvetica', 'normal');
        doc.setFontSize(9);
        doc.setTextColor(107, 83, 68);
        doc.text(`Generated: ${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}`, 14, 20);

        // Line separator
        doc.setDrawColor(107, 83, 68);
        doc.setLineWidth(0.3);
        doc.line(14, 22, pageWidth - 14, 22);

        // Layout: 60% image, 38% legend (more space for legend to prevent cutoff)
        const imageAreaWidth = pageWidth * 0.58;
        const legendAreaWidth = pageWidth * 0.38;
        const margin = 14;

        // Image area with margin for labels
        const labelMarginSpace = 25; // Space for labels around image
        const imgMaxWidth = imageAreaWidth - margin * 2 - labelMarginSpace * 2;
        const imgMaxHeight = pageHeight - 40 - labelMarginSpace * 2;
        const imgAspect = currentImage.width / currentImage.height;

        let imgWidth, imgHeight;
        if (imgAspect > imgMaxWidth / imgMaxHeight) {
            imgWidth = imgMaxWidth;
            imgHeight = imgWidth / imgAspect;
        } else {
            imgHeight = imgMaxHeight;
            imgWidth = imgHeight * imgAspect;
        }

        const imgX = margin + labelMarginSpace + (imageAreaWidth - margin * 2 - labelMarginSpace * 2 - imgWidth) / 2;
        const imgY = 30 + labelMarginSpace + (imgMaxHeight - imgHeight) / 2;

        // Convert image to data URL for PDF
        const tempCanvas = document.createElement('canvas');
        tempCanvas.width = currentImage.width;
        tempCanvas.height = currentImage.height;
        const tempCtx = tempCanvas.getContext('2d');
        tempCtx.drawImage(currentImage, 0, 0);
        const imageDataUrl = tempCanvas.toDataURL('image/png');

        // Draw image
        doc.addImage(imageDataUrl, 'PNG', imgX, imgY, imgWidth, imgHeight);

        // Draw extraction points with lines to labels outside image
        const placedLabels = [];
        const imageRect = { x: imgX, y: imgY, width: imgWidth, height: imgHeight };

        // Organize pins by edge for better distribution
        const pinsByEdge = { top: [], right: [], bottom: [], left: [] };

        [...colorHistory].reverse().forEach((entry) => {
            if (entry.x === undefined || entry.y === undefined) return;

            const pinX = imgX + (entry.x / currentImage.width) * imgWidth;
            const pinY = imgY + (entry.y / currentImage.height) * imgHeight;
            const pinNumber = dmcToNumber.get(entry.dmc_id);

            // Determine closest edge
            const distToTop = pinY - imageRect.y;
            const distToBottom = (imageRect.y + imageRect.height) - pinY;
            const distToLeft = pinX - imageRect.x;
            const distToRight = (imageRect.x + imageRect.width) - pinX;
            const minDist = Math.min(distToTop, distToBottom, distToLeft, distToRight);

            let edge;
            if (minDist === distToTop) edge = 'top';
            else if (minDist === distToBottom) edge = 'bottom';
            else if (minDist === distToLeft) edge = 'left';
            else edge = 'right';

            pinsByEdge[edge].push({ x: pinX, y: pinY, number: pinNumber, entry: entry });
        });

        // Sort pins on each edge
        pinsByEdge.top.sort((a, b) => a.x - b.x);
        pinsByEdge.bottom.sort((a, b) => a.x - b.x);
        pinsByEdge.left.sort((a, b) => a.y - b.y);
        pinsByEdge.right.sort((a, b) => a.y - b.y);

        // Draw pins and labels for each edge
        const labelSize = 12; // Compact label size
        const edgeMargin = 6; // Distance from image edge

        Object.entries(pinsByEdge).forEach(([edge, pins]) => {
            pins.forEach((pin, index) => {
                // Calculate label position
                let labelX, labelY;
                const spacing = edge === 'top' || edge === 'bottom'
                    ? imgWidth / (pins.length + 1)
                    : imgHeight / (pins.length + 1);

                if (edge === 'top') {
                    labelX = imgX + spacing * (index + 1);
                    labelY = imageRect.y - edgeMargin - labelSize / 2;
                } else if (edge === 'bottom') {
                    labelX = imgX + spacing * (index + 1);
                    labelY = imageRect.y + imageRect.height + edgeMargin + labelSize / 2;
                } else if (edge === 'left') {
                    labelX = imageRect.x - edgeMargin - labelSize / 2;
                    labelY = imgY + spacing * (index + 1);
                } else { // right
                    labelX = imageRect.x + imageRect.width + edgeMargin + labelSize / 2;
                    labelY = imgY + spacing * (index + 1);
                }

                // Draw connecting line
                doc.setDrawColor(107, 83, 68);
                doc.setLineWidth(0.3);
                doc.setLineDash([1, 1]);
                doc.line(pin.x, pin.y, labelX, labelY);
                doc.setLineDash([]);

                // Draw extraction point marker - smaller for precision
                const markerSize = 0.4; // Very small marker for exact location

                // Outer white glow
                doc.setFillColor(255, 255, 255);
                doc.circle(pin.x, pin.y, markerSize + 0.3, 'F');

                // Inner colored marker
                const dmcColor = hexToRgb(pin.entry.hex);
                doc.setFillColor(dmcColor.r, dmcColor.g, dmcColor.b);
                doc.circle(pin.x, pin.y, markerSize, 'F');

                // Border
                doc.setDrawColor(107, 83, 68);
                doc.setLineWidth(0.3);
                doc.circle(pin.x, pin.y, markerSize, 'S');

                // Draw label with color swatch integrated
                // Label is composed of color swatch (left) and number (right)
                const swatchWidth = labelSize * 0.4; // 40% for color swatch
                const numberWidth = labelSize * 0.6; // 60% for number

                // Color swatch background
                doc.setFillColor(dmcColor.r, dmcColor.g, dmcColor.b);
                doc.roundedRect(
                    labelX - labelSize / 2,
                    labelY - labelSize / 2,
                    swatchWidth,
                    labelSize,
                    2, 2, 'F'
                );

                // Number background (white)
                doc.setFillColor(255, 255, 255);
                doc.rect(
                    labelX - labelSize / 2 + swatchWidth,
                    labelY - labelSize / 2,
                    numberWidth,
                    labelSize,
                    'F'
                );

                // Right side rounded corner
                doc.setFillColor(255, 255, 255);
                doc.roundedRect(
                    labelX - labelSize / 2 + swatchWidth,
                    labelY - labelSize / 2,
                    numberWidth,
                    labelSize,
                    2, 2, 'F'
                );

                // Label border
                doc.setDrawColor(107, 83, 68);
                doc.setLineWidth(0.8);
                doc.roundedRect(labelX - labelSize / 2, labelY - labelSize / 2, labelSize, labelSize, 2, 2, 'S');

                // Number text (positioned on the white side)
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(8);
                doc.setTextColor(44, 36, 22);
                doc.text(
                    pin.number.toString(),
                    labelX - labelSize / 2 + swatchWidth + numberWidth / 2,
                    labelY,
                    { align: 'center', baseline: 'middle' }
                );
            });
        });

        // Legend area
        const legendX = imageAreaWidth + margin;
        const legendY = 32;

        // Legend header
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(14);
        doc.setTextColor(44, 36, 22);
        doc.text('Color Legend', legendX, legendY);

        // Legend subtitle
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(8);
        doc.setTextColor(107, 83, 68);
        doc.text(`${uniqueDmcEntries.length} unique colors`, legendX, legendY + 5);

        // Legend line separator
        doc.setDrawColor(107, 83, 68);
        doc.setLineWidth(0.2);
        doc.line(legendX, legendY + 7, pageWidth - margin, legendY + 7);

        // Detect crowded areas for detail view (will be on separate page)
        const allPins = [];
        Object.values(pinsByEdge).forEach(pins => {
            pins.forEach(pin => allPins.push(pin));
        });

        // Find crowded groups (pins within 10mm of each other)
        const crowdedGroups = [];
        const crowdThreshold = 10; // mm
        const processed = new Set();

        allPins.forEach((pin1, i) => {
            if (processed.has(i)) return;

            const group = [pin1];
            allPins.forEach((pin2, j) => {
                if (i !== j && !processed.has(j)) {
                    const dist = Math.sqrt(
                        Math.pow(pin1.x - pin2.x, 2) + Math.pow(pin1.y - pin2.y, 2)
                    );
                    if (dist < crowdThreshold) {
                        group.push(pin2);
                        processed.add(j);
                    }
                }
            });

            if (group.length > 1) {
                crowdedGroups.push(group);
                processed.add(i);
            }
        });

        // Store crowded groups data for detail page (after Color Details page)
        const crowdedGroupsData = crowdedGroups.length > 0 ? {
            groups: crowdedGroups,
            imageRect: { x: imgX, y: imgY, width: imgWidth, height: imgHeight }
        } : null;

        // Legend entries with improved spacing
        const entryStartY = legendY + 12;
        const entryHeight = 10;
        const maxEntriesPerPage = Math.floor((pageHeight - entryStartY - 10) / entryHeight);

        uniqueDmcEntries.slice(0, maxEntriesPerPage).forEach((item, index) => {
            const y = entryStartY + index * entryHeight;
            const pinNumber = index + 1;

            // Stripe background
            if (index % 2 === 0) {
                doc.setFillColor(252, 250, 248);
                doc.rect(legendX, y - 1.5, legendAreaWidth - 2, entryHeight, 'F');
            }

            // Number (compact)
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(9);
            doc.setTextColor(44, 36, 22);
            doc.text(pinNumber.toString(), legendX + 2, y + 3.5, { baseline: 'middle' });

            // Color swatch (slightly larger)
            const swatchSize = 7;
            const swatchX = legendX + 11;
            const dmcColor = hexToRgb(item.hex);

            doc.setFillColor(dmcColor.r, dmcColor.g, dmcColor.b);
            doc.roundedRect(swatchX, y + 0.5, swatchSize, swatchSize, 1, 1, 'F');

            // Swatch border
            doc.setDrawColor(180, 180, 180);
            doc.setLineWidth(0.25);
            doc.roundedRect(swatchX, y + 0.5, swatchSize, swatchSize, 1, 1, 'S');

            // DMC number (compact)
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(8);
            doc.setTextColor(44, 36, 22);
            const dmcText = `${item.dmc_id}`;
            doc.text(dmcText, swatchX + swatchSize + 3, y + 3.5, { baseline: 'middle' });

            // Delta E (right aligned, more compact)
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(7);
            doc.setTextColor(100, 100, 100);
            const deltaText = `Δ${item.deltaE.toFixed(1)}`;
            doc.text(deltaText, legendX + legendAreaWidth - 4, y + 3.5, { align: 'right', baseline: 'middle' });
        });

        // ===== PAGE 2: Detailed Color Table =====
        if (uniqueDmcEntries.length > 0) {
            doc.addPage('p');

            // Page header
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(16);
            doc.setTextColor(44, 36, 22);
            doc.text('Color Details', 14, 14);

            doc.setFont('helvetica', 'normal');
            doc.setFontSize(9);
            doc.setTextColor(107, 83, 68);
            doc.text(`${uniqueDmcEntries.length} unique colors extracted`, 14, 20);

            // Line separator
            doc.setDrawColor(107, 83, 68);
            doc.setLineWidth(0.3);
            doc.line(14, 22, 210 - 14, 22);

            // Table data
            const tableColumn = ["No.", "Color", "DMC", "Name", "DMC Hex", "Picked Hex", "ΔE"];

            const tableRows = uniqueDmcEntries.map((item, index) => {
                const pinNumber = index + 1;
                return [
                    pinNumber.toString(),
                    '',
                    item.dmc_id,
                    item.name_en,
                    item.hex,
                    item.picked_hex,
                    item.deltaE.toFixed(2)
                ];
            });

            const tableMargin = 14;
            const tableWidth = 210 - (tableMargin * 2);

            doc.autoTable({
                head: [tableColumn],
                body: tableRows,
                startY: 26,
                margin: { left: tableMargin, right: tableMargin },
                tableWidth: tableWidth,
                theme: 'striped',
                styles: {
                    fontSize: 9,
                    cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
                    lineColor: [220, 220, 220],
                    lineWidth: 0.1,
                    textColor: [44, 36, 22],
                    font: 'helvetica'
                },
                headStyles: {
                    fillColor: [107, 83, 68],
                    textColor: [255, 255, 255],
                    fontSize: 9,
                    fontStyle: 'bold',
                    halign: 'center',
                    cellPadding: { top: 5, right: 3, bottom: 5, left: 3 }
                },
                alternateRowStyles: {
                    fillColor: [250, 248, 245]
                },
                columnStyles: {
                    0: { cellWidth: 10, halign: 'center', fontStyle: 'bold' },
                    1: { cellWidth: 20, halign: 'center' },
                    2: { cellWidth: 16, halign: 'center', fontStyle: 'bold' },
                    3: { cellWidth: 'auto', halign: 'left' },
                    4: { cellWidth: 22, halign: 'center', fontSize: 8 },
                    5: { cellWidth: 22, halign: 'center', fontSize: 8 },
                    6: { cellWidth: 14, halign: 'center', fontSize: 8 }
                },
                didDrawCell: function (data) {
                    if (data.column.index === 1 && data.section === 'body') {
                        const rowIndex = data.row.index;
                        const item = uniqueDmcEntries[rowIndex];

                        const swatchWidth = 8;
                        const swatchHeight = data.cell.height - 4;
                        const startX = data.cell.x + (data.cell.width - swatchWidth * 2 - 2) / 2;
                        const startY = data.cell.y + 2;

                        // Picked color (left)
                        const pickedColor = hexToRgb(item.picked_hex);
                        doc.setFillColor(pickedColor.r, pickedColor.g, pickedColor.b);
                        doc.roundedRect(startX, startY, swatchWidth, swatchHeight, 1, 1, 'F');

                        doc.setDrawColor(180, 180, 180);
                        doc.setLineWidth(0.25);
                        doc.roundedRect(startX, startY, swatchWidth, swatchHeight, 1, 1, 'S');

                        // DMC color (right)
                        const dmcColor = hexToRgb(item.hex);
                        doc.setFillColor(dmcColor.r, dmcColor.g, dmcColor.b);
                        doc.roundedRect(startX + swatchWidth + 2, startY, swatchWidth, swatchHeight, 1, 1, 'F');

                        doc.setDrawColor(180, 180, 180);
                        doc.setLineWidth(0.25);
                        doc.roundedRect(startX + swatchWidth + 2, startY, swatchWidth, swatchHeight, 1, 1, 'S');
                    }
                }
            });

            // Footer
            const finalY = doc.lastAutoTable.finalY || 50;

            doc.setDrawColor(107, 83, 68);
            doc.setLineWidth(0.2);
            doc.line(14, finalY + 6, 210 - 14, finalY + 6);

            doc.setFont('helvetica', 'bold');
            doc.setFontSize(9);
            doc.setTextColor(107, 83, 68);
            doc.text('Color Accuracy (ΔE)', 14, finalY + 12);

            doc.setFont('helvetica', 'normal');
            doc.setFontSize(8);
            doc.setTextColor(120, 120, 120);
            doc.text('0-2: Almost identical  •  2-5: Close match  •  5-10: Noticeable  •  10+: Different', 14, finalY + 17);
        }

        // ===== PAGE 3+: Detail Views for Crowded Areas =====
        if (crowdedGroupsData && crowdedGroupsData.groups.length > 0) {
            // Sort by size and process each group
            const sortedGroups = [...crowdedGroupsData.groups].sort((a, b) => b.length - a.length);

            sortedGroups.forEach((group, groupIndex) => {
                // Add new landscape page for each crowded area
                doc.addPage('l');

                const detailPageWidth = 297;
                const detailPageHeight = 210;
                const detailMargin = 14;

                // Page header
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(16);
                doc.setTextColor(44, 36, 22);
                doc.text(`Detail View ${groupIndex + 1} - Magnified Area`, detailMargin, 14);

                doc.setFont('helvetica', 'normal');
                doc.setFontSize(9);
                doc.setTextColor(107, 83, 68);
                doc.text(`${group.length} extraction points in close proximity (magnified 8x)`, detailMargin, 20);

                // Line separator
                doc.setDrawColor(107, 83, 68);
                doc.setLineWidth(0.3);
                doc.line(detailMargin, 22, detailPageWidth - detailMargin, 22);

                // Calculate bounding box of the group
                const xs = group.map(p => p.x);
                const ys = group.map(p => p.y);
                const minX = Math.min(...xs) - 8;
                const maxX = Math.max(...xs) + 8;
                const minY = Math.min(...ys) - 8;
                const maxY = Math.max(...ys) + 8;
                const groupWidth = maxX - minX;
                const groupHeight = maxY - minY;

                // Detail view dimensions (use most of the page)
                const detailMaxWidth = detailPageWidth - detailMargin * 2;
                const detailMaxHeight = detailPageHeight - 40;
                const detailAspect = groupWidth / groupHeight;

                let detailWidth, detailHeight;
                if (detailAspect > detailMaxWidth / detailMaxHeight) {
                    detailWidth = detailMaxWidth;
                    detailHeight = detailWidth / detailAspect;
                } else {
                    detailHeight = detailMaxHeight;
                    detailWidth = detailHeight * detailAspect;
                }

                const detailX = detailMargin + (detailMaxWidth - detailWidth) / 2;
                const detailY = 30 + (detailMaxHeight - detailHeight) / 2;

                // Draw cropped image section at high magnification
                const imgRect = crowdedGroupsData.imageRect;
                const srcX = ((minX - imgRect.x) / imgRect.width) * currentImage.width;
                const srcY = ((minY - imgRect.y) / imgRect.height) * currentImage.height;
                const srcW = (groupWidth / imgRect.width) * currentImage.width;
                const srcH = (groupHeight / imgRect.height) * currentImage.height;

                // Create high-res cropped canvas
                const detailCanvas = document.createElement('canvas');
                detailCanvas.width = Math.min(srcW, currentImage.width);
                detailCanvas.height = Math.min(srcH, currentImage.height);
                const detailCtx = detailCanvas.getContext('2d');

                detailCtx.drawImage(
                    currentImage,
                    Math.max(0, srcX),
                    Math.max(0, srcY),
                    Math.min(srcW, currentImage.width - Math.max(0, srcX)),
                    Math.min(srcH, currentImage.height - Math.max(0, srcY)),
                    0,
                    0,
                    detailCanvas.width,
                    detailCanvas.height
                );

                const detailImageUrl = detailCanvas.toDataURL('image/png');

                // Background
                doc.setFillColor(245, 245, 245);
                doc.roundedRect(detailX - 2, detailY - 2, detailWidth + 4, detailHeight + 4, 4, 4, 'F');

                // Draw image
                doc.addImage(detailImageUrl, 'PNG', detailX, detailY, detailWidth, detailHeight);

                // Border
                doc.setDrawColor(107, 83, 68);
                doc.setLineWidth(1);
                doc.roundedRect(detailX, detailY, detailWidth, detailHeight, 3, 3, 'S');

                // Calculate scale for positioning pins
                const scale = detailWidth / groupWidth;

                // Calculate pin positions and prepare for smart label placement
                const labelSize = 18;
                const placedLabels = []; // Track placed label positions
                const allPinPositions = group.map(p => ({
                    x: detailX + ((p.x - minX) * scale),
                    y: detailY + ((p.y - minY) * scale)
                }));

                // Helper function to check if two rectangles overlap
                function rectOverlap(r1, r2, padding = 5) {
                    return !(r1.x + r1.width + padding < r2.x ||
                             r2.x + r2.width + padding < r1.x ||
                             r1.y + r1.height + padding < r2.y ||
                             r2.y + r2.height + padding < r1.y);
                }

                // Helper function to check if label overlaps any pin
                function overlapsAnyPin(labelRect, currentPinX, currentPinY) {
                    const pinRadius = 3; // Safety radius around pins
                    for (const pin of allPinPositions) {
                        // Skip the current pin
                        if (Math.abs(pin.x - currentPinX) < 0.1 && Math.abs(pin.y - currentPinY) < 0.1) {
                            continue;
                        }
                        // Check if label overlaps this pin
                        if (labelRect.x - pinRadius < pin.x && pin.x < labelRect.x + labelRect.width + pinRadius &&
                            labelRect.y - pinRadius < pin.y && pin.y < labelRect.y + labelRect.height + pinRadius) {
                            return true;
                        }
                    }
                    return false;
                }

                // Helper function to find best label position with multiple distance attempts
                function findBestLabelPosition(pinX, pinY, labelSize) {
                    // 8 directions
                    const directions = [
                        { angle: -90, name: 'top' },           // Up
                        { angle: -45, name: 'top-right' },     // Up-right
                        { angle: 0, name: 'right' },           // Right
                        { angle: 45, name: 'bottom-right' },   // Down-right
                        { angle: 90, name: 'bottom' },         // Down
                        { angle: 135, name: 'bottom-left' },   // Down-left
                        { angle: 180, name: 'left' },          // Left
                        { angle: -135, name: 'top-left' }      // Up-left
                    ];

                    // Try increasing distances: 12mm, 18mm, 24mm, 30mm
                    const distances = [12, 18, 24, 30];

                    for (const distance of distances) {
                        for (const dir of directions) {
                            const angleRad = (dir.angle * Math.PI) / 180;
                            const offsetX = distance * Math.cos(angleRad);
                            const offsetY = distance * Math.sin(angleRad);

                            const labelCenterX = pinX + offsetX;
                            const labelCenterY = pinY + offsetY;

                            const labelRect = {
                                x: labelCenterX - labelSize / 2,
                                y: labelCenterY - labelSize / 2,
                                width: labelSize,
                                height: labelSize
                            };

                            // Check if within bounds
                            if (labelRect.x < detailX || labelRect.x + labelSize > detailX + detailWidth ||
                                labelRect.y < detailY || labelRect.y + labelSize > detailY + detailHeight) {
                                continue;
                            }

                            // Check overlap with existing labels
                            let hasLabelOverlap = false;
                            for (const placed of placedLabels) {
                                if (rectOverlap(labelRect, placed, 5)) {
                                    hasLabelOverlap = true;
                                    break;
                                }
                            }

                            if (hasLabelOverlap) continue;

                            // Check overlap with any pins
                            if (overlapsAnyPin(labelRect, pinX, pinY)) {
                                continue;
                            }

                            // Found a good position!
                            return {
                                offset: { x: offsetX, y: offsetY },
                                rect: labelRect,
                                centerX: labelCenterX,
                                centerY: labelCenterY
                            };
                        }
                    }

                    // Fallback: place it far above if no position found
                    const fallbackOffsetY = -30;
                    return {
                        offset: { x: 0, y: fallbackOffsetY },
                        rect: {
                            x: pinX - labelSize / 2,
                            y: pinY + fallbackOffsetY - labelSize / 2,
                            width: labelSize,
                            height: labelSize
                        },
                        centerX: pinX,
                        centerY: pinY + fallbackOffsetY
                    };
                }

                // Draw pins with smart label placement
                group.forEach((pin) => {
                    const localX = detailX + ((pin.x - minX) * scale);
                    const localY = detailY + ((pin.y - minY) * scale);

                    const dmcColor = hexToRgb(pin.entry.hex);

                    // Large marker for detail view
                    const markerSize = 1.5;

                    // Outer white glow
                    doc.setFillColor(255, 255, 255);
                    doc.circle(localX, localY, markerSize + 0.8, 'F');

                    // Inner colored marker
                    doc.setFillColor(dmcColor.r, dmcColor.g, dmcColor.b);
                    doc.circle(localX, localY, markerSize, 'F');

                    // Border
                    doc.setDrawColor(44, 36, 22);
                    doc.setLineWidth(0.6);
                    doc.circle(localX, localY, markerSize, 'S');

                    // Find best label position
                    const labelPlacement = findBestLabelPosition(localX, localY, labelSize);
                    placedLabels.push(labelPlacement.rect);

                    const swatchWidth = labelSize * 0.4;
                    const numberWidth = labelSize * 0.6;
                    const labelCenterX = labelPlacement.centerX;
                    const labelCenterY = labelPlacement.centerY;

                    // Color swatch background
                    doc.setFillColor(dmcColor.r, dmcColor.g, dmcColor.b);
                    doc.roundedRect(
                        labelCenterX - labelSize / 2,
                        labelCenterY - labelSize / 2,
                        swatchWidth,
                        labelSize,
                        2, 2, 'F'
                    );

                    // Number background (white)
                    doc.setFillColor(255, 255, 255);
                    doc.roundedRect(
                        labelCenterX - labelSize / 2 + swatchWidth,
                        labelCenterY - labelSize / 2,
                        numberWidth,
                        labelSize,
                        2, 2, 'F'
                    );

                    // Label border (bold)
                    doc.setDrawColor(44, 36, 22);
                    doc.setLineWidth(1);
                    doc.roundedRect(
                        labelCenterX - labelSize / 2,
                        labelCenterY - labelSize / 2,
                        labelSize,
                        labelSize,
                        2, 2, 'S'
                    );

                    // Number text (large and bold)
                    doc.setFont('helvetica', 'bold');
                    doc.setFontSize(11);
                    doc.setTextColor(44, 36, 22);
                    doc.text(
                        pin.number.toString(),
                        labelCenterX - labelSize / 2 + swatchWidth + numberWidth / 2,
                        labelCenterY,
                        { align: 'center', baseline: 'middle' }
                    );

                    // Connecting line from marker to label
                    doc.setDrawColor(107, 83, 68);
                    doc.setLineWidth(0.5);
                    doc.setLineDash([1, 1]);
                    doc.line(
                        localX,
                        localY,
                        labelCenterX,
                        labelCenterY
                    );
                    doc.setLineDash([]);
                });

                // Add reference info at bottom
                doc.setFont('helvetica', 'italic');
                doc.setFontSize(8);
                doc.setTextColor(120, 120, 120);
                doc.text(
                    'This magnified view shows extraction points that are close together for easier identification.',
                    detailPageWidth / 2,
                    detailPageHeight - 6,
                    { align: 'center' }
                );
            });
        }

        // Save PDF
        const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        const filename = `dmc_pro_${dateStr}.pdf`;

        console.log('PDF Generated. Saving as:', filename);
        doc.save(filename);

        console.log('PDF download initiated.');

    } catch (e) {
        console.error('PDF Export Failed:', e);
        alert('PDF出力に失敗しました:\n' + e.message);
    }
}

init();
