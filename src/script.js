
// --- DOM 引用 ---
var areaDataManager = new ExcelDataManager();
var calcBtn = document.getElementById('calcBtn');
var resetBtn = document.getElementById('resetBtn');
var areaDataManager = new ExcelDataManager();

var rad2deg = function (r) { return r * 180 / Math.PI; };
var deg2rad = function (d) { return d * Math.PI / 180; };


// ============================================================
//  配置驱动辅助函数（第一步重构）
//  使用 window.MODULES_CONFIG 替代硬编码
// ============================================================

/**
 * 获取当前模块的配置
 */
function getCurrentModuleConfig() {
    if (!window.MODULES_CONFIG) {
        console.warn('⚠️ MODULES_CONFIG 未加载，请检查 modules-config.js');
        return null;
    }
    return window.MODULES_CONFIG.getModuleConfig(currentSub);
}

/**
 * 根据当前模块配置，收集所有输入框的值
 * 自动处理特殊输入（如 arrangement 下拉框）
 */
function collectInputsForCurrentModule() {
    const config = getCurrentModuleConfig();
    if (!config) return null;

    const values = {};

    config.inputs.forEach(input => {
        // 直接通过完整 ID 获取元素
        let el = document.getElementById(input.id);

        // 如果是下拉框（select），特殊处理一下值解析
        if (el) {
            let val = el.value;

            if (input.type === 'number') {
                const parsed = parseFloat(val);
                values[input.id] = isNaN(parsed) ? (input.default || 0) : parsed;
            } else if (input.type === 'select') {
                // 排列分布：尝试解析 JSON（如果是空字符串则返回 null）
                if (val && val.startsWith('{')) {
                    try { values[input.id] = JSON.parse(val); }
                    catch (e) { values[input.id] = val; }
                } else {
                    values[input.id] = val;
                }
            } else {
                values[input.id] = val;
            }
        } else {
            // 如果 DOM 元素不存在，使用默认值并警告
            console.warn(`⚠️ 未找到输入元素: ${input.id}，使用默认值`);
            values[input.id] = input.default || '';
        }
    });

    return values;
}

/**
 * 根据当前模块配置，动态渲染输入框
 * 替代所有硬编码的 #size-inputs / #custom-inputs 等
 */
function renderInputsForModule(moduleId) {
    const config = getCurrentModuleConfig();
    if (!config) {
        console.warn('⚠️ 未找到模块配置:', moduleId);
        return;
    }

    const container = document.getElementById('dynamic-input-area');
    if (!container) {
        console.warn('⚠️ 未找到 #dynamic-input-area 容器');
        return;
    }

    let html = '';
    const inputs = config.inputs || [];

    inputs.forEach(input => {
        const id = input.id;
        const label = input.label;
        const type = input.type || 'text';
        const defaultValue = input.default !== undefined ? input.default : '';

        // 特殊处理：下拉选择框（排列分布）
        if (type === 'select') {
            // 先渲染一个空 select，后面由 updateArrangementOptions 填充选项
            html += `
                <div class="form-item">
                    <label for="${id}">${label}</label>
                    <select id="${id}"></select>
                </div>
            `;
        } else {
            // 普通输入框（text / number）
            const stepAttr = (type === 'number') ? ' step="any"' : '';
            html += `
                <div class="form-item">
                    <label for="${id}">${label}</label>
                    <input type="${type}" id="${id}" value="${defaultValue}" placeholder="请输入..." ${stepAttr} />
                </div>
            `;
        }
    });

    container.innerHTML = html;

    // ---- 后处理：绑定特殊事件 ----
    // 1. 如果当前模块有相机个数输入框，绑定排列更新事件
    const countInput = document.getElementById('customCameraCount');
    if (countInput) {
        // 移除旧监听器，避免重复绑定（用新函数替换）
        const newHandler = function () {
            updateArrangementOptions();
        };
        // 保存引用以便后续清理（可选）
        countInput.removeEventListener('input', countInput._arrangeHandler);
        countInput.removeEventListener('change', countInput._arrangeHandler);
        countInput._arrangeHandler = newHandler;
        countInput.addEventListener('input', newHandler);
        countInput.addEventListener('change', newHandler);
    }

    // 2. 如果当前模块有排列分布下拉框，立即填充选项
    if (inputs.some(inp => inp.id === 'arrangementSelect')) {
        updateArrangementOptions();
    }

    // 3. 触发自定义事件，通知其他模块（如 3D 场景）输入已更新（可选）
    const event = new CustomEvent('inputsRendered', { detail: { moduleId: moduleId } });
    document.dispatchEvent(event);

    // ---- 新增：工作距离同步 ----
    if (typeof window.syncDistCamToArea === 'function') {
        setTimeout(function () {
            // 1. 先从 line-distCam 同步（用于手动输入）
            window.syncDistCamToArea();

            // 2. 如果有缓存的最新面阵工作距离，则覆盖（确保点击卡片的值生效）
            if (window._latestAreaWd !== undefined && window._latestAreaWd !== null && window._latestAreaWd !== '') {
                if (typeof window.updateAreaDistCam === 'function') {
                    window.updateAreaDistCam(window._latestAreaWd);
                }
            }
        }, 20);
    }
}


/**
 * 根据当前模块配置和计算结果，更新 UI 结果区
 * 替代大量的 document.getElementById('xxx').innerText
 */
function renderResultsForCurrentModule(resultData) {
    const config = getCurrentModuleConfig();
    if (!config) return;

    const map = config.resultMap;
    if (!map) return;

    // 遍历结果映射，更新对应的 DOM 元素
    for (const [key, elementId] of Object.entries(map)) {
        const el = document.getElementById(elementId);
        if (!el) continue;

        let value = resultData[key];
        // 如果值是数字，保留一位小数并添加单位（由调用方决定）
        if (typeof value === 'number' && !isNaN(value)) {
            // 这里不自动加单位，因为有些是角度（°）有些是长度（mm）
            // 由 calculate 函数在赋值时直接传格式化好的字符串
            el.innerText = value;
        } else if (value !== undefined && value !== null) {
            el.innerText = value;
        } else {
            el.innerText = '--';
        }
    }
}


// --- 子状态 ---
var currentGroup = 'size';
var currentSub = 'size-face';

// --- 辅助函数：清空指定2D画布 ---
function clearCanvas(canvasId) {
    var canvas = document.getElementById(canvasId);
    if (canvas) {
        var ctx = canvas.getContext('2d');
        if (ctx) {
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            canvas.width = 800;
            canvas.height = 200;
            canvas.style.width = '';
            canvas.style.height = '';
        }
    }
}

// --- 清除指定子功能的结果数据 ---
function clearSubResults(subId) {
    const config = getCurrentModuleConfig();
    if (!config) return;

    // 通用清空：所有结果字段设为 '--'
    for (const [key, elementId] of Object.entries(config.resultMap)) {
        const el = document.getElementById(elementId);
        if (el) el.innerText = '--';
    }

    // 特殊处理：定制面光需要清空画布
    if (subId === 'size-custom') {
        clearCanvas('scCanTotal');
    }

    document.getElementById('commonFovH').innerText = '--';
    document.getElementById('commonFovV').innerText = '--';
    document.getElementById('commonFovD').innerText = '--';
    document.getElementById('commonSizeText').innerText = '--';

    clearThreeScene();
}

// ============================================================
//  搜索选择器的初始化和控制函数
// ============================================================
function initSearchSelect(wrapper, type) {
    type = type || 'camera';
    var input = wrapper.querySelector('.search-select-input');
    var dropdown = wrapper.querySelector('.search-select-dropdown');
    var hiddenSelect = wrapper.querySelector('select');
    var selectId = wrapper.dataset.selectId;

    var items = [];

    function renderDropdown(filterText) {
        dropdown.innerHTML = '';
        var matched = [];
        if (filterText) {
            var lower = filterText.toLowerCase();
            for (var i = 0; i < items.length; i++) {
                if (items[i].name.toLowerCase().indexOf(lower) !== -1) {
                    matched.push(items[i]);
                }
            }
        } else {
            matched = items.slice(0);
        }

        if (matched.length === 0) {
            var li = document.createElement('li');
            li.className = 'no-match';
            li.textContent = '无匹配型号';
            dropdown.appendChild(li);
        } else {
            for (var j = 0; j < matched.length; j++) {
                var li = document.createElement('li');
                li.textContent = matched[j].name;
                li.dataset.row = JSON.stringify(matched[j].row);
                li.addEventListener('mousedown', function (e) {
                    e.preventDefault();
                    var name = this.textContent;
                    var row = JSON.parse(this.dataset.row);
                    input.value = name;
                    dropdown.classList.remove('show');
                    // 不再自动填充，由外部统一处理
                    hiddenSelect.value = JSON.stringify(row);
                    input.dataset.row = JSON.stringify(row);
                    var evt = new Event('change', { bubbles: true });
                    hiddenSelect.dispatchEvent(evt);
                });
                dropdown.appendChild(li);
            }
        }
    }

    function getPrefixForSelect(id) {
        if (id === 'modelSelect') return 'size-';
        if (id === 'custom-modelSelect') return 'custom-';
        if (id === 'spot-modelSelect') return 'spot-';
        if (id === 'coax-modelSelect') return 'coax-';
        return '';
    }

    input.addEventListener('input', function () {
        var val = this.value.trim();
        renderDropdown(val);
        if (val.length > 0) {
            dropdown.classList.add('show');
        } else {
            dropdown.classList.remove('show');
        }
    });

    document.addEventListener('click', function (e) {
        if (!wrapper.contains(e.target)) {
            dropdown.classList.remove('show');
        }
    });

    input.addEventListener('keydown', function (e) {
        var lis = dropdown.querySelectorAll('li:not(.no-match)');
        if (lis.length === 0) return;
        var activeIdx = -1;
        for (var i = 0; i < lis.length; i++) {
            if (lis[i].classList.contains('active')) {
                activeIdx = i;
                break;
            }
        }
        if (e.key === 'ArrowDown') {
            e.preventDefault();
            var nextIdx = (activeIdx + 1) % lis.length;
            lis.forEach(function (li) { li.classList.remove('active'); });
            lis[nextIdx].classList.add('active');
            lis[nextIdx].scrollIntoView({ block: 'nearest' });
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            var prevIdx = (activeIdx - 1 + lis.length) % lis.length;
            lis.forEach(function (li) { li.classList.remove('active'); });
            lis[prevIdx].classList.add('active');
            lis[prevIdx].scrollIntoView({ block: 'nearest' });
        } else if (e.key === 'Enter') {
            e.preventDefault();
            var activeLi = dropdown.querySelector('li.active');
            if (activeLi) {
                activeLi.click();
            } else if (lis.length === 1) {
                lis[0].click();
            }
        } else if (e.key === 'Escape') {
            dropdown.classList.remove('show');
        }
    });

    wrapper.updateItems = function (newItems) {
        items = newItems.slice();
        var currentVal = input.value.trim();
        if (currentVal) {
            renderDropdown(currentVal);
        } else {
            dropdown.classList.remove('show');
        }
    };

    var initialItems = (type === 'lens') ? areaDataManager.getLensItems() : areaDataManager.getCameraItems();
    wrapper.updateItems(initialItems);

    dropdown.classList.remove('show');

    wrapper._input = input;
    wrapper._dropdown = dropdown;
    wrapper._hiddenSelect = hiddenSelect;
}

function initAllSearchSelects() {
    // 只初始化公共的相机和镜头选择器
    var cameraWrapper = document.querySelector('[data-select-id="common-modelSelect"]');
    if (cameraWrapper) initSearchSelect(cameraWrapper, 'camera');
    var lensWrapper = document.querySelector('[data-select-id="common-lensSelect"]');
    if (lensWrapper) initSearchSelect(lensWrapper, 'lens');
}

function updateAllSearchSelects() {
    if (!areaDataManager.isLoaded) return;
    var items = areaDataManager.getCameraItems();
    var wrapper = document.querySelector('[data-select-id="common-modelSelect"]');
    if (wrapper && wrapper.updateItems) wrapper.updateItems(items);
}
function updateLensSearchSelects() {
    if (!areaDataManager.isLoaded) return;
    var items = areaDataManager.getLensItems();
    var wrapper = document.querySelector('[data-select-id="common-lensSelect"]');
    if (wrapper && wrapper.updateItems) wrapper.updateItems(items);
}

// ============================================================
//  两级菜单切换
// ============================================================
document.querySelectorAll('.sub-tab-btn.level1').forEach(function (btn) {
    btn.addEventListener('click', function () {
        var group = this.dataset.group;
        document.querySelectorAll('.sub-tab-btn.level1').forEach(function (b) { b.classList.remove('active'); });
        this.classList.add('active');
        document.getElementById('level2-size').style.display = (group === 'size') ? 'flex' : 'none';
        document.getElementById('level2-spot').style.display = (group === 'spot') ? 'flex' : 'none';
        var firstSub = document.querySelector('#level2-' + group + ' .sub-tab-btn.level2');
        if (firstSub) {
            clearSubResults(currentSub);
            firstSub.click();
        }
        currentGroup = group;
    });
});

document.querySelectorAll('.sub-tab-btn.level2').forEach(function (btn) {
    btn.addEventListener('click', function () {
        // 1. 切换激活状态
        const parent = this.closest('.sub-tab-level2');
        parent.querySelectorAll('.sub-tab-btn.level2').forEach(function (b) { b.classList.remove('active'); });
        this.classList.add('active');

        // 2. 更新当前子模块 ID
        currentSub = this.dataset.sub;

        // 3. ⭐ 动态渲染输入框（取代所有 display:none 切换）
        renderInputsForModule(currentSub);

        // 4. 切换结果区显示
        document.querySelectorAll('.sub-result').forEach(function (el) { el.style.display = 'none'; });
        const target = document.querySelector('.sub-result-' + currentSub);
        if (target) target.style.display = 'block';

        // 5. 清空旧结果，切换 3D 容器
        clearSubResults(currentSub);
        switchThreeContainer(currentSub);

        // 6. （可选）触发公共计算器状态更新
        console.log('🔁 切换到模块:', currentSub);
    });
});


// --- Excel文件数据更新 ---
function updateAreaData() {
    var fileInput = document.getElementById('fileUpload');
    var file = fileInput.files[0];
    if (!file) {
        alert('请先选择Excel文件');
        return;
    }
    var statusSpan = document.getElementById('dataStatus');
    if (statusSpan) {
        statusSpan.textContent = '加载中...';
        statusSpan.style.color = '#ff9800';
    }

    // 显示加载遮罩（如果存在）
    var tip = document.getElementById('loadingTip');
    if (tip) tip.classList.add('show');

    areaDataManager.loadExcel(file)
        .then(function () {
            if (statusSpan) {
                statusSpan.textContent = areaDataManager.getVersionInfo();
                statusSpan.style.color = '#4caf50';
            }
            updateAllSearchSelects();
            updateLensSearchSelects();
            areaDataManager.saveToLocal();
            alert('数据更新成功！\n文件: ' + areaDataManager.fileName + '\n相机: ' + areaDataManager.cameraData.length + '条, 镜头: ' + areaDataManager.lensData.length + '条');

            // ----- 修复：恢复界面可交互 -----
            // 1. 隐藏加载遮罩
            if (tip) tip.classList.remove('show');
            // 2. 启用所有输入框（如果有被禁用的）
            document.querySelectorAll('input, textarea, select, button').forEach(el => el.disabled = false);
            // 3. 移除可能残留的覆盖层（如果有模态背景等）
            document.querySelectorAll('.modal-backdrop, .overlay, .loading-mask').forEach(el => el.remove());
            // 4. 确保动态输入区域可用（不需额外操作）
        })
        .catch(function (error) {
            if (statusSpan) {
                statusSpan.textContent = '加载失败';
                statusSpan.style.color = '#f44336';
            }
            alert('数据加载失败: ' + error.message);

            // 同样恢复界面
            if (tip) tip.classList.remove('show');
            document.querySelectorAll('input, textarea, select, button').forEach(el => el.disabled = false);
            document.querySelectorAll('.modal-backdrop, .overlay, .loading-mask').forEach(el => el.remove());
        });
}
// 点击更新按钮 → 触发隐藏的文件选择框
document.getElementById('updateDataBtn').addEventListener('click', function () {
    document.getElementById('fileUpload').click();
});

// 文件选择后自动加载
document.getElementById('fileUpload').addEventListener('change', updateAreaData);


// --- 主工具切换 ---
function switchTool(toolVal) {
    var areaWrap = document.getElementById('areaWrap');
    var lineWrap = document.getElementById('lineSelectorWrap');

    if (toolVal === 'area') {
        // 显示面阵，隐藏镜头
        if (areaWrap) {
            areaWrap.style.display = 'flex';
            areaWrap.style.flexDirection = 'row';
            areaWrap.style.flexWrap = 'nowrap';
        }
        if (lineWrap) lineWrap.style.display = 'none';

        // 显示输入卡片和结果卡片
        document.getElementById('inputWrap').style.display = 'block';
        document.getElementById('resultCard').style.display = 'block';

        // 根据当前模块动态渲染输入框
        renderInputsForModule(currentSub);

        // 刷新数据状态
        if (areaDataManager.isLoaded) {
            var statusSpan = document.getElementById('dataStatus');
            if (statusSpan) {
                statusSpan.textContent = areaDataManager.getVersionInfo();
                statusSpan.style.color = '#4caf50';
            }
            updateAllSearchSelects();
            updateLensSearchSelects();
        }

        // 显示对应的结果区域
        document.querySelectorAll('.sub-result').forEach(function (el) { el.style.display = 'none'; });
        var target = document.querySelector('.sub-result-' + currentSub);
        if (target) target.style.display = 'block';
        clearSubResults(currentSub);
        switchThreeContainer(currentSub);

    } else {
        // 显示镜头，隐藏面阵
        if (areaWrap) areaWrap.style.display = 'none';
        if (lineWrap) lineWrap.style.display = 'block';

        // 隐藏输入和结果卡片
        document.getElementById('inputWrap').style.display = 'none';
        document.getElementById('resultCard').style.display = 'none';
    }
}

// --- 排列更新 ---
function getArrangements(n) {
    var arr = [];
    for (var i = 1; i <= Math.sqrt(n); i++) {
        if (n % i === 0) {
            var j = n / i;
            arr.push({ rows: i, cols: j });
            if (i !== j) arr.push({ rows: j, cols: i });
        }
    }
    arr.sort(function (a, b) {
        return a.rows - b.rows || a.cols - b.cols;
    });
    return arr;
}

function updateArrangementOptions() {
    var countInput = document.getElementById('customCameraCount');
    var select = document.getElementById('arrangementSelect');
    var n = parseInt(countInput.value) || 1;
    if (n < 1) n = 1;
    var arrangements = getArrangements(n);
    select.innerHTML = '';
    if (arrangements.length === 0) {
        select.innerHTML = '<option value="">-- 无有效排列 --</option>';
        return;
    }
    arrangements.forEach(function (item, index) {
        var opt = document.createElement('option');
        opt.value = JSON.stringify(item);
        opt.textContent = item.rows + ' × ' + item.cols;
        if (index === 0) opt.selected = true;
        select.appendChild(opt);
    });
}



// --- 2D绘图辅助 ---
/**
 * 绘制定制面光的光源尺寸矩形，并显示相机开孔（二维阵列），包含孔间距和边距标注，以及开孔尺寸标注
 */
function drawTotalRect(ctx, canvas, totalW, totalH, holeSize, rows, cols, overlap, singleW, singleH) {
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    var rect = canvas.getBoundingClientRect();
    var dpr = window.devicePixelRatio || 1;
    var cssWidth = rect.width || canvas.clientWidth || 800;
    var cssHeight = rect.height || canvas.clientHeight || 600;
    if (cssWidth < 50) cssWidth = 800;
    if (cssHeight < 50) cssHeight = 600;

    canvas.width = cssWidth * dpr;
    canvas.height = cssHeight * dpr;
    canvas.style.width = cssWidth + 'px';
    canvas.style.height = cssHeight + 'px';
    ctx.scale(dpr, dpr);

    var w = cssWidth;
    var h = cssHeight;

    var leftPad = 65;
    var rightPad = 80;
    var topPad = 35;
    var bottomPad = 35;

    var usableW = w - leftPad - rightPad;
    var usableH = h - topPad - bottomPad;
    if (totalW <= 0 || totalH <= 0) return;

    var scaleX = usableW / totalW;
    var scaleY = usableH / totalH;
    var scale = Math.min(scaleX, scaleY);

    var drawW = totalW * scale;
    var drawH = totalH * scale;

    var offsetX = leftPad + (usableW - drawW) / 2;
    var offsetY = topPad + (usableH - drawH) / 2;

    ctx.fillStyle = "rgba(100, 180, 255, 0.25)";
    ctx.fillRect(offsetX, offsetY, drawW, drawH);
    ctx.strokeStyle = "rgba(200, 200, 200, 0.3)";
    ctx.lineWidth = 1;
    ctx.setLineDash([6, 6]);
    ctx.strokeRect(offsetX, offsetY, drawW, drawH);
    ctx.setLineDash([]);

    var firstX = offsetX + singleW * scale;
    var lastX = offsetX + drawW - singleW * scale;
    if (cols === 1) {
        firstX = offsetX + drawW / 2;
        lastX = firstX;
    }
    var stepX = (cols > 1) ? (lastX - firstX) / (cols - 1) : 0;

    var firstY = offsetY + singleH * scale;
    var lastY = offsetY + drawH - singleH * scale;
    if (rows === 1) {
        firstY = offsetY + drawH / 2;
        lastY = firstY;
    }
    var stepY = (rows > 1) ? (lastY - firstY) / (rows - 1) : 0;

    var holeRadiusPx = 0;
    if (holeSize > 0 && rows > 0 && cols > 0) {
        holeRadiusPx = (holeSize * scale) / 2;
        var minRadius = 3;
        if (holeRadiusPx < minRadius) holeRadiusPx = minRadius;
        var maxRadius = Math.min(drawW / (cols + 2), drawH / (rows + 2), 25);
        if (holeRadiusPx > maxRadius) holeRadiusPx = maxRadius;

        function drawCircle(cx, cy, radius) {
            ctx.beginPath();
            ctx.arc(cx, cy, radius, 0, 2 * Math.PI);
            ctx.fillStyle = "rgba(255, 80, 80, 0.5)";
            ctx.fill();
            ctx.strokeStyle = "#cc0000";
            ctx.lineWidth = 2;
            ctx.stroke();
            ctx.strokeStyle = "#cc0000";
            ctx.lineWidth = 1.5;
            ctx.beginPath();
            ctx.moveTo(cx - radius * 0.5, cy);
            ctx.lineTo(cx + radius * 0.5, cy);
            ctx.moveTo(cx, cy - radius * 0.5);
            ctx.lineTo(cx, cy + radius * 0.5);
            ctx.stroke();
        }

        for (var r = 0; r < rows; r++) {
            var cy = firstY + r * stepY;
            for (var c = 0; c < cols; c++) {
                var cx = firstX + c * stepX;
                drawCircle(cx, cy, holeRadiusPx);
            }
        }
    }

    if (holeSize > 0.1 && rows > 0 && cols > 0 && holeRadiusPx > 0) {
        var firstCx = firstX;
        var firstCy = firstY;
        var labelText = "Φ" + holeSize.toFixed(1) + " mm";

        ctx.save();
        ctx.font = "bold 13px Arial";
        var metrics = ctx.measureText(labelText);
        var textWidth = metrics.width;
        var textHeight = 16;
        var padding = 6;
        var arrowLen = 10;

        var labelX, labelY, arrowX, arrowY, align;
        var rightSpace = (offsetX + drawW) - (firstCx + holeRadiusPx + 10);
        var leftSpace = (firstCx - holeRadiusPx - 10) - offsetX;
        var bottomSpace = (offsetY + drawH) - (firstCy + holeRadiusPx + 10);
        var topSpace = (firstCy - holeRadiusPx - 10) - offsetY;

        var avoidBottom = (firstCy + holeRadiusPx + 30 > offsetY + drawH - 10);
        var avoidRight = (firstCx + holeRadiusPx + 30 > offsetX + drawW - 10);

        if (rightSpace > textWidth + padding * 2 + 20 && !avoidRight) {
            labelX = firstCx + holeRadiusPx + arrowLen + 4;
            labelY = firstCy;
            arrowX = firstCx + holeRadiusPx + 2;
            arrowY = firstCy;
            align = 'left';
        } else if (leftSpace > textWidth + padding * 2 + 20) {
            labelX = firstCx - holeRadiusPx - arrowLen - 4;
            labelY = firstCy;
            arrowX = firstCx - holeRadiusPx - 2;
            arrowY = firstCy;
            align = 'right';
        } else if (bottomSpace > textHeight + padding * 2 + 20 && !avoidBottom) {
            labelX = firstCx;
            labelY = firstCy + holeRadiusPx + arrowLen + 4;
            arrowX = firstCx;
            arrowY = firstCy + holeRadiusPx + 2;
            align = 'center';
        } else if (topSpace > textHeight + padding * 2 + 20) {
            labelX = firstCx;
            labelY = firstCy - holeRadiusPx - arrowLen - 4;
            arrowX = firstCx;
            arrowY = firstCy - holeRadiusPx - 2;
            align = 'center';
        } else {
            ctx.restore();
            return;
        }

        ctx.strokeStyle = "#cc0000";
        ctx.lineWidth = 1.2;
        ctx.beginPath();
        ctx.moveTo(arrowX, arrowY);
        ctx.lineTo(labelX, labelY);
        ctx.stroke();

        ctx.fillStyle = "#cc0000";
        ctx.beginPath();
        ctx.arc(arrowX, arrowY, 2.5, 0, 2 * Math.PI);
        ctx.fill();

        var bboxX, bboxY, bboxW, bboxH;
        if (align === 'left') {
            bboxX = labelX - padding;
            bboxY = labelY - textHeight / 2 - padding;
            bboxW = textWidth + padding * 2;
            bboxH = textHeight + padding * 2;
        } else if (align === 'right') {
            bboxX = labelX - textWidth - padding;
            bboxY = labelY - textHeight / 2 - padding;
            bboxW = textWidth + padding * 2;
            bboxH = textHeight + padding * 2;
        } else {
            bboxX = labelX - textWidth / 2 - padding;
            bboxY = labelY - textHeight / 2 - padding;
            bboxW = textWidth + padding * 2;
            bboxH = textHeight + padding * 2;
        }

        bboxX = Math.max(0, Math.min(bboxX, w - bboxW));
        bboxY = Math.max(0, Math.min(bboxY, h - bboxH));

        ctx.shadowColor = "rgba(0,0,0,0.08)";
        ctx.shadowBlur = 6;
        ctx.fillStyle = "rgba(255, 255, 255, 0.92)";
        ctx.strokeStyle = "rgba(200, 200, 200, 0.6)";
        ctx.lineWidth = 0.8;
        ctx.beginPath();
        ctx.roundRect(bboxX, bboxY, bboxW, bboxH, 4);
        ctx.fill();
        ctx.stroke();
        ctx.shadowBlur = 0;

        ctx.fillStyle = "#cc0000";
        ctx.font = "bold 13px Arial";
        ctx.textAlign = align;
        ctx.textBaseline = 'middle';
        ctx.fillText(labelText, labelX, labelY);

        ctx.restore();
    }

    function drawHDimLine(x1, x2, y, label, color) {
        if (Math.abs(x2 - x1) < 1) return;
        var arrowSize = 5;
        ctx.save();
        ctx.strokeStyle = color || "#444";
        ctx.lineWidth = 1.2;
        ctx.fillStyle = color || "#444";
        ctx.font = "11px Arial";
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        ctx.beginPath();
        ctx.moveTo(x1, y);
        ctx.lineTo(x2, y);
        ctx.stroke();
        ctx.beginPath();
        ctx.moveTo(x1, y);
        ctx.lineTo(x1 + arrowSize, y - arrowSize);
        ctx.lineTo(x1 + arrowSize, y + arrowSize);
        ctx.closePath();
        ctx.fill();
        ctx.beginPath();
        ctx.moveTo(x2, y);
        ctx.lineTo(x2 - arrowSize, y - arrowSize);
        ctx.lineTo(x2 - arrowSize, y + arrowSize);
        ctx.closePath();
        ctx.fill();
        ctx.fillStyle = color || "#333";
        ctx.textBaseline = 'top';
        ctx.fillText(label, (x1 + x2) / 2, y + 3);
        ctx.restore();
    }

    function drawVDimLine(x, y1, y2, label, color) {
        if (Math.abs(y2 - y1) < 1) return;
        var arrowSize = 5;
        ctx.save();
        ctx.strokeStyle = color || "#444";
        ctx.lineWidth = 1.2;
        ctx.fillStyle = color || "#444";
        ctx.font = "11px Arial";
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        ctx.beginPath();
        ctx.moveTo(x, y1);
        ctx.lineTo(x, y2);
        ctx.stroke();
        ctx.beginPath();
        ctx.moveTo(x, y1);
        ctx.lineTo(x - arrowSize, y1 + arrowSize);
        ctx.lineTo(x + arrowSize, y1 + arrowSize);
        ctx.closePath();
        ctx.fill();
        ctx.beginPath();
        ctx.moveTo(x, y2);
        ctx.lineTo(x - arrowSize, y2 - arrowSize);
        ctx.lineTo(x + arrowSize, y2 - arrowSize);
        ctx.closePath();
        ctx.fill();
        ctx.fillStyle = color || "#333";
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        ctx.fillText(label, x - 4, (y1 + y2) / 2);
        ctx.restore();
    }

    var tickSize = 6;

    var dimYBottom = offsetY + drawH + 18;
    ctx.strokeStyle = "#666";
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(offsetX, dimYBottom - tickSize);
    ctx.lineTo(offsetX, dimYBottom + tickSize);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(offsetX + drawW, dimYBottom - tickSize);
    ctx.lineTo(offsetX + drawW, dimYBottom + tickSize);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(offsetX, dimYBottom);
    ctx.lineTo(offsetX + drawW, dimYBottom);
    ctx.stroke();
    ctx.fillStyle = "#333";
    ctx.font = "bold 14px Arial";
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    ctx.fillText(totalW.toFixed(1) + ' mm', offsetX + drawW / 2, dimYBottom + 4);

    var dimXRight = offsetX + drawW + 12;
    ctx.strokeStyle = "#666";
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(dimXRight - tickSize, offsetY);
    ctx.lineTo(dimXRight + tickSize, offsetY);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(dimXRight - tickSize, offsetY + drawH);
    ctx.lineTo(dimXRight + tickSize, offsetY + drawH);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(dimXRight, offsetY);
    ctx.lineTo(dimXRight, offsetY + drawH);
    ctx.stroke();
    ctx.fillStyle = "#333";
    ctx.font = "bold 14px Arial";
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.fillText(totalH.toFixed(1) + ' mm', dimXRight + 4, offsetY + drawH / 2);

    var dimYTop = offsetY - 8;
    var dimXLeft = offsetX - 8;
    var colorH = "#0066cc";
    var colorV = "#009933";

    if (cols > 0) {
        var holePositionsX = [];
        for (var c = 0; c < cols; c++) {
            var cx = firstX + c * stepX;
            holePositionsX.push(cx);
        }
        if (holePositionsX.length > 0) {
            var leftEdge = offsetX;
            var rightEdge = offsetX + drawW;
            var firstHoleX = holePositionsX[0];
            var lastHoleX = holePositionsX[holePositionsX.length - 1];

            drawHDimLine(leftEdge, firstHoleX, dimYTop, singleW.toFixed(1) + ' mm', colorH);

            if (cols > 1) {
                var spacingH = (totalW - 2 * singleW) / (cols - 1);
                for (var i = 0; i < cols - 1; i++) {
                    var x1 = firstHoleX + i * (lastHoleX - firstHoleX) / (cols - 1);
                    var x2 = x1 + (lastHoleX - firstHoleX) / (cols - 1);
                    drawHDimLine(x1, x2, dimYTop, spacingH.toFixed(1) + ' mm', colorH);
                }
            }
            drawHDimLine(lastHoleX, rightEdge, dimYTop, singleW.toFixed(1) + ' mm', colorH);
        }
    }

    if (rows > 0) {
        var holePositionsY = [];
        for (var r = 0; r < rows; r++) {
            var cy = firstY + r * stepY;
            holePositionsY.push(cy);
        }
        if (holePositionsY.length > 0) {
            var topEdge = offsetY;
            var bottomEdge = offsetY + drawH;
            var firstHoleY = holePositionsY[0];
            var lastHoleY = holePositionsY[holePositionsY.length - 1];

            drawVDimLine(dimXLeft, topEdge, firstHoleY, singleH.toFixed(1) + ' mm', colorV);

            if (rows > 1) {
                var spacingV = (totalH - 2 * singleH) / (rows - 1);
                for (var j = 0; j < rows - 1; j++) {
                    var y1 = firstHoleY + j * (lastHoleY - firstHoleY) / (rows - 1);
                    var y2 = y1 + (lastHoleY - firstHoleY) / (rows - 1);
                    drawVDimLine(dimXLeft, y1, y2, spacingV.toFixed(1) + ' mm', colorV);
                }
            }
            drawVDimLine(dimXLeft, lastHoleY, bottomEdge, singleH.toFixed(1) + ' mm', colorV);
        }
    }
}

function drawSpotRect(ctx, canvas, width, height) {
    var rect = canvas.getBoundingClientRect();
    var dpr = window.devicePixelRatio || 1;
    var cssWidth = rect.width || canvas.clientWidth || 800;
    var cssHeight = rect.height || canvas.clientHeight || 400;
    if (cssWidth < 50) cssWidth = 800;
    if (cssHeight < 50) cssHeight = 400;
    canvas.width = cssWidth * dpr;
    canvas.height = cssHeight * dpr;
    canvas.style.width = cssWidth + 'px';
    canvas.style.height = cssHeight + 'px';
    ctx.scale(dpr, dpr);

    var w = cssWidth;
    var h = cssHeight;

    var pad = 60;
    var drawW = w - 2 * pad;
    var drawH = h - 2 * pad;
    if (drawW <= 0 || drawH <= 0) return;

    var scaleX = drawW / width;
    var scaleY = drawH / height;
    var scale = Math.min(scaleX, scaleY);
    var rw = width * scale;
    var rh = height * scale;

    var offX = (w - rw) / 2;
    var offY = (h - rh) / 2;

    ctx.fillStyle = 'rgba(0, 120, 255, 0.25)';
    ctx.fillRect(offX, offY, rw, rh);
    ctx.strokeStyle = '#0066cc';
    ctx.lineWidth = 2;
    ctx.strokeRect(offX, offY, rw, rh);

    var dimY = offY + rh + 20;
    ctx.strokeStyle = '#333';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(offX, dimY);
    ctx.lineTo(offX + rw, dimY);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(offX, dimY - 5);
    ctx.lineTo(offX, dimY + 5);
    ctx.moveTo(offX + rw, dimY - 5);
    ctx.lineTo(offX + rw, dimY + 5);
    ctx.stroke();
    ctx.fillStyle = '#333';
    ctx.font = 'bold 14px Arial';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    ctx.fillText(width.toFixed(1) + ' mm', offX + rw / 2, dimY + 5);

    var dimX = offX + rw + 20;
    ctx.beginPath();
    ctx.moveTo(dimX, offY);
    ctx.lineTo(dimX, offY + rh);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(dimX - 5, offY);
    ctx.lineTo(dimX + 5, offY);
    ctx.moveTo(dimX - 5, offY + rh);
    ctx.lineTo(dimX + 5, offY + rh);
    ctx.stroke();
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.fillText(height.toFixed(1) + ' mm', dimX + 5, offY + rh / 2);
}

// ============================================================
//  策略模式：计算策略集合
//  每个策略接收 (params) 返回 { result, threeParams }
// ============================================================
const calcStrategies = {
    /**
     * 标准策略：平面同轴、条光、背光、光斑-平面同轴、光斑-环光
     * 计算视场角 + 光源尺寸
     */
    standard: function (params) {
        const { cameraRow, lensRow, dCam, dLight, sw, sh, config } = params;
        const formulaKey = areaDataManager.lensColumnMap.formula || '计算公式';
        const formula = lensRow[formulaKey] || '';
        if (!formula) { alert('所选镜头未配置计算公式'); return null; }

        let fn;
        try { fn = new Function('WD', 'CMOS', 'return (' + formula + ');'); }
        catch (e) { alert('计算公式语法错误: ' + e.message); return null; }

        const fovW = fn(dCam, sw);
        const fovH = fn(dCam, sh);
        if (isNaN(fovW) || isNaN(fovH) || fovW <= 0 || fovH <= 0) {
            alert('计算出的视野无效，请检查公式或输入值');
            return null;
        }

        const fovH_angle = rad2deg(2 * Math.atan(fovW / (2 * dCam)));
        const fovV_angle = rad2deg(2 * Math.atan(fovH / (2 * dCam)));
        const fovD_angle = rad2deg(2 * Math.atan(Math.hypot(fovW, fovH) / (2 * dCam)));
        const wLight = fovW * (dCam + dLight) / dCam;
        const hLight = fovH * (dCam + dLight) / dCam;

        const result = {
            fovH: fovH_angle.toFixed(1) + ' °',
            fovV: fovV_angle.toFixed(1) + ' °',
            fovD: fovD_angle.toFixed(1) + ' °',
            sizeText: fovW.toFixed(1) + ' mm × ' + fovH.toFixed(1) + ' mm'
        };

        // 根据模块特有字段
        const sub = config.id || currentSub;
        if (sub === 'size-face' || sub === 'size-bar' || sub === 'size-back' ||
            sub === 'size-coax' || sub === 'size-dome') {
            result.lightW = wLight.toFixed(1) + ' mm';
            result.lightH = hLight.toFixed(1) + ' mm';
        } else if (sub === 'spot-face' || sub === 'spot-ring') {
            result.spotText = wLight.toFixed(1) + ' mm × ' + hLight.toFixed(1) + ' mm';
        }

        return {
            result: result,
            threeParams: { fovH_angle, fovV_angle, fovD_angle, fovW, fovH, wLight, hLight, dCam, dLight, sw, sh }
        };
    },

    /**
     * 环光策略：尺寸-环光
     * 增加内外径计算
     */
    ring: function (params) {
        const base = calcStrategies.standard(params);
        if (!base) return null;

        const { fovW, fovH, dCam, dLight } = base.threeParams;
        const fovD_len = Math.hypot(fovW, fovH);
        const innerDiam = Math.max(0, fovD_len * (dCam - dLight) / dCam);
        const outerDiam = fovD_len * (dCam + dLight) / dCam;

        base.result.innerDiam = innerDiam.toFixed(1) + ' mm';
        base.result.outerDiam = outerDiam.toFixed(1) + ' mm';

        return base;
    },

    /**
     * 定制面光策略
     * 多相机阵列 + 开孔 + 画布绘制
     */
    custom: function (params) {
        const base = calcStrategies.standard(params);
        if (!base) return null;

        const { fovW, fovH, dCam, dLight, sw, sh } = base.threeParams;
        const customCameraCount = params.customCameraCount || 1;
        const arrangementVal = params.arrangementVal || '';
        const holeSize = params.holeSize || 0;

        let rows = 1, cols = 1;
        if (arrangementVal) {
            try {
                const obj = JSON.parse(arrangementVal);
                rows = obj.rows || 1;
                cols = obj.cols || 1;
            } catch (e) {
                rows = 1;
                cols = customCameraCount;
            }
        } else {
            rows = 1;
            cols = customCameraCount;
        }

        const overlap = dLight;
        const totalW = cols * fovW - (cols - 1) * overlap;
        const totalH = rows * fovH - (rows - 1) * overlap;
        const lightW = totalW + fovW;
        const lightH = totalH + fovH;

        base.result.lightW = lightW.toFixed(1) + ' mm';
        base.result.lightH = lightH.toFixed(1) + ' mm';
        base.result.totalSize = totalW.toFixed(1) + ' mm × ' + totalH.toFixed(1) + ' mm';

        // 保存画布参数，供外部绘制
        base._canvasParams = { lightW, lightH, holeSize, rows, cols, overlap, fovW, fovH };

        return base;
    },

    /**
     * 同轴光策略
     * 特殊逻辑：传感器 + 发光区 + 光斑投影
     */
    coax: function (params) {
        const { cameraRow, lensRow, dCam, dLight, sw, sh, manualLightLen, manualLightWid } = params;

        const hasSensor = (sw > 0 && sh > 0 && lensRow !== null);
        const hasManualLight = !(isNaN(manualLightLen) || isNaN(manualLightWid) || manualLightLen <= 0 || manualLightWid <= 0);

        if (!hasSensor && !hasManualLight) {
            alert('请提供完整的传感器参数（长边、短边）或手动输入发光区尺寸（长边、短边）');
            return null;
        }

        let lightLen, lightWid, fovW, fovH, wReal, hReal, fovH_angle, fovV_angle, fovD_angle;

        if (hasSensor) {
            const formulaKey = areaDataManager.lensColumnMap.formula || '计算公式';
            const formula = lensRow[formulaKey] || '';
            if (!formula) { alert('所选镜头未配置计算公式'); return null; }

            let fn;
            try { fn = new Function('WD', 'CMOS', 'return (' + formula + ');'); }
            catch (e) { alert('计算公式语法错误: ' + e.message); return null; }

            fovW = fn(dCam, sw);
            fovH = fn(dCam, sh);
            if (isNaN(fovW) || isNaN(fovH) || fovW <= 0 || fovH <= 0) {
                alert('计算出的视野无效');
                return null;
            }
            wReal = fovW;
            hReal = fovH;
            fovH_angle = rad2deg(2 * Math.atan(fovW / (2 * dCam)));
            fovV_angle = rad2deg(2 * Math.atan(fovH / (2 * dCam)));
            fovD_angle = rad2deg(2 * Math.atan(Math.hypot(fovW, fovH) / (2 * dCam)));

            if (hasManualLight) {
                lightLen = manualLightLen;
                lightWid = manualLightWid;
            } else {
                lightLen = fovW * (dCam + dLight) / dCam;
                lightWid = fovH * (dCam + dLight) / dCam;
            }
        } else {
            lightLen = manualLightLen;
            lightWid = manualLightWid;
            wReal = null; hReal = null;
        }

        const scale = dCam / (dCam + dLight);
        const spotW = (lightLen * scale) * 1.2;
        const spotH = (lightWid * scale) * 1.2;

        const result = {
            fovH: (fovH_angle !== undefined) ? fovH_angle.toFixed(1) + ' °' : '--',
            fovV: (fovV_angle !== undefined) ? fovV_angle.toFixed(1) + ' °' : '--',
            fovD: (fovD_angle !== undefined) ? fovD_angle.toFixed(1) + ' °' : '--',
            sizeText: (wReal !== null && hReal !== null) ? wReal.toFixed(1) + ' mm × ' + hReal.toFixed(1) + ' mm' : '--',
            lightText: lightLen.toFixed(1) + ' mm × ' + lightWid.toFixed(1) + ' mm',
            spotText: spotW.toFixed(1) + ' mm × ' + spotH.toFixed(1) + ' mm'
        };

        return {
            result: result,
            threeParams: { spotW, spotH, dCam, dLight, lightLen, lightWid, fovH_angle, fovV_angle, fovD_angle, wReal, hReal },
            isCoax: true
        };
    }
};


// --- 计算主函数（策略模式版） ---
function calculate() {
    var activeBtn = document.querySelector('.tool-toggle-btn.active');
    var tool = activeBtn.dataset.tool;

    if (tool === 'line') {
        alert('线扫功能已移除');
        return;
    }

    document.getElementById('resultCard').style.display = 'block';

    // ===== 1. 获取配置 =====
    const config = getCurrentModuleConfig();
    if (!config) {
        alert('未找到当前模块配置');
        return;
    }
    console.log('📦 当前模块配置:', config.label, config);

    // ===== 2. 收集输入 =====
    const inputs = collectInputsForCurrentModule();
    console.log('📥 收集到的输入:', inputs);

    // ===== 3. 获取相机和镜头行数据 =====
    var cameraSelect = document.getElementById('common-modelSelect');
    var cameraRow = null;
    if (cameraSelect && cameraSelect.value) {
        try { cameraRow = JSON.parse(cameraSelect.value); } catch (e) { }
    } else {
        var camSearch = document.getElementById('common-modelSelectSearch');
        if (camSearch && camSearch.dataset.row) {
            try { cameraRow = JSON.parse(camSearch.dataset.row); } catch (e) { }
        }
    }
    var lensSelect = document.getElementById('common-lensSelect');
    var lensRow = null;
    if (lensSelect && lensSelect.value) {
        try { lensRow = JSON.parse(lensSelect.value); } catch (e) { }
    } else {
        var lensSearch = document.getElementById('common-lensSearch');
        if (lensSearch && lensSearch.dataset.row) {
            try { lensRow = JSON.parse(lensSearch.dataset.row); } catch (e) { }
        }
    }

    // ===== 4. 准备策略参数 =====
    // 从相机行读取传感器尺寸
    var sw = parseFloat(cameraRow && cameraRow['传感器长边']) || 0;
    var sh = parseFloat(cameraRow && cameraRow['传感器短边']) || 0;

    // 获取工作距离（不同模块的输入 ID 不同，使用配置中的 inputs）
    var dCam = 0, dLight = 0;
    config.inputs.forEach(input => {
        if (input.id.endsWith('distCam')) {
            dCam = parseFloat(document.getElementById(input.id)?.value) || 0;
        }
        if (input.id.endsWith('distLight')) {
            dLight = parseFloat(document.getElementById(input.id)?.value) || 0;
        }
    });

    // 同轴光特殊：需要手动发光区尺寸
    var manualLightLen = 0, manualLightWid = 0;
    if (config.calcType === 'coax') {
        manualLightLen = parseFloat(document.getElementById('coax-lightLen')?.value) || 0;
        manualLightWid = parseFloat(document.getElementById('coax-lightWid')?.value) || 0;
        if (isNaN(dCam) || isNaN(dLight) || dCam <= 0 || dLight <= 0) {
            alert('相机工作距离和光源工作距离必须为大于0的数字');
            return;
        }
    }

    // 定制面光特殊：需要相机个数、排列、开孔尺寸
    var customCameraCount = 1, arrangementVal = '', holeSize = 0;
    if (config.calcType === 'custom') {
        customCameraCount = parseFloat(document.getElementById('customCameraCount')?.value) || 1;
        arrangementVal = document.getElementById('arrangementSelect')?.value || '';
        holeSize = parseFloat(document.getElementById('customHoleSize')?.value) || 0;
        if (isNaN(dLight) || dLight < 0) dLight = 0;
    }

    // 验证：非同轴光必须有相机和镜头
    if (config.calcType !== 'coax') {
        if (!cameraRow) { alert('请选择相机型号'); return; }
        if (!lensRow) { alert('请选择镜头型号'); return; }
        if (sw <= 0 || sh <= 0 || isNaN(dCam) || dCam <= 0) {
            alert('传感器数据缺失或相机工作距离无效，请确认相机型号已选且数据完整');
            return;
        }
    }

    // ===== 5. 构建参数对象 =====
    const params = {
        config: config,
        cameraRow: cameraRow,
        lensRow: lensRow,
        dCam: dCam,
        dLight: dLight,
        sw: sw,
        sh: sh,
        manualLightLen: manualLightLen,
        manualLightWid: manualLightWid,
        customCameraCount: customCameraCount,
        arrangementVal: arrangementVal,
        holeSize: holeSize
    };

    // ===== 6. 执行策略 =====
    const strategy = calcStrategies[config.calcType];
    if (!strategy) {
        alert('未知的计算类型: ' + config.calcType);
        return;
    }

    const resultData = strategy(params);
    if (!resultData) return;

    // ===== 7. 渲染结果 =====
    renderResultsForCurrentModule(resultData.result);


    // 更新公共预览
    const common = {
        fovH: document.getElementById('commonFovH'),
        fovV: document.getElementById('commonFovV'),
        fovD: document.getElementById('commonFovD'),
        sizeText: document.getElementById('commonSizeText')
    };
    if (resultData.result) {
        common.fovH.innerText = resultData.result.fovH || '--';
        common.fovV.innerText = resultData.result.fovV || '--';
        common.fovD.innerText = resultData.result.fovD || '--';
        common.sizeText.innerText = resultData.result.sizeText || '--';
    }

    // ===== 8. 3D 场景更新 =====
    if (resultData.isCoax) {
        // 同轴光 3D
        const tp = resultData.threeParams;
        clearThreeScene();
        coaxSceneObjects = [];
        coaxDragParams = null;
        updateThreeSceneCoax(tp.spotW, tp.spotH, tp.dCam, tp.dLight, tp.lightLen, tp.lightWid,
            tp.fovH_angle, tp.fovV_angle, tp.fovD_angle, tp.wReal, tp.hReal);
        initCoaxDrag();
    } else {
        // 常规 3D
        const tp = resultData.threeParams;
        updateThreeScene(tp.fovH_angle, tp.fovV_angle, tp.fovD_angle, tp.fovW, tp.fovH,
            tp.wLight, tp.hLight, tp.dCam, tp.dLight, tp.sw, tp.sh);
    }

    // ===== 9. 定制面光画布绘制 =====
    if (resultData._canvasParams) {
        const cp = resultData._canvasParams;
        const scCtxTotal = document.getElementById('scCanTotal').getContext('2d');
        drawTotalRect(scCtxTotal, document.getElementById('scCanTotal'),
            cp.lightW, cp.lightH, cp.holeSize, cp.rows, cp.cols, cp.overlap, cp.fovW, cp.fovH);
    }
}


// --- 重置 ---
function resetAll() {
    document.querySelectorAll('input').forEach(el => el.value = '');
    clearSubResults(currentSub);
    document.getElementById('resultCard').style.display = 'none';
    document.getElementById('commonFovH').innerText = '--';
    document.getElementById('commonFovV').innerText = '--';
    document.getElementById('commonFovD').innerText = '--';
    document.getElementById('commonSizeText').innerText = '--';

    // 清空所有 2D Canvas（跳过 WebGL Canvas）
    document.querySelectorAll('canvas').forEach(c => {
        try {
            const ctx = c.getContext('2d');
            if (ctx) {
                ctx.clearRect(0, 0, c.width, c.height);
                c.width = 800;
                c.height = 200;
            }
        } catch (e) {
            // 忽略非 2D Canvas 或异常
        }
    });
}

// --- 页面初始化 ---
window.onload = function () {

    // 1. 尝试从用户数据目录加载（仅当在 nw.js 环境中）
    let dataLoaded = false;
    if (typeof require !== 'undefined' && typeof nw !== 'undefined') {
        try {
            // 即使文件不存在，loadFromFile 会返回 false
            dataLoaded = areaDataManager.loadFromFile();
            if (dataLoaded) {
                console.log('✅ 从用户数据目录加载数据成功');
            }
        } catch (e) {
            console.warn('加载用户数据文件失败:', e.message);
        }
    }

    // 2. 如果文件加载失败，尝试从 localStorage 恢复
    if (!dataLoaded) {
        const hasCache = areaDataManager.loadFromLocal();
        if (hasCache) {
            dataLoaded = true;
            console.log('✅ 从 localStorage 恢复数据成功');
        }
    }

    // 3. 更新界面状态（与之前相同）
    const statusSpan = document.getElementById('dataStatus');
    if (dataLoaded) {
        if (statusSpan) {
            statusSpan.textContent = '✅ 已加载: ' + areaDataManager.getVersionInfo();
            statusSpan.style.color = '#4caf50';
        }
        // 填充下拉选择器
        updateAllSearchSelects();
        updateLensSearchSelects();
        if (window.lineSelector && typeof window.lineSelector.onDataUpdated === 'function') {
            window.lineSelector.onDataUpdated();
        }
        document.getElementById('loadingTip').classList.remove('show');
    } else {
        if (statusSpan) {
            statusSpan.textContent = '⚠️ 未找到数据文件，请上传 Excel';
            statusSpan.style.color = '#ff9800';
        }
        const tip = document.getElementById('loadingTip');
        tip.classList.add('show');
        tip.innerHTML = '📂 请点击「更新数据」按钮上传 Excel 文件（支持 .xlsx / .xls）';
        tip.style.color = '#ff9800';
    }

    // 初始化模块配置系统（打印已注册的模块）
    if (window.MODULES_CONFIG) {
        console.log('✅ 模块配置系统已加载，共注册 ' + Object.keys(window.MODULES_CONFIG.modules).length + ' 个模块');
        console.log('   📋 模块列表:', Object.keys(window.MODULES_CONFIG.modules).join(', '));
    } else {
        console.warn('⚠️ 未找到 MODULES_CONFIG，请检查 modules-config.js 是否加载');
    }


    initAllSearchSelects();

    // 公共相机选择后，自动填充所有子模块的传感器输入
    document.getElementById('common-modelSelect').addEventListener('change', function () {
        if (!this.value) return;
        try {
            var row = JSON.parse(this.value);
            console.log('选中的相机数据:', row); // 便于调试

            // 直接从 row 中提取传感器长边和短边（使用已知列名）
            var senW = row['传感器长边'];   // 根据你的实际列名，可能与映射一致
            var senH = row['传感器短边'];
            console.log('传感器长边:', senW, '短边:', senH);

            // 遍历所有子模块前缀，填充对应的输入框
            var prefixes = ['size-', 'custom-', 'spot-', 'coax-'];
            for (var i = 0; i < prefixes.length; i++) {
                var prefix = prefixes[i];
                var elW = document.getElementById(prefix + 'senW');
                var elH = document.getElementById(prefix + 'senH');
                if (elW) elW.value = (senW !== undefined && senW !== null) ? senW : '';
                if (elH) elH.value = (senH !== undefined && senH !== null) ? senH : '';
            }
        } catch (e) {
            console.warn('填充传感器失败:', e);
        }
    });

    const hasCache = areaDataManager.loadFromLocal();
    if (hasCache) {
        const statusSpan = document.getElementById('dataStatus');
        if (statusSpan) {
            statusSpan.textContent = areaDataManager.getVersionInfo();
            statusSpan.style.color = '#4caf50';
        }
        updateAllSearchSelects();
        updateLensSearchSelects();
    }
    document.getElementById('loadingTip').classList.remove('show');
    document.querySelector('.sub-tab-btn.level1[data-group="size"]').classList.add('active');
    document.querySelector('#level2-size .sub-tab-btn.level2[data-sub="size-face"]').classList.add('active');
    document.getElementById('level2-size').style.display = 'flex';
    document.getElementById('level2-spot').style.display = 'none';
    document.querySelector('.sub-result-size-face').style.display = 'block';

    var btns = document.querySelectorAll('.tool-toggle-btn');
    for (var i = 0; i < btns.length; i++) {
        btns[i].addEventListener('click', function () {
            var allBtns = document.querySelectorAll('.tool-toggle-btn');
            for (var j = 0; j < allBtns.length; j++) {
                allBtns[j].classList.remove('active');
            }
            this.classList.add('active');
            switchTool(this.dataset.tool);
        });
    }

    initThreeScene();
    clearSubResults(currentSub);
    setTimeout(function () {
        var container = document.getElementById(activeContainerId);
        if (container) updateRendererSize(container);
    }, 100);

    // 页面加载后，根据当前激活的按钮显示对应模块
    var activeBtn = document.querySelector('.tool-toggle-btn.active');
    if (activeBtn) {
        switchTool(activeBtn.dataset.tool);
    }
};

// ---- 初始化：根据当前激活的子模块渲染输入框 ----
// 默认激活的是 'size-face'（由 HTML 初始状态决定）
const defaultSub = 'size-face';
renderInputsForModule(defaultSub);
console.log('✅ 初始输入渲染完成，模块:', defaultSub);


setTimeout(function () {
    var tip = document.getElementById('loadingTip');
    if (tip.classList.contains('show')) {
        tip.innerHTML = '⚠️ 加载较慢，如无法使用Excel功能，请刷新页面';
        tip.style.color = '#ff9800';
    }
}, 5000);

calcBtn.onclick = calculate;
resetBtn.onclick = resetAll;

window.addEventListener('resize', function () {
    if (document.getElementById('resultCard').style.display === 'block') {
        var currentContainer = document.getElementById(activeContainerId);
        if (currentContainer) updateRendererSize(currentContainer);
    }
});