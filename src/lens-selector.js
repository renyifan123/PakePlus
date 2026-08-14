// ================================================================
//  镜头选型器 - Excel 数据驱动版
//  依赖: window.areaDataManager (Excel数据管理器)
//  支持两种模式：计算工作距离 (wd) 或 计算视野 (fov)
//  所有镜头参数（公式、接圈阈值、靶面）均从 Excel 读取
// ================================================================
(function () {
    'use strict';

    // ----- 全局状态 -----
    var _calcMode = 'wd';                // 'wd' 或 'fov'
    var _selectedCameraRow = null;      // 当前选中的相机行数据
    var _cameraResW = 0;
    var _cameraResH = 0;
    var _cameraTargetSurface = null;    // 当前相机的靶面
    var _lastResults = null;            // 上次计算结果，用于更新UI

    var _currentCmosW = 0;
    var _currentCmosH = 0;
    var _displayFovW = 0;   // 用于显示的完整视野长边
    var _displayFovH = 0;   // 用于显示的完整视野短边



    // ----- DOM 引用 -----
    var lineDistCam = document.getElementById('line-distCam');
    var lineFovW = document.getElementById('line-fovW');
    var lineFovH = document.getElementById('line-fovH');
    var lineCalcBtn = document.getElementById('lineCalcBtn');
    var lineResetBtn = document.getElementById('lineResetBtn');
    var lineResSize = document.getElementById('lineResSize');
    var lineResPrecision = document.getElementById('lineResPrecision');
    var lensCountLabel = document.getElementById('lensCountLabel');


    // ============================================================
    //  1. 数据获取工具函数
    // ============================================================
    function getLensData() {
        if (window.areaDataManager && window.areaDataManager.isLoaded) {
            return window.areaDataManager.lensData || [];
        }
        return [];
    }

    function getLensColumnMap() {
        if (window.areaDataManager && window.areaDataManager.isLoaded) {
            return window.areaDataManager.lensColumnMap || {};
        }
        return {};
    }

    function getCameraData() {
        if (window.areaDataManager && window.areaDataManager.isLoaded) {
            return window.areaDataManager.cameraData || [];
        }
        return [];
    }

    function getCameraColumnMap() {
        if (window.areaDataManager && window.areaDataManager.isLoaded) {
            return window.areaDataManager.cameraColumnMap || {};
        }
        return {};
    }

    // ============================================================
    //  2. 核心计算引擎（动态公式 + 接圈阈值）
    // ============================================================

    /**
     * 执行公式字符串
     * @param {string} formulaStr - 公式字符串，如 "(FOV * 49.7) / CMOS + 20.2"
     * @param {number} cmos - CMOS尺寸（长边或短边，根据上下文）
     * @param {number} fovOrWd - 视野（WD模式）或工作距离（FOV模式）
     * @param {string} mode - 'wd' 或 'fov'
     * @returns {number} 计算结果
     */
    function computeWithFormula(formulaStr, cmos, fovOrWd, mode) {
        try {
            var fn = new Function('CMOS', 'FOV', 'WD', 'return (' + formulaStr + ');');
            if (mode === 'wd') {
                // 计算工作距离，公式中使用 FOV 和 CMOS
                return fn(cmos, fovOrWd, 0);
            } else {
                // 计算视野，公式中使用 WD 和 CMOS
                return fn(cmos, 0, fovOrWd);
            }
        } catch (e) {
            console.warn('公式执行失败:', formulaStr, e);
            return NaN;
        }
    }

    /**
     * 解析接圈阈值字符串
     * 支持格式: ">231:0|174:5|143:10|123:15|111:20|101:25"
     *          或 "231:0|174:5|143:10" (默认 >= )
     * @param {string} thresholdStr
     * @param {number} value - 计算出的 WD 或 FOV 值
     * @returns {number} 接圈长度，-1 表示不合适
     */
    function parseExtensionTube(thresholdStr, value) {
        if (!thresholdStr) return -1;
        // 尝试解析 JSON 格式（兼容旧数据）
        try {
            var jsonData = JSON.parse(thresholdStr);
            if (Array.isArray(jsonData) && jsonData.length > 0) {
                var sorted = jsonData.slice().sort(function (a, b) { return b.maxWD - a.maxWD; });
                for (var i = 0; i < sorted.length; i++) {
                    if (value > sorted[i].maxWD) {
                        return sorted[i].tube;
                    }
                }
                return -1;
            }
        } catch (e) { }

        var pairs = thresholdStr.split('|');
        var rules = [];
        for (var i = 0; i < pairs.length; i++) {
            var parts = pairs[i].split(':');
            if (parts.length !== 2) continue;
            var left = parts[0].replace(/\s/g, '');
            var tube = parseInt(parts[1]);
            // 检测区间格式：min~max
            if (left.indexOf('~') !== -1) {
                var range = left.split('~');
                if (range.length === 2) {
                    var min = parseFloat(range[0]);
                    var max = parseFloat(range[1]);
                    rules.push({ type: 'range', min: min, max: max, tube: tube });
                    continue;
                }
            }
            // 处理 > 或 >=
            var threshold = parseFloat(left.replace('>', ''));
            var operator = left.indexOf('>') !== -1 ? 'gt' : 'ge';
            rules.push({ type: operator, threshold: threshold, tube: tube });
        }
        // 按顺序匹配（区间优先，但这里按顺序，因为区间更具体，可以放在前面）
        for (var j = 0; j < rules.length; j++) {
            var rule = rules[j];
            if (rule.type === 'range') {
                if (value >= rule.min && value <= rule.max) {
                    return rule.tube;
                }
            } else if (rule.type === 'gt') {
                if (value > rule.threshold) return rule.tube;
            } else if (rule.type === 'ge') {
                if (value >= rule.threshold) return rule.tube;
            }
        }
        return -1;
    }

    /**
     * 获取接圈阈值列名（自动检测或从 lensColumnMap 读取）
     */
    function getTubeThresholdKey() {
        var map = getLensColumnMap();
        if (map.tubeThreshold) return map.tubeThreshold;
        // 自动检测
        var lensData = getLensData();
        if (lensData.length > 0) {
            var firstRow = lensData[0];
            var keys = Object.keys(firstRow);
            for (var i = 0; i < keys.length; i++) {
                if (keys[i].indexOf('接圈阈值') !== -1 || keys[i].indexOf('TubeThreshold') !== -1) {
                    return keys[i];
                }
            }
        }
        return '接圈阈值'; // 默认
    }

    // ============================================================
    //  3. 镜头数据遍历与计算
    // ============================================================

    /**
     * 遍历所有镜头，执行公式计算
     * @param {number} cmosParam - 用于计算的 CMOS 值（长边或短边）
     * @param {number} fovOrWdParam - 视野值（WD模式）或工作距离值（FOV模式）
     * @param {string} mode - 'wd' 或 'fov'
     * @returns {object} 按系列分组的结果对象
     */
    function calculateLenses(cmosW, cmosH, fovOrWdParam, mode) {
        var lensData = getLensData();
        var results = {};  // 动态对象

        if (!lensData || lensData.length === 0) {
            return results;
        }

        var colMap = getLensColumnMap();
        var nameKey = colMap.name || '镜头型号';
        var targetKey = colMap.target || '镜头靶面';
        // 如果自动检测失败，尝试从第一行数据中查找包含"靶面"的列
        if (!colMap.target && lensData.length > 0) {
            var firstRow = lensData[0];
            var keys = Object.keys(firstRow);
            for (var i = 0; i < keys.length; i++) {
                if (keys[i].indexOf('靶面') !== -1 || keys[i].toLowerCase().indexOf('target') !== -1) {
                    targetKey = keys[i];
                    break;
                }
            }
        }

        var focalKey = colMap.focal || '焦距';
        var threadKey = colMap.threadSpec || '螺纹规格';
        var formulaKey = (mode === 'wd')
            ? (colMap.distFormula || '距离公式')
            : (colMap.formula || '视野公式');
        var tubeKey = getTubeThresholdKey();

        lensData.forEach(function (row) {
            var model = row[nameKey] || '';
            var series = extractSeries(model);
            // 若系列为'其他'，也保留，后面会渲染
            if (!results[series]) {
                results[series] = [];
            }

            var formula = row[formulaKey] || '';
            if (!formula) {
                results[series].push({
                    model: model,
                    focal: row[focalKey] || '',
                    resultW: NaN,
                    resultH: NaN,
                    extensionTube: -1,
                    isAvailable: false,
                    target: row[targetKey] || ''
                });
                return;
            }

            // 计算长边结果
            var resultW = computeWithFormula(formula, cmosW, fovOrWdParam, mode);
            var resultH = NaN;
            if (mode === 'fov' && cmosH > 0) {
                resultH = computeWithFormula(formula, cmosH, fovOrWdParam, mode);
            }

            // ---- 接圈判断（根据模式选择判断依据） ----
            var extensionTube = -1;
            var isAvailable = false;
            var thresholdStr = row[tubeKey] || '';

            if (mode === 'wd') {
                if (!isNaN(resultW) && isFinite(resultW) && resultW > 0) {
                    extensionTube = thresholdStr ? parseExtensionTube(thresholdStr, resultW) : 0;
                    isAvailable = (extensionTube !== -1);
                }
            } else {
                if (fovOrWdParam > 0) {
                    extensionTube = thresholdStr ? parseExtensionTube(thresholdStr, fovOrWdParam) : 0;
                    isAvailable = (extensionTube !== -1);
                }
            }

            results[series].push({
                model: model,
                focal: row[focalKey] || '',
                resultW: resultW,
                resultH: resultH,
                extensionTube: extensionTube,
                isAvailable: isAvailable,
                target: row[targetKey] || '',
                thread: row[threadKey] || ''
            });
        });

        return results;
    }

    // ============================================================
    //  4. UI 更新函数
    // ============================================================

    /**
     * 更新镜头列表显示
     * @param {object} results - 按系列分组的结果对象
     * @param {string} mode - 'wd' 或 'fov'
     */
    function updateResultsUI(results, mode) {
        var leftCol = document.getElementById('colLeft');
        var rightCol = document.getElementById('colRight');
        // 清空左右列
        leftCol.innerHTML = '';
        rightCol.innerHTML = '';

        var seriesKeys = Object.keys(results);
        if (seriesKeys.length === 0) {
            leftCol.innerHTML = '<div style="padding:20px;text-align:center;color:#b0b8c4;">暂无计算结果</div>';
            return;
        }

        // 使用 organizeSeries 排序
        var organized = organizeSeries(seriesKeys);
        var leftSeries = organized.left;
        var rightSeries = organized.right;

        // 生成左列
        leftSeries.forEach(function (seriesKey) {
            var items = results[seriesKey] || [];
            var groupHtml = createResultGroupHtml(seriesKey, items, mode);
            leftCol.innerHTML += groupHtml;
        });

        // 生成右列
        rightSeries.forEach(function (seriesKey) {
            var items = results[seriesKey] || [];
            var groupHtml = createResultGroupHtml(seriesKey, items, mode);
            rightCol.innerHTML += groupHtml;
        });

        // 靶面匹配（全局更新）
        checkTargetSurfaceMatch(results);
    }

    // 辅助函数：生成计算结果系列组的 HTML
    function createResultGroupHtml(seriesKey, items, mode) {

        // 获取该系列第一支镜头的靶面
        var seriesTarget = '';
        for (var i = 0; i < items.length; i++) {
            if (items[i].target) {
                seriesTarget = items[i].target;
                break;
            }
        }
        var availableCount = items.filter(function (item) { return item.isAvailable; }).length;
        var total = items.length;
        var html = '<div class="lens-series-group" data-series="' + seriesKey + '">';
        // 标题中显示靶面
        html += '<div class="lens-series-header"><span>🔹 ' + seriesKey + ' 系列' + (seriesTarget ? ' ' + seriesTarget : '') + '</span><span class="badge">' + availableCount + '/' + total + '</span></div>';
        html += '<div class="lens-series-body" style="display:flex; flex-wrap:wrap; gap:10px; padding:10px 12px;">';

        items.forEach(function (item) {
            var label = item.focal ? (seriesKey + '-' + item.focal) : item.model;

            var userFovW = parseFloat(lineFovW ? lineFovW.value : 0);
            var userFovH = parseFloat(lineFovH ? lineFovH.value : 0);
            var userDist = parseFloat(lineDistCam ? lineDistCam.value : 0);

            // 工作距离
            var wdDisplay;
            if (mode === 'wd') {
                wdDisplay = (item.isAvailable && !isNaN(item.resultW) && isFinite(item.resultW)) ?
                    item.resultW.toFixed(1) + ' mm' : '不合适';
            } else {
                wdDisplay = (userDist > 0) ? userDist.toFixed(1) + ' mm' : '--';
            }
            var wdValue = (mode === 'wd') ?
                (item.isAvailable && !isNaN(item.resultW) && isFinite(item.resultW) ? item.resultW.toFixed(2) : '') :
                (userDist > 0 ? userDist.toFixed(2) : '');

            // 接圈
            var ringText = item.isAvailable ? (item.extensionTube + ' mm') : '--';

            // 视野
            var fovDisplay;
            if (mode === 'wd') {
                var dispW = _displayFovW || 0;
                var dispH = _displayFovH || 0;
                fovDisplay = (dispW > 0 && dispH > 0) ? (dispW.toFixed(1) + '×' + dispH.toFixed(1) + ' mm') : '--';
            } else {
                if (item.isAvailable && !isNaN(item.resultW) && isFinite(item.resultW)) {
                    var w = item.resultW.toFixed(1);
                    var h = (!isNaN(item.resultH) && isFinite(item.resultH)) ? item.resultH.toFixed(1) : '--';
                    fovDisplay = w + '×' + h + ' mm';
                } else {
                    fovDisplay = '不合适';
                }
            }

            var fovAttrW = (mode === 'fov' && item.isAvailable && typeof item.resultW === 'number' && isFinite(item.resultW)) ? item.resultW.toFixed(2) : '';
            var fovAttrH = (mode === 'fov' && item.isAvailable && typeof item.resultH === 'number' && isFinite(item.resultH)) ? item.resultH.toFixed(2) : '';

            var statusClass = item.isAvailable ? 'ok' : 'unavailable';
            var cardClass = 'lens-card' + (item.isAvailable ? '' : ' unavailable-card');

            html += '<div class="' + cardClass + '" data-model="' + item.model + '" data-fovw="' + fovAttrW + '" data-fovh="' + fovAttrH + '" data-thread="' + (item.thread || '') + '" data-wd="' + wdValue + '">' +
                '<div class="lens-name" title="' + item.model + '">' + label + '</div>' +
                '<div class="lens-detail">' +
                '<div class="row"><span class="label">距离</span><span class="value ' + statusClass + '">' + wdDisplay + '</span></div>' +
                '<div class="row"><span class="label">接圈</span><span class="value ring">' + ringText + '</span></div>' +
                '<div class="row"><span class="label">视野</span><span class="value">' + fovDisplay + '</span></div>' +
                '</div>' +
                '</div>';
        });

        html += '</div></div>';
        return html;
    }

    /**
     * 更新结果区的视野尺寸和像素精度
     */
    function updateResultArea(fovW, fovH) {
        console.log('📊 updateResultArea 被调用');
        console.log('  传入视野: fovW=' + fovW + ', fovH=' + fovH);
        console.log('  全局分辨率: _cameraResW=' + _cameraResW + ', _cameraResH=' + _cameraResH);

        if (lineResSize) {
            if (fovW > 0 && fovH > 0) {
                lineResSize.textContent = fovW.toFixed(2) + ' × ' + fovH.toFixed(2) + ' mm';
            } else if (fovW > 0) {
                lineResSize.textContent = fovW.toFixed(2) + ' mm';
            } else {
                lineResSize.textContent = '--';
            }
        }

        if (lineResPrecision) {
            var resW = _cameraResW || 0;
            var resH = _cameraResH || 0;
            console.log('  用于计算的 resW=' + resW + ', resH=' + resH);

            // 只使用长边计算像素精度
            if (resW > 0 && fovW > 0) {
                var precision = fovW / resW;
                console.log('  计算精度: ' + fovW + ' / ' + resW + ' = ' + precision);
                lineResPrecision.textContent = precision.toFixed(4) + ' mm/pixel';
            } else {
                console.log('  精度计算条件不满足: resW=' + resW + ', fovW=' + fovW);
                lineResPrecision.textContent = '-- (无分辨率)';
            }
        } else {
            console.warn('lineResPrecision 元素未找到');
        }
    }

    // ============================================================
    //  5. 靶面匹配检查 (与C#逻辑一致)
    // ============================================================

    function parseSurfaceSize(surface) {
        try {
            var clean = String(surface).replace('"', '').trim();
            if (clean.indexOf('/') !== -1) {
                var parts = clean.split('/');
                return parseFloat(parts[0]) / parseFloat(parts[1]);
            }
            return parseFloat(clean);
        } catch (e) {
            return 0;
        }
    }

    function getTargetSurfaceBySize(width, height) {
        var cameras = getCameraData();
        var colMap = getCameraColumnMap();
        var chipKey = colMap.chipSize || colMap.ChipSize || '芯片尺寸';
        var targetKey = colMap.target || colMap.TargetSurface || '靶面';

        var targetChipSize = width + '×' + height;
        for (var i = 0; i < cameras.length; i++) {
            var row = cameras[i];
            var chip = row[chipKey] || '';
            if (chip === targetChipSize) {
                return row[targetKey] || '';
            }
        }
        // 备选：模糊匹配
        for (var j = 0; j < cameras.length; j++) {
            var row2 = cameras[j];
            var chip2 = row2[chipKey] || '';
            if (chip2.replace(/\s/g, '') === targetChipSize.replace(/\s/g, '')) {
                return row2[targetKey] || '';
            }
        }
        return null;
    }

    /**
     * 检查靶面匹配，并在系列标题上显示警告
     */
    function checkTargetSurfaceMatch(results) {
        if (!results) return;

        var targetSurface = _cameraTargetSurface;
        if (!targetSurface) {
            var colMap = getCameraColumnMap();
            var senWKey = colMap.senW || '传感器长边';
            var senHKey = colMap.senH || '传感器短边';
            if (_selectedCameraRow) {
                var w = _selectedCameraRow[senWKey];
                var h = _selectedCameraRow[senHKey];
                if (w && h) {
                    targetSurface = getTargetSurfaceBySize(String(w).trim(), String(h).trim());
                }
            }
            if (!targetSurface) {
                clearTargetWarnings();
                return;
            }
        }

        var cameraSize = parseSurfaceSize(targetSurface);

        // 遍历每个系列（动态获取所有系列组）
        document.querySelectorAll('.lens-series-group').forEach(function (group) {
            var header = group.querySelector('.lens-series-header');
            var badge = group.querySelector('.badge');
            if (!header) return;

            var seriesKey = group.dataset.series;
            var items = results[seriesKey] || [];
            var lensTarget = null;
            for (var i = 0; i < items.length; i++) {
                if (items[i].target) {
                    lensTarget = items[i].target;
                    break;
                }
            }
            if (!lensTarget) {
                // 如果结果中没有，尝试从原始数据中查找
                var lensData = getLensData();
                var colMap2 = getLensColumnMap();
                var targetKey2 = colMap2.target || '镜头靶面';
                var nameKey2 = colMap2.name || '镜头型号';
                for (var j = 0; j < lensData.length; j++) {
                    var row = lensData[j];
                    var model = row[nameKey2] || '';
                    if (extractSeries(model) === seriesKey) {
                        lensTarget = row[targetKey2];
                        if (lensTarget) break;
                    }
                }
            }
            if (!lensTarget) return;

            var lensSize = parseSurfaceSize(lensTarget);
            var isMismatch = (cameraSize > lensSize);

            if (isMismatch) {
                header.style.color = '#cc0000';
                header.style.background = '#fff0f0';
                if (badge) {
                    badge.style.color = '#cc0000';
                    badge.style.background = '#ffd0d0';
                }
                var textNode = header.childNodes[0];
                if (textNode && textNode.textContent) {
                    var clean = textNode.textContent.replace(/⚠️/g, '').replace(/靶面不匹配/g, '').trim();
                    // 如果还没有提示，则追加
                    if (clean.indexOf('⚠️') === -1 && clean.indexOf('靶面不匹配') === -1) {
                        textNode.textContent = clean + ' ⚠️靶面不匹配';
                    }
                }
            } else {
                // 清除标记
                header.style.color = '';
                header.style.background = '#f0f4ff';
                if (badge) {
                    badge.style.color = '#4080ff';
                    badge.style.background = '#4080ff20';
                }
                var textNode2 = header.childNodes[0];
                if (textNode2 && textNode2.textContent) {
                    var clean2 = textNode2.textContent.replace(/⚠️/g, '').replace(/靶面不匹配/g, '').trim();
                    textNode2.textContent = clean2;
                }
            }
        });
    }

    function clearTargetWarnings() {
        document.querySelectorAll('.lens-series-group').forEach(function (group) {
            var header = group.querySelector('.lens-series-header');
            var badge = group.querySelector('.badge');
            if (header) {
                header.style.color = '';
                header.style.background = '#f0f4ff';
                var textNode = header.childNodes[0];
                if (textNode && textNode.textContent) {
                    textNode.textContent = textNode.textContent.replace(/⚠️/g, '').trim();
                }
            }
            if (badge) {
                badge.style.color = '#4080ff';
                badge.style.background = '#4080ff20';
            }
        });
    }

    // ============================================================
    //  6. 主计算流程
    // ============================================================

    function doLineCalc() {
        // 1. 获取输入
        var dist = parseFloat(lineDistCam ? lineDistCam.value : 0);
        var fovW = parseFloat(lineFovW ? lineFovW.value : 0);
        var fovH = parseFloat(lineFovH ? lineFovH.value : 0);

        // ----- 清除之前选中的状态 -----
        document.getElementById('selectedLensDisplay').textContent = '点击镜头卡片选择';
        var threadEl = document.getElementById('lineResThread');
        if (threadEl) threadEl.textContent = '--';
        var wdEl = document.getElementById('lineResWD');
        if (wdEl) wdEl.textContent = '--';
        document.querySelectorAll('.lens-card.selected').forEach(function (c) {
            c.classList.remove('selected');
        });

        console.log('📥 输入: 长边=', fovW, '短边=', fovH, '工作距离=', dist);

        // 2. 自动判断计算模式
        // 如果视野有效，优先计算工作距离；否则如果工作距离有效，计算视野
        var hasFov = (fovW > 0 || fovH > 0);
        var hasDist = (dist > 0);

        if (!hasFov && !hasDist) {
            alert('请至少输入一组参数：视野（长边+短边）或工作距离');
            return;
        }

        // 决定模式：有视野 => wd，否则 => fov
        var mode = hasFov ? 'wd' : 'fov';
        _calcMode = mode;   // <--- 添加这一行，更新全局模式

        // 3. 获取CMOS尺寸
        var cmosW = 0, cmosH = 0;
        if (_selectedCameraRow) {
            var colMap = getCameraColumnMap();
            var senWKey = colMap.senW || '传感器长边';
            var senHKey = colMap.senH || '传感器短边';
            cmosW = parseFloat(_selectedCameraRow[senWKey]) || 0;
            cmosH = parseFloat(_selectedCameraRow[senHKey]) || 0;
        }
        if (cmosW <= 0 || cmosH <= 0) {
            var mainSenW = document.getElementById('size-senW');
            var mainSenH = document.getElementById('size-senH');
            if (mainSenW && mainSenH) {
                cmosW = parseFloat(mainSenW.value) || 0;
                cmosH = parseFloat(mainSenH.value) || 0;
            }
        }
        // 保存到全局
        _currentCmosW = cmosW;
        _currentCmosH = cmosH;
        // 计算用于显示的完整视野（根据CMOS比例补齐）
        // 计算用于显示的完整视野（根据CMOS比例补齐）
        var displayFovW = fovW;
        var displayFovH = fovH;
        if (cmosW > 0 && cmosH > 0) {
            var isFovWValid = (fovW > 0);
            var isFovHValid = (fovH > 0);

            if (isFovWValid && !isFovHValid) {
                // 只输入了长边 → 计算短边
                displayFovH = fovW * (cmosH / cmosW);
            } else if (isFovHValid && !isFovWValid) {
                // 只输入了短边 → 计算长边
                displayFovW = fovH * (cmosW / cmosH);
            } else if (isFovWValid && isFovHValid) {
                // 同时输入长边和短边 → 根据CMOS比例调整，保持与CMOS一致
                var cmosRatio = cmosW / cmosH;
                var fovRatio = fovW / fovH;
                if (cmosRatio > fovRatio) {
                    // CMOS更宽，以短边为准，计算长边
                    displayFovW = fovH * (cmosW / cmosH);
                    displayFovH = fovH;
                } else {
                    // 否则以长边为准，计算短边
                    displayFovW = fovW;
                    displayFovH = fovW * (cmosH / cmosW);
                }
            }
        }
        _displayFovW = displayFovW;
        _displayFovH = displayFovH;
        console.log('📐 显示视野: W=', _displayFovW, 'H=', _displayFovH);
        if (cmosW <= 0 || cmosH <= 0) {
            alert('请先选择相机型号或输入CMOS尺寸（长边和短边）');
            return;
        }

        // 4. 根据模式准备参数
        // 4. 根据模式准备参数
        var cmosParam, fovOrWdParam;
        if (mode === 'wd') {
            // 计算工作距离：需要视野
            var hasCMOSW = (cmosW > 0);
            var hasCMOSH = (cmosH > 0);
            var hasFovW = (fovW > 0);
            var hasFovH = (fovH > 0);

            // 模拟 C# 逻辑：根据输入组合选择基准
            if (hasCMOSW && hasCMOSH && hasFovW && hasFovH) {
                var cmosRatio = cmosW / cmosH;
                var fovRatio = fovW / fovH;
                if (cmosRatio > fovRatio) {
                    cmosParam = cmosH;
                    fovOrWdParam = fovH;
                } else {
                    cmosParam = cmosW;
                    fovOrWdParam = fovW;
                }
            } else if (hasCMOSW && hasCMOSH && hasFovW && !hasFovH) {
                // 只输入了视野长边 → 使用 CMOS 长边
                cmosParam = cmosW;
                fovOrWdParam = fovW;
            } else if (hasCMOSW && hasCMOSH && !hasFovW && hasFovH) {
                // 只输入了视野短边 → 使用 CMOS 短边
                cmosParam = cmosH;
                fovOrWdParam = fovH;
            } else if (hasCMOSW && !hasCMOSH && hasFovW) {
                // 只有 CMOS 长边 + 视野长边
                cmosParam = cmosW;
                fovOrWdParam = fovW;
            } else if (!hasCMOSW && hasCMOSH && hasFovH) {
                // 只有 CMOS 短边 + 视野短边
                cmosParam = cmosH;
                fovOrWdParam = fovH;
            } else {
                alert('无法确定计算基准，请确保CMOS尺寸和视野尺寸至少有一对匹配');
                return;
            }

            if (cmosParam <= 0 || fovOrWdParam <= 0) {
                alert('计算参数无效');
                return;
            }
        } else {
            // 计算视野：需要工作距离
            // 使用CMOS长边作为基准（与C#一致）
            if (cmosW <= 0) {
                alert('需要有效的CMOS长边来计算视野');
                return;
            }
            cmosParam = cmosW;
            fovOrWdParam = dist;
        }
        // 计算用 CMOS 参数
        var cmosWParam, cmosHParam;
        if (mode === 'wd') {
            // 工作距离模式：只用一个选定的 CMOS（长边或短边）
            cmosWParam = cmosParam;
            cmosHParam = 0;
        } else {
            // 视野模式：需要长边和短边分别计算
            cmosWParam = cmosW;
            cmosHParam = cmosH;
        }
        var results = calculateLenses(cmosWParam, cmosHParam, fovOrWdParam, mode);
        updateResultsUI(results, mode);

        // 6. 更新结果区
        // 6. 更新结果区
        if (mode === 'wd') {
            updateResultArea(_displayFovW, _displayFovH);
        } else {
            if (lineResSize) lineResSize.textContent = '-- (点击镜头查看)';
            if (lineResPrecision) lineResPrecision.textContent = '--';
        }

        // 7. 靶面匹配
        checkTargetSurfaceMatch(results);
        _lastResults = results;
    }

    // ============================================================
    //  7. 重置功能
    // ============================================================

    function resetLine() {
        if (lineDistCam) lineDistCam.value = '';
        if (lineFovW) lineFovW.value = '';
        if (lineFovH) lineFovH.value = '';
        if (lineResSize) lineResSize.textContent = '--';
        if (lineResPrecision) lineResPrecision.textContent = '--';

        // 重置镜头列表为初始状态（显示所有镜头，无计算结果）
        renderLensList(getLensData());

        // 清除靶面标记
        clearTargetWarnings();

        // 清空相机选择
        var input = document.getElementById('line-modelSelectSearch');
        if (input) input.value = '';
        var hidden = document.getElementById('line-modelSelect');
        if (hidden) hidden.value = '';
        _selectedCameraRow = null;
        _cameraResW = 0;
        _cameraResH = 0;
        _cameraTargetSurface = null;
        _lastResults = null;

        var displayEl = document.getElementById('selectedLensDisplay');
        if (displayEl) displayEl.textContent = '点击镜头卡片选择';
        //重置螺纹规格显示
        var threadEl = document.getElementById('lineResThread');
        if (threadEl) threadEl.textContent = '--';

        // 重置工作距离
        var wdEl = document.getElementById('lineResWD');
        if (wdEl) wdEl.textContent = '--';

        // 清除高亮
        document.querySelectorAll('.lens-card.selected').forEach(function (c) {
            c.classList.remove('selected');
        });
        // 重置显示框
        document.getElementById('selectedLensDisplay').textContent = '点击镜头卡片选择';


        console.log('🔄 镜头选型器已重置');
    }

    // ============================================================
    //  8. 渲染镜头列表（静态展示所有镜头）
    // ============================================================

    function renderLensList(lensData) {
        var leftCol = document.getElementById('colLeft');
        var rightCol = document.getElementById('colRight');
        var emptyHint = document.getElementById('lensEmptyHint');
        var countLabel = document.getElementById('lensCountLabel');

        // 清空左右列
        leftCol.innerHTML = '';
        rightCol.innerHTML = '';

        if (!lensData || lensData.length === 0) {
            if (emptyHint) emptyHint.style.display = 'block';
            if (countLabel) countLabel.textContent = '共 0 支镜头';
            return;
        }
        if (emptyHint) emptyHint.style.display = 'none';

        var colMap = getLensColumnMap();
        var nameKey = colMap.name || '镜头型号';
        var targetKey = colMap.target || '镜头靶面';
        var focalKey = colMap.focal || '焦距';

        // 按系列分组
        var groups = {};
        lensData.forEach(function (row) {
            var model = row[nameKey] || '';
            var series = extractSeries(model);
            if (!groups[series]) groups[series] = [];
            groups[series].push(row);
        });

        var seriesKeys = Object.keys(groups);
        var organized = organizeSeries(seriesKeys);
        var leftSeries = organized.left;
        var rightSeries = organized.right;
        var totalCount = 0;

        // 渲染左列
        leftSeries.forEach(function (series) {
            var items = groups[series];
            totalCount += items.length;
            var groupHtml = createSeriesGroupHtml(series, items, nameKey, targetKey, focalKey);
            leftCol.innerHTML += groupHtml;
        });

        // 渲染右列
        rightSeries.forEach(function (series) {
            var items = groups[series];
            totalCount += items.length;
            var groupHtml = createSeriesGroupHtml(series, items, nameKey, targetKey, focalKey);
            rightCol.innerHTML += groupHtml;
        });

        if (countLabel) countLabel.textContent = '共 ' + totalCount + ' 支镜头';

        // 靶面匹配（清空警告，因为还没计算）
        clearTargetWarnings();
    }

    // 辅助函数：生成静态系列组的HTML
    function createSeriesGroupHtml(series, items, nameKey, targetKey, focalKey) {
        // 获取该系列第一支镜头的靶面
        var seriesTarget = '';
        for (var i = 0; i < items.length; i++) {
            if (items[i][targetKey]) {
                seriesTarget = items[i][targetKey];
                break;
            }
        }
        var badgeCount = items.length;
        var html = '<div class="lens-series-group" data-series="' + series + '">';
        // 标题中显示靶面（不带"靶面"二字，前面加空格）
        html += '<div class="lens-series-header"><span>🔹 ' + series + ' 系列' + (seriesTarget ? ' ' + seriesTarget : '') + '</span><span class="badge">' + badgeCount + '</span></div>';
        html += '<div class="lens-series-body" style="display:flex; flex-wrap:wrap; gap:10px; padding:10px 12px;">';
        items.forEach(function (row) {
            var name = row[nameKey] || '--';
            var focal = row[focalKey] || '--';
            var target = row[targetKey] || '--';
            html += '<div class="lens-card" data-model="' + name + '">' +
                '<div class="lens-name" title="' + name + '">' + name + '</div>' +
                '<div class="lens-detail">' +
                '<div class="row"><span class="label">焦距</span><span class="value">' + focal + 'mm</span></div>' +
                '<div class="row"><span class="label">靶面</span><span class="value">' + target + '</span></div>' +
                '</div>' +
                '</div>';
        });
        html += '</div></div>';
        return html;
    }

    // ============================================================
    //  9. 提取系列前缀
    // ============================================================

    function extractSeries(modelName) {
        if (!modelName) return '其他';
        var s = String(modelName).trim();
        var upper = s.toUpperCase();

        // 方法1：匹配 "CST-XX" 或 "XX-" 格式，提取第一个连字符后的字母部分
        // 例如 "CST-TE3528-25M" -> "TE"
        var match = upper.match(/^[A-Z]+-([A-Z]+)/);
        if (match) {
            return match[1];
        }

        // 方法2：尝试匹配 "XX-" 或 "XX_"，直接提取前缀
        var match2 = upper.match(/^([A-Z]+)[\-_\s]/);
        if (match2) {
            return match2[1];
        }

        // 方法3：如果包含 "CST-"，则提取 "CST-" 后的连续字母
        var cstMatch = upper.match(/^CST-([A-Z]+)/);
        if (cstMatch) {
            return cstMatch[1];
        }

        // 方法4：拆分连字符或下划线，取第一个有效部分（跳过 "CST"）
        var parts = upper.split(/[-_\s]+/);
        if (parts.length >= 2) {
            // 如果第一个部分是 "CST"，则取第二个
            if (parts[0] === 'CST' && parts[1]) {
                return parts[1];
            }
            // 否则取第一个非空部分
            for (var i = 0; i < parts.length; i++) {
                if (parts[i] && parts[i].length > 0 && !/^\d+$/.test(parts[i])) {
                    return parts[i];
                }
            }
        }
        // 特殊处理：若匹配到特定前缀，强制归并
        if (upper.indexOf('SPECIAL') !== -1) return 'SPECIAL';

        // 方法5：最后的备选——如果全部失败，返回 "其他"
        return '其他';
    }

    /**
 * 组织系列到左右列
 * 固定系列：左列 YB, TB, ZB, M；右列 YD, TE, ZD, ZM
 * 其他系列按字母序交替分配到左右列
 * @param {Array} seriesKeys - 所有系列名称列表
 * @returns {Object} { left: [], right: [] }
 */
    function organizeSeries(seriesKeys) {
        var FIXED_LEFT = ['YB', 'TB', 'ZB', 'M'];
        var FIXED_RIGHT = ['YD', 'TE', 'ZD', 'ZM'];

        var fixedLeft = [];
        var fixedRight = [];
        var others = [];

        seriesKeys.forEach(function (key) {
            if (FIXED_LEFT.indexOf(key) !== -1) {
                fixedLeft.push(key);
            } else if (FIXED_RIGHT.indexOf(key) !== -1) {
                fixedRight.push(key);
            } else {
                others.push(key);
            }
        });

        // 按固定顺序排序固定系列（保持原顺序）
        fixedLeft.sort(function (a, b) {
            return FIXED_LEFT.indexOf(a) - FIXED_LEFT.indexOf(b);
        });
        fixedRight.sort(function (a, b) {
            return FIXED_RIGHT.indexOf(a) - FIXED_RIGHT.indexOf(b);
        });

        // 其他系列按字母序排序
        others.sort(function (a, b) {
            return a.localeCompare(b);
        });

        // 交替分配其他系列到左右列，从左列开始
        var left = fixedLeft.slice();
        var right = fixedRight.slice();
        var toLeft = true;
        others.forEach(function (key) {
            if (toLeft) {
                left.push(key);
            } else {
                right.push(key);
            }
            toLeft = !toLeft;
        });

        return { left: left, right: right };
    }

    // ============================================================
    //  10. 相机型号搜索选择器 (复用面阵计算器的数据)
    // ============================================================

    function initLineSearchSelect() {
        var wrapper = document.querySelector('[data-select-id="line-modelSelect"]');
        if (!wrapper) return;
        var input = wrapper.querySelector('.search-select-input');
        var dropdown = wrapper.querySelector('.search-select-dropdown');
        var hiddenSelect = wrapper.querySelector('select');

        var data = getCameraData();
        var colMap = getCameraColumnMap();
        var nameKey = colMap.name || '型号';
        var items = [];
        data.forEach(function (row) {
            var name = row[nameKey] || '';
            if (name) items.push({ name: String(name).trim(), row: row });
        });
        var unique = {};
        var uniqueItems = [];
        items.forEach(function (item) {
            if (!unique[item.name]) {
                unique[item.name] = true;
                uniqueItems.push(item);
            }
        });
        uniqueItems.sort(function (a, b) { return a.name.localeCompare(b.name, 'zh-Hans-CN'); });

        function renderDropdown(filterText) {
            dropdown.innerHTML = '';
            var matched = filterText ? uniqueItems.filter(function (item) {
                return item.name.toLowerCase().indexOf(filterText.toLowerCase()) !== -1;
            }) : uniqueItems.slice(0);
            if (matched.length === 0) {
                var li = document.createElement('li');
                li.className = 'no-match';
                li.textContent = '无匹配型号';
                dropdown.appendChild(li);
            } else {
                matched.forEach(function (item) {
                    var li = document.createElement('li');
                    li.textContent = item.name;
                    li.dataset.row = JSON.stringify(item.row);
                    li.addEventListener('mousedown', function (e) {
                        e.preventDefault();
                        input.value = this.textContent;
                        dropdown.classList.remove('show');
                        if (hiddenSelect) hiddenSelect.value = JSON.stringify(item.row);
                        input.dataset.row = JSON.stringify(item.row);
                        applyCameraToLineInputs(item.row);
                    });
                    dropdown.appendChild(li);
                });
            }
        }

        input.addEventListener('input', function () {
            var val = this.value.trim();
            renderDropdown(val);
            if (val.length > 0) dropdown.classList.add('show');
            else dropdown.classList.remove('show');
        });

        document.addEventListener('click', function (e) {
            if (!wrapper.contains(e.target)) dropdown.classList.remove('show');
        });

        input.addEventListener('keydown', function (e) {
            var lis = dropdown.querySelectorAll('li:not(.no-match)');
            if (lis.length === 0) return;
            var activeIdx = -1;
            lis.forEach(function (li, idx) { if (li.classList.contains('active')) activeIdx = idx; });
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                var next = (activeIdx + 1) % lis.length;
                lis.forEach(function (li) { li.classList.remove('active'); });
                lis[next].classList.add('active');
                lis[next].scrollIntoView({ block: 'nearest' });
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                var prev = (activeIdx - 1 + lis.length) % lis.length;
                lis.forEach(function (li) { li.classList.remove('active'); });
                lis[prev].classList.add('active');
                lis[prev].scrollIntoView({ block: 'nearest' });
            } else if (e.key === 'Enter') {
                e.preventDefault();
                var activeLi = dropdown.querySelector('li.active');
                if (activeLi) activeLi.click();
                else if (lis.length === 1) lis[0].click();
            } else if (e.key === 'Escape') {
                dropdown.classList.remove('show');
            }
        });

        wrapper.updateItems = function (newItems) {
            uniqueItems = newItems.slice();
            var val = input.value.trim();
            if (val) renderDropdown(val);
            else dropdown.classList.remove('show');
        };
        dropdown.classList.remove('show');

        // 如果有选中值，自动应用
        if (hiddenSelect && hiddenSelect.value) {
            try {
                var row = JSON.parse(hiddenSelect.value);
                applyCameraToLineInputs(row);
            } catch (e) { }
        }
    }

    // ----- 应用相机数据到输入 -----
    function applyCameraToLineInputs(row) {
        if (!row) return;
        _selectedCameraRow = row;
        console.log('📷 选中相机行数据:', row);
        var colMap = getCameraColumnMap();
        console.log('📊 相机列映射:', colMap);
        var senWKey = colMap.senW || '传感器长边';
        var senHKey = colMap.senH || '传感器短边';
        console.log('🔑 传感器列名:', senWKey, senHKey);
        var resWKey = colMap.resW || '分辨率长边';
        var resHKey = colMap.resH || '分辨率短边';
        var resW = row[resWKey] || 0;
        var resH = row[resHKey] || 0;
        console.log('📐 传感器尺寸:', resW, resH);
        _cameraResW = parseFloat(resW) || 0;
        _cameraResH = parseFloat(resH) || 0;


        // 获取靶面 - 优先使用列映射中的 target
        var targetKey = colMap.target || colMap.TargetSurface || '靶面';
        _cameraTargetSurface = row[targetKey] || null;
        // 如果仍为空，再尝试从芯片尺寸推断
        if (!_cameraTargetSurface) {
            var chipKey = colMap.chipSize || colMap.ChipSize || '芯片尺寸';
            var chip = row[chipKey] || '';
            if (chip) {
                var parts = chip.split('×');
                if (parts.length === 2) {
                    _cameraTargetSurface = getTargetSurfaceBySize(parts[0].trim(), parts[1].trim());
                }
            }
        }

    }

    // ============================================================
    //  11. 数据更新回调
    // ============================================================

    function onDataUpdated() {
        var lensData = getLensData();
        renderLensList(lensData);
        initLineSearchSelect();

        var hidden = document.getElementById('line-modelSelect');
        if (hidden && hidden.value) {
            try {
                applyCameraToLineInputs(JSON.parse(hidden.value));
            } catch (e) { }
        }

        // 如果有上次的计算结果，重新应用
        if (_lastResults) {
            updateResultsUI(_lastResults, _calcMode);
            checkTargetSurfaceMatch(_lastResults);
        }


        document.getElementById('selectedLensDisplay').textContent = '点击镜头卡片选择';
        document.getElementById('lineResThread').textContent = '--';
        document.getElementById('lineResWD').textContent = '--';
        document.querySelectorAll('.lens-card.selected').forEach(function (c) {
            c.classList.remove('selected');
        });


    }

    // ============================================================
    //  12. 工具切换控制
    // ============================================================

    function setupToolToggle() {
        var btns = document.querySelectorAll('.tool-toggle-btn');
        btns.forEach(function (btn) {
            btn.addEventListener('click', function () {
                var isLine = (this.dataset.tool === 'line');
                var areaWrap = document.getElementById('areaWrap');
                var lineWrap = document.getElementById('lineSelectorWrap');

                // 根据点击的按钮控制显示
                if (areaWrap) areaWrap.style.display = isLine ? 'none' : 'flex';
                if (lineWrap) lineWrap.style.display = isLine ? 'block' : 'none';

                btns.forEach(function (b) { b.classList.remove('active'); });
                this.classList.add('active');

                if (isLine) {
                    setTimeout(function () {
                        onDataUpdated();
                        initLineSearchSelect();
                    }, 100);
                }
            });
        });

        // 默认状态：根据 active 按钮初始化
        var activeBtn = document.querySelector('.tool-toggle-btn.active');
        if (activeBtn) {
            var isLine = (activeBtn.dataset.tool === 'line');
            var areaWrap = document.getElementById('areaWrap');
            var lineWrap = document.getElementById('lineSelectorWrap');
            if (areaWrap) areaWrap.style.display = isLine ? 'none' : 'flex';
            if (lineWrap) lineWrap.style.display = isLine ? 'block' : 'none';
        } else {
            // 默认显示镜头选型器
            var areaWrap = document.getElementById('areaWrap');
            var lineWrap = document.getElementById('lineSelectorWrap');
            if (areaWrap) areaWrap.style.display = 'none';
            if (lineWrap) lineWrap.style.display = 'block';
        }
    }

    // ============================================================
    //  13. 监听数据源更新
    // ============================================================

    function hookDataUpdate() {
        var updateBtn = document.getElementById('updateDataBtn');
        if (updateBtn) {
            updateBtn.addEventListener('click', function () { setTimeout(onDataUpdated, 500); });
        }
        var fileInput = document.getElementById('fileUpload');
        if (fileInput) {
            fileInput.addEventListener('change', function () { setTimeout(onDataUpdated, 600); });
        }
        if (window.updateAreaData) {
            var orig = window.updateAreaData;
            window.updateAreaData = function () {
                orig.apply(this, arguments);
                setTimeout(onDataUpdated, 400);
            };
        }
    }

    // ============================================================
    //  14. 模式切换 (供外部调用)
    // ============================================================

    function setCalcMode(mode) {
        if (mode === 'wd' || mode === 'fov') {
            _calcMode = mode;
            // 更新UI上的按钮状态（如果有）
            var radios = document.querySelectorAll('input[name="calcMode"]');
            radios.forEach(function (radio) {
                radio.checked = (radio.value === mode);
            });
            console.log('📐 计算模式切换为:', mode === 'wd' ? '工作距离' : '视野');
        }
    }

    // ============================================================
    //  14.5 相机型号标签点击 → 弹出相机列表
    // ============================================================
    function bindCameraLabelClick() {
        var modelLabel = document.querySelector('.input-area .form-item:first-child label');
        if (!modelLabel) return;

        modelLabel.style.cursor = 'pointer';
        modelLabel.title = '点击查看所有相机列表';
        modelLabel.addEventListener('click', function () {
            // 获取所有相机数据（去重）
            var items = [];
            if (window.areaDataManager && window.areaDataManager.isLoaded) {
                items = window.areaDataManager.getCameraItems();
            }
            if (!items || items.length === 0) {
                alert('暂无相机数据，请先上传 Excel 文件。');
                return;
            }

            // 构建模态框
            var overlay = document.createElement('div');
            overlay.style.cssText = `
                position: fixed;
                top: 0; left: 0;
                width: 100%; height: 100%;
                background: rgba(0,0,0,0.5);
                display: flex;
                align-items: center;
                justify-content: center;
                z-index: 9999;
                backdrop-filter: blur(2px);
            `;
            var modal = document.createElement('div');
            modal.style.cssText = `
                background: #fff;
                border-radius: 16px;
                max-width: 780px;
                width: 90%;
                max-height: 80vh;
                padding: 24px 28px;
                box-shadow: 0 20px 60px rgba(0,0,0,0.3);
                display: flex;
                flex-direction: column;
                overflow: hidden;
            `;
            var header = document.createElement('div');
            header.style.cssText = `
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 16px;
                flex-shrink: 0;
            `;
            header.innerHTML = `
                <h3 style="margin:0; font-size:20px; font-weight:600; color:#1d2129;">📷 相机列表</h3>
                <button style="background:transparent; border:none; font-size:24px; cursor:pointer; color:#999; padding:0 8px;">&times;</button>
            `;
            var closeBtn = header.querySelector('button');
            closeBtn.addEventListener('click', function () {
                if (document.body.contains(overlay)) {
                    document.body.removeChild(overlay);
                }
            });

            var tableWrap = document.createElement('div');
            tableWrap.style.cssText = `
                overflow-y: auto;
                flex: 1;
            `;
            var table = document.createElement('table');
            table.style.cssText = `
                width: 100%;
                border-collapse: collapse;
                font-size: 14px;
            `;
            // 表头
            var thead = document.createElement('thead');
            thead.innerHTML = `
                <tr style="background:#f0f4ff; border-bottom:2px solid #e8eaed;">
                    <th style="padding:10px 12px; text-align:left; position:sticky; top:0; background:#f0f4ff;">型号</th>
                    <th style="padding:10px 12px; text-align:left; position:sticky; top:0; background:#f0f4ff;">传感器尺寸</th>
                    <th style="padding:10px 12px; text-align:left; position:sticky; top:0; background:#f0f4ff;">分辨率</th>
                    <th style="padding:10px 12px; text-align:left; position:sticky; top:0; background:#f0f4ff;">靶面</th>
                </tr>
            `;
            table.appendChild(thead);

            var tbody = document.createElement('tbody');
            var colMap = getCameraColumnMap();
            var nameKey = colMap.name || '型号';
            var senWKey = colMap.senW || '传感器长边';
            var senHKey = colMap.senH || '传感器短边';
            var resWKey = colMap.resW || '分辨率长边';
            var resHKey = colMap.resH || '分辨率短边';
            var targetKey = colMap.target || '靶面';

            items.forEach(function (item) {
                var row = item.row;
                var name = row[nameKey] || '--';
                var senW = row[senWKey] || '--';
                var senH = row[senHKey] || '--';
                var resW = row[resWKey] || '--';
                var resH = row[resHKey] || '--';
                var target = row[targetKey] || '--';

                var tr = document.createElement('tr');
                tr.style.cssText = `
                    cursor: pointer;
                    border-bottom: 1px solid #f0f2f5;
                    transition: background 0.15s;
                `;
                tr.addEventListener('mouseenter', function () {
                    this.style.background = '#f5f8ff';
                });
                tr.addEventListener('mouseleave', function () {
                    this.style.background = '';
                });
                tr.addEventListener('click', function () {
                    // 填充到搜索框
                    var searchInput = document.getElementById('line-modelSelectSearch');
                    if (searchInput) {
                        searchInput.value = name;
                    }
                    var hiddenSelect = document.getElementById('line-modelSelect');
                    if (hiddenSelect) {
                        hiddenSelect.value = JSON.stringify(row);
                    }
                    // 直接应用相机数据
                    applyCameraToLineInputs(row);
                    // 触发隐藏 select 的 change 事件
                    if (hiddenSelect) {
                        var evt = new Event('change', { bubbles: true });
                        hiddenSelect.dispatchEvent(evt);
                    }
                    // 关闭模态框
                    if (document.body.contains(overlay)) {
                        document.body.removeChild(overlay);
                    }
                });

                tr.innerHTML = `
                    <td style="padding:10px 12px; font-weight:500;">${name}</td>
                    <td style="padding:10px 12px;">${senW} × ${senH} mm</td>
                    <td style="padding:10px 12px;">${resW} × ${resH}</td>
                    <td style="padding:10px 12px;">${target}</td>
                `;
                tbody.appendChild(tr);
            });
            table.appendChild(tbody);
            tableWrap.appendChild(table);
            modal.appendChild(header);
            modal.appendChild(tableWrap);
            overlay.appendChild(modal);
            document.body.appendChild(overlay);

            // 点击遮罩关闭
            overlay.addEventListener('click', function (e) {
                if (e.target === overlay && document.body.contains(overlay)) {
                    document.body.removeChild(overlay);
                }
            });
            // ESC 键关闭
            var escHandler = function (e) {
                if (e.key === 'Escape') {
                    if (document.body.contains(overlay)) {
                        document.body.removeChild(overlay);
                        document.removeEventListener('keydown', escHandler);
                    }
                }
            };
            document.addEventListener('keydown', escHandler);
        });
    }

    // ============================================================
    //  15. 初始化
    // ============================================================

    function init() {
        if (typeof XLSX === 'undefined') {
            setTimeout(init, 500);
            return;
        }

        setupToolToggle();
        initLineSearchSelect();

        if (lineCalcBtn) lineCalcBtn.addEventListener('click', doLineCalc);
        if (lineResetBtn) lineResetBtn.addEventListener('click', resetLine);

        // 模式切换监听（如果页面有单选按钮）
        var radios = document.querySelectorAll('input[name="calcMode"]');
        radios.forEach(function (radio) {
            radio.addEventListener('change', function () {
                if (this.checked) {
                    setCalcMode(this.value);
                }
            });
        });

        hookDataUpdate();
        setTimeout(onDataUpdated, 300);

        // 默认模式
        setCalcMode('wd');


        // 确保所有 lens-series-body 为 flex 容器
        document.querySelectorAll('.lens-series-body').forEach(function (el) {
            el.style.display = 'flex';
            el.style.flexWrap = 'wrap';
            el.style.gap = '10px';
        });

        // 在 init 函数中，添加事件监听（替换原来的）
        var lineWrap = document.getElementById('lineSelectorWrap');
        lineWrap.addEventListener('click', function (e) {
            var card = e.target.closest('.lens-card');
            if (card) {
                // 高亮
                document.querySelectorAll('.lens-card.selected').forEach(function (c) {
                    c.classList.remove('selected');
                });
                card.classList.add('selected');
                // 更新选中镜头显示框
                var model = card.dataset.model || card.querySelector('.lens-name')?.textContent || '';
                if (model) {
                    document.getElementById('selectedLensDisplay').textContent = model;
                }
                // 更新螺纹规格
                var thread = card.dataset.thread || '';
                var threadEl = document.getElementById('lineResThread');
                if (threadEl) {
                    threadEl.textContent = thread || '--';
                }
                // 更新工作距离
                var wd = card.dataset.wd || '';
                var wdEl = document.getElementById('lineResWD');
                if (wdEl) {
                    wdEl.textContent = wd ? wd + ' mm' : '--';
                }

                // 如果是 fov 模式，更新顶部结果区
                if (_calcMode === 'fov') {
                    var fovW = parseFloat(card.dataset.fovw) || 0;
                    var fovH = parseFloat(card.dataset.fovh) || 0;
                    console.log('📌 点击卡片，fovW=' + fovW + ', fovH=' + fovH);
                    if (fovW > 0) {
                        // 只要有长边视野，就更新（短边可能为0，但 updateResultArea 能处理）
                        updateResultArea(fovW, fovH);
                    } else {
                        // 如果没有视野数据，显示“该镜头不合适”
                        if (lineResSize) lineResSize.textContent = '-- (该镜头不合适)';
                        if (lineResPrecision) lineResPrecision.textContent = '--';
                    }
                }
            }
        });

        // 在 init 中，添加事件委托（可放在已有点击事件之后）
        document.getElementById('lineSelectorWrap').addEventListener('click', function (e) {
            var header = e.target.closest('.lens-series-header');
            if (header) {
                var body = header.nextElementSibling;
                if (body && body.classList.contains('lens-series-body')) {
                    if (body.style.display === 'none') {
                        body.style.display = 'flex';
                    } else {
                        body.style.display = 'none';
                    }
                }
            }
        });


        // ---- 添加上传按钮（独立更新） ----
        var top = document.querySelector('.line-selector-top');
        if (!top) return;

        // 确保 top 是 flex 容器
        top.style.display = 'flex';
        top.style.alignItems = 'stretch';
        top.style.gap = '20px'; // 各列间距

        // 找到 .result-area
        var resultArea = document.querySelector('.line-selector-top .result-area');
        if (!resultArea) return;

        // 创建按钮容器，放在 result-area 左边
        var btnWrapper = document.createElement('div');
        btnWrapper.style.cssText = `
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 0 10px;
        background: #fff;
        border-radius: 14px;
        border: 1px solid #eef0f4;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    `;

        var fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = '.xlsx,.xls';
        fileInput.style.display = 'none';
        fileInput.id = 'line-fileUpload';

        var uploadBtn = document.createElement('button');
        uploadBtn.textContent = '📂 更新数据';
        uploadBtn.className = 'line-upload-btn';
        uploadBtn.style.cssText = `
        padding: 10px 14px;
        font-size: 14px;
        border: none;
        border-radius: 8px;
        background: #f0f4ff;
        color: #1d2129;
        cursor: pointer;
        white-space: nowrap;
        font-weight: 500;
        transition: 0.2s;
    `;
        uploadBtn.onmouseover = () => uploadBtn.style.background = '#e0e8ff';
        uploadBtn.onmouseout = () => uploadBtn.style.background = '#f0f4ff';

        btnWrapper.appendChild(fileInput);
        btnWrapper.appendChild(uploadBtn);

        // 将按钮容器插入到 resultArea 的前面
        top.insertBefore(btnWrapper, resultArea);

        // 事件绑定
        uploadBtn.addEventListener('click', function () {
            fileInput.click();
        });

        fileInput.addEventListener('change', function (e) {
            var file = e.target.files[0];
            if (!file) return;
            uploadBtn.textContent = '⏳ 加载中...';
            uploadBtn.disabled = true;
            window.areaDataManager.loadExcel(file)
                .then(function () {
                    // ========== 添加成功弹窗（与面阵一致） ==========
                    alert('数据更新成功！\n文件: ' + window.areaDataManager.fileName + '\n相机: ' + window.areaDataManager.cameraData.length + '条, 镜头: ' + window.areaDataManager.lensData.length + '条');

                    uploadBtn.textContent = '✅ 已更新';
                    setTimeout(function () {
                        uploadBtn.textContent = '📂 更新数据';
                        uploadBtn.disabled = false;
                    }, 2000);
                    // 刷新镜头选型器
                    if (window.lineSelector && window.lineSelector.onDataUpdated) {
                        window.lineSelector.onDataUpdated();
                    } else {
                        onDataUpdated();
                    }
                    // 同步面阵状态（如果有）
                    var statusSpan = document.getElementById('dataStatus');
                    if (statusSpan) {
                        statusSpan.textContent = window.areaDataManager.getVersionInfo();
                        statusSpan.style.color = '#4caf50';
                    }
                })
                .catch(function (err) {
                    alert('加载失败: ' + err.message);
                    uploadBtn.textContent = '📂 更新数据';
                    uploadBtn.disabled = false;
                });
            fileInput.value = '';
        });
        // 绑定相机型号标签点击事件
        bindCameraLabelClick();

        console.log('🔍 镜头选型器已初始化 (Excel驱动版)');
    }

    if (document.readyState === 'complete' || document.readyState === 'interactive') {
        init();
    } else {
        document.addEventListener('DOMContentLoaded', init);
    }

    // ============================================================
    //  16. 暴露公共接口
    // ============================================================

    window.lineSelector = {
        renderLensList: renderLensList,
        onDataUpdated: onDataUpdated,
        doLineCalc: doLineCalc,
        resetLine: resetLine,
        setCalcMode: setCalcMode,
        getCalcMode: function () { return _calcMode; }
    };

})();