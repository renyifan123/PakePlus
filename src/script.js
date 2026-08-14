// 全局镜头图片
const lensImage = new Image();
lensImage.src = "lens.png";
lensImage.onload = () => console.log("镜头图片加载完成");
lensImage.onerror = () => console.warn("未找到lens.png，将使用默认方块");
var gltfLoader = new THREE.GLTFLoader();
var cameraModel = null;   // 相机模型（或组合模型）
var lensModel = null;     // 镜头模型（如果分开）
var externalModels = [];  // 用于存储已加载的模型组，方便管理
var lightModel = null;   // 背光专用的光源模型

// Excel管理类
class ExcelDataManager {
    constructor() {
        // ---- 相机数据（Sheet1） ----
        this.cameraData = [];
        this.cameraColumnMap = {};
        this.cameraModelNames = [];
        this.cameraModelRowMap = {};
        // ---- 镜头数据（Sheet2） ----
        this.lensData = [];
        this.lensColumnMap = {};
        this.lensModelNames = [];
        this.lensModelRowMap = {};
        // ---- 通用 ----
        this.fileName = '';
        this.loadTime = '';
        this.isLoaded = false;
        // 兼容旧代码：指向相机数据
        this.data = [];
        this.columnMap = {};
        this.modelNames = [];
        this.modelRowMap = {};
    }

    loadExcel(file) {
        var self = this;
        return new Promise(function (resolve, reject) {
            var reader = new FileReader();
            reader.onload = function (ev) {
                try {
                    var data = new Uint8Array(ev.target.result);
                    var workbook = XLSX.read(data, { type: 'array' });
                    var sheetNames = workbook.SheetNames;

                    // 重置数据
                    self.cameraData = [];
                    self.lensData = [];

                    for (var i = 0; i < sheetNames.length; i++) {
                        var name = sheetNames[i];
                        var sheet = workbook.Sheets[name];
                        var json = XLSX.utils.sheet_to_json(sheet);
                        if (!json || json.length === 0) continue;

                        // 根据工作表名称分发
                        if (name === 'Sheet1') {
                            self.cameraData = json;
                            self.cameraColumnMap = self.detectColumns(json, 'camera');
                            self.buildCameraIndex();
                        } else if (name === 'Sheet2') {
                            self.lensData = json;
                            self.lensColumnMap = self.detectColumns(json, 'lens');
                            self.buildLensIndex();
                        }
                        // Sheet3 空表忽略
                    }

                    if (self.cameraData.length === 0 && self.lensData.length === 0) {
                        reject(new Error('Excel文件没有有效数据'));
                        return;
                    }

                    // 兼容旧代码：默认指向相机数据
                    self.data = self.cameraData;
                    self.columnMap = self.cameraColumnMap;
                    self.modelNames = self.cameraModelNames;
                    self.modelRowMap = self.cameraModelRowMap;

                    self.fileName = file.name;
                    self.loadTime = new Date().toLocaleString();
                    self.isLoaded = true;
                    resolve(self.data);
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = function () {
                reject(new Error('读取文件失败'));
            };
            reader.readAsArrayBuffer(file);
        });
    }

    // 检测列名（根据表格类型）
    detectColumns(jsonData, type) {
        if (!jsonData || jsonData.length === 0) return {};
        var firstRow = jsonData[0];
        var keys = Object.keys(firstRow);
        var findKey = function (patterns) {
            for (var i = 0; i < patterns.length; i++) {
                var p = patterns[i];
                for (var j = 0; j < keys.length; j++) {
                    var k = keys[j];
                    if (k.indexOf(p) !== -1 || k.toLowerCase().indexOf(p.toLowerCase()) !== -1) {
                        return k;
                    }
                }
            }
            return '';
        };

        if (type === 'camera') {
            return {
                name: findKey(['型号']),
                senW: findKey(['传感器长边']),
                senH: findKey(['传感器短边']),
                resW: findKey(['分辨率长边']),
                resH: findKey(['分辨率短边']),
                target: findKey(['相机靶面', '靶面'])
            };
        } else if (type === 'lens') {
            return {
                name: findKey(['镜头型号']),
                target: findKey(['镜头靶面', '靶面']),
                focal: findKey(['焦距', 'focal', 'f']),
                formula: findKey(['视野公式']),
                distFormula: findKey(['距离公式']),
                tubeThreshold: findKey(['接圈阈值', 'TubeThreshold']),
                threadSpec: findKey(['螺纹规格', '螺纹', 'thread'])
            };
        }
        return {};
    }

    // 建立相机型号索引
    buildCameraIndex() {
        var self = this;
        var nameKey = this.cameraColumnMap.name || '型号';
        this.cameraModelNames = [];
        this.cameraModelRowMap = {};
        for (var i = 0; i < this.cameraData.length; i++) {
            var row = this.cameraData[i];
            var displayName = row[nameKey] || ('型号' + (i + 1));
            displayName = String(displayName).trim();
            if (!displayName) continue;
            if (!this.cameraModelRowMap[displayName]) {
                this.cameraModelRowMap[displayName] = row;
                this.cameraModelNames.push(displayName);
            }
        }
        this.cameraModelNames.sort(function (a, b) {
            return a.localeCompare(b, 'zh-Hans-CN', { sensitivity: 'base' });
        });
    }

    // 建立镜头型号索引
    buildLensIndex() {
        var self = this;
        var nameKey = this.lensColumnMap.name || '镜头型号';
        this.lensModelNames = [];
        this.lensModelRowMap = {};
        for (var i = 0; i < this.lensData.length; i++) {
            var row = this.lensData[i];
            var displayName = row[nameKey] || ('镜头型号' + (i + 1));
            displayName = String(displayName).trim();
            if (!displayName) continue;
            if (!this.lensModelRowMap[displayName]) {
                this.lensModelRowMap[displayName] = row;
                this.lensModelNames.push(displayName);
            }
        }
        this.lensModelNames.sort(function (a, b) {
            return a.localeCompare(b, 'zh-Hans-CN', { sensitivity: 'base' });
        });
    }

    // 获取相机型号列表（用于搜索选择器）
    getCameraItems() {
        var items = [];
        for (var i = 0; i < this.cameraModelNames.length; i++) {
            var name = this.cameraModelNames[i];
            items.push({ name: name, row: this.cameraModelRowMap[name] });
        }
        return items;
    }

    // 获取镜头型号列表
    getLensItems() {
        var items = [];
        for (var i = 0; i < this.lensModelNames.length; i++) {
            var name = this.lensModelNames[i];
            items.push({ name: name, row: this.lensModelRowMap[name] });
        }
        return items;
    }

    // 相机数据填充输入框
    applyCameraToInputs(row, prefix) {
        var map = this.cameraColumnMap;
        var inputMap = {
            'senW': map.senW,
            'senH': map.senH,
            'resW': map.resW,
            'resH': map.resH,
            'target': map.target
        };
        var oldMap = this.detectColumns(this.cameraData, 'camera');
        var focalKey = this.findKeyInRow(this.cameraData[0] || {}, ['焦距', 'focal', 'f']);
        var distCamKey = this.findKeyInRow(this.cameraData[0] || {}, ['相机工作距离', 'cam_dist', '工作距离', '距离']);
        var distLightKey = this.findKeyInRow(this.cameraData[0] || {}, ['光源工作距离', 'light_dist', '光源距离']);

        for (var id in inputMap) {
            var colKey = inputMap[id];
            var element = document.getElementById(prefix + id);
            if (element && colKey && row[colKey] !== undefined && row[colKey] !== null) {
                element.value = row[colKey];
            }
        }
        if (focalKey) {
            var focalEl = document.getElementById(prefix + 'focal');
            if (focalEl && row[focalKey] !== undefined && row[focalKey] !== null) {
                focalEl.value = row[focalKey];
            }
        }
        if (distCamKey) {
            var distCamEl = document.getElementById(prefix + 'distCam');
            if (distCamEl && row[distCamKey] !== undefined && row[distCamKey] !== null) {
                distCamEl.value = row[distCamKey];
            }
        }
        if (distLightKey) {
            var distLightEl = document.getElementById(prefix + 'distLight');
            if (distLightEl && row[distLightKey] !== undefined && row[distLightKey] !== null) {
                distLightEl.value = row[distLightKey];
            }
        }
    }

    // 辅助：在行数据中查找键
    findKeyInRow(row, patterns) {
        if (!row) return '';
        var keys = Object.keys(row);
        for (var i = 0; i < patterns.length; i++) {
            var p = patterns[i];
            for (var j = 0; j < keys.length; j++) {
                var k = keys[j];
                if (k.indexOf(p) !== -1 || k.toLowerCase().indexOf(p.toLowerCase()) !== -1) {
                    return k;
                }
            }
        }
        return '';
    }

    // ---- 兼容旧代码的方法 ----
    autoDetectColumns() {
        return this.cameraColumnMap;
    }

    buildModelIndex() {
        this.buildCameraIndex();
        this.modelNames = this.cameraModelNames;
        this.modelRowMap = this.cameraModelRowMap;
    }

    getVersionInfo() {
        if (!this.isLoaded) return '未加载';
        return this.fileName + ' | ' + this.loadTime + ' | 相机:' + this.cameraData.length + '条, 镜头:' + this.lensData.length + '条';
    }

    populateSelect(selectElement) {
        selectElement.innerHTML = '<option value="">-- 请选择型号 --</option>';
        if (!this.isLoaded || !this.cameraData || this.cameraData.length === 0) return;
        var nameKey = this.cameraColumnMap.name || '型号';
        for (var i = 0; i < this.cameraData.length; i++) {
            var row = this.cameraData[i];
            var opt = document.createElement('option');
            opt.value = JSON.stringify(row);
            var displayName = row[nameKey] || ('型号' + (i + 1));
            opt.textContent = displayName;
            selectElement.appendChild(opt);
        }
    }

    applyToInputs(row, prefix) {
        this.applyCameraToInputs(row, prefix);
    }

    saveToLocal() {
        const saveObj = {
            cameraData: this.cameraData,
            lensData: this.lensData,
            cameraColumnMap: this.cameraColumnMap,
            lensColumnMap: this.lensColumnMap,
            fileName: this.fileName,
            loadTime: this.loadTime,
            data: this.cameraData,
            columnMap: this.cameraColumnMap
        };
        localStorage.setItem("excelCameraData", JSON.stringify(saveObj));
    }

    loadFromLocal() {
        const str = localStorage.getItem("excelCameraData");
        if (!str) return false;
        try {
            const obj = JSON.parse(str);
            this.cameraData = obj.cameraData || obj.data || [];
            this.lensData = obj.lensData || [];
            this.cameraColumnMap = obj.cameraColumnMap || obj.columnMap || {};
            this.lensColumnMap = obj.lensColumnMap || {};
            this.fileName = obj.fileName || '';
            this.loadTime = obj.loadTime || '';
            this.isLoaded = true;
            this.buildCameraIndex();
            this.buildLensIndex();
            this.data = this.cameraData;
            this.columnMap = this.cameraColumnMap;
            this.modelNames = this.cameraModelNames;
            this.modelRowMap = this.cameraModelRowMap;
            return true;
        } catch (e) {
            localStorage.removeItem("excelCameraData");
            return false;
        }
    }

    clearLocal() {
        localStorage.removeItem("excelCameraData");
    }
}

// --- DOM 引用 ---
var areaDataManager = new ExcelDataManager();
var calcBtn = document.getElementById('calcBtn');
var resetBtn = document.getElementById('resetBtn');

var rad2deg = function (r) { return r * 180 / Math.PI; };
var deg2rad = function (d) { return d * Math.PI / 180; };

// Three.js 相关变量
var threeScene, threeCamera, threeRenderer, threeControls;
var labelRenderer;
var labelObjects = [];
var activeContainerId = 'three-sf';
var modelsLoaded = false;
// 同轴光拖拽相关全局变量
var dragControls = null;
var draggableSphere = null;
var coaxSceneObjects = [];
var coaxDragParams = null;
var coaxDistLabel = null;

// --- 子状态 ---
var currentGroup = 'size';
var currentSub = 'size-face';

// --- 辅助函数：清空指定2D画布 ---
function clearCanvas(canvasId) {
    var canvas = document.getElementById(canvasId);
    if (canvas) {
        var ctx = canvas.getContext('2d');
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        canvas.width = 800;
        canvas.height = 200;
        canvas.style.width = '';
        canvas.style.height = '';
    }
}

// --- 辅助函数：清空3D场景中的动态物体 ---
function clearThreeScene() {
    console.log('🔍 clearThreeScene 被调用');

    if (threeScene) {
        var toRemove = [];
        threeScene.children.forEach(function (child) {
            if (child.type !== 'AmbientLight' && child.type !== 'DirectionalLight' &&
                child.type !== 'GridHelper' && child.type !== 'AxesHelper' &&
                child.type !== 'PerspectiveCamera' && child.type !== 'Scene') {
                toRemove.push(child);
            }
        });
        toRemove.forEach(function (child) {
            threeScene.remove(child);
            if (child.geometry) child.geometry.dispose();
            if (child.material) child.material.dispose();
        });

        if (draggableSphere) {
            while (draggableSphere.children.length > 0) {
                var child = draggableSphere.children[0];
                draggableSphere.remove(child);
                if (child.geometry) child.geometry.dispose();
                if (child.material) child.material.dispose();
            }
            threeScene.remove(draggableSphere);
            draggableSphere = null;
        }
        if (dragControls) {
            dragControls.dispose();
            dragControls = null;
        }
        coaxSceneObjects = [];
        coaxDragParams = null;

        modelsLoaded = false;
        cameraModel = null;
        lensModel = null;
    }

    labelObjects.forEach(function (obj) {
        if (threeScene) threeScene.remove(obj);
    });
    labelObjects = [];

    if (coaxDistLabel) {
        threeScene.remove(coaxDistLabel);
        coaxDistLabel = null;
    }

    if (lightModel) {
        threeScene.remove(lightModel);
        lightModel = null;
    }
}

// --- 清除指定子功能的结果数据 ---
function clearSubResults(subId) {
    if (subId === 'size-face') {
        document.getElementById('sfFovH').innerText = '--';
        document.getElementById('sfFovV').innerText = '--';
        document.getElementById('sfFovD').innerText = '--';
        document.getElementById('sfSizeText').innerText = '--';
        document.getElementById('sfLightW').innerText = '--';
        document.getElementById('sfLightH').innerText = '--';
    } else if (subId === 'size-bar') {
        document.getElementById('sbFovH').innerText = '--';
        document.getElementById('sbFovV').innerText = '--';
        document.getElementById('sbFovD').innerText = '--';
        document.getElementById('sbSizeText').innerText = '--';
        document.getElementById('sbLightW').innerText = '--';
        document.getElementById('sbLightH').innerText = '--';
    } else if (subId === 'size-ring') {
        document.getElementById('srFovH').innerText = '--';
        document.getElementById('srFovV').innerText = '--';
        document.getElementById('srFovD').innerText = '--';
        document.getElementById('srSizeText').innerText = '--';
        document.getElementById('srInnerDiam').innerText = '--';
        document.getElementById('srOuterDiam').innerText = '--';
    } else if (subId === 'size-custom') {
        document.getElementById('scFovH').innerText = '--';
        document.getElementById('scFovV').innerText = '--';
        document.getElementById('scFovD').innerText = '--';
        document.getElementById('scSizeText').innerText = '--';
        document.getElementById('scLightW').innerText = '--';
        document.getElementById('scLightH').innerText = '--';
        document.getElementById('scTotalSize').innerText = '--';
        clearCanvas('scCanTotal');
    } else if (subId === 'size-back') {
        document.getElementById('bfFovH').innerText = '--';
        document.getElementById('bfFovV').innerText = '--';
        document.getElementById('bfFovD').innerText = '--';
        document.getElementById('bfSizeText').innerText = '--';
        document.getElementById('bfLightW').innerText = '--';
        document.getElementById('bfLightH').innerText = '--';
    } else if (subId === 'spot-face') {
        document.getElementById('spfFovH').innerText = '--';
        document.getElementById('spfFovV').innerText = '--';
        document.getElementById('spfFovD').innerText = '--';
        document.getElementById('spfSizeText').innerText = '--';
        document.getElementById('spfSpotText').innerText = '--';
    } else if (subId === 'spot-coax') {
        document.getElementById('spcFovH').innerText = '--';
        document.getElementById('spcFovV').innerText = '--';
        document.getElementById('spcFovD').innerText = '--';
        document.getElementById('spcSizeText').innerText = '--';
        document.getElementById('spcLightText').innerText = '--';
        document.getElementById('spcSpotText').innerText = '--';
    } else if (subId === 'spot-ring') {
        document.getElementById('sprFovH').innerText = '--';
        document.getElementById('sprFovV').innerText = '--';
        document.getElementById('sprFovD').innerText = '--';
        document.getElementById('sprSizeText').innerText = '--';
        document.getElementById('sprSpotText').innerText = '--';
    }

    clearThreeScene();
}

function loadExternalModels() {
    if (modelsLoaded) return;
    gltfLoader.load('models/camera.glb', function (gltf) {
        cameraModel = gltf.scene;
        cameraModel.userData.isModel = true;
        cameraModel.position.set(0, 0, -22);
        cameraModel.rotation.set(-Math.PI / 2, 0, 0);
        threeScene.add(cameraModel);
        modelsLoaded = true;
    }, undefined, function (error) {
        console.error('相机模型加载失败', error);
    });

    gltfLoader.load('models/lens.glb', function (gltf) {
        lensModel = gltf.scene;
        lensModel.userData.isModel = true;
        lensModel.position.set(0, 0, -17);
        lensModel.rotation.set(Math.PI / 2, 0, 0);
        threeScene.add(lensModel);
    }, undefined, function (error) {
        console.error('镜头模型加载失败', error);
    });
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
                    var isLens = selectId && selectId.endsWith('lensSelect');
                    if (!isLens) {
                        var prefix = getPrefixForSelect(selectId);
                        if (prefix) {
                            areaDataManager.applyToInputs(row, prefix);
                        }
                    }
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
    var cameraWrappers = document.querySelectorAll('.search-select-wrapper:not([data-select-id$="lensSelect"])');
    cameraWrappers.forEach(function (wrapper) {
        initSearchSelect(wrapper, 'camera');
    });
    var lensWrappers = document.querySelectorAll('.search-select-wrapper[data-select-id$="lensSelect"]');
    lensWrappers.forEach(function (wrapper) {
        initSearchSelect(wrapper, 'lens');
    });
}

function updateLensSearchSelects() {
    if (!areaDataManager.isLoaded) return;
    var items = areaDataManager.getLensItems();
    var wrappers = document.querySelectorAll('.search-select-wrapper[data-select-id$="lensSelect"]');
    wrappers.forEach(function (wrapper) {
        if (wrapper.updateItems) {
            wrapper.updateItems(items);
        }
    });
}

function updateAllSearchSelects() {
    if (!areaDataManager.isLoaded) return;
    var items = areaDataManager.getCameraItems();
    var wrappers = document.querySelectorAll('.search-select-wrapper:not([data-select-id$="lensSelect"])');
    wrappers.forEach(function (wrapper) {
        if (wrapper.updateItems) {
            wrapper.updateItems(items);
        }
    });
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
        var parent = this.closest('.sub-tab-level2');
        parent.querySelectorAll('.sub-tab-btn.level2').forEach(function (b) { b.classList.remove('active'); });
        this.classList.add('active');
        currentSub = this.dataset.sub;

        // -------- 控制输入显示 --------
        document.getElementById('size-inputs').style.display = 'none';
        document.getElementById('custom-inputs').style.display = 'none';
        document.getElementById('spot-inputs').style.display = 'none';
        document.getElementById('coax-inputs').style.display = 'none';

        if (currentSub === 'size-face' || currentSub === 'size-bar' || currentSub === 'size-ring' || currentSub === 'size-back') {
            document.getElementById('size-inputs').style.display = 'grid';
        } else if (currentSub === 'size-custom') {
            document.getElementById('custom-inputs').style.display = 'grid';
            updateArrangementOptions();
        } else if (currentSub === 'spot-face' || currentSub === 'spot-ring') {
            document.getElementById('spot-inputs').style.display = 'grid';
        } else if (currentSub === 'spot-coax') {
            document.getElementById('coax-inputs').style.display = 'grid';
        }

        if (currentSub === 'size-custom') {
            document.querySelectorAll('.custom-item').forEach(function (el) {
                el.classList.remove('hide-item');
            });
        } else {
            document.querySelectorAll('.custom-item').forEach(function (el) {
                el.classList.add('hide-item');
            });
        }

        document.querySelectorAll('.sub-result').forEach(function (el) { el.style.display = 'none'; });
        var target = document.querySelector('.sub-result-' + currentSub);
        if (target) target.style.display = 'block';
        clearSubResults(currentSub);
        switchThreeContainer(currentSub);
    });
});

// --- 切换3D容器 ---
function switchThreeContainer(subId) {
    var containerMap = {
        'size-face': 'three-sf',
        'size-bar': 'three-sb',
        'size-ring': 'three-sr',
        'size-custom': 'three-sc',
        'size-back': 'three-bf',
        'spot-face': 'three-spf',
        'spot-coax': 'three-spc',
        'spot-ring': 'three-spr'
    };
    var newContainerId = containerMap[subId] || 'three-sf';
    if (activeContainerId === newContainerId) return;

    var oldContainer = document.getElementById(activeContainerId);
    var newContainer = document.getElementById(newContainerId);
    if (!newContainer) return;

    if (oldContainer && newContainer && threeRenderer) {
        var rendererDom = threeRenderer.domElement;
        var labelDom = labelRenderer.domElement;
        if (rendererDom && rendererDom.parentNode === oldContainer) {
            newContainer.appendChild(rendererDom);
        }
        if (labelDom && labelDom.parentNode === oldContainer) {
            newContainer.appendChild(labelDom);
        }
        activeContainerId = newContainerId;

        if (!newContainer._resizeObserver) {
            var resizeObserver = new ResizeObserver(function () {
                var currentContainer = document.getElementById(activeContainerId);
                if (currentContainer) updateRendererSize(currentContainer);
            });
            resizeObserver.observe(newContainer);
            newContainer._resizeObserver = resizeObserver;
        }

        var attempts = 0;
        var maxAttempts = 5;
        function tryUpdate() {
            attempts++;
            var width = newContainer.clientWidth || newContainer.offsetWidth;
            var height = newContainer.clientHeight || newContainer.offsetHeight;
            if ((width > 0 && height > 0) || attempts >= maxAttempts) {
                updateRendererSize(newContainer);
                return;
            }
            requestAnimationFrame(function () {
                tryUpdate();
            });
        }
        setTimeout(tryUpdate, 0);
    }
}

// --- 尺寸更新函数 ---
function updateRendererSize(container) {
    if (!container) container = document.getElementById(activeContainerId);
    if (!container) return;

    var width = container.clientWidth || container.offsetWidth;
    var height = container.clientHeight || container.offsetHeight;

    if (width === 0 || height === 0) {
        var rect = container.getBoundingClientRect();
        width = rect.width || 800;
        height = rect.height || 500;
    }

    if (width === 0 || height === 0) {
        var parent = container.parentElement;
        if (parent) {
            width = parent.clientWidth || 800;
            height = parent.clientHeight || 500;
        }
    }

    width = Math.max(width, 1);
    height = Math.max(height, 1);

    if (threeRenderer) {
        threeRenderer.setSize(width, height);
        threeRenderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    }
    if (labelRenderer) {
        labelRenderer.setSize(width, height);
    }
    if (threeCamera) {
        threeCamera.aspect = width / height;
        threeCamera.updateProjectionMatrix();
    }
}

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
    areaDataManager.loadExcel(file)
        .then(function () {
            if (statusSpan) {
                statusSpan.textContent = areaDataManager.getVersionInfo();
                statusSpan.style.color = '#4caf50';
            }
            areaDataManager.populateSelect(document.getElementById('modelSelect'));
            areaDataManager.populateSelect(document.getElementById('custom-modelSelect'));
            areaDataManager.populateSelect(document.getElementById('spot-modelSelect'));
            areaDataManager.populateSelect(document.getElementById('coax-modelSelect'));
            updateAllSearchSelects();
            updateLensSearchSelects();
            areaDataManager.saveToLocal();
            alert('数据更新成功！\n文件: ' + areaDataManager.fileName + '\n相机: ' + areaDataManager.cameraData.length + '条, 镜头: ' + areaDataManager.lensData.length + '条');
        })
        .catch(function (error) {
            if (statusSpan) {
                statusSpan.textContent = '加载失败';
                statusSpan.style.color = '#f44336';
            }
            alert('数据加载失败: ' + error.message);
        });
}

// 点击更新按钮 → 触发隐藏的文件选择框
document.getElementById('updateDataBtn').addEventListener('click', function () {
    document.getElementById('fileUpload').click();
});

// 文件选择后自动加载
document.getElementById('fileUpload').addEventListener('change', updateAreaData);

// 隐藏select的change事件（保留兼容）
document.getElementById('modelSelect').addEventListener('change', function () {
    if (!this.value || !areaDataManager.isLoaded) return;
    try {
        var row = JSON.parse(this.value);
    } catch (error) { }
});

document.getElementById('custom-modelSelect').addEventListener('change', function () {
    if (!this.value || !areaDataManager.isLoaded) return;
    try {
        var row = JSON.parse(this.value);
    } catch (error) { }
});

document.getElementById('spot-modelSelect').addEventListener('change', function () {
    if (!this.value || !areaDataManager.isLoaded) return;
    try {
        var row = JSON.parse(this.value);
    } catch (error) { }
});

document.getElementById('coax-modelSelect').addEventListener('change', function () {
    if (!this.value || !areaDataManager.isLoaded) return;
    try {
        var row = JSON.parse(this.value);
    } catch (error) { }
});

// --- 主工具切换 ---
function switchTool(toolVal) {
    if (toolVal === 'area') {
        // 显示输入卡片和结果卡片
        document.getElementById('inputWrap').style.display = 'block';
        document.getElementById('resultCard').style.display = 'block';

        // 恢复area相关显示
        var areaItems = document.querySelectorAll('.area-item');
        var lineItems = document.querySelectorAll('.line-item');
        var inputWrap = document.getElementById('inputWrap');
        inputWrap.classList.remove('line-scan-wrap');
        areaItems.forEach(function (el) { el.classList.remove('hide-item'); });
        lineItems.forEach(function (el) { el.classList.add('hide-item'); });
        document.querySelector('.sub-tab-wrap').style.display = 'flex';
        document.getElementById('line-inputs') && (document.getElementById('line-inputs').style.display = 'none');

        if (areaDataManager.isLoaded) {
            var statusSpan = document.getElementById('dataStatus');
            if (statusSpan) {
                statusSpan.textContent = areaDataManager.getVersionInfo();
                statusSpan.style.color = '#4caf50';
            }
            areaDataManager.populateSelect(document.getElementById('modelSelect'));
            areaDataManager.populateSelect(document.getElementById('custom-modelSelect'));
            areaDataManager.populateSelect(document.getElementById('spot-modelSelect'));
            areaDataManager.populateSelect(document.getElementById('coax-modelSelect'));
            updateAllSearchSelects();
            updateLensSearchSelects();
        }

        // 根据当前子选项卡显示对应的输入组
        document.getElementById('size-inputs').style.display = 'none';
        document.getElementById('custom-inputs').style.display = 'none';
        document.getElementById('spot-inputs').style.display = 'none';
        document.getElementById('coax-inputs').style.display = 'none';

        if (currentSub === 'size-face' || currentSub === 'size-bar' || currentSub === 'size-ring' || currentSub === 'size-back') {
            document.getElementById('size-inputs').style.display = 'grid';
        } else if (currentSub === 'size-custom') {
            document.getElementById('custom-inputs').style.display = 'grid';
            updateArrangementOptions();
        } else if (currentSub === 'spot-face' || currentSub === 'spot-ring') {
            document.getElementById('spot-inputs').style.display = 'grid';
        } else if (currentSub === 'spot-coax') {
            document.getElementById('coax-inputs').style.display = 'grid';
        }

        document.querySelectorAll('.sub-result').forEach(function (el) { el.style.display = 'none'; });
        var target = document.querySelector('.sub-result-' + currentSub);
        if (target) target.style.display = 'block';
        clearSubResults(currentSub);
        switchThreeContainer(currentSub);
        document.getElementById('resultCard').style.display = 'block';

        document.querySelectorAll('.coax-item').forEach(function (el) {
            el.classList.add('hide-item');
        });

        if (currentSub === 'size-custom') {
            document.querySelectorAll('.custom-item').forEach(function (el) {
                el.classList.remove('hide-item');
            });
        } else {
            document.querySelectorAll('.custom-item').forEach(function (el) {
                el.classList.add('hide-item');
            });
        }
    } else {
        // 线扫模式：隐藏输入和结果卡片，页面空白
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

document.getElementById('customCameraCount').addEventListener('change', function () {
    updateArrangementOptions();
});

// --- Three.js 初始化 ---
function initThreeScene() {
    var container = document.getElementById(activeContainerId);
    if (!container) {
        console.warn('初始容器未找到，使用 three-sf');
        container = document.getElementById('three-sf');
        if (!container) return;
    }

    threeScene = new THREE.Scene();
    threeScene.background = new THREE.Color(0x1a1a2e);

    threeCamera = new THREE.PerspectiveCamera(45, 1, 0.1, 2000);
    threeCamera.position.set(500, 300, 500);
    threeCamera.lookAt(0, 0, 200);

    threeRenderer = new THREE.WebGLRenderer({ antialias: true });
    threeRenderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    container.appendChild(threeRenderer.domElement);

    labelRenderer = new THREE.CSS2DRenderer();
    labelRenderer.domElement.style.position = 'absolute';
    labelRenderer.domElement.style.top = '0';
    labelRenderer.domElement.style.left = '0';
    labelRenderer.domElement.style.width = '100%';
    labelRenderer.domElement.style.height = '100%';
    labelRenderer.domElement.style.pointerEvents = 'none';
    labelRenderer.domElement.style.background = 'transparent';
    container.appendChild(labelRenderer.domElement);

    threeControls = new THREE.OrbitControls(threeCamera, threeRenderer.domElement);
    threeControls.enableDamping = true;
    threeControls.dampingFactor = 0.05;
    threeControls.target.set(0, 0, 200);
    threeControls.zoomToCursor = true;

    var ambientLight = new THREE.AmbientLight(0x404060);
    threeScene.add(ambientLight);
    var dirLight = new THREE.DirectionalLight(0xffffff, 1);
    dirLight.position.set(1, 2, 1);
    threeScene.add(dirLight);
    var backLight = new THREE.DirectionalLight(0xffffff, 0.5);
    backLight.position.set(-1, 0, -1);
    threeScene.add(backLight);

    updateRendererSize(container);

    if (window.ResizeObserver) {
        var resizeObserver = new ResizeObserver(function () {
            var currentContainer = document.getElementById(activeContainerId);
            if (currentContainer) updateRendererSize(currentContainer);
        });
        resizeObserver.observe(container);
        container._resizeObserver = resizeObserver;
    } else {
        window.addEventListener('resize', function () {
            var currentContainer = document.getElementById(activeContainerId);
            if (currentContainer) updateRendererSize(currentContainer);
        });
    }

    animateThree();

    console.log('Three.js 场景初始化成功（带 ResizeObserver）');
}

function animateThree() {
    requestAnimationFrame(animateThree);
    if (threeControls) threeControls.update();
    if (threeRenderer && threeScene && threeCamera) {
        threeRenderer.render(threeScene, threeCamera);
    }
    if (labelRenderer && threeScene && threeCamera) {
        labelRenderer.render(threeScene, threeCamera);
    }
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

// --- 3D场景更新 ---
function updateThreeScene(fovH, fovV, fovD, wReal, hReal, wLight, hLight, dCam, dLight, sw, sh, f) {
    console.log('updateThreeScene 被调用');

    loadExternalModels();

    if (threeScene) {
        var toRemove = [];
        threeScene.children.forEach(function (child) {
            if (child.type !== 'AmbientLight' && child.type !== 'DirectionalLight' &&
                child.type !== 'GridHelper' && child.type !== 'AxesHelper' &&
                child.type !== 'PerspectiveCamera' && child.type !== 'Scene' &&
                !child.userData.isModel) {
                toRemove.push(child);
            }
        });
        toRemove.forEach(function (child) {
            threeScene.remove(child);
            if (child.geometry) child.geometry.dispose();
            if (child.material) child.material.dispose();
        });
    }

    labelObjects.forEach(function (obj) {
        if (threeScene) threeScene.remove(obj);
    });
    labelObjects = [];

    if (!wReal || !hReal || !dCam) {
        console.warn('无效数据，跳过 3D 更新');
        return;
    }

    function addRect(ptsArray, color, dashed) {
        var points = ptsArray.concat(ptsArray[0]);
        var geom = new THREE.BufferGeometry().setFromPoints(points);
        var mat;
        if (dashed) {
            mat = new THREE.LineDashedMaterial({ color: color, dashSize: 8, gapSize: 6 });
        } else {
            mat = new THREE.LineBasicMaterial({ color: color });
        }
        var line = new THREE.Line(geom, mat);
        if (dashed) line.computeLineDistances();
        threeScene.add(line);
    }

    function addLine(p1, p2, color, dashed) {
        var points = [p1, p2];
        var geom = new THREE.BufferGeometry().setFromPoints(points);
        var mat;
        if (dashed) {
            mat = new THREE.LineDashedMaterial({ color: color, dashSize: 6, gapSize: 5 });
        } else {
            mat = new THREE.LineBasicMaterial({ color: color });
        }
        var line = new THREE.Line(geom, mat);
        if (dashed) line.computeLineDistances();
        threeScene.add(line);
    }

    var halfW = wReal / 2, halfH = hReal / 2, zCam = dCam;
    var halfLW = wLight / 2, halfLH = hLight / 2, zLight = dCam + dLight;
    var pts = [
        new THREE.Vector3(-halfW, -halfH, zCam),
        new THREE.Vector3(halfW, -halfH, zCam),
        new THREE.Vector3(halfW, halfH, zCam),
        new THREE.Vector3(-halfW, halfH, zCam)
    ];
    var isRing = (currentSub === 'size-ring' || currentSub === 'spot-ring');
    var isBackLight = (currentSub === 'size-back');

    var origin = new THREE.Vector3(0, 0, 0);
    for (var i = 0; i < 4; i++) addLine(origin, pts[i], 0x88ccff, false);
    addRect(pts, 0xffaa00, false);

    if (isRing) {
        var fovD_len = Math.hypot(wReal, hReal);

        // 计算在工件平面（Z = dCam）上的投影直径
        // 注意：光源位于 Z = dCam + dLight，投影到 Z = dCam 平面时，需要缩放
        // 但由于我们在 Z = dCam - dLight 处绘制投影圆（视锥截面），
        // 实际上我们想要的是光源在视野平面上的投影大小。
        // 正确公式：投影直径 = 视野对角线 * (dCam + dLight) / dCam
        var outerDiam = fovD_len * (dCam + dLight) / dCam;
        var innerDiam = fovD_len * (dCam - dLight) / dCam;

        if (innerDiam < 0) innerDiam = 0;
        var outerRadius = outerDiam / 2;
        var innerRadius = innerDiam / 2;

        function addCircle(center, radius, color, dashed) {
            var segments = 40;
            var points = [];
            for (var i = 0; i <= segments; i++) {
                var theta = (i / segments) * Math.PI * 2;
                var x = center.x + radius * Math.cos(theta);
                var y = center.y + radius * Math.sin(theta);
                var z = center.z;
                points.push(new THREE.Vector3(x, y, z));
            }
            var geom = new THREE.BufferGeometry().setFromPoints(points);
            var mat;
            if (dashed) {
                mat = new THREE.LineDashedMaterial({ color: color, dashSize: 6, gapSize: 5 });
            } else {
                mat = new THREE.LineBasicMaterial({ color: color });
            }
            var line = new THREE.Line(geom, mat);
            if (dashed) line.computeLineDistances();
            threeScene.add(line);
        }

        var projZ = dCam - dLight;
        if (projZ < 0) projZ = 0;
        var centerProj = new THREE.Vector3(0, 0, projZ);
        addCircle(centerProj, outerRadius, 0xff3300, true);
        if (innerRadius > 0.1) {
            addCircle(centerProj, innerRadius, 0xff3300, true);
        }
    } else {
        var ptsLight = [
            new THREE.Vector3(-halfLW, -halfLH, zLight),
            new THREE.Vector3(halfLW, -halfLH, zLight),
            new THREE.Vector3(halfLW, halfLH, zLight),
            new THREE.Vector3(-halfLW, halfLH, zLight)
        ];
        for (var i = 0; i < 4; i++) {
            addLine(pts[i], ptsLight[i], 0xff8800, true);
        }

        if (isBackLight) {
            if (lightModel) {
                threeScene.remove(lightModel);
                lightModel = null;
            }

            var reqZ = zLight;
            console.log('🔦 加载光源模型，目标位置 Z:', reqZ);

            gltfLoader.load(
                'models/light.glb',
                function (gltf) {
                    var model = gltf.scene;
                    model.updateMatrixWorld(true);

                    var box = new THREE.Box3().setFromObject(model);
                    var size = box.getSize(new THREE.Vector3());

                    var group = new THREE.Group();
                    group.position.set(-53, -200, zLight + 10);

                    var scaleX = (wLight * 1.2) / size.x;
                    var scaleY = (hLight * 1.2) / size.z;
                    var scaleZ = 1;
                    group.scale.set(scaleX, scaleY, scaleZ);

                    group.position.set(
                        -53 * group.scale.x,
                        -200 * group.scale.y,
                        zLight + 10
                    );
                    group.add(model);

                    threeScene.add(group);
                    lightModel = group;
                    window.lightGroup = group;
                    window.lightModelRaw = model;
                },
                undefined,
                function (error) {
                    console.error('❌ 光源模型加载失败', error);
                    addRect(ptsLight, 0xff3300, true);
                }
            );
        } else {
            addRect(ptsLight, 0xff3300, true);
            if (lightModel) {
                threeScene.remove(lightModel);
                lightModel = null;
            }
        }

        if (!isBackLight) {
            var projZ = dCam - dLight;
            if (projZ < 0) projZ = 0;
            var projPts = [
                new THREE.Vector3(-halfLW, -halfLH, projZ),
                new THREE.Vector3(halfLW, -halfLH, projZ),
                new THREE.Vector3(halfLW, halfLH, projZ),
                new THREE.Vector3(-halfLW, halfLH, projZ)
            ];
            addRect(projPts, 0xff3300, true);
            for (var i = 0; i < 4; i++) {
                addLine(ptsLight[i], projPts[i], 0xff3300, true);
            }
        }
    }

    function createLabel(text, position, colorClass) {
        var div = document.createElement('div');
        div.className = 'label-3d';
        div.innerHTML = text;
        if (colorClass) {
            var valueSpan = div.querySelector('.label-value');
            if (valueSpan) valueSpan.className = 'label-value ' + colorClass;
        }
        var label = new THREE.CSS2DObject(div);
        label.position.copy(position);
        threeScene.add(label);
        labelObjects.push(label);
        return label;
    }

    var offsetField = Math.max(40, hReal * 0.25) + 20;
    var offsetLight = Math.max(35, hLight * 0.2);
    var offsetProj = Math.max(30, hLight * 0.15) - 5;

    var fieldText = '<span class="label-title">📷 视野</span><br>' +
        '<span class="label-value field">' + wReal.toFixed(1) + ' × ' + hReal.toFixed(1) + ' mm</span>';
    createLabel(fieldText, new THREE.Vector3(-wReal * 1.1, halfH + offsetField + 50, dCam));

    var isRing = (currentSub === 'size-ring' || currentSub === 'spot-ring');
    var innerDiam, outerDiam;
    if (isRing) {
        var diagSensor = Math.hypot(sw, sh);
        innerDiam = diagSensor * (dCam - dLight) / f;
        outerDiam = diagSensor * (dCam + dLight) / f;
        if (innerDiam < 0) innerDiam = 0;
    }

    if (!isRing) {
        var lightText = '<span class="label-title">💡 光源</span><br>' +
            '<span class="label-value light">' + wLight.toFixed(1) + ' × ' + hLight.toFixed(1) + ' mm</span>';
        createLabel(lightText, new THREE.Vector3(-wLight, -halfLH + -offsetLight, zLight));
    }

    var projLabelText;
    if (!isBackLight) {
        if (isRing) {
            projLabelText = '<span class="label-title">⬇️ 环形投影</span><br>' +
                '<span class="label-value light">内径：' + innerDiam.toFixed(2) + ' mm</span> &nbsp;|&nbsp; ' +
                '<span class="label-value light">外径：' + outerDiam.toFixed(2) + ' mm</span>';
        } else {
            projLabelText = '<span class="label-title">⬇️ 光源投影</span><br>' +
                '<span class="label-value light">' + wLight.toFixed(1) + ' × ' + hLight.toFixed(1) + ' mm</span>';
        }
        createLabel(projLabelText, new THREE.Vector3(-wLight, 0, projZ));
    }

    var centerZ = (dCam + dLight - 80) / 2;
    threeControls.target.set(0, 0, centerZ);

    var offsetY = dCam + dLight + 200;
    var offsetZ = 0;
    threeCamera.position.set(0, centerZ + offsetY, centerZ + offsetZ);

    threeCamera.lookAt(0, 0, centerZ);
    threeControls.update();

    console.log('3D 场景更新完成');
}

/**
 * 同轴光3D场景绘制
 */
function updateThreeSceneCoax(spotW, spotH, camDist, lightDist, lightLen, lightWid, fovH, fovV, fovD, wReal, hReal, skipSphere) {

    loadExternalModels();
    coaxDragParams = {
        spotW, spotH, camDist, lightDist, lightLen, lightWid,
        fovH, fovV, fovD, wReal, hReal,
        hasSensor: (wReal !== null && wReal > 0 && hReal !== null && hReal > 0)
    };

    if (threeScene) {
        var toRemove = [];
        threeScene.children.forEach(function (child) {
            if (child === draggableSphere) return;
            if (child === coaxDistLabel) return;
            if (child.type !== 'AmbientLight' && child.type !== 'DirectionalLight' &&
                child.type !== 'GridHelper' && child.type !== 'AxesHelper' &&
                child.type !== 'PerspectiveCamera' && child.type !== 'Scene' &&
                !child.userData.isModel) {
                toRemove.push(child);
            }
        });
        toRemove.forEach(function (child) {
            threeScene.remove(child);
            if (child.geometry) child.geometry.dispose();
            if (child.material) child.material.dispose();
        });
        coaxSceneObjects = [];
    }

    function addRect(ptsArray, color, dashed) {
        var points = ptsArray.concat(ptsArray[0]);
        var geom = new THREE.BufferGeometry().setFromPoints(points);
        var mat = dashed ? new THREE.LineDashedMaterial({ color: color, dashSize: 8, gapSize: 6 })
            : new THREE.LineBasicMaterial({ color: color });
        var line = new THREE.Line(geom, mat);
        if (dashed) line.computeLineDistances();
        threeScene.add(line);
        coaxSceneObjects.push(line);
        return line;
    }

    function addLine(p1, p2, color, dashed) {
        var points = [p1, p2];
        var geom = new THREE.BufferGeometry().setFromPoints(points);
        var mat = dashed ? new THREE.LineDashedMaterial({ color: color, dashSize: 6, gapSize: 5 })
            : new THREE.LineBasicMaterial({ color: color });
        var line = new THREE.Line(geom, mat);
        if (dashed) line.computeLineDistances();
        threeScene.add(line);
        coaxSceneObjects.push(line);
        return line;
    }

    function createLabel(text, position) {
        var div = document.createElement('div');
        div.className = 'label-3d';
        div.innerHTML = text;
        var label = new THREE.CSS2DObject(div);
        label.position.copy(position);
        threeScene.add(label);
        coaxSceneObjects.push(label);
        labelObjects.push(label);
        return label;
    }

    var origin = new THREE.Vector3(0, 0, 0);
    var zWork = camDist;
    var zLight = camDist - lightDist;
    if (zLight < 0) zLight = 0;

    var hasSensor = (wReal !== null && hReal !== null && wReal > 0 && hReal > 0);
    var coneW = hasSensor ? wReal : spotW;
    var coneH = hasSensor ? hReal : spotH;
    var halfConeW = coneW / 2,
        halfConeH = coneH / 2;
    var ptsCone = [
        new THREE.Vector3(-halfConeW, -halfConeH, zWork),
        new THREE.Vector3(halfConeW, -halfConeH, zWork),
        new THREE.Vector3(halfConeW, halfConeH, zWork),
        new THREE.Vector3(-halfConeW, halfConeH, zWork)
    ];
    for (var i = 0; i < 4; i++) {
        addLine(origin, ptsCone[i], 0x88ccff, false);
    }
    addRect(ptsCone, 0x00aaff, false);

    var offsetField = Math.max(40, Math.max(hReal, spotH) * 0.25) + 20;
    var offsetSpot = Math.max(35, spotH * 0.2);
    var offsetLight = Math.max(30, lightWid * 0.15) - 5;

    if (hasSensor) {
        var fieldText = '<span class="label-title">📷 视野</span><br>' +
            '<span class="label-value field">' + wReal.toFixed(1) + ' × ' + hReal.toFixed(1) + ' mm</span>';
        createLabel(fieldText, new THREE.Vector3(-wReal, halfConeH + offsetField + 50, zWork));
    }

    var halfLL = lightLen / 2,
        halfLW = lightWid / 2;
    var ptsLight = [
        new THREE.Vector3(-halfLL, -halfLW, zLight),
        new THREE.Vector3(halfLL, -halfLW, zLight),
        new THREE.Vector3(halfLL, halfLW, zLight),
        new THREE.Vector3(-halfLL, halfLW, zLight)
    ];
    addRect(ptsLight, 0xff8800, true);

    var halfSW = spotW / 2,
        halfSH = spotH / 2;
    var ptsSpot = [
        new THREE.Vector3(-halfSW, -halfSH, zWork),
        new THREE.Vector3(halfSW, -halfSH, zWork),
        new THREE.Vector3(halfSW, halfSH, zWork),
        new THREE.Vector3(-halfSW, halfSH, zWork)
    ];
    addRect(ptsSpot, 0x00ff88, true);

    var spotGeometry = new THREE.BufferGeometry();
    var vertices = [
        -halfSW, -halfSH, zWork,
        halfSW, -halfSH, zWork,
        halfSW, halfSH, zWork,
        -halfSW, -halfSH, zWork,
        halfSW, halfSH, zWork,
        -halfSW, halfSH, zWork
    ];
    spotGeometry.setAttribute('position', new THREE.Float32BufferAttribute(vertices, 3));
    spotGeometry.computeVertexNormals();
    var spotMaterial = new THREE.MeshBasicMaterial({
        color: 0xffffff,
        transparent: true,
        opacity: 0.15,
        side: THREE.DoubleSide,
        depthWrite: false
    });
    var spotMesh = new THREE.Mesh(spotGeometry, spotMaterial);
    threeScene.add(spotMesh);
    coaxSceneObjects.push(spotMesh);

    for (var i = 0; i < 4; i++) {
        addLine(ptsLight[i], ptsSpot[i], 0xffaa44, true);
    }

    var spotText = '<span class="label-title">💡 有效光斑</span><br>' +
        '<span class="label-value field">' + spotW.toFixed(1) + ' × ' + spotH.toFixed(1) + ' mm</span>';
    createLabel(spotText, new THREE.Vector3(-spotW - 50, -halfSH + -offsetSpot, (zWork - 30)));

    var lightText = '<span class="label-title">💡 发光区</span><br>' +
        '<span class="label-value light">' + lightLen.toFixed(1) + ' × ' + lightWid.toFixed(1) + ' mm</span>';
    createLabel(lightText, new THREE.Vector3(-lightLen, 0, zLight));

    if (!skipSphere) {
        var targetZ = zWork * 0.5;
        var camY = (1.4 * camDist) + 350;
        var camZ = 50 + 0.5 * camDist;

        threeControls.target.set(0, 0, targetZ);
        threeCamera.position.set(0, camY, camZ);
        threeCamera.lookAt(0, 0, targetZ);
        threeControls.update();
    }
}

// 初始化同轴光拖拽功能
function initCoaxDrag() {
    if (draggableSphere) {
        threeScene.remove(draggableSphere);
        draggableSphere = null;
    }
    if (dragControls) {
        dragControls.dispose();
        dragControls = null;
    }
    if (coaxDistLabel) {
        threeScene.remove(coaxDistLabel);
        coaxDistLabel = null;
    }

    var params = coaxDragParams;
    if (!params) return;
    var zLight = params.camDist - params.lightDist;
    if (zLight < 0) zLight = 0;

    var sphereGeom = new THREE.SphereGeometry(8, 32, 32);
    var sphereMat = new THREE.MeshStandardMaterial({
        color: 0xff2200,
        emissive: 0xff2200,
        emissiveIntensity: 0.3,
    });
    draggableSphere = new THREE.Mesh(sphereGeom, sphereMat);
    draggableSphere.position.set(0, 0, zLight);
    threeScene.add(draggableSphere);

    var labelDiv = document.createElement('div');
    labelDiv.className = 'label-3d';
    labelDiv.style.background = 'rgba(255,50,0,0.7)';
    labelDiv.style.color = '#fff';
    labelDiv.id = 'coax-dist-label';
    labelDiv.textContent = '光源距离: ' + params.lightDist.toFixed(1) + ' mm';
    coaxDistLabel = new THREE.CSS2DObject(labelDiv);
    coaxDistLabel.position.set(-150, -100, -80);
    threeScene.add(coaxDistLabel);

    dragControls = new THREE.DragControls([draggableSphere], threeCamera, threeRenderer.domElement);
    dragControls.addEventListener('dragstart', function () {
        threeControls.enabled = false;
    });
    dragControls.addEventListener('drag', onCoaxDrag);
    dragControls.addEventListener('dragend', function () {
        threeControls.enabled = true;
    });
}

function onCoaxDrag(event) {
    var pos = event.object.position;
    pos.x = 0;
    pos.y = 0;
    var minZ = 0.1,
        maxZ = coaxDragParams.camDist;
    if (pos.z < minZ) pos.z = minZ;
    if (pos.z > maxZ) pos.z = maxZ;

    var newDLight = coaxDragParams.camDist - pos.z;
    if (newDLight < 0) newDLight = 0;

    document.getElementById('coax-distLight').value = newDLight.toFixed(1);

    coaxDragParams.lightDist = newDLight;
    if (coaxDistLabel) {
        coaxDistLabel.element.textContent = '光源距离: ' + newDLight.toFixed(1) + ' mm';
    }
    var scale = coaxDragParams.camDist / (coaxDragParams.camDist + newDLight);
    var newSpotW = (coaxDragParams.lightLen * scale) * 1.2;
    var newSpotH = (coaxDragParams.lightWid * scale) * 1.2;

    document.getElementById('spcSpotText').innerHTML = newSpotW.toFixed(2) + ' mm × ' + newSpotH.toFixed(2) + ' mm';

    updateThreeSceneCoax(
        newSpotW, newSpotH,
        coaxDragParams.camDist,
        newDLight,
        coaxDragParams.lightLen,
        coaxDragParams.lightWid,
        coaxDragParams.fovH,
        coaxDragParams.fovV,
        coaxDragParams.fovD,
        coaxDragParams.wReal,
        coaxDragParams.hReal,
        true
    );
}

// --- 计算主函数 ---
function calculate() {
    var activeBtn = document.querySelector('.tool-toggle-btn.active');
    var tool = activeBtn.dataset.tool;

    if (tool === 'line') {
        alert('线扫功能已移除');
        return;
    }

    document.getElementById('resultCard').style.display = 'block';

    if (currentSub === 'spot-coax') {
        clearThreeScene();
        coaxSceneObjects = [];
        coaxDragParams = null;
        var sw = parseFloat(document.getElementById('coax-senW').value);
        var sh = parseFloat(document.getElementById('coax-senH').value);
        var dCam = parseFloat(document.getElementById('coax-distCam').value);
        var dLight = parseFloat(document.getElementById('coax-distLight').value);
        var manualLightLen = parseFloat(document.getElementById('coax-lightLen').value);
        var manualLightWid = parseFloat(document.getElementById('coax-lightWid').value);

        if (isNaN(dCam) || isNaN(dLight) || dCam <= 0 || dLight <= 0) {
            alert('相机工作距离和光源工作距离必须为大于0的数字');
            return;
        }

        var hasSensor = !(isNaN(sw) || isNaN(sh) || sw <= 0 || sh <= 0);
        var hasManualLight = !(isNaN(manualLightLen) || isNaN(manualLightWid) || manualLightLen <= 0 || manualLightWid <= 0);

        if (!hasSensor && !hasManualLight) {
            alert('请提供完整的传感器参数（长边、短边）或手动输入发光区尺寸（长边、短边）');
            return;
        }

        var lightLen, lightWid;
        var fovW, fovH, wReal, hReal, fovH_angle, fovV_angle, fovD_angle;

        if (hasSensor) {
            var lensSelect = document.getElementById('coax-lensSelect');
            var lensRow = null;
            if (lensSelect && lensSelect.value) {
                lensRow = JSON.parse(lensSelect.value);
            } else {
                // 备选：从搜索输入框的 dataset 读取
                var searchInput = document.getElementById('coax-lensSearch');
                if (searchInput && searchInput.dataset.row) {
                    lensRow = JSON.parse(searchInput.dataset.row);
                }
            }
            if (!lensRow) {
                alert('请从下拉列表中选择镜头型号（不能仅输入文字）');
                return;
            }
            var formulaKey = areaDataManager.lensColumnMap.formula || '计算公式';
            var formula = lensRow[formulaKey] || '';
            if (!formula) {
                alert('所选镜头未配置计算公式');
                return;
            }
            var fn;
            try {
                fn = new Function('WD', 'CMOS', 'return (' + formula + ');');
            } catch (e) {
                alert('计算公式语法错误: ' + e.message);
                return;
            }
            fovW = fn(dCam, sw);
            fovH = fn(dCam, sh);
            if (isNaN(fovW) || isNaN(fovH) || fovW <= 0 || fovH <= 0) {
                alert('计算出的视野无效');
                return;
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

        var scale = dCam / (dCam + dLight);
        var spotW = (lightLen * scale) * 1.2;
        var spotH = (lightWid * scale) * 1.2;

        document.getElementById('spcFovH').innerText = (fovH_angle !== undefined) ? fovH_angle.toFixed(2) + ' °' : '--';
        document.getElementById('spcFovV').innerText = (fovV_angle !== undefined) ? fovV_angle.toFixed(2) + ' °' : '--';
        document.getElementById('spcFovD').innerText = (fovD_angle !== undefined) ? fovD_angle.toFixed(2) + ' °' : '--';
        document.getElementById('spcSizeText').innerText = (wReal !== null && hReal !== null) ? wReal.toFixed(2) + ' mm × ' + hReal.toFixed(2) + ' mm' : '--';
        document.getElementById('spcLightText').innerText = lightLen.toFixed(2) + ' mm × ' + lightWid.toFixed(2) + ' mm';
        document.getElementById('spcSpotText').innerText = spotW.toFixed(2) + ' mm × ' + spotH.toFixed(2) + ' mm';

        updateThreeSceneCoax(spotW, spotH, dCam, dLight, lightLen, lightWid, fovH_angle, fovV_angle, fovD_angle, wReal, hReal);
        initCoaxDrag();
        return;
    }

    var prefix = '';
    if (currentSub === 'size-face' || currentSub === 'size-bar' || currentSub === 'size-ring' || currentSub === 'size-back') {
        prefix = 'size-';
    } else if (currentSub === 'size-custom') {
        prefix = 'custom-';
    } else if (currentSub === 'spot-face' || currentSub === 'spot-ring') {
        prefix = 'spot-';
    } else {
        alert('未知的计算模式');
        return;
    }

    var sw = parseFloat(document.getElementById(prefix + 'senW').value);
    var sh = parseFloat(document.getElementById(prefix + 'senH').value);
    var dCam = parseFloat(document.getElementById(prefix + 'distCam').value);
    var dLight = parseFloat(document.getElementById(prefix + 'distLight').value);

    if (currentSub === 'size-custom') {
        if (isNaN(dLight) || dLight < 0) dLight = 0;
    }

    if (isNaN(sw) || isNaN(sh) || isNaN(dCam) || sw <= 0 || sh <= 0 || dCam <= 0) {
        alert('传感器长边、短边、相机工作距离必须为大于0的数字');
        return;
    }

    var lensSelectId = prefix + 'lensSelect';
    var lensSelect = document.getElementById(lensSelectId);
    var lensRow = null;

    if (lensSelect && lensSelect.value) {
        lensRow = JSON.parse(lensSelect.value);
    } else {
        var searchInputId = prefix + 'lensSearch';
        var searchInput = document.getElementById(searchInputId);
        if (searchInput && searchInput.dataset.row) {
            lensRow = JSON.parse(searchInput.dataset.row);
        }
    }

    if (!lensRow) {
        alert('请选择镜头型号');
        return;
    }
    var formulaKey = areaDataManager.lensColumnMap.formula || '计算公式';
    var formula = lensRow[formulaKey] || '';
    if (!formula) {
        alert('所选镜头未配置计算公式，请检查 Excel 中的“计算公式”列');
        return;
    }
    var fn;
    try {
        fn = new Function('WD', 'CMOS', 'return (' + formula + ');');
    } catch (e) {
        alert('计算公式语法错误: ' + e.message);
        return;
    }
    var fovW = fn(dCam, sw);
    var fovH = fn(dCam, sh);
    if (isNaN(fovW) || isNaN(fovH) || fovW <= 0 || fovH <= 0) {
        alert('计算出的视野无效，请检查公式或输入值');
        return;
    }
    var fovH_angle = rad2deg(2 * Math.atan(fovW / (2 * dCam)));
    var fovV_angle = rad2deg(2 * Math.atan(fovH / (2 * dCam)));
    var fovD_angle = rad2deg(2 * Math.atan(Math.hypot(fovW, fovH) / (2 * dCam)));
    var wLight = fovW * (dCam + dLight) / dCam;
    var hLight = fovH * (dCam + dLight) / dCam;

    if (currentSub === 'size-face') {
        document.getElementById('sfFovH').innerText = fovH_angle.toFixed(2) + ' °';
        document.getElementById('sfFovV').innerText = fovV_angle.toFixed(2) + ' °';
        document.getElementById('sfFovD').innerText = fovD_angle.toFixed(2) + ' °';
        document.getElementById('sfSizeText').innerText = fovW.toFixed(2) + ' mm × ' + fovH.toFixed(2) + ' mm';
        document.getElementById('sfLightW').innerText = wLight.toFixed(2) + ' mm';
        document.getElementById('sfLightH').innerText = hLight.toFixed(2) + ' mm';
    } else if (currentSub === 'size-bar') {
        document.getElementById('sbFovH').innerText = fovH_angle.toFixed(2) + ' °';
        document.getElementById('sbFovV').innerText = fovV_angle.toFixed(2) + ' °';
        document.getElementById('sbFovD').innerText = fovD_angle.toFixed(2) + ' °';
        document.getElementById('sbSizeText').innerText = fovW.toFixed(2) + ' mm × ' + fovH.toFixed(2) + ' mm';
        document.getElementById('sbLightW').innerText = wLight.toFixed(2) + ' mm';
        document.getElementById('sbLightH').innerText = hLight.toFixed(2) + ' mm';
    } else if (currentSub === 'size-ring') {
        document.getElementById('srFovH').innerText = fovH_angle.toFixed(2) + ' °';
        document.getElementById('srFovV').innerText = fovV_angle.toFixed(2) + ' °';
        document.getElementById('srFovD').innerText = fovD_angle.toFixed(2) + ' °';
        document.getElementById('srSizeText').innerText = fovW.toFixed(2) + ' mm × ' + fovH.toFixed(2) + ' mm';
        var fovD_len = Math.hypot(fovW, fovH);
        var innerDiam = Math.max(0, fovD_len * (dCam - dLight) / dCam);
        var outerDiam = fovD_len * (dCam + dLight) / dCam;
        document.getElementById('srInnerDiam').innerText = innerDiam.toFixed(2) + ' mm';
        document.getElementById('srOuterDiam').innerText = outerDiam.toFixed(2) + ' mm';
    } else if (currentSub === 'size-custom') {
        var customCameraCount = parseFloat(document.getElementById('customCameraCount').value) || 1;
        var arrangementSelect = document.getElementById('arrangementSelect');
        var arrangementVal = arrangementSelect.value;
        var rows = 1, cols = 1;
        if (arrangementVal) {
            try {
                var obj = JSON.parse(arrangementVal);
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

        var overlap = dLight;
        var holeSize = parseFloat(document.getElementById('customHoleSize').value) || 0;

        var totalW = cols * fovW - (cols - 1) * overlap;
        var totalH = rows * fovH - (rows - 1) * overlap;
        var lightW = totalW + fovW;
        var lightH = totalH + fovH;

        document.getElementById('scFovH').innerText = fovH_angle.toFixed(2) + ' °';
        document.getElementById('scFovV').innerText = fovV_angle.toFixed(2) + ' °';
        document.getElementById('scFovD').innerText = fovD_angle.toFixed(2) + ' °';
        document.getElementById('scSizeText').innerText = fovW.toFixed(2) + ' mm × ' + fovH.toFixed(2) + ' mm';
        document.getElementById('scLightW').innerText = lightW.toFixed(2) + ' mm';
        document.getElementById('scLightH').innerText = lightH.toFixed(2) + ' mm';
        document.getElementById('scTotalSize').innerText = totalW.toFixed(2) + ' mm × ' + totalH.toFixed(2) + ' mm';

        var scCtxTotal = document.getElementById('scCanTotal').getContext('2d');
        drawTotalRect(scCtxTotal, document.getElementById('scCanTotal'), lightW, lightH, holeSize, rows, cols, overlap, fovW, fovH);
    } else if (currentSub === 'size-back') {
        document.getElementById('bfFovH').innerText = fovH_angle.toFixed(2) + ' °';
        document.getElementById('bfFovV').innerText = fovV_angle.toFixed(2) + ' °';
        document.getElementById('bfFovD').innerText = fovD_angle.toFixed(2) + ' °';
        document.getElementById('bfSizeText').innerText = fovW.toFixed(2) + ' mm × ' + fovH.toFixed(2) + ' mm';
        document.getElementById('bfLightW').innerText = wLight.toFixed(2) + ' mm';
        document.getElementById('bfLightH').innerText = hLight.toFixed(2) + ' mm';
    } else if (currentSub === 'spot-face') {
        document.getElementById('spfFovH').innerText = fovH_angle.toFixed(2) + ' °';
        document.getElementById('spfFovV').innerText = fovV_angle.toFixed(2) + ' °';
        document.getElementById('spfFovD').innerText = fovD_angle.toFixed(2) + ' °';
        document.getElementById('spfSizeText').innerText = fovW.toFixed(2) + ' mm × ' + fovH.toFixed(2) + ' mm';
        document.getElementById('spfSpotText').innerText = wLight.toFixed(2) + ' mm × ' + hLight.toFixed(2) + ' mm';
    } else if (currentSub === 'spot-ring') {
        document.getElementById('sprFovH').innerText = fovH_angle.toFixed(2) + ' °';
        document.getElementById('sprFovV').innerText = fovV_angle.toFixed(2) + ' °';
        document.getElementById('sprFovD').innerText = fovD_angle.toFixed(2) + ' °';
        document.getElementById('sprSizeText').innerText = fovW.toFixed(2) + ' mm × ' + fovH.toFixed(2) + ' mm';
        document.getElementById('sprSpotText').innerText = wLight.toFixed(2) + ' mm × ' + hLight.toFixed(2) + ' mm';
    }

    updateThreeScene(fovH_angle, fovV_angle, fovD_angle, fovW, fovH, wLight, hLight, dCam, dLight, sw, sh, dCam);
}

// --- 重置 ---
function resetAll() {
    var inputs = document.querySelectorAll('input');
    for (var i = 0; i < inputs.length; i++) {
        inputs[i].value = '';
    }

    document.getElementById('sfFovH').innerText = '--';
    document.getElementById('sfFovV').innerText = '--';
    document.getElementById('sfFovD').innerText = '--';
    document.getElementById('sfSizeText').innerText = '--';
    document.getElementById('sfLightW').innerText = '--';
    document.getElementById('sfLightH').innerText = '--';

    document.getElementById('sbFovH').innerText = '--';
    document.getElementById('sbFovV').innerText = '--';
    document.getElementById('sbFovD').innerText = '--';
    document.getElementById('sbSizeText').innerText = '--';
    document.getElementById('sbLightW').innerText = '--';
    document.getElementById('sbLightH').innerText = '--';

    document.getElementById('bfFovH').innerText = '--';
    document.getElementById('bfFovV').innerText = '--';
    document.getElementById('bfFovD').innerText = '--';
    document.getElementById('bfSizeText').innerText = '--';
    document.getElementById('bfLightW').innerText = '--';
    document.getElementById('bfLightH').innerText = '--';

    document.getElementById('spfFovH').innerText = '--';
    document.getElementById('spfFovV').innerText = '--';
    document.getElementById('spfFovD').innerText = '--';
    document.getElementById('spfSizeText').innerText = '--';
    document.getElementById('spfSpotText').innerText = '--';

    document.getElementById('sprFovH').innerText = '--';
    document.getElementById('sprFovV').innerText = '--';
    document.getElementById('sprFovD').innerText = '--';
    document.getElementById('sprSizeText').innerText = '--';
    document.getElementById('sprSpotText').innerText = '--';

    document.getElementById('spcFovH').innerText = '--';
    document.getElementById('spcFovV').innerText = '--';
    document.getElementById('spcFovD').innerText = '--';
    document.getElementById('spcSizeText').innerText = '--';
    document.getElementById('spcLightText').innerText = '--';
    document.getElementById('spcSpotText').innerText = '--';

    document.getElementById('coax-senW').value = '';
    document.getElementById('coax-senH').value = '';
    document.getElementById('coax-distCam').value = '';
    document.getElementById('coax-distLight').value = '';
    document.getElementById('coax-lightLen').value = '';
    document.getElementById('coax-lightWid').value = '';

    if (dragControls) {
        dragControls.dispose();
        dragControls = null;
    }
    if (draggableSphere) {
        threeScene.remove(draggableSphere);
        draggableSphere = null;
    }
    coaxSceneObjects = [];
    coaxDragParams = null;

    document.getElementById('size-senW').value = '';
    document.getElementById('size-senH').value = '';

    document.getElementById('spot-lightLen').value = '';
    document.getElementById('spot-lightWid').value = '';

    document.getElementById('srFovH').innerText = '--';
    document.getElementById('srFovV').innerText = '--';
    document.getElementById('srFovD').innerText = '--';
    document.getElementById('srSizeText').innerText = '--';
    document.getElementById('srInnerDiam').innerText = '--';
    document.getElementById('srOuterDiam').innerText = '--';

    document.getElementById('scFovH').innerText = '--';
    document.getElementById('scFovV').innerText = '--';
    document.getElementById('scFovD').innerText = '--';
    document.getElementById('scSizeText').innerText = '--';
    document.getElementById('scLightW').innerText = '--';
    document.getElementById('scLightH').innerText = '--';
    document.getElementById('scTotalSize').innerText = '--';

    document.getElementById('resultCard').style.display = 'none';

    var allCanvases = document.querySelectorAll('canvas');
    allCanvases.forEach(function (c) {
        var ctx = c.getContext('2d');
        ctx.clearRect(0, 0, c.width, c.height);
        c.width = 800;
        c.height = 200;
        c.style.width = '';
        c.style.height = '';
    });

    if (threeScene) {
        var toRemove = [];
        threeScene.children.forEach(function (child) {
            if (child.type !== 'AmbientLight' && child.type !== 'DirectionalLight' &&
                child.type !== 'GridHelper' && child.type !== 'AxesHelper' &&
                child.type !== 'PerspectiveCamera' && child.type !== 'Scene') {
                toRemove.push(child);
            }
        });
        toRemove.forEach(function (child) {
            threeScene.remove(child);
            if (child.geometry) child.geometry.dispose();
            if (child.material) child.material.dispose();
        });
    }
    labelObjects.forEach(function (obj) {
        if (threeScene) threeScene.remove(obj);
    });
    labelObjects = [];
}

// --- 页面初始化 ---
window.onload = function () {
    initAllSearchSelects();

    const hasCache = areaDataManager.loadFromLocal();
    if (hasCache) {
        const statusSpan = document.getElementById('dataStatus');
        if (statusSpan) {
            statusSpan.textContent = areaDataManager.getVersionInfo();
            statusSpan.style.color = '#4caf50';
        }
        areaDataManager.populateSelect(document.getElementById('modelSelect'));
        areaDataManager.populateSelect(document.getElementById('custom-modelSelect'));
        areaDataManager.populateSelect(document.getElementById('spot-modelSelect'));
        areaDataManager.populateSelect(document.getElementById('coax-modelSelect'));
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
};

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