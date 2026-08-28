// Excel数据管理类
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
                    var k = keys[j].replace(/\s/g, ''); // 去除所有空格
                    var pClean = p.replace(/\s/g, '');
                    if (k.indexOf(pClean) !== -1 || k.toLowerCase().indexOf(pClean.toLowerCase()) !== -1) {
                        return keys[j]; // 返回原始列名
                    }
                }
            }
            return '';
        }

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
                threadSpec: findKey(['螺纹规格', '螺纹', 'thread']),
                // ---- 新增远心镜头字段 ----
                lensType: findKey(['镜头类型', '类型']),
                magnification: findKey(['放大倍率', '倍率', 'mag']),
                fixedWD: findKey(['固定工作距离', '固定WD', 'fixedWD'])
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

    // 在 class ExcelDataManager { ... } 内部添加

    parseWorkbook(workbook) {
        var sheetNames = workbook.SheetNames;
        this.cameraData = [];
        this.lensData = [];
        for (var i = 0; i < sheetNames.length; i++) {
            var name = sheetNames[i];
            var sheet = workbook.Sheets[name];
            var json = XLSX.utils.sheet_to_json(sheet);
            if (!json || json.length === 0) continue;
            if (name === 'Sheet1') {
                this.cameraData = json;
                this.cameraColumnMap = this.detectColumns(json, 'camera');
            } else if (name === 'Sheet2') {
                this.lensData = json;
                this.lensColumnMap = this.detectColumns(json, 'lens');
            }
        }
        this.buildCameraIndex();
        this.buildLensIndex();
        // 兼容旧代码
        this.data = this.cameraData;
        this.columnMap = this.cameraColumnMap;
        this.modelNames = this.cameraModelNames;
        this.modelRowMap = this.cameraModelRowMap;
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



    // ---------- 获取数据目录路径 ----------
    getDataDir() {
        // 如果是 nw.js 环境
        if (typeof nw !== 'undefined' && nw.App && nw.App.dataPath) {
            return nw.App.dataPath;
        }
        // 如果是 Electron 环境（备选）
        if (typeof process !== 'undefined' && process.env && process.env.APPDATA) {
            const path = require('path');
            return path.join(process.env.APPDATA, 'SheetJSApp'); // 自定义应用名
        }
        // 开发环境 fallback
        return process.cwd();
    }

    // ---------- 获取数据文件完整路径 ----------
    getDataFilePath() {
        const path = require('path');
        return path.join(this.getDataDir(), 'data.xlsx');
    }

    // ---------- 从文件加载（使用数据目录） ----------
    loadFromFile() {
        const fs = require('fs');
        const filePath = this.getDataFilePath();
        if (!fs.existsSync(filePath)) {
            console.warn('数据文件不存在:', filePath);
            return false;
        }
        const buffer = fs.readFileSync(filePath);
        const workbook = XLSX.read(buffer, { type: 'buffer' });
        this.parseWorkbook(workbook);
        this.isLoaded = true;
        this.fileName = 'data.xlsx (本地)';
        this.loadTime = new Date().toLocaleString();
        this.buildCameraIndex();
        this.buildLensIndex();
        this.saveToLocal(); // 同时更新 localStorage 备份
        return true;
    }

    // ---------- 保存到文件（写入数据目录） ----------
    saveToFile() {
        if (!this.isLoaded) return false;
        const fs = require('fs');
        const path = require('path');
        const dir = this.getDataDir();
        // 确保目录存在
        if (!fs.existsSync(dir)) {
            fs.mkdirSync(dir, { recursive: true });
        }
        const filePath = path.join(dir, 'data.xlsx');
        // 创建 workbook
        const wb = XLSX.utils.book_new();
        const ws1 = XLSX.utils.json_to_sheet(this.cameraData);
        const ws2 = XLSX.utils.json_to_sheet(this.lensData);
        XLSX.utils.book_append_sheet(wb, ws1, 'Sheet1');
        XLSX.utils.book_append_sheet(wb, ws2, 'Sheet2');
        const out = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
        fs.writeFileSync(filePath, out);
        return true;
    }

    // ---------- 修改 loadExcel：上传后自动保存到数据目录 ----------
    loadExcel(file) {
        const self = this;
        return new Promise(function (resolve, reject) {
            const reader = new FileReader();
            reader.onload = function (ev) {
                try {
                    const data = new Uint8Array(ev.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    self.parseWorkbook(workbook);
                    self.isLoaded = true;
                    self.fileName = file.name + ' (已保存本地)';
                    self.loadTime = new Date().toLocaleString();
                    self.buildCameraIndex();
                    self.buildLensIndex();
                    self.saveToLocal();
                    // 保存到数据目录
                    if (typeof require !== 'undefined') {
                        try {
                            const fs = require('fs');
                            const path = require('path');
                            const dir = self.getDataDir();
                            if (!fs.existsSync(dir)) {
                                fs.mkdirSync(dir, { recursive: true });
                            }
                            const targetPath = path.join(dir, 'data.xlsx');
                            fs.writeFileSync(targetPath, Buffer.from(data));
                            console.log('数据已保存到用户目录:', targetPath);
                        } catch (e) {
                            console.warn('无法写入文件:', e.message);
                        }
                    }
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










}