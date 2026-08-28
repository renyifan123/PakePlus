// ============================================================
//  three-scene.js  -  Three.js 场景渲染模块
//  依赖：THREE（全局），currentSub（来自 script.js）
// ============================================================

// --- 全局镜头图片与模型加载器 ---
const lensImage = new Image();
lensImage.src = "lens.png";
lensImage.onload = () => console.log("镜头图片加载完成");
lensImage.onerror = () => console.warn("未找到lens.png，将使用默认方块");
var gltfLoader = new THREE.GLTFLoader();
var cameraModel = null;
var lensModel = null;
var externalModels = [];
var lightModel = null;

// --- Three.js 场景核心变量 ---
var threeScene, threeCamera, threeRenderer, threeControls;
var labelRenderer;
var labelObjects = [];
var activeContainerId = 'three-sf';
var modelsLoaded = false;

// --- 同轴光拖拽相关 ---
var dragControls = null;
var draggableSphere = null;
var coaxSceneObjects = [];
var coaxDragParams = null;
var coaxDistLabel = null;
var currentModuleModel = null; // 当前加载的模块模型

// 在 updateThreeScene 函数内部顶部添加
var MODEL_PATH_MAP = {
    'size-face': 'models/size-face.glb',
    'size-bar': 'models/size-bar.glb',
    'size-ring': 'models/size-ring.glb',
    'size-coax': 'models/size-coax.glb',
    'size-dome': 'models/size-dome.glb',
    'spot-face': 'models/spot-face.glb',
    'spot-coax': 'models/spot-coax.glb',
    'spot-ring': 'models/spot-ring.glb'
};


// 每个模块的变换配置（位置、旋转、缩放系数）
const MODEL_TRANSFORM_MAP = {
    'size-face': {
        offset: [0, 0, 0],// 相对于投影面中心的偏移 (x, y, z)
        rotation: [0, 0, 0], // 固定旋转 (弧度)
        scaleFactorX: 1.1,
        scaleFactorY: 1.1,
        opacity: 0.1,      // 新增
        axisMap: { width: 'x', height: 'z' }   // 新增
    },
    'size-bar': {
        offset: [0, 0, 0],
        rotation: [0, 0, 0],
        scaleFactorX: 0.6,
        scaleFactorY: 0.6,
        opacity: 1.0,     // 新增
        axisMap: { width: 'x', height: 'z' }   // 新增
    },
    'size-ring': {
        offset: [0, 0, 0],
        rotation: [0, Math.PI / 2, 0],
        scaleFactorX: 1.2,   // 环光使用外径，也可调整
        scaleFactorY: 1.2,
        opacity: 0.1,      // 新增
        axisMap: { width: 'x', height: 'y' }   // 新增
    },
    'size-coax': {
        offset: [0, 0, 0],
        rotation: [Math.PI / 2, 0, 0],
        scaleFactorX: 0.6,
        scaleFactorY: 0.6,
        opacity: 1.0,     // 新增
        axisMap: { width: 'x', height: 'z' }   // 新增
    },
    'size-dome': {
        offset: [0, 0, 0],
        rotation: [0, 0, 0],
        scaleFactorX: 1,
        scaleFactorY: 1,
        opacity: 1.0,      // 新增
        axisMap: { width: 'x', height: 'z' }   // 新增
    },
    'spot-face': {
        offset: [0, 0, 0],
        rotation: [0, 0, 0],
        scaleFactorX: 0.6,
        scaleFactorY: 0.6,
        opacity: 1.0,     // 新增
        axisMap: { width: 'x', height: 'z' }   // 新增
    },
    'spot-coax': {
        offset: [0, 0, 0],
        rotation: [0, 0, 0],
        scaleFactorX: 0.6,
        scaleFactorY: 0.6,
        opacity: 1.0,     // 新增
        axisMap: { width: 'x', height: 'z' }   // 新增
    },
    'spot-ring': {
        offset: [0, 0, 0],
        rotation: [0, 0, 0],
        scaleFactorX: 0.6,
        scaleFactorY: 0.6,
        opacity: 1.0,     // 新增
        axisMap: { width: 'x', height: 'z' }   // 新增
    }
};

function applyModelScale(model, wLight, hLight, tf, axisMap) {
    const box = new THREE.Box3().setFromObject(model);
    const size = box.getSize(new THREE.Vector3());
    const axis = axisMap || { width: 'x', height: 'z' };

    // 从 size 中取出对应轴的值
    const sizeW = size[axis.width];
    const sizeH = size[axis.height];
    if (sizeW === 0 || sizeH === 0) return;

    const scaleX = (wLight * tf.scaleFactorX) / sizeW;
    const scaleY = (hLight * tf.scaleFactorY) / sizeH;
    // 第三个轴保持 1（或根据模型需要调整）
    model.scale.set(scaleX, scaleY, 1);
}


// ============================================================
//  辅助函数：设置模型透明度
// ============================================================
function setModelOpacity(model, opacity) {
    console.log('设置透明度:', opacity);
    if (!model) return;
    model.traverse(function (child) {
        if (child.isMesh) {
            const materials = Array.isArray(child.material) ? child.material : [child.material];
            materials.forEach(mat => {
                mat.transparent = true;
                mat.opacity = opacity;
                mat.depthWrite = false;    // 关键：防止深度遮挡
                mat.side = THREE.DoubleSide; // 双面渲染，便于观察
                mat.needsUpdate = true;
            });
        }
    });
}

// ============================================================
//  辅助函数：清空场景动态物体
// ============================================================
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

// ============================================================
//  加载外部模型（相机、镜头）
// ============================================================
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
//  切换 3D 容器
// ============================================================
function switchThreeContainer(subId) {
    const config = window.MODULES_CONFIG?.getModuleConfig(subId);
    const newContainerId = config?.threeContainer || 'three-sf';


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

// ============================================================
//  更新渲染器尺寸
// ============================================================
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

// ============================================================
//  初始化 Three.js 场景
// ============================================================
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

// ============================================================
//  动画循环
// ============================================================
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

// ============================================================
//  主 3D 场景更新（面光、条光、环光、背光等）
// ============================================================
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
    var isCircular = (currentSub === 'size-ring' || currentSub === 'spot-ring' || currentSub === 'size-dome');
    var isRingOnly = (currentSub === 'size-ring' || currentSub === 'spot-ring');
    var isBackLight = (currentSub === 'size-back');

    var origin = new THREE.Vector3(0, 0, 0);
    for (var i = 0; i < 4; i++) addLine(origin, pts[i], 0x88ccff, false);
    addRect(pts, 0xffaa00, false);

    if (isCircular) {
        var fovD_len = Math.hypot(wReal, hReal);
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
        if (isRingOnly && innerRadius > 0.1) {
            addCircle(centerProj, innerRadius, 0xff3300, true);
        }

        // 在 addCircle 完成后，添加模型（环光专用）
        var modelPath = MODEL_PATH_MAP[currentSub];
        if (modelPath) {
            gltfLoader.load(modelPath, function (gltf) {
                const model = gltf.scene;
                const tf = MODEL_TRANSFORM_MAP[currentSub] || { offset: [0, 0, 0], rotation: [0, 0, 0], scaleFactorX: 1, scaleFactorY: 1, opacity: 0.1 };
                const axisMap = tf.axisMap || { width: 'x', height: 'z' };

                // 设置位置、旋转
                const basePos = new THREE.Vector3(0, 0, projZ);
                model.position.copy(basePos).add(new THREE.Vector3(tf.offset[0], tf.offset[1], tf.offset[2]));
                model.rotation.set(tf.rotation[0], tf.rotation[1], tf.rotation[2]);

                // ---- 环光等比缩放：使用外径作为目标直径 ----
                const box = new THREE.Box3().setFromObject(model);
                const size = box.getSize(new THREE.Vector3());
                // 根据 axisMap 提取模型原始宽度和高度（环光理想情况下两者相等）
                const sizeW = size[axisMap.width];
                const sizeH = size[axisMap.height];
                // 取较大值作为模型特征尺寸（避免因模型不对称导致变形）
                const modelSize = Math.max(sizeW, sizeH);
                if (modelSize > 0) {
                    // 统一缩放系数：使用 scaleFactorX 作为微调（可同时调整 scaleFactorX 和 scaleFactorY，但这里只用一个）
                    const uniformScale = (outerDiam * tf.scaleFactorX) / modelSize;
                    // 三轴等比缩放
                    model.scale.set(uniformScale, uniformScale, uniformScale);
                }

                if (tf.opacity !== undefined && tf.opacity < 1) {
                    setModelOpacity(model, tf.opacity);
                }
                threeScene.add(model);
            });
        }

    } else {
        var ptsLight = [
            new THREE.Vector3(-halfLW, -halfLH, zLight),
            new THREE.Vector3(halfLW, -halfLH, zLight),
            new THREE.Vector3(halfLW, halfLH, zLight),
            new THREE.Vector3(-halfLW, halfLH, zLight)
        ];
        // ---- 背光模式：
        if (isBackLight) {

            // 1. 视野面到光源面的连线（保留）
            for (var i = 0; i < 4; i++) {
                addLine(pts[i], ptsLight[i], 0xff8800, true);
            }
            if (lightModel) {
                threeScene.remove(lightModel);
                lightModel = null;
            }

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
            // 保留投影面矩形
            addRect(projPts, 0xff3300, true);
            // 改为从视野面四个角连线到投影面四个角
            for (var i = 0; i < 4; i++) {
                addLine(pts[i], projPts[i], 0xff3300, true);
            }
            // ---- 普通光源模型 ----
            // ---- 加载模型 ----
            var modelPath = MODEL_PATH_MAP[currentSub];
            if (modelPath) {
                gltfLoader.load(modelPath, function (gltf) {
                    const model = gltf.scene;
                    const tf = MODEL_TRANSFORM_MAP[currentSub] || { offset: [0, 0, 0], rotation: [0, 0, 0], scaleFactorX: 0.6, scaleFactorY: 0.6, opacity: 1.0 };
                    const axisMap = tf.axisMap || { width: 'x', height: 'z' };

                    // 设置位置、旋转
                    const basePos = new THREE.Vector3(0, 0, projZ);
                    model.position.copy(basePos).add(new THREE.Vector3(tf.offset[0], tf.offset[1], tf.offset[2]));
                    model.rotation.set(tf.rotation[0], tf.rotation[1], tf.rotation[2]);

                    // 应用缩放（使用 axisMap）
                    applyModelScale(model, wLight, hLight, tf, axisMap);

                    if (tf.opacity !== undefined && tf.opacity < 1) {
                        setModelOpacity(model, tf.opacity);
                    }
                    threeScene.add(model);
                });
            }
        }
    }
    /*
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
    */

    var innerDiam, outerDiam;
    if (isCircular) {
        var fovD_len = Math.hypot(wReal, hReal);
        innerDiam = fovD_len * (dCam - dLight) / dCam;
        outerDiam = fovD_len * (dCam + dLight) / dCam;
        if (innerDiam < 0) innerDiam = 0;
    }


    /*
        var projLabelText;
        if (!isBackLight) {
            if (isCircular) {
                var labelTitle = (currentSub === 'size-dome') ? '⬇️ 圆顶投影' : '⬇️ 环形投影';
                var innerPart = '';
                if (isRingOnly) {
                    innerPart = '<span class="label-value light">内径：' + innerDiam.toFixed(2) + ' mm</span> &nbsp;|&nbsp; ';
                }
                projLabelText = '<span class="label-title">' + labelTitle + '</span><br>' +
                    innerPart +
                    '<span class="label-value light">外径：' + outerDiam.toFixed(2) + ' mm</span>';
            } else {
                projLabelText = '<span class="label-title">⬇️ 光源投影</span><br>' +
                    '<span class="label-value light">' + wLight.toFixed(1) + ' × ' + hLight.toFixed(1) + ' mm</span>';
            }
            createLabel(projLabelText, new THREE.Vector3(-wLight, 0, projZ));
        }
    */
    // ---- 根据模式设置视口中心 ----
    var centerZ, offsetY, offsetZ;
    if (isBackLight) {
        // 背光：聚焦于视野面 (z = dCam) 和光源面 (z = zLight) 之间
        centerZ = (dCam + dLight - 80) / 2;          // zLight = dCam + dLight
        offsetY = dCam + dLight + 200;  // 适当降低高度，使光源模型可见
        offsetZ = 0;
    } else {
        // 非背光：聚焦于视野面 (z = dCam) 和投影面 (z = projZ) 之间
        centerZ = dCam / 2;
        offsetY = dCam + dLight + 200;          // 保持原有高度
        offsetZ = 0;
    }
    threeControls.target.set(0, 0, centerZ);
    threeCamera.position.set(0, centerZ + offsetY + 100, centerZ + offsetZ + 50);
    threeCamera.lookAt(0, 0, centerZ);
    threeControls.update();

    console.log('3D 场景更新完成');
}

// ============================================================
//  同轴光 3D 场景绘制
// ============================================================
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
    /*
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
    */
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
    /*
        if (hasSensor) {
            var fieldText = '<span class="label-title">📷 视野</span><br>' +
                '<span class="label-value field">' + wReal.toFixed(1) + ' × ' + hReal.toFixed(1) + ' mm</span>';
            createLabel(fieldText, new THREE.Vector3(-wReal, halfConeH + offsetField + 50, zWork));
        }
    */
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
    /*
        var spotText = '<span class="label-title">💡 有效光斑</span><br>' +
            '<span class="label-value field">' + spotW.toFixed(1) + ' × ' + spotH.toFixed(1) + ' mm</span>';
        createLabel(spotText, new THREE.Vector3(-spotW - 50, -halfSH + -offsetSpot, (zWork - 30)));
    
        var lightText = '<span class="label-title">💡 发光区</span><br>' +
            '<span class="label-value light">' + lightLen.toFixed(1) + ' × ' + lightWid.toFixed(1) + ' mm</span>';
        createLabel(lightText, new THREE.Vector3(-lightLen, 0, zLight));
    */
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

// ============================================================
//  同轴光拖拽功能
// ============================================================
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

// ============================================================
//  视角切换（正视图 / 俯视图 / 左视图）
// ============================================================
function setView(view) {
    if (!threeControls || !threeCamera || !threeScene) return;

    var target = threeControls.target.clone();
    var distance = threeCamera.position.distanceTo(target);
    if (distance < 1) distance = 500; // 防零

    var dir = new THREE.Vector3();
    switch (view) {
        case 'front': dir.set(0, 1, 0); break;   // 从Z正方向看
        case 'top': dir.set(0, 0, -1); break;   // 从Y正方向看
        case 'left': dir.set(1, 0, 0); break;  // 从X负方向看
        default: return;
    }

    var newPos = target.clone().add(dir.multiplyScalar(distance));
    threeCamera.position.copy(newPos);
    threeCamera.lookAt(target);
    threeControls.update();
}

// 暴露到全局，供HTML按钮调用
window.setView = setView;