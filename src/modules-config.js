window.MODULES_CONFIG = {
  groups: {
    size: { label: '尺寸计算', icon: '📏' },
    spot: { label: '光斑计算', icon: '💡' }
  },

  modules: {
    // ========== 尺寸计算组 ==========
    'size-face': {
      group: 'size',
      label: '平面同轴',
      icon: '🔲',
      // 不再使用 inputPrefix，id 直接写完整 DOM id
      inputs: [
        { id: 'size-distCam', label: '相机工作距离（mm）', type: 'number' },
        { id: 'size-distLight', label: '光源工作距离（mm）', type: 'number' }
      ],
      resultMap: {
        fovH: 'sfFovH',
        fovV: 'sfFovV',
        fovD: 'sfFovD',
        sizeText: 'sfSizeText',
        lightW: 'sfLightW',
        lightH: 'sfLightH'
      },
      threeContainer: 'three-sf',
      calcType: 'standard'
    },

    'size-bar': {
      group: 'size',
      label: '条光',
      icon: '➖',
      inputs: [
        { id: 'size-distCam', label: '相机工作距离（mm）', type: 'number' },
        { id: 'size-distLight', label: '光源工作距离（mm）', type: 'number' }
      ],
      resultMap: {
        fovH: 'sbFovH',
        fovV: 'sbFovV',
        fovD: 'sbFovD',
        sizeText: 'sbSizeText',
        lightW: 'sbLightW',
        lightH: 'sbLightH'
      },
      threeContainer: 'three-sb',
      calcType: 'standard'
    },

    'size-ring': {
      group: 'size',
      label: '环光',
      icon: '⭕',
      inputs: [
        { id: 'size-distCam', label: '相机工作距离（mm）', type: 'number' },
        { id: 'size-distLight', label: '光源工作距离（mm）', type: 'number' }
      ],
      resultMap: {
        fovH: 'srFovH',
        fovV: 'srFovV',
        fovD: 'srFovD',
        sizeText: 'srSizeText',
        innerDiam: 'srInnerDiam',
        outerDiam: 'srOuterDiam'
      },
      threeContainer: 'three-sr',
      calcType: 'ring'
    },

    'size-custom': {
      group: 'size',
      label: '定制面光',
      icon: '🔧',
      inputs: [
        { id: 'custom-distCam', label: '相机工作距离（mm）', type: 'number' },
        { id: 'custom-distLight', label: '视野重叠区域（mm）', type: 'number' },
        { id: 'customCameraCount', label: '相机个数', type: 'number', default: 1 },
        // 下拉框直接匹配 DOM id
        { id: 'arrangementSelect', label: '排列分布', type: 'select' },
        { id: 'customHoleSize', label: '开孔尺寸（mm）', type: 'number', default: 0 }
      ],
      resultMap: {
        fovH: 'scFovH',
        fovV: 'scFovV',
        fovD: 'scFovD',
        sizeText: 'scSizeText',
        lightW: 'scLightW',
        lightH: 'scLightH',
        totalSize: 'scTotalSize'
      },
      threeContainer: 'three-sc',
      calcType: 'custom'
    },

    'size-back': {
      group: 'size',
      label: '背光',
      icon: '💡',
      inputs: [
        { id: 'size-distCam', label: '相机工作距离（mm）', type: 'number' },
        { id: 'size-distLight', label: '光源工作距离（mm）', type: 'number' }
      ],
      resultMap: {
        fovH: 'bfFovH',
        fovV: 'bfFovV',
        fovD: 'bfFovD',
        sizeText: 'bfSizeText',
        lightW: 'bfLightW',
        lightH: 'bfLightH'
      },
      threeContainer: 'three-bf',
      calcType: 'standard'
    },



    'size-coax': {
      group: 'size',
      label: '同轴光',
      icon: '🔲',
      inputs: [
        { id: 'size-distCam', label: '相机工作距离（mm）', type: 'number' },
        { id: 'size-distLight', label: '光源工作距离（mm）', type: 'number' }
      ],
      resultMap: {
        fovH: 'coaxFovH',
        fovV: 'coaxFovV',
        fovD: 'coaxFovD',
        sizeText: 'coaxSizeText',
        lightW: 'coaxLightW',
        lightH: 'coaxLightH'
      },
      threeContainer: 'three-coax',
      calcType: 'standard'
    },

    'size-dome': {
      group: 'size',
      label: '圆顶光',
      icon: '🔮',
      inputs: [
        { id: 'size-distCam', label: '相机工作距离（mm）', type: 'number' },
        { id: 'size-distLight', label: '光源工作距离（mm）', type: 'number' }
      ],
      resultMap: {
        fovH: 'domeFovH',
        fovV: 'domeFovV',
        fovD: 'domeFovD',
        sizeText: 'domeSizeText',
        lightW: 'domeLightW',
        lightH: 'domeLightH'
      },
      threeContainer: 'three-dome',
      calcType: 'standard'
    },

    // ========== 光斑计算组 ==========
    'spot-face': {
      group: 'spot',
      label: '平面同轴',
      icon: '📐',
      inputs: [
        { id: 'spot-distCam', label: '相机工作距离（mm）', type: 'number' },
        { id: 'spot-distLight', label: '光源工作距离（mm）', type: 'number' }
      ],
      resultMap: {
        fovH: 'spfFovH',
        fovV: 'spfFovV',
        fovD: 'spfFovD',
        sizeText: 'spfSizeText',
        spotText: 'spfSpotText'
      },
      threeContainer: 'three-spf',
      calcType: 'standard'
    },

    'spot-coax': {
      group: 'spot',
      label: '同轴光',
      icon: '💡',
      inputs: [
        { id: 'coax-distCam', label: '相机工作距离（mm）', type: 'number' },
        { id: 'coax-distLight', label: '光源工作距离（mm）', type: 'number' },
        { id: 'coax-lightLen', label: '发光区长边（mm）', type: 'number' },
        { id: 'coax-lightWid', label: '发光区短边（mm）', type: 'number' }
      ],
      resultMap: {
        fovH: 'spcFovH',
        fovV: 'spcFovV',
        fovD: 'spcFovD',
        sizeText: 'spcSizeText',
        lightText: 'spcLightText',
        spotText: 'spcSpotText'
      },
      threeContainer: 'three-spc',
      calcType: 'coax'
    },

    'spot-ring': {
      group: 'spot',
      label: '环光',
      icon: '⭕',
      inputs: [
        { id: 'spot-distCam', label: '相机工作距离（mm）', type: 'number' },
        { id: 'spot-distLight', label: '光源工作距离（mm）', type: 'number' }
      ],
      resultMap: {
        fovH: 'sprFovH',
        fovV: 'sprFovV',
        fovD: 'sprFovD',
        sizeText: 'sprSizeText',
        spotText: 'sprSpotText'
      },
      threeContainer: 'three-spr',
      calcType: 'standard'
    }
  },

  // ---- 工具函数 ----
  getModuleConfig: function (moduleId) {
    return this.modules[moduleId] || null;
  },
  getModulesByGroup: function (groupId) {
    const result = [];
    for (const [id, cfg] of Object.entries(this.modules)) {
      if (cfg.group === groupId) result.push(id);
    }
    return result;
  }
};