// 内容脚本 - 注入到页面，负责渲染右侧面板

(function() {
  'use strict';

  // 避免重复注入
  if (window.pdfToWordExtensionInjected) {
    return;
  }
  window.pdfToWordExtensionInjected = true;

  let panelContainer = null;
  let isPanelVisible = false;

  // 创建面板容器
  function createPanel() {
    if (panelContainer) {
      return panelContainer;
    }

    // 创建Shadow DOM容器，避免样式污染
    const host = document.createElement('div');
    host.id = 'pdf-to-word-extension-host';
    document.body.appendChild(host);

    const shadow = host.attachShadow({ mode: 'open' });

    // 创建面板HTML
    const panel = document.createElement('div');
    panel.id = 'pdf-to-word-panel';
    panel.className = 'pdf-to-word-panel';
    panel.innerHTML = `
      <div class="panel-header">
        <h3>PDF 转 Word</h3>
        <button class="close-btn" title="关闭">×</button>
      </div>
      <div class="panel-body">
        <div class="upload-area" id="upload-area">
          <div class="upload-icon">📄</div>
          <p class="upload-text">点击或拖拽PDF文件到此处</p>
          <p class="upload-hint">支持 .pdf 格式</p>
          <input type="file" id="file-input" accept=".pdf,application/pdf" hidden>
        </div>
        <div class="file-info" id="file-info" style="display: none;">
          <div class="file-icon">📄</div>
          <div class="file-details">
            <p class="file-name" id="file-name"></p>
            <p class="file-size" id="file-size"></p>
          </div>
          <button class="remove-file" id="remove-file" title="移除">×</button>
        </div>
        <div class="options-container" id="options-container">
          <label class="option-item">
            <input type="checkbox" id="ocr-checkbox">
            <span>使用OCR识别（可识别手写/扫描内容）</span>
          </label>
          <label class="option-item" style="margin-top: 8px;">
            <input type="checkbox" id="filter-header-checkbox" checked>
            <span>过滤页眉/logo区域</span>
          </label>
        </div>
        <div class="ocr-engine-container" id="ocr-engine-container" style="display: none;">
          <label class="engine-label">OCR引擎：</label>
          <select id="ocr-engine-select" class="engine-select">
            <option value="tesseract">Tesseract.js（本地，较慢但免费）</option>
            <option value="baidu">百度AI（在线，更快需API）</option>
          </select>
          <a href="#" class="settings-link" id="settings-link">⚙️ API设置</a>
        </div>
        <div class="settings-panel" id="settings-panel" style="display: none;">
          <h4>百度AI OCR 设置</h4>
          <input type="text" id="api-key-input" placeholder="API Key" class="settings-input">
          <input type="password" id="secret-key-input" placeholder="Secret Key" class="settings-input">
          <div class="settings-buttons">
            <button class="settings-btn save" id="save-settings">保存</button>
            <button class="settings-btn cancel" id="cancel-settings">取消</button>
          </div>
          <p class="settings-hint">获取API密钥：<a href="https://ai.baidu.com/tech/ocr" target="_blank">百度AI开放平台</a></p>
        </div>
        <div class="progress-container" id="progress-container" style="display: none;">
          <div class="progress-bar">
            <div class="progress-fill" id="progress-fill"></div>
          </div>
          <p class="progress-text" id="progress-text">准备转换...</p>
        </div>
        <button class="convert-btn" id="convert-btn" disabled>
          <span class="btn-text">开始转换</span>
        </button>
        <div class="result-container" id="result-container" style="display: none;">
          <p class="result-text" id="result-text"></p>
        </div>
      </div>
    `;

    // 添加样式
    const style = document.createElement('style');
    style.textContent = getPanelStyles();

    shadow.appendChild(style);
    shadow.appendChild(panel);

    panelContainer = { host, shadow, panel };

    // 绑定事件
    bindPanelEvents(panel, shadow);

    return panelContainer;
  }

  // 面板样式
  function getPanelStyles() {
    return `
      * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
      }

      .pdf-to-word-panel {
        position: fixed;
        top: 0;
        right: -400px;
        width: 400px;
        height: 100vh;
        background: #ffffff;
        box-shadow: -2px 0 10px rgba(0, 0, 0, 0.1);
        z-index: 2147483647;
        display: flex;
        flex-direction: column;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
        transition: right 0.3s ease;
      }

      .pdf-to-word-panel.visible {
        right: 0;
      }

      .panel-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 16px 20px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
      }

      .panel-header h3 {
        font-size: 18px;
        font-weight: 600;
      }

      .close-btn {
        background: none;
        border: none;
        color: white;
        font-size: 24px;
        cursor: pointer;
        padding: 0;
        width: 32px;
        height: 32px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 4px;
        transition: background 0.2s;
      }

      .close-btn:hover {
        background: rgba(255, 255, 255, 0.2);
      }

      .panel-body {
        flex: 1;
        padding: 20px;
        overflow-y: auto;
      }

      .upload-area {
        border: 2px dashed #d0d0d0;
        border-radius: 12px;
        padding: 40px 20px;
        text-align: center;
        cursor: pointer;
        transition: all 0.3s;
        background: #f8f9fa;
      }

      .upload-area:hover {
        border-color: #667eea;
        background: #f0f2ff;
      }

      .upload-area.dragover {
        border-color: #667eea;
        background: #e8ebff;
        transform: scale(1.02);
      }

      .upload-icon {
        font-size: 48px;
        margin-bottom: 12px;
      }

      .upload-text {
        font-size: 16px;
        color: #333;
        margin-bottom: 8px;
      }

      .upload-hint {
        font-size: 13px;
        color: #999;
      }

      .file-info {
        display: flex;
        align-items: center;
        padding: 16px;
        background: #f0f2ff;
        border-radius: 8px;
        margin-bottom: 20px;
        border: 1px solid #667eea;
      }

      .file-icon {
        font-size: 32px;
        margin-right: 12px;
      }

      .file-details {
        flex: 1;
        overflow: hidden;
      }

      .file-name {
        font-size: 14px;
        font-weight: 500;
        color: #333;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }

      .file-size {
        font-size: 12px;
        color: #666;
        margin-top: 4px;
      }

      .remove-file {
        background: none;
        border: none;
        font-size: 20px;
        color: #999;
        cursor: pointer;
        padding: 4px 8px;
        border-radius: 4px;
        transition: all 0.2s;
      }

      .remove-file:hover {
        color: #ff4d4f;
        background: #fff1f0;
      }

      .progress-container {
        margin: 20px 0;
      }

      .progress-bar {
        width: 100%;
        height: 8px;
        background: #e8e8e8;
        border-radius: 4px;
        overflow: hidden;
      }

      .progress-fill {
        height: 100%;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        width: 0%;
        transition: width 0.3s ease;
      }

      .progress-text {
        text-align: center;
        font-size: 13px;
        color: #666;
        margin-top: 8px;
      }

      .convert-btn {
        width: 100%;
        padding: 14px 24px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 8px;
        font-size: 16px;
        font-weight: 500;
        cursor: pointer;
        transition: all 0.3s;
        margin-top: 10px;
      }

      .convert-btn:hover:not(:disabled) {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
      }

      .convert-btn:disabled {
        opacity: 0.6;
        cursor: not-allowed;
      }

      .convert-btn.converting {
        pointer-events: none;
      }

      .result-container {
        margin-top: 20px;
        padding: 16px;
        border-radius: 8px;
        text-align: center;
      }

      .result-container.success {
        background: #f6ffed;
        border: 1px solid #b7eb8f;
      }

      .result-container.error {
        background: #fff2f0;
        border: 1px solid #ffccc7;
      }

      .result-text {
        font-size: 14px;
        line-height: 1.6;
      }

      .result-container.success .result-text {
        color: #52c41a;
      }

      .result-container.error .result-text {
        color: #ff4d4f;
      }

      .download-link {
        display: inline-block;
        margin-top: 12px;
        padding: 10px 20px;
        background: #52c41a;
        color: white;
        text-decoration: none;
        border-radius: 6px;
        font-weight: 500;
        transition: background 0.2s;
      }

      .download-link:hover {
        background: #389e0d;
      }

      .options-container {
        margin: 16px 0;
        padding: 12px;
        background: #f5f5f5;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: space-between;
      }

      .option-item {
        display: flex;
        align-items: center;
        gap: 8px;
        font-size: 13px;
        color: #333;
        cursor: pointer;
      }

      .option-item input[type="checkbox"] {
        width: 16px;
        height: 16px;
        cursor: pointer;
      }

      .settings-link {
        font-size: 12px;
        color: #667eea;
        text-decoration: none;
        cursor: pointer;
      }

      .settings-link:hover {
        text-decoration: underline;
      }

      .settings-panel {
        margin: 12px 0;
        padding: 16px;
        background: #f8f9fa;
        border: 1px solid #e0e0e0;
        border-radius: 8px;
      }

      .settings-panel h4 {
        margin-bottom: 12px;
        font-size: 14px;
        color: #333;
      }

      .settings-input {
        width: 100%;
        padding: 10px 12px;
        margin-bottom: 10px;
        border: 1px solid #d0d0d0;
        border-radius: 6px;
        font-size: 13px;
      }

      .settings-input:focus {
        outline: none;
        border-color: #667eea;
      }

      .settings-buttons {
        display: flex;
        gap: 10px;
        margin-top: 12px;
      }

      .settings-btn {
        flex: 1;
        padding: 10px;
        border: none;
        border-radius: 6px;
        font-size: 13px;
        cursor: pointer;
        transition: background 0.2s;
      }

      .settings-btn.save {
        background: #667eea;
        color: white;
      }

      .settings-btn.save:hover {
        background: #5a6fd6;
      }

      .settings-btn.cancel {
        background: #e0e0e0;
        color: #333;
      }

      .settings-btn.cancel:hover {
        background: #d0d0d0;
      }

      .settings-hint {
        margin-top: 12px;
        font-size: 12px;
        color: #666;
      }

      .settings-hint a {
        color: #667eea;
      }

      .ocr-engine-container {
        margin: 12px 0;
        padding: 12px;
        background: #f0f7ff;
        border-radius: 8px;
        display: flex;
        align-items: center;
        gap: 10px;
        flex-wrap: wrap;
      }

      .engine-label {
        font-size: 13px;
        color: #333;
      }

      .engine-select {
        flex: 1;
        min-width: 200px;
        padding: 8px 12px;
        border: 1px solid #d0d0d0;
        border-radius: 6px;
        font-size: 13px;
        background: white;
      }

      .engine-select:focus {
        outline: none;
        border-color: #667eea;
      }
    `;
  }

  // 绑定面板事件
  function bindPanelEvents(panel, shadow) {
    const uploadArea = panel.querySelector('#upload-area');
    const fileInput = panel.querySelector('#file-input');
    const closeBtn = panel.querySelector('.close-btn');
    const convertBtn = panel.querySelector('#convert-btn');
    const removeFileBtn = panel.querySelector('#remove-file');
    const ocrCheckbox = panel.querySelector('#ocr-checkbox');
    const filterHeaderCheckbox = panel.querySelector('#filter-header-checkbox');
    const ocrEngineContainer = panel.querySelector('#ocr-engine-container');
    const ocrEngineSelect = panel.querySelector('#ocr-engine-select');
    const settingsLink = panel.querySelector('#settings-link');
    const settingsPanel = panel.querySelector('#settings-panel');
    const saveSettingsBtn = panel.querySelector('#save-settings');
    const cancelSettingsBtn = panel.querySelector('#cancel-settings');
    const apiKeyInput = panel.querySelector('#api-key-input');
    const secretKeyInput = panel.querySelector('#secret-key-input');

    let currentFile = null;
    let useOCR = false;
    let ocrEngine = 'tesseract';
    let filterHeader = true;

    // 加载保存的设置
    loadSettings();

    // OCR选项
    ocrCheckbox.addEventListener('change', (e) => {
      useOCR = e.target.checked;
      ocrEngineContainer.style.display = useOCR ? 'flex' : 'none';
    });

    // 过滤页眉选项
    filterHeaderCheckbox.addEventListener('change', (e) => {
      filterHeader = e.target.checked;
    });

    // OCR引擎选择
    ocrEngineSelect.addEventListener('change', (e) => {
      ocrEngine = e.target.value;
    });

    // 设置链接
    settingsLink.addEventListener('click', (e) => {
      e.preventDefault();
      settingsPanel.style.display = settingsPanel.style.display === 'none' ? 'block' : 'none';
    });

    // 保存设置
    saveSettingsBtn.addEventListener('click', async () => {
      const apiKey = apiKeyInput.value.trim();
      const secretKey = secretKeyInput.value.trim();

      try {
        // 使用chrome.storage保存设置
        if (typeof chrome !== 'undefined' && chrome.storage) {
          await chrome.storage.local.set({
            'baidu_ocr_api_key': apiKey,
            'baidu_ocr_secret_key': secretKey
          });
        } else {
          // 降级到localStorage
          localStorage.setItem('baidu_ocr_api_key', apiKey);
          localStorage.setItem('baidu_ocr_secret_key', secretKey);
        }

        settingsPanel.style.display = 'none';
        showResult('设置已保存！', 'success');
        setTimeout(() => {
          const resultContainer = panel.querySelector('#result-container');
          resultContainer.style.display = 'none';
        }, 1500);
      } catch (error) {
        console.error('保存设置失败:', error);
        showResult('保存设置失败: ' + error.message, 'error');
      }
    });

    // 取消设置
    cancelSettingsBtn.addEventListener('click', () => {
      settingsPanel.style.display = 'none';
      loadSettings();
    });

    // 加载设置
    async function loadSettings() {
      try {
        let savedApiKey = '';
        let savedSecretKey = '';

        // 使用chrome.storage获取设置
        if (typeof chrome !== 'undefined' && chrome.storage) {
          const result = await chrome.storage.local.get(['baidu_ocr_api_key', 'baidu_ocr_secret_key']);
          savedApiKey = result.baidu_ocr_api_key || '';
          savedSecretKey = result.baidu_ocr_secret_key || '';
        }

        // 如果chrome.storage没有，尝试localStorage
        if (!savedApiKey) {
          savedApiKey = localStorage.getItem('baidu_ocr_api_key') || '';
        }
        if (!savedSecretKey) {
          savedSecretKey = localStorage.getItem('baidu_ocr_secret_key') || '';
        }

        if (savedApiKey) apiKeyInput.value = savedApiKey;
        if (savedSecretKey) secretKeyInput.value = savedSecretKey;
      } catch (error) {
        console.error('加载设置失败:', error);
      }
    }

    // 关闭按钮
    closeBtn.addEventListener('click', () => {
      togglePanel(false);
    });

    // 点击上传区域
    uploadArea.addEventListener('click', () => {
      fileInput.click();
    });

    // 文件选择
    fileInput.addEventListener('change', (e) => {
      if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
      }
    });

    // 拖拽上传
    uploadArea.addEventListener('dragover', (e) => {
      e.preventDefault();
      uploadArea.classList.add('dragover');
    });

    uploadArea.addEventListener('dragleave', () => {
      uploadArea.classList.remove('dragover');
    });

    uploadArea.addEventListener('drop', (e) => {
      e.preventDefault();
      uploadArea.classList.remove('dragover');
      const files = e.dataTransfer.files;
      if (files.length > 0 && files[0].type === 'application/pdf') {
        handleFile(files[0]);
      }
    });

    // 移除文件
    removeFileBtn.addEventListener('click', () => {
      currentFile = null;
      fileInput.value = '';
      updateFileDisplay();
    });

    // 转换按钮
    convertBtn.addEventListener('click', () => {
      if (currentFile) {
        startConversion(currentFile);
      }
    });

    // 处理文件
    function handleFile(file) {
      if (file.type !== 'application/pdf') {
        showResult('请选择PDF文件', 'error');
        return;
      }
      currentFile = file;
      updateFileDisplay();
    }

    // 更新文件显示
    function updateFileDisplay() {
      const uploadArea = panel.querySelector('#upload-area');
      const fileInfo = panel.querySelector('#file-info');
      const fileName = panel.querySelector('#file-name');
      const fileSize = panel.querySelector('#file-size');
      const convertBtn = panel.querySelector('#convert-btn');
      const resultContainer = panel.querySelector('#result-container');

      if (currentFile) {
        uploadArea.style.display = 'none';
        fileInfo.style.display = 'flex';
        fileName.textContent = currentFile.name;
        fileSize.textContent = formatFileSize(currentFile.size);
        convertBtn.disabled = false;
        resultContainer.style.display = 'none';
      } else {
        uploadArea.style.display = 'block';
        fileInfo.style.display = 'none';
        convertBtn.disabled = true;
      }
    }

    // 开始转换
    async function startConversion(file) {
      const convertBtn = panel.querySelector('#convert-btn');
      const progressContainer = panel.querySelector('#progress-container');
      const progressFill = panel.querySelector('#progress-fill');
      const progressText = panel.querySelector('#progress-text');

      convertBtn.disabled = true;
      convertBtn.classList.add('converting');
      progressContainer.style.display = 'block';

      try {
        updateProgress(10, '正在初始化...');

        // 确保处理函数可用
        if (typeof window.parsePDF !== 'function') {
          throw new Error('PDF处理模块未加载，请刷新页面重试');
        }

        updateProgress(15, '正在读取PDF文件...');

        // 读取PDF文件
        const arrayBuffer = await file.arrayBuffer();

        updateProgress(30, '正在解析PDF内容...');

        // 解析PDF
        const pdfData = await window.parsePDF(arrayBuffer, (progress) => {
          updateProgress(30 + progress * 0.4, useOCR ? `正在使用${ocrEngine === 'tesseract' ? 'Tesseract.js' : '百度AI'}进行OCR识别...` : '正在提取内容...');
        }, useOCR, ocrEngine, filterHeader);

        updateProgress(70, '正在生成Word文档...');

        // 生成Word
        const docxBlob = await window.generateDocx(pdfData, (progress) => {
          updateProgress(70 + progress * 0.2, '正在生成文档...');
        });

        updateProgress(100, '转换完成！');

        // 调试：显示提取的文本统计
        let totalChars = 0;
        pdfData.pages.forEach(page => {
          page.text.forEach(item => {
            totalChars += item.text.length;
          });
        });
        console.log('PDF提取统计:', {
          页数: pdfData.pages.length,
          总字符数: totalChars,
          第一页文本: pdfData.pages[0]?.text.slice(0, 20).map(t => t.text).join(' ')
        });

        // 下载文件
        const fileName = file.name.replace('.pdf', '.docx');
        downloadFile(docxBlob, fileName);

        showResult(`转换成功！<br>提取了 ${totalChars} 个字符<br><a href="#" class="download-link" id="download-again">重新下载</a>`, 'success');

        // 重新下载链接
        panel.querySelector('#download-again').addEventListener('click', (e) => {
          e.preventDefault();
          downloadFile(docxBlob, fileName);
        });

      } catch (error) {
        console.error('转换失败:', error);
        showResult('转换失败: ' + error.message, 'error');
      } finally {
        convertBtn.disabled = false;
        convertBtn.classList.remove('converting');

        // 终止Tesseract worker以释放内存
        if (useOCR && ocrEngine === 'tesseract' && window.terminateTesseract) {
          window.terminateTesseract().catch(console.error);
        }
      }

      function updateProgress(percent, text) {
        progressFill.style.width = percent + '%';
        progressText.textContent = text;
      }
    }

    // 显示结果
    function showResult(message, type) {
      const resultContainer = panel.querySelector('#result-container');
      const resultText = panel.querySelector('#result-text');
      const progressContainer = panel.querySelector('#progress-container');

      progressContainer.style.display = 'none';
      resultContainer.style.display = 'block';
      resultContainer.className = 'result-container ' + type;
      resultText.innerHTML = message;
    }

    // 格式化文件大小
    function formatFileSize(bytes) {
      if (bytes === 0) return '0 Bytes';
      const k = 1024;
      const sizes = ['Bytes', 'KB', 'MB', 'GB'];
      const i = Math.floor(Math.log(bytes) / Math.log(k));
      return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    // 下载文件
    function downloadFile(blob, fileName) {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }
  }

  // 显示/隐藏面板
  function togglePanel(show) {
    const panel = createPanel();
    isPanelVisible = show !== undefined ? show : !isPanelVisible;

    if (isPanelVisible) {
      panel.panel.classList.add('visible');
    } else {
      panel.panel.classList.remove('visible');
    }
  }

  // 监听来自后台脚本的消息
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'togglePanel') {
      togglePanel();
      sendResponse({ success: true, visible: isPanelVisible });
    }
    return true;
  });

  // 点击面板外部关闭 - 已禁用，用户需要手动点击关闭按钮
  // document.addEventListener('click', (e) => {
  //   if (!isPanelVisible || !panelContainer) return;
  //
  //   const host = panelContainer.host;
  //   const panel = panelContainer.panel;
  //
  //   if (!host.contains(e.target) && !panel.contains(e.target)) {
  //     const rect = panel.getBoundingClientRect();
  //     if (e.clientX < rect.left) {
  //       togglePanel(false);
  //     }
  //   }
  // });

  console.log('PDF转Word插件内容脚本已加载');
})();
