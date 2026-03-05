// PDF 解析和 Word 生成模块

// 设置 PDF.js worker 路径
if (typeof pdfjsLib !== 'undefined') {
  pdfjsLib.GlobalWorkerOptions.workerSrc = chrome.runtime.getURL('pdf.worker.min.js');
}

// Tesseract.js 配置
let tesseractWorker = null;

// 初始化 Tesseract.js worker
async function initTesseract() {
  if (tesseractWorker) return tesseractWorker;

  if (typeof Tesseract === 'undefined') {
    throw new Error('Tesseract.js 库未加载');
  }

  try {
    console.log('正在初始化Tesseract.js...');

    // 配置worker路径
    const workerPath = chrome.runtime.getURL('tesseract.worker.min.js');
    console.log('Worker路径:', workerPath);

    // 创建worker，使用多个CDN备选加载语言包
    // 备选CDN: jsdelivr, unpkg, github raw
    const langPaths = [
      'https://tessdata.projectnaptha.com/4.0.0_best',
      'https://cdn.jsdelivr.net/npm/tesseract.js-core@4.0.0/tessdata',
      'https://unpkg.com/tesseract.js-core@4.0.0/tessdata'
    ];

    let lastError = null;
    for (const langPath of langPaths) {
      try {
        console.log(`尝试从 ${langPath} 加载语言包...`);

        tesseractWorker = await Tesseract.createWorker('chi_sim+eng', 1, {
          workerPath: workerPath,
          langPath: langPath,
          logger: (m) => {
            console.log('Tesseract日志:', m);
          },
          errorHandler: (err) => {
            console.error('Tesseract Worker错误:', err);
          },
        });

        console.log('Tesseract.js初始化成功');
        return tesseractWorker;
      } catch (e) {
        console.warn(`从 ${langPath} 加载失败:`, e.message);
        lastError = e;
        // 继续尝试下一个CDN
      }
    }

    // 所有CDN都失败
    throw new Error(`OCR引擎初始化失败，无法加载语言包。${lastError?.message || ''}`);
  } catch (error) {
    console.error('Tesseract初始化失败:', error);
    console.error('错误详情:', error.stack);
    throw new Error('OCR引擎初始化失败: ' + error.message);
  }
}

// 终止 Tesseract worker
async function terminateTesseract() {
  if (tesseractWorker) {
    await tesseractWorker.terminate();
    tesseractWorker = null;
  }
}

// 将PDF页面转换为图片
async function pdfPageToImage(page, scale = 2) {
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  const context = canvas.getContext('2d');

  canvas.width = viewport.width;
  canvas.height = viewport.height;

  await page.render({
    canvasContext: context,
    viewport: viewport
  }).promise;

  return canvas.toDataURL('image/png');
}

// 图片预处理 - 提高对比度和清晰度，有助于OCR识别手写体
function preprocessImageForOCR(imageDataUrl) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');

      canvas.width = img.width;
      canvas.height = img.height;

      // 绘制原图
      ctx.drawImage(img, 0, 0);

      // 获取图像数据
      const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
      const data = imageData.data;

      // 图像增强：增加对比度、灰度化、二值化
      const contrast = 1.5; // 对比度因子
      const threshold = 128; // 二值化阈值

      for (let i = 0; i < data.length; i += 4) {
        // 灰度化
        const gray = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];

        // 增加对比度
        const contrasted = ((gray - 128) * contrast) + 128;

        // 限制在0-255范围内
        const clamped = Math.max(0, Math.min(255, contrasted));

        // 简单的二值化（可选，根据需要调整）
        // const final = clamped > threshold ? 255 : 0;
        const final = clamped;

        data[i] = final;     // R
        data[i + 1] = final; // G
        data[i + 2] = final; // B
        // Alpha保持不变
      }

      // 将处理后的数据写回canvas
      ctx.putImageData(imageData, 0, 0);

      // 返回处理后的图片
      resolve(canvas.toDataURL('image/png'));
    };

    img.onerror = reject;
    img.src = imageDataUrl;
  });
}

// 过滤掉不需要识别的内容（如logo、页眉页脚等）
function filterOCRResults(items, pageHeight) {
  // 定义要过滤的关键词（常见logo文字）
  const skipPatterns = [
    /^YIDATEC$/i,
    /^YIDATEC\s*\d+$/i,
    /^%?\s*F?\d+$/i,  // 类似 "F100018015" 的乱码
    /^\d{3,}$/,       // 纯数字且长度大于等于3
    /^[@#%^&*]+$/,    // 纯特殊字符
  ];

  return items.filter(item => {
    const text = item.text?.trim() || '';

    // 过滤空文本
    if (!text) return false;

    // 过滤匹配skipPatterns的文本
    for (const pattern of skipPatterns) {
      if (pattern.test(text)) {
        console.log('过滤掉logo/乱码:', text);
        return false;
      }
    }

    // 过滤顶部10%区域的文字（通常是页眉/logo）
    if (item.y > pageHeight * 0.9) {
      console.log('过滤掉页眉文字:', text);
      return false;
    }

    return true;
  });
}

// 使用 Tesseract.js 进行OCR识别
async function recognizeWithTesseract(imageDataUrl, pageHeight = 1000, filterHeader = true) {
  const worker = await initTesseract();

  try {
    console.log('开始Tesseract.js识别...');
    const result = await worker.recognize(imageDataUrl);
    console.log('Tesseract.js识别结果:', result);

    // 解析识别结果 - 优先使用lines（整行）来保持格式
    const words = [];

    // 优先从 lines 获取（整行文本，保持段落格式）
    if (result.data.lines && result.data.lines.length > 0) {
      console.log('从lines提取:', result.data.lines.length);
      for (const line of result.data.lines) {
        if (line.text && line.text.trim()) {
          words.push({
            text: line.text.trim(),
            fontSize: 12,
            fontName: '',
            x: line.bbox?.x0 || 0,
            y: line.bbox?.y0 || 0,
            width: (line.bbox?.x1 || 0) - (line.bbox?.x0 || 0),
            height: (line.bbox?.y1 || 0) - (line.bbox?.y0 || 0),
            hasEOL: true, // 标记为行尾，这样段落处理会正确处理
            dir: 'ltr',
            isOCR: true
          });
        }
      }
    }
    // 尝试从 words 获取（如果lines不可用）
    else if (result.data.words && result.data.words.length > 0) {
      console.log('从words提取:', result.data.words.length);
      // 将words按行分组
      const lineGroups = [];
      const tolerance = 10;

      for (const word of result.data.words) {
        if (!word.text || !word.text.trim()) continue;

        let found = false;
        for (const group of lineGroups) {
          if (Math.abs(word.bbox?.y0 - group.y) <= tolerance) {
            group.words.push(word);
            found = true;
            break;
          }
        }
        if (!found) {
          lineGroups.push({ y: word.bbox?.y0 || 0, words: [word] });
        }
      }

      // 每行合并为一个文本项
      for (const group of lineGroups) {
        group.words.sort((a, b) => (a.bbox?.x0 || 0) - (b.bbox?.x0 || 0));
        const lineText = group.words.map(w => w.text).join('');
        words.push({
          text: lineText,
          fontSize: 12,
          fontName: '',
          x: group.words[0]?.bbox?.x0 || 0,
          y: group.y,
          width: (group.words[group.words.length - 1]?.bbox?.x1 || 0) - (group.words[0]?.bbox?.x0 || 0),
          height: group.words[0]?.bbox?.y1 - group.words[0]?.bbox?.y0 || 20,
          hasEOL: true,
          dir: 'ltr',
          isOCR: true
        });
      }
    }
    // 最后尝试从 text 获取（整页文本，按行分割）
    else if (result.data.text) {
      console.log('从text提取');
      const lines = result.data.text.split('\n').filter(l => l.trim());
      for (let i = 0; i < lines.length; i++) {
        words.push({
          text: lines[i].trim(),
          fontSize: 12,
          fontName: '',
          x: 0,
          y: i * 30, // 估算行高
          width: lines[i].length * 12,
          height: 30,
          hasEOL: true,
          dir: 'ltr',
          isOCR: true
        });
      }
    }

    console.log('Tesseract.js提取到文字数（过滤前）:', words.length);

    // 过滤掉logo和乱码
    const filteredWords = filterHeader ? filterOCRResults(words, pageHeight) : words;
    console.log('Tesseract.js提取到文字数（过滤后）:', filteredWords.length);

    return filteredWords;
  } catch (error) {
    console.error('OCR识别失败:', error);
    throw error;
  }
}

// 百度AI OCR 配置
const BAIDU_OCR_CONFIG = {
  API_KEY: '',
  SECRET_KEY: '',
  OCR_API_URL: 'https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic',
  // 使用高精度识别API，更适合手写体
  ACCURATE_API_URL: 'https://aip.baidubce.com/rest/2.0/ocr/v1/accurate'
};

// 从chrome.storage获取百度API配置
async function getBaiduCredentials() {
  return new Promise((resolve) => {
    if (typeof chrome !== 'undefined' && chrome.storage) {
      chrome.storage.local.get(['baidu_ocr_api_key', 'baidu_ocr_secret_key'], (result) => {
        resolve({
          apiKey: result.baidu_ocr_api_key || BAIDU_OCR_CONFIG.API_KEY,
          secretKey: result.baidu_ocr_secret_key || BAIDU_OCR_CONFIG.SECRET_KEY
        });
      });
    } else {
      // 降级到localStorage
      resolve({
        apiKey: localStorage.getItem('baidu_ocr_api_key') || BAIDU_OCR_CONFIG.API_KEY,
        secretKey: localStorage.getItem('baidu_ocr_secret_key') || BAIDU_OCR_CONFIG.SECRET_KEY
      });
    }
  });
}

// 获取百度AI Access Token
async function getBaiduAccessToken() {
  const { apiKey, secretKey } = await getBaiduCredentials();

  if (!apiKey || !secretKey) {
    throw new Error('请先配置百度AI OCR API密钥');
  }

  const url = `https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=${apiKey}&client_secret=${secretKey}`;

  try {
    const response = await fetch(url, { method: 'POST' });
    const data = await response.json();

    if (data.error) {
      throw new Error(`获取AccessToken失败: ${data.error_description}`);
    }

    return data.access_token;
  } catch (error) {
    throw new Error(`百度AI认证失败: ${error.message}`);
  }
}

// 使用百度OCR
async function recognizeWithBaidu(imageDataUrl, pageHeight = 1000, filterHeader = true) {
  console.log('开始百度OCR识别...');
  const accessToken = await getBaiduAccessToken();
  console.log('获取到百度AccessToken');

  const base64Data = imageDataUrl.split(',')[1];
  console.log('图片base64长度:', base64Data.length);

  // 使用高精度识别API（更适合手写体）
  const url = `${BAIDU_OCR_CONFIG.ACCURATE_API_URL}?access_token=${accessToken}`;

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      body: `image=${encodeURIComponent(base64Data)}&recognize_granularity=big&detect_direction=true&vertexes_location=true`
    });

    const data = await response.json();
    console.log('百度OCR返回结果:', data);

    if (data.error_code) {
      throw new Error(`OCR识别失败: ${data.error_msg}`);
    }

    const ocrResults = data.words_result || [];
    console.log('百度OCR识别到文字数:', ocrResults.length);

    // 将结果按行分组（百度OCR返回的结果需要按Y坐标排序）
    const results = ocrResults.map(result => ({
      text: result.words,
      fontSize: 12,
      fontName: '',
      x: result.location?.left || 0,
      y: result.location?.top || 0,
      width: result.location?.width || 0,
      height: result.location?.height || 0,
      hasEOL: true, // 标记为行尾
      dir: 'ltr',
      isOCR: true
    }));

    // 按Y坐标排序
    results.sort((a, b) => a.y - b.y);

    // 过滤掉logo和乱码
    const filteredResults = filterHeader ? filterOCRResults(results, pageHeight) : results;
    console.log('百度OCR识别到文字数（过滤后）:', filteredResults.length);

    return filteredResults;
  } catch (error) {
    console.error('百度OCR请求失败:', error);
    throw error;
  }
}

// 解析 PDF 文件
async function parsePDF(arrayBuffer, onProgress, useOCR = false, ocrEngine = 'tesseract', filterHeader = true) {
  try {
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const numPages = pdf.numPages;
    const pages = [];

    // 初始化OCR引擎（如果需要）
    if (useOCR && ocrEngine === 'tesseract') {
      try {
        onProgress?.(0.05);
        await initTesseract();
      } catch (e) {
        console.warn('Tesseract初始化失败，将使用普通文本提取:', e.message);
        useOCR = false;
      }
    }

    for (let i = 1; i <= numPages; i++) {
      const page = await pdf.getPage(i);
      const viewport = page.getViewport({ scale: 1.0 });

      // 获取文本内容
      const textContent = await page.getTextContent();
      const items = textContent.items;

      // 提取普通文本
      const textItems = items.map(item => ({
        text: item.str,
        fontSize: item.fontSize || 12,
        fontName: item.fontName || '',
        x: item.transform ? item.transform[4] : 0,
        y: item.transform ? item.transform[5] : 0,
        width: item.width || 0,
        height: item.height || 0,
        hasEOL: item.hasEOL || false,
        dir: item.dir || 'ltr',
        isOCR: false
      }));

      const allItems = [...textItems];

      // 如果需要OCR
      if (useOCR) {
        try {
          onProgress?.(0.3 + (i - 1) / numPages * 0.2);

          console.log(`第${i}页: 正在转换为图片用于OCR...`);
          // 将页面转为图片 - 使用更高分辨率提高OCR准确度
          const imageDataUrl = await pdfPageToImage(page, 3); // scale=3 提高清晰度

          console.log(`第${i}页: 正在预处理图片...`);
          // 图片预处理（对比度增强、灰度化）
          const processedImageUrl = await preprocessImageForOCR(imageDataUrl);

          onProgress?.(0.4 + (i - 1) / numPages * 0.1);

          let ocrItems = [];

          if (ocrEngine === 'tesseract') {
            console.log(`第${i}页: 使用Tesseract.js进行OCR识别...`);
            ocrItems = await recognizeWithTesseract(processedImageUrl, viewport.height, filterHeader);
          } else if (ocrEngine === 'baidu') {
            console.log(`第${i}页: 使用百度AI OCR进行识别...`);
            // 百度OCR对预处理后的图片效果可能更好
            ocrItems = await recognizeWithBaidu(processedImageUrl, viewport.height, filterHeader);
          }

          console.log(`第${i}页: 普通文本${textItems.length}个, OCR识别${ocrItems.length}个`);

          // 如果OCR识别到文字，优先使用OCR结果
          // 因为OCR可以识别手写/扫描内容，而普通PDF提取可能提取不到
          if (ocrItems.length > 0) {
            allItems.length = 0;
            allItems.push(...ocrItems);
            console.log(`第${i}页: 已使用OCR结果替换普通文本，OCR文本样例:`,
              ocrItems.slice(0, 3).map(item => item.text).join(', '));
          } else {
            console.warn(`第${i}页: OCR未识别到任何文字，保留普通PDF文本`);
          }

        } catch (e) {
          console.error(`第${i}页OCR识别失败，使用普通文本:`, e);
          // 显示更详细的错误信息
          if (e.message.includes('API')) {
            console.error('API配置错误，请检查百度API密钥');
          }
        }
      }

      // 按位置排序
      allItems.sort((a, b) => {
        if (Math.abs(a.y - b.y) < 5) {
          return a.x - b.x;
        }
        return b.y - a.y;
      });

      const pageData = {
        pageNum: i,
        text: allItems,
        viewport: viewport
      };

      pages.push(pageData);

      if (onProgress) {
        onProgress(0.5 + i / numPages * 0.5);
      }
    }

    return {
      numPages,
      pages
    };

  } catch (error) {
    console.error('PDF解析错误:', error);
    throw new Error('PDF文件解析失败: ' + error.message);
  }
}

// 将 PDF 数据转换为 Word 文档
async function generateDocx(pdfData, onProgress) {
  try {
    const children = [];

    for (let pageIndex = 0; pageIndex < pdfData.pages.length; pageIndex++) {
      const page = pdfData.pages[pageIndex];
      const textItems = page.text;

      if (textItems.length === 0) continue;

      // 按 Y 坐标分组（同一行的文本）
      const lineGroups = groupByLines(textItems);

      // 处理每一行
      for (const line of lineGroups) {
        const paragraph = createParagraph(line, page.viewport);
        if (paragraph) {
          children.push(paragraph);
        }
      }

      // 页面分隔
      if (pageIndex < pdfData.pages.length - 1) {
        children.push(new docx.Paragraph({
          text: '',
          pageBreakBefore: true
        }));
      }

      if (onProgress) {
        onProgress((pageIndex + 1) / pdfData.pages.length);
      }
    }

    // 创建文档
    const doc = new docx.Document({
      sections: [{
        properties: {
          page: {
            margin: {
              top: 1440,
              right: 1440,
              bottom: 1440,
              left: 1440
            }
          }
        },
        children: children
      }]
    });

    const blob = await docx.Packer.toBlob(doc);
    return blob;

  } catch (error) {
    console.error('Word生成错误:', error);
    throw new Error('Word文档生成失败: ' + error.message);
  }
}

// 按行分组文本
function groupByLines(textItems) {
  const tolerance = 4;
  const lines = [];

  const validItems = textItems.filter(item => item.text && item.text.trim().length > 0);

  if (validItems.length === 0) return lines;

  const sorted = [...validItems].sort((a, b) => b.y - a.y);
  const yGroups = [];

  for (const item of sorted) {
    let found = false;
    for (const group of yGroups) {
      if (Math.abs(item.y - group.y) <= tolerance) {
        group.items.push(item);
        found = true;
        break;
      }
    }
    if (!found) {
      yGroups.push({ y: item.y, items: [item] });
    }
  }

  for (const group of yGroups) {
    group.items.sort((a, b) => a.x - b.x);
    lines.push(group);
  }

  lines.sort((a, b) => b.y - a.y);

  return lines;
}

// 创建段落
function createParagraph(line, viewport) {
  const items = line.items;
  if (items.length === 0) return null;

  // 检查是否是OCR结果（OCR结果通常已经按行组织好）
  const isOCRLine = items.some(item => item.isOCR);

  let textContent = '';

  if (isOCRLine && items.length === 1) {
    // 单个OCR行结果，直接使用
    textContent = items[0].text.trim();
  } else {
    // 多个items，需要合并
    const textParts = [];
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      textParts.push(item.text);
      if (i < items.length - 1) {
        const nextItem = items[i + 1];
        const gap = nextItem.x - (item.x + item.width);
        if (item.hasEOL || gap > 5) {
          textParts.push(' ');
        }
      }
    }
    textContent = textParts.join('').trim();
  }

  if (!textContent) return null;

  const avgFontSize = items.reduce((sum, item) => sum + (item.fontSize || 12), 0) / items.length;
  const isBold = items.some(item =>
    /Bold|Heavy|Black|粗|加粗/.test(item.fontName || '') ||
    (item.fontSize && item.fontSize >= 14)
  );

  const alignment = detectAlignment(items, viewport);

  // 创建TextRun
  let runs = [];

  if (isOCRLine && items.length === 1) {
    // OCR单行结果，简化处理
    runs = [new docx.TextRun({
      text: textContent,
      size: 24, // 12pt
      font: '宋体'
    })];
  } else {
    // 普通PDF文本，保留格式
    runs = items.map(item => {
      const itemIsBold = /Bold|Heavy|Black|粗|加粗/.test(item.fontName || '');
      const itemIsItalic = /Italic|Oblique|斜/.test(item.fontName || '');
      const itemIsUnderline = /Underline|下划线/.test(item.fontName || '');

      return new docx.TextRun({
        text: item.text,
        bold: itemIsBold,
        italics: itemIsItalic,
        underline: itemIsUnderline ? {} : undefined,
        size: Math.round((item.fontSize || 12) * 2),
        font: getFontName(item.fontName)
      });
    });
  }

  let isTitle = false;
  let isHeading = false;

  // OCR结果的avgFontSize可能不准确，所以降低阈值
  if (!isOCRLine && avgFontSize >= 16) {
    isTitle = true;
  } else if (!isOCRLine && avgFontSize >= 14) {
    isHeading = true;
  }

  // 居中且较短的文本可能是标题
  if (textContent.length < 40 && alignment === docx.AlignmentType.CENTER) {
    isTitle = true;
  }

  // 根据文本内容判断是否为标题/章节
  if (/^[第][一二三四五六七八九十\d]+[章节条]/.test(textContent) ||
      /^[一二三四五六七八九十]+[、.．\s]/.test(textContent) ||
      /^\d+[.．]\s*/.test(textContent) ||
      /^（[一二三四五六七八九十]）/.test(textContent) ||
      /^[\d一二三四五六七八九十]+[、.．\s]*$/.test(textContent)) {
    isHeading = true;
  }

  // OCR结果通常不需要特殊标题处理，保持简单格式
  if (isOCRLine) {
    isTitle = false;
    isHeading = false;
  }

  const paragraphConfig = {
    children: runs,
    alignment: alignment,
    spacing: {
      before: isTitle ? 240 : (isHeading ? 200 : 100),
      after: isTitle ? 200 : (isHeading ? 160 : 100),
      line: 360
    }
  };

  if (isTitle) {
    paragraphConfig.heading = docx.HeadingLevel.HEADING_1;
  } else if (isHeading) {
    paragraphConfig.heading = docx.HeadingLevel.HEADING_2;
  }

  return new docx.Paragraph(paragraphConfig);
}

// 检测对齐方式
function detectAlignment(items, viewport) {
  if (!items.length || !viewport) return docx.AlignmentType.LEFT;

  const firstItem = items[0];
  const lastItem = items[items.length - 1];

  const lineStart = firstItem.x;
  const lineEnd = lastItem.x + lastItem.width;
  const lineWidth = lineEnd - lineStart;
  const pageWidth = viewport.width;

  const leftMargin = lineStart;
  const rightMargin = pageWidth - lineEnd;

  const textContent = items.map(i => i.text).join('');

  if (Math.abs(leftMargin - rightMargin) < 80) {
    return docx.AlignmentType.CENTER;
  }

  if (rightMargin < 40 && leftMargin > 80) {
    return docx.AlignmentType.RIGHT;
  }

  if (lineWidth > pageWidth * 0.6 && textContent.length > 15) {
    return docx.AlignmentType.JUSTIFIED;
  }

  return docx.AlignmentType.LEFT;
}

// 获取字体名称
function getFontName(fontName) {
  if (!fontName) return '宋体';

  const name = fontName.toLowerCase();

  if (name.includes('simsun') || name.includes('song') || name.includes('宋')) return '宋体';
  if (name.includes('simhei') || name.includes('hei') || name.includes('黑')) return '黑体';
  if (name.includes('kai') || name.includes('楷')) return '楷体';
  if (name.includes('fang') || name.includes('仿')) return '仿宋';
  if (name.includes('microsoft yahei') || name.includes('微软雅黑')) return '微软雅黑';
  if (name.includes('arial')) return 'Arial';
  if (name.includes('times')) return 'Times New Roman';
  if (name.includes('courier')) return 'Courier New';
  if (name.includes('helvetica')) return 'Helvetica';

  return '宋体';
}

// 导出函数
if (typeof window !== 'undefined') {
  window.parsePDF = parsePDF;
  window.generateDocx = generateDocx;
  window.terminateTesseract = terminateTesseract;
  window.BAIDU_OCR_CONFIG = BAIDU_OCR_CONFIG;
}
