const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');
const util = require('util');
const execFilePromise = util.promisify(execFile);

const WORD_EXTENSIONS = new Set(['.doc', '.docx']);
const DOCX_EXTENSION = '.docx';
const DOC_CONVERTER_URL = (process.env.DOC_CONVERTER_URL || 'http://127.0.0.1:5051').replace(/\/$/, '');
const DOC_CONVERTER_TIMEOUT_MS = Number(process.env.DOC_CONVERTER_TIMEOUT_MS || 20000);
const NETWORK_ERROR_CODES = new Set([
  'ECONNREFUSED',
  'ECONNRESET',
  'EHOSTUNREACH',
  'ENETUNREACH',
  'ETIMEDOUT'
]);

async function convertDocWithLibreOffice(inputPath, outputDir) {
  let command = 'libreoffice';
  let fallbackCommand = 'soffice';

  if (process.platform === 'win32') {
    const commonPaths = [
      'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
      'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe'
    ];
    for (const p of commonPaths) {
      if (fs.existsSync(p)) {
        command = p;
        fallbackCommand = null;
        break;
      }
    }
  }

  try {
    await execFilePromise(command, [
      '--headless',
      '--convert-to',
      'docx',
      inputPath,
      '--outdir',
      outputDir
    ]);
  } catch (error) {
    if (fallbackCommand) {
      try {
        await execFilePromise(fallbackCommand, [
          '--headless',
          '--convert-to',
          'docx',
          inputPath,
          '--outdir',
          outputDir
        ]);
      } catch (innerError) {
        throw new Error(`LibreOffice/Soffice conversion failed: ${innerError.message}`);
      }
    } else {
      throw new Error(`LibreOffice conversion failed: ${error.message}`);
    }
  }
}

async function convertDocWithService(inputPath, outputPath) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), DOC_CONVERTER_TIMEOUT_MS);

  try {
    const fileBuffer = await fs.promises.readFile(inputPath);
    const form = new FormData();
    form.append(
      'file',
      new Blob([fileBuffer], { type: 'application/msword' }),
      path.basename(inputPath)
    );

    const response = await fetch(`${DOC_CONVERTER_URL}/api/convert-doc`, {
      method: 'POST',
      body: form,
      signal: controller.signal
    });

    if (!response.ok) {
      let detail = '';
      try {
        const data = await response.json();
        detail = data.error || '';
      } catch {
        detail = await response.text();
      }
      throw new Error(detail || `Microservice chuyển đổi trả về HTTP ${response.status}.`);
    }

    const arrayBuffer = await response.arrayBuffer();
    await fs.promises.writeFile(outputPath, Buffer.from(arrayBuffer));
  } catch (error) {
    if (error.name === 'AbortError') {
      throw new Error('Microservice chuyển đổi không phản hồi trong thời gian cho phép.');
    }

    const errorCode = error.cause?.code || error.code;
    if (NETWORK_ERROR_CODES.has(errorCode) || error.message === 'fetch failed') {
      throw new Error(
        `Không thể kết nối tới microservice chuyển đổi .NET tại ${DOC_CONVERTER_URL}. ` +
        'Hãy kiểm tra service .NET đã chạy chưa và endpoint /health có phản hồi hay không.'
      );
    }

    throw new Error(`Không thể chuyển file .doc sang .docx bằng microservice .NET. Chi tiết: ${error.message}`);
  } finally {
    clearTimeout(timeout);
  }
}

async function ensureDocxForTemplate(inputPath, options = {}) {
  const ext = path.extname(inputPath).toLowerCase();
  if (ext === DOCX_EXTENSION) {
    return {
      outputPath: inputPath,
      cleanup: () => {}
    };
  }

  if (!WORD_EXTENSIONS.has(ext)) {
    throw new Error('Hệ thống chỉ chấp nhận file Word định dạng .doc hoặc .docx');
  }

  const templateId = options.templateId || `tmp-${Date.now()}`;
  const outputDir = options.outputDir || path.dirname(inputPath);
  const conversionDir = path.join(outputDir, `_conversion-${templateId}`);
  const outputPath = path.join(
    conversionDir,
    `${path.basename(inputPath, ext)}${DOCX_EXTENSION}`
  );

  await fs.promises.mkdir(conversionDir, { recursive: true });

  try {
    // Try converting using LibreOffice first since it's locally installed and very fast
    try {
      await convertDocWithLibreOffice(inputPath, conversionDir);
    } catch (libreError) {
      console.warn('LibreOffice conversion failed or not found, falling back to .NET microservice:', libreError.message);
      // Fallback to .NET microservice
      await convertDocWithService(inputPath, outputPath);
    }

    // Double-check the file exists at the expected path
    if (!fs.existsSync(outputPath)) {
      throw new Error('Không tìm thấy file .docx sau khi chuyển đổi.');
    }

    return {
      outputPath,
      cleanup: () => {
        if (fs.existsSync(conversionDir)) {
          fs.rmSync(conversionDir, { recursive: true, force: true });
        }
      }
    };
  } catch (error) {
    if (fs.existsSync(conversionDir)) {
      fs.rmSync(conversionDir, { recursive: true, force: true });
    }
    throw error;
  }
}

module.exports = {
  DOC_CONVERTER_URL,
  WORD_EXTENSIONS,
  ensureDocxForTemplate
};
