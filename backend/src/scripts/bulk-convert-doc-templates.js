const fs = require('fs');
const path = require('path');
const { ensureDocxForTemplate } = require('../services/document-conversion');

function collectDocFiles(rootDir) {
  const entries = fs.readdirSync(rootDir, { withFileTypes: true });
  const files = [];

  for (const entry of entries) {
    const fullPath = path.join(rootDir, entry.name);
    if (entry.isDirectory()) {
      files.push(...collectDocFiles(fullPath));
      continue;
    }

    if (path.extname(entry.name).toLowerCase() === '.doc') {
      files.push(fullPath);
    }
  }

  return files;
}

async function main() {
  const sourceDir = process.argv[2];
  const outputDir = process.argv[3];

  if (!sourceDir || !outputDir) {
    console.error('Cách dùng: node src/scripts/bulk-convert-doc-templates.js <sourceDir> <outputDir>');
    process.exit(1);
  }

  const absoluteSourceDir = path.resolve(sourceDir);
  const absoluteOutputDir = path.resolve(outputDir);

  if (!fs.existsSync(absoluteSourceDir)) {
    console.error(`Không tìm thấy thư mục nguồn: ${absoluteSourceDir}`);
    process.exit(1);
  }

  const docFiles = collectDocFiles(absoluteSourceDir);
  if (docFiles.length === 0) {
    console.log('Không tìm thấy file .doc nào để chuyển đổi.');
    return;
  }

  fs.mkdirSync(absoluteOutputDir, { recursive: true });

  let convertedCount = 0;
  const failedFiles = [];

  for (const docFile of docFiles) {
    const relativePath = path.relative(absoluteSourceDir, docFile);
    const destinationDir = path.join(absoluteOutputDir, path.dirname(relativePath));
    const destinationFile = path.join(
      destinationDir,
      `${path.basename(docFile, '.doc')}.docx`
    );

    fs.mkdirSync(destinationDir, { recursive: true });

    try {
      const conversion = await ensureDocxForTemplate(docFile, {
        templateId: `bulk-${convertedCount + 1}`,
        outputDir: destinationDir
      });

      fs.copyFileSync(conversion.outputPath, destinationFile);
      conversion.cleanup();
      convertedCount += 1;
      console.log(`[OK] ${relativePath} -> ${path.relative(absoluteOutputDir, destinationFile)}`);
    } catch (error) {
      failedFiles.push({ file: relativePath, error: error.message });
      console.error(`[FAIL] ${relativePath}: ${error.message}`);
    }
  }

  console.log(`Hoàn tất: ${convertedCount}/${docFiles.length} file đã được chuyển sang .docx.`);

  if (failedFiles.length > 0) {
    process.exitCode = 1;
  }
}

main().catch((error) => {
  console.error('Lỗi batch convert:', error);
  process.exit(1);
});
