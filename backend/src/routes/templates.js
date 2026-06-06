const express = require('express');
const router = express.Router();
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const templateController = require('../controllers/templateController');
const { WORD_EXTENSIONS } = require('../services/document-conversion');

// Ensure upload directory exists
const UPLOADS_DIR = path.join(__dirname, '..', '..', 'uploads');
const TEMPLATES_DIR = path.join(UPLOADS_DIR, 'templates');
if (!fs.existsSync(TEMPLATES_DIR)) {
  fs.mkdirSync(TEMPLATES_DIR, { recursive: true });
}

function normalizeVietnameseFileName(fileName) {
  if (!fileName) return '';
  try {
    const normalized = Buffer.from(fileName, 'latin1').toString('utf8');
    const replacementCount = (normalized.match(/\uFFFD/g) || []).length;
    const originalReplacementCount = (fileName.match(/\uFFFD/g) || []).length;
    if (replacementCount > originalReplacementCount) {
      return fileName;
    }
    return normalized;
  } catch {
    return fileName;
  }
}

// Multer storage
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, TEMPLATES_DIR);
  },
  filename: (req, file, cb) => {
    const normalizedName = normalizeVietnameseFileName(file.originalname);
    cb(null, `${Date.now()}-${normalizedName}`);
  }
});

const upload = multer({
  storage: storage,
  limits: {
    files: 10,
    fileSize: 1 * 1024 * 1024
  },
  fileFilter: (req, file, cb) => {
    const normalizedName = normalizeVietnameseFileName(file.originalname);
    const ext = path.extname(normalizedName).toLowerCase();
    if (WORD_EXTENSIONS.has(ext)) {
      cb(null, true);
    } else {
      cb(new Error('Hệ thống chỉ chấp nhận file Word định dạng .doc hoặc .docx'));
    }
  }
});

// Routes
router.get('/', templateController.getTemplates);

router.post('/', (req, res, next) => {
  upload.fields([
    { name: 'templateFile', maxCount: 10 },
    { name: 'templateFiles', maxCount: 10 }
  ])(req, res, (err) => {
    if (err) {
      return res.status(400).json({ error: err.message });
    }
    next();
  });
}, templateController.uploadTemplates);

router.get('/:id/form', templateController.getTemplateForm);
router.get('/:id/tables', templateController.getTemplateTables);
router.post('/:id/inject-table', templateController.injectTemplateTable);
router.put('/:id/fields', templateController.updateTemplateFields);
router.get('/:id/links', templateController.getTemplateLinks);
router.post('/:id/link', templateController.linkTemplate);
router.post('/:id/unlink', templateController.unlinkTemplate);
router.get('/:id/parent-fields', templateController.getParentFields);
router.post('/:templateId/fields/:fieldId/restore', templateController.restoreTemplateField);
router.get('/:id/download-original', templateController.downloadOriginalTemplate);
router.delete('/:id', templateController.deleteTemplate);
router.put('/:id', templateController.updateTemplateName);
router.put('/:id/category', templateController.updateTemplateCategory);
router.put('/:id/repeated', templateController.toggleTemplateRepeated);
router.put('/:id/status', templateController.toggleTemplateStatus);
router.post('/:id/duplicate', templateController.duplicateTemplate);
router.get('/:id/export', templateController.exportTemplateFields);
router.post('/:id/import', templateController.importTemplateFields);

module.exports = router;
