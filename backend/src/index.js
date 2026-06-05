const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const { v4: uuidv4 } = require('uuid');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const archiver = require('archiver');
require('dotenv').config();

const { pool, testConnection } = require('./db');
const { scanPlaceholders, mergeDocument, mergeDocumentToBuffer, injectPlaceholders, restorePlaceholder, scanTables, injectTablePlaceholders } = require('./parser');
const { WORD_EXTENSIONS, ensureDocxForTemplate } = require('./services/document-conversion');

const DEFAULT_TEMPLATE_CATEGORIES = [
  {
    name: 'Hợp đồng',
    children: ['Chuyển nhượng', 'Tặng cho', 'Thế chấp', 'Ủy quyền']
  },
  {
    name: 'Đất đai',
    children: ['Sang tên', 'Biến động', 'Đăng ký']
  },
  {
    name: 'Thuế',
    children: ['TNCN', 'Lệ phí trước bạ', 'Phi nông nghiệp']
  },
  {
    name: 'Biểu mẫu nội bộ',
    children: []
  }
];

const app = express();
const PORT = process.env.PORT || 5000;

// Middleware
app.use(cors());
app.use(express.json());

// Đảm bảo các thư mục upload tồn tại
const UPLOADS_DIR = path.join(__dirname, '..', 'uploads');
const TEMPLATES_DIR = path.join(UPLOADS_DIR, 'templates');
const OUTPUTS_DIR = path.join(UPLOADS_DIR, 'outputs');

[UPLOADS_DIR, TEMPLATES_DIR, OUTPUTS_DIR].forEach(dir => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

// Cấu hình lưu trữ file upload bằng Multer
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, TEMPLATES_DIR);
  },
  filename: (req, file, cb) => {
    const normalizedName = normalizeVietnameseFileName(file.originalname);
    // Lưu tạm bằng tên gốc kèm timestamp trước khi đổi tên chính thức theo template_id
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

function cleanupFileIfExists(filePath) {
  if (filePath && fs.existsSync(filePath)) {
    fs.unlinkSync(filePath);
  }
}

function normalizeVietnameseFileName(fileName) {
  if (!fileName) {
    return '';
  }

  try {
    const normalized = Buffer.from(fileName, 'latin1').toString('utf8');
    const replacementCount = (normalized.match(/�/g) || []).length;
    const originalReplacementCount = (fileName.match(/�/g) || []).length;

    if (replacementCount > originalReplacementCount) {
      return fileName;
    }

    return normalized;
  } catch {
    return fileName;
  }
}

async function createTemplateFromUploadedFile(file, options = {}) {
  const normalizedOriginalName = normalizeVietnameseFileName(file.originalname);
  const originalExt = path.extname(normalizedOriginalName).toLowerCase();
  const templateId = uuidv4();
  const templateName = options.templateName || path.basename(normalizedOriginalName, originalExt);
  const categoryId = options.categoryId || null;
  const officeId = options.officeId || DEFAULT_OFFICE_ID;

  let sourceDocxPath = file.path;
  let cleanupConvertedFile = () => {};

  try {
    if (originalExt === '.doc') {
      const conversion = await ensureDocxForTemplate(file.path, {
        templateId,
        outputDir: TEMPLATES_DIR
      });
      sourceDocxPath = conversion.outputPath;
      cleanupConvertedFile = conversion.cleanup;
      cleanupFileIfExists(file.path);
    }

    const finalFileName = `${templateId}.docx`;
    const finalPath = path.join(TEMPLATES_DIR, finalFileName);
    fs.renameSync(sourceDocxPath, finalPath);
    cleanupConvertedFile();

    const relativeFilePath = path.relative(path.join(__dirname, '..'), finalPath);

    let variables = [];
    try {
      variables = scanPlaceholders(finalPath);
    } catch (parseErr) {
      cleanupFileIfExists(finalPath);
      throw parseErr;
    }

    const conn = await pool.getConnection();
    try {
      await conn.beginTransaction();

      await conn.query(
        'INSERT INTO templates (id, office_id, category_id, name, file_path, status) VALUES (?, ?, ?, ?, ?, ?)',
        [templateId, officeId, categoryId, templateName, relativeFilePath, 'draft']
      );

      for (let i = 0; i < variables.length; i++) {
        const key = variables[i];
        const fieldId = uuidv4();
        const defaultLabel = key
          .replace(/_/g, ' ')
          .replace(/\b\w/g, c => c.toUpperCase());

        await conn.query(
          `INSERT INTO template_fields (id, template_id, key_name, field_type, label, is_required, order_index)
           VALUES (?, ?, ?, ?, ?, ?, ?)`,
          [fieldId, templateId, key, 'text', defaultLabel, 1, i]
        );
      }

      await conn.commit();

      return {
        templateId,
        originalFileName: normalizedOriginalName,
        name: templateName,
        variablesCount: variables.length,
        variables
      };
    } catch (dbErr) {
      await conn.rollback();
      cleanupFileIfExists(finalPath);
      throw dbErr;
    } finally {
      conn.release();
    }
  } catch (error) {
    cleanupConvertedFile();
    cleanupFileIfExists(file.path);
    throw error;
  }
}

// Seed data Văn phòng công chứng mặc định
const DEFAULT_OFFICE_ID = 'd3b07384-d113-4ec6-a5d6-c0c2a05d2ed1';

async function seedDefaultOffice() {
  try {
    const [rows] = await pool.query('SELECT id FROM notary_offices WHERE id = ?', [DEFAULT_OFFICE_ID]);
    if (rows.length === 0) {
      await pool.query(
        `INSERT INTO notary_offices (id, name, email, password_hash, phone, status) 
         VALUES (?, ?, ?, ?, ?, ?)`,
        [
          DEFAULT_OFFICE_ID,
          'Văn phòng Công chứng Trung tâm',
          'admin@congchungso.vn',
          '$2b$10$NotaryOfficeDefaultPasswordHashHere', // password giả định
          '0901234567',
          'active'
        ]
      );
      console.log('🌱 Đã tự động tạo văn phòng công chứng mặc định.');
    }
  } catch (error) {
    console.error('Không thể tạo dữ liệu văn phòng công chứng mặc định:', error.message);
  }
}

async function seedDefaultTemplateCategories() {
  try {
    const [rows] = await pool.query('SELECT COUNT(*) AS cnt FROM template_categories');
    if (rows[0].cnt > 0) {
      return;
    }

    for (let i = 0; i < DEFAULT_TEMPLATE_CATEGORIES.length; i++) {
      const root = DEFAULT_TEMPLATE_CATEGORIES[i];
      const rootId = uuidv4();

      await pool.query(
        'INSERT INTO template_categories (id, name, parent_id, sort_order) VALUES (?, ?, ?, ?)',
        [rootId, root.name, null, i]
      );

      for (let j = 0; j < root.children.length; j++) {
        await pool.query(
          'INSERT INTO template_categories (id, name, parent_id, sort_order) VALUES (?, ?, ?, ?)',
          [uuidv4(), root.children[j], rootId, j]
        );
      }
    }

    console.log('✅ Đã tạo cây danh mục biểu mẫu mặc định.');
  } catch (error) {
    console.error('Không thể tạo danh mục biểu mẫu mặc định:', error.message);
  }
}

// ------------------- API ROUTES -------------------

app.get('/api/template-categories', async (req, res) => {
  try {
    const [rows] = await pool.query(`
      SELECT c.id, c.name, c.parent_id, c.sort_order, c.created_at,
             (SELECT COUNT(*) FROM template_categories cc WHERE cc.parent_id = c.id) AS children_count,
             (SELECT COUNT(*) FROM templates t WHERE t.category_id = c.id) AS templates_count
      FROM template_categories c
      ORDER BY c.parent_id IS NOT NULL, c.sort_order ASC, c.name ASC
    `);
    res.json(rows);
  } catch (error) {
    console.error('Lỗi API GET /api/template-categories:', error);
    res.status(500).json({ error: 'Không thể tải danh mục biểu mẫu.' });
  }
});

app.post('/api/template-categories', async (req, res) => {
  try {
    const name = (req.body.name || '').trim();
    const parentId = req.body.parentId || null;

    if (!name) {
      return res.status(400).json({ error: 'Tên danh mục không được để trống.' });
    }

    if (parentId) {
      const [parents] = await pool.query('SELECT id FROM template_categories WHERE id = ?', [parentId]);
      if (parents.length === 0) {
        return res.status(400).json({ error: 'Danh mục cha không tồn tại.' });
      }
    }

    const [existing] = await pool.query(
      'SELECT id FROM template_categories WHERE name = ? AND ((parent_id IS NULL AND ? IS NULL) OR parent_id = ?)',
      [name, parentId, parentId]
    );

    if (existing.length > 0) {
      return res.status(409).json({ error: 'Danh mục này đã tồn tại trong cùng cấp.' });
    }

    const [[{ nextSortOrder }]] = await pool.query(
      'SELECT COALESCE(MAX(sort_order), -1) + 1 AS nextSortOrder FROM template_categories WHERE ((parent_id IS NULL AND ? IS NULL) OR parent_id = ?)',
      [parentId, parentId]
    );

    const category = {
      id: uuidv4(),
      name,
      parent_id: parentId,
      sort_order: nextSortOrder
    };

    await pool.query(
      'INSERT INTO template_categories (id, name, parent_id, sort_order) VALUES (?, ?, ?, ?)',
      [category.id, category.name, category.parent_id, category.sort_order]
    );

    res.status(201).json(category);
  } catch (error) {
    console.error('Lỗi API POST /api/template-categories:', error);
    res.status(500).json({ error: 'Không thể tạo danh mục biểu mẫu.' });
  }
});

// 1. Tải lên Template mới và bóc tách biến
app.post('/api/templates', (req, res, next) => {
  upload.fields([
    { name: 'templateFile', maxCount: 10 },
    { name: 'templateFiles', maxCount: 10 }
  ])(req, res, (err) => {
    if (err) {
      return res.status(400).json({ error: err.message });
    }
    next();
  });
}, async (req, res) => {
  try {
    const files = [
      ...(req.files?.templateFiles || []),
      ...(req.files?.templateFile || [])
    ];

    if (files.length === 0) {
      return res.status(400).json({ error: 'Vui lòng chọn ít nhất một file .doc hoặc .docx mẫu.' });
    }

    if (files.length > 10) {
      files.forEach(file => cleanupFileIfExists(file.path));
      return res.status(400).json({ error: 'Mỗi lần chỉ được tải tối đa 10 file.' });
    }

    const categoryId = req.body.categoryId || null;
    const officeId = req.body.officeId || DEFAULT_OFFICE_ID;

    if (categoryId) {
      const [categories] = await pool.query('SELECT id FROM template_categories WHERE id = ?', [categoryId]);
      if (categories.length === 0) {
        files.forEach(file => cleanupFileIfExists(file.path));
        return res.status(400).json({ error: 'Danh mục biểu mẫu không tồn tại.' });
      }
    }

    const requestedName = (req.body.name || '').trim();
    const results = [];

    for (let index = 0; index < files.length; index++) {
      const file = files[index];
      try {
        const created = await createTemplateFromUploadedFile(file, {
          templateName: files.length === 1 && requestedName ? requestedName : undefined,
          categoryId,
          officeId
        });
        results.push({
          fileName: created.originalFileName,
          status: 'success',
          ...created
        });
      } catch (error) {
        results.push({
          fileName: normalizeVietnameseFileName(file.originalname),
          status: 'failed',
          error: error.message || 'Không thể xử lý file này.'
        });
      }
    }

    const successResults = results.filter(result => result.status === 'success');
    const failedResults = results.filter(result => result.status === 'failed');
    const allSucceeded = failedResults.length === 0;

    return res.status(allSucceeded ? 201 : successResults.length > 0 ? 207 : 400).json({
      message: allSucceeded
        ? `Tải lên thành công ${successResults.length} biểu mẫu.`
        : successResults.length > 0
          ? `Đã xử lý ${successResults.length}/${results.length} biểu mẫu.`
          : 'Không có biểu mẫu nào được xử lý thành công.',
      totalFiles: results.length,
      successCount: successResults.length,
      failureCount: failedResults.length,
      results
    });
  } catch (error) {
    console.error('Lỗi API POST /api/templates:', error);
    res.status(500).json({ error: 'Đã xảy ra lỗi trên hệ thống khi xử lý upload file.' });
  }
});

// 2. Lấy danh sách tất cả các biểu mẫu
// Helper: simple slugify Vietnamese (remove diacritics, spaces to underscores, lower case)
function slugifyVietnamese(str) {
  const from = "ÁÀÂÄÃÅáàâäãåÉÈÊËéèêëÍÌÎÏíìîïÓÒÔÖÕØóòôöõøÚÙÛÜúùûüÑñÇçÝŸýÿŽžšđâăêôơưĂÂÊÔƠƯĐ";
  const to   = "AAAAAAaaaaaaEEEEeeeeIIIIiiiiOOOOOOooooooUUUUuuuuNnCcYYyyZzsd aaeoo uAAE OOUU D";
  const mapping = {};
  for (let i = 0; i < from.length; i++) mapping[from[i]] = to[i];
  const slug = str
    .split('')
    .map(ch => mapping[ch] || ch)
    .join('')
    .replace(/[^\w\s-]/g, '')
    .trim()
    .replace(/\s+/g, '_')
    .toLowerCase();
  return slug;
}

// Simple field dictionary for common Vietnamese headers
const FIELD_DICT = {
  'stt': { name: 'stt', type: 'number' },
  'tên': { name: 'fullname', type: 'text' },
  'họ tên': { name: 'fullname', type: 'text' },
  'địa chỉ': { name: 'address', type: 'text' },
  'mã số thuế': { name: 'tax_id', type: 'text' },
  'giấy tờ pháp nhân': { name: 'id_number', type: 'text' },
  'diện tích': { name: 'land_area', type: 'number' },
  'tỷ lệ': { name: 'ownership_ratio', type: 'text' },
  'ghi chú': { name: 'note', type: 'text' }
};

app.get('/api/templates', async (req, res) => {
  try {
    const [rows] = await pool.query(`
      SELECT t.id, t.name, t.file_path, t.status, t.created_at, t.parent_template_id, t.is_repeated,
             t.category_id, c.name AS category_name, c.parent_id AS category_parent_id,
             (SELECT COUNT(*) FROM template_fields WHERE template_id = t.id) as fields_count,
             (SELECT COUNT(*) FROM templates WHERE parent_template_id = t.id) as children_count
      FROM templates t
      LEFT JOIN template_categories c ON c.id = t.category_id
      ORDER BY t.created_at DESC
    `);
    res.json(rows);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Không thể tải danh sách biểu mẫu.' });
  }
});
// 3. Lấy cấu hình form động
app.get('/api/templates/:id/form', async (req, res) => {
  try {
    const templateId = req.params.id;
    const [templates] = await pool.query('SELECT * FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu yêu cầu.' });
    }
    const [fields] = await pool.query(
      'SELECT id, key_name, field_type, label, is_required, order_index, replace_text, paragraph_context, parent_field_key FROM template_fields WHERE template_id = ? ORDER BY order_index ASC',
      [templateId]
    );
    res.json({ template: templates[0], fields });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Lỗi tải cấu hình form.' });
  }
});

// ---------- New API: Get table structure and suggested fields ----------
app.get('/api/templates/:id/tables', async (req, res) => {
  try {
    const templateId = req.params.id;
    const [templates] = await pool.query('SELECT * FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu.' });
    }
    const template = templates[0];
    const absolutePath = path.join(__dirname, '..', template.file_path);
    const tables = scanTables(absolutePath);
    // Build suggested fields for each table header
    const tablesWithSuggestions = tables.map(tbl => {
      const suggestedFields = tbl.headers.map((header, colIdx) => {
        const lower = header.trim().toLowerCase();
        const dict = FIELD_DICT[lower];
        const key_name = dict ? dict.name : slugifyVietnamese(header);
        const field_type = dict ? dict.type : 'text';
        return { colIndex: colIdx, header, key_name, field_type, label: header };
      });
      return { tableIndex: tbl.tableIndex, headers: tbl.headers, rows: tbl.rows, suggestedFields };
    });
    res.json({ tables: tablesWithSuggestions });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Lỗi lấy cấu trúc bảng.' });
  }
});

// ---------- New API: Inject placeholders into table rows ----------
app.post('/api/templates/:id/inject-table', async (req, res) => {
  try {
    const templateId = req.params.id;
    const { tableIndex, fields, selectedRows } = req.body; // fields: array of {colIndex, key_name, label, field_type, is_required}, selectedRows: array of 1-indexed numbers
    if (typeof tableIndex !== 'number' || !Array.isArray(fields) || !Array.isArray(selectedRows)) {
      return res.status(400).json({ error: 'Dữ liệu yêu cầu không hợp lệ.' });
    }
    const [templates] = await pool.query('SELECT * FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu.' });
    }
    const template = templates[0];
    const absolutePath = path.join(__dirname, '..', template.file_path);
    // Scan tables to locate target table
    const tables = scanTables(absolutePath);
    const targetTable = tables.find(t => t.tableIndex === tableIndex);
    if (!targetTable) {
      return res.status(404).json({ error: 'Bảng mục tiêu không tồn tại.' });
    }
    // Inject placeholders into target table cells in DOCX
    try {
      injectTablePlaceholders(absolutePath, absolutePath, tableIndex, fields, selectedRows);
    } catch (injectErr) {
      console.error('Inject error:', injectErr);
      return res.status(500).json({ error: 'Không thể gài placeholder vào bảng của file.' });
    }
    // Save each generated field into DB
    const conn = await pool.getConnection();
    try {
      await conn.beginTransaction();
      for (const r of selectedRows) {
        for (const f of fields) {
          const fieldId = uuidv4();
          const key_name = `${f.key_name}_${r}`;
          const label = f.label || f.key_name;
          await conn.query(
            `INSERT INTO template_fields (id, template_id, key_name, field_type, label, is_required, order_index, replace_text, paragraph_context) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
            [fieldId, templateId, key_name, f.field_type, label, f.is_required ? 1 : 0, 0, null, null]
          );
        }
      }
      await conn.commit();
      res.json({ message: 'Đã gài placeholder và tạo các field thành công.' });
    } catch (dbErr) {
      await conn.rollback();
      console.error(dbErr);
      res.status(500).json({ error: 'Lỗi lưu các field vào DB.' });
    } finally {
      conn.release();
    }
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Lỗi xử lý inject-table.' });
  }
});

// 4. Cập nhật cấu hình trường dữ liệu (hỗ trợ tạo biến nhanh từ bôi đen)
app.put('/api/templates/:id/fields', async (req, res) => {
  try {
    const templateId = req.params.id;
    const { fields } = req.body; // Mảng chứa các đối tượng field: { id, key_name, label, field_type, is_required, order_index, replace_text }

    if (!Array.isArray(fields)) {
      return res.status(400).json({ error: 'Cấu trúc trường dữ liệu gửi lên không hợp lệ.' });
    }

    // 1. Kiểm tra template tồn tại để lấy đường dẫn file gốc
    const [templates] = await pool.query('SELECT * FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Biểu mẫu không tồn tại trên hệ thống.' });
    }
    const template = templates[0];
    const absoluteTemplatePath = path.join(__dirname, '..', template.file_path);

    // 2. Kiểm tra xem có trường nào được tạo nhanh bằng cách bôi đen text không
    const replacements = fields
      .filter(f => f.replace_text && typeof f.replace_text === 'string' && f.replace_text.trim() !== '')
      .map(f => ({
        searchText: f.replace_text.trim(),
        key_name: f.key_name,
        paragraph_context: f.paragraph_context
      }));

    // Nếu có bôi đen bóc tách, tiến hành sửa đổi trực tiếp vào file mẫu .docx gốc
    if (replacements.length > 0) {
      try {
        injectPlaceholders(absoluteTemplatePath, absoluteTemplatePath, replacements);
      } catch (err) {
        console.error('Lỗi khi ghi đè biến động vào docx:', err);
        return res.status(500).json({ error: 'Không thể ghi đè biến trực quan vào file Word: ' + err.message });
      }
    }

    const conn = await pool.getConnection();
    try {
      await conn.beginTransaction();

      // Xóa tất cả cấu hình fields cũ của template này
      await conn.query('DELETE FROM template_fields WHERE template_id = ?', [templateId]);

      // Thêm mới lại toàn bộ danh sách fields đã cập nhật
      for (const field of fields) {
        const fieldId = (field.id && !field.id.startsWith('temp-')) ? field.id : uuidv4();
        await conn.query(
          `INSERT INTO template_fields (id, template_id, key_name, field_type, label, is_required, order_index, replace_text, paragraph_context, parent_field_key)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
          [
            fieldId,
            templateId,
            field.key_name,
            field.field_type,
            field.label,
            field.is_required ? 1 : 0,
            field.order_index,
            field.replace_text || null,
            field.paragraph_context || null,
            field.parent_field_key || null
          ]
        );
      }

      // Kích hoạt trạng thái active cho template khi đã cấu hình xong fields
      await conn.query("UPDATE templates SET status = 'active' WHERE id = ?", [templateId]);

      await conn.commit();
      res.json({ message: 'Lưu cấu hình và gài biến thành công!' });
    } catch (dbErr) {
      await conn.rollback();
      throw dbErr;
    } finally {
      conn.release();
    }
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Đã xảy ra lỗi khi lưu cấu hình form.' });
  }
});

// 4.1. Lấy danh sách các biểu mẫu con đang liên kết với biểu mẫu gốc
app.get('/api/templates/:id/links', async (req, res) => {
  try {
    const { id } = req.params;
    const [rows] = await pool.query(
      'SELECT id, name, file_path, status, created_at, is_repeated FROM templates WHERE parent_template_id = ?',
      [id]
    );
    res.json(rows);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Không thể tải danh sách biểu mẫu con liên kết.' });
  }
});

// 4.2. Liên kết biểu mẫu con vào biểu mẫu gốc
app.post('/api/templates/:id/link', async (req, res) => {
  try {
    const parentId = req.params.id;
    const { childTemplateId } = req.body;
    if (!childTemplateId) {
      return res.status(400).json({ error: 'Vui lòng cung cấp ID biểu mẫu con cần liên kết.' });
    }

    // 1. Không cho liên kết chính nó
    if (parentId === childTemplateId) {
      return res.status(400).json({ error: 'Không thể liên kết biểu mẫu với chính nó.' });
    }

    // 2. Kiểm tra parent tồn tại
    const [parentRows] = await pool.query(
      'SELECT id, parent_template_id FROM templates WHERE id = ?',
      [parentId]
    );
    if (parentRows.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu gốc.' });
    }

    // 3. Kiểm tra child tồn tại
    const [childRows] = await pool.query(
      'SELECT id, parent_template_id FROM templates WHERE id = ?',
      [childTemplateId]
    );
    if (childRows.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu con cần liên kết.' });
    }
    const child = childRows[0];

    // 4. File con đã có file gốc khác chưa? (mỗi file con chỉ thuộc 1 file gốc)
    if (child.parent_template_id && child.parent_template_id !== parentId) {
      return res.status(409).json({
        error: 'Biểu mẫu này đã được liên kết với một biểu mẫu gốc khác. Vui lòng hủy liên kết cũ trước.'
      });
    }
    if (child.parent_template_id === parentId) {
      return res.status(409).json({ error: 'Biểu mẫu con này đã được liên kết với biểu mẫu gốc rồi.' });
    }

    // 5. File gốc dự kiến lại đang là file con của template khác? (hệ thống chỉ hỗ trợ 2 cấp Master-Child)
    if (parentRows[0].parent_template_id) {
      return res.status(409).json({
        error: 'Biểu mẫu gốc này đang là file con của một biểu mẫu khác. Hệ thống chỉ hỗ trợ liên kết 2 cấp (Gốc → Con).'
      });
    }

    // 6. File con dự kiến lại đang là file gốc của các template khác? (tránh nesting nhiều cấp)
    const [grandChildren] = await pool.query(
      'SELECT COUNT(*) AS cnt FROM templates WHERE parent_template_id = ?',
      [childTemplateId]
    );
    if (grandChildren[0].cnt > 0) {
      return res.status(409).json({
        error: 'Biểu mẫu này đang là file gốc của các biểu mẫu con khác nên không thể trở thành file con. Hệ thống chỉ hỗ trợ liên kết 2 cấp.'
      });
    }

    // 7. Ngăn vòng lặp liên kết (A → B → ... → A)
    let cursor = parentRows[0].parent_template_id;
    const visited = new Set([parentId]);
    while (cursor) {
      if (cursor === childTemplateId) {
        return res.status(409).json({ error: 'Liên kết này tạo ra vòng lặp giữa các biểu mẫu. Thao tác bị từ chối.' });
      }
      if (visited.has(cursor)) break; // an toàn nếu dữ liệu cũ đã có vòng lặp
      visited.add(cursor);
      const [next] = await pool.query('SELECT parent_template_id FROM templates WHERE id = ?', [cursor]);
      cursor = next.length > 0 ? next[0].parent_template_id : null;
    }

    await pool.query('UPDATE templates SET parent_template_id = ? WHERE id = ?', [parentId, childTemplateId]);
    res.json({ message: 'Đã liên kết biểu mẫu thành công!' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Lỗi thiết lập liên kết biểu mẫu.' });
  }
});

// 4.3. Hủy liên kết biểu mẫu con
app.post('/api/templates/:id/unlink', async (req, res) => {
  try {
    const { id } = req.params;
    await pool.query('UPDATE templates SET parent_template_id = NULL WHERE id = ?', [id]);
    res.json({ message: 'Đã hủy liên kết biểu mẫu thành công!' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Lỗi hủy liên kết biểu mẫu.' });
  }
});

// 4.4. Lấy danh sách trường dữ liệu của biểu mẫu gốc để biểu mẫu con ánh xạ
app.get('/api/templates/:id/parent-fields', async (req, res) => {
  try {
    const childId = req.params.id;
    const [templates] = await pool.query('SELECT parent_template_id FROM templates WHERE id = ?', [childId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu.' });
    }
    const parentId = templates[0].parent_template_id;
    if (!parentId) {
      return res.json([]);
    }
    const [fields] = await pool.query(
      'SELECT id, key_name, label FROM template_fields WHERE template_id = ? ORDER BY order_index ASC',
      [parentId]
    );
    res.json(fields);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Lỗi lấy danh sách trường của biểu mẫu gốc.' });
  }
});

// 4.5. Khôi phục biến thành văn bản gốc trong Word và xóa cấu hình biến
app.post('/api/templates/:templateId/fields/:fieldId/restore', async (req, res) => {
  try {
    const { templateId, fieldId } = req.params;

    // Lấy thông tin field
    const [fields] = await pool.query(
      'SELECT key_name, replace_text FROM template_fields WHERE id = ? AND template_id = ?',
      [fieldId, templateId]
    );

    if (fields.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy cấu hình biến cần khôi phục.' });
    }

    const { key_name, replace_text } = fields[0];
    
    // Nếu biến được tạo thủ công (không có chữ gốc replace_text)
    if (!replace_text) {
      await pool.query('DELETE FROM template_fields WHERE id = ?', [fieldId]);
      return res.json({ message: `Đã xóa cấu hình biến thủ công {{${key_name}}} thành công!` });
    }

    // Lấy thông tin template
    const [templates] = await pool.query('SELECT file_path FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu tương ứng.' });
    }

    const template = templates[0];
    const absoluteTemplatePath = path.join(__dirname, '..', template.file_path);

    // Thực hiện khôi phục trong file word
    try {
      const content = fs.readFileSync(absoluteTemplatePath, "binary");
      const zip = new PizZip(content);
      
      const xmlFiles = zip.file(/word\/.*\.xml/);
      let modified = false;
      const placeholder = `{{${key_name}}}`;
      
      for (const file of xmlFiles) {
        let docXml = file.asText();
        if (docXml.includes(placeholder)) {
          docXml = restorePlaceholder(docXml, key_name, replace_text);
          zip.file(file.name, docXml);
          modified = true;
        }
      }

      if (modified) {
        const buffer = zip.generate({ type: "nodebuffer" });
        fs.writeFileSync(absoluteTemplatePath, buffer);
        console.log(`[Restoring] Đã khôi phục biến {{${key_name}}} thành "${replace_text}"`);
      } else {
        console.log(`[Restoring] Không tìm thấy biến {{${key_name}}} trong các file XML.`);
      }
    } catch (err) {
      console.error('Lỗi khi khôi phục biến trong file docx:', err);
      return res.status(500).json({ error: 'Không thể ghi đè khôi phục tệp Word: ' + err.message });
    }

    // Xóa biến khỏi database
    await pool.query('DELETE FROM template_fields WHERE id = ?', [fieldId]);

    res.json({ message: `Đã khôi phục biến {{${key_name}}} thành văn bản gốc thành công!` });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Lỗi khôi phục biến.' });
  }
});

// 5. Khách hàng nộp Form -> Lưu dữ liệu dạng JSON phẳng -> Điền ngược lại docx
const prepareValuesForTemplate = (fields, data, contextLabel = '') => {
  const padded = {};
  const missingRequired = [];

  fields.forEach(f => {
    let rawVal = undefined;
    let isMapped = false;
    if (f.parent_field_key) {
      rawVal = data[f.parent_field_key];
      isMapped = true;
    } else {
      rawVal = data[f.key_name];
    }

    const hasValue =
      (typeof rawVal === 'string' && rawVal.trim() !== '') ||
      (typeof rawVal === 'number') ||
      (typeof rawVal === 'boolean');

    if (typeof rawVal === 'string' && rawVal.trim() !== '') {
      padded[f.key_name] = ` ${rawVal.trim()} `;
    } else if (typeof rawVal === 'number') {
      padded[f.key_name] = ` ${rawVal} `;
    } else if (typeof rawVal === 'boolean') {
      padded[f.key_name] = rawVal ? ' Có ' : ' Không ';
    } else if (f.replace_text && f.replace_text.trim() !== '') {
      padded[f.key_name] = ` ${f.replace_text.trim()} `;
    } else {
      padded[f.key_name] = '';
    }

    if (
      f.is_required &&
      !hasValue &&
      !(f.replace_text && f.replace_text.trim() !== '')
    ) {
      missingRequired.push({
        key_name: f.key_name,
        label: f.label || f.key_name,
        mappedFrom: isMapped ? f.parent_field_key : null,
        context: contextLabel
      });
    }
  });

  for (const key in data) {
    if (!(key in padded)) {
      const val = data[key];
      if (typeof val === 'string' && val.trim() !== '') {
        padded[key] = ` ${val.trim()} `;
      } else if (typeof val === 'boolean') {
        padded[key] = val ? ' Có ' : ' Không ';
      } else {
        padded[key] = val;
      }
    }
  }

  return { padded, missingRequired };
};

const prepareValuesForSingleRecord = (fields, recordData, parentData, contextLabel = '') => {
  const padded = {};
  const missingRequired = [];

  fields.forEach(f => {
    let rawVal = undefined;
    let isMapped = false;
    if (f.parent_field_key) {
      rawVal = parentData[f.parent_field_key];
      isMapped = true;
    } else {
      rawVal = recordData[f.key_name];
    }

    const hasValue =
      (typeof rawVal === 'string' && rawVal.trim() !== '') ||
      (typeof rawVal === 'number') ||
      (typeof rawVal === 'boolean');

    if (typeof rawVal === 'string' && rawVal.trim() !== '') {
      padded[f.key_name] = ` ${rawVal.trim()} `;
    } else if (typeof rawVal === 'number') {
      padded[f.key_name] = ` ${rawVal} `;
    } else if (typeof rawVal === 'boolean') {
      padded[f.key_name] = rawVal ? ' Có ' : ' Không ';
    } else if (f.replace_text && f.replace_text.trim() !== '') {
      padded[f.key_name] = ` ${f.replace_text.trim()} `;
    } else {
      padded[f.key_name] = '';
    }

    if (
      f.is_required &&
      !hasValue &&
      !(f.replace_text && f.replace_text.trim() !== '')
    ) {
      missingRequired.push({
        key_name: f.key_name,
        label: f.label || f.key_name,
        mappedFrom: isMapped ? f.parent_field_key : null,
        context: contextLabel
      });
    }
  });

  for (const key in recordData) {
    if (!(key in padded) && key !== '_id') {
      const val = recordData[key];
      if (typeof val === 'string' && val.trim() !== '') {
        padded[key] = ` ${val.trim()} `;
      } else if (typeof val === 'boolean') {
        padded[key] = val ? ' Có ' : ' Không ';
      } else {
        padded[key] = val;
      }
    }
  }

  return { padded, missingRequired };
};

app.post('/api/submissions', async (req, res) => {
  try {
    const { templateId, customerName, customerPhone, values, selectedChildIds } = req.body;

    if (!templateId || !values || typeof values !== 'object') {
      return res.status(400).json({ error: 'Dữ liệu nộp hồ sơ không đầy đủ hoặc sai định dạng.' });
    }

    // 1. Kiểm tra template tồn tại
    const [templates] = await pool.query('SELECT * FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Biểu mẫu không tồn tại trên hệ thống.' });
    }
    const template = templates[0];

    // Lấy danh sách các child templates liên kết
    let [childTemplates] = await pool.query(
      "SELECT * FROM templates WHERE parent_template_id = ? AND status = 'active'",
      [templateId]
    );

    if (Array.isArray(selectedChildIds)) {
      childTemplates = childTemplates.filter(t => selectedChildIds.includes(t.id));
    }

    const submissionId = uuidv4();
    let submissionStatus = 'completed';

    // Map repeated child records to parent suffix variables
    for (const child of childTemplates) {
      if (child.is_repeated) {
        const recordsList = values[child.id];
        if (Array.isArray(recordsList)) {
          const [childFields] = await pool.query(
            'SELECT key_name, parent_field_key FROM template_fields WHERE template_id = ?',
            [child.id]
          );
          
          recordsList.forEach((recordData, rIdx) => {
            childFields.forEach(cf => {
              const targetKey = cf.parent_field_key || cf.key_name;
              const sourceVal = recordData[cf.key_name];
              if (sourceVal !== undefined) {
                values[`${targetKey}_${rIdx + 1}`] = sourceVal;
              }
            });
          });
        }
      }
    }

    const [templateFields] = await pool.query(
      'SELECT key_name, is_required, replace_text, label FROM template_fields WHERE template_id = ?',
      [templateId]
    );
    const masterResult = prepareValuesForTemplate(templateFields, values, template.name);

    // Thu thập tất cả các trường bắt buộc bị thiếu (Master + Children)
    const allMissingRequired = [...masterResult.missingRequired];

    // Chuẩn bị dữ liệu cho tất cả file con và thu thập trường thiếu
    for (const child of childTemplates) {
      const [childFields] = await pool.query(
        'SELECT key_name, is_required, replace_text, parent_field_key, label FROM template_fields WHERE template_id = ?',
        [child.id]
      );
      
      if (child.is_repeated) {
        const recordsList = values[child.id];
        const recordsArray = Array.isArray(recordsList) ? recordsList : [{}];
        
        for (let rIdx = 0; rIdx < recordsArray.length; rIdx++) {
          const recordVal = recordsArray[rIdx];
          const singleResult = prepareValuesForSingleRecord(childFields, recordVal, values, `${child.name} (Bản ghi ${rIdx + 1})`);
          allMissingRequired.push(...singleResult.missingRequired);
        }
      } else {
        const childResult = prepareValuesForTemplate(childFields, values, child.name);
        allMissingRequired.push(...childResult.missingRequired);
      }
    }

    // Kiểm tra trường bắt buộc trước khi lưu
    if (allMissingRequired.length > 0) {
      return res.status(400).json({
        error: 'Thiếu thông tin bắt buộc trong hồ sơ',
        missingFields: allMissingRequired.map(m => ({
          label: m.label,
          key: m.key_name,
          mappedFrom: m.mappedFrom,
          template: m.context
        }))
      });
    }

    // 4. Lưu kết quả nộp hồ sơ vào Database (Chỉ lưu JSON dữ liệu, KHÔNG sinh file vật lý)
    await pool.query(
      `INSERT INTO document_submissions (id, template_id, customer_name, customer_phone, status, values_json, output_file_path, completed_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        submissionId,
        templateId,
        customerName || 'Khách hàng vãng lai',
        customerPhone || null,
        submissionStatus,
        JSON.stringify({ values, selectedChildIds }), // Lưu thông tin điền và danh sách file con được chọn
        null, // Không lưu file vật lý
        submissionStatus === 'completed' ? new Date() : null
      ]
    );

    res.status(201).json({
      message: 'Hồ sơ đã được lưu thành công trên hệ thống!',
      submissionId,
      downloadUrl: `/api/submissions/${submissionId}/download`
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Đã xảy ra lỗi trên hệ thống khi nộp hồ sơ.' });
  }
});

// 5b. Lấy danh sách toàn bộ hồ sơ đã nộp (dành cho vai trò công chứng/admin)
app.get('/api/submissions', async (req, res) => {
  try {
    const [rows] = await pool.query(`
      SELECT s.id, s.template_id, s.customer_name, s.customer_phone, s.status, s.completed_at, t.name as template_name
      FROM document_submissions s
      JOIN templates t ON s.template_id = t.id
      ORDER BY s.completed_at DESC
    `);
    res.json(rows);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Không thể tải lịch sử hồ sơ đã nhận.' });
  }
});

// 5c. Lấy chi tiết hồ sơ đã nộp và các trường cấu hình template tương ứng để map nhãn
// Hàm trộn văn bản và lấy text trần trực tiếp từ RAM
function getMergedDocumentText(templatePath, dataJson) {
  try {
    const content = fs.readFileSync(templatePath, 'binary');
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: {
        start: '{{',
        end: '}}'
      }
    });
    doc.render(dataJson);
    return doc.getFullText();
  } catch (error) {
    console.error("Lỗi khi lấy văn bản đã trộn:", error);
    return "Không thể trích xuất văn bản từ biểu mẫu.";
  }
}

// 5c. Lấy chi tiết hồ sơ đã nộp và các trường cấu hình template tương ứng để map nhãn
app.get('/api/submissions/:id/detail', async (req, res) => {
  try {
    const submissionId = req.params.id;
    const [submissions] = await pool.query(
      'SELECT s.id, s.template_id, s.customer_name, s.customer_phone, s.status, s.values_json, s.completed_at, t.name as template_name, t.file_path FROM document_submissions s JOIN templates t ON s.template_id = t.id WHERE s.id = ?',
      [submissionId]
    );

    if (submissions.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy hồ sơ.' });
    }

    const submission = submissions[0];
    if (typeof submission.values_json === 'string') {
      try {
        submission.values_json = JSON.parse(submission.values_json);
      } catch (e) {
        console.error("Failed to parse values_json in detail:", e);
      }
    }
    const templateId = submission.template_id;

    // Lấy cấu hình các trường của template chính
    const [masterFields] = await pool.query(
      'SELECT key_name, label, field_type, template_id FROM template_fields WHERE template_id = ? ORDER BY order_index ASC',
      [templateId]
    );

    // Lấy cấu hình các trường của các template con liên kết
    const [childFields] = await pool.query(
      'SELECT f.key_name, f.label, f.field_type, f.template_id, t.name as template_name, t.is_repeated FROM template_fields f JOIN templates t ON f.template_id = t.id WHERE t.parent_template_id = ? ORDER BY t.name, f.order_index ASC',
      [templateId]
    );

    // Tự động phân tích và sinh text cho văn bản chính
    const { values, selectedChildIds } = submission.values_json || {};
    const masterResult = prepareValuesForTemplate(masterFields, values || {}, submission.template_name);
    const absoluteTemplatePath = path.join(__dirname, '..', submission.file_path);
    
    let mergedText = "";
    if (fs.existsSync(absoluteTemplatePath)) {
      mergedText = getMergedDocumentText(absoluteTemplatePath, masterResult.padded);
    } else {
      mergedText = "Tệp biểu mẫu gốc không tồn tại trên hệ thống.";
    }

    // Sinh text cho các văn bản con được chọn
    let [childTemplates] = await pool.query(
      "SELECT id, name, file_path, is_repeated FROM templates WHERE parent_template_id = ? AND status = 'active'",
      [templateId]
    );
    if (Array.isArray(selectedChildIds)) {
      childTemplates = childTemplates.filter(t => selectedChildIds.includes(t.id));
    }

    const childDocumentsText = [];
    for (const child of childTemplates) {
      const [cFields] = await pool.query(
        'SELECT key_name, is_required, replace_text, parent_field_key, label FROM template_fields WHERE template_id = ?',
        [child.id]
      );
      const absoluteChildTemplatePath = path.join(__dirname, '..', child.file_path);
      
      if (fs.existsSync(absoluteChildTemplatePath)) {
        if (child.is_repeated) {
          const recordsList = (values || {})[child.id];
          const recordsArray = Array.isArray(recordsList) ? recordsList : [{}];
          recordsArray.forEach((recordVal, rIdx) => {
            const prep = prepareValuesForSingleRecord(cFields, recordVal, values || {}, `${child.name} (Bản ghi ${rIdx + 1})`);
            const textContent = getMergedDocumentText(absoluteChildTemplatePath, prep.padded);
            childDocumentsText.push({
              name: `${child.name} (Bản ghi ${rIdx + 1})`,
              text: textContent
            });
          });
        } else {
          const prep = prepareValuesForTemplate(cFields, values || {}, child.name);
          const textContent = getMergedDocumentText(absoluteChildTemplatePath, prep.padded);
          childDocumentsText.push({
            name: child.name,
            text: textContent
          });
        }
      }
    }

    res.json({
      submission,
      masterFields,
      childFields,
      mergedText,
      childDocumentsText
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Không thể tải chi tiết hồ sơ.' });
  }
});

// 6. Download file kết quả (Trộn động dữ liệu và trả về file ZIP hoặc DOCX tức thì từ RAM)
app.get('/api/submissions/:id/download', async (req, res) => {
  try {
    const submissionId = req.params.id;
    const [rows] = await pool.query(
      'SELECT s.values_json, t.name as template_name, t.id as template_id, t.file_path FROM document_submissions s JOIN templates t ON s.template_id = t.id WHERE s.id = ?',
      [submissionId]
    );

    if (rows.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy hồ sơ.' });
    }

    let valuesJson = rows[0].values_json;
    if (typeof valuesJson === 'string') {
      try {
        valuesJson = JSON.parse(valuesJson);
      } catch (e) {
        console.error("Failed to parse values_json in download:", e);
      }
    }
    const { values, selectedChildIds } = valuesJson || {};
    const safeValues = values || {};
    const parentTemplate = rows[0];

    // Lấy child templates liên kết
    let [childTemplates] = await pool.query(
      "SELECT id, name, file_path, is_repeated FROM templates WHERE parent_template_id = ? AND status = 'active'",
      [parentTemplate.template_id]
    );
    if (Array.isArray(selectedChildIds)) {
      childTemplates = childTemplates.filter(t => selectedChildIds.includes(t.id));
    }

    // Map repeated child records to parent suffix variables
    for (const child of childTemplates) {
      if (child.is_repeated) {
        const recordsList = safeValues[child.id];
        if (Array.isArray(recordsList)) {
          const [childFields] = await pool.query(
            'SELECT key_name, parent_field_key FROM template_fields WHERE template_id = ?',
            [child.id]
          );
          recordsList.forEach((recordData, rIdx) => {
            childFields.forEach(cf => {
              const targetKey = cf.parent_field_key || cf.key_name;
              const sourceVal = recordData[cf.key_name];
              if (sourceVal !== undefined) {
                safeValues[`${targetKey}_${rIdx + 1}`] = sourceVal;
              }
            });
          });
        }
      }
    }

    const [parentFields] = await pool.query(
      'SELECT key_name, is_required, replace_text, label FROM template_fields WHERE template_id = ?',
      [parentTemplate.template_id]
    );
    const masterResult = prepareValuesForTemplate(parentFields, safeValues, parentTemplate.template_name);
    const absoluteTemplatePath = path.join(__dirname, '..', parentTemplate.file_path);

    if (!fs.existsSync(absoluteTemplatePath)) {
      return res.status(400).json({ error: `Tệp biểu mẫu gốc của '${parentTemplate.template_name}' không tồn tại trên máy chủ.` });
    }

    const masterBuffer = mergeDocumentToBuffer(absoluteTemplatePath, masterResult.padded);
    const safeName = parentTemplate.template_name.replace(/[^a-zA-Z0-9À-ỹ\s-_]/g, '');

    // Nếu không có file con liên kết -> Gửi thẳng file docx
    if (childTemplates.length === 0) {
      res.setHeader('Content-Disposition', `attachment; filename=${encodeURIComponent(`${safeName}_HoanThinh.docx`)}`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      return res.send(masterBuffer);
    }

    // Có file con -> Tạo ZIP trực tiếp trả về cho client
    res.setHeader('Content-Disposition', `attachment; filename=${encodeURIComponent(`${safeName}_HoanThinh.zip`)}`);
    res.setHeader('Content-Type', 'application/zip');

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.pipe(res);

    // Thêm master
    archive.append(masterBuffer, { name: `${parentTemplate.template_name}.docx` });

    // Thêm children
    for (const child of childTemplates) {
      const [childFields] = await pool.query(
        'SELECT key_name, is_required, replace_text, parent_field_key, label FROM template_fields WHERE template_id = ?',
        [child.id]
      );

      const absoluteChildTemplatePath = path.join(__dirname, '..', child.file_path);
      if (!fs.existsSync(absoluteChildTemplatePath)) {
        console.warn(`Child template file not found: ${absoluteChildTemplatePath}`);
        continue;
      }

      if (child.is_repeated) {
        const recordsList = safeValues[child.id];
        const recordsArray = Array.isArray(recordsList) ? recordsList : [{}];
        for (let rIdx = 0; rIdx < recordsArray.length; rIdx++) {
          const prep = prepareValuesForSingleRecord(childFields, recordsArray[rIdx], safeValues, `${child.name} (Bản ghi ${rIdx + 1})`);
          const childBuffer = mergeDocumentToBuffer(absoluteChildTemplatePath, prep.padded);
          archive.append(childBuffer, { name: `${child.name}_Căn_${rIdx + 1}.docx` });
        }
      } else {
        const prep = prepareValuesForTemplate(childFields, safeValues, child.name);
        const childBuffer = mergeDocumentToBuffer(absoluteChildTemplatePath, prep.padded);
        archive.append(childBuffer, { name: `${child.name}.docx` });
      }
    }

    archive.finalize();
  } catch (error) {
    console.error("Lỗi khi tải trọn bộ hồ sơ:", error);
    res.status(500).json({ error: 'Không thể tải trọn bộ hồ sơ.' });
  }
});

// 6a. Lấy danh sách tên các file con trong gói hồ sơ nộp (ZIP hoặc đơn lẻ) được sinh động từ DB
app.get('/api/submissions/:id/files', async (req, res) => {
  try {
    const submissionId = req.params.id;
    const [rows] = await pool.query(
      'SELECT s.values_json, t.name as template_name, t.id as template_id FROM document_submissions s JOIN templates t ON s.template_id = t.id WHERE s.id = ?',
      [submissionId]
    );

    if (rows.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy hồ sơ.' });
    }

    let valuesJson = rows[0].values_json;
    if (typeof valuesJson === 'string') {
      try {
        valuesJson = JSON.parse(valuesJson);
      } catch (e) {
        console.error("Failed to parse values_json in files:", e);
      }
    }
    const { values, selectedChildIds } = valuesJson || {};
    const safeValues = values || {};
    const parentName = rows[0].template_name;

    // Lấy child templates
    let [childTemplates] = await pool.query(
      "SELECT id, name, is_repeated FROM templates WHERE parent_template_id = ? AND status = 'active'",
      [rows[0].template_id]
    );

    if (Array.isArray(selectedChildIds)) {
      childTemplates = childTemplates.filter(t => selectedChildIds.includes(t.id));
    }

    if (childTemplates.length === 0) {
      return res.json([`${parentName}.docx`]);
    }

    const fileNames = [`${parentName}.docx`];
    for (const child of childTemplates) {
      if (child.is_repeated) {
        const recordsList = safeValues[child.id];
        const recordsArray = Array.isArray(recordsList) ? recordsList : [{}];
        for (let rIdx = 0; rIdx < recordsArray.length; rIdx++) {
          fileNames.push(`${child.name}_Căn_${rIdx + 1}.docx`);
        }
      } else {
        fileNames.push(`${child.name}.docx`);
      }
    }

    return res.json(fileNames);
  } catch (error) {
    console.error("Lỗi khi liệt kê danh sách file:", error);
    res.status(500).json({ error: 'Không thể liệt kê danh sách file.' });
  }
});

// 6b. Tải một file con cụ thể trong gói hồ sơ nộp, tự động trộn và xuất trực tiếp từ RAM
app.get('/api/submissions/:id/download-file', async (req, res) => {
  try {
    const submissionId = req.params.id;
    const filename = req.query.filename;
    if (!filename) {
      return res.status(400).json({ error: 'Thiếu tên file cần tải.' });
    }

    const [rows] = await pool.query(
      'SELECT s.values_json, t.name as template_name, t.id as template_id, t.file_path FROM document_submissions s JOIN templates t ON s.template_id = t.id WHERE s.id = ?',
      [submissionId]
    );

    if (rows.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy hồ sơ.' });
    }

    let valuesJson = rows[0].values_json;
    if (typeof valuesJson === 'string') {
      try {
        valuesJson = JSON.parse(valuesJson);
      } catch (e) {
        console.error("Failed to parse values_json in download-file:", e);
      }
    }
    const { values, selectedChildIds } = valuesJson || {};
    const safeValues = values || {};
    const parentTemplate = rows[0];

    // 1. Kiểm tra nếu file yêu cầu chính là file master mẹ
    if (filename === `${parentTemplate.template_name}.docx`) {
      let [childTemplatesForSuffix] = await pool.query(
        "SELECT id, is_repeated FROM templates WHERE parent_template_id = ? AND status = 'active'",
        [parentTemplate.template_id]
      );
      if (Array.isArray(selectedChildIds)) {
        childTemplatesForSuffix = childTemplatesForSuffix.filter(t => selectedChildIds.includes(t.id));
      }
      for (const child of childTemplatesForSuffix) {
        if (child.is_repeated) {
          const recordsList = safeValues[child.id];
          if (Array.isArray(recordsList)) {
            const [childFields] = await pool.query(
              'SELECT key_name, parent_field_key FROM template_fields WHERE template_id = ?',
              [child.id]
            );
            recordsList.forEach((recordData, rIdx) => {
              childFields.forEach(cf => {
                const targetKey = cf.parent_field_key || cf.key_name;
                const sourceVal = recordData[cf.key_name];
                if (sourceVal !== undefined) {
                  safeValues[`${targetKey}_${rIdx + 1}`] = sourceVal;
                }
              });
            });
          }
        }
      }

      const [parentFields] = await pool.query(
        'SELECT key_name, is_required, replace_text, label FROM template_fields WHERE template_id = ?',
        [parentTemplate.template_id]
      );
      const masterResult = prepareValuesForTemplate(parentFields, safeValues, parentTemplate.template_name);
      const absoluteTemplatePath = path.join(__dirname, '..', parentTemplate.file_path);

      if (!fs.existsSync(absoluteTemplatePath)) {
        return res.status(400).json({ error: `Tệp biểu mẫu gốc '${parentTemplate.template_name}' không tồn tại trên máy chủ.` });
      }

      const buffer = mergeDocumentToBuffer(absoluteTemplatePath, masterResult.padded);

      res.setHeader('Content-Disposition', `attachment; filename=${encodeURIComponent(filename)}`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      return res.send(buffer);
    }

    // 2. Kiểm tra các file con
    let [childTemplates] = await pool.query(
      "SELECT id, name, file_path, is_repeated FROM templates WHERE parent_template_id = ? AND status = 'active'",
      [parentTemplate.template_id]
    );
    if (Array.isArray(selectedChildIds)) {
      childTemplates = childTemplates.filter(t => selectedChildIds.includes(t.id));
    }

    for (const child of childTemplates) {
      if (child.is_repeated) {
        const recordsList = safeValues[child.id];
        const recordsArray = Array.isArray(recordsList) ? recordsList : [{}];
        for (let rIdx = 0; rIdx < recordsArray.length; rIdx++) {
          if (filename === `${child.name}_Căn_${rIdx + 1}.docx`) {
            const [childFields] = await pool.query(
              'SELECT key_name, is_required, replace_text, parent_field_key, label FROM template_fields WHERE template_id = ?',
              [child.id]
            );
            const prep = prepareValuesForSingleRecord(childFields, recordsArray[rIdx], safeValues, `${child.name} (Bản ghi ${rIdx + 1})`);
            const absoluteChildTemplatePath = path.join(__dirname, '..', child.file_path);

            if (!fs.existsSync(absoluteChildTemplatePath)) {
              return res.status(400).json({ error: `Tệp biểu mẫu con '${child.name}' không tồn tại trên máy chủ.` });
            }

            const buffer = mergeDocumentToBuffer(absoluteChildTemplatePath, prep.padded);

            res.setHeader('Content-Disposition', `attachment; filename=${encodeURIComponent(filename)}`);
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            return res.send(buffer);
          }
        }
      } else {
        if (filename === `${child.name}.docx`) {
          const [childFields] = await pool.query(
            'SELECT key_name, is_required, replace_text, parent_field_key, label FROM template_fields WHERE template_id = ?',
            [child.id]
          );
          const prep = prepareValuesForTemplate(childFields, safeValues, child.name);
          const absoluteChildTemplatePath = path.join(__dirname, '..', child.file_path);

          if (!fs.existsSync(absoluteChildTemplatePath)) {
            return res.status(400).json({ error: `Tệp biểu mẫu con '${child.name}' không tồn tại trên máy chủ.` });
          }

          const buffer = mergeDocumentToBuffer(absoluteChildTemplatePath, prep.padded);

          res.setHeader('Content-Disposition', `attachment; filename=${encodeURIComponent(filename)}`);
          res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
          return res.send(buffer);
        }
      }
    }

    return res.status(404).json({ error: 'Không tìm thấy file tương ứng trong hồ sơ.' });
  } catch (error) {
    console.error("Lỗi khi tải file lẻ:", error);
    res.status(500).json({ error: 'Không thể tải file lẻ.' });
  }
});

// 7. Download file mẫu gốc
app.get('/api/templates/:id/download-original', async (req, res) => {
  try {
    const templateId = req.params.id;
    const [rows] = await pool.query('SELECT file_path, name FROM templates WHERE id = ?', [templateId]);

    if (rows.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy file mẫu.' });
    }

    const absolutePath = path.join(__dirname, '..', rows[0].file_path);
    if (!fs.existsSync(absolutePath)) {
      return res.status(404).json({ error: 'File mẫu gốc không tồn tại trên hệ thống.' });
    }

    res.download(absolutePath, `${rows[0].name}.docx`);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Không thể tải file mẫu.' });
  }
});

// 8. Xóa biểu mẫu (Hard delete: xóa CSDL cascade và xóa các tệp vật lý liên quan)
app.delete('/api/templates/:id', async (req, res) => {
  try {
    const templateId = req.params.id;

    // 1. Lấy thông tin file path của template
    const [templates] = await pool.query('SELECT file_path FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu cần xóa.' });
    }
    const template = templates[0];

    // 2. Lấy danh sách các submission để xóa file output của chúng
    const [submissions] = await pool.query(
      'SELECT output_file_path FROM document_submissions WHERE template_id = ?',
      [templateId]
    );

    // 3. Xóa record trong DB (sẽ tự động cascade delete template_fields và document_submissions)
    await pool.query('DELETE FROM templates WHERE id = ?', [templateId]);

    // 4. Xóa file template vật lý
    const absoluteTemplatePath = path.join(__dirname, '..', template.file_path);
    if (fs.existsSync(absoluteTemplatePath)) {
      try {
        fs.unlinkSync(absoluteTemplatePath);
      } catch (err) {
        console.error(`Không thể xóa file template vật lý ${absoluteTemplatePath}:`, err);
      }
    }

    // 5. Xóa các file output vật lý của submissions
    for (const sub of submissions) {
      if (sub.output_file_path) {
        const absoluteOutputPath = path.join(__dirname, '..', sub.output_file_path);
        if (fs.existsSync(absoluteOutputPath)) {
          try {
            fs.unlinkSync(absoluteOutputPath);
          } catch (err) {
            console.error(`Không thể xóa file output vật lý ${absoluteOutputPath}:`, err);
          }
        }
      }
    }

    res.json({ message: 'Đã xóa biểu mẫu và toàn bộ dữ liệu liên quan thành công!' });
  } catch (error) {
    console.error('Lỗi khi xóa biểu mẫu:', error);
    res.status(500).json({ error: 'Đã xảy ra lỗi trên hệ thống khi xóa biểu mẫu.' });
  }
});

// 9. Cập nhật thông tin biểu mẫu (Đổi tên)
app.put('/api/templates/:id', async (req, res) => {
  try {
    const templateId = req.params.id;
    const { name } = req.body;

    if (!name || name.trim() === '') {
      return res.status(400).json({ error: 'Tên biểu mẫu không được để trống.' });
    }

    const [result] = await pool.query(
      'UPDATE templates SET name = ? WHERE id = ?',
      [name.trim(), templateId]
    );

    if (result.affectedRows === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu cần cập nhật.' });
    }

    res.json({ message: 'Đã cập nhật tên biểu mẫu thành công!' });
  } catch (error) {
    console.error('Lỗi khi cập nhật tên biểu mẫu:', error);
    res.status(500).json({ error: 'Đã xảy ra lỗi trên hệ thống khi cập nhật tên biểu mẫu.' });
  }
});

app.put('/api/templates/:id/category', async (req, res) => {
  try {
    const templateId = req.params.id;
    const categoryId = req.body.categoryId || null;

    if (categoryId) {
      const [categories] = await pool.query('SELECT id FROM template_categories WHERE id = ?', [categoryId]);
      if (categories.length === 0) {
        return res.status(400).json({ error: 'Danh mục biểu mẫu không tồn tại.' });
      }
    }

    const [result] = await pool.query(
      'UPDATE templates SET category_id = ? WHERE id = ?',
      [categoryId, templateId]
    );

    if (result.affectedRows === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu cần cập nhật.' });
    }

    res.json({ message: 'Đã cập nhật danh mục biểu mẫu thành công!' });
  } catch (error) {
    console.error('Lỗi khi cập nhật danh mục biểu mẫu:', error);
    res.status(500).json({ error: 'Đã xảy ra lỗi trên hệ thống khi cập nhật danh mục biểu mẫu.' });
  }
});

// 9.2 Cập nhật chế độ lặp biểu mẫu con
app.put('/api/templates/:id/repeated', async (req, res) => {
  try {
    const templateId = req.params.id;
    const { isRepeated } = req.body;

    await pool.query(
      'UPDATE templates SET is_repeated = ? WHERE id = ?',
      [isRepeated ? 1 : 0, templateId]
    );

    res.json({ message: 'Đã cập nhật chế độ lặp biểu mẫu thành công!' });
  } catch (error) {
    console.error('Lỗi khi cập nhật chế độ lặp:', error);
    res.status(500).json({ error: 'Đã xảy ra lỗi trên hệ thống khi cập nhật chế độ lặp.' });
  }
});

// 10. Sao chép biểu mẫu (duplicate)
app.post('/api/templates/:id/duplicate', async (req, res) => {
  const conn = await pool.getConnection();
  try {
    const originalTemplateId = req.params.id;

    // Lấy thông tin biểu mẫu gốc
    const [templates] = await conn.query('SELECT * FROM templates WHERE id = ?', [originalTemplateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu gốc cần sao chép.' });
    }
    const originalTemplate = templates[0];
    const newTemplateId = uuidv4();
    
    // Tạo đường dẫn file mẫu mới
    const ext = path.extname(originalTemplate.file_path) || '.docx';
    const newFileName = `${newTemplateId}${ext}`;
    const newFilePath = path.join(TEMPLATES_DIR, newFileName);
    const absoluteOriginalPath = path.join(__dirname, '..', originalTemplate.file_path);

    // Sao chép file vật lý
    if (fs.existsSync(absoluteOriginalPath)) {
      fs.copyFileSync(absoluteOriginalPath, newFilePath);
    } else {
      return res.status(404).json({ error: 'Tệp tin Word của biểu mẫu gốc không tồn tại vật lý.' });
    }

    const relativeNewFilePath = path.relative(path.join(__dirname, '..'), newFilePath);

    // Bắt đầu transaction để lưu vào DB
    await conn.beginTransaction();

    // Insert dòng mới vào bảng templates
    const newName = `[Sao chép] ${originalTemplate.name}`;
    await conn.query(
      'INSERT INTO templates (id, office_id, category_id, name, file_path, status, parent_template_id) VALUES (?, ?, ?, ?, ?, ?, ?)',
      [newTemplateId, originalTemplate.office_id, originalTemplate.category_id || null, newName, relativeNewFilePath, 'draft', null]
    );

    // Lấy các trường cấu hình của biểu mẫu gốc
    const [fields] = await conn.query(
      'SELECT key_name, field_type, label, is_required, order_index, replace_text, paragraph_context FROM template_fields WHERE template_id = ?',
      [originalTemplateId]
    );

    // Insert các trường tương ứng cho biểu mẫu mới
    for (const field of fields) {
      const newFieldId = uuidv4();
      await conn.query(
        `INSERT INTO template_fields (id, template_id, key_name, field_type, label, is_required, order_index, replace_text, paragraph_context, parent_field_key)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [
          newFieldId,
          newTemplateId,
          field.key_name,
          field.field_type,
          field.label,
          field.is_required,
          field.order_index,
          field.replace_text,
          field.paragraph_context,
          null
        ]
      );
    }

    await conn.commit();
    res.status(201).json({
      message: 'Sao chép biểu mẫu thành công!',
      newTemplateId,
      name: newName
    });
  } catch (error) {
    await conn.rollback();
    console.error('Lỗi khi sao chép biểu mẫu:', error);
    res.status(500).json({ error: 'Đã xảy ra lỗi trên hệ thống khi sao chép biểu mẫu.' });
  } finally {
    conn.release();
  }
});

// 11. Xuất cấu hình trường (export config)
app.get('/api/templates/:id/export', async (req, res) => {
  try {
    const templateId = req.params.id;
    const [templates] = await pool.query('SELECT name FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu.' });
    }
    const [fields] = await pool.query(
      'SELECT key_name, field_type, label, is_required, order_index, replace_text, paragraph_context, parent_field_key FROM template_fields WHERE template_id = ? ORDER BY order_index ASC',
      [templateId]
    );
    res.json({
      templateName: templates[0].name,
      fields
    });
  } catch (error) {
    console.error('Lỗi khi xuất cấu hình trường:', error);
    res.status(500).json({ error: 'Không thể xuất cấu hình trường của biểu mẫu này.' });
  }
});

// 12. Nhập cấu hình trường (import config)
app.post('/api/templates/:id/import', async (req, res) => {
  const conn = await pool.getConnection();
  try {
    const templateId = req.params.id;
    const { fields } = req.body;

    if (!Array.isArray(fields)) {
      return res.status(400).json({ error: 'Dữ liệu cấu hình trường nhập vào không hợp lệ (yêu cầu mảng).' });
    }

    // Kiểm tra template tồn tại
    const [templates] = await conn.query('SELECT id FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu đích để nhập cấu hình.' });
    }

    await conn.beginTransaction();

    // Xóa các trường cấu hình cũ
    await conn.query('DELETE FROM template_fields WHERE template_id = ?', [templateId]);

    // Thêm mới các trường cấu hình từ JSON
    for (const field of fields) {
      const fieldId = uuidv4();
      await conn.query(
        `INSERT INTO template_fields (id, template_id, key_name, field_type, label, is_required, order_index, replace_text, paragraph_context, parent_field_key)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [
          fieldId,
          templateId,
          field.key_name,
          field.field_type || 'text',
          field.label || field.key_name,
          field.is_required ? 1 : 0,
          field.order_index || 0,
          field.replace_text || null,
          field.paragraph_context || null,
          field.parent_field_key || null
        ]
      );
    }

    await conn.commit();
    res.json({ message: 'Nhập cấu hình các trường thành công!' });
  } catch (error) {
    await conn.rollback();
    console.error('Lỗi khi nhập cấu hình trường:', error);
    res.status(500).json({ error: 'Đã xảy ra lỗi trên hệ thống khi nhập cấu hình trường.' });
  } finally {
    conn.release();
  }
});



async function migrateDatabase() {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS notary_offices (
        id CHAR(36) PRIMARY KEY,
        name VARCHAR(255) NOT NULL,
        email VARCHAR(191) NOT NULL UNIQUE,
        password_hash VARCHAR(255) NOT NULL,
        phone VARCHAR(20),
        status VARCHAR(20) NOT NULL DEFAULT 'active',
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP
      )
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS template_categories (
        id CHAR(36) PRIMARY KEY,
        name VARCHAR(255) NOT NULL,
        parent_id CHAR(36),
        sort_order INTEGER DEFAULT 0,
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        CONSTRAINT fk_template_categories_parent
          FOREIGN KEY (parent_id) REFERENCES template_categories(id) ON DELETE SET NULL
      )
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS templates (
        id CHAR(36) PRIMARY KEY,
        office_id CHAR(36) NOT NULL,
        category_id CHAR(36),
        name VARCHAR(255) NOT NULL,
        file_path VARCHAR(500) NOT NULL,
        status VARCHAR(20) NOT NULL DEFAULT 'draft',
        parent_template_id CHAR(36),
        is_repeated BOOLEAN NOT NULL DEFAULT FALSE,
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        CONSTRAINT fk_templates_office
          FOREIGN KEY (office_id) REFERENCES notary_offices(id) ON DELETE CASCADE,
        CONSTRAINT fk_templates_category
          FOREIGN KEY (category_id) REFERENCES template_categories(id) ON DELETE SET NULL,
        CONSTRAINT fk_templates_parent
          FOREIGN KEY (parent_template_id) REFERENCES templates(id) ON DELETE SET NULL
      )
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS template_fields (
        id CHAR(36) PRIMARY KEY,
        template_id CHAR(36) NOT NULL,
        key_name VARCHAR(100) NOT NULL,
        field_type VARCHAR(20) NOT NULL DEFAULT 'text',
        label VARCHAR(255) NOT NULL,
        is_required BOOLEAN NOT NULL DEFAULT TRUE,
        order_index INTEGER DEFAULT 0,
        replace_text VARCHAR(500),
        paragraph_context TEXT,
        parent_field_key VARCHAR(100),
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        CONSTRAINT fk_template_fields_template
          FOREIGN KEY (template_id) REFERENCES templates(id) ON DELETE CASCADE,
        CONSTRAINT uk_template_fields_template_key UNIQUE (template_id, key_name)
      )
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS document_submissions (
        id CHAR(36) PRIMARY KEY,
        template_id CHAR(36) NOT NULL,
        customer_name VARCHAR(255),
        customer_phone VARCHAR(20),
        status VARCHAR(20) NOT NULL DEFAULT 'pending',
        values_json JSONB NOT NULL,
        output_file_path VARCHAR(500),
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        completed_at TIMESTAMPTZ,
        CONSTRAINT fk_document_submissions_template
          FOREIGN KEY (template_id) REFERENCES templates(id) ON DELETE CASCADE
      )
    `);

    const [templatesColumns] = await pool.query(`
      SELECT column_name
      FROM information_schema.columns
      WHERE table_schema = 'public' AND table_name = 'templates'
    `);
    const templatesColumnNames = templatesColumns.map(c => c.column_name);

    if (!templatesColumnNames.includes('parent_template_id')) {
      await pool.query('ALTER TABLE templates ADD COLUMN parent_template_id CHAR(36)');
      await pool.query(`
        ALTER TABLE templates
        ADD CONSTRAINT fk_templates_parent
        FOREIGN KEY (parent_template_id) REFERENCES templates(id) ON DELETE SET NULL
      `);
    }

    if (!templatesColumnNames.includes('is_repeated')) {
      await pool.query('ALTER TABLE templates ADD COLUMN is_repeated BOOLEAN NOT NULL DEFAULT FALSE');
    }

    if (!templatesColumnNames.includes('category_id')) {
      await pool.query('ALTER TABLE templates ADD COLUMN category_id CHAR(36)');
      await pool.query(`
        ALTER TABLE templates
        ADD CONSTRAINT fk_templates_category
        FOREIGN KEY (category_id) REFERENCES template_categories(id) ON DELETE SET NULL
      `);
    }

    const [fieldColumns] = await pool.query(`
      SELECT column_name
      FROM information_schema.columns
      WHERE table_schema = 'public' AND table_name = 'template_fields'
    `);
    const fieldColumnNames = fieldColumns.map(c => c.column_name);

    if (!fieldColumnNames.includes('replace_text')) {
      await pool.query('ALTER TABLE template_fields ADD COLUMN replace_text VARCHAR(500)');
    }

    if (!fieldColumnNames.includes('paragraph_context')) {
      await pool.query('ALTER TABLE template_fields ADD COLUMN paragraph_context TEXT');
    }

    if (!fieldColumnNames.includes('parent_field_key')) {
      await pool.query('ALTER TABLE template_fields ADD COLUMN parent_field_key VARCHAR(100)');
    }

    console.log('✅ Database migration hoàn tất trên PostgreSQL.');
  } catch (error) {
    console.error('❌ Lỗi khi chạy migration database:', error.message);
  }
}

// Khởi chạy server sau khi kiểm tra DB
async function startServer() {
  await testConnection();
  await migrateDatabase();
  await seedDefaultOffice();
  await seedDefaultTemplateCategories();
  
  app.listen(PORT, () => {
    console.log(`🚀 Backend server running at http://localhost:${PORT}`);
  });
}

startServer();

