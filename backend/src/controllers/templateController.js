const { v4: uuidv4 } = require('uuid');
const path = require('path');
const fs = require('fs');
const archiver = require('archiver');
const PizZip = require('pizzip');
const { pool } = require('../db');
const {
  scanPlaceholders,
  restorePlaceholder,
  scanTables,
  injectTablePlaceholders,
  injectPlaceholders
} = require('../parser');
const { WORD_EXTENSIONS, ensureDocxForTemplate } = require('../services/document-conversion');
const storage = require('../services/storage');

const DEFAULT_OFFICE_ID = 'd3b07384-d113-4ec6-a5d6-c0c2a05d2ed1';
const BACKEND_ROOT = path.join(__dirname, '..', '..');

// Resolve storageKey to a local path for processing.
// For Supabase paths: downloads to a temp file, returns cleanup fn.
// For legacy local paths: returns abs path with no-op cleanup.
async function resolveToLocal(storageKey) {
  if (storage.isSupabasePath(storageKey)) {
    const { tmpPath, cleanup } = await storage.downloadToTemp(storageKey);
    return { localPath: tmpPath, cleanup };
  }
  const localPath = path.join(BACKEND_ROOT, storageKey.replace(/\\/g, '/'));
  return { localPath, cleanup: () => {} };
}

// Like resolveToLocal but also re-uploads the (possibly modified) temp file on finish.
async function resolveForModify(storageKey, fn) {
  if (storage.isSupabasePath(storageKey)) {
    const { tmpPath, cleanup } = await storage.downloadToTemp(storageKey);
    try {
      fn(tmpPath);
      await storage.uploadFile(tmpPath, storageKey);
    } finally {
      cleanup();
    }
  } else {
    const localPath = path.join(BACKEND_ROOT, storageKey.replace(/\\/g, '/'));
    fn(localPath);
  }
}

// Delete file from Supabase or local disk.
function deleteStorageFile(storageKey) {
  if (storage.isSupabasePath(storageKey)) {
    return storage.deleteFile(storageKey);
  }
  const absPath = path.join(BACKEND_ROOT, storageKey.replace(/\\/g, '/'));
  cleanupFileIfExists(absPath);
  return Promise.resolve();
}

// Helpers
function cleanupFileIfExists(filePath) {
  if (filePath && fs.existsSync(filePath)) {
    fs.unlinkSync(filePath);
  }
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

function slugifyVietnamese(str) {
  const from = "ÁÀÂÄÃÅáàâäãåÉÈÊËéèêëÍÌÎÏíìîïÓÒÔÖÕØóòôöõøÚÙÛÜúùûüÑñÇçÝŸýÿŽžšđâăêôơưĂÂÊÔƠƯĐ";
  const to   = "AAAAAAaaaaaaEEEEeeeeIIIIiiiiOOOOOOooooooUUUUuuuuNnCcYYyyZzsd aaeoo uAAE OOUU D";
  const mapping = {};
  for (let i = 0; i < from.length; i++) mapping[from[i]] = to[i];
  return str
    .split('')
    .map(ch => mapping[ch] || ch)
    .join('')
    .replace(/[^\w\s-]/g, '')
    .trim()
    .replace(/\s+/g, '_')
    .toLowerCase();
}

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

async function createTemplateFromUploadedFile(file, options = {}) {
  const normalizedOriginalName = normalizeVietnameseFileName(file.originalname);
  const originalExt = path.extname(normalizedOriginalName).toLowerCase();
  const templateId = uuidv4();
  const templateName = options.templateName || path.basename(normalizedOriginalName, originalExt);
  const categoryId = options.categoryId || null;
  const officeId = options.officeId || DEFAULT_OFFICE_ID;
  const templatesDir = path.dirname(file.path);

  let sourceDocxPath = file.path;
  let cleanupConvertedFile = () => {};

  try {
    if (originalExt === '.doc') {
      const conversion = await ensureDocxForTemplate(file.path, {
        templateId,
        outputDir: templatesDir
      });
      sourceDocxPath = conversion.outputPath;
      cleanupConvertedFile = conversion.cleanup;
      cleanupFileIfExists(file.path);
    }

    const finalFileName = `${templateId}.docx`;
    const finalPath = path.join(templatesDir, finalFileName);
    fs.renameSync(sourceDocxPath, finalPath);
    cleanupConvertedFile();

    let variables = [];
    try {
      variables = scanPlaceholders(finalPath);
    } catch (parseErr) {
      cleanupFileIfExists(finalPath);
      throw parseErr;
    }

    // Upload to Supabase Storage and remove local copy
    const supabaseKey = `templates/${templateId}.docx`;
    await storage.uploadFile(finalPath, supabaseKey);
    cleanupFileIfExists(finalPath);

    const conn = await pool.getConnection();
    try {
      await conn.beginTransaction();

      await conn.query(
        'INSERT INTO templates (id, office_id, category_id, name, storage_key, status) VALUES (?, ?, ?, ?, ?, ?)',
        [templateId, officeId, categoryId, templateName, supabaseKey, 'draft']
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

// Controller functions

async function getTemplates(req, res) {
  try {
    const [rows] = await pool.query(`
      SELECT t.id, t.name, t.storage_key, t.status, t.created_at, t.parent_template_id, t.is_repeated,
             t.category_id, c.name AS category_name, c.parent_id AS category_parent_id,
             (SELECT COUNT(*) FROM template_fields WHERE template_id = t.id) as fields_count,
             (SELECT COUNT(*) FROM templates WHERE parent_template_id = t.id) as children_count
      FROM templates t
      LEFT JOIN template_categories c ON c.id = t.category_id
      WHERE t.parent_template_id IS NULL
      ORDER BY t.created_at DESC
    `);
    res.json(rows);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Không thể tải danh sách biểu mẫu.' });
  }
}

async function uploadTemplates(req, res) {
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
}

async function getTemplateForm(req, res) {
  try {
    const templateId = req.params.id;
    const [templates] = await pool.query('SELECT * FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu yêu cầu.' });
    }
    const [fields] = await pool.query(
      'SELECT id, key_name, field_type, label, is_required, order_index, replace_text, paragraph_context, parent_field_key, occurrence_index FROM template_fields WHERE template_id = ? ORDER BY order_index ASC',
      [templateId]
    );
    res.json({ template: templates[0], fields });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Lỗi tải cấu hình form.' });
  }
}

async function getTemplateTables(req, res) {
  try {
    const templateId = req.params.id;
    const [templates] = await pool.query('SELECT * FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu.' });
    }
    const template = templates[0];
    const { localPath: absolutePath, cleanup } = await resolveToLocal(template.storage_key);
    let tables;
    try {
      tables = scanTables(absolutePath);
    } finally {
      cleanup();
    }
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
}

async function injectTemplateTable(req, res) {
  try {
    const templateId = req.params.id;
    const { tableIndex, fields, selectedRows } = req.body;
    if (typeof tableIndex !== 'number' || !Array.isArray(fields) || !Array.isArray(selectedRows)) {
      return res.status(400).json({ error: 'Dữ liệu yêu cầu không hợp lệ.' });
    }
    const [templates] = await pool.query('SELECT * FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu.' });
    }
    const template = templates[0];
    const { localPath: absolutePath, cleanup: cleanupLocal } = await resolveToLocal(template.storage_key);
    let targetTable;
    try {
      const tables = scanTables(absolutePath);
      targetTable = tables.find(t => t.tableIndex === tableIndex);
    } finally {
      cleanupLocal();
    }
    if (!targetTable) {
      return res.status(404).json({ error: 'Bảng mục tiêu không tồn tại.' });
    }
    try {
      await resolveForModify(template.storage_key, localPath => {
        injectTablePlaceholders(localPath, localPath, tableIndex, fields, selectedRows);
      });
    } catch (injectErr) {
      console.error('Inject error:', injectErr);
      return res.status(500).json({ error: 'Không thể gài placeholder vào bảng của file.' });
    }
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
}

async function updateTemplateFields(req, res) {
  try {
    const templateId = req.params.id;
    const { fields } = req.body;

    if (!Array.isArray(fields)) {
      return res.status(400).json({ error: 'Cấu trúc trường dữ liệu gửi lên không hợp lệ.' });
    }

    const [templates] = await pool.query('SELECT * FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Biểu mẫu không tồn tại trên hệ thống.' });
    }
    const { storage_key } = templates[0];

    const replacements = fields
      .filter(f => f.replace_text && typeof f.replace_text === 'string' && f.replace_text.trim() !== '')
      .map(f => ({
        searchText: f.replace_text.trim(),
        key_name: f.key_name,
        paragraph_context: f.paragraph_context,
        occurrence_index: f.occurrence_index ?? 0
      }));

    if (replacements.length > 0) {
      try {
        await resolveForModify(storage_key, localPath => {
          injectPlaceholders(localPath, localPath, replacements);
        });
      } catch (err) {
        console.error('Lỗi khi ghi đè biến động vào docx:', err);
        return res.status(500).json({ error: 'Không thể ghi đè biến trực quan vào file Word: ' + err.message });
      }
    }

    const conn = await pool.getConnection();
    try {
      await conn.beginTransaction();

      await conn.query('DELETE FROM template_fields WHERE template_id = ?', [templateId]);

      for (const field of fields) {
        const fieldId = (field.id && !field.id.startsWith('temp-')) ? field.id : uuidv4();
        await conn.query(
          `INSERT INTO template_fields (id, template_id, key_name, field_type, label, is_required, order_index, replace_text, paragraph_context, parent_field_key, occurrence_index)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
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
            field.parent_field_key || null,
            field.occurrence_index ?? 0
          ]
        );
      }

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
    console.error('updateTemplateFields error:', error);
    const msg = error.message || '';
    if (msg.includes('too long') || msg.includes('value too long') || msg.includes('Data too long')) {
      return res.status(400).json({ error: 'Một trường có dữ liệu quá dài. Vui lòng rút ngắn nhãn hiển thị hoặc chọn đoạn văn ngắn hơn khi tạo biến.' });
    }
    if (msg.includes('duplicate') || msg.includes('Duplicate') || msg.includes('unique') || msg.includes('UNIQUE')) {
      return res.status(400).json({ error: 'Có hai trường trùng mã biến (key_name). Vui lòng kiểm tra lại danh sách field.' });
    }
    res.status(500).json({ error: 'Đã xảy ra lỗi khi lưu cấu hình form: ' + msg });
  }
}

async function getTemplateLinks(req, res) {
  try {
    const { id } = req.params;
    const [rows] = await pool.query(
      'SELECT id, name, storage_key, status, created_at, is_repeated FROM templates WHERE parent_template_id = ?',
      [id]
    );
    res.json(rows);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Không thể tải danh sách biểu mẫu con liên kết.' });
  }
}

async function linkTemplate(req, res) {
  try {
    const parentId = req.params.id;
    const { childTemplateId } = req.body;
    if (!childTemplateId) {
      return res.status(400).json({ error: 'Vui lòng cung cấp ID biểu mẫu con cần liên kết.' });
    }

    if (parentId === childTemplateId) {
      return res.status(400).json({ error: 'Không thể liên kết biểu mẫu với chính nó.' });
    }

    const [parentRows] = await pool.query(
      'SELECT id, parent_template_id FROM templates WHERE id = ?',
      [parentId]
    );
    if (parentRows.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu gốc.' });
    }

    const [childRows] = await pool.query(
      'SELECT id, parent_template_id FROM templates WHERE id = ?',
      [childTemplateId]
    );
    if (childRows.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu con cần liên kết.' });
    }
    const child = childRows[0];

    if (child.parent_template_id && child.parent_template_id !== parentId) {
      return res.status(409).json({
        error: 'Biểu mẫu này đã được liên kết với một biểu mẫu gốc khác. Vui lòng hủy liên kết cũ trước.'
      });
    }
    if (child.parent_template_id === parentId) {
      return res.status(409).json({ error: 'Biểu mẫu con này đã được liên kết với biểu mẫu gốc rồi.' });
    }

    if (parentRows[0].parent_template_id) {
      return res.status(409).json({
        error: 'Biểu mẫu gốc này đang là file con của một biểu mẫu khác. Hệ thống chỉ hỗ trợ liên kết 2 cấp (Gốc → Con).'
      });
    }

    const [grandChildren] = await pool.query(
      'SELECT COUNT(*) AS cnt FROM templates WHERE parent_template_id = ?',
      [childTemplateId]
    );
    if (grandChildren[0].cnt > 0) {
      return res.status(409).json({
        error: 'Biểu mẫu này đang là file gốc của các biểu mẫu con khác nên không thể trở thành file con. Hệ thống chỉ hỗ trợ liên kết 2 cấp.'
      });
    }

    let cursor = parentRows[0].parent_template_id;
    const visited = new Set([parentId]);
    while (cursor) {
      if (cursor === childTemplateId) {
        return res.status(409).json({ error: 'Liên kết này tạo ra vòng lặp giữa các biểu mẫu. Thao tác bị từ chối.' });
      }
      if (visited.has(cursor)) break;
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
}

async function unlinkTemplate(req, res) {
  try {
    const { id } = req.params;
    await pool.query('UPDATE templates SET parent_template_id = NULL WHERE id = ?', [id]);
    res.json({ message: 'Đã hủy liên kết biểu mẫu thành công!' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Lỗi hủy liên kết biểu mẫu.' });
  }
}

async function getParentFields(req, res) {
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
}

async function restoreTemplateField(req, res) {
  try {
    const { templateId, fieldId } = req.params;
    const [fields] = await pool.query(
      'SELECT key_name, replace_text FROM template_fields WHERE id = ? AND template_id = ?',
      [fieldId, templateId]
    );

    if (fields.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy cấu hình biến cần khôi phục.' });
    }

    const { key_name, replace_text } = fields[0];
    
    if (!replace_text) {
      await pool.query('DELETE FROM template_fields WHERE id = ?', [fieldId]);
      return res.json({ message: `Đã xóa cấu hình biến thủ công {{${key_name}}} thành công!` });
    }

    const [templates] = await pool.query('SELECT storage_key FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu tương ứng.' });
    }

    const { storage_key } = templates[0];

    try {
      const srcBuffer = await storage.getFileBuffer(storage_key, BACKEND_ROOT);
      const zip = new PizZip(srcBuffer);
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
        if (storage.isSupabasePath(storage_key)) {
          await storage.uploadBuffer(buffer, storage_key);
        } else {
          fs.writeFileSync(path.join(BACKEND_ROOT, storage_key.replace(/\\/g, '/')), buffer);
        }
      }
    } catch (err) {
      console.error('Lỗi khi khôi phục biến trong file docx:', err);
      return res.status(500).json({ error: 'Không thể ghi đè khôi phục tệp Word: ' + err.message });
    }

    await pool.query('DELETE FROM template_fields WHERE id = ?', [fieldId]);
    res.json({ message: `Đã khôi phục biến {{${key_name}}} thành văn bản gốc thành công!` });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Lỗi khôi phục biến.' });
  }
}

async function downloadOriginalTemplate(req, res) {
  try {
    const templateId = req.params.id;
    const [rows] = await pool.query('SELECT storage_key, name FROM templates WHERE id = ?', [templateId]);

    if (rows.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy file mẫu.' });
    }

    const { storage_key, name } = rows[0];
    try {
      const buffer = await storage.getFileBuffer(storage_key, BACKEND_ROOT);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(name)}.docx`);
      res.send(buffer);
    } catch (err) {
      return res.status(404).json({ error: 'File mẫu gốc không tồn tại trên hệ thống.' });
    }
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Không thể tải file mẫu.' });
  }
}

async function deleteTemplate(req, res) {
  try {
    const templateId = req.params.id;
    const [templates] = await pool.query('SELECT storage_key FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu cần xóa.' });
    }
    const template = templates[0];

    // Collect children before deleting (ON DELETE SET NULL would orphan them)
    const [children] = await pool.query(
      'SELECT id, storage_key FROM templates WHERE parent_template_id = ?',
      [templateId]
    );

    const [submissions] = await pool.query(
      'SELECT output_file_path FROM document_submissions WHERE template_id = ?',
      [templateId]
    );

    // Delete children first (their own fields/submissions cascade via FK)
    for (const child of children) {
      await pool.query('DELETE FROM templates WHERE id = ?', [child.id]);
      await deleteStorageFile(child.storage_key);
    }

    await pool.query('DELETE FROM templates WHERE id = ?', [templateId]);
    await deleteStorageFile(template.storage_key);

    for (const sub of submissions) {
      if (sub.output_file_path) {
        const absoluteOutputPath = path.join(BACKEND_ROOT, sub.output_file_path.replace(/\\/g, '/'));
        if (fs.existsSync(absoluteOutputPath)) {
          try { fs.unlinkSync(absoluteOutputPath); } catch (err) {
            console.error(`Không thể xóa file output ${absoluteOutputPath}:`, err);
          }
        }
      }
    }

    res.json({ message: 'Đã xóa biểu mẫu và toàn bộ dữ liệu liên quan thành công!' });
  } catch (error) {
    console.error('Lỗi khi xóa biểu mẫu:', error);
    res.status(500).json({ error: 'Đã xảy ra lỗi trên hệ thống khi xóa biểu mẫu.' });
  }
}

async function updateTemplateName(req, res) {
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
}

async function updateTemplateCategory(req, res) {
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
}

async function toggleTemplateRepeated(req, res) {
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
}

async function toggleTemplateStatus(req, res) {
  try {
    const templateId = req.params.id;
    const { status } = req.body;

    if (status !== 'active' && status !== 'draft') {
      return res.status(400).json({ error: 'Trạng thái không hợp lệ.' });
    }

    await pool.query(
      'UPDATE templates SET status = ? WHERE id = ?',
      [status, templateId]
    );

    res.json({ message: 'Đã cập nhật trạng thái biểu mẫu thành công!' });
  } catch (error) {
    console.error('Lỗi khi cập nhật trạng thái:', error);
    res.status(500).json({ error: 'Đã xảy ra lỗi trên hệ thống khi cập nhật trạng thái.' });
  }
}

async function duplicateTemplate(req, res) {
  const conn = await pool.getConnection();
  try {
    const originalTemplateId = req.params.id;
    const [templates] = await conn.query('SELECT * FROM templates WHERE id = ?', [originalTemplateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu gốc cần sao chép.' });
    }
    const originalTemplate = templates[0];
    const newTemplateId = uuidv4();
    const newSupabaseKey = `templates/${newTemplateId}.docx`;

    let fileBuffer;
    try {
      fileBuffer = await storage.getFileBuffer(originalTemplate.storage_key, BACKEND_ROOT);
    } catch (err) {
      return res.status(404).json({ error: 'Tệp tin Word của biểu mẫu gốc không tồn tại.' });
    }
    await storage.uploadBuffer(fileBuffer, newSupabaseKey);

    await conn.beginTransaction();

    const newName = `[Sao chép] ${originalTemplate.name}`;
    await conn.query(
      'INSERT INTO templates (id, office_id, category_id, name, storage_key, status, parent_template_id) VALUES (?, ?, ?, ?, ?, ?, ?)',
      [newTemplateId, originalTemplate.office_id, originalTemplate.category_id || null, newName, newSupabaseKey, 'draft', null]
    );

    const [fields] = await conn.query(
      'SELECT key_name, field_type, label, is_required, order_index, replace_text, paragraph_context FROM template_fields WHERE template_id = ?',
      [originalTemplateId]
    );

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
}

async function exportTemplateFields(req, res) {
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
}

async function importTemplateFields(req, res) {
  const conn = await pool.getConnection();
  try {
    const templateId = req.params.id;
    const { fields } = req.body;

    if (!Array.isArray(fields)) {
      return res.status(400).json({ error: 'Dữ liệu cấu hình trường nhập vào không hợp lệ (yêu cầu mảng).' });
    }

    const [templates] = await conn.query('SELECT id FROM templates WHERE id = ?', [templateId]);
    if (templates.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy biểu mẫu đích để nhập cấu hình.' });
    }

    await conn.beginTransaction();
    await conn.query('DELETE FROM template_fields WHERE template_id = ?', [templateId]);

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
}

async function uploadChildTemplates(req, res) {
  const { id: parentId } = req.params;
  try {
    const [parentRows] = await pool.query('SELECT id, name FROM templates WHERE id = ?', [parentId]);
    if (!parentRows.length) return res.status(404).json({ error: 'Template cha không tồn tại' });

    const files = [...(req.files?.templateFile || []), ...(req.files?.templateFiles || [])];
    if (!files.length) return res.status(400).json({ error: 'Không có file nào được tải lên' });

    const inserted = [];
    const failed = [];
    for (const file of files) {
      const normalizedOriginal = normalizeVietnameseFileName(file.originalname);
      const baseName = path.basename(normalizedOriginal, path.extname(normalizedOriginal));
      try {
        const templatesDir = path.dirname(file.path);
        const id = uuidv4();
        let sourceDocxPath = file.path;
        let cleanupConverted = () => {};

        // Convert .doc → .docx if needed
        const ext = path.extname(file.path).toLowerCase();
        if (ext === '.doc') {
          const conversion = await ensureDocxForTemplate(file.path, { templateId: id, outputDir: templatesDir });
          sourceDocxPath = conversion.outputPath;
          cleanupConverted = conversion.cleanup;
          cleanupFileIfExists(file.path);
        }

        const finalDocxPath = path.join(templatesDir, `${id}.docx`);
        fs.renameSync(sourceDocxPath, finalDocxPath);
        cleanupConverted();

        // Scan placeholders while file is still local
        const variables = scanPlaceholders(finalDocxPath);

        // Upload to Supabase and remove local copy
        const storageKey = `templates/${id}.docx`;
        await storage.uploadFile(finalDocxPath, storageKey);
        cleanupFileIfExists(finalDocxPath);

        const conn = await pool.getConnection();
        try {
          await conn.beginTransaction();
          await conn.query(
            `INSERT INTO templates (id, office_id, name, storage_key, status, parent_template_id) VALUES (?, ?, ?, ?, 'active', ?)`,
            [id, DEFAULT_OFFICE_ID, baseName, storageKey, parentId]
          );
          for (let i = 0; i < variables.length; i++) {
            const key = variables[i];
            const defaultLabel = key.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
            await conn.query(
              `INSERT INTO template_fields (id, template_id, key_name, field_type, label, is_required, order_index) VALUES (?, ?, ?, 'text', ?, true, ?)`,
              [uuidv4(), id, key, defaultLabel, i]
            );
          }
          await conn.commit();
        } catch (dbErr) {
          await conn.rollback();
          // File already uploaded; clean up from Supabase on DB failure
          await storage.deleteFile(storageKey).catch(() => {});
          throw dbErr;
        } finally {
          conn.release();
        }

        inserted.push({ id, name: baseName, status: 'draft', is_repeated: false, variablesCount: variables.length });
      } catch (err) {
        cleanupFileIfExists(file.path);
        failed.push({ name: baseName, error: err.message });
      }
    }
    res.json({
      message: `Đã thêm ${inserted.length} phụ lục${failed.length ? `, ${failed.length} lỗi` : ''}`,
      templates: inserted,
      failed
    });
  } catch (error) {
    console.error('uploadChildTemplates:', error);
    res.status(500).json({ error: 'Không thể thêm phụ lục: ' + error.message });
  }
}

async function downloadAllAsZip(req, res) {
  const { id } = req.params;
  try {
    const [parentRows] = await pool.query('SELECT * FROM templates WHERE id = ?', [id]);
    if (!parentRows.length) return res.status(404).json({ error: 'Template không tồn tại' });
    const parent = parentRows[0];

    const [children] = await pool.query(
      'SELECT * FROM templates WHERE parent_template_id = ?', [id]
    );

    const all = [parent, ...children];

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(parent.name)}.zip`);

    const archive = archiver('zip', { zlib: { level: 6 } });
    archive.on('error', err => { console.error(err); res.status(500).end(); });
    archive.pipe(res);

    const usedNames = {};
    for (const tmpl of all) {
      try {
        const buffer = await storage.getFileBuffer(tmpl.storage_key, BACKEND_ROOT);
        const ext = path.extname(tmpl.storage_key) || '.docx';
        let fileName = `${tmpl.name}${ext}`;
        if (usedNames[fileName]) { usedNames[fileName]++; fileName = `${tmpl.name}_${usedNames[fileName]}${ext}`; }
        else usedNames[fileName] = 1;
        archive.append(buffer, { name: fileName });
      } catch (err) {
        console.error(`Skip ${tmpl.name} in ZIP:`, err.message);
      }
    }

    await archive.finalize();
  } catch (error) {
    console.error('downloadAllAsZip:', error);
    if (!res.headersSent) res.status(500).json({ error: 'Không thể tạo file ZIP' });
  }
}

module.exports = {
  getTemplates,
  uploadTemplates,
  getTemplateForm,
  getTemplateTables,
  injectTemplateTable,
  updateTemplateFields,
  getTemplateLinks,
  linkTemplate,
  unlinkTemplate,
  getParentFields,
  restoreTemplateField,
  downloadOriginalTemplate,
  deleteTemplate,
  updateTemplateName,
  updateTemplateCategory,
  toggleTemplateRepeated,
  toggleTemplateStatus,
  duplicateTemplate,
  exportTemplateFields,
  importTemplateFields,
  uploadChildTemplates,
  downloadAllAsZip
};
