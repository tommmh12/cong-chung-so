const { v4: uuidv4 } = require('uuid');
const path = require('path');
const fs = require('fs');
const archiver = require('archiver');
const { pool } = require('../db');
const { mergeDocumentToBuffer } = require('../parser');

// Helper to prepare values for templates
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

// POST /api/submissions
async function createSubmission(req, res) {
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

    // Chỉ xử lý child templates đang active
    let [childTemplates] = await pool.query(
      'SELECT * FROM templates WHERE parent_template_id = ? AND status = \'active\'',
      [templateId]
    );

    if (selectedChildIds && selectedChildIds.length > 0) {
      childTemplates = childTemplates.filter(t => selectedChildIds.includes(t.id));
    } else {
      childTemplates = [];
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

    // Lưu kết quả nộp hồ sơ vào Database (Chỉ lưu JSON dữ liệu, KHÔNG sinh file vật lý)
    await pool.query(
      `INSERT INTO document_submissions (id, template_id, customer_name, customer_phone, status, values_json, output_file_path, completed_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        submissionId,
        templateId,
        customerName || 'Khách hàng vãng lai',
        customerPhone || null,
        submissionStatus,
        JSON.stringify({ values, selectedChildIds }),
        null,
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
}

// GET /api/submissions
async function getSubmissions(req, res) {
  try {
    const [rows] = await pool.query(`
      SELECT s.id, s.template_id, s.customer_name, s.customer_phone, s.status, s.created_at, s.completed_at, t.name as template_name, s.values_json
      FROM document_submissions s
      JOIN templates t ON s.template_id = t.id
      ORDER BY s.created_at DESC
    `);

    const results = [];
    for (const row of rows) {
      let values = {};
      let selectedChildIds = [];
      try {
        const parsed = typeof row.values_json === 'string' ? JSON.parse(row.values_json) : (row.values_json || {});
        values = parsed.values || {};
        selectedChildIds = parsed.selectedChildIds || [];
      } catch (e) {
        console.warn('Lỗi parse values_json:', e);
      }

      // Chỉ hiện child templates đang active trong danh sách hồ sơ
      let [childTemplates] = await pool.query(
        'SELECT id, name, is_repeated FROM templates WHERE parent_template_id = ? AND status = \'active\'',
        [row.template_id]
      );

      if (selectedChildIds.length > 0) {
        childTemplates = childTemplates.filter(t => selectedChildIds.includes(t.id));
      } else {
        childTemplates = [];
      }

      const fileNames = [`${row.template_name}.docx`];
      for (const child of childTemplates) {
        if (child.is_repeated) {
          const recordsList = values[child.id];
          const recordsArray = Array.isArray(recordsList) ? recordsList : [{}];
          for (let rIdx = 0; rIdx < recordsArray.length; rIdx++) {
            fileNames.push(`${child.name}_Căn_${rIdx + 1}.docx`);
          }
        } else {
          fileNames.push(`${child.name}.docx`);
        }
      }

      results.push({
        id: row.id,
        template_id: row.template_id,
        customer_name: row.customer_name,
        customer_phone: row.customer_phone,
        status: row.status,
        created_at: row.created_at,
        completed_at: row.completed_at,
        template_name: row.template_name,
        values_json: row.values_json,
        files: fileNames
      });
    }

    res.json(results);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Không thể tải lịch sử hồ sơ đã nhận.' });
  }
}

// GET /api/submissions/:id/download
async function downloadSubmission(req, res) {
  try {
    const submissionId = req.params.id;
    const [rows] = await pool.query(
      'SELECT s.values_json, t.name as template_name, t.id as template_id, t.storage_key FROM document_submissions s JOIN templates t ON s.template_id = t.id WHERE s.id = ?',
      [submissionId]
    );

    if (rows.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy hồ sơ.' });
    }

    const parsedVJ = typeof rows[0].values_json === 'string' ? JSON.parse(rows[0].values_json) : (rows[0].values_json || {});
    const values = parsedVJ.values || {};
    const selectedChildIds = parsedVJ.selectedChildIds || [];
    const parentTemplate = rows[0];

    // Chỉ tải child templates đang active
    let [childTemplates] = await pool.query(
      'SELECT id, name, storage_key, is_repeated FROM templates WHERE parent_template_id = ? AND status = \'active\'',
      [parentTemplate.template_id]
    );
    if (selectedChildIds.length > 0) {
      childTemplates = childTemplates.filter(t => selectedChildIds.includes(t.id));
    }

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

    const [parentFields] = await pool.query(
      'SELECT key_name, is_required, replace_text, label FROM template_fields WHERE template_id = ?',
      [parentTemplate.template_id]
    );
    const masterResult = prepareValuesForTemplate(parentFields, values, parentTemplate.template_name);
    const absoluteTemplatePath = path.join(__dirname, '..', '..', parentTemplate.storage_key.replace(/\\/g, '/'));

    const masterBuffer = mergeDocumentToBuffer(absoluteTemplatePath, masterResult.padded);
    const safeName = parentTemplate.template_name.replace(/[^a-zA-Z0-9À-ỹ\s-_]/g, '');

    if (childTemplates.length === 0) {
      res.setHeader('Content-Disposition', `attachment; filename=${encodeURIComponent(`${safeName}_HoanThinh.docx`)}`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      return res.send(masterBuffer);
    }

    res.setHeader('Content-Disposition', `attachment; filename=${encodeURIComponent(`${safeName}_HoanThinh.zip`)}`);
    res.setHeader('Content-Type', 'application/zip');

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.pipe(res);

    archive.append(masterBuffer, { name: `${parentTemplate.template_name}.docx` });

    for (const child of childTemplates) {
      const [childFields] = await pool.query(
        'SELECT key_name, is_required, replace_text, parent_field_key, label FROM template_fields WHERE template_id = ?',
        [child.id]
      );

      if (child.is_repeated) {
        const recordsList = values[child.id];
        const recordsArray = Array.isArray(recordsList) ? recordsList : [{}];
        for (let rIdx = 0; rIdx < recordsArray.length; rIdx++) {
          const prep = prepareValuesForSingleRecord(childFields, recordsArray[rIdx], values, `${child.name} (Bản ghi ${rIdx + 1})`);
          const absoluteChildTemplatePath = path.join(__dirname, '..', '..', child.storage_key.replace(/\\/g, '/'));
          const childBuffer = mergeDocumentToBuffer(absoluteChildTemplatePath, prep.padded);
          archive.append(childBuffer, { name: `${child.name}_Căn_${rIdx + 1}.docx` });
        }
      } else {
        const prep = prepareValuesForTemplate(childFields, values, child.name);
        const absoluteChildTemplatePath = path.join(__dirname, '..', '..', child.storage_key.replace(/\\/g, '/'));
        const childBuffer = mergeDocumentToBuffer(absoluteChildTemplatePath, prep.padded);
        archive.append(childBuffer, { name: `${child.name}.docx` });
      }
    }

    archive.finalize();
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Không thể tải trọn bộ hồ sơ.' });
  }
}

// GET /api/submissions/:id/files
async function getSubmissionFiles(req, res) {
  try {
    const submissionId = req.params.id;
    const [rows] = await pool.query(
      'SELECT s.values_json, t.name as template_name, t.id as template_id FROM document_submissions s JOIN templates t ON s.template_id = t.id WHERE s.id = ?',
      [submissionId]
    );

    if (rows.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy hồ sơ.' });
    }

    const parsedVJ2 = typeof rows[0].values_json === 'string' ? JSON.parse(rows[0].values_json) : (rows[0].values_json || {});
    const values = parsedVJ2.values || {};
    const selectedChildIds = parsedVJ2.selectedChildIds || [];
    const parentName = rows[0].template_name;

    let [childTemplates] = await pool.query(
      'SELECT id, name, is_repeated FROM templates WHERE parent_template_id = ? AND status = \'active\'',
      [rows[0].template_id]
    );

    if (selectedChildIds.length > 0) {
      childTemplates = childTemplates.filter(t => selectedChildIds.includes(t.id));
    } else {
      childTemplates = [];
    }

    if (childTemplates.length === 0) {
      return res.json([`${parentName}.docx`]);
    }

    const fileNames = [`${parentName}.docx`];
    for (const child of childTemplates) {
      if (child.is_repeated) {
        const recordsList = values[child.id];
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
    console.error(error);
    res.status(500).json({ error: 'Không thể liệt kê danh sách file.' });
  }
}

// GET /api/submissions/:id/download-file
async function downloadSubmissionFile(req, res) {
  try {
    const submissionId = req.params.id;
    const filename = (req.query.filename || '').normalize('NFC');
    if (!filename) {
      return res.status(400).json({ error: 'Thiếu tên file cần tải.' });
    }

    const [rows] = await pool.query(
      'SELECT s.values_json, t.name as template_name, t.id as template_id, t.storage_key FROM document_submissions s JOIN templates t ON s.template_id = t.id WHERE s.id = ?',
      [submissionId]
    );

    if (rows.length === 0) {
      return res.status(404).json({ error: 'Không tìm thấy hồ sơ.' });
    }

    const parsed = typeof rows[0].values_json === 'string' ? JSON.parse(rows[0].values_json) : (rows[0].values_json || {});
    const values = parsed.values || {};
    const selectedChildIds = parsed.selectedChildIds || [];
    const parentTemplate = rows[0];

    // 1. Kiểm tra nếu file yêu cầu chính là file master mẹ
    if (filename === `${parentTemplate.template_name}.docx`.normalize('NFC')) {
      let [childTemplatesForSuffix] = await pool.query(
        'SELECT id, is_repeated FROM templates WHERE parent_template_id = ?',
        [parentTemplate.template_id]
      );
      if (selectedChildIds.length > 0) {
        childTemplatesForSuffix = childTemplatesForSuffix.filter(t => selectedChildIds.includes(t.id));
      }
      for (const child of childTemplatesForSuffix) {
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

      const [parentFields] = await pool.query(
        'SELECT key_name, is_required, replace_text, label FROM template_fields WHERE template_id = ?',
        [parentTemplate.template_id]
      );
      const masterResult = prepareValuesForTemplate(parentFields, values, parentTemplate.template_name);
      const absoluteTemplatePath = path.join(__dirname, '..', '..', parentTemplate.storage_key.replace(/\\/g, '/'));
      const buffer = mergeDocumentToBuffer(absoluteTemplatePath, masterResult.padded);

      res.setHeader('Content-Disposition', `attachment; filename=${encodeURIComponent(filename)}`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      return res.send(buffer);
    }

    // 2. Kiểm tra các file con (chỉ active)
    let [childTemplates] = await pool.query(
      'SELECT id, name, storage_key, is_repeated FROM templates WHERE parent_template_id = ? AND status = \'active\'',
      [parentTemplate.template_id]
    );
    if (selectedChildIds.length > 0) {
      childTemplates = childTemplates.filter(t => selectedChildIds.includes(t.id));
    }

    for (const child of childTemplates) {
      if (child.is_repeated) {
        const recordsList = values[child.id];
        const recordsArray = Array.isArray(recordsList) ? recordsList : [{}];
        for (let rIdx = 0; rIdx < recordsArray.length; rIdx++) {
          if (filename === `${child.name}_Căn_${rIdx + 1}.docx`.normalize('NFC')) {
            const [childFields] = await pool.query(
              'SELECT key_name, is_required, replace_text, parent_field_key, label FROM template_fields WHERE template_id = ?',
              [child.id]
            );
            const prep = prepareValuesForSingleRecord(childFields, recordsArray[rIdx], values, `${child.name} (Bản ghi ${rIdx + 1})`);
            const absoluteChildTemplatePath = path.join(__dirname, '..', '..', child.storage_key.replace(/\\/g, '/'));
            const buffer = mergeDocumentToBuffer(absoluteChildTemplatePath, prep.padded);

            res.setHeader('Content-Disposition', `attachment; filename=${encodeURIComponent(filename)}`);
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            return res.send(buffer);
          }
        }
      } else {
        if (filename === `${child.name}.docx`.normalize('NFC')) {
          const [childFields] = await pool.query(
            'SELECT key_name, is_required, replace_text, parent_field_key, label FROM template_fields WHERE template_id = ?',
            [child.id]
          );
          const prep = prepareValuesForTemplate(childFields, values, child.name);
          const absoluteChildTemplatePath = path.join(__dirname, '..', '..', child.storage_key.replace(/\\/g, '/'));
          const buffer = mergeDocumentToBuffer(absoluteChildTemplatePath, prep.padded);

          res.setHeader('Content-Disposition', `attachment; filename=${encodeURIComponent(filename)}`);
          res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
          return res.send(buffer);
        }
      }
    }

    return res.status(404).json({ error: 'Không tìm thấy file tương ứng trong hồ sơ.' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Không thể tải file lẻ.' });
  }
}

module.exports = {
  createSubmission,
  getSubmissions,
  downloadSubmission,
  getSubmissionFiles,
  downloadSubmissionFile
};
