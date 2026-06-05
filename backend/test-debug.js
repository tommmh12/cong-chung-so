const { pool } = require('./src/db');
const path = require('path');
const fs = require('fs');
const { mergeDocumentToBuffer } = require('./src/parser');

// Helper emulating backend preparation
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

async function testDownload() {
  try {
    console.log("=== FETCHING ALL SUBMISSIONS ===");
    const [rows] = await pool.query(
      `SELECT id, template_id, customer_name, values_json, status, completed_at FROM document_submissions`
    );

    for (const sub of rows) {
      try {
        console.log(`\n--- Submission: ${sub.id} | Template: ${sub.template_name} | Customer: ${sub.customer_name} ---`);
        let valuesJson = sub.values_json;
        if (typeof valuesJson === 'string') {
          try {
            valuesJson = JSON.parse(valuesJson);
          } catch (e) {
            console.error("Parse values_json failed:", e);
          }
        }
        const selectedChildIds = ['12c78301-f26c-488e-86af-20df502eb3f4'];
        const values = { ho_va_te: 'Đỗ Hoàng' };
        console.log("values:", values);
        console.log("selectedChildIds:", selectedChildIds);

        // Fetch parent template file path and name
        const [parentTemplates] = await pool.query(
          'SELECT name, file_path, id FROM templates WHERE id = $1',
          [sub.template_id]
        );
        if (parentTemplates.length === 0) {
          console.log("Parent template not found for submission:", sub.template_id);
          continue;
        }
        const parentTemplate = parentTemplates[0];

        // Let's run download simulation for this submission
        const [parentFields] = await pool.query(
          'SELECT key_name, is_required, replace_text, label FROM template_fields WHERE template_id = $1',
          [parentTemplate.id]
        );
        console.log(`Parent fields: ${parentFields.length}`);

        // Lấy child templates liên kết
        let [childTemplates] = await pool.query(
          "SELECT id, name, file_path, is_repeated FROM templates WHERE parent_template_id = $1 AND status = 'active'",
          [parentTemplate.id]
        );
        if (Array.isArray(selectedChildIds)) {
          childTemplates = childTemplates.filter(t => selectedChildIds.includes(t.id));
        }
        console.log(`Child templates: ${childTemplates.length}`);

        // Map repeated child records to parent suffix variables
        for (const child of childTemplates) {
          if (child.is_repeated) {
            const recordsList = values ? values[child.id] : null;
            if (Array.isArray(recordsList)) {
              const [childFields] = await pool.query(
                'SELECT key_name, parent_field_key FROM template_fields WHERE template_id = $1',
                [child.id]
              );
              recordsList.forEach((recordData, rIdx) => {
                childFields.forEach(cf => {
                  const targetKey = cf.parent_field_key || cf.key_name;
                  const sourceVal = recordData[cf.key_name];
                  if (sourceVal !== undefined && values) {
                    values[`${targetKey}_${rIdx + 1}`] = sourceVal;
                  }
                });
              });
            }
          }
        }

        console.log("Preparing master values...");
        const masterResult = prepareValuesForTemplate(parentFields, values || {}, parentTemplate.name);
        
        const absoluteTemplatePath = path.join(__dirname, 'src', '..', parentTemplate.file_path);
        console.log("absoluteTemplatePath:", absoluteTemplatePath);
        console.log("File exists?", fs.existsSync(absoluteTemplatePath));

        if (fs.existsSync(absoluteTemplatePath)) {
          try {
            const masterBuffer = mergeDocumentToBuffer(absoluteTemplatePath, masterResult.padded);
            console.log(`Master render succeeded! Buffer size: ${masterBuffer.length}`);
          } catch (e) {
            console.error("Master render failed:", e.message);
          }
        }

        // Children render
        for (const child of childTemplates) {
          const [childFields] = await pool.query(
            'SELECT key_name, is_required, replace_text, parent_field_key, label FROM template_fields WHERE template_id = $1',
            [child.id]
          );
          const absoluteChildTemplatePath = path.join(__dirname, 'src', '..', child.file_path);
          console.log(`- Child template: ${child.name} at ${child.file_path}. Exists? ${fs.existsSync(absoluteChildTemplatePath)}`);
          
          if (fs.existsSync(absoluteChildTemplatePath)) {
            try {
              if (child.is_repeated) {
                const recordsList = values ? values[child.id] : null;
                const recordsArray = Array.isArray(recordsList) ? recordsList : [{}];
                for (let rIdx = 0; rIdx < recordsArray.length; rIdx++) {
                  const prep = prepareValuesForSingleRecord(childFields, recordsArray[rIdx], values || {}, `${child.name} (Bản ghi ${rIdx + 1})`);
                  const childBuffer = mergeDocumentToBuffer(absoluteChildTemplatePath, prep.padded);
                  console.log(`  - Repeated child render ${rIdx + 1} succeeded! Buffer size: ${childBuffer.length}`);
                }
              } else {
                const prep = prepareValuesForTemplate(childFields, values || {}, child.name);
                const childBuffer = mergeDocumentToBuffer(absoluteChildTemplatePath, prep.padded);
                console.log(`  - Simple child render succeeded! Buffer size: ${childBuffer.length}`);
              }
            } catch (e) {
              console.error("  - Child render failed:", e.message);
            }
          }
        }
      } catch (subErr) {
        console.error(`Error processing submission ${sub.id}:`, subErr.message);
      }
    }
  } catch (err) {
    console.error("SIMULATION CRASHED WITH ERROR:", err);
  } finally {
    await pool.end();
  }
}

// Emulate prepareValuesForSingleRecord
const prepareValuesForSingleRecord = (fields, recordData, parentData, contextLabel = '') => {
  const padded = {};
  fields.forEach(f => {
    let rawVal = recordData[f.key_name];
    if (rawVal === undefined && f.parent_field_key) {
      rawVal = parentData[f.parent_field_key];
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
  });

  for (const key in recordData) {
    if (!(key in padded)) {
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
  return { padded };
};

testDownload();

