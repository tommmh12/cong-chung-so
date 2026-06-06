const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");

function unescapeXml(safe) {
  if (!safe) return "";
  return safe
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#160;/g, ' ')
    .replace(/&nbsp;/g, ' ');
}

function getParagraphText(pXml) {
  const elementRegex = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\b[^>]*>|<w:br\b[^>]*>|<w:cr\b[^>]*>/g;
  let text = "";
  let match;
  while ((match = elementRegex.exec(pXml)) !== null) {
    if (/^<w:t[>\s]/.test(match[0])) {
      const textVal = match[1] ? match[1].replace(/<[^>]+>/g, "") : "";
      text += unescapeXml(textVal);
    } else {
      text += " ";
    }
  }
  return text;
}

/**
 * Quét toàn bộ file .docx để lấy danh sách biến động dạng {{ten_bien}}
 * @param {string} filePath - Đường dẫn đến file docx gốc
 * @returns {Array<string>} Danh sách tên biến không trùng lặp
 */
function scanPlaceholders(filePath) {
  try {
    const content = fs.readFileSync(filePath, "binary");
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: {
        start: '{{',
        end: '}}'
      }
    });

    const text = doc.getFullText();
    // Regex quét các cụm ký tự nằm trong {{ }}
    // Hỗ trợ ký tự tiếng Việt không dấu, chữ, số và dấu gạch dưới
    const regex = /\{\{([a-zA-Z0-9_]+)\}\}/g;
    const variables = new Set();
    let match;

    while ((match = regex.exec(text)) !== null) {
      variables.add(match[1]);
    }
    
    return Array.from(variables);
  } catch (error) {
    console.error("Lỗi khi quét file DOCX:", error);
    if (error.message && error.message.includes('filetype for this file could not be identified')) {
      throw new Error('File không phải là tài liệu Word (.docx) hợp lệ. Vui lòng kiểm tra lại file — có thể là file theme (.thmx), file hỏng, hoặc file không phải định dạng Word.');
    }
    throw new Error("Không thể đọc và phân tích file Word. Vui lòng kiểm tra định dạng file.");
  }
}

function collapseConsecutiveSpacesAcrossTags(xmlContent) {
  const parts = [];
  const tRegex = /(<w:t\b[^>]*>)([\s\S]*?)(<\/w:t>)/g;
  let lastIndex = 0;
  let match;
  
  // Phân tách XML thành các thẻ <w:t> và các đoạn text/xml ở giữa chúng
  while ((match = tRegex.exec(xmlContent)) !== null) {
    parts.push({
      type: 'raw',
      content: xmlContent.substring(lastIndex, match.index)
    });
    
    parts.push({
      type: 't',
      openTag: match[1],
      text: match[2].replace(/ {2,}/g, ' '), // Thu gọn khoảng trắng kép trong cùng 1 tag
      closeTag: match[3]
    });
    
    lastIndex = tRegex.lastIndex;
  }
  
  parts.push({
    type: 'raw',
    content: xmlContent.substring(lastIndex)
  });
  
  // Thu gọn khoảng trắng thừa chuyển giao giữa các thẻ <w:t> liên tiếp
  let lastTPart = null;
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    if (part.type === 't') {
      if (lastTPart) {
        const lastText = lastTPart.text;
        const currentText = part.text;
        
        // Nếu tag trước kết thúc bằng dấu cách và tag này bắt đầu bằng dấu cách
        if (lastText.endsWith(' ') && currentText.startsWith(' ')) {
          part.text = currentText.substring(1); // Xóa dấu cách ở đầu tag hiện tại
        }
      }
      lastTPart = part;
    } else if (part.type === 'raw') {
      // Reset trạng thái tag liên tiếp nếu đi qua các phần tử ngắt dòng hoặc đoạn văn mới
      const raw = part.content;
      if (raw.includes('</w:p>') || raw.includes('<w:p') || raw.includes('<w:br') || raw.includes('<w:tab')) {
        lastTPart = null;
      }
    }
  }
  
  // Dựng lại toàn bộ chuỗi XML
  return parts.map(part => {
    if (part.type === 't') {
      let openTag = part.openTag;
      // Tự động gài xml:space="preserve" để Word giữ khoảng trắng
      if ((part.text.startsWith(' ') || part.text.endsWith(' ')) && !openTag.includes('xml:space=')) {
        openTag = openTag.replace('<w:t', '<w:t xml:space="preserve"');
      }
      return `${openTag}${part.text}${part.closeTag}`;
    }
    return part.content;
  }).join('');
}

/**
 * Ghi đè dữ liệu JSON phẳng vào template file .docx để xuất ra file kết quả
 * @param {string} templatePath - Đường dẫn file mẫu gốc
 * @param {string} outputPath - Đường dẫn lưu file .docx kết quả
 * @param {Object} dataJson - Dữ liệu dạng phẳng { key: value } để map vào các placeholder
 */
function mergeDocument(templatePath, outputPath, dataJson) {
  try {
    // Đảm bảo thư mục đầu ra tồn tại
    const outputDir = path.dirname(outputPath);
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);
    
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: {
        start: '{{',
        end: '}}'
      }
    });
    
    // Ghi dữ liệu JSON phẳng trực tiếp vào các biến tương ứng
    doc.render(dataJson);
    
    // Thu gọn nhiều dấu cách liên tiếp trong các thẻ <w:t> ở toàn bộ các file XML của Word
    const zipObj = doc.getZip();
    zipObj.file(/word\/.*\.xml/).forEach((file) => {
      const xmlContent = file.asText();
      const newXmlContent = collapseConsecutiveSpacesAcrossTags(xmlContent);
      zipObj.file(file.name, newXmlContent);
    });
    
    const buf = zipObj.generate({
      type: "nodebuffer",
      compression: "DEFLATE",
    });
    
    fs.writeFileSync(outputPath, buf);
    console.log(`✅ Xuất file thành công: ${outputPath}`);
  } catch (error) {
    console.error("Lỗi khi map dữ liệu vào DOCX:", error);
    throw new Error("Lỗi trong quá trình render dữ liệu vào file Word: " + error.message);
  }
}

function mergeDocumentToBuffer(templatePath, dataJson) {
  try {
    const content = fs.readFileSync(templatePath, "binary");
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
    
    const zipObj = doc.getZip();
    zipObj.file(/word\/.*\.xml/).forEach((file) => {
      const xmlContent = file.asText();
      const newXmlContent = collapseConsecutiveSpacesAcrossTags(xmlContent);
      zipObj.file(file.name, newXmlContent);
    });
    
    return zipObj.generate({
      type: "nodebuffer",
      compression: "DEFLATE",
    });
  } catch (error) {
    console.error("Lỗi khi map dữ liệu vào DOCX thành buffer:", error);
    throw new Error("Lỗi render dữ liệu: " + error.message);
  }
}

/**
 * Thay thế các chuỗi văn bản tĩnh trong file .docx bằng các thẻ biến {{placeholder}}
 * @param {string} filePath - Đường dẫn file gốc
 * @param {string} outputPath - Đường dẫn file kết quả sau khi gài biến
 * @param {Array<{searchText: string, key_name: string}>} replacements - Danh sách các cụm từ cần thay thế thành biến
 */
/**
 * Thay thế các chuỗi văn bản tĩnh trong file .docx bằng các thẻ biến {{placeholder}}
 * @param {string} filePath - Đường dẫn file gốc
 * @param {string} outputPath - Đường dẫn file kết quả sau khi gài biến
 * @param {Array<{searchText: string, key_name: string, paragraph_context?: string}>} replacements - Danh sách các cụm từ cần thay thế thành biến
 */
function injectPlaceholders(filePath, outputPath, replacements) {
  try {
    const content = fs.readFileSync(filePath, "binary");
    const zip = new PizZip(content);
    
    // Quét qua tất cả các file XML trong thư mục word/ của docx (bao gồm cả document, headers, footers)
    const xmlFiles = zip.file(/word\/.*\.xml/);
    
    // Tạo bản đồ ánh xạ các placeholder hiện tại về chuỗi gốc để hỗ trợ so khớp ngữ cảnh trong cùng một đoạn văn
    const placeholderMap = {};
    for (const rep of replacements) {
      if (rep.key_name && rep.searchText) {
        placeholderMap[`{{${rep.key_name}}}`] = rep.searchText;
      }
    }
    
    for (const file of xmlFiles) {
      let docXml = file.asText();
      let modified = false;
      
      for (const rep of replacements) {
        const { searchText, key_name, paragraph_context } = rep;
        if (!searchText || !key_name) continue;

        const placeholder = `{{${key_name}}}`;
        // Nếu file XML đã có sẵn placeholder này rồi, bỏ qua không chạy thay thế
        if (docXml.includes(placeholder)) {
          continue;
        }

        // Thực hiện chuẩn hóa paragraph XML và thay thế có đối chiếu ngữ cảnh
        const newXml = normalizeAndReplaceText(docXml, searchText, placeholder, paragraph_context, placeholderMap);
        if (newXml !== docXml) {
          docXml = newXml;
          modified = true;
        }
      }
      
      if (modified) {
        zip.file(file.name, docXml);
      }
    }

    const buffer = zip.generate({ type: "nodebuffer" });
    fs.writeFileSync(outputPath, buffer);
    console.log(`[Tagging] Đã cập nhật file mẫu và gài các biến vào các XML:`, replacements);
  } catch (error) {
    console.error("Lỗi khi gài biến vào file DOCX:", error);
    throw new Error("Không thể ghi đè biến vào cấu trúc tệp Word: " + error.message);
  }
}

function mergeAdjacentRunsInXml(pXml) {
  // Regex matches consecutive runs with text and merges them if their styles (rPr) are identical.
  const runPairRegex = /(<w:r\b[^>]*>)(?:<w:rPr\b[^>]*>([\s\S]*?)<\/w:rPr>)?(<w:t\b[^>]*>)([\s\S]*?)<\/w:t><\/w:r>(\s*)<w:r\b[^>]*>(?:<w:rPr\b[^>]*>([\s\S]*?)<\/w:rPr>)?(<w:t\b[^>]*>)([\s\S]*?)<\/w:t><\/w:r>/g;
  
  const cleanPr = (pr) => {
    return (pr || "")
      .replace(/<w:lang\b[^>]*>/g, "")
      .replace(/<w:noProof\b[^>]*>/g, "")
      .replace(/<w:proofErr\b[^>]*>/g, "")
      .replace(/\s+/g, ' ')
      .trim();
  };

  let xml = pXml;
  let changed = true;
  while (changed) {
    changed = false;
    xml = xml.replace(runPairRegex, (match, r1Open, r1Pr, t1Open, t1Text, spacing, r2Pr, t2Open, t2Text) => {
      const r1PrNorm = cleanPr(r1Pr);
      const r2PrNorm = cleanPr(r2Pr);
      
      if (r1PrNorm === r2PrNorm) {
        changed = true;
        const mergedText = t1Text + t2Text;
        const xmlSpace = (mergedText.startsWith(' ') || mergedText.endsWith(' ')) ? ' xml:space="preserve"' : '';
        const newR1Pr = r1Pr ? `<w:rPr>${r1Pr}</w:rPr>` : '';
        return `${r1Open}${newR1Pr}<w:t${xmlSpace}>${mergedText}</w:t></w:r>`;
      }
      return match;
    });
  }
  return xml;
}

/**
 * Quét các paragraph XML, gộp các Run bị chia cắt và thay thế văn bản tĩnh thành biến
 */
function normalizeAndReplaceText(docXml, searchText, placeholder, paragraphContext, placeholderMap) {
  // Loại bỏ các thẻ kiểm tra chính tả/ngữ pháp (proofErr) làm chia nhỏ các run
  docXml = docXml.replace(/<w:proofErr\b[^>]*>/g, "");

  const pRegex = /<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g;
  
  // Chuẩn hóa searchText để tạo biểu thức chính quy (Regex) tìm kiếm mềm dẻo về khoảng trắng
  const escapedSearch = searchText.trim().replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
  const searchRegexPattern = escapedSearch.replace(/\s+/g, '[\\s\\u00a0]+');
  const searchRegex = new RegExp(searchRegexPattern, 'g');

  return docXml.replace(pRegex, (pXml) => {
    // Trích xuất text trần đoạn văn dùng helper mới (có quy đổi Tab/Br/Cr)
    const pText = getParagraphText(pXml);
    const normPText = pText.replace(/[\s\u00a0\t\r\n]+/g, " ");
    const normSearchText = searchText.trim().replace(/[\s\u00a0\t\r\n]+/g, " ");

    // Khôi phục tạm thời các placeholder về chữ gốc để phục vụ kiểm tra chứa từ khóa
    let checkText = normPText;
    if (placeholderMap) {
      for (const [plc, originalText] of Object.entries(placeholderMap)) {
        if (checkText.includes(plc)) {
          const cleanOriginal = originalText.replace(/[\s\u00a0\t\r\n]+/g, " ").trim();
          checkText = checkText.replaceAll(plc, cleanOriginal);
        }
      }
    }

    // Nếu văn bản trần có chứa cụm từ cần thay thế
    if (checkText.includes(normSearchText)) {
      // Nếu có paragraphContext truyền lên từ client, ta phải so khớp ngữ cảnh của đoạn văn
      if (paragraphContext) {
        // Chuẩn hóa khoảng trắng để so sánh chính xác hơn
        let cleanText = pText.replace(/[\s\u00a0\t\r\n]+/g, " ").trim();
        const cleanContext = paragraphContext.replace(/[\s\u00a0\t\r\n]+/g, " ").trim();
        
        // Khôi phục các placeholder về văn bản gốc để so khớp ngữ cảnh chính xác
        if (placeholderMap) {
          for (const [plc, originalText] of Object.entries(placeholderMap)) {
            if (cleanText.includes(plc)) {
              const cleanOriginal = originalText.replace(/[\s\u00a0\t\r\n]+/g, " ").trim();
              cleanText = cleanText.replaceAll(plc, cleanOriginal);
            }
          }
        }
        
        // Đoạn văn trong XML khớp nếu nó chứa ngữ cảnh hoặc ngữ cảnh chứa nó
        const isMatch = cleanText.includes(cleanContext) || cleanContext.includes(cleanText);
        if (!isMatch) {
          return pXml; // Không khớp ngữ cảnh đoạn văn bôi đen, bỏ qua paragraph này
        }
      }

      console.log(`[Tagging] Khớp đoạn văn bôi đen: "${pText}"`);
      
      // 1. Thử gộp các run kề nhau có cùng định dạng (style) trong XML
      const pXmlMerged = mergeAdjacentRunsInXml(pXml);
      
      // 2. Tìm xem có run nào chứa toàn bộ searchText không
      const rRegex = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g;
      let matchedRunIndex = -1;
      const runs = [];
      let rMatch;
      while ((rMatch = rRegex.exec(pXmlMerged)) !== null) {
        runs.push({
          full: rMatch[0],
          inner: rMatch[1]
        });
      }
      
      const parsedRuns = runs.map(run => {
        const tMatch = run.inner.match(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/);
        const tText = tMatch ? tMatch[1] : null;
        return {
          full: run.full,
          inner: run.inner,
          tText
        };
      });
      
      for (let i = 0; i < parsedRuns.length; i++) {
        if (parsedRuns[i].tText) {
          const rawTText = unescapeXml(parsedRuns[i].tText);
          const normTText = rawTText.replace(/[\s\u00a0\t\r\n]+/g, " ");
          if (normTText.includes(normSearchText)) {
            matchedRunIndex = i;
            break;
          }
        }
      }
      
      if (matchedRunIndex !== -1) {
        console.log(`[Tagging] Tìm thấy run chứa từ khóa. Tiến hành thay thế giữ style.`);
        const targetRun = parsedRuns[matchedRunIndex];
        const rawTText = unescapeXml(targetRun.tText);
        const newTText = rawTText.replace(searchRegex, placeholder);
        const escapedNewTText = escapeXml(newTText);
        const xmlSpace = (escapedNewTText.startsWith(' ') || escapedNewTText.endsWith(' ')) ? ' xml:space="preserve"' : '';
        
        const newRunInner = targetRun.inner.replace(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/, `<w:t${xmlSpace}>${escapedNewTText}</w:t>`);
        // Use functional replacement to be completely safe against special characters like $
        const newRunFull = targetRun.full.replace(targetRun.inner, () => newRunInner);
          
        let runCount = 0;
        const reconstructedParagraph = pXmlMerged.replace(/<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g, (originalRun) => {
          if (runCount === matchedRunIndex) {
            runCount++;
            return newRunFull;
          }
          runCount++;
          return originalRun;
        });
        
        return reconstructedParagraph;
      }
      
      // 3. Fallback: Nếu không có run đơn lẻ nào chứa trọn vẹn, ta gộp cả paragraph và áp style run đầu tiên
      console.log(`[Tagging] Fallback: Không tìm thấy run chứa trọn vẹn từ khóa. Tiến hành gộp và áp style run đầu tiên.`);
      const newText = pText.replace(searchRegex, placeholder);
      
      // Tìm run đầu tiên có thuộc tính rPr để lấy style thừa kế
      let rPrXml = "";
      for (const pr of parsedRuns) {
        const matchPr = pr.inner.match(/<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/);
        if (matchPr) {
          rPrXml = matchPr[0];
          break;
        }
      }
      
      const pPrMatch = pXml.match(/<w:pPr\b[^>]*>([\s\S]*?)<\/w:pPr>/);
      const pPrXml = pPrMatch ? pPrMatch[0] : "";
      
      const newInnerXml = `${pPrXml}<w:r>${rPrXml}<w:t xml:space="preserve">${escapeXml(newText)}</w:t></w:r>`;
      const openTagMatch = pXml.match(/^<w:p\b[^>]*>/);
      const openTag = openTagMatch ? openTagMatch[0] : "<w:p>";
      return `${openTag}${newInnerXml}</w:p>`;
    }
    
    return pXml;
  });
}

function restorePlaceholder(docXml, key_name, replace_text) {
  const placeholder = `{{${key_name}}}`;
  const pRegex = /<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g;
  
  return docXml.replace(pRegex, (pXml) => {
    if (pXml.includes(placeholder)) {
      const rRegex = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g;
      return pXml.replace(rRegex, (rXml) => {
        if (rXml.includes(placeholder)) {
          return rXml.replace(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/, (tXml, tText) => {
            const newTText = tText.replaceAll(placeholder, replace_text);
            const xmlSpace = (newTText.startsWith(' ') || newTText.endsWith(' ')) ? ' xml:space="preserve"' : '';
            return `<w:t${xmlSpace}>${escapeXml(newTText)}</w:t>`;
          });
        }
        return rXml;
      });
    }
    return pXml;
  });
}

function escapeXml(unsafe) {
  return unsafe.replace(/[<>&'"]/g, (c) => {
    switch (c) {
      case '<': return '&lt;';
      case '>': return '&gt;';
      case '&': return '&amp;';
      case '\'': return '&apos;';
      case '"': return '&quot;';
    }
  });
}

function scanTables(filePath) {
  // Read DOCX and extract table structures
  const content = fs.readFileSync(filePath, "binary");
  const zip = new PizZip(content);
  const docXml = zip.file("word/document.xml").asText();

  // Match all tables
  const tableRegex = /<w:tbl[\s\S]*?<\/w:tbl>/g;
  const tables = [];
  let tblMatch;
  let tableIndex = 0;
  while ((tblMatch = tableRegex.exec(docXml)) !== null) {
    const tblXml = tblMatch[0];
    // Match rows
    const rowRegex = /<w:tr[\s\S]*?<\/w:tr>/g;
    const rows = [];
    let trMatch;
    while ((trMatch = rowRegex.exec(tblXml)) !== null) {
      const trXml = trMatch[0];
      // Match cells
      const cellRegex = /<w:tc[\s\S]*?<\/w:tc>/g;
      const cells = [];
      let tcMatch;
      while ((tcMatch = cellRegex.exec(trXml)) !== null) {
        const tcXml = tcMatch[0];
        // Extract visible text inside <w:t> tags
        const textMatches = tcXml.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g) || [];
        const cellText = textMatches.map(t => t.replace(/<[^>]+>/g, "")).join("");
        cells.push(cellText.trim());
      }
      rows.push({ cells });
    }
    if (rows.length > 0) {
      const headers = rows[0].cells;
      const dataRows = rows.slice(1).map((r, idx) => ({ rowIndex: idx + 1, cells: r.cells }));
      tables.push({
        tableIndex,
        headers,
        rows: dataRows
      });
      tableIndex++;
    }
  }
  return tables;
}

function injectTablePlaceholders(filePath, outputPath, tableIndex, fields, selectedRows) {
  try {
    const content = fs.readFileSync(filePath, "binary");
    const zip = new PizZip(content);
    let docXml = zip.file("word/document.xml").asText();

    // Match all tables
    const tableRegex = /<w:tbl[\s\S]*?<\/w:tbl>/g;
    const tables = [];
    let match;
    while ((match = tableRegex.exec(docXml)) !== null) {
      tables.push({
        xml: match[0],
        index: match.index,
        length: match[0].length
      });
    }

    if (tableIndex < 0 || tableIndex >= tables.length) {
      throw new Error(`Không tìm thấy bảng với chỉ mục ${tableIndex}`);
    }

    const targetTable = tables[tableIndex];
    let tblXml = targetTable.xml;

    // Match all rows in this table
    const rowRegex = /<w:tr[\s\S]*?<\/w:tr>/g;
    const rows = [];
    let rowMatch;
    while ((rowMatch = rowRegex.exec(tblXml)) !== null) {
      rows.push({
        xml: rowMatch[0],
        index: rowMatch.index,
        length: rowMatch[0].length
      });
    }

    // We keep headers intact (row 0 is header). Data rows start at row 1.
    // selectedRows: array of row numbers (1-indexed) e.g., [1, 2]
    const sortedRows = [...selectedRows].sort((a, b) => b - a); // Sort descending to preserve index offsets!
    for (const r of sortedRows) {
      if (r < 1 || r >= rows.length) continue;
      const targetRow = rows[r];
      let rowXml = targetRow.xml;

      // Match all cells in this row
      const cellRegex = /<w:tc[\s\S]*?<\/w:tc>/g;
      const cells = [];
      let cellMatch;
      while ((cellMatch = cellRegex.exec(rowXml)) !== null) {
        cells.push({
          xml: cellMatch[0],
          index: cellMatch.index,
          length: cellMatch[0].length
        });
      }

      // Replace cells from right to left (descending colIndex)
      const sortedFields = [...fields].sort((a, b) => b.colIndex - a.colIndex);

      for (const f of sortedFields) {
        const colIdx = f.colIndex;
        if (colIdx < 0 || colIdx >= cells.length) continue;

        const targetCell = cells[colIdx];
        const cellXml = targetCell.xml;

        // Extract w:tcPr (cell properties)
        const tcPrMatch = cellXml.match(/<w:tcPr\b[^>]*>[\s\S]*?<\/w:tcPr>/);
        const tcPrXml = tcPrMatch ? tcPrMatch[0] : "";

        // Extract first w:pPr (paragraph properties) to preserve alignment/styles
        const pPrMatch = cellXml.match(/<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/);
        const pPrXml = pPrMatch ? pPrMatch[0] : "";

        // Extract first w:rPr (run properties) to preserve font/bold style
        const rPrMatch = cellXml.match(/<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/);
        const rPrXml = rPrMatch ? rPrMatch[0] : "";

        const placeholderKey = `${f.key_name}_${r}`;
        const newCellContent = `<w:p>${pPrXml}<w:r>${rPrXml}<w:t xml:space="preserve">{{${placeholderKey}}}</w:t></w:r></w:p>`;
        const newCellXml = `<w:tc>${tcPrXml}${newCellContent}</w:tc>`;

        // Replace cell XML in rowXml
        rowXml = rowXml.substring(0, targetCell.index) + newCellXml + rowXml.substring(targetCell.index + targetCell.length);
      }

      // Replace row XML in tblXml
      tblXml = tblXml.substring(0, targetRow.index) + rowXml + tblXml.substring(targetRow.index + targetRow.length);
    }

    // Replace table XML in docXml
    docXml = docXml.substring(0, targetTable.index) + tblXml + docXml.substring(targetTable.index + targetTable.length);

    zip.file("word/document.xml", docXml);
    const buffer = zip.generate({ type: "nodebuffer" });
    fs.writeFileSync(outputPath, buffer);
    console.log(`[Table Detection] Successfully injected table placeholders for Table #${tableIndex + 1}`);
  } catch (error) {
    console.error("Lỗi khi gài biến bảng:", error);
    throw new Error("Không thể ghi đè biến bảng vào tệp Word: " + error.message);
  }
}

module.exports = {
  scanPlaceholders,
  mergeDocument,
  mergeDocumentToBuffer,
  injectPlaceholders,
  restorePlaceholder,
  scanTables,
  injectTablePlaceholders
};
