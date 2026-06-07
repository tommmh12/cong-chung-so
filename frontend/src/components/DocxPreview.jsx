import { useState, useEffect, useRef } from 'react';
import * as docx from 'docx-preview';

const API_BASE = 'http://localhost:5000/api';

// Helper to find all Text Nodes under an element recursively
function getTextNodes(node) {
  const textNodes = [];
  function traverse(n) {
    if (n.nodeType === 3) { // Node.TEXT_NODE
      textNodes.push(n);
    } else {
      for (let i = 0; i < n.childNodes.length; i++) {
        traverse(n.childNodes[i]);
      }
    }
  }
  traverse(node);
  return textNodes;
}

// Create styled replacement element safely
function createReplacementElement(field, val, isModal) {
  const replEl = document.createElement('span');
  if (!isModal) {
    replEl.setAttribute('data-field-key', field.key_name);
  }
  
  if (val !== undefined && val !== null && val.toString().trim() !== '') {
    if (isModal) {
      replEl.className = "bg-emerald-50 text-emerald-800 px-1 rounded border border-emerald-200 mx-0.5";
    } else {
      replEl.className = "cursor-pointer hover:ring-2 hover:ring-emerald-500 transition-all bg-emerald-50 text-emerald-800 px-1 rounded border border-emerald-200 inline-block mx-0.5";
    }
    replEl.textContent = val;
  } else {
    // Unfilled: muted gray italic, inherits document font
    replEl.style.cssText = 'color:#bbb;font-style:italic;font-weight:normal;font-size:0.85em;border-bottom:1px dotted #ddd;padding-bottom:1px;';
    if (!isModal) {
      replEl.style.cursor = 'pointer';
      replEl.onmouseenter = () => { replEl.style.color = '#999'; replEl.style.borderBottomColor = '#aaa'; };
      replEl.onmouseleave = () => { replEl.style.color = '#bbb'; replEl.style.borderBottomColor = '#ddd'; };
    }
    replEl.textContent = field.label;
  }
  return replEl;
}

// Safely replaces placeholders in DOM tree text nodes without corrupting HTML markup
function replacePlaceholdersInElement(el, fieldsList, data, isModal) {
  let textNodes = getTextNodes(el);
  let concatenatedText = textNodes.map(n => n.textContent).join('');
  
  if (!concatenatedText.includes('{')) return;
  
  let found = true;
  while (found) {
    found = false;
    textNodes = getTextNodes(el);
    concatenatedText = textNodes.map(n => n.textContent).join('');
    
    let match = null;
    for (let i = 0; i < fieldsList.length; i++) {
      const field = fieldsList[i];
      // Match {{key_name}} allowing optional whitespace inside the curly braces
      const regex = new RegExp('{\\s*{\\s*' + field.key_name + '\\s*}\\s*}', 'i');
      const m = regex.exec(concatenatedText);
      if (m) {
        match = { 
          field, 
          index: m.index, 
          length: m[0].length 
        };
        break; // Match first, split nodes, then loop again
      }
    }
    
    if (match) {
      found = true;
      const { field, index, length } = match;
      const startPos = index;
      const endPos = index + length;
      
      let currentLen = 0;
      let startNode = null;
      let startOffset = 0;
      let endNode = null;
      let endOffset = 0;
      
      const nodesToClear = [];
      
      for (let i = 0; i < textNodes.length; i++) {
        const node = textNodes[i];
        const nodeLen = node.textContent.length;
        
        if (startNode === null && currentLen + nodeLen >= startPos) {
          startNode = node;
          startOffset = startPos - currentLen;
        }
        
        if (endNode === null && currentLen + nodeLen >= endPos) {
          endNode = node;
          endOffset = endPos - currentLen;
          
          if (startNode && startNode !== endNode) {
            const startIdx = textNodes.indexOf(startNode);
            for (let j = startIdx + 1; j < i; j++) {
              nodesToClear.push(textNodes[j]);
            }
          }
          break;
        }
        
        currentLen += nodeLen;
      }
      
      if (startNode && endNode) {
        const val = data ? data[field.key_name] : undefined;
        const replEl = createReplacementElement(field, val, isModal);
        
        // Clear all intermediate text node text
        nodesToClear.forEach(node => {
          node.textContent = '';
        });
        
        if (startNode === endNode) {
          const parent = startNode.parentNode;
          const text = startNode.textContent;
          const beforeText = text.substring(0, startOffset);
          const afterText = text.substring(endOffset);
          
          const beforeNode = document.createTextNode(beforeText);
          const afterNode = document.createTextNode(afterText);
          
          parent.insertBefore(beforeNode, startNode);
          parent.insertBefore(replEl, startNode);
          parent.insertBefore(afterNode, startNode);
          parent.removeChild(startNode);
        } else {
          const startText = startNode.textContent;
          startNode.textContent = startText.substring(0, startOffset);
          
          const endParent = endNode.parentNode;
          const endText = endNode.textContent;
          endNode.textContent = endText.substring(endOffset);
          
          endParent.insertBefore(replEl, endNode);
        }
      }
    }
  }
}

// Helper: Simulated A4 sheet render component inside PDF Modal
export function DocxPreviewInModal({ fileUrl, liveData, fields }) {
  const containerRef = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const originalHtmlRef = useRef('');
  const prevPropsRef = useRef('');
  const liveDataRef = useRef(liveData);
  const fieldsRef = useRef(fields);
  useEffect(() => { liveDataRef.current = liveData; }, [liveData]);
  useEffect(() => { fieldsRef.current = fields; }, [fields]);

  const applyLiveData = (data, fieldsList) => {
    if (!originalHtmlRef.current || !containerRef.current) return;

    prevPropsRef.current = JSON.stringify({ fields: fieldsList, liveData: data });

    containerRef.current.innerHTML = originalHtmlRef.current;

    const elements = containerRef.current.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li, td');

    elements.forEach(el => {
      replacePlaceholdersInElement(el, fieldsList, data, true);
    });
  };

  useEffect(() => {
    if (!fileUrl) return;

    let isMounted = true;
    async function loadDocx() {
      setLoading(true);
      setError(null);
      originalHtmlRef.current = '';
      prevPropsRef.current = '';
      try {
        const response = await fetch(fileUrl);
        if (!response.ok) throw new Error("Không thể tải file mẫu.");
        const blob = await response.blob();

        if (isMounted && containerRef.current) {
          containerRef.current.innerHTML = "";
          await docx.renderAsync(blob, containerRef.current, null, {
            className: "docx-rendered-modal",
            inWrapper: false,
            ignoreWidth: true,
            ignoreHeight: true,
            debug: false
          });

          originalHtmlRef.current = containerRef.current.innerHTML;
          const latestFields = fieldsRef.current;
          const latestLiveData = liveDataRef.current;
          if (latestFields && latestFields.length > 0) {
            applyLiveData(latestLiveData || {}, latestFields);
          }
        }
      } catch (err) {
        if (isMounted) setError(err.message);
      } finally {
        if (isMounted) setLoading(false);
      }
    }

    loadDocx();
    return () => {
      isMounted = false;
    };
  }, [fileUrl]);

  useEffect(() => {
    if (fields && fields.length > 0 && originalHtmlRef.current) {
      const serialized = JSON.stringify({ fields, liveData });
      if (serialized === prevPropsRef.current) {
        return; // Guard redundant runs
      }
      applyLiveData(liveData || {}, fields);
    }
  }, [liveData, fields]);

  return (
    <div className="w-full text-black">
      {loading && (
        <div className="h-full min-h-[150px] flex items-center justify-center text-zinc-400 text-xs font-sans">
          <div className="w-4 h-4 border-2 border-emerald-500/30 border-t-emerald-500 rounded-full animate-spin mr-2"></div>
          Đang chuẩn bị trang in...
        </div>
      )}
      {error && (
        <div className="text-rose-500 text-xs py-6 text-center font-sans">⚠️ Lỗi: {error}</div>
      )}
      <div ref={containerRef} className="docx-container-modal" style={{ display: loading || error ? 'none' : 'block' }} />
    </div>
  );
}

// Child wrapper for Modal preview that loads fields dynamically
export function DocxPreviewInModalChild({ child, previewKey, formData, recordData }) {
  const [childFields, setChildFields] = useState([]);
  const [resolvedData, setResolvedData] = useState({});
  const [loadingFields, setLoadingFields] = useState(true);

  useEffect(() => {
    async function loadForm() {
      try {
        const res = await fetch(`${API_BASE}/templates/${child.id}/form`);
        if (res.ok) {
          const data = await res.json();
          setChildFields(data.fields || []);
        }
      } catch (err) {
        console.error(err);
      } finally {
        setLoadingFields(false);
      }
    }
    loadForm();
  }, [child.id]);

  useEffect(() => {
    if (childFields.length > 0) {
      const resolved = {};
      childFields.forEach(f => {
        if (f.parent_field_key) {
          resolved[f.key_name] = formData[f.parent_field_key] || '';
        } else if (recordData) {
          resolved[f.key_name] = recordData[f.key_name] || '';
        } else {
          resolved[f.key_name] = formData[f.key_name] || '';
        }
      });
      setResolvedData(resolved);
    }
  }, [childFields, formData, recordData]);

  if (loadingFields) {
    return (
      <div className="h-full min-h-[150px] flex items-center justify-center text-zinc-400 text-xs font-sans">
        <div className="w-4 h-4 border-2 border-emerald-500/30 border-t-emerald-500 rounded-full animate-spin mr-2"></div>
        Đang nạp cấu trúc tài liệu...
      </div>
    );
  }

  return (
    <DocxPreviewInModal
      fileUrl={`${API_BASE}/templates/${child.id}/download-original?t=${previewKey}`}
      liveData={resolvedData}
      fields={childFields}
    />
  );
}

// Main inline document viewer with highlighted tag navigation
export function DocxPreview({ fileUrl, title, liveData, fields, onTableRowClick, hideToolbar }) {
  const containerRef = useRef(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const originalHtmlRef = useRef('');
  const prevPropsRef = useRef('');
  // Always keep refs pointing to latest props so async loadDocx uses current values
  const liveDataRef = useRef(liveData);
  const fieldsRef = useRef(fields);
  useEffect(() => { liveDataRef.current = liveData; }, [liveData]);
  useEffect(() => { fieldsRef.current = fields; }, [fields]);

  const applyLiveData = (data, fieldsList) => {
    if (!originalHtmlRef.current || !containerRef.current) return;

    // Save current parameters in ref to avoid redundant DOM refreshes
    prevPropsRef.current = JSON.stringify({ fields: fieldsList, liveData: data });

    // First, restore the clean original HTML structure
    containerRef.current.innerHTML = originalHtmlRef.current;

    // Find all paragraphs, headings, list items, table cells
    const elements = containerRef.current.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li, td');

    elements.forEach(el => {
      replacePlaceholdersInElement(el, fieldsList, data, false);
    });
  };

  useEffect(() => {
    if (!fileUrl) return;

    let isMounted = true;
    async function loadDocx() {
      setLoading(true);
      setError(null);
      originalHtmlRef.current = '';
      prevPropsRef.current = '';
      try {
        const response = await fetch(fileUrl);
        if (!response.ok) throw new Error("Không thể tải file tài liệu để xem trước.");
        const blob = await response.blob();

        if (isMounted && containerRef.current) {
          containerRef.current.innerHTML = "";
          await docx.renderAsync(blob, containerRef.current, null, {
            className: "docx-rendered",
            inWrapper: true,
            ignoreWidth: false,
            ignoreHeight: true,
            debug: false
          });

          originalHtmlRef.current = containerRef.current.innerHTML;
          // Use refs to get latest liveData/fields regardless of when DOCX finishes loading
          const latestFields = fieldsRef.current;
          const latestLiveData = liveDataRef.current;
          if (latestFields && latestFields.length > 0) {
            applyLiveData(latestLiveData || {}, latestFields);
          }
        }
      } catch (err) {
        if (isMounted) setError(err.message);
      } finally {
        if (isMounted) setLoading(false);
      }
    }

    loadDocx();
    return () => {
      isMounted = false;
    };
  }, [fileUrl]);

  useEffect(() => {
    if (fields && fields.length > 0 && originalHtmlRef.current) {
      const serialized = JSON.stringify({ fields, liveData });
      if (serialized === prevPropsRef.current) {
        return; // Guard redundant runs
      }
      applyLiveData(liveData || {}, fields);
    }
  }, [liveData, fields]);

  const handleContainerClick = (e) => {
    const target = e.target.closest('[data-field-key]');
    if (target) {
      const fieldKey = target.getAttribute('data-field-key');
      if (fieldKey) {
        const inputEl = document.getElementById(`field-input-${fieldKey}`);
        const configEl = document.getElementById(`field-config-${fieldKey}`);
        
        if (inputEl) {
          inputEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
          inputEl.focus();
          inputEl.classList.add('ring-2', 'ring-primary-500', 'border-primary-500');
          setTimeout(() => {
            inputEl.classList.remove('ring-2', 'ring-primary-500', 'border-primary-500');
          }, 1550);
        } else if (configEl) {
          configEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
          configEl.classList.add('border-emerald-500', 'bg-emerald-50');
          setTimeout(() => {
            configEl.classList.remove('border-emerald-500', 'bg-emerald-50');
          }, 1550);
        }
      }
    } else if (onTableRowClick) {
      const trTarget = e.target.closest('tr');
      if (trTarget) {
        const tableEl = trTarget.closest('table');
        if (tableEl && containerRef.current) {
          const tables = Array.from(containerRef.current.querySelectorAll('table'));
          const tableIndex = tables.indexOf(tableEl);
          const rows = Array.from(tableEl.querySelectorAll('tr'));
          const rowIndex = rows.indexOf(trTarget);
          
          if (tableIndex !== -1 && rowIndex !== -1) {
            onTableRowClick(tableIndex, rowIndex, trTarget);
          }
        }
      }
    }
  };

  return (
    <div
      onClick={handleContainerClick}
      style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0, background: '#f2f4f6', overflow: 'hidden' }}
    >
      {/* Preview toolbar */}
      <div style={{
        display: hideToolbar ? 'none' : 'flex', alignItems: 'center', justifyContent: 'space-between',
        padding: '8px 16px',
        background: '#ffffff',
        borderBottom: '1px solid #e6e8ea',
        flexShrink: 0,
        gap: 8,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <span className="material-symbols-outlined" style={{ fontSize: 18, color: '#45464d' }}>description</span>
          <span style={{ fontSize: 12, fontWeight: 600, color: '#191c1e', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 320 }}>
            {title || 'Xem trước tài liệu'}
          </span>
        </div>
        {loading && (
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: '#76777d' }}>
            <div style={{ width: 14, height: 14, border: '2px solid #e6e8ea', borderTopColor: '#000000', borderRadius: '50%', animation: 'lx-spin 0.65s linear infinite', flexShrink: 0 }} />
            Đang tải...
          </div>
        )}
      </div>

      {/* Document area — gray "desktop" so white A4 pages stand out */}
      <div style={{ flex: 1, overflowY: 'auto', overflowX: 'auto', background: '#e2e5e9' }}>
        {error && (
          <div style={{ margin: 24, textAlign: 'center', color: '#ba1a1a', padding: '48px 24px', fontSize: 13, background: '#ffffff', border: '1px solid #c6c6cd', borderRadius: 6 }}>
            <span className="material-symbols-outlined" style={{ fontSize: 32, display: 'block', marginBottom: 8 }}>error_outline</span>
            Lỗi tải bản xem trước: {error}
          </div>
        )}
        {!error && loading && (
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10, padding: '60px 24px', color: '#45464d', fontSize: 13 }}>
            <div style={{ width: 20, height: 20, border: '2px solid #c6c6cd', borderTopColor: '#191c1e', borderRadius: '50%', animation: 'lx-spin 0.65s linear infinite', flexShrink: 0 }} />
            Đang nạp văn bản mẫu...
          </div>
        )}
        <div ref={containerRef} className="docx-container" style={{ display: loading || error ? 'none' : 'block' }} />
      </div>
    </div>
  );
}
