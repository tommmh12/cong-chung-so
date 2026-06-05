import { useState, useEffect, useRef, useCallback, Fragment } from 'react';
import * as docx from 'docx-preview';

const API_BASE = 'http://localhost:5000/api';
const SUPPORTED_WORD_EXTENSIONS = ['doc', 'docx'];
const MAX_UPLOAD_FILES = 10;
const MAX_UPLOAD_FILE_SIZE = 1024 * 1024;

import { renderAsync } from 'docx-preview';

const DocxFilePreview = ({ subId, filename }) => {
  const containerRef = useRef(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  useEffect(() => {
    let active = true;
    if (!subId || !filename) return;

    const loadAndRender = async () => {
      setLoading(true);
      setError(null);
      try {
        if (containerRef.current) {
          containerRef.current.innerHTML = "";
        }
        
        const url = `http://localhost:5000/api/submissions/${subId}/download-file?filename=${encodeURIComponent(filename)}`;
        const res = await fetch(url);
        if (!res.ok) {
          throw new Error("Không thể tải file Word từ máy chủ.");
        }
        const arrayBuffer = await res.arrayBuffer();
        
        if (active && containerRef.current) {
          await renderAsync(arrayBuffer, containerRef.current, null, {
            className: "docx-rendered",
            inWrapper: true,
            ignoreWidth: false,
            ignoreHeight: false,
            ignoreFonts: false,
            breakPages: true,
            experimental: true
          });
        }
      } catch (err) {
        console.error("Docx render error:", err);
        if (active) {
          setError(err.message || "Lỗi khi kết xuất file docx.");
        }
      } finally {
        if (active) {
          setLoading(false);
        }
      }
    };

    loadAndRender();

    return () => {
      active = false;
    };
  }, [subId, filename]);

  return (
    <div className="space-y-3 font-sans">
      {loading && (
        <div className="py-12 text-center text-slate-400 text-xs flex flex-col items-center justify-center gap-2">
          <div className="w-5 h-5 border-2 border-emerald-500/30 border-t-emerald-500 rounded-full animate-spin"></div>
          <span>Đang tạo bản xem trước tài liệu Word (.docx)...</span>
        </div>
      )}
      
      {error && (
        <div className="py-6 text-center text-rose-500 text-xs italic border border-dashed border-rose-200 rounded-xl bg-rose-50/30">
          ⚠️ {error}
        </div>
      )}

      <div 
        ref={containerRef} 
        className="docx-container overflow-y-auto max-h-[600px] bg-slate-100/60 p-4 border border-slate-200/80 rounded-xl scrollbar-thin shadow-inner"
        style={{ display: loading || error ? 'none' : 'block' }}
      />
    </div>
  );
};

const getFileExtension = (fileName) => fileName.split('.').pop().toLowerCase();
const isSupportedWordFile = (fileName) => SUPPORTED_WORD_EXTENSIONS.includes(getFileExtension(fileName));
const getWordBaseName = (fileName) => fileName.replace(/\.(docx|doc)$/i, '');

const generateId = () => {
  return `temp-${Date.now()}-${Math.random().toString(36).substring(2, 11)}`;
};

const getFieldStep = (field) => {
  const key = (field.key_name || '').toLowerCase();
  const label = (field.label || '').toLowerCase();
  const childGroup = (field.groupName || '').toLowerCase();

  if (
    key.includes('tncn') || key.includes('truocba') || key.includes('lptb') || key.includes('thuế') || key.includes('thue_') ||
    label.includes('thuế tncn') || label.includes('trước bạ') || label.includes('lptb') || label.includes('lệ phí') ||
    childGroup.includes('thuế') || childGroup.includes('lệ phí trước bạ') || childGroup.includes('trước bạ')
  ) {
    return 2;
  }

  if (
    key.includes('ct_') || key.includes('matdo') || key.includes('caytrong') || key.includes('lamnghiep') || key.includes('cây') || key.includes('chăm sóc') || key.includes('kinh doanh') ||
    label.includes('cây trồng') || label.includes('mật độ') || label.includes('lâm nghiệp') || label.includes('cây loại') ||
    childGroup.includes('cây trồng') || childGroup.includes('phụ lục cây trồng') || childGroup.includes('lâm sản')
  ) {
    return 3;
  }

  if (
    key.includes('pnn') || key.includes('phinongnghiep') || key.includes('sở hữu') ||
    label.includes('pnn') || label.includes('phi nông nghiệp') || label.includes('tỷ lệ sở hữu') ||
    childGroup.includes('pnn') || childGroup.includes('phi nông nghiệp')
  ) {
    return 4;
  }

  if (field.groupName) {
    return 4;
  }
  
  return 1;
};

// Simulated A4 sheet render component inside PDF Modal
function DocxPreviewInModal({ fileUrl, liveData, fields }) {
  const containerRef = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const originalHtmlRef = useRef('');

  const applyLiveData = (data, fieldsList) => {
    if (!originalHtmlRef.current || !containerRef.current) return;
    
    let html = originalHtmlRef.current;
    const cleanHtmlRegex = /\{(<(?:span|font|strong|em|u|b|i)\b[^>]*>|<\/(?:span|font|strong|em|u|b|i)>|\s)*\{(<(?:span|font|strong|em|u|b|i)\b[^>]*>|<\/(?:span|font|strong|em|u|b|i)>|\s)*([a-zA-Z0-9_]+)(<(?:span|font|strong|em|u|b|i)\b[^>]*>|<\/(?:span|font|strong|em|u|b|i)>|\s)*\}(<(?:span|font|strong|em|u|b|i)\b[^>]*>|<\/(?:span|font|strong|em|u|b|i)>|\s)*\}/gi;
    html = html.replace(cleanHtmlRegex, '{{$3}}');

    fieldsList.forEach(field => {
      const val = data ? data[field.key_name] : undefined;
      const placeholder = `{{${field.key_name}}}`;
      
      let replacement;
      if (val !== undefined && val !== null && val.toString().trim() !== '') {
        replacement = ` <span class="bg-emerald-50 text-emerald-800 px-1 py-0.5 rounded border border-emerald-250 font-sans mx-0.5 text-[13px] font-semibold">${val}</span> `;
      } else {
        replacement = ` <span class="bg-zinc-50 text-zinc-500 border border-dashed border-zinc-350 px-1 py-0.5 rounded mx-0.5 text-[11px] font-sans font-normal italic">${field.label}</span> `;
      }
      
      html = html.replaceAll(placeholder, replacement);
    });
    
    containerRef.current.innerHTML = html;
  };

  useEffect(() => {
    if (!fileUrl) return;

    let isMounted = true;
    async function loadDocx() {
      setLoading(true);
      setError(null);
      originalHtmlRef.current = '';
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
          if (fields && fields.length > 0) {
            applyLiveData(liveData || {}, fields);
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
function DocxPreviewInModalChild({ child, previewKey, formData, recordData }) {
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
      fileUrl={`http://localhost:5000/api/templates/${child.id}/download-original?t=${previewKey}`}
      liveData={resolvedData}
      fields={childFields}
    />
  );
}


// React component to preview docx files with live client-side data replacement
function DocxPreview({ fileUrl, title, liveData, fields, onTableRowClick }) {
  const containerRef = useRef(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const originalHtmlRef = useRef('');

  const applyLiveData = (data, fieldsList) => {
    if (!originalHtmlRef.current || !containerRef.current) return;
    
    let html = originalHtmlRef.current;

    // Normalize split placeholders like <span>{</span><span>{</span><span>variable</span><span>}</span><span>}</span>
    const cleanHtmlRegex = /\{(<(?:span|font|strong|em|u|b|i)\b[^>]*>|<\/(?:span|font|strong|em|u|b|i)>|\s)*\{(<(?:span|font|strong|em|u|b|i)\b[^>]*>|<\/(?:span|font|strong|em|u|b|i)>|\s)*([a-zA-Z0-9_]+)(<(?:span|font|strong|em|u|b|i)\b[^>]*>|<\/(?:span|font|strong|em|u|b|i)>|\s)*\}(<(?:span|font|strong|em|u|b|i)\b[^>]*>|<\/(?:span|font|strong|em|u|b|i)>|\s)*\}/gi;
    html = html.replace(cleanHtmlRegex, '{{$3}}');

    fieldsList.forEach(field => {
      const val = data ? data[field.key_name] : undefined;
      const placeholder = `{{${field.key_name}}}`;
      
      let replacement;
      if (val !== undefined && val !== null && val.toString().trim() !== '') {
        // Highlight green for filled variables with data-field-key and cursor-pointer style
        replacement = ` <span data-field-key="${field.key_name}" class="cursor-pointer hover:ring-2 hover:ring-emerald-500 hover:scale-[1.03] transition-all bg-emerald-50 text-emerald-800 px-1 py-0.5 rounded border border-emerald-200 inline-block mx-0.5 text-[13px] font-sans">${val}</span> `;
      } else {
        // Highlight dashed zinc for unfilled variables with data-field-key and cursor-pointer style
        replacement = ` <span data-field-key="${field.key_name}" class="cursor-pointer hover:ring-2 hover:ring-zinc-400 hover:scale-[1.03] transition-all bg-zinc-50 text-zinc-650 border border-dashed border-zinc-300 px-1.5 py-0.5 rounded-lg text-[11px] inline-block mx-0.5 font-sans">${field.label}</span> `;
      }
      
      html = html.replaceAll(placeholder, replacement);
    });
    
    containerRef.current.innerHTML = html;
  };

  useEffect(() => {
    if (!fileUrl) return;

    let isMounted = true;
    async function loadDocx() {
      setLoading(true);
      setError(null);
      originalHtmlRef.current = '';
      try {
        const response = await fetch(fileUrl);
        if (!response.ok) throw new Error("Không thể tải file tài liệu để xem trước.");
        const blob = await response.blob();
        
        if (isMounted && containerRef.current) {
          containerRef.current.innerHTML = "";
          await docx.renderAsync(blob, containerRef.current, null, {
            className: "docx-rendered",
            inWrapper: false,
            ignoreWidth: true,
            ignoreHeight: true,
            debug: false
          });
          
          originalHtmlRef.current = containerRef.current.innerHTML;
          if (fields && fields.length > 0) {
            applyLiveData(liveData || {}, fields);
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
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fileUrl]);

  // Listen to live data changes to run DOM updates instantly
  useEffect(() => {
    if (fields && fields.length > 0 && originalHtmlRef.current) {
      applyLiveData(liveData || {}, fields);
    }
  }, [liveData, fields]);

  // Click on variable on document -> scrolls and focuses to input field or config block
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
          
          // Flash input briefly
          inputEl.classList.add('ring-2', 'ring-primary-500', 'border-primary-500');
          setTimeout(() => {
            inputEl.classList.remove('ring-2', 'ring-primary-500', 'border-primary-500');
          }, 1500);
        } else if (configEl) {
          configEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
          
          // Flash config card briefly
          configEl.classList.add('border-emerald-500', 'bg-emerald-50');
          setTimeout(() => {
            configEl.classList.remove('border-emerald-500', 'bg-emerald-50');
          }, 1500);
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
    <div onClick={handleContainerClick} className="flex flex-col h-full bg-zinc-900 border border-zinc-800 rounded-lg overflow-hidden shadow-xl">
      <div className="bg-zinc-950/70 px-4 py-3 border-b border-zinc-800 flex justify-between items-center text-xs font-display">
        <span className="font-semibold text-zinc-300 flex items-center gap-2">
          <svg className="w-4 h-4 text-emerald-500 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
            <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
          </svg>
          Xem trước: {title || "Tài liệu"}
        </span>
        {loading && <span className="text-emerald-400 animate-pulse font-medium">Đang tải bản xem trước...</span>}
      </div>
      <div className="flex-1 overflow-auto p-4 bg-zinc-800 text-black max-h-[650px] min-h-[500px]">
        {error && (
          <div className="h-full min-h-[400px] flex items-center justify-center text-zinc-400 text-sm">
            <span className="mr-2">⚠️</span> {error}
          </div>
        )}
        {!error && loading && (
          <div className="h-full min-h-[400px] flex items-center justify-center text-zinc-400 text-sm">
            <div className="w-6 h-6 border-2 border-emerald-500/30 border-t-emerald-500 rounded-full animate-spin mr-2"></div>
            Đang giải nén văn bản...
          </div>
        )}
        <div ref={containerRef} className="docx-container" style={{ display: loading || error ? 'none' : 'block' }} />
      </div>
    </div>
  );
}

export default function App() {
  const [userRole, setUserRole] = useState('user'); // user | congchung
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [sidebarActiveMenu, setSidebarActiveMenu] = useState('templates'); // templates | guide | submissions | stats
  const [submissionHistory, setSubmissionHistory] = useState([]);
  const [isLoadingSubmissions, setIsLoadingSubmissions] = useState(false);
  const [expandedSubmissionId, setExpandedSubmissionId] = useState(null);
  const [submissionFilesMap, setSubmissionFilesMap] = useState({});
  const [loadingFilesForSubId, setLoadingFilesForSubId] = useState(null);
  const [submissionDetailsMap, setSubmissionDetailsMap] = useState({});
  const [loadingDetailsForSubId, setLoadingDetailsForSubId] = useState(null);
  const [detailActiveTab, setDetailActiveTab] = useState('fields'); // fields | preview
  const [activePreviewFilename, setActivePreviewFilename] = useState(null);

  const [activeView, setActiveView] = useState('dashboard'); // dashboard | config | fill | success
  const [currentStep, setCurrentStep] = useState(1);
  const totalSteps = 4;
  const [showPdfModal, setShowPdfModal] = useState(false);
  const [selectedChildIds, setSelectedChildIds] = useState([]);
  const [hasUnsavedManual, setHasUnsavedManual] = useState(false);
  const [previewKey, setPreviewKey] = useState(0);
  const [templates, setTemplates] = useState([]);
  const [categories, setCategories] = useState([]);
  const [selectedTemplate, setSelectedTemplate] = useState(null);
  const [fields, setFields] = useState([]);
  const [fieldsHistory, setFieldsHistory] = useState([]);
  const [historyIndex, setHistoryIndex] = useState(-1);

  const updateFieldsAndHistory = (newFields) => {
    setFields(newFields);
    const nextHistory = fieldsHistory.slice(0, historyIndex + 1);
    setFieldsHistory([...nextHistory, newFields]);
    setHistoryIndex(nextHistory.length);
  };

  const handleUndo = () => {
    if (historyIndex > 0) {
      const prevIndex = historyIndex - 1;
      setHistoryIndex(prevIndex);
      setFields(fieldsHistory[prevIndex]);
      setHasUnsavedManual(true);
      showNotification('Đã hoàn tác thao tác cấu hình!');
    }
  };

  const handleResetForm = () => {
    if (window.confirm("Bạn có chắc chắn muốn nhập lại toàn bộ thông tin? Dữ liệu đã nhập sẽ bị xóa.")) {
      const resetData = {};
      fields.forEach(f => {
        resetData[f.key_name] = f.field_type === 'boolean' ? false : '';
      });
      setFormData(resetData);
      setCustomerName('');
      setCustomerPhone('');
      setCurrentStep(1);
      showNotification("Đã reset toàn bộ form điền.");
    }
  };

  const handleAddRepeatedRecord = (childTempId) => {
    const records = formData[childTempId] || [];
    const newRecord = { _id: `rec-${Date.now()}-${records.length}` };
    fields.forEach(f => {
      if (f.childTemplateId === childTempId) {
        newRecord[f.key_name] = f.field_type === 'boolean' ? false : '';
      }
    });
    setFormData(prev => ({
      ...prev,
      [childTempId]: [...records, newRecord]
    }));
    showNotification('Đã thêm một tài sản/căn nhà mới!');
  };

  const handleRemoveRepeatedRecord = (childTempId, recordIndex) => {
    const records = formData[childTempId] || [];
    if (records.length <= 1) {
      showNotification('Phải có ít nhất một tài sản/căn nhà.', 'error');
      return;
    }
    if (window.confirm('Bạn có chắc muốn xóa tài sản/căn nhà này?')) {
      const updated = records.filter((_, idx) => idx !== recordIndex);
      setFormData(prev => ({
        ...prev,
        [childTempId]: updated
      }));
      showNotification('Đã xóa tài sản/căn nhà.');
    }
  };

  const handleUpdateRepeatedField = (childTempId, recordIndex, keyName, value) => {
    setFormData(prev => {
      const records = [...(prev[childTempId] || [])];
      if (records[recordIndex]) {
        records[recordIndex] = {
          ...records[recordIndex],
          [keyName]: value
        };
      }
      return {
        ...prev,
        [childTempId]: records
      };
    });
  };

  const handleToggleRepeated = async (childId, isRepeated) => {
    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${childId}/repeated`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ isRepeated })
      });
      if (!res.ok) throw new Error('Không thể cập nhật chế độ lặp');
      showNotification('Đã cập nhật chế độ lặp biểu mẫu thành công!');
      
      setLinkedChildren(prev => prev.map(t => t.id === childId ? { ...t, is_repeated: isRepeated ? 1 : 0 } : t));
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };



  // States for linking templates
  const [showLinkModal, setShowLinkModal] = useState(false);
  const [linkingTemplate, setLinkingTemplate] = useState(null);
  const [linkedChildren, setLinkedChildren] = useState([]);
  const [availableToLink, setAvailableToLink] = useState([]);
  const [parentFields, setParentFields] = useState([]);
  const [linkSearch, setLinkSearch] = useState('');

  // States for dashboard search & filter
  const [dashboardSearch, setDashboardSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('all'); // all | active | draft
  const [sortBy, setSortBy] = useState('newest'); // newest | oldest | name
  const [selectedCategoryId, setSelectedCategoryId] = useState('all');
  const [expandedCategoryIds, setExpandedCategoryIds] = useState(new Set());
  
  // States for editing template title
  const [editingTemplateId, setEditingTemplateId] = useState(null);
  const [editingNameValue, setEditingNameValue] = useState('');
  
  // States for uploading
  const [uploadFiles, setUploadFiles] = useState([]);
  const [uploadName, setUploadName] = useState('');
  const [uploadCategoryId, setUploadCategoryId] = useState('');
  const [isUploading, setIsUploading] = useState(false);
  const [isDragActive, setIsDragActive] = useState(false);
  const [uploadResults, setUploadResults] = useState([]);
  const [newCategoryName, setNewCategoryName] = useState('');
  const [newCategoryParentId, setNewCategoryParentId] = useState('');
  const [mobileFillTab, setMobileFillTab] = useState('form'); // form | preview
  const [mobileSuccessTab, setMobileSuccessTab] = useState('success'); // success | preview

  const getLogicalGroupName = (field) => {
    if (field.groupName) {
      return `Biểu mẫu con: ${field.groupName}`;
    }
    
    const key = (field.key_name || '').toLowerCase();
    const label = (field.label || '').toLowerCase();
    
    if (
      key.includes('ban_') || key.includes('chuyennhuong_') || key.includes('tangcho_') ||
      key.includes('bena_') || label.includes('bên bán') || label.includes('chuyển nhượng') ||
      label.includes('bên a') || label.includes('tặng cho')
    ) {
      return '1. Bên chuyển nhượng / Bên bán / Bên tặng cho (Bên A)';
    }
    
    if (
      key.includes('mua_') || key.includes('nhanchuyennhuong_') || key.includes('nhantangcho_') ||
      key.includes('benb_') || label.includes('bên mua') || label.includes('nhận chuyển nhượng') ||
      label.includes('bên b') || label.includes('nhận tặng cho')
    ) {
      return '2. Bên nhận chuyển nhượng / Bên mua / Bên nhận (Bên B)';
    }
    
    if (
      key.includes('taisan_') || key.includes('dat_') || key.includes('nha_') ||
      key.includes('xe_') || key.includes('batdongsan_') || label.includes('tài sản') ||
      label.includes('đất') || label.includes('nhà') || label.includes('bất động sản') ||
      label.includes('căn hộ') || label.includes('xe') || label.includes('biển số')
    ) {
      return '3. Thông tin tài sản (Đất đai, Nhà ở, Phương tiện...)';
    }
    
    if (
      key.includes('hopdong_') || key.includes('gia_') || key.includes('thanhtoan_') ||
      key.includes('phi_') || key.includes('ngayky_') || label.includes('hợp đồng') ||
      label.includes('giá') || label.includes('thanh toán') || label.includes('phí') ||
      label.includes('ngày ký')
    ) {
      return '4. Thông tin hợp đồng & Thanh toán';
    }
    
    return '5. Thông tin khác';
  };

  const getGroupedFields = (fieldsList) => {
    const groups = {};
    fieldsList.forEach(field => {
      const gName = getLogicalGroupName(field);
      if (!groups[gName]) groups[gName] = [];
      groups[gName].push(field);
    });
    return groups;
  };


  // States for submitting form
  const [customerName, setCustomerName] = useState('');
  const [customerPhone, setCustomerPhone] = useState('');
  const [formData, setFormData] = useState({});
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submissionResult, setSubmissionResult] = useState(null);
  const [successSelectedFiles, setSuccessSelectedFiles] = useState([]);
  const [successPreviewFilename, setSuccessPreviewFilename] = useState('');

  // States for submission progress tracking (for ZIP export with multiple files)
  const [submissionProgress, setSubmissionProgress] = useState({ current: 0, total: 0, currentFile: '' });

  // States for live preview switching during filling
  const [previewTemplateId, setPreviewTemplateId] = useState(null);
  const [previewFields, setPreviewFields] = useState([]);

  // Shared states
  const [isLoading, setIsLoading] = useState(false);
  const [notification, setNotification] = useState(null);

  // States for table detection
  const [detectedTables, setDetectedTables] = useState([]);
  const [isLoadingTables, setIsLoadingTables] = useState(false);
  const [activeTableIndex, setActiveTableIndex] = useState(null);
  const [tableFieldEdits, setTableFieldEdits] = useState([]);
  const [selectedTableRows, setSelectedTableRows] = useState([]);
  const [isInjectingTable, setIsInjectingTable] = useState(false);

  // States for resizable split panel
  const [leftWidth, setLeftWidth] = useState(40);
  const [isMobile, setIsMobile] = useState(typeof window !== 'undefined' ? window.innerWidth < 1024 : false);
  const splitContainerRef = useRef(null);

  useEffect(() => {
    const handleResize = () => {
      setIsMobile(window.innerWidth < 1024);
    };
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  const startResizing = (e) => {
    e.preventDefault();
    const handleMouseMove = (moveEvent) => {
      if (splitContainerRef.current) {
        const containerRect = splitContainerRef.current.getBoundingClientRect();
        const newLeftWidth = ((moveEvent.clientX - containerRect.left) / containerRect.width) * 100;
        // Giới hạn chiều rộng từ 25% đến 75%
        if (newLeftWidth >= 25 && newLeftWidth <= 75) {
          setLeftWidth(newLeftWidth);
        }
      }
    };

    const handleMouseUp = () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };

    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
  };

  // States for text selection and quick variable tagging
  const [selectedText, setSelectedText] = useState('');
  const [paragraphContext, setParagraphContext] = useState('');
  const [quickKey, setQuickKey] = useState('');
  const [quickLabel, setQuickLabel] = useState('');
  const [quickType, setQuickType] = useState('text');

  useEffect(() => {
    const handleSelection = () => {
      const selection = window.getSelection();
      const text = selection ? selection.toString().trim() : '';
      const container = document.querySelector('.docx-container');
      
      if (text && container && container.contains(selection.anchorNode)) {
        // Tìm element chứa node bôi đen
        let parentElement = selection.anchorNode.parentElement;
        // Đi ngược lên cây DOM tìm thẻ Paragraph (P hoặc DIV) chứa văn bản này
        while (
          parentElement && 
          parentElement.tagName !== 'P' && 
          parentElement.tagName !== 'DIV' && 
          parentElement !== container
        ) {
          parentElement = parentElement.parentElement;
        }
        
        const contextText = parentElement ? parentElement.textContent.trim() : '';
        setSelectedText(text);
        setParagraphContext(contextText);
      }
    };

    document.addEventListener('selectionchange', handleSelection);
    return () => document.removeEventListener('selectionchange', handleSelection);
  }, []);

  const handleAddQuickField = (e) => {
    e.preventDefault();
    if (!quickKey || !quickLabel) {
      showNotification('Vui lòng điền đầy đủ tên biến và nhãn hiển thị', 'error');
      return;
    }

    const cleanKey = quickKey.trim().toLowerCase().replace(/[^a-z0-9_]/g, '_');
    if (fields.some(f => f.key_name === cleanKey)) {
      showNotification('Tên biến này đã tồn tại trong biểu mẫu', 'error');
      return;
    }

    const newField = {
      id: generateId(),
      key_name: cleanKey,
      label: quickLabel.trim(),
      field_type: quickType,
      is_required: 1,
      order_index: fields.length,
      replace_text: selectedText,
      paragraph_context: paragraphContext
    };

    updateFieldsAndHistory([...fields, newField]);
    setSelectedText('');
    setParagraphContext('');
    setQuickKey('');
    setQuickLabel('');
    showNotification(`Đã bóc tách thành công chữ "${selectedText}" thành biến {{${cleanKey}}}!`);
    window.getSelection()?.removeAllRanges();
  };

  const [showManualAdd, setShowManualAdd] = useState(false);

  const handleOpenManualAdd = () => {
    setShowManualAdd(true);
    setSelectedText('');
    setParagraphContext('');
    setQuickKey('');
    setQuickLabel('');
    setQuickType('text');
  };

  const fieldListRef = useRef(null);

  const handleAddManualField = (e) => {
    e.preventDefault();
    if (!quickKey || !quickLabel) {
      showNotification('Vui lòng điền đầy đủ tên biến và nhãn hiển thị', 'error');
      return;
    }

    const cleanKey = quickKey.trim().toLowerCase().replace(/[^a-z0-9_]/g, '_');
    if (fields.some(f => f.key_name === cleanKey)) {
      showNotification('Tên biến này đã tồn tại trong biểu mẫu', 'error');
      return;
    }

    const newField = {
      id: generateId(),
      key_name: cleanKey,
      label: quickLabel.trim(),
      field_type: quickType,
      is_required: 1,
      order_index: fields.length,
      replace_text: null,
      paragraph_context: null
    };

    updateFieldsAndHistory([...fields, newField]);
    setHasUnsavedManual(true);
    setShowManualAdd(false);
    setQuickKey('');
    setQuickLabel('');
    showNotification(`✅ Đã thêm biến {{${cleanKey}}}! Nhấn "Lưu & Kích hoạt" để lưu lại.`);

    // Cuộn xuống field mới trong container danh sách
    setTimeout(() => {
      const container = fieldListRef.current;
      const el = document.getElementById(`field-config-${cleanKey}`);
      if (el) {
        if (container) {
          container.scrollTop = el.offsetTop - container.offsetTop - 16;
        } else {
          el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
        el.style.transition = 'box-shadow 0.3s ease';
        el.style.boxShadow = '0 0 0 3px #f59e0b, 0 0 0 5px rgba(245,158,11,0.2)';
        setTimeout(() => { el.style.boxShadow = ''; }, 2000);
      }
    }, 120);
  };


  const showNotification = useCallback((message, type = 'success') => {
    setNotification({ message, type });
    setTimeout(() => setNotification(null), 5000);
  }, []);

  const fetchCategories = useCallback(async () => {
    try {
      const res = await fetch(`${API_BASE}/template-categories`);
      if (!res.ok) throw new Error('Không thể lấy danh mục biểu mẫu');
      const data = await res.json();
      setCategories(data);
    } catch (err) {
      showNotification(err.message, 'error');
    }
  }, [showNotification]);

  const fetchTemplates = useCallback(async () => {
    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates`);
      if (!res.ok) throw new Error('Không thể lấy danh sách biểu mẫu');
      const data = await res.json();
      setTemplates(data);
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  }, [showNotification, setIsLoading, setTemplates]);

  // Load templates on mount
  useEffect(() => {
    let active = true;
    Promise.resolve().then(() => {
      if (active) {
        fetchTemplates();
        fetchCategories();
      }
    });
    return () => {
      active = false;
    };
  }, [fetchTemplates, fetchCategories]);

  const fetchSubmissions = useCallback(async () => {
    setIsLoadingSubmissions(true);
    try {
      const res = await fetch(`${API_BASE}/submissions`);
      if (!res.ok) throw new Error('Không thể tải danh sách hồ sơ đã nhận.');
      const data = await res.json();
      setSubmissionHistory(data);
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoadingSubmissions(false);
    }
  }, [showNotification]);

  // Load submissions automatically when switching to congchung role or submissions view
  useEffect(() => {
    if (userRole === 'congchung' && sidebarActiveMenu === 'submissions') {
      fetchSubmissions();
    }
  }, [userRole, sidebarActiveMenu, fetchSubmissions]);

  const handleToggleExpandSubmission = async (subId) => {
    if (expandedSubmissionId === subId) {
      setExpandedSubmissionId(null);
      setActivePreviewFilename(null);
      return;
    }
    setExpandedSubmissionId(subId);
    setActivePreviewFilename(null);
    if (!submissionFilesMap[subId]) {
      setLoadingFilesForSubId(subId);
      try {
        const res = await fetch(`${API_BASE}/submissions/${subId}/files`);
        if (res.ok) {
          const files = await res.json();
          setSubmissionFilesMap(prev => ({ ...prev, [subId]: files }));
        }
      } catch (err) {
        console.error("Lỗi khi tải danh sách file:", err);
      } finally {
        setLoadingFilesForSubId(null);
      }
    }
    if (!submissionDetailsMap[subId]) {
      setLoadingDetailsForSubId(subId);
      try {
        const res = await fetch(`${API_BASE}/submissions/${subId}/detail`);
        if (res.ok) {
          const detail = await res.json();
          setSubmissionDetailsMap(prev => ({ ...prev, [subId]: detail }));
        }
      } catch (err) {
        console.error("Lỗi khi tải chi tiết hồ sơ:", err);
      } finally {
        setLoadingDetailsForSubId(null);
      }
    }
  };

  const handleOpenLinkModal = async (template) => {
    setLinkingTemplate(template);
    setIsLoading(true);
    try {
      const resLinks = await fetch(`${API_BASE}/templates/${template.id}/links`);
      if (!resLinks.ok) throw new Error('Không thể lấy danh sách file liên kết');
      const linksData = await resLinks.json();
      setLinkedChildren(linksData);

      // Available to link
      const available = templates.filter(t => 
        t.id !== template.id && 
        !t.parent_template_id && 
        !linksData.some(linked => linked.id === t.id)
      );
      setAvailableToLink(available);
      setShowLinkModal(true);
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleLinkChild = async (childId) => {
    if (!linkingTemplate) return;
    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${linkingTemplate.id}/link`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ childTemplateId: childId })
      });
      if (!res.ok) throw new Error('Không thể thiết lập liên kết');
      showNotification('Đã liên kết biểu mẫu con thành công!');
      
      const childObj = templates.find(t => t.id === childId);
      if (childObj) {
        setLinkedChildren(prev => [...prev, childObj]);
        setAvailableToLink(prev => prev.filter(t => t.id !== childId));
        setTemplates(prev => prev.map(t => t.id === childId ? { ...t, parent_template_id: linkingTemplate.id } : t));
      }
      fetchTemplates();
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleUnlinkChild = async (childId) => {
    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${childId}/unlink`, {
        method: 'POST'
      });
      if (!res.ok) throw new Error('Không thể hủy liên kết');
      showNotification('Đã hủy liên kết biểu mẫu con!');
      
      setLinkedChildren(prev => prev.filter(t => t.id !== childId));
      const childObj = templates.find(t => t.id === childId);
      if (childObj) {
        setAvailableToLink(prev => [...prev, { ...childObj, parent_template_id: null }]);
        setTemplates(prev => prev.map(t => t.id === childId ? { ...t, parent_template_id: null } : t));
      }
      fetchTemplates();
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const validateUploadFiles = useCallback((incomingFiles) => {
    if (!incomingFiles || incomingFiles.length === 0) {
      return { validFiles: [], error: 'Vui lòng chọn ít nhất một file .doc hoặc .docx.' };
    }

    if (incomingFiles.length > MAX_UPLOAD_FILES) {
      return { validFiles: [], error: `Mỗi lần chỉ được tải tối đa ${MAX_UPLOAD_FILES} file.` };
    }

    for (const file of incomingFiles) {
      if (!isSupportedWordFile(file.name)) {
        return { validFiles: [], error: 'Hệ thống chỉ chấp nhận file Word định dạng .doc hoặc .docx' };
      }

      if (file.size > MAX_UPLOAD_FILE_SIZE) {
        return { validFiles: [], error: `File ${file.name} vượt quá giới hạn 1MB.` };
      }
    }

    return { validFiles: incomingFiles, error: null };
  }, []);

  const handleSelectUploadFiles = useCallback((fileList) => {
    const incomingFiles = Array.from(fileList || []);
    const { validFiles, error } = validateUploadFiles(incomingFiles);

    if (error) {
      showNotification(error, 'error');
      return false;
    }

    setUploadFiles(validFiles);
    setUploadResults([]);
    if (validFiles.length === 1 && !uploadName) {
      setUploadName(getWordBaseName(validFiles[0].name));
    }
    if (validFiles.length > 1 && uploadName) {
      setUploadName('');
    }
    return true;
  }, [showNotification, uploadName, validateUploadFiles]);

  const handleFileUpload = async (e) => {
    e.preventDefault();
    if (uploadFiles.length === 0) {
      showNotification('Vui lòng chọn ít nhất một file .doc hoặc .docx', 'error');
      return;
    }

    const { error } = validateUploadFiles(uploadFiles);
    if (error) {
      showNotification(error, 'error');
      return;
    }

    setIsUploading(true);
    setUploadResults([]);
    const formDataObj = new FormData();
    uploadFiles.forEach(file => {
      formDataObj.append('templateFiles', file);
    });
    if (uploadFiles.length === 1 && uploadName.trim()) {
      formDataObj.append('name', uploadName.trim());
    }
    if (uploadCategoryId) {
      formDataObj.append('categoryId', uploadCategoryId);
    }

    try {
      const res = await fetch(`${API_BASE}/templates`, {
        method: 'POST',
        body: formDataObj,
      });

      if (!res.ok) {
        let errMsg = 'Upload file thất bại';
        try {
          const data = await res.json();
          errMsg = data.error || errMsg;
        } catch (parseErr) {
          console.warn('Lỗi parse JSON:', parseErr);
        }
        throw new Error(errMsg);
      }

      const data = await res.json();
      const results = Array.isArray(data.results) ? data.results : [];
      setUploadResults(results);

      if ((data.successCount || 0) > 0) {
        showNotification(data.message || 'Tải file mẫu lên và trích xuất các biến thành công!');
      } else {
        showNotification(data.message || 'Không có file nào được tải thành công.', 'error');
      }

      setUploadFiles([]);
      setUploadName('');
      setUploadCategoryId('');
      fetchTemplates();

      if (results.length === 1 && results[0].status === 'success') {
        handleOpenConfig(results[0].templateId);
      }
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsUploading(false);
    }
  };

  const handleCreateCategory = async (e) => {
    e.preventDefault();
    if (!newCategoryName.trim()) {
      showNotification('Tên danh mục không được để trống', 'error');
      return;
    }

    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/template-categories`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name: newCategoryName.trim(),
          parentId: newCategoryParentId || null
        })
      });

      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || 'Không thể tạo danh mục biểu mẫu');
      }

      showNotification('Đã tạo danh mục biểu mẫu!');
      setNewCategoryName('');
      setNewCategoryParentId('');
      fetchCategories();
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleUpdateTemplateCategory = async (templateId, categoryId) => {
    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${templateId}/category`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ categoryId: categoryId || null })
      });

      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || 'Không thể cập nhật danh mục biểu mẫu');
      }

      setTemplates(prev => prev.map(t => {
        if (t.id !== templateId) return t;
        const category = categories.find(c => c.id === categoryId);
        return {
          ...t,
          category_id: categoryId || null,
          category_name: category ? category.name : null,
          category_parent_id: category ? category.parent_id : null
        };
      }));
      showNotification('Đã cập nhật danh mục biểu mẫu!');
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleSaveTemplateName = async (templateId) => {
    if (!editingNameValue || editingNameValue.trim() === '') {
      showNotification('Tên biểu mẫu không được để trống', 'error');
      return;
    }

    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${templateId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: editingNameValue.trim() }),
      });

      if (!res.ok) {
        let errMsg = 'Cập nhật tên biểu mẫu thất bại';
        try {
          const data = await res.json();
          errMsg = data.error || errMsg;
        } catch (parseErr) {
          console.warn('Lỗi parse JSON:', parseErr);
        }
        throw new Error(errMsg);
      }

      showNotification('Cập nhật tên biểu mẫu thành công!');
      setEditingTemplateId(null);
      fetchTemplates();
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleDeleteTemplate = async (templateId, templateName) => {
    if (!window.confirm(`⚠️ CẢNH BÁO CỰC KỲ QUAN TRỌNG:\n\nBạn có chắc chắn muốn xóa biểu mẫu "${templateName}" không?\nHành động này sẽ xóa hoàn toàn:\n1. Biểu mẫu và cấu hình các trường biến.\n2. Tất cả hồ sơ đã nộp của khách hàng liên quan đến biểu mẫu này.\n3. Các tệp tin gốc và tệp kết quả vật lý trên ổ đĩa.\n\nHành động này KHÔNG THỂ KHÔI PHỤC.`)) {
      return;
    }

    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${templateId}`, {
        method: 'DELETE',
      });

      if (!res.ok) {
        let errMsg = 'Xóa biểu mẫu thất bại';
        try {
          const data = await res.json();
          errMsg = data.error || errMsg;
        } catch (parseErr) {
          console.warn('Lỗi parse JSON:', parseErr);
        }
        throw new Error(errMsg);
      }

      showNotification('Đã xóa biểu mẫu thành công!');
      fetchTemplates();
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleDuplicateTemplate = async (templateId) => {
    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${templateId}/duplicate`, {
        method: 'POST',
      });

      if (!res.ok) {
        let errMsg = 'Sao chép biểu mẫu thất bại';
        try {
          const data = await res.json();
          errMsg = data.error || errMsg;
        } catch (parseErr) {
          console.warn('Lỗi parse JSON:', parseErr);
        }
        throw new Error(errMsg);
      }

      showNotification('Đã sao chép biểu mẫu thành công!');
      fetchTemplates();
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleExportConfig = async () => {
    if (!selectedTemplate) return;
    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${selectedTemplate.id}/export`);
      if (!res.ok) throw new Error('Không thể xuất cấu hình.');
      const data = await res.json();
      
      const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(data.fields, null, 2));
      const downloadAnchor = document.createElement('a');
      downloadAnchor.setAttribute("href", dataStr);
      const safeName = selectedTemplate.name.replace(/[^a-zA-Z0-9À-ỹ\s-_]/g, '').trim().replace(/\s+/g, '_');
      downloadAnchor.setAttribute("download", `Config_${safeName}.json`);
      document.body.appendChild(downloadAnchor);
      downloadAnchor.click();
      downloadAnchor.remove();
      showNotification('Đã xuất file cấu hình thành công!');
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleImportConfig = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const importedFields = JSON.parse(event.target.result);
        if (!Array.isArray(importedFields)) {
          throw new Error('Định dạng file cấu hình không hợp lệ. Phải là một mảng các trường.');
        }

        const isValid = importedFields.every(f => typeof f.key_name === 'string');
        if (!isValid) {
          throw new Error('Dữ liệu các trường trong file thiếu thông tin bắt buộc (key_name).');
        }

        setIsLoading(true);
        const res = await fetch(`${API_BASE}/templates/${selectedTemplate.id}/import`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ fields: importedFields }),
        });

        if (!res.ok) {
          const data = await res.json();
          throw new Error(data.error || 'Import cấu hình thất bại.');
        }

        showNotification('Nhập cấu hình mới thành công! Đang tải lại...');
        const formRes = await fetch(`${API_BASE}/templates/${selectedTemplate.id}/form`);
        if (formRes.ok) {
          const formData = await formRes.json();
          setFields(formData.fields || []);
          setFieldsHistory([formData.fields || []]);
          setHistoryIndex(0);
        }
      } catch (err) {
        showNotification(err.message, 'error');
      } finally {
        setIsLoading(false);
        e.target.value = null;
      }
    };
    reader.readAsText(file);
  };

  // ---------- Table Detection Handlers ----------
  const fetchDetectedTables = async (templateId) => {
    setIsLoadingTables(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${templateId}/tables`);
      if (!res.ok) throw new Error('Không thể quét bảng từ tài liệu');
      const data = await res.json();
      setDetectedTables(data.tables || []);
    } catch (err) {
      console.error('Table scan error:', err);
      setDetectedTables([]);
    } finally {
      setIsLoadingTables(false);
    }
  };

  const handleOpenTablePanel = (table) => {
    if (activeTableIndex === table.tableIndex) {
      setActiveTableIndex(null);
      return;
    }
    setActiveTableIndex(table.tableIndex);
    setTableFieldEdits(table.suggestedFields.map(f => ({ ...f })));
    setSelectedTableRows(table.rows.map(r => r.rowIndex));
  };

  const handleTableRowClick = (tableIndex, rowIndex) => {
    const table = detectedTables.find(t => t.tableIndex === tableIndex);
    if (table) {
      setActiveTableIndex(tableIndex);
      setTableFieldEdits(table.suggestedFields.map(f => ({ ...f })));
      if (rowIndex >= 1) {
        setSelectedTableRows([rowIndex]);
      } else {
        setSelectedTableRows(table.rows.map(r => r.rowIndex));
      }
      showNotification(`🔍 Đã nhận diện Bảng #${tableIndex + 1}! Hãy thiết lập các biến ở Panel bên trái.`);
      
      setTimeout(() => {
        const accordionEl = document.getElementById(`table-accordion-${tableIndex}`);
        if (accordionEl) {
          accordionEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
          accordionEl.classList.add('ring-2', 'ring-amber-500');
          setTimeout(() => {
            accordionEl.classList.remove('ring-2', 'ring-amber-500');
          }, 2000);
        }
      }, 100);
    }
  };

  const handleTableFieldEdit = (colIdx, key, value) => {
    setTableFieldEdits(prev => prev.map(f =>
      f.colIndex === colIdx ? { ...f, [key]: value } : f
    ));
  };

  const handleInjectTable = async () => {
    if (!selectedTemplate || activeTableIndex === null) return;
    if (selectedTableRows.length === 0) {
      showNotification('Vui lòng chọn ít nhất một dòng để gài biến.', 'error');
      return;
    }
    setIsInjectingTable(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${selectedTemplate.id}/inject-table`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          tableIndex: activeTableIndex,
          fields: tableFieldEdits.map(f => ({
            colIndex: f.colIndex,
            key_name: f.key_name,
            label: f.label,
            field_type: f.field_type,
            is_required: true
          })),
          selectedRows: selectedTableRows
        })
      });
      if (!res.ok) {
        const errData = await res.json().catch(() => ({}));
        throw new Error(errData.error || 'Lỗi gài placeholder bảng');
      }
      showNotification('Đã gài placeholder và tạo các field cho bảng thành công!');
      setActiveTableIndex(null);
      setDetectedTables([]);
      setPreviewKey(prev => prev + 1);
      // Refresh fields list and re-scan tables
      handleOpenConfig(selectedTemplate.id);
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsInjectingTable(false);
    }
  };

  const handleOpenConfig = async (templateId) => {
    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${templateId}/form`);
      if (!res.ok) {
        let errMsg = 'Không thể tải biểu mẫu';
        try {
          const data = await res.json();
          errMsg = data.error || errMsg;
        } catch (parseErr) {
          console.warn('Lỗi parse JSON:', parseErr);
        }
        throw new Error(errMsg);
      }
      const data = await res.json();
      
      setSelectedTemplate(data.template);
      setFields(data.fields);

      // Tải danh sách trường từ biểu mẫu gốc (nếu có)
      if (data.template.parent_template_id) {
        try {
          const resParent = await fetch(`${API_BASE}/templates/${templateId}/parent-fields`);
          if (resParent.ok) {
            const parentFieldsData = await resParent.json();
            setParentFields(parentFieldsData);
          } else {
            setParentFields([]);
          }
        } catch (parentErr) {
          console.error(parentErr);
          setParentFields([]);
        }
      } else {
        setParentFields([]);
      }

      setActiveView('config');
      // Auto-fetch detected tables for the template
      fetchDetectedTables(templateId);
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleFieldChange = (index, key, value) => {
    const updated = [...fields];
    updated[index] = { ...updated[index], [key]: value };
    updateFieldsAndHistory(updated);
  };

  // Phát hiện các trùng lặp khi ánh xạ biến con vào biến mẹ
  // Trả về: Map<parent_field_key, [field, ...]> chỉ chứa các parent_field_key bị nhiều child cùng trỏ vào
  const getMappingConflicts = () => {
    const map = new Map();
    fields.forEach(f => {
      if (!f.parent_field_key) return;
      const list = map.get(f.parent_field_key) || [];
      list.push(f);
      map.set(f.parent_field_key, list);
    });
    const conflicts = new Map();
    map.forEach((list, key) => {
      if (list.length > 1) conflicts.set(key, list);
    });
    return conflicts;
  };

  const handleSaveConfig = async () => {
    // Validate ánh xạ trước khi gửi
    const conflicts = getMappingConflicts();
    if (conflicts.size > 0) {
      const lines = [];
      conflicts.forEach((list, parentKey) => {
        const childKeys = list.map(f => `{{${f.key_name}}}`).join(', ');
        lines.push(`{{${parentKey}}} ← ${childKeys}`);
      });
      const ok = window.confirm(
        `Có ${conflicts.size} biến của file gốc đang được nhiều biến con cùng ánh xạ:\n\n` +
        lines.join('\n') +
        `\n\nKhi điền form, các biến con này sẽ nhận cùng giá trị từ biến gốc — có thể không phải ý bạn muốn.\n\nNhấn OK để vẫn lưu, hoặc Hủy để chỉnh lại.`
      );
      if (!ok) return;
    }

    // Cảnh báo khi biến con trùng tên với biến mẹ nhưng vẫn cố ánh xạ thủ công vào biến mẹ khác
    const matchingButRemapped = fields.filter(f => {
      if (!f.parent_field_key) return false;
      const sameNameInParent = parentFields.some(pf => pf.key_name === f.key_name);
      return sameNameInParent && f.parent_field_key !== f.key_name;
    });
    if (matchingButRemapped.length > 0) {
      const list = matchingButRemapped.map(f => `{{${f.key_name}}} → {{${f.parent_field_key}}}`).join('\n');
      const ok = window.confirm(
        `Một số biến con có cùng tên với biến mẹ nhưng đang được ánh xạ thủ công vào biến khác:\n\n` +
        list +
        `\n\nĐiều này sẽ ghi đè cơ chế tự động đồng bộ theo tên trùng. Bạn có chắc muốn lưu?`
      );
      if (!ok) return;
    }

    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${selectedTemplate.id}/fields`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields }),
      });
      if (!res.ok) {
        let errMsg = 'Lưu cấu hình thất bại';
        try {
          const data = await res.json();
          errMsg = data.error || errMsg;
        } catch (parseErr) {
          console.warn('Lỗi parse JSON:', parseErr);
        }
        throw new Error(errMsg);
      }

      setHasUnsavedManual(false);
      showNotification('Lưu cấu hình biểu mẫu và kích hoạt thành công!');
      fetchTemplates();
      setActiveView('dashboard');
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleRestoreField = async (fieldId, idx) => {
    const field = fields[idx];
    if (!field.id || field.id.startsWith('temp-')) {
      const updatedFields = fields.filter((_, i) => i !== idx);
      updateFieldsAndHistory(updatedFields);
      showNotification("Đã xóa cấu hình biến tạm.");
      return;
    }

    if (!window.confirm(`Bạn có chắc chắn muốn khôi phục biến {{${field.key_name}}} thành chữ gốc "${field.replace_text || ''}" và xóa cấu hình biến này khỏi tệp Word không?`)) {
      return;
    }
    
    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${selectedTemplate.id}/fields/${field.id}/restore`, {
        method: 'POST'
      });
      if (!res.ok) {
        let errMsg = 'Khôi phục biến thất bại';
        try {
          const data = await res.json();
          errMsg = data.error || errMsg;
        } catch (parseErr) {
          console.warn('Lỗi parse JSON:', parseErr);
        }
        throw new Error(errMsg);
      }
      const data = await res.json();
      
      showNotification(data.message);
      
      const updatedFields = fields.filter((_, i) => i !== idx);
      updateFieldsAndHistory(updatedFields);
      
      setPreviewKey(prev => prev + 1);
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleDeleteField = (idx) => {
    const field = fields[idx];
    const confirmMsg = field.replace_text 
      ? `Bạn có chắc chắn muốn xóa cấu hình biến {{${field.key_name}}} (vẫn giữ nguyên {{${field.key_name}}} trong file Word) không?`
      : `Bạn có chắc chắn muốn xóa biến thủ công {{${field.key_name}}} khỏi danh sách cấu hình không?`;
    if (window.confirm(confirmMsg)) {
      const updated = fields.filter((_, i) => i !== idx);
      updateFieldsAndHistory(updated);
      showNotification("Đã tạm xóa cấu hình biến. Hãy nhấn 'Lưu & Kích hoạt' để đồng bộ.");
    }
  };


  const handleOpenFill = async (templateId) => {
    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${templateId}/form`);
      if (!res.ok) {
        let errMsg = 'Không thể tải biểu mẫu';
        try {
          const data = await res.json();
          errMsg = data.error || errMsg;
        } catch (parseErr) {
          console.warn('Lỗi parse JSON:', parseErr);
        }
        throw new Error(errMsg);
      }
      const data = await res.json();

      setSelectedTemplate(data.template);
      
      const masterFields = data.fields;
      let finalFields = [...masterFields];
      
      // Lấy danh sách file con liên kết để gom trường riêng biệt
      let currentLinks = [];
      try {
        const resLinks = await fetch(`${API_BASE}/templates/${templateId}/links`);
        if (resLinks.ok) {
          const linksData = await resLinks.json();
          currentLinks = linksData;
          setLinkedChildren(linksData);
          setSelectedChildIds(linksData.map(c => c.id));
          for (const child of linksData) {
            const resChildForm = await fetch(`${API_BASE}/templates/${child.id}/form`);
            if (resChildForm.ok) {
              const childData = await resChildForm.json();
              // Lọc: bỏ trường trùng tên biến với master hoặc trường đã ánh xạ vào biến mẹ
              const childUniqueFields = childData.fields.filter(cf => {
                const matchesMasterKey = masterFields.some(mf => mf.key_name === cf.key_name);
                const isMapped = cf.parent_field_key !== null && cf.parent_field_key !== '';
                return !matchesMasterKey && !isMapped;
              }).map(cf => ({
                ...cf,
                groupName: child.name,
                childTemplateId: child.id
              }));
              finalFields = [...finalFields, ...childUniqueFields];
            }
          }
        } else {
          setLinkedChildren([]);
          setSelectedChildIds([]);
        }
      } catch (linkErr) {
        console.error('Không thể tải các biểu mẫu con:', linkErr);
        setLinkedChildren([]);
        setSelectedChildIds([]);
      }

      setFields(finalFields);
      setPreviewTemplateId(data.template.id);
      setPreviewFields(data.fields);
      
      // Khởi tạo form với giá trị mặc định
      const initialForm = {};
      
      // Khởi tạo các trường thường (không thuộc file lặp)
      finalFields.forEach(f => {
        const isRepeatedChildField = f.childTemplateId && currentLinks.find(c => c.id === f.childTemplateId)?.is_repeated === 1;
        if (!isRepeatedChildField) {
          initialForm[f.key_name] = f.field_type === 'boolean' ? false : '';
        }
      });
      
      // Khởi tạo các mảng lặp cho các file lặp (Phương án 3)
      currentLinks.forEach(child => {
        if (child.is_repeated === 1) {
          const firstRecord = { _id: `rec-${Date.now()}-0` };
          finalFields.forEach(f => {
            if (f.childTemplateId === child.id) {
              firstRecord[f.key_name] = f.field_type === 'boolean' ? false : '';
            }
          });
          initialForm[child.id] = [firstRecord];
        }
      });
      
      setFormData(initialForm);

      setCustomerName('');
      setCustomerPhone('');
      setCurrentStep(1);
      setShowPdfModal(false);
      setActiveView('fill');
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleSwitchPreview = async (templateId) => {
    setPreviewTemplateId(templateId);
    try {
      const res = await fetch(`${API_BASE}/templates/${templateId}/form`);
      if (res.ok) {
        const data = await res.json();
        setPreviewFields(data.fields);
      }
    } catch (err) {
      console.error(err);
      showNotification('Không thể tải bản xem trước cho tài liệu này', 'error');
    }
  };

  const getResolvedPreviewData = () => {
    if (!previewTemplateId || !selectedTemplate || previewTemplateId === selectedTemplate.id) {
      return formData;
    }
    const resolved = {};
    const childTemp = linkedChildren.find(c => c.id === previewTemplateId);
    const isRepeated = childTemp?.is_repeated === 1;

    previewFields.forEach(f => {
      if (f.parent_field_key) {
        resolved[f.key_name] = formData[f.parent_field_key] || '';
      } else if (isRepeated) {
        const record = formData[previewTemplateId]?.[0] || {};
        resolved[f.key_name] = record[f.key_name] || '';
      } else {
        resolved[f.key_name] = formData[f.key_name] || '';
      }
    });
    return resolved;
  };

  const getCategoryChildren = (parentId = null) => {
    return categories
      .filter(category => (category.parent_id || null) === parentId)
      .sort((a, b) => {
        if ((a.sort_order || 0) !== (b.sort_order || 0)) {
          return (a.sort_order || 0) - (b.sort_order || 0);
        }
        return a.name.localeCompare(b.name, 'vi');
      });
  };

  const getFlattenedCategoryOptions = (parentId = null, depth = 0) => {
    return getCategoryChildren(parentId).flatMap(category => ([
      {
        id: category.id,
        label: `${'\u00A0\u00A0'.repeat(depth)}${depth > 0 ? '↳ ' : ''}${category.name}`
      },
      ...getFlattenedCategoryOptions(category.id, depth + 1)
    ]));
  };

  const getDescendantCategoryIds = (categoryId) => {
    const childIds = getCategoryChildren(categoryId).map(category => category.id);
    return childIds.reduce(
      (allIds, childId) => [...allIds, childId, ...getDescendantCategoryIds(childId)],
      []
    );
  };

  const toggleCategoryExpanded = (categoryId) => {
    setExpandedCategoryIds(prev => {
      const next = new Set(prev);
      if (next.has(categoryId)) {
        next.delete(categoryId);
      } else {
        next.add(categoryId);
      }
      return next;
    });
  };


  const getFilteredTemplates = () => {
    const query = dashboardSearch.trim().toLowerCase();
    const categoryScope = selectedCategoryId === 'all'
      ? null
      : selectedCategoryId === 'uncategorized'
        ? 'uncategorized'
        : [selectedCategoryId, ...getDescendantCategoryIds(selectedCategoryId)];

    return templates
      .filter(t => {
        const matchesSearch = !query || t.name.toLowerCase().includes(query);
        const matchesStatus = userRole === 'user'
          ? t.status === 'active'
          : (statusFilter === 'all' || t.status === statusFilter);
        const matchesCategory = categoryScope === null
          ? true
          : categoryScope === 'uncategorized'
            ? !t.category_id
            : categoryScope.includes(t.category_id);
        const matchesParent = userRole === 'user'
          ? !t.parent_template_id
          : true;
        return matchesSearch && matchesStatus && matchesCategory && matchesParent;
      })
      .sort((a, b) => {
        if (sortBy === 'name') return a.name.localeCompare(b.name, 'vi');
        if (sortBy === 'oldest') return new Date(a.created_at) - new Date(b.created_at);
        return new Date(b.created_at) - new Date(a.created_at);
      });
  };

  const renderCategoryTree = (parentId = null, depth = 0) => {
    return getCategoryChildren(parentId).map(category => {
      const childCategories = getCategoryChildren(category.id);
      const hasChildren = childCategories.length > 0;
      const isExpanded = expandedCategoryIds.has(category.id);
      const descendantIds = [category.id, ...getDescendantCategoryIds(category.id)];
      const templateCount = templates.filter(t => descendantIds.includes(t.category_id)).length;
      const isActive = selectedCategoryId === category.id;

      return (
        <div key={category.id} className="space-y-1">
          <div
            className={`flex items-center gap-2 px-2 py-1.5 rounded-xl border transition-all ${
              isActive
                ? 'bg-emerald-50 border-emerald-300'
                : 'bg-white border-slate-200 hover:border-emerald-200'
            }`}
            style={{ marginLeft: `${depth * 14}px` }}
          >
            <button
              type="button"
              onClick={() => hasChildren && toggleCategoryExpanded(category.id)}
              className={`w-6 h-6 rounded-md flex items-center justify-center text-xs font-bold transition-colors ${
                hasChildren
                  ? 'text-slate-500 hover:bg-slate-100 cursor-pointer'
                  : 'text-slate-300 cursor-default'
              }`}
              disabled={!hasChildren}
              title={hasChildren ? (isExpanded ? 'Thu gọn nhánh' : 'Mở rộng nhánh') : 'Không có nhánh con'}
            >
              {hasChildren ? (isExpanded ? '▾' : '▸') : '•'}
            </button>
            <button
              type="button"
              onClick={() => {
                setSelectedCategoryId(category.id);
                if (hasChildren && !isExpanded) {
                  setExpandedCategoryIds(prev => new Set(prev).add(category.id));
                }
              }}
              className="flex-1 text-left min-w-0 cursor-pointer"
            >
              <div className="flex items-center justify-between gap-2">
                <span className={`text-xs font-semibold truncate ${isActive ? 'text-emerald-800' : 'text-slate-700'}`}>{category.name}</span>
                <span className="text-[10px] font-bold text-slate-400">{templateCount}</span>
              </div>
            </button>
          </div>
          {hasChildren && isExpanded ? renderCategoryTree(category.id, depth + 1) : null}
        </div>
      );
    });
  };

  const handleFormSubmit = async (e) => {
    e.preventDefault();
    setIsSubmitting(true);
    setSubmissionProgress({ current: 0, total: 0, message: 'Chuẩn bị dữ liệu...' });

    try {
      // Tính toán số bước progress: 1 master + số child files + 2 (validate + finalize)
      const childCount = selectedChildIds.length;
      const totalSteps = 2 + childCount + 2;

      setSubmissionProgress({ current: 1, total: totalSteps, message: 'Chuẩn bị dữ liệu...' });
      await new Promise(r => setTimeout(r, 300));

      if (childCount > 0) {
        setSubmissionProgress({ current: 2, total: totalSteps, message: 'Xử lý file gốc...' });
        await new Promise(r => setTimeout(r, 200));
      }

      const res = await fetch(`${API_BASE}/submissions`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          templateId: selectedTemplate.id,
          customerName,
          customerPhone,
          values: formData,
          selectedChildIds
        }),
      });

      if (!res.ok) {
        let errMsg = 'Lỗi gửi hồ sơ';
        try {
          const data = await res.json();
          errMsg = data.error || errMsg;
        } catch (parseErr) {
          console.warn('Lỗi parse JSON:', parseErr);
        }
        throw new Error(errMsg);
      }
      const data = await res.json();

      // Simulate child file processing
      if (childCount > 0) {
        for (let i = 0; i < childCount; i++) {
          setSubmissionProgress({
            current: 3 + i,
            total: totalSteps,
            message: `Xử lý file con ${i + 1}/${childCount}...`
          });
          await new Promise(r => setTimeout(r, 250));
        }
      }

      setSubmissionProgress({ current: totalSteps - 1, total: totalSteps, message: 'Đóng gói file...' });
      await new Promise(r => setTimeout(r, 200));

      // Initialize the list of files generated
      const generatedFiles = [`${selectedTemplate.name}.docx`];
      const activeChildren = linkedChildren.filter(child => selectedChildIds.includes(child.id));
      activeChildren.forEach(child => {
        if (child.is_repeated) {
          const records = formData[child.id] || [{}];
          records.forEach((_, rIdx) => {
            generatedFiles.push(`${child.name}_Căn_${rIdx + 1}.docx`);
          });
        } else {
          generatedFiles.push(`${child.name}.docx`);
        }
      });

      setSuccessSelectedFiles(generatedFiles);
      setSuccessPreviewFilename(`${selectedTemplate.name}.docx`);

      setSubmissionResult({
        submissionId: data.submissionId,
        downloadUrl: `http://localhost:5000${data.downloadUrl}`,
        customerName,
        customerPhone,
        templateName: selectedTemplate.name,
        files: generatedFiles
      });
      setActiveView('success');
      showNotification('Đã tạo thành công hồ sơ và xuất tài liệu!');
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsSubmitting(false);
      setSubmissionProgress({ current: 0, total: 0, message: '' });
    }
  };

  const handleDownloadSelectedFiles = () => {
    if (!submissionResult || successSelectedFiles.length === 0) {
      showNotification('Vui lòng chọn ít nhất một tài liệu để tải.', 'error');
      return;
    }
    successSelectedFiles.forEach((filename, index) => {
      setTimeout(() => {
        const link = document.createElement('a');
        link.href = `http://localhost:5000/api/submissions/${submissionResult.submissionId}/download-file?filename=${encodeURIComponent(filename)}`;
        link.setAttribute('download', filename);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
      }, index * 250);
    });
    showNotification(`Đang tiến hành tải ${successSelectedFiles.length} tài liệu...`);
  };

  // Reset preview tab if selected child template is deselected
  useEffect(() => {
    if (selectedTemplate && previewTemplateId && previewTemplateId !== selectedTemplate.id) {
      if (!selectedChildIds.includes(previewTemplateId)) {
        handleSwitchPreview(selectedTemplate.id);
      }
    }
  }, [selectedChildIds, selectedTemplate, previewTemplateId]);

  const hasActiveFields = (stepNum) => {
    if (stepNum === 1) return true;
    return fields.some(f => 
      getFieldStep(f) === stepNum && 
      (!f.childTemplateId || selectedChildIds.includes(f.childTemplateId))
    );
  };

  return (
    <div className="app-layout">
      {/* Toast Notification */}
      {notification && (
        <div className={`fixed top-5 right-5 z-50 px-6 py-3 rounded-lg shadow-xl font-medium transition-all transform duration-300 translate-y-0 ${
          notification.type === 'error' 
            ? 'bg-rose-500/90 text-white border border-rose-400' 
            : 'bg-emerald-500/90 text-white border border-emerald-400'
        }`}>
          <div className="flex items-center gap-2">
            <span>{notification.type === 'error' ? '⚠️' : '✅'}</span>
            <span>{notification.message}</span>
          </div>
        </div>
      )}

      {/* Left Sidebar */}
      <aside className={`sidebar ${sidebarCollapsed ? 'sidebar-collapsed' : ''} select-none`}>
        {/* Sidebar Brand/Logo */}
        <div className="h-16 flex items-center px-6 border-b border-slate-800 gap-3 overflow-hidden">
          <div className="bg-emerald-500 text-slate-900 p-2 rounded-xl shrink-0">
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2.5">
              <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
          </div>
          {!sidebarCollapsed && (
            <div className="min-w-0 transition-opacity duration-200">
              <h1 className="text-sm font-extrabold tracking-tight text-white font-display">
                Công Chứng <span className="text-emerald-400 font-extrabold">Số</span>
              </h1>
              <span className="text-[8px] text-slate-500 font-bold uppercase tracking-wider block">Notary Digital Hub</span>
            </div>
          )}
        </div>

        {/* Sidebar Menu Items */}
        <div className="flex-1 py-4 overflow-y-auto space-y-1">
          {userRole === 'user' ? (
            <>
              <button
                type="button"
                onClick={() => {
                  setSidebarActiveMenu('templates');
                  setActiveView('dashboard');
                }}
                className={`sidebar-menu-item w-[calc(100%-24px)] text-left ${
                  activeView === 'dashboard' && sidebarActiveMenu === 'templates' ? 'active' : ''
                }`}
              >
                <svg className="w-4 h-4 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M4 6h16M4 12h16M4 18h7" />
                </svg>
                {!sidebarCollapsed && <span>Duyệt biểu mẫu</span>}
              </button>
              <button
                type="button"
                onClick={() => {
                  setSidebarActiveMenu('guide');
                  setActiveView('dashboard');
                }}
                className={`sidebar-menu-item w-[calc(100%-24px)] text-left ${
                  activeView === 'dashboard' && sidebarActiveMenu === 'guide' ? 'active' : ''
                }`}
              >
                <svg className="w-4 h-4 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M12 6.253v13m0-13C10.832 5.477 9.246 5 7.5 5S4.168 5.477 3 6.253v13C4.168 18.477 5.754 18 7.5 18s3.332.477 4.5 1.253m0-13C13.168 5.477 14.754 5 16.5 5c1.747 0 3.332.477 4.5 1.253v13C19.832 18.477 18.247 18 16.5 18c-1.746 0-3.332.477-4.5 1.253" />
                </svg>
                {!sidebarCollapsed && <span>Hướng dẫn dịch vụ</span>}
              </button>
            </>
          ) : (
            <>
              <button
                type="button"
                onClick={() => {
                  setSidebarActiveMenu('templates');
                  setActiveView('dashboard');
                }}
                className={`sidebar-menu-item w-[calc(100%-24px)] text-left ${
                  activeView === 'dashboard' && sidebarActiveMenu === 'templates' ? 'active' : ''
                }`}
              >
                <svg className="w-4 h-4 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M4 6h16M4 12h16M4 18h7" />
                </svg>
                {!sidebarCollapsed && <span>Thư viện biểu mẫu</span>}
              </button>
              <button
                type="button"
                onClick={() => {
                  setSidebarActiveMenu('import');
                  setActiveView('dashboard');
                }}
                className={`sidebar-menu-item w-[calc(100%-24px)] text-left ${
                  activeView === 'dashboard' && sidebarActiveMenu === 'import' ? 'active' : ''
                }`}
              >
                <svg className="w-4 h-4 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M12 16V4m0 0l-4 4m4-4l4 4M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2" />
                </svg>
                {!sidebarCollapsed && <span>Nhập biểu mẫu</span>}
              </button>
              <button
                type="button"
                onClick={() => {
                  setSidebarActiveMenu('categories');
                  setActiveView('dashboard');
                }}
                className={`sidebar-menu-item w-[calc(100%-24px)] text-left ${
                  activeView === 'dashboard' && sidebarActiveMenu === 'categories' ? 'active' : ''
                }`}
              >
                <svg className="w-4 h-4 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M3 7h5l2 2h11v8a2 2 0 01-2 2H5a2 2 0 01-2-2V7z" />
                </svg>
                {!sidebarCollapsed && <span>Quản lý danh mục</span>}
              </button>
              <button
                type="button"
                onClick={() => {
                  setSidebarActiveMenu('submissions');
                  setActiveView('dashboard');
                }}
                className={`sidebar-menu-item w-[calc(100%-24px)] text-left ${
                  activeView === 'dashboard' && sidebarActiveMenu === 'submissions' ? 'active' : ''
                }`}
              >
                <svg className="w-4 h-4 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01" />
                </svg>
                {!sidebarCollapsed && <span>Hồ sơ đã nhận</span>}
              </button>
              <button
                type="button"
                onClick={() => {
                  setSidebarActiveMenu('stats');
                  setActiveView('dashboard');
                }}
                className={`sidebar-menu-item w-[calc(100%-24px)] text-left ${
                  activeView === 'dashboard' && sidebarActiveMenu === 'stats' ? 'active' : ''
                }`}
              >
                <svg className="w-4 h-4 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
                </svg>
                {!sidebarCollapsed && <span>Thống kê báo cáo</span>}
              </button>
            </>
          )}
        </div>

        {/* Sidebar Collapse Toggle Button */}
        <div className="p-4 border-t border-slate-800 flex justify-center">
          <button
            type="button"
            onClick={() => setSidebarCollapsed(!sidebarCollapsed)}
            className="w-10 h-10 rounded-xl bg-slate-800/50 hover:bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center transition-colors cursor-pointer"
          >
            {sidebarCollapsed ? '➔' : '➔'.split('').reverse().join('')}
          </button>
        </div>
      </aside>

      {/* Main Container */}
      <div className="main-container">
        {/* Top Navbar */}
        <nav className="navbar select-none">
          {/* Left: View State Header / Breadcrumb */}
          <div className="flex items-center gap-3">
            {activeView !== 'dashboard' && (
              <button
                type="button"
                onClick={() => setActiveView('dashboard')}
                className="p-2 hover:bg-slate-100 rounded-xl transition-colors cursor-pointer text-slate-600 shrink-0"
              >
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M10 19l-7-7m0 0l7-7m-7 7h18" />
                </svg>
              </button>
            )}
            <div>
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest font-sans">
                {userRole === 'user' ? 'Khách Hàng' : 'Văn Phòng Công Chứng'}
              </span>
              <h2 className="text-sm font-extrabold text-slate-800 mt-0.5 tracking-tight font-display">
                {activeView === 'config' ? 'Chi tiết biểu mẫu' : activeView === 'fill' ? 'Điền hồ sơ giao dịch' : activeView === 'success' ? 'Biên nhận kết quả' : sidebarActiveMenu === 'templates' ? 'Thư viện biểu mẫu' : sidebarActiveMenu === 'import' ? 'Nhập biểu mẫu' : sidebarActiveMenu === 'categories' ? 'Quản lý danh mục biểu mẫu' : sidebarActiveMenu === 'guide' ? 'Hướng dẫn nộp hồ sơ' : sidebarActiveMenu === 'submissions' ? 'Danh sách hồ sơ nhận được' : 'Thống kê tổng quan'}
              </h2>
            </div>
          </div>

          {/* Middle: Role Switcher Segmented Tabs */}
          <div className="role-switcher">
            <button
              type="button"
              onClick={() => {
                setUserRole('user');
                setSidebarActiveMenu('templates');
                setActiveView('dashboard');
              }}
              className={`role-switcher-btn ${userRole === 'user' ? 'active' : ''}`}
            >
              🙋 Khách hàng
            </button>
            <button
              type="button"
              onClick={() => {
                setUserRole('congchung');
                setSidebarActiveMenu('templates');
                setActiveView('dashboard');
              }}
              className={`role-switcher-btn ${userRole === 'congchung' ? 'active' : ''}`}
            >
              💼 Văn phòng
            </button>
          </div>

          {/* Right: Actions */}
          <div className="flex items-center gap-2">
            {activeView === 'fill' && (
              <>
                <button
                  type="button"
                  onClick={handleResetForm}
                  className="px-3.5 py-2 text-xs font-semibold text-slate-700 hover:text-slate-900 bg-slate-100 hover:bg-slate-200 transition-colors rounded-xl border border-slate-200 shadow-sm cursor-pointer"
                >
                  ↺ Nhập lại
                </button>
                <button
                  type="button"
                  onClick={() => setShowPdfModal(true)}
                  className="px-4 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 transition-colors rounded-xl border border-emerald-500 shadow-sm hover:shadow-md cursor-pointer"
                >
                  👁 Xem trước
                </button>
              </>
            )}
            {userRole === 'congchung' && (
              <div className="hidden md:flex items-center gap-2 pl-3 border-l border-slate-200">
                <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse shrink-0"></span>
                <span className="text-xs text-slate-600 font-bold">Notary Admin</span>
              </div>
            )}
          </div>
        </nav>

        {/* Content Area */}
        <main className="content-area">
          {activeView === 'dashboard' && sidebarActiveMenu === 'templates' && (
            <div className="space-y-8 animate-fade-up">
              {/* Stats Bar (Only for Admin/congchung) */}
              {userRole === 'congchung' && (
                <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
                  <div className="glass-panel glass-panel-hover p-5 flex items-center gap-4 bg-white">
                    <div className="p-3 bg-slate-100 text-slate-700 border border-slate-200 rounded-2xl">
                      <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                        <path strokeLinecap="round" strokeLinejoin="round" d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10" />
                      </svg>
                    </div>
                    <div>
                      <div className="text-2xl font-extrabold text-slate-900 font-display tracking-tight">{templates.length}</div>
                      <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Tổng số biểu mẫu</div>
                    </div>
                  </div>
                  <div className="glass-panel glass-panel-hover p-5 flex items-center gap-4 bg-white">
                    <div className="p-3 bg-emerald-50 text-emerald-700 border border-emerald-100 rounded-2xl">
                      <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                        <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                      </svg>
                    </div>
                    <div>
                      <div className="text-2xl font-extrabold text-slate-900 font-display tracking-tight">
                        {templates.filter(t => t.status === 'active').length}
                      </div>
                      <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Biểu mẫu hoạt động</div>
                    </div>
                  </div>
                  <div className="glass-panel glass-panel-hover p-5 flex items-center gap-4 bg-white">
                    <div className="p-3 bg-emerald-50 text-emerald-700 border border-emerald-100 rounded-2xl">
                      <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                        <path strokeLinecap="round" strokeLinejoin="round" d="M13 10V3L4 14h7v7l9-11h-7z" />
                      </svg>
                    </div>
                    <div>
                      <div className="text-2xl font-extrabold text-emerald-600 font-display tracking-tight">Sẵn sàng</div>
                      <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Trạng thái hệ thống</div>
                    </div>
                  </div>
                </div>
              )}

              {/* Template Library Layout */}
              <div className="space-y-4">
                {/* Upload Area (Only for Admin/congchung) */}
                {false && userRole === 'congchung' && (
                  <div className="glass-panel p-6 bg-white">
                    <h2 className="text-sm font-bold mb-2 text-slate-800 flex items-center gap-2 tracking-tight font-display">
                      <svg className="w-4 h-4 text-emerald-500 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                        <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
                      </svg>
                      Tải lên biểu mẫu mới
                    </h2>
                    <p className="text-xs text-slate-400 mb-6 leading-relaxed">
                      Tải lên biểu mẫu Word (`.doc` hoặc `.docx`) chứa các thẻ biến dạng `{"{{ten_bien}}"}` để tự động dựng cấu hình form trực tuyến.
                    </p>

                    <form onSubmit={handleFileUpload} className="space-y-4">
                      <div>
                        <label className="clean-label">Tên biểu mẫu (khi tải 1 file)</label>
                        <input
                          type="text"
                          value={uploadName}
                          onChange={(e) => setUploadName(e.target.value)}
                          placeholder="Để trống nếu tải nhiều file"
                          className="clean-input w-full text-xs"
                        />
                      </div>

                      <div>
                        <label className="clean-label">Danh mục biểu mẫu</label>
                        <select
                          value={uploadCategoryId}
                          onChange={(e) => setUploadCategoryId(e.target.value)}
                          className="clean-input w-full text-xs cursor-pointer"
                        >
                          <option value="">Chưa phân loại</option>
                          {getFlattenedCategoryOptions().map(category => (
                            <option key={category.id} value={category.id}>
                              {category.label}
                            </option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <label className="clean-label">Chọn file tài liệu (.doc, .docx)</label>
                        <div
                          className={`border-2 border-dashed border-slate-200 hover:border-emerald-400 rounded-2xl p-6 text-center cursor-pointer transition-all relative ${
                            isDragActive ? 'drag-active scale-[1.005]' : 'bg-slate-50/50'
                          }`}
                          onDragOver={(e) => {
                            e.preventDefault();
                            setIsDragActive(true);
                          }}
                          onDragLeave={(e) => {
                            e.preventDefault();
                            setIsDragActive(false);
                          }}
                          onDrop={(e) => {
                            e.preventDefault();
                            setIsDragActive(false);
                            if (e.dataTransfer.files?.length) {
                              handleSelectUploadFiles(e.dataTransfer.files);
                            }
                          }}
                        >
                          <input
                            type="file"
                            multiple
                            accept=".doc,.docx"
                            onChange={(e) => {
                              if (e.target.files?.length) {
                                handleSelectUploadFiles(e.target.files);
                              }
                            }}
                            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                          />
                          <div className="space-y-2">
                            <svg className="w-9 h-9 text-slate-300 mx-auto shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="1.5">
                              <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m3.75 9v6m3-3H9m1.5-12H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
                            </svg>
                            <div className="text-xs text-slate-700 font-semibold">
                              {uploadFiles.length > 0 ? `Đã chọn ${uploadFiles.length} file` : 'Kéo thả file hoặc click để chọn'}
                            </div>
                            <div className="text-[10px] text-slate-400">Hỗ trợ tối đa 10 file .doc hoặc .docx, mỗi file không quá 1MB</div>
                          </div>
                        </div>
                      </div>

                      {uploadFiles.length > 0 && (
                        <div className="rounded-2xl border border-slate-200 bg-slate-50/70 p-4">
                          <div className="flex items-center justify-between gap-3 mb-3">
                            <p className="text-[11px] font-bold uppercase tracking-wider text-slate-500">Danh sách file sẽ tải lên</p>
                            <button
                              type="button"
                              onClick={() => {
                                setUploadFiles([]);
                                setUploadName('');
                              }}
                              className="text-[11px] font-semibold text-slate-500 hover:text-slate-700 cursor-pointer"
                            >
                              Xóa danh sách
                            </button>
                          </div>
                          <div className="space-y-2 max-h-40 overflow-y-auto">
                            {uploadFiles.map(file => (
                              <div key={`${file.name}-${file.size}-${file.lastModified}`} className="flex items-center justify-between gap-3 rounded-xl bg-white px-3 py-2 border border-slate-200 text-xs">
                                <span className="truncate text-slate-700 font-medium">{file.name}</span>
                                <span className="shrink-0 text-slate-400">{(file.size / 1024).toFixed(0)} KB</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}

                      <button
                        type="submit"
                        disabled={isUploading}
                        className="btn-premium w-full bg-emerald-600 hover:bg-emerald-700 text-white font-semibold py-2.5 text-xs transition-all shadow-sm hover:shadow-md flex items-center justify-center gap-2 disabled:opacity-50 cursor-pointer"
                      >
                        {isUploading ? (
                          <>
                            <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin"></div>
                            Đang xử lý đợt upload...
                          </>
                        ) : (
                          'Tải lên & Trích xuất'
                        )}
                      </button>
                    </form>

                    {uploadResults.length > 0 && (
                      <div className="mt-4 rounded-2xl border border-slate-200 bg-slate-50/60 p-4 space-y-3">
                        <div className="flex items-center justify-between gap-3">
                          <h3 className="text-xs font-bold text-slate-800 tracking-tight">Kết quả đợt upload</h3>
                          <span className="text-[11px] text-slate-500">{uploadResults.filter(item => item.status === 'success').length}/{uploadResults.length} thành công</span>
                        </div>
                        <div className="space-y-2">
                          {uploadResults.map((result) => (
                            <div key={`${result.fileName}-${result.templateId || result.error}`} className="rounded-xl border border-slate-200 bg-white p-3 text-xs">
                              <div className="flex items-start justify-between gap-3">
                                <div className="min-w-0">
                                  <div className="font-semibold text-slate-800 truncate">{result.fileName}</div>
                                  <div className={`mt-1 ${result.status === 'success' ? 'text-emerald-600' : 'text-rose-600'}`}>
                                    {result.status === 'success'
                                      ? `${result.variablesCount} biến, tên biểu mẫu: ${result.name}`
                                      : result.error}
                                  </div>
                                </div>
                                {result.status === 'success' && (
                                  <button
                                    type="button"
                                    onClick={() => handleOpenConfig(result.templateId)}
                                    className="shrink-0 text-[11px] font-semibold text-emerald-700 hover:text-emerald-800 cursor-pointer"
                                  >
                                    Mở chi tiết
                                  </button>
                                )}
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}

                    <div className="mt-6 pt-5 border-t border-slate-200 space-y-3">
                      <div>
                        <h3 className="text-xs font-bold text-slate-800 tracking-tight">Tạo danh mục mới</h3>
                        <p className="text-[10px] text-slate-400 mt-1">Dùng danh mục để tổ chức biểu mẫu theo cây thư mục ảo.</p>
                      </div>
                      <form onSubmit={handleCreateCategory} className="space-y-3">
                        <select
                          value={newCategoryParentId}
                          onChange={(e) => setNewCategoryParentId(e.target.value)}
                          className="clean-input w-full text-xs cursor-pointer"
                        >
                          <option value="">Danh mục gốc</option>
                          {getFlattenedCategoryOptions().map(category => (
                            <option key={category.id} value={category.id}>
                              {category.label}
                            </option>
                          ))}
                        </select>
                        <input
                          type="text"
                          value={newCategoryName}
                          onChange={(e) => setNewCategoryName(e.target.value)}
                          placeholder="Ví dụ: Chuyển nhượng đất"
                          className="clean-input w-full text-xs"
                        />
                        <button
                          type="submit"
                          className="btn-premium w-full bg-white hover:bg-slate-50 text-slate-700 font-semibold py-2.5 text-xs border border-slate-200 transition-all cursor-pointer"
                        >
                          Thêm danh mục
                        </button>
                      </form>
                    </div>
                  </div>
                )}

                {/* Template List */}
                <div className="space-y-4">
                  <div className="flex items-end justify-between gap-3 flex-wrap">
                    <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2 tracking-tight font-display">
                      <svg className="w-4 h-4 text-emerald-500 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                        <path strokeLinecap="round" strokeLinejoin="round" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01" />
                      </svg>
                      Danh sách biểu mẫu hiện có
                      <span className="text-[10px] font-bold text-slate-400 bg-slate-100 border border-slate-200 px-1.5 py-0.5 rounded-md ml-1">
                        {getFilteredTemplates().length}/{templates.length}
                      </span>
                    </h2>
                    {userRole === 'congchung' && (
                      <button
                        type="button"
                        onClick={() => {
                          setSidebarActiveMenu('import');
                          setActiveView('dashboard');
                        }}
                        className="btn-premium bg-emerald-600 hover:bg-emerald-700 text-white font-semibold py-2.5 px-4 text-xs transition-all shadow-sm hover:shadow-md cursor-pointer"
                      >
                        Nhập biểu mẫu
                      </button>
                    )}
                  </div>

                  {/* Search & Filter Toolbar */}
                  {templates.length > 0 && (
                    <div className="flex flex-col sm:flex-row gap-2.5 sm:items-center">
                      <div className="relative flex-1">
                        <svg className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                          <path strokeLinecap="round" strokeLinejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
                        </svg>
                        <input
                          type="text"
                          value={dashboardSearch}
                          onChange={(e) => setDashboardSearch(e.target.value)}
                          placeholder="Tìm theo tên biểu mẫu..."
                          className="w-full pl-9 pr-3 py-2.5 text-xs bg-white border border-slate-200 rounded-xl focus:outline-none focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 transition-all"
                        />
                        {dashboardSearch && (
                          <button
                            type="button"
                            onClick={() => setDashboardSearch('')}
                            className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-700 text-base font-bold w-5 h-5 rounded-full flex items-center justify-center hover:bg-slate-100 cursor-pointer"
                          >
                            ×
                          </button>
                        )}
                      </div>
                      <select
                        value={statusFilter}
                        onChange={(e) => setStatusFilter(e.target.value)}
                        className="px-3 py-2.5 text-xs bg-white border border-slate-200 rounded-xl focus:outline-none focus:border-emerald-500 cursor-pointer text-slate-700 font-semibold"
                      >
                        <option value="all">Tất cả trạng thái</option>
                        <option value="active">Đang hoạt động</option>
                        <option value="draft">Bản nháp</option>
                      </select>
                      <select
                        value={sortBy}
                        onChange={(e) => setSortBy(e.target.value)}
                        className="px-3 py-2.5 text-xs bg-white border border-slate-200 rounded-xl focus:outline-none focus:border-emerald-500 cursor-pointer text-slate-700 font-semibold"
                      >
                        <option value="newest">Mới nhất</option>
                        <option value="oldest">Cũ nhất</option>
                        <option value="name">Tên A → Z</option>
                      </select>
                    </div>
                  )}

                  {isLoading && templates.length === 0 ? (
                    <div className="glass-panel p-12 text-center text-slate-400 bg-white">
                      <div className="w-8 h-8 border-2 border-emerald-500/30 border-t-emerald-500 rounded-full animate-spin mx-auto mb-4"></div>
                      Đang đồng bộ biểu mẫu từ máy chủ...
                    </div>
                  ) : templates.length === 0 ? (
                    <div className="glass-panel p-12 text-center text-slate-400 bg-white">
                      <div className="text-3xl mb-3">📁</div>
                      Chưa có biểu mẫu nào được cấu hình.
                    </div>
                  ) : (
                    <div className="grid grid-cols-1 xl:grid-cols-[280px_minmax(0,1fr)] gap-4 items-start">
                      <div className="glass-panel p-4 bg-white space-y-3 sticky top-4">
                        <div className="flex items-center justify-between gap-2">
                          <h3 className="text-xs font-bold text-slate-800 tracking-tight font-display">Cây danh mục</h3>
                          <span className="text-[10px] font-bold text-slate-400">{categories.length} mục</span>
                        </div>
                        <div className="space-y-2">
                          <button
                            type="button"
                            onClick={() => setSelectedCategoryId('all')}
                            className={`w-full text-left px-3 py-2 rounded-xl border transition-all cursor-pointer ${
                              selectedCategoryId === 'all'
                                ? 'bg-slate-900 border-slate-900 text-white'
                                : 'bg-white border-slate-200 text-slate-700 hover:border-emerald-200 hover:bg-emerald-50/40'
                            }`}
                          >
                            <div className="flex items-center justify-between gap-2">
                              <span className="text-xs font-semibold">Tất cả danh mục</span>
                              <span className={`text-[10px] font-bold ${selectedCategoryId === 'all' ? 'text-white/80' : 'text-slate-400'}`}>{templates.length}</span>
                            </div>
                          </button>
                          <button
                            type="button"
                            onClick={() => setSelectedCategoryId('uncategorized')}
                            className={`w-full text-left px-3 py-2 rounded-xl border transition-all cursor-pointer ${
                              selectedCategoryId === 'uncategorized'
                                ? 'bg-amber-50 border-amber-300 text-amber-800'
                                : 'bg-white border-slate-200 text-slate-700 hover:border-amber-200 hover:bg-amber-50/40'
                            }`}
                          >
                            <div className="flex items-center justify-between gap-2">
                              <span className="text-xs font-semibold">Chưa phân loại</span>
                              <span className="text-[10px] font-bold text-slate-400">{templates.filter(t => !t.category_id).length}</span>
                            </div>
                          </button>
                        </div>
                        <div className="space-y-1 max-h-[520px] overflow-y-auto pr-1">
                          {renderCategoryTree()}
                        </div>
                      </div>

                      {getFilteredTemplates().length === 0 ? (
                        <div className="glass-panel p-12 text-center text-slate-400 bg-white">
                          <div className="text-3xl mb-3">🔍</div>
                          Không tìm thấy biểu mẫu nào khớp với bộ lọc hiện tại.
                        </div>
                      ) : (
                        <div className={`grid grid-cols-1 ${userRole === 'congchung' ? 'md:grid-cols-2' : 'md:grid-cols-3'} gap-4`}>
                          {getFilteredTemplates().map((temp) => (
                        <div key={temp.id} className="glass-panel glass-panel-hover p-5 flex flex-col justify-between bg-white">
                          <div>
                            <div className="flex justify-between items-start gap-2 mb-3">
                              {userRole === 'congchung' && editingTemplateId === temp.id ? (
                                <div className="flex items-center gap-1.5 flex-1 min-w-0">
                                  <input
                                    type="text"
                                    value={editingNameValue}
                                    onChange={(e) => setEditingNameValue(e.target.value)}
                                    className="flex-1 bg-white border border-zinc-300 rounded-lg px-2.5 py-1 text-xs text-zinc-900 focus:outline-none focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20"
                                    autoFocus
                                    onKeyDown={(e) => {
                                      if (e.key === 'Enter') handleSaveTemplateName(temp.id);
                                      else if (e.key === 'Escape') setEditingTemplateId(null);
                                    }}
                                  />
                                  <button
                                    onClick={() => handleSaveTemplateName(temp.id)}
                                    className="text-emerald-600 hover:text-emerald-800 text-xs p-1 cursor-pointer"
                                    title="Lưu"
                                  >
                                    ✓
                                  </button>
                                  <button
                                    onClick={() => setEditingTemplateId(null)}
                                    className="text-rose-600 hover:text-rose-800 text-xs p-1 cursor-pointer"
                                    title="Hủy"
                                  >
                                    ✕
                                  </button>
                                </div>
                              ) : (
                                <div className="flex items-center gap-1.5 flex-1 min-w-0">
                                  <h3 className="font-bold text-zinc-800 truncate text-xs md:text-sm tracking-wide flex-1 font-display">{temp.name}</h3>
                                  {userRole === 'congchung' && (
                                    <>
                                      <button
                                        onClick={() => {
                                          setEditingTemplateId(temp.id);
                                          setEditingNameValue(temp.name);
                                        }}
                                        className="text-zinc-400 hover:text-zinc-600 text-xs p-1 shrink-0 cursor-pointer"
                                        title="Đổi tên"
                                      >
                                        <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                                          <path strokeLinecap="round" strokeLinejoin="round" d="M15.232 5.232l3.536 3.536m-2.036-5.036a2.5 2.5 0 113.536 3.536L6.5 21.036H3v-3.572L16.732 3.732z" />
                                        </svg>
                                      </button>
                                      <button
                                        onClick={() => handleDuplicateTemplate(temp.id)}
                                        className="text-zinc-400 hover:text-emerald-600 text-xs p-1 shrink-0 cursor-pointer"
                                        title="Sao chép biểu mẫu"
                                      >
                                        <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                                          <rect x="9" y="9" width="13" height="13" rx="2" ry="2" />
                                          <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" strokeLinecap="round" strokeLinejoin="round" />
                                        </svg>
                                      </button>
                                      <button
                                        onClick={() => handleDeleteTemplate(temp.id, temp.name)}
                                        className="text-zinc-400 hover:text-rose-600 text-xs p-1 shrink-0 cursor-pointer"
                                        title="Xóa biểu mẫu"
                                      >
                                        <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                                          <path strokeLinecap="round" strokeLinejoin="round" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                                        </svg>
                                      </button>
                                    </>
                                  )}
                                </div>
                              )}
                              <div className="flex flex-col gap-1 items-end shrink-0">
                                <span className={`px-2 py-0.5 rounded-md text-[8px] font-bold uppercase tracking-wider ${
                                  temp.status === 'active' 
                                    ? 'bg-emerald-50 text-emerald-700 border border-emerald-200' 
                                    : 'bg-amber-50 text-amber-700 border border-amber-200'
                                }`}>
                                  {temp.status === 'active' ? 'Hoạt động' : 'Bản nháp'}
                                </span>
                                {temp.parent_template_id && (
                                  <span className="px-2 py-0.5 rounded-md text-[8px] font-bold uppercase tracking-wider bg-zinc-100 text-zinc-700 border border-zinc-200">
                                    File Con
                                  </span>
                                )}
                                {temp.children_count > 0 && (
                                  <span className="px-2 py-0.5 rounded-md text-[8px] font-bold uppercase tracking-wider bg-emerald-50 text-emerald-800 border border-emerald-200">
                                    File Gốc
                                  </span>
                                )}
                              </div>
                            </div>
                            <div className="text-[10px] text-zinc-400 space-y-1 mb-4 font-sans">
                              <div className="flex justify-between">
                                <span>Ngày tạo:</span>
                                <span className="text-zinc-600 font-semibold">{temp.created_at.substring(0, 10)}</span>
                              </div>
                              <div className="flex justify-between">
                                <span>Danh mục:</span>
                                <span className="text-zinc-700 font-semibold truncate max-w-[140px]" title={temp.category_name || 'Chưa phân loại'}>
                                  {temp.category_name || 'Chưa phân loại'}
                                </span>
                              </div>
                              <div className="flex justify-between">
                                <span>Biến động:</span>
                                <span className="text-zinc-800 font-bold">{temp.fields_count} biến</span>
                              </div>
                              {temp.parent_template_id && (
                                <div className="flex justify-between text-zinc-600 font-semibold">
                                  <span>Thuộc file gốc:</span>
                                  <span className="truncate max-w-[120px]" title={templates.find(t => t.id === temp.parent_template_id)?.name || 'File Gốc'}>
                                    {templates.find(t => t.id === temp.parent_template_id)?.name || 'File Gốc'}
                                  </span>
                                </div>
                              )}
                              {temp.children_count > 0 && (
                                <div className="flex justify-between text-emerald-700 font-semibold">
                                  <span>File liên kết:</span>
                                  <span>{temp.children_count} file con</span>
                                </div>
                              )}
                            </div>
                          </div>

                          <div>
                            <div className={`grid ${userRole === 'congchung' ? 'grid-cols-2' : 'grid-cols-1'} gap-2 pt-3 border-t border-zinc-100`}>
                              {userRole === 'congchung' && (
                                <button
                                  onClick={() => handleOpenConfig(temp.id)}
                                  className="btn-premium text-[10px] bg-white hover:bg-zinc-50 text-zinc-700 font-semibold py-2 px-3 rounded-lg border border-zinc-250 hover:border-zinc-350 transition-colors flex items-center justify-center gap-1.5 cursor-pointer"
                                >
                                  Chi tiết
                                </button>
                              )}
                              <button
                                onClick={() => handleOpenFill(temp.id)}
                                className="btn-premium text-[10px] bg-emerald-600 hover:bg-emerald-700 text-white font-semibold py-2 px-3 rounded-lg border border-emerald-500 shadow-sm transition-colors flex items-center justify-center gap-1.5 cursor-pointer"
                              >
                                Điền hồ sơ
                              </button>
                            </div>
                          </div>
                        </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}

          {activeView === 'dashboard' && sidebarActiveMenu === 'import' && userRole === 'congchung' && (
            <div className="grid grid-cols-1 xl:grid-cols-[minmax(0,1.2fr)_360px] gap-6 animate-fade-up">
              <div className="glass-panel p-6 bg-white space-y-5">
                <div>
                  <span className="text-[11px] text-emerald-600 font-bold uppercase tracking-wider font-display">Workspace nhập liệu</span>
                  <h2 className="text-base font-bold text-slate-800 mt-1 font-display">Nhập biểu mẫu mới</h2>
                  <p className="text-xs text-slate-400 mt-2 leading-relaxed">
                    Khu vực này chỉ dành cho việc tạo biểu mẫu mới. Sau khi tải lên xong, hệ thống sẽ chuyển sang trang chi tiết để cấu hình trường dữ liệu.
                  </p>
                </div>

                <form onSubmit={handleFileUpload} className="space-y-4">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                      <label className="clean-label">Tên biểu mẫu (khi tải 1 file)</label>
                      <input
                        type="text"
                        value={uploadName}
                        onChange={(e) => setUploadName(e.target.value)}
                        placeholder="Để trống nếu tải nhiều file"
                        className="clean-input w-full text-xs"
                      />
                    </div>
                    <div>
                      <label className="clean-label">Danh mục biểu mẫu</label>
                      <select
                        value={uploadCategoryId}
                        onChange={(e) => setUploadCategoryId(e.target.value)}
                        className="clean-input w-full text-xs cursor-pointer"
                      >
                        <option value="">Chưa phân loại</option>
                        {getFlattenedCategoryOptions().map(category => (
                          <option key={category.id} value={category.id}>
                            {category.label}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>

                  <div>
                    <label className="clean-label">Tệp Word nguồn (.doc, .docx)</label>
                    <div
                      className={`border-2 border-dashed border-slate-200 hover:border-emerald-400 rounded-2xl p-8 text-center cursor-pointer transition-all relative ${
                        isDragActive ? 'drag-active scale-[1.005]' : 'bg-slate-50/50'
                      }`}
                      onDragOver={(e) => {
                        e.preventDefault();
                        setIsDragActive(true);
                      }}
                      onDragLeave={(e) => {
                        e.preventDefault();
                        setIsDragActive(false);
                      }}
                      onDrop={(e) => {
                        e.preventDefault();
                        setIsDragActive(false);
                        if (e.dataTransfer.files?.length) {
                          handleSelectUploadFiles(e.dataTransfer.files);
                        }
                      }}
                    >
                      <input
                        type="file"
                        multiple
                        accept=".doc,.docx"
                        onChange={(e) => {
                          if (e.target.files?.length) {
                            handleSelectUploadFiles(e.target.files);
                          }
                        }}
                        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                      />
                      <div className="space-y-2">
                        <svg className="w-10 h-10 text-slate-300 mx-auto shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="1.5">
                          <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m3.75 9v6m3-3H9m1.5-12H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
                        </svg>
                        <div className="text-sm text-slate-700 font-semibold">
                          {uploadFiles.length > 0 ? `Đã chọn ${uploadFiles.length} file` : 'Kéo thả file hoặc click để chọn'}
                        </div>
                        <div className="text-[11px] text-slate-400">Hỗ trợ tối đa 10 file .doc hoặc .docx, mỗi file không quá 1MB</div>
                      </div>
                    </div>
                  </div>

                  {uploadFiles.length > 0 && (
                    <div className="rounded-2xl border border-slate-200 bg-slate-50/70 p-4">
                      <div className="flex items-center justify-between gap-3 mb-3">
                        <p className="text-[11px] font-bold uppercase tracking-wider text-slate-500">Danh sách file sẽ tải lên</p>
                        <button
                          type="button"
                          onClick={() => {
                            setUploadFiles([]);
                            setUploadName('');
                          }}
                          className="text-[11px] font-semibold text-slate-500 hover:text-slate-700 cursor-pointer"
                        >
                          Xóa danh sách
                        </button>
                      </div>
                      <div className="space-y-2 max-h-48 overflow-y-auto">
                        {uploadFiles.map(file => (
                          <div key={`${file.name}-${file.size}-${file.lastModified}`} className="flex items-center justify-between gap-3 rounded-xl bg-white px-3 py-2 border border-slate-200 text-xs">
                            <span className="truncate text-slate-700 font-medium">{file.name}</span>
                            <span className="shrink-0 text-slate-400">{(file.size / 1024).toFixed(0)} KB</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  <div className="flex flex-wrap gap-3">
                    <button
                      type="submit"
                      disabled={isUploading}
                      className="btn-premium bg-emerald-600 hover:bg-emerald-700 text-white font-semibold py-2.5 px-4 text-xs transition-all shadow-sm hover:shadow-md flex items-center justify-center gap-2 disabled:opacity-50 cursor-pointer"
                    >
                      {isUploading ? 'Đang xử lý đợt upload...' : 'Tải lên biểu mẫu'}
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setUploadFiles([]);
                        setUploadName('');
                        setUploadCategoryId('');
                        setUploadResults([]);
                      }}
                      className="btn-premium bg-white hover:bg-slate-50 text-slate-700 font-semibold py-2.5 px-4 text-xs border border-slate-200 transition-all cursor-pointer"
                    >
                      Làm trống
                    </button>
                  </div>
                </form>

                {uploadResults.length > 0 && (
                  <div className="rounded-2xl border border-slate-200 bg-slate-50/60 p-4 space-y-3">
                    <div className="flex items-center justify-between gap-3">
                      <h3 className="text-xs font-bold text-slate-800 tracking-tight">Kết quả đợt upload</h3>
                      <span className="text-[11px] text-slate-500">{uploadResults.filter(item => item.status === 'success').length}/{uploadResults.length} thành công</span>
                    </div>
                    <div className="space-y-2">
                      {uploadResults.map((result) => (
                        <div key={`${result.fileName}-${result.templateId || result.error}`} className="rounded-xl border border-slate-200 bg-white p-3 text-xs">
                          <div className="flex items-start justify-between gap-3">
                            <div className="min-w-0">
                              <div className="font-semibold text-slate-800 truncate">{result.fileName}</div>
                              <div className={`mt-1 ${result.status === 'success' ? 'text-emerald-600' : 'text-rose-600'}`}>
                                {result.status === 'success'
                                  ? `${result.variablesCount} biến, tên biểu mẫu: ${result.name}`
                                  : result.error}
                              </div>
                            </div>
                            {result.status === 'success' && (
                              <button
                                type="button"
                                onClick={() => handleOpenConfig(result.templateId)}
                                className="shrink-0 text-[11px] font-semibold text-emerald-700 hover:text-emerald-800 cursor-pointer"
                              >
                                Mở chi tiết
                              </button>
                            )}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}

          {activeView === 'dashboard' && sidebarActiveMenu === 'categories' && userRole === 'congchung' && (
            <div className="grid grid-cols-1 xl:grid-cols-[minmax(0,1.25fr)_380px] gap-6 animate-fade-up">
              <div className="glass-panel p-6 bg-white space-y-5">
                <div>
                  <span className="text-[11px] text-emerald-600 font-bold uppercase tracking-wider font-display">Danh mục thư viện</span>
                  <h2 className="text-base font-bold text-slate-800 mt-1 font-display">Cây danh mục biểu mẫu</h2>
                  <p className="text-xs text-slate-400 mt-2 leading-relaxed">
                    Mỗi nhánh có thể mở rộng hoặc thu gọn riêng. Cây này chỉ quản lý tổ chức thư viện, không thay thế quan hệ file gốc và file con trong nghiệp vụ.
                  </p>
                </div>

                <div className="flex flex-wrap gap-3">
                  <button
                    type="button"
                    onClick={() => setExpandedCategoryIds(new Set(categories.map(category => category.id)))}
                    className="btn-premium bg-white hover:bg-slate-50 text-slate-700 font-semibold py-2 px-3 text-xs border border-slate-200 transition-all cursor-pointer"
                  >
                    Mở tất cả
                  </button>
                  <button
                    type="button"
                    onClick={() => setExpandedCategoryIds(new Set())}
                    className="btn-premium bg-white hover:bg-slate-50 text-slate-700 font-semibold py-2 px-3 text-xs border border-slate-200 transition-all cursor-pointer"
                  >
                    Thu gọn tất cả
                  </button>
                </div>

                <div className="border border-slate-200 rounded-2xl p-4 bg-slate-50/60">
                  <div className="space-y-2">
                    <button
                      type="button"
                      onClick={() => setSelectedCategoryId('all')}
                      className={`w-full text-left px-3 py-2 rounded-xl border transition-all cursor-pointer ${
                        selectedCategoryId === 'all'
                          ? 'bg-slate-900 border-slate-900 text-white'
                          : 'bg-white border-slate-200 text-slate-700 hover:border-emerald-200 hover:bg-emerald-50/40'
                      }`}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="text-xs font-semibold">Tất cả danh mục</span>
                        <span className={`text-[10px] font-bold ${selectedCategoryId === 'all' ? 'text-white/80' : 'text-slate-400'}`}>{templates.length}</span>
                      </div>
                    </button>
                    <div className="space-y-1 max-h-[540px] overflow-y-auto pr-1">
                      {renderCategoryTree()}
                    </div>
                  </div>
                </div>
              </div>

              <div className="glass-panel p-6 bg-white space-y-4 h-fit">
                <div>
                  <span className="text-[11px] text-slate-500 font-bold uppercase tracking-wider font-display">Tạo danh mục</span>
                  <h3 className="text-sm font-bold text-slate-800 mt-1 font-display">Thêm nhánh mới vào cây</h3>
                  <p className="text-xs text-slate-400 mt-2 leading-relaxed">
                    Bạn có thể tạo danh mục gốc hoặc gắn nó vào một danh mục cha để tạo cây nhiều tầng, ví dụ: Hợp đồng → Đất đai → Chuyển nhượng.
                  </p>
                </div>

                <form onSubmit={handleCreateCategory} className="space-y-3">
                  <select
                    value={newCategoryParentId}
                    onChange={(e) => setNewCategoryParentId(e.target.value)}
                    className="clean-input w-full text-xs cursor-pointer"
                  >
                    <option value="">Danh mục gốc</option>
                    {getFlattenedCategoryOptions().map(category => (
                      <option key={category.id} value={category.id}>
                        {category.label}
                      </option>
                    ))}
                  </select>
                  <input
                    type="text"
                    value={newCategoryName}
                    onChange={(e) => setNewCategoryName(e.target.value)}
                    placeholder="Ví dụ: Hợp đồng thế chấp"
                    className="clean-input w-full text-xs"
                  />
                  <button
                    type="submit"
                    className="btn-premium w-full bg-emerald-600 hover:bg-emerald-700 text-white font-semibold py-2.5 text-xs transition-all cursor-pointer"
                  >
                    Thêm danh mục
                  </button>
                </form>

                <div className="border-t border-slate-100 pt-4 space-y-2">
                  <div className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Ghi chú sử dụng</div>
                  <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs text-slate-600 leading-relaxed">
                    Dùng nút tam giác ở mỗi node để mở hoặc thu nhánh. Khi chọn một danh mục, thư viện biểu mẫu sẽ lọc theo toàn bộ nhánh con của danh mục đó.
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Guide View */}
          {activeView === 'dashboard' && sidebarActiveMenu === 'guide' && (
            <div className="glass-panel p-8 bg-white animate-fade-up max-w-3xl mx-auto space-y-6">
              <div className="text-center pb-6 border-b border-slate-100">
                <h3 className="text-lg font-bold text-slate-900 font-display">Hướng dẫn sử dụng dịch vụ Công chứng số</h3>
                <p className="text-xs text-slate-400 mt-2">Dựng hồ sơ pháp lý hoàn chỉnh chỉ với 4 bước trực tuyến</p>
              </div>

              <div className="space-y-6 font-sans">
                <div className="flex gap-4">
                  <div className="w-8 h-8 rounded-xl bg-emerald-50 border border-emerald-100 text-emerald-600 font-bold flex items-center justify-center shrink-0">1</div>
                  <div className="space-y-1">
                    <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wide font-display">Bước 1: Chọn biểu mẫu phù hợp</h4>
                    <p className="text-xs text-slate-500 leading-relaxed">
                      Duyệt qua danh mục các biểu mẫu giao dịch phổ biến tại mục <strong>Duyệt biểu mẫu</strong>. Nhấn nút <strong>Điền hồ sơ</strong> để bắt đầu.
                    </p>
                  </div>
                </div>

                <div className="flex gap-4">
                  <div className="w-8 h-8 rounded-xl bg-emerald-50 border border-emerald-100 text-emerald-600 font-bold flex items-center justify-center shrink-0">2</div>
                  <div className="space-y-1">
                    <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wide font-display">Bước 2: Điền thông tin cá nhân & giao dịch</h4>
                    <p className="text-xs text-slate-500 leading-relaxed">
                      Nhập Tên, Số điện thoại để văn phòng tiện liên hệ. Sau đó làm theo trình Wizard 4 bước được phân loại khoa học (Bên bán, Bên mua, Tài sản, Hợp đồng) để hoàn thiện nội dung.
                    </p>
                  </div>
                </div>

                <div className="flex gap-4">
                  <div className="w-8 h-8 rounded-xl bg-emerald-50 border border-emerald-100 text-emerald-600 font-bold flex items-center justify-center shrink-0">3</div>
                  <div className="space-y-1">
                    <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wide font-display">Bước 3: Chọn thêm biểu mẫu phụ lục con</h4>
                    <p className="text-xs text-slate-500 leading-relaxed">
                      Hệ thống tự động liên kết các biểu mẫu con đi kèm (như Tờ khai thuế TNCN, Tờ khai Lệ phí trước bạ). Chọn những biểu mẫu con cần thiết, dữ liệu từ file gốc sẽ tự động đồng bộ sang.
                    </p>
                  </div>
                </div>

                <div className="flex gap-4">
                  <div className="w-8 h-8 rounded-xl bg-emerald-50 border border-emerald-100 text-emerald-600 font-bold flex items-center justify-center shrink-0">4</div>
                  <div className="space-y-1">
                    <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wide font-display">Bước 4: Xuất tài liệu & nộp hồ sơ</h4>
                    <p className="text-xs text-slate-500 leading-relaxed">
                      Xem trước toàn bộ tài liệu đã điền chuẩn PDF. Nhấn <strong>Xác nhận & Xuất File</strong> để tải trọn bộ hồ sơ dưới dạng `.zip` hoặc `.docx` đã điền hoàn hảo. Hồ sơ của bạn sẽ được tự động gửi đến Văn phòng công chứng để chuẩn bị bản in vật lý.
                    </p>
                  </div>
                </div>
              </div>

              <div className="bg-slate-50 p-4 border border-slate-200 rounded-2xl text-[11px] text-slate-500 flex items-start gap-2 leading-relaxed">
                <span>ℹ️</span>
                <span>Mọi thắc mắc trong quá trình nộp hồ sơ, quý khách vui lòng liên hệ hotline <strong>090.123.4567</strong> hoặc đến trực tiếp Văn phòng Công chứng Trung tâm để được tư vấn miễn phí.</span>
              </div>
            </div>
          )}

          {/* Submissions View */}
          {activeView === 'dashboard' && sidebarActiveMenu === 'submissions' && (
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
              {/* Left Column: Submissions List Table */}
              <div className={`glass-panel p-6 bg-white animate-fade-up space-y-4 transition-all duration-300 ${expandedSubmissionId ? 'lg:col-span-6' : 'lg:col-span-12'}`}>
                <div className="flex justify-between items-center pb-4 border-b border-slate-100">
                  <h3 className="text-base font-bold text-slate-800 font-display">Lịch sử nhận hồ sơ</h3>
                  <button
                    onClick={fetchSubmissions}
                    className="px-3.5 py-1.5 text-xs font-semibold text-slate-700 bg-slate-50 hover:bg-slate-100 rounded-xl border border-slate-200 transition-colors cursor-pointer flex items-center gap-1"
                  >
                    ↺ Tải lại
                  </button>
                </div>

                {isLoadingSubmissions ? (
                  <div className="py-12 text-center text-slate-400">
                    <div className="w-6 h-6 border-2 border-emerald-500/30 border-t-emerald-500 rounded-full animate-spin mx-auto mb-3"></div>
                    Đang tải danh sách hồ sơ...
                  </div>
                ) : submissionHistory.length === 0 ? (
                  <div className="py-12 text-center text-slate-400">
                    <div className="text-3xl mb-2">📥</div>
                    Chưa nhận được hồ sơ nào từ khách hàng.
                  </div>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs font-sans border-collapse">
                      <thead>
                        <tr className="border-b border-slate-200 text-slate-500 font-bold uppercase tracking-wider">
                          <th className="py-3.5 px-4 w-[35%]">Khách hàng</th>
                          {!expandedSubmissionId && <th className="py-3.5 px-4 w-[20%]">Số điện thoại</th>}
                          <th className="py-3.5 px-4 w-[35%]">Biểu mẫu nộp</th>
                          {!expandedSubmissionId && <th className="py-3.5 px-4 w-[15%]">Ngày nộp</th>}
                          <th className="py-3.5 px-4 text-right w-[15%]">Thao tác</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 text-slate-700 font-medium">
                        {submissionHistory.map(sub => {
                          const isExpanded = expandedSubmissionId === sub.id;
                          return (
                            <tr 
                              key={sub.id}
                              className={`hover:bg-slate-50/70 transition-colors cursor-pointer ${
                                isExpanded ? 'bg-emerald-50/40 border-l-2 border-emerald-500 font-semibold' : ''
                              }`}
                              onClick={() => handleToggleExpandSubmission(sub.id)}
                            >
                              <td className="py-4 px-4 font-bold text-slate-900 flex items-center gap-2">
                                {sub.customer_name}
                              </td>
                              {!expandedSubmissionId && <td className="py-4 px-4 font-mono">{sub.customer_phone || 'N/A'}</td>}
                              <td className="py-4 px-4 max-w-[180px] truncate">{sub.template_name}</td>
                              {!expandedSubmissionId && <td className="py-4 px-4 text-slate-500">{new Date(sub.completed_at).toLocaleString('vi-VN')}</td>}
                              <td className="py-4 px-4 text-right space-x-1.5" onClick={e => e.stopPropagation()}>
                                <button
                                  type="button"
                                  onClick={() => handleToggleExpandSubmission(sub.id)}
                                  className={`px-2 py-1 text-[10px] font-bold rounded-lg border transition-all cursor-pointer inline-flex items-center gap-1 shadow-sm ${
                                    isExpanded 
                                      ? 'text-slate-700 bg-slate-150 border-slate-350 hover:bg-slate-205' 
                                      : 'text-slate-700 bg-slate-50 border-slate-200 hover:bg-slate-100'
                                  }`}
                                >
                                  👁️ Xem
                                </button>
                                <a
                                  href={`http://localhost:5000/api/submissions/${sub.id}/download`}
                                  download
                                  className="px-2 py-1 text-[10px] font-bold text-emerald-700 bg-emerald-50 hover:bg-emerald-100 border border-emerald-250 rounded-lg transition-colors cursor-pointer inline-flex items-center gap-1 shadow-sm"
                                >
                                  📥 Tải
                                </a>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>

              {/* Right Column: Tab/Panel next to the list to see detailed what was filled */}
              {expandedSubmissionId && (() => {
                const sub = submissionHistory.find(s => s.id === expandedSubmissionId);
                if (!sub) return null;
                return (
                  <div className="lg:col-span-6 glass-panel p-6 bg-white border border-slate-200 rounded-2xl shadow-md space-y-5 max-h-[85vh] overflow-y-auto animate-fade-right">
                    <div className="flex justify-between items-center pb-3 border-b border-slate-150">
                      <div className="space-y-0.5">
                        <span className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Chi tiết hồ sơ nộp</span>
                        <h4 className="text-sm font-black text-slate-900 font-display truncate max-w-[280px]" title={sub.customer_name}>
                          {sub.customer_name}
                        </h4>
                      </div>
                      <button
                        type="button"
                        onClick={() => setExpandedSubmissionId(null)}
                        className="w-7 h-7 rounded-full bg-slate-50 hover:bg-slate-150 text-slate-500 border border-slate-200 flex items-center justify-center text-xs font-bold cursor-pointer transition-colors shadow-sm"
                        title="Đóng chi tiết"
                      >
                        ✕
                      </button>
                    </div>

                    <div className="space-y-4">
                      {/* Meta info card */}
                      <div className="bg-slate-50 border border-slate-150 rounded-xl p-3.5 space-y-2 text-[11px] text-slate-600 font-sans shadow-sm">
                        <div className="flex justify-between">
                          <span>Số điện thoại:</span>
                          <strong className="text-slate-800 font-mono">{sub.customer_phone || 'N/A'}</strong>
                        </div>
                        <div className="flex justify-between">
                          <span>Ngày nộp:</span>
                          <strong className="text-slate-800">{new Date(sub.completed_at).toLocaleString('vi-VN')}</strong>
                        </div>
                        <div className="flex justify-between items-center">
                          <span>Trạng thái:</span>
                          <span className={`px-2 py-0.5 rounded-full text-[9px] font-bold uppercase ${
                            sub.status === 'completed' 
                              ? 'bg-emerald-50 text-emerald-700 border border-emerald-200' 
                              : 'bg-rose-50 text-rose-700 border border-rose-200'
                          }`}>
                            {sub.status === 'completed' ? 'Hoàn thành' : 'Thất bại'}
                          </span>
                        </div>
                      </div>

                      {/* Detail form data */}
                      {loadingDetailsForSubId === sub.id ? (
                        <div className="py-12 text-center text-slate-400 text-xs font-sans">
                          <div className="w-5 h-5 border-2 border-emerald-500/30 border-t-emerald-500 rounded-full animate-spin mx-auto mb-2.5"></div>
                          Đang tải dữ liệu hồ sơ...
                        </div>
                      ) : !submissionDetailsMap[sub.id] ? (
                        <div className="py-8 text-center text-slate-400 text-xs italic font-sans border border-dashed rounded-xl border-slate-200">
                          Không tải được dữ liệu chi tiết từ server.
                        </div>
                      ) : (() => {
                        const detail = submissionDetailsMap[sub.id];
                        const values = detail.submission?.values_json?.values || {};
                        const masterFields = detail.masterFields || [];
                        const childFields = detail.childFields || [];
                        const selectedChildIds = detail.submission?.values_json?.selectedChildIds || [];

                        // Group child fields by template_id
                        const childTemplatesMap = {};
                        childFields.forEach(field => {
                          if (!childTemplatesMap[field.template_id]) {
                            childTemplatesMap[field.template_id] = {
                              name: field.template_name,
                              is_repeated: field.is_repeated === 1,
                              fields: []
                            };
                          }
                          childTemplatesMap[field.template_id].fields.push(field);
                        });

                        return (
                          <div className="space-y-4">
                            {/* Tabs selector */}
                            <div className="flex border-b border-slate-200/80 pb-0.5">
                              <button
                                type="button"
                                onClick={() => setDetailActiveTab('fields')}
                                className={`flex-1 pb-2 text-[11px] font-bold text-center border-b-2 transition-all duration-150 cursor-pointer ${
                                  detailActiveTab === 'fields'
                                    ? 'border-emerald-500 text-emerald-600'
                                    : 'border-transparent text-slate-400 hover:text-slate-600'
                                }`}
                              >
                                📋 Dữ liệu nhập
                              </button>
                              <button
                                type="button"
                                onClick={() => setDetailActiveTab('preview')}
                                className={`flex-1 pb-2 text-[11px] font-bold text-center border-b-2 transition-all duration-150 cursor-pointer ${
                                  detailActiveTab === 'preview'
                                    ? 'border-emerald-500 text-emerald-600'
                                    : 'border-transparent text-slate-400 hover:text-slate-600'
                                }`}
                              >
                                👁️ Văn bản xem trước
                              </button>
                            </div>

                            {detailActiveTab === 'fields' ? (
                              <div className="space-y-4">
                                {/* Master fields */}
                                <div className="bg-white rounded-xl border border-slate-150 p-4 space-y-3.5 shadow-sm">
                                  <div className="text-[10px] font-bold text-slate-800 border-b border-slate-100 pb-2 uppercase tracking-wide font-display">
                                    📋 Biểu mẫu chính: {detail.submission?.template_name}
                                  </div>
                                  {masterFields.length === 0 ? (
                                    <div className="text-xs text-slate-400 italic font-sans">Không có thông tin trường điền.</div>
                                  ) : (
                                    <div className="grid grid-cols-1 gap-y-3.5">
                                      {masterFields.map(field => {
                                        const val = values[field.key_name];
                                        const displayVal = typeof val === 'boolean' 
                                          ? (val ? '✓ Có / Đồng ý' : '✗ Không') 
                                          : (val || '—');
                                        return (
                                          <div key={field.key_name} className="flex flex-col gap-0.5 text-xs font-sans border-b border-slate-50 last:border-b-0 pb-1.5 last:pb-0">
                                            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">{field.label}</span>
                                            <span className="font-semibold text-slate-800 break-words mt-0.5">{displayVal}</span>
                                          </div>
                                        );
                                      })}
                                    </div>
                                  )}
                                </div>

                                {/* Non-repeated child fields */}
                                {Object.entries(childTemplatesMap)
                                  .filter(([childId]) => selectedChildIds.includes(childId) && !childTemplatesMap[childId].is_repeated)
                                  .map(([childId, childData]) => (
                                    <div key={childId} className="bg-white rounded-xl border border-slate-150 p-4 space-y-3.5 shadow-sm">
                                      <div className="text-[10px] font-bold text-slate-800 border-b border-slate-100 pb-2 uppercase tracking-wide font-display">
                                        📄 Biểu mẫu con: {childData.name}
                                      </div>
                                      <div className="grid grid-cols-1 gap-y-3.5">
                                        {childData.fields.map(field => {
                                          const val = values[field.key_name];
                                          const displayVal = typeof val === 'boolean' 
                                            ? (val ? '✓ Có / Đồng ý' : '✗ Không') 
                                            : (val || '—');
                                          return (
                                            <div key={field.key_name} className="flex flex-col gap-0.5 text-xs font-sans border-b border-slate-50 last:border-b-0 pb-1.5 last:pb-0">
                                              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">{field.label}</span>
                                              <span className="font-semibold text-slate-800 break-words mt-0.5">{displayVal}</span>
                                            </div>
                                          );
                                        })}
                                      </div>
                                    </div>
                                  ))}

                                {/* Repeated child fields */}
                                {Object.entries(childTemplatesMap)
                                  .filter(([childId]) => selectedChildIds.includes(childId) && childTemplatesMap[childId].is_repeated)
                                  .map(([childId, childData]) => {
                                    const records = values[childId] || [];
                                    return (
                                      <div key={childId} className="bg-white rounded-xl border border-slate-150 p-4 space-y-3.5 shadow-sm">
                                        <div className="text-[10px] font-bold text-slate-800 border-b border-slate-100 pb-2 flex items-center justify-between uppercase tracking-wide font-display">
                                          <span>📂 Biểu mẫu con: {childData.name} (Lặp lại)</span>
                                          <span className="text-[9px] font-bold text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded-lg select-none">
                                            {records.length} bản ghi
                                          </span>
                                        </div>
                                        {records.length === 0 ? (
                                          <div className="text-xs text-slate-400 italic font-sans">Không có bản ghi nào.</div>
                                        ) : (
                                          <div className="space-y-4">
                                            {records.map((record, rIdx) => (
                                              <div key={rIdx} className="bg-slate-50/50 border border-slate-200/80 rounded-xl p-3.5 space-y-3">
                                                <div className="text-[10px] font-bold text-slate-500 uppercase tracking-wider border-b border-slate-200/50 pb-1.5 font-display">
                                                  Bản ghi #{rIdx + 1}
                                                </div>
                                                <div className="grid grid-cols-1 gap-y-3">
                                                  {childData.fields.map(field => {
                                                    const val = record[field.key_name];
                                                    const displayVal = typeof val === 'boolean' 
                                                      ? (val ? '✓ Có / Đồng ý' : '✗ Không') 
                                                      : (val || '—');
                                                    return (
                                                      <div key={field.key_name} className="flex flex-col gap-0.5 text-xs font-sans border-b border-slate-100/30 last:border-b-0 pb-1 last:pb-0">
                                                        <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">{field.label}</span>
                                                        <span className="font-semibold text-slate-700 break-words mt-0.5">{displayVal}</span>
                                                      </div>
                                                    );
                                                  })}
                                                </div>
                                              </div>
                                            ))}
                                          </div>
                                        )}
                                      </div>
                                    );
                                  })}
                              </div>
                            ) : (
                              <div className="space-y-4">
                                {(() => {
                                  const subFiles = submissionFilesMap[sub.id] || [];
                                  const selectedFile = activePreviewFilename || subFiles[0];
                                  return (
                                    <div className="space-y-4">
                                      {/* File tabs selector if there are multiple files */}
                                      {subFiles.length > 1 && (
                                        <div className="flex flex-wrap gap-1.5 bg-slate-50 border border-slate-200/60 p-2 rounded-xl">
                                          <div className="w-full text-[9px] font-bold text-slate-400 uppercase tracking-wider mb-1 px-1 font-display">
                                            Danh sách tệp tin trong hồ sơ:
                                          </div>
                                          {subFiles.map(filename => (
                                            <button
                                              key={filename}
                                              type="button"
                                              onClick={() => setActivePreviewFilename(filename)}
                                              className={`px-3 py-1.5 rounded-lg text-[10px] font-bold transition-all duration-150 cursor-pointer ${
                                                selectedFile === filename
                                                  ? 'bg-emerald-600 text-white shadow-sm border border-emerald-500'
                                                  : 'bg-white hover:bg-slate-50 text-slate-600 hover:text-slate-800 border border-slate-200'
                                              }`}
                                            >
                                              📄 {filename}
                                            </button>
                                          ))}
                                        </div>
                                      )}
                                      
                                      {/* Document preview container */}
                                      {selectedFile ? (
                                        <div className="bg-white rounded-xl border border-slate-150 p-4 space-y-3 shadow-sm">
                                          <div className="text-[10px] font-bold text-slate-800 border-b border-slate-100 pb-2 uppercase tracking-wide font-display flex justify-between items-center">
                                            <span>👁️ Xem trước: {selectedFile}</span>
                                            <span className="text-[9px] font-semibold text-slate-400 font-sans">(Trộn động từ DB)</span>
                                          </div>
                                          <DocxFilePreview subId={sub.id} filename={selectedFile} />
                                        </div>
                                      ) : (
                                        <div className="bg-white rounded-xl border border-slate-150 p-6 text-center text-slate-400 text-xs italic font-sans shadow-sm">
                                          Không tìm thấy tệp tin nào để xem trước.
                                        </div>
                                      )}
                                    </div>
                                  );
                                })()}
                              </div>
                            )}
                          </div>
                        );
                      })()}

                      {/* Generated files list */}
                      <div className="bg-white rounded-xl border border-slate-150 p-4 space-y-3.5 shadow-sm font-sans">
                        <div className="text-[10px] font-bold text-slate-800 border-b border-slate-100 pb-2 uppercase tracking-wide font-display">
                          📄 Tệp kết quả sinh động (.docx):
                        </div>
                        {loadingFilesForSubId === sub.id ? (
                          <div className="py-4 text-center text-slate-400 text-xs flex items-center justify-center gap-1.5">
                            <div className="w-3.5 h-3.5 border border-emerald-500/30 border-t-emerald-500 rounded-full animate-spin"></div>
                            Đang quét danh sách file...
                          </div>
                        ) : !submissionFilesMap[sub.id] || submissionFilesMap[sub.id].length === 0 ? (
                          <div className="py-2 text-slate-400 italic text-[11px] font-sans">Không tìm thấy file chi tiết.</div>
                        ) : (
                          <div className="flex flex-col gap-2">
                            {submissionFilesMap[sub.id].map(filename => (
                              <div key={filename} className="flex justify-between items-center bg-slate-50 border border-slate-150 p-2.5 rounded-xl">
                                <span className="text-[11px] font-semibold text-slate-700 truncate pr-3" title={filename}>{filename}</span>
                                <a
                                  href={`http://localhost:5000/api/submissions/${sub.id}/download-file?filename=${encodeURIComponent(filename)}`}
                                  download
                                  className="p-1 hover:bg-emerald-50 border border-transparent hover:border-emerald-250 text-emerald-600 rounded transition-colors shrink-0 cursor-pointer"
                                  title={`Tải file ${filename}`}
                                >
                                  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2.2">
                                    <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                                  </svg>
                                </a>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                );
              })()}
            </div>
          )}

          {/* Stats View */}
          {activeView === 'dashboard' && sidebarActiveMenu === 'stats' && (
            <div className="glass-panel p-6 bg-white animate-fade-up space-y-6">
              <div className="flex justify-between items-center pb-4 border-b border-slate-100">
                <h3 className="text-base font-bold text-slate-800 font-display">Thống kê tổng quan hoạt động</h3>
                <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider bg-slate-50 border border-slate-200 px-3 py-1 rounded-lg">Realtime</span>
              </div>

              {/* Stats Metrics Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4">
                <div className="glass-panel p-5 bg-white border border-slate-150 hover:border-slate-250 hover:shadow-sm transition-all flex flex-col justify-between min-h-[100px]">
                  <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Tổng biểu mẫu</span>
                  <div className="flex items-baseline gap-2 mt-2">
                    <span className="text-2xl font-black text-slate-900 tracking-tight font-display">{templates.length}</span>
                    <span className="text-[10px] font-bold text-slate-500">tệp mẫu</span>
                  </div>
                </div>

                <div className="glass-panel p-5 bg-white border border-slate-150 hover:border-slate-250 hover:shadow-sm transition-all flex flex-col justify-between min-h-[100px]">
                  <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Biểu mẫu hoạt động</span>
                  <div className="flex items-baseline gap-2 mt-2">
                    <span className="text-2xl font-black text-emerald-600 tracking-tight font-display">
                      {templates.filter(t => t.status === 'active').length}
                    </span>
                    <span className="text-[10px] font-bold text-emerald-500">active</span>
                  </div>
                </div>

                <div className="glass-panel p-5 bg-white border border-slate-150 hover:border-slate-250 hover:shadow-sm transition-all flex flex-col justify-between min-h-[100px]">
                  <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Tổng hồ sơ nhận</span>
                  <div className="flex items-baseline gap-2 mt-2">
                    <span className="text-2xl font-black text-slate-900 tracking-tight font-display">{submissionHistory.length}</span>
                    <span className="text-[10px] font-bold text-slate-500">lượt nộp</span>
                  </div>
                </div>

                <div className="glass-panel p-5 bg-white border border-slate-150 hover:border-slate-250 hover:shadow-sm transition-all flex flex-col justify-between min-h-[100px]">
                  <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Hồ sơ hoàn thành</span>
                  <div className="flex items-baseline gap-2 mt-2">
                    <span className="text-2xl font-black text-emerald-600 tracking-tight font-display">
                      {submissionHistory.filter(s => s.status === 'completed').length}
                    </span>
                    <span className="text-[10px] font-bold text-emerald-500">thành công</span>
                  </div>
                </div>
              </div>

              {/* Detail analytics panels */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-5 pt-2">
                {/* Phân tích tỷ lệ nộp hồ sơ */}
                <div className="glass-panel p-5 bg-white border border-slate-200 space-y-4">
                  <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wide font-display">Biểu mẫu được dùng nhiều nhất</h4>
                  {submissionHistory.length === 0 ? (
                    <div className="py-8 text-center text-slate-400 text-xs italic">Chưa có đủ số liệu hoạt động.</div>
                  ) : (
                    <div className="space-y-3 font-sans">
                      {(() => {
                        const usageMap = {};
                        submissionHistory.forEach(s => {
                          usageMap[s.template_name] = (usageMap[s.template_name] || 0) + 1;
                        });
                        const sorted = Object.entries(usageMap)
                          .sort((a, b) => b[1] - a[1])
                          .slice(0, 5);
                        const maxCount = sorted[0]?.[1] || 1;

                        return sorted.map(([name, count]) => {
                          const percent = Math.round((count / maxCount) * 100);
                          return (
                            <div key={name} className="space-y-1.5">
                              <div className="flex justify-between text-xs font-bold text-slate-700">
                                <span className="truncate pr-4">{name}</span>
                                <span className="text-slate-900">{count} lượt</span>
                              </div>
                              <div className="h-2 w-full bg-slate-100 rounded-full overflow-hidden">
                                <div className="h-full bg-emerald-500 rounded-full" style={{ width: `${percent}%` }}></div>
                              </div>
                            </div>
                          );
                        });
                      })()}
                    </div>
                  )}
                </div>

                {/* Hoạt động gần đây */}
                <div className="glass-panel p-5 bg-white border border-slate-200 space-y-4">
                  <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wide font-display">Hoạt động nộp gần đây</h4>
                  {submissionHistory.length === 0 ? (
                    <div className="py-8 text-center text-slate-400 text-xs italic">Chưa có lịch sử nộp hồ sơ.</div>
                  ) : (
                    <div className="space-y-3.5 max-h-[220px] overflow-y-auto pr-1 text-xs font-sans">
                      {submissionHistory.slice(0, 5).map(sub => (
                        <div key={sub.id} className="flex justify-between items-start gap-2.5 pb-3 border-b border-slate-100 last:border-b-0 last:pb-0">
                          <div className="space-y-0.5 min-w-0">
                            <div className="font-bold text-slate-800 truncate">{sub.customer_name}</div>
                            <div className="text-[10px] text-slate-400">Đã nộp: <strong className="text-slate-600">{sub.template_name}</strong></div>
                          </div>
                          <div className="text-right shrink-0">
                            <span className="text-[10px] text-slate-400 block">{new Date(sub.completed_at).toLocaleTimeString('vi-VN')}</span>
                            <span className="text-[9px] font-extrabold text-emerald-600 uppercase mt-0.5 inline-block">Thành công</span>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* Link Management Modal */}
          {showLinkModal && linkingTemplate && (() => {
            const filteredAvailable = availableToLink.filter(t => 
              t.name.toLowerCase().includes(linkSearch.toLowerCase())
            );
            return (
              <div className="fixed inset-0 z-50 flex items-center justify-center bg-zinc-900/60 backdrop-blur-sm p-4">
                <div className="bg-white border border-zinc-200 w-full max-w-xl rounded-lg shadow-2xl flex flex-col max-h-[85vh] overflow-hidden animate-fade-up">
                  {/* Header */}
                  <div className="bg-zinc-900 text-white px-6 py-4 flex justify-between items-center border-b border-zinc-850">
                    <div>
                      <h3 className="font-bold text-sm text-emerald-500 uppercase tracking-wider font-display flex items-center gap-2">
                        <svg className="w-4 h-4 text-emerald-500 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                          <path strokeLinecap="round" strokeLinejoin="round" d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101m-.758-4.899a4 4 0 005.656 0l4-4a4 4 0 00-5.656-5.656l-1.1 1.1" />
                        </svg>
                        Quản lý liên kết biểu mẫu
                      </h3>
                      <p className="text-xs text-zinc-400 mt-1 line-clamp-1">File gốc: <strong>{linkingTemplate.name}</strong></p>
                    </div>
                    <button 
                      onClick={() => { setShowLinkModal(false); setLinkingTemplate(null); setLinkSearch(''); }}
                      className="text-zinc-400 hover:text-white text-xl font-bold p-1 select-none cursor-pointer transition-colors"
                    >
                      ×
                    </button>
                  </div>
                  
                  {/* Body */}
                  <div className="p-6 overflow-y-auto space-y-6 flex-1 font-sans">
                    
                    {/* Cấu trúc cây Master-Child (Task 10) */}
                    <div className="bg-zinc-50 p-4 border border-zinc-200 rounded-lg space-y-3 shadow-inner">
                      <h4 className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider font-display flex items-center gap-1.5">
                        <svg className="w-4 h-4 text-zinc-400 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                          <path strokeLinecap="round" strokeLinejoin="round" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-2 0c0 .993-.241 1.929-.668 2.754l-1.524-1.525a3.997 3.997 0 00.192-1.229 4 4 0 10-4 4 3.997 3.997 0 001.23-.192l1.524 1.524A7.965 7.965 0 0110 18a8 8 0 01-8-8 8 8 0 018-8 8 8 0 018 8z" />
                        </svg>
                        Sơ đồ cấu trúc tài liệu (Master-Child)
                      </h4>
                      
                      <div className="pl-2 space-y-2">
                        {/* Gốc (Master) */}
                        <div className="flex items-center gap-2">
                          <span className="w-2.5 h-2.5 rounded-full bg-emerald-600 border border-white ring-2 ring-emerald-100 shrink-0"></span>
                          <span className="text-xs font-bold text-zinc-800 bg-emerald-50 px-2.5 py-1 rounded-lg border border-emerald-100 max-w-full truncate">{linkingTemplate.name} (Gốc)</span>
                        </div>
                        
                        {/* Các nhánh con (Children) */}
                        <div className="pl-1">
                          {linkedChildren.length === 0 ? (
                            <div className="flex items-center gap-2 pl-4 py-1">
                              <span className="w-1.5 h-5 border-l-2 border-b-2 border-dashed border-zinc-200 rounded-bl shrink-0 -mt-3"></span>
                              <span className="text-[11px] text-zinc-400 italic">Chưa liên kết biểu mẫu con</span>
                            </div>
                          ) : (
                            <div className="space-y-0">
                              {linkedChildren.map((child, index) => {
                                const isLast = index === linkedChildren.length - 1;
                                return (
                                  <div key={child.id} className="flex items-start gap-2 pl-4 min-h-[32px]">
                                    {/* Connector Line */}
                                    <div className="flex shrink-0 relative w-4 h-full min-h-[32px]">
                                      <div className={`absolute left-0 top-0 w-0.5 border-l-2 border-zinc-200 ${isLast ? 'h-3.5' : 'h-full'}`}></div>
                                      <div className="absolute left-0 top-3.5 w-3.5 h-0.5 border-t-2 border-zinc-200"></div>
                                    </div>
                                    <div className="flex items-center gap-1.5 pt-1.5 max-w-full min-w-0">
                                      <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 shrink-0"></span>
                                      <span className="text-xs text-zinc-700 font-semibold truncate bg-white px-2 py-0.5 rounded border border-zinc-200">{child.name}</span>
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          )}
                        </div>
                      </div>
                    </div>

                    {/* Linked templates section */}
                    <div className="space-y-3">
                      <h4 className="text-xs font-bold text-zinc-700 uppercase tracking-wider flex items-center gap-1.5 font-display">
                        <svg className="w-4 h-4 text-zinc-500 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                          <path strokeLinecap="round" strokeLinejoin="round" d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z" />
                        </svg>
                        Biểu mẫu con đang liên kết ({linkedChildren.length})
                      </h4>
                      
                      {linkedChildren.length === 0 ? (
                        <div className="text-xs text-zinc-400 text-center py-5 border border-dashed border-zinc-200 bg-zinc-50 rounded-lg font-sans">
                          Chưa có biểu mẫu con nào được liên kết với file gốc này.
                        </div>
                      ) : (
                        <div className="space-y-2 max-h-[150px] overflow-y-auto pr-1">
                          {linkedChildren.map(child => (
                            <div key={child.id} className="flex flex-col sm:flex-row sm:items-center justify-between bg-zinc-50 border border-zinc-200 px-3 py-2 rounded-lg gap-2">
                              <div className="flex flex-col min-w-0 flex-1">
                                <span className="text-xs text-zinc-700 font-bold truncate pr-3" title={child.name}>{child.name}</span>
                                <label className="inline-flex items-center gap-1.5 mt-1 cursor-pointer select-none">
                                  <input
                                    type="checkbox"
                                    checked={!!child.is_repeated}
                                    onChange={(e) => handleToggleRepeated(child.id, e.target.checked)}
                                    className="w-3.5 h-3.5 text-emerald-600 bg-white border-zinc-300 rounded accent-emerald-600 cursor-pointer"
                                  />
                                  <span className="text-[10px] text-zinc-500 font-medium">Chế độ lặp (Nhập nhiều căn nhà/tài sản)</span>
                                </label>
                              </div>
                              <button
                                onClick={() => handleUnlinkChild(child.id)}
                                className="bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-[10px] font-bold px-2.5 py-1.5 rounded-lg transition-colors cursor-pointer self-end sm:self-center"
                              >
                                Hủy liên kết
                              </button>
                            </div>
                          ))}

                        </div>
                      )}
                    </div>
                    
                    {/* Available to link section */}
                    <div className="space-y-3 pt-5 border-t border-zinc-150">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <h4 className="text-xs font-bold text-zinc-700 uppercase tracking-wider flex items-center gap-1.5 font-display">
                          <svg className="w-4 h-4 text-zinc-500 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                            <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                          </svg>
                          Thêm biểu mẫu liên kết sẵn có ({filteredAvailable.length})
                        </h4>
                        {/* Search input box (Task 17) */}
                        <input
                          type="text"
                          placeholder="Tìm kiếm biểu mẫu..."
                          value={linkSearch}
                          onChange={(e) => setLinkSearch(e.target.value)}
                          className="bg-white border border-zinc-300 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20 rounded-lg px-2.5 py-1 text-xs text-zinc-950 focus:outline-none w-full sm:w-48 shadow-sm transition-all"
                        />
                      </div>
                      
                      {filteredAvailable.length === 0 ? (
                        <div className="text-xs text-zinc-400 text-center py-4 bg-zinc-50 border border-zinc-200 rounded-lg">
                          Không có biểu mẫu nào hợp lệ để liên kết.
                        </div>
                      ) : (
                        <div className="space-y-2 max-h-[180px] overflow-y-auto pr-1">
                          {filteredAvailable.map(avail => (
                            <div key={avail.id} className="flex justify-between items-center border border-zinc-200 px-3 py-2 rounded-lg hover:border-zinc-350 bg-white shadow-sm transition-colors">
                              <div className="flex flex-col min-w-0 pr-3">
                                <span className="text-xs text-zinc-700 font-semibold truncate">{avail.name}</span>
                                <span className="text-[10px] text-zinc-400 mt-0.5">{avail.fields_count} biến</span>
                              </div>
                              <button
                                onClick={() => handleLinkChild(avail.id)}
                                className="bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] font-bold px-3 py-1.5 rounded-lg transition-colors shadow-sm cursor-pointer border border-emerald-500 flex items-center gap-1"
                              >
                                <span>+</span> Liên kết
                              </button>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>
                  
                  {/* Footer */}
                  <div className="bg-zinc-50 px-6 py-3.5 border-t border-zinc-200 flex justify-end">
                    <button
                      onClick={() => { setShowLinkModal(false); setLinkingTemplate(null); setLinkSearch(''); }}
                      className="bg-zinc-900 hover:bg-zinc-800 text-white text-xs font-semibold py-2 px-4 rounded-lg transition-all cursor-pointer shadow-sm"
                    >
                      Đóng
                    </button>
                  </div>
                </div>
              </div>
            );
          })()}

      {/* Field Configurator View */}
      {activeView === 'config' && selectedTemplate && (
        <div className="w-full flex-1 flex flex-col pl-4 lg:pl-16 xl:pl-24 pr-0 animate-fade-up font-sans">
          <div ref={splitContainerRef} className={`flex ${isMobile ? 'flex-col pr-4' : 'flex-row'} gap-4 items-stretch relative min-h-[600px] w-full`}>
            {/* Cấu hình các trường bên trái */}
            <div
              style={{ width: isMobile ? '100%' : `${leftWidth}%` }}
              className="glass-panel p-6 space-y-6 h-fit flex flex-col bg-white"
            >
            <div className="flex flex-col gap-4 pb-4 border-b border-slate-100">
              <div>
                <span className="text-[11px] text-emerald-600 font-bold uppercase tracking-wider font-display">Chi tiết biểu mẫu</span>
                <h2 className="text-base font-bold text-slate-800 mt-1 line-clamp-1 font-display">{selectedTemplate.name}</h2>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 bg-slate-50 border border-slate-200 rounded-xl p-3">
                <div className="space-y-1">
                  <label className="text-[10px] font-bold uppercase tracking-wider text-slate-500">Danh mục thư viện</label>
                  <select
                    value={selectedTemplate.category_id || ''}
                    onChange={(e) => {
                      handleUpdateTemplateCategory(selectedTemplate.id, e.target.value);
                      setSelectedTemplate(prev => prev ? { ...prev, category_id: e.target.value || null } : prev);
                    }}
                    className="w-full px-3 py-2 text-xs bg-white border border-slate-200 rounded-lg focus:outline-none focus:border-emerald-500 cursor-pointer text-slate-700 font-semibold"
                  >
                    <option value="">Chưa phân loại</option>
                    {getFlattenedCategoryOptions().map(category => (
                      <option key={category.id} value={category.id}>
                        {category.label}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] font-bold uppercase tracking-wider text-slate-500">Quan hệ biểu mẫu</label>
                  <button
                    type="button"
                    onClick={() => handleOpenLinkModal(selectedTemplate)}
                    disabled={!!selectedTemplate.parent_template_id}
                    className={`w-full px-3 py-2 text-xs rounded-lg border transition-all cursor-pointer font-semibold ${
                      selectedTemplate.parent_template_id
                        ? 'bg-zinc-100 text-zinc-400 border-zinc-200 cursor-not-allowed'
                        : 'bg-white hover:bg-emerald-50 text-slate-700 border-slate-200 hover:border-emerald-300'
                    }`}
                  >
                    {selectedTemplate.parent_template_id ? 'Biểu mẫu con không quản lý file liên kết' : `Liên kết file con (${selectedTemplate.children_count || 0})`}
                  </button>
                </div>
              </div>
              <div className="flex flex-wrap gap-2">
                <button
                  onClick={() => {
                    setSidebarActiveMenu('templates');
                    setActiveView('dashboard');
                  }}
                  className="btn-premium bg-white hover:bg-slate-50 text-slate-700 font-semibold py-1.5 px-3 text-xs border border-slate-200 hover:border-slate-300 transition-colors cursor-pointer"
                >
                  Quay lại thư viện
                </button>
                <button
                  onClick={handleSaveConfig}
                  className="btn-premium bg-emerald-600 hover:bg-emerald-700 text-white font-semibold py-1.5 px-4 text-xs transition-all shadow-sm hover:shadow-md cursor-pointer"
                >
                  Lưu & Kích hoạt
                </button>
                
                <button
                  onClick={handleUndo}
                  disabled={historyIndex <= 0}
                  className={`btn-premium flex items-center gap-1.5 py-1.5 px-3 text-xs font-semibold border transition-colors cursor-pointer ${
                    historyIndex > 0
                      ? 'bg-white hover:bg-slate-50 border-slate-200 text-slate-700'
                      : 'bg-zinc-50 border-zinc-150 text-zinc-300 cursor-not-allowed opacity-50'
                  }`}
                  title="Hoàn tác thay đổi"
                >
                  <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M3 10h10a8 8 0 018 8v2M3 10l6 6m-6-6l6-6" />
                  </svg>
                  Hoàn tác
                </button>

                <button
                  onClick={handleExportConfig}
                  className="btn-premium bg-white hover:bg-slate-50 text-slate-700 font-semibold py-1.5 px-3 text-xs border border-slate-200 hover:border-slate-300 transition-colors cursor-pointer flex items-center gap-1.5"
                  title="Xuất cấu hình sang file JSON"
                >
                  <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
                  </svg>
                  Xuất
                </button>

                <label
                  className="btn-premium bg-white hover:bg-slate-50 text-slate-700 font-semibold py-1.5 px-3 text-xs border border-slate-200 hover:border-slate-300 transition-colors cursor-pointer flex items-center gap-1.5"
                  title="Nhập cấu hình từ file JSON"
                >
                  <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                  </svg>
                  Nhập
                  <input
                    type="file"
                    accept=".json"
                    onChange={handleImportConfig}
                    className="hidden"
                  />
                </label>

                {hasUnsavedManual && (
                  <span className="flex items-center text-amber-600 bg-amber-50 border border-amber-200 text-[10px] font-bold px-2 py-0.5 rounded-lg animate-pulse shrink-0">
                    ⚠️ Có thay đổi chưa lưu
                  </span>
                )}
              </div>
            </div>

            <p className="text-xs text-zinc-400">
              Thiết lập Nhãn, Kiểu và thứ tự cho các biến của biểu mẫu.
            </p>

            <div className="bg-emerald-50/50 border border-emerald-100 p-3 rounded-lg text-[11px] text-emerald-800 flex items-start gap-1.5 leading-relaxed">
              <span className="text-sm">💡</span>
              <span><strong>Mẹo:</strong> Bạn có thể bôi đen chữ trên bản xem trước Word ở bên phải để biến nó thành biến động ngay lập tức!</span>
            </div>

            {/* Hộp thoại bóc tách biến nhanh khi bôi đen */}
            {selectedText && (
              <div className="bg-zinc-50 border border-zinc-200 p-4 rounded-lg space-y-3 shadow-sm">
                <div className="flex justify-between items-center">
                  <h3 className="text-xs font-bold text-zinc-700 uppercase tracking-wider font-display">✨ Tạo biến nhanh từ chữ chọn</h3>
                  <button 
                    onClick={() => {
                      setSelectedText('');
                      window.getSelection()?.removeAllRanges();
                    }}
                    className="text-zinc-400 hover:text-zinc-600 text-xs font-bold cursor-pointer"
                  >
                    Đóng
                  </button>
                </div>
                
                <div className="bg-white p-2 rounded-lg border border-zinc-200 font-medium text-xs text-zinc-700 line-clamp-2 italic">
                  Đang chọn: <span className="text-zinc-900 font-bold">"{selectedText}"</span>
                </div>
                
                <form onSubmit={handleAddQuickField} className="space-y-2">
                  <div>
                    <label className="text-[9px] text-zinc-500 font-semibold uppercase tracking-wider block mb-1">Mã biến (Không dấu/khoảng trắng)</label>
                    <input
                      type="text"
                      required
                      value={quickKey}
                      onChange={(e) => setQuickKey(e.target.value)}
                      placeholder="Ví dụ: ho_ten_chu_dat"
                      className="bg-white border border-zinc-300 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20 rounded-lg px-2.5 py-1.5 w-full text-xs text-zinc-900 focus:outline-none"
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="text-[9px] text-zinc-500 font-semibold uppercase tracking-wider block mb-1">Nhãn hiển thị</label>
                      <input
                        type="text"
                        required
                        value={quickLabel}
                        onChange={(e) => setQuickLabel(e.target.value)}
                        placeholder="Ví dụ: Họ tên chủ đất"
                        className="bg-white border border-zinc-300 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20 rounded-lg px-2.5 py-1.5 w-full text-xs text-zinc-900 focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="text-[9px] text-zinc-500 font-semibold uppercase tracking-wider block mb-1">Kiểu dữ liệu</label>
                      <select
                        value={quickType}
                        onChange={(e) => setQuickType(e.target.value)}
                        className="bg-white border border-zinc-300 focus:border-emerald-500 rounded-lg px-2.5 py-1.5 w-full text-xs text-zinc-900 focus:outline-none"
                      >
                        <option value="text">Chữ (Text)</option>
                        <option value="date">Ngày (Date)</option>
                        <option value="number">Số (Number)</option>
                        <option value="boolean">Đúng/Sai</option>
                      </select>
                    </div>
                  </div>
                  
                  <button
                    type="submit"
                    className="w-full bg-zinc-900 hover:bg-zinc-800 text-white font-semibold py-2 rounded-lg text-xs transition-all shadow-sm flex items-center justify-center gap-1.5 cursor-pointer border border-zinc-800"
                  >
                    <span>+</span> Tạo biến động
                  </button>
                </form>
              </div>
            )}

            {/* Nút + Form thêm biến thủ công (không cần bôi đen) */}
            <div className="border-t border-zinc-100 pt-3">
              {!showManualAdd ? (
                <button
                  type="button"
                  onClick={handleOpenManualAdd}
                  className="w-full flex items-center justify-center gap-1.5 border border-dashed border-zinc-300 hover:border-emerald-400 hover:bg-emerald-50/50 text-zinc-500 hover:text-emerald-600 text-xs font-semibold py-2 rounded-lg transition-all cursor-pointer"
                >
                  <span className="text-base leading-none">＋</span> Thêm biến thủ công (không cần bôi đen)
                </button>
              ) : (
                <div className="bg-emerald-50/30 border border-emerald-100 p-4 rounded-lg space-y-3 shadow-sm text-zinc-700">
                  <div className="flex justify-between items-center">
                    <h3 className="text-xs font-bold text-emerald-800 uppercase tracking-wider font-display">✏️ Thêm biến thủ công</h3>
                    <button
                      type="button"
                      onClick={() => setShowManualAdd(false)}
                      className="text-zinc-400 hover:text-zinc-650 text-xs font-semibold cursor-pointer"
                    >
                      Đóng
                    </button>
                  </div>
                  <p className="text-[10px] text-emerald-700 leading-relaxed font-sans">
                    Dùng khi DOCX có dòng như <em>"Nơi cấp:"</em> nhưng không có <code className="bg-emerald-100 px-1.5 py-0.5 rounded text-emerald-800 font-mono text-[9px]">{'{{bien}}'}</code>.
                    Sau khi thêm, hãy gõ <code className="bg-emerald-100 px-1.5 py-0.5 rounded text-emerald-800 font-mono text-[9px]">{`{{ten_bien}}`}</code> vào file Word tương ứng.
                  </p>
                  <form onSubmit={handleAddManualField} className="space-y-2">
                    <div className="grid grid-cols-2 gap-2">
                      <div>
                        <label className="text-[9px] text-zinc-550 font-semibold uppercase tracking-wider block mb-1">Mã biến <span className="text-rose-500">*</span></label>
                        <input
                          type="text"
                          required
                          value={quickKey}
                          onChange={(e) => setQuickKey(e.target.value)}
                          placeholder="vd: noi_cap"
                          className="bg-white border border-zinc-300 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20 rounded-lg px-2.5 py-1.5 w-full text-xs text-zinc-900 focus:outline-none"
                        />
                      </div>
                      <div>
                        <label className="text-[9px] text-zinc-550 font-semibold uppercase tracking-wider block mb-1">Nhãn hiển thị <span className="text-rose-500">*</span></label>
                        <input
                          type="text"
                          required
                          value={quickLabel}
                          onChange={(e) => setQuickLabel(e.target.value)}
                          placeholder="vd: Nơi cấp"
                          className="bg-white border border-zinc-300 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20 rounded-lg px-2.5 py-1.5 w-full text-xs text-zinc-900 focus:outline-none"
                        />
                      </div>
                    </div>
                    <div>
                      <label className="text-[9px] text-zinc-550 font-semibold uppercase tracking-wider block mb-1">Kiểu dữ liệu</label>
                      <select
                        value={quickType}
                        onChange={(e) => setQuickType(e.target.value)}
                        className="bg-white border border-zinc-300 focus:border-emerald-500 rounded-lg px-2.5 py-1.5 w-full text-xs text-zinc-900 focus:outline-none"
                      >
                        <option value="text">Chữ (Text)</option>
                        <option value="date">Ngày (Date)</option>
                        <option value="number">Số (Number)</option>
                        <option value="boolean">Đúng/Sai</option>
                      </select>
                    </div>
                    <button
                      type="submit"
                      className="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-semibold py-2 rounded-lg text-xs transition-all shadow-sm flex items-center justify-center gap-1.5 cursor-pointer border border-emerald-500"
                    >
                      <span>＋</span> Tạo biến thủ công
                    </button>
                  </form>
                </div>
              )}
            </div>

            {fields.length === 0 ? (
              <div className="text-center py-8 text-zinc-400 text-xs">
                Không quét được biến động nào dạng `{"{{ten_bien}}"}` từ file. Vui lòng kiểm tra lại file Word.
              </div>
            ) : (
              <div className="space-y-4 max-h-[500px] overflow-y-auto pr-1">
                {fields.map((field, idx) => (
                  <div id={`field-config-${field.key_name}`} key={field.id} className="bg-zinc-50/50 p-4 rounded-lg border border-zinc-200 space-y-3 hover:border-zinc-300 transition-all duration-300">
                    <div className="flex justify-between items-center">
                      <span className="font-mono text-xs font-bold text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded-lg border border-emerald-150">{`{{${field.key_name}}}`}</span>
                      <div className="flex items-center gap-1.5">
                        <input
                          type="checkbox"
                          id={`req-${field.id}`}
                          checked={!!field.is_required}
                          onChange={(e) => handleFieldChange(idx, 'is_required', e.target.checked)}
                          className="w-3.5 h-3.5 text-emerald-600 bg-white border-zinc-300 rounded accent-emerald-600 cursor-pointer"
                        />
                        <label htmlFor={`req-${field.id}`} className="text-[10px] font-semibold text-zinc-500 cursor-pointer">Bắt buộc</label>
                      </div>
                    </div>
                    
                    <div className="grid grid-cols-2 gap-2">
                      <div>
                        <label className="text-[10px] text-zinc-400 font-semibold uppercase tracking-wider block mb-1">Nhãn hiển thị</label>
                        <input
                          type="text"
                          value={field.label}
                          onChange={(e) => handleFieldChange(idx, 'label', e.target.value)}
                          className="bg-white border border-zinc-300 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20 rounded-lg px-2.5 py-1.5 w-full text-xs text-zinc-900 focus:outline-none"
                        />
                      </div>
                      <div>
                        <label className="text-[10px] text-zinc-400 font-semibold uppercase tracking-wider block mb-1">Kiểu dữ liệu</label>
                        <select
                          value={field.field_type}
                          onChange={(e) => handleFieldChange(idx, 'field_type', e.target.value)}
                          className="bg-white border border-zinc-300 focus:border-emerald-500 rounded-lg px-2 py-1.5 w-full text-xs text-zinc-900 focus:outline-none"
                        >
                          <option value="text">Chữ (Text)</option>
                          <option value="date">Ngày (Date)</option>
                          <option value="number">Số (Number)</option>
                          <option value="boolean">Đúng/Sai</option>
                        </select>
                      </div>
                    </div>

                    {parentFields && parentFields.length > 0 && (() => {
                      const conflicts = getMappingConflicts();
                      const isConflict = field.parent_field_key && conflicts.has(field.parent_field_key);
                      const isMissingParent = field.parent_field_key && !parentFields.some(pf => pf.key_name === field.parent_field_key);
                      const sameNameInParent = parentFields.some(pf => pf.key_name === field.key_name);
                      const isAutoOverride = field.parent_field_key && sameNameInParent && field.parent_field_key !== field.key_name;
                      return (
                      <div className={`pt-2 border-t border-dashed ${isMissingParent ? 'border-rose-300' : isConflict || isAutoOverride ? 'border-amber-300' : 'border-zinc-200'}`}>
                        <label className="text-[10px] text-emerald-700 font-semibold uppercase tracking-wider block mb-1 flex items-center gap-1">
                          <svg className="w-3 h-3 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2.5">
                            <path strokeLinecap="round" strokeLinejoin="round" d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101m-.758-4.899a4 4 0 005.656 0l4-4a4 4 0 00-5.656-5.656l-1.1 1.1" />
                          </svg>
                          Đồng bộ từ trường File Gốc
                          {sameNameInParent && !field.parent_field_key && (
                            <span className="ml-auto text-[9px] font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-1.5 py-0.5 rounded">
                              tự đồng bộ theo tên
                            </span>
                          )}
                        </label>
                        <select
                          value={field.parent_field_key || ''}
                          onChange={(e) => handleFieldChange(idx, 'parent_field_key', e.target.value || null)}
                          className={`bg-white border rounded-xl px-2 py-1.5 w-full text-xs text-slate-900 focus:outline-none font-sans transition-colors ${
                            isMissingParent
                              ? 'border-rose-400 focus:border-rose-500 ring-2 ring-rose-100 text-rose-700'
                              : isConflict
                              ? 'border-amber-400 focus:border-amber-500 ring-2 ring-amber-100'
                              : 'border-emerald-200 focus:border-emerald-500'
                          }`}
                        >
                          <option value="">-- Tự điền độc lập (Không đồng bộ) --</option>
                          {parentFields.map(pf => (
                            <option key={pf.id} value={pf.key_name}>
                              {pf.label} ({`{{${pf.key_name}}}`})
                            </option>
                          ))}
                          {isMissingParent && (
                            <option value={field.parent_field_key} disabled>
                              ❌ [Lỗi] Thiếu trường cha: {`{{${field.parent_field_key}}}`}
                            </option>
                          )}
                        </select>
                        {isMissingParent && (
                          <div className="mt-1.5 text-[10px] text-rose-700 bg-rose-50 border border-rose-200 rounded-lg px-2 py-1 flex items-start gap-1.5 animate-pulse">
                            <span>⚠️</span>
                            <span>
                              Lỗi: Trường cha <code className="font-mono">{`{{${field.parent_field_key}}}`}</code> đã bị xóa hoặc không tồn tại ở biểu mẫu gốc.
                            </span>
                          </div>
                        )}
                        {isConflict && !isMissingParent && (
                          <div className="mt-1.5 text-[10px] text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-2 py-1 flex items-start gap-1.5">
                            <span>⚠️</span>
                            <span>
                              Có <strong>{conflicts.get(field.parent_field_key).length}</strong> biến con đang cùng ánh xạ vào{' '}
                              <code className="font-mono">{`{{${field.parent_field_key}}}`}</code>. Các biến này sẽ nhận cùng một giá trị.
                            </span>
                          </div>
                        )}
                        {isAutoOverride && !isConflict && (
                          <div className="mt-1.5 text-[10px] text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-2 py-1 flex items-start gap-1.5">
                            <span>⚠️</span>
                            <span>
                              Biến này đã có biến cùng tên trong file gốc nhưng bạn đang ánh xạ thủ công sang biến khác.
                            </span>
                          </div>
                        )}
                      </div>
                      );
                    })()}

                    {/* Văn bản thay thế gốc & các nút thao tác nâng cao */}
                    <div className="space-y-2 pt-2 border-t border-zinc-200">
                      <div className="flex justify-between items-center gap-1">
                        {field.replace_text ? (
                          <label className="text-[10px] text-zinc-400 font-semibold uppercase tracking-wider">Chữ gốc thay thế</label>
                        ) : (
                          <span className="text-[10px] text-zinc-400 italic">Biến được tạo thủ công</span>
                        )}
                        <div className="flex gap-1.5">
                          {field.replace_text && (
                            <button
                              type="button"
                              onClick={() => handleRestoreField(field.id, idx)}
                              className="text-[8px] text-zinc-600 hover:text-zinc-900 transition-colors font-bold uppercase tracking-wider bg-zinc-100 hover:bg-zinc-200 border border-zinc-250 px-2 py-0.5 rounded-lg cursor-pointer font-sans"
                              title="Khôi phục chữ gốc này vào tệp Word và xóa biến"
                            >
                              Khôi phục gốc
                            </button>
                          )}
                          <button
                            type="button"
                            onClick={() => handleDeleteField(idx)}
                            className="text-[8px] text-rose-600 hover:text-rose-800 transition-colors font-bold uppercase tracking-wider bg-rose-50 border border-rose-100 px-2 py-0.5 rounded-lg cursor-pointer font-sans"
                            title="Chỉ xóa biến khỏi danh sách cấu hình"
                          >
                            Xóa cấu hình
                          </button>
                        </div>
                      </div>
                      {field.replace_text && (
                        <>
                          <input
                            type="text"
                            value={field.replace_text || ''}
                            onChange={(e) => handleFieldChange(idx, 'replace_text', e.target.value)}
                            className="bg-white border border-zinc-300 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20 rounded-lg px-2.5 py-1.5 w-full text-xs text-zinc-900 focus:outline-none"
                            placeholder="Chữ gốc..."
                          />
                          {field.paragraph_context && (
                            <div className="text-[9px] text-zinc-400 leading-relaxed italic bg-white p-2.5 rounded-lg border border-zinc-200 line-clamp-2">
                              <strong>Ngữ cảnh:</strong> "{field.paragraph_context}"
                            </div>
                          )}
                        </>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            )}

            {/* ---------- Table Detection Panel ---------- */}
            {isLoadingTables && (
              <div className="flex items-center gap-2 text-xs text-zinc-400 mt-4 py-3">
                <div className="w-4 h-4 border-2 border-zinc-300 border-t-zinc-650 rounded-full animate-spin"></div>
                Đang quét bảng trong tài liệu...
              </div>
            )}

            {!isLoadingTables && detectedTables.length > 0 && (
              <div className="space-y-3 mt-6 pt-5 border-t border-zinc-200">
                <h3 className="text-xs font-bold text-zinc-700 uppercase tracking-wider flex items-center gap-2 font-display">
                  <svg className="w-4 h-4 text-emerald-500 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M3 10h18M3 14h18m-9-4v8m-7 0h14a2 2 0 002-2V8a2 2 0 00-2-2H5a2 2 0 00-2 2v8a2 2 0 002 2z" />
                  </svg>
                  Bảng phát hiện trong tài liệu
                  <span className="text-[9px] font-semibold text-white bg-emerald-500 px-1.5 py-0.5 rounded-md leading-none">{detectedTables.length}</span>
                </h3>

                {detectedTables.map(table => (
                  <div id={`table-accordion-${table.tableIndex}`} key={table.tableIndex} className="border border-zinc-200 rounded-lg overflow-hidden bg-white shadow-sm transition-all duration-300">
                    {/* Table accordion header */}
                    <button
                      type="button"
                      onClick={() => handleOpenTablePanel(table)}
                      className="w-full flex items-center justify-between px-3 py-2.5 bg-zinc-50 hover:bg-zinc-100 transition-colors text-left cursor-pointer"
                    >
                      <div className="flex items-center gap-2 min-w-0">
                        <span className="text-xs font-bold text-zinc-700 shrink-0">Bảng #{table.tableIndex + 1}</span>
                        <span className="text-[10px] text-zinc-400 truncate">{table.headers.length} cột · {table.rows.length} dòng</span>
                      </div>
                      <span className={`text-zinc-400 text-[9px] transition-transform duration-200 shrink-0 ${activeTableIndex === table.tableIndex ? 'rotate-180' : ''}`}>▼</span>
                    </button>

                    {/* Column header chips */}
                    <div className="px-3 py-2 flex flex-wrap gap-1 border-t border-zinc-100 bg-white">
                      {table.headers.map((h, i) => (
                        <span key={i} className="text-[9px] bg-zinc-100 text-zinc-700 border border-zinc-200 px-1.5 py-0.5 rounded-md font-medium">{h}</span>
                      ))}
                    </div>

                    {/* Expanded config panel */}
                    {activeTableIndex === table.tableIndex && (
                      <div className="px-3 py-3 border-t border-zinc-200 bg-emerald-50/10 space-y-3">
                        <div className="text-[9px] text-zinc-400 font-semibold uppercase tracking-wider">Cấu hình cột → biến</div>

                        {/* Column-to-field mapping */}
                        <div className="space-y-2">
                          {tableFieldEdits.map(f => (
                            <div key={f.colIndex} className="grid grid-cols-[1fr_1.2fr_0.8fr] gap-1.5 items-center">
                              <div className="text-[10px] text-zinc-600 font-medium truncate px-1" title={f.header}>
                                {f.header}
                              </div>
                              <input
                                type="text"
                                value={f.key_name}
                                onChange={e => handleTableFieldEdit(f.colIndex, 'key_name', e.target.value)}
                                className="bg-white border border-zinc-300 focus:border-emerald-500 rounded-lg px-2 py-1 text-[11px] text-zinc-900 focus:outline-none"
                                placeholder="key_name"
                              />
                               <select
                                 value={f.field_type}
                                 onChange={e => handleTableFieldEdit(f.colIndex, 'field_type', e.target.value)}
                                 className="bg-white border border-zinc-300 focus:border-emerald-500 rounded-lg px-1.5 py-1.5 text-[11px] text-zinc-900 focus:outline-none"
                               >
                                 <option value="text">Chữ</option>
                                 <option value="number">Số</option>
                                 <option value="date">Ngày</option>
                               </select>
                            </div>
                          ))}
                        </div>

                        {/* Row Checkboxes Selector */}
                        <div className="pt-2.5 border-t border-zinc-200/60 space-y-2 font-sans">
                           <div className="flex justify-between items-center">
                             <label className="text-[10px] text-zinc-500 font-semibold uppercase tracking-wider block font-display">Chọn dòng áp dụng:</label>
                             <button
                               type="button"
                               onClick={() => {
                                 if (selectedTableRows.length === table.rows.length) {
                                   setSelectedTableRows([]);
                                 } else {
                                   setSelectedTableRows(table.rows.map(r => r.rowIndex));
                                 }
                               }}
                               className="text-[9px] text-emerald-600 hover:text-emerald-800 font-bold uppercase tracking-wider select-none btn-premium cursor-pointer"
                             >
                               {selectedTableRows.length === table.rows.length ? 'Bỏ chọn hết' : 'Chọn tất cả'}
                             </button>
                           </div>
                           <div className="flex flex-wrap gap-1.5 py-1">
                             {table.rows.map(r => {
                               const isChecked = selectedTableRows.includes(r.rowIndex);
                               return (
                                 <button
                                   key={r.rowIndex}
                                   type="button"
                                   onClick={() => {
                                     if (isChecked) {
                                       setSelectedTableRows(prev => prev.filter(idx => idx !== r.rowIndex));
                                     } else {
                                       setSelectedTableRows(prev => [...prev, r.rowIndex]);
                                     }
                                   }}
                                   className={`px-2.5 py-1 rounded-lg text-[10px] font-bold border transition-all btn-premium select-none cursor-pointer ${
                                     isChecked 
                                       ? 'bg-emerald-50 text-emerald-800 border-emerald-250 shadow-sm ring-1 ring-emerald-200/30 scale-[1.02]' 
                                       : 'bg-white text-zinc-600 border-zinc-250 hover:border-zinc-350'
                                   }`}
                                 >
                                   {isChecked ? '✓ ' : ''}Dòng {r.rowIndex}
                                 </button>
                               );
                             })}
                           </div>
                         </div>

                        {/* Inject button */}
                        <button
                           type="button"
                           onClick={handleInjectTable}
                           disabled={isInjectingTable}
                           className="btn-premium w-full bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2 rounded-lg text-xs transition-all shadow-sm flex items-center justify-center gap-1.5 disabled:opacity-50 cursor-pointer border border-emerald-500"
                         >
                           {isInjectingTable ? (
                             <>
                               <div className="w-3.5 h-3.5 border-2 border-white/30 border-t-white rounded-full animate-spin"></div>
                               Đang xử lý...
                             </>
                           ) : (
                             '⚡ Gài placeholder vào bảng'
                           )}
                         </button>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )}
          </div>

            {/* Thanh kéo giãn (Resizer Handle) */}
            {!isMobile && (
              <div 
                onMouseDown={startResizing}
                className="w-2 hover:w-3 bg-zinc-200/80 hover:bg-emerald-600 cursor-col-resize self-stretch transition-all rounded-md flex items-center justify-center select-none shadow-inner group mx-1"
                title="Kéo chuột sang trái/phải để thay đổi kích thước"
              >
                <div className="w-0.5 h-8 bg-zinc-400 group-hover:bg-white rounded-full transition-colors" />
              </div>
            )}

            {/* Bản xem trước tài liệu bên phải */}
            <div 
              style={{ width: isMobile ? '100%' : `${100 - leftWidth}%` }} 
              className="flex flex-col flex-1"
            >
              <DocxPreview 
                fileUrl={`http://localhost:5000/api/templates/${selectedTemplate.id}/download-original?t=${previewKey}`} 
                title={selectedTemplate.name}
                fields={fields}
                onTableRowClick={handleTableRowClick}
              />
            </div>
          </div>
        </div>
      )}

      {/* Dynamic Form Client View */}
      {activeView === 'fill' && selectedTemplate && (
        <div className="w-full flex-1 flex flex-col pl-4 lg:pl-16 xl:pl-24 pr-4 gap-4 animate-fade-up font-sans">
          {/* Mobile Tab Switcher */}
          {isMobile && (
            <div className="flex border border-slate-200 bg-white p-1.5 rounded-2xl shadow-sm">
              <button
                type="button"
                onClick={() => setMobileFillTab('form')}
                className={`flex-1 text-center py-2.5 text-xs font-bold uppercase tracking-wider rounded-xl transition-all ${mobileFillTab === 'form' ? 'bg-slate-900 text-white shadow' : 'text-slate-500 hover:text-slate-800'}`}
              >
                📝 Điền hồ sơ
              </button>
              <button
                type="button"
                onClick={() => setMobileFillTab('preview')}
                className={`flex-1 text-center py-2.5 text-xs font-bold uppercase tracking-wider rounded-xl transition-all ${mobileFillTab === 'preview' ? 'bg-slate-900 text-white shadow' : 'text-slate-500 hover:text-slate-800'}`}
              >
                👁 Xem tài liệu
              </button>
            </div>
          )}

          <div ref={splitContainerRef} className={`flex ${isMobile ? 'flex-col' : 'flex-row'} gap-4 items-stretch relative min-h-[600px] w-full`}>
            {/* Form điền thông tin bên trái */}
            <div
              style={{ width: isMobile ? '100%' : `${leftWidth}%` }}
              className={`glass-panel p-6 md:p-8 space-y-6 h-fit flex-col bg-white ${isMobile && mobileFillTab !== 'form' ? 'hidden' : 'flex'}`}
            >
              <div className="text-center pb-5 border-b border-slate-100">
                <span className="text-[9px] text-emerald-600 font-extrabold uppercase tracking-widest font-display">HỒ SƠ KHÁCH HÀNG ĐIỀN TRỰC TUYẾN</span>
                <h2 className="text-sm font-bold text-zinc-800 mt-2 line-clamp-2 uppercase tracking-wide leading-relaxed font-display">{selectedTemplate.name}</h2>
                <p className="text-xs text-zinc-400 mt-2">Vui lòng nhập đầy đủ thông tin bên dưới để tự động tạo hợp đồng.</p>
              </div>

              <form onSubmit={handleFormSubmit} className="space-y-6">
              {/* Step indicator header */}
              <div className="flex items-center justify-between px-1 mb-2 select-none font-sans shrink-0">
                {[
                  { step: 1, label: 'Hợp đồng' },
                  { step: 2, label: 'Thuế TNCN' },
                  { step: 3, label: 'Cây trồng' },
                  { step: 4, label: 'Đất PNN' }
                ].map((item, index) => {
                  const stepActive = hasActiveFields(item.step);
                  return (
                    <div key={item.step} className="flex items-center flex-1 last:flex-none">
                      <div className="flex flex-col items-center gap-1 relative">
                        <button
                          type="button"
                          disabled={!stepActive}
                          onClick={() => stepActive && setCurrentStep(item.step)}
                          className={`w-6 h-6 rounded-full border flex items-center justify-center text-[10px] font-bold transition-all ${
                            !stepActive
                              ? 'bg-zinc-100 border-zinc-200 text-zinc-350 cursor-not-allowed opacity-50'
                              : currentStep === item.step
                              ? 'bg-emerald-600 border-emerald-600 text-white ring-4 ring-emerald-50 cursor-pointer'
                              : currentStep > item.step
                              ? 'bg-emerald-50 border-emerald-300 text-emerald-600 cursor-pointer'
                              : 'bg-white border-slate-200 text-slate-400 hover:border-slate-350 cursor-pointer'
                          }`}
                        >
                          {currentStep > item.step && stepActive ? '✓' : item.step}
                        </button>
                        <span className={`text-[9px] font-bold tracking-tight whitespace-nowrap ${
                          !stepActive
                            ? 'text-zinc-300'
                            : currentStep === item.step
                            ? 'text-emerald-700 font-extrabold'
                            : 'text-slate-400'
                        }`}>
                          {item.label}
                        </span>
                      </div>
                      {index < 3 && (
                        <div className={`h-0.5 flex-1 mx-1.5 -mt-3.5 transition-all duration-300 ${currentStep > item.step && hasActiveFields(item.step + 1) ? 'bg-emerald-400' : 'bg-slate-100'}`} />
                      )}
                    </div>
                  );
                })}
              </div>

                 {/* Tên khách hàng & Số điện thoại (Chỉ hiện ở bước 1) */}
                 {currentStep === 1 && (
                  <>
                    <div className="bg-zinc-50 p-5 rounded-lg border border-zinc-200 shadow-inner space-y-4">
                      <h3 className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider font-display">Thông tin liên hệ</h3>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div>
                          <label className="block text-[10px] font-bold text-zinc-500 uppercase tracking-wider mb-1.5 font-display">Họ và tên khách hàng</label>
                          <input
                            type="text"
                            required
                            value={customerName}
                            onChange={(e) => setCustomerName(e.target.value)}
                            placeholder="Ví dụ: Nguyễn Văn A"
                            className="w-full bg-white border border-zinc-300 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20 rounded-lg px-4 py-3 text-sm text-zinc-900 focus:outline-none transition-all shadow-sm"
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-zinc-500 uppercase tracking-wider mb-1.5 font-display">Số điện thoại liên hệ</label>
                          <input
                            type="text"
                            required
                            value={customerPhone}
                            onChange={(e) => setCustomerPhone(e.target.value)}
                            placeholder="Ví dụ: 0912345678"
                            className="w-full bg-white border border-zinc-300 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20 rounded-lg px-4 py-3 text-sm text-zinc-900 focus:outline-none transition-all shadow-sm"
                          />
                        </div>
                      </div>
                    </div>

                    {/* Lựa chọn tài liệu đi kèm (Chỉ hiện ở bước 1) */}
                    {linkedChildren && linkedChildren.length > 0 && (
                      <div className="bg-zinc-50 p-5 rounded-lg border border-zinc-200 shadow-inner space-y-4">
                        <div className="flex justify-between items-center">
                          <h3 className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider font-display">Các tài liệu trong bộ hồ sơ</h3>
                          <span className="text-[9px] font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded">
                            Đang chọn {1 + selectedChildIds.length}/{1 + linkedChildren.length}
                          </span>
                        </div>
                        <div className="grid grid-cols-1 gap-2.5">
                          {/* Master Document Card (Always Selected & Disabled) */}
                          <div className="flex items-center gap-3 bg-white p-3.5 rounded-xl border border-emerald-200 ring-1 ring-emerald-50 select-none opacity-90">
                            <input
                              type="checkbox"
                              checked={true}
                              disabled={true}
                              className="w-4 h-4 text-emerald-600 bg-emerald-50 border-emerald-300 rounded accent-emerald-600 cursor-not-allowed"
                            />
                            <div className="flex-1 min-w-0">
                              <span className="block text-xs font-bold text-zinc-800 truncate">👑 {selectedTemplate.name} (Tài liệu chính)</span>
                              <span className="block text-[9.5px] text-zinc-400 mt-0.5">Biểu mẫu bắt buộc phải có</span>
                            </div>
                          </div>

                          {/* Child Document Cards */}
                          {linkedChildren.map(child => {
                            const isChecked = selectedChildIds.includes(child.id);
                            return (
                              <div 
                                key={child.id} 
                                onClick={() => {
                                  if (isChecked) {
                                    setSelectedChildIds(prev => prev.filter(id => id !== child.id));
                                  } else {
                                    setSelectedChildIds(prev => [...prev, child.id]);
                                  }
                                }}
                                className={`flex items-center gap-3 p-3.5 rounded-xl border transition-all cursor-pointer duration-200 select-none ${
                                  isChecked 
                                    ? 'bg-white border-emerald-500 shadow-sm ring-1 ring-emerald-100 scale-[1.01]' 
                                    : 'bg-white border-zinc-200 hover:border-zinc-350 hover:scale-[1.005]'
                                }`}
                              >
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  readOnly
                                  className="w-4 h-4 text-emerald-600 bg-white border-zinc-350 rounded accent-emerald-600 cursor-pointer pointer-events-none"
                                />
                                <div className="flex-1 min-w-0">
                                  <span className="block text-xs font-bold text-zinc-800 truncate">{child.name}</span>
                                  <span className="block text-[9.5px] text-zinc-400 mt-0.5">
                                    {isChecked ? '✓ Đang bật (Nhấn để tắt)' : '✗ Đang tắt (Nhấn để bật)'}
                                  </span>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    )}
                  </>
                )}

                {/* Danh sách các trường động */}
                <div className="space-y-4">
                  <div className="space-y-6 max-h-[380px] overflow-y-auto pr-1">
                    {(() => {
                      const stepFields = fields.filter(f => 
                        getFieldStep(f) === currentStep &&
                        (!f.childTemplateId || selectedChildIds.includes(f.childTemplateId))
                      );
                      if (stepFields.length === 0) {
                        return (
                          <div className="text-center py-12 text-slate-400 text-xs bg-slate-50 border border-slate-100 rounded-xl font-sans">
                            💡 Tài liệu liên quan tới bước này đã bị bỏ chọn.
                            <br /> Bạn có thể nhấn "Tiếp theo" để tiếp tục.
                          </div>
                        );
                      }

                      return Object.entries(getGroupedFields(stepFields)).sort(([a], [b]) => a.localeCompare(b)).map(([groupName, groupFields]) => {
                        const firstField = groupFields[0];
                        const childTemp = firstField?.childTemplateId 
                          ? linkedChildren.find(c => c.id === firstField.childTemplateId) 
                          : null;
                        const isRepeated = childTemp?.is_repeated === 1;

                        if (isRepeated) {
                          const childId = childTemp.id;
                          const records = formData[childId] || [];

                          return (
                            <div key={groupName} className="space-y-4 pt-3 border-t first:border-t-0 border-zinc-200/60 font-sans">
                              <div className="flex justify-between items-center px-1 border-l-2 border-emerald-600 pl-2">
                                <h3 className="text-[10px] font-bold text-emerald-700 uppercase tracking-wider font-display">
                                  {groupName.startsWith('Biểu mẫu con:') ? groupName.substring(13).trim() : groupName} (Nhập nhiều bản ghi)
                                </h3>
                                <span className="text-[9px] font-bold text-emerald-800 bg-emerald-50 border border-emerald-250 px-2 py-0.5 rounded-lg select-none">
                                  {records.length} tài sản
                                </span>
                              </div>

                              <div className="space-y-4 pl-3">
                                {records.map((record, rIdx) => (
                                  <div key={record._id || rIdx} className="bg-zinc-50/50 p-4 rounded-xl border border-zinc-200 shadow-sm relative space-y-4">
                                    <div className="flex justify-between items-center border-b border-zinc-200/60 pb-2">
                                      <span className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider font-display">Tài sản / Căn nhà #{rIdx + 1}</span>
                                      {records.length > 1 && (
                                        <button
                                          type="button"
                                          onClick={() => handleRemoveRepeatedRecord(childId, rIdx)}
                                          className="text-[9px] text-rose-600 hover:text-rose-800 font-bold uppercase tracking-wider bg-rose-50 border border-rose-100 hover:bg-rose-100 transition-colors px-2 py-1 rounded-lg cursor-pointer"
                                        >
                                          Xóa bản ghi
                                        </button>
                                      )}
                                    </div>

                                    <div className="space-y-3.5">
                                      {groupFields.map((field) => (
                                        <div key={field.id} className="space-y-1.5">
                                          <label className="block text-[10px] font-bold text-zinc-500 uppercase tracking-wide font-display">
                                            {field.label}
                                            {!!field.is_required && <span className="text-rose-500 ml-1">*</span>}
                                          </label>
                                          
                                          {field.field_type === 'text' && (
                                            <input
                                              id={`field-input-${field.key_name}-${rIdx}`}
                                              type="text"
                                              required={!!field.is_required}
                                              value={record[field.key_name] || ''}
                                              onChange={(e) => handleUpdateRepeatedField(childId, rIdx, field.key_name, e.target.value)}
                                              className="w-full bg-white border border-zinc-300 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20 rounded-lg px-4 py-2.5 text-sm text-zinc-900 focus:outline-none transition-all shadow-sm"
                                            />
                                          )}

                                          {field.field_type === 'date' && (
                                            <input
                                              id={`field-input-${field.key_name}-${rIdx}`}
                                              type="date"
                                              required={!!field.is_required}
                                              value={record[field.key_name] || ''}
                                              onChange={(e) => handleUpdateRepeatedField(childId, rIdx, field.key_name, e.target.value)}
                                              className="w-full bg-white border border-zinc-300 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20 rounded-lg px-4 py-2.5 text-sm text-zinc-900 focus:outline-none transition-all shadow-sm"
                                            />
                                          )}

                                          {field.field_type === 'number' && (
                                            <input
                                              id={`field-input-${field.key_name}-${rIdx}`}
                                              type="number"
                                              required={!!field.is_required}
                                              value={record[field.key_name] || ''}
                                              onChange={(e) => handleUpdateRepeatedField(childId, rIdx, field.key_name, e.target.value)}
                                              className="w-full bg-white border border-zinc-300 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20 rounded-lg px-4 py-2.5 text-sm text-zinc-900 focus:outline-none transition-all shadow-sm"
                                            />
                                          )}

                                          {field.field_type === 'boolean' && (
                                            <div className="flex items-center gap-3 bg-white border border-zinc-200 px-4 py-2.5 rounded-lg shadow-sm">
                                              <input
                                                id={`field-input-${field.key_name}-${rIdx}`}
                                                type="checkbox"
                                                checked={!!record[field.key_name]}
                                                onChange={(e) => handleUpdateRepeatedField(childId, rIdx, field.key_name, e.target.checked)}
                                                className="w-4 h-4 text-emerald-600 bg-white border-zinc-300 rounded accent-emerald-600 focus:ring-emerald-500/40 cursor-pointer"
                                              />
                                              <label htmlFor={`field-input-${field.key_name}-${rIdx}`} className="text-xs text-zinc-650 font-semibold cursor-pointer select-none">
                                                Kích hoạt / Xác nhận tùy chọn này
                                              </label>
                                            </div>
                                          )}
                                        </div>
                                      ))}
                                    </div>
                                  </div>
                                ))}

                                <button
                                  type="button"
                                  onClick={() => handleAddRepeatedRecord(childId)}
                                  className="w-full flex items-center justify-center gap-1.5 border border-dashed border-emerald-350 hover:border-emerald-500 hover:bg-emerald-50 text-emerald-600 hover:text-emerald-700 text-xs font-semibold py-2.5 rounded-xl transition-all cursor-pointer shadow-sm select-none"
                                >
                                  ＋ Thêm căn nhà/Tài sản mới
                                </button>
                              </div>
                            </div>
                          );
                        }

                        return (
                          <div key={groupName} className="space-y-4 pt-3 border-t first:border-t-0 border-zinc-200/60 font-sans">
                            <h3 className="text-[10px] font-bold text-emerald-700 uppercase tracking-wider px-1 border-l-2 border-emerald-600 pl-2 font-display">
                              {groupName.startsWith('1.') || groupName.startsWith('2.') || groupName.startsWith('3.') || groupName.startsWith('4.') || groupName.startsWith('5.')
                                ? groupName.substring(3) 
                                : groupName.startsWith('Biểu mẫu con:')
                                ? groupName.substring(13).trim()
                                : groupName}
                            </h3>
                            
                            {groupFields.map((field) => (
                              <div key={field.id} className="space-y-1.5 pl-3">
                                <label className="block text-[10px] font-bold text-zinc-500 uppercase tracking-wide font-display">
                                  {field.label}
                                  {!!field.is_required && <span className="text-rose-500 ml-1">*</span>}
                                </label>
                                
                                {field.field_type === 'text' && (
                                  <input
                                    id={`field-input-${field.key_name}`}
                                    type="text"
                                    required={!!field.is_required}
                                    value={formData[field.key_name] || ''}
                                    onChange={(e) => setFormData({ ...formData, [field.key_name]: e.target.value })}
                                    className="w-full bg-white border border-zinc-300 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20 rounded-lg px-4 py-2.5 text-sm text-zinc-900 focus:outline-none transition-all shadow-sm"
                                  />
                                )}

                                {field.field_type === 'date' && (
                                  <input
                                    id={`field-input-${field.key_name}`}
                                    type="date"
                                    required={!!field.is_required}
                                    value={formData[field.key_name] || ''}
                                    onChange={(e) => setFormData({ ...formData, [field.key_name]: e.target.value })}
                                    className="w-full bg-white border border-zinc-300 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20 rounded-lg px-4 py-2.5 text-sm text-zinc-900 focus:outline-none transition-all shadow-sm"
                                  />
                                )}

                                {field.field_type === 'number' && (
                                  <input
                                    id={`field-input-${field.key_name}`}
                                    type="number"
                                    required={!!field.is_required}
                                    value={formData[field.key_name] || ''}
                                    onChange={(e) => setFormData({ ...formData, [field.key_name]: e.target.value })}
                                    className="w-full bg-white border border-zinc-300 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500/20 rounded-lg px-4 py-2.5 text-sm text-zinc-900 focus:outline-none transition-all shadow-sm"
                                  />
                                )}

                                {field.field_type === 'boolean' && (
                                  <div className="flex items-center gap-3 bg-zinc-50 border border-zinc-200 px-4 py-2.5 rounded-lg shadow-sm">
                                    <input
                                      id={`field-input-${field.key_name}`}
                                      type="checkbox"
                                      checked={!!formData[field.key_name]}
                                      onChange={(e) => setFormData({ ...formData, [field.key_name]: e.target.checked })}
                                      className="w-4 h-4 text-emerald-600 bg-white border-zinc-300 rounded accent-emerald-600 focus:ring-emerald-500/40 cursor-pointer"
                                    />
                                    <label htmlFor={`field-input-${field.key_name}`} className="text-xs text-zinc-650 font-semibold cursor-pointer select-none">
                                      Kích hoạt / Xác nhận tùy chọn này
                                    </label>
                                  </div>
                                )}
                              </div>
                            ))}
                          </div>
                        );
                      });

                    })()}
                  </div>
                </div>

                <div className="pt-4 border-t border-zinc-200 flex justify-between gap-3">
                  {currentStep > 1 ? (
                    <button
                      type="button"
                      onClick={() => {
                        let prev = currentStep - 1;
                        while (prev >= 1) {
                          const prevStepFields = fields.filter(f => 
                            getFieldStep(f) === prev && 
                            (!f.childTemplateId || selectedChildIds.includes(f.childTemplateId))
                          );
                          if (prevStepFields.length > 0 || prev === 1) {
                            break;
                          }
                          prev--;
                        }
                        setCurrentStep(Math.max(prev, 1));
                      }}
                      className="btn-premium px-4 py-2.5 text-xs font-semibold text-slate-700 bg-white hover:bg-slate-50 border border-slate-200 rounded-xl transition-all cursor-pointer flex items-center gap-1"
                    >
                      ← Quay lại
                    </button>
                  ) : (
                    <button
                      type="button"
                      onClick={() => {
                        if (window.confirm("Bạn có chắc muốn hủy hồ sơ và quay lại? Dữ liệu đã nhập sẽ bị mất.")) {
                          setActiveView('dashboard');
                        }
                      }}
                      className="btn-premium px-4 py-2.5 text-xs font-semibold text-rose-600 bg-white hover:bg-rose-50 border border-rose-200 rounded-xl transition-all cursor-pointer"
                    >
                      Hủy hồ sơ
                    </button>
                  )}

                  {currentStep < totalSteps ? (
                    <button
                      type="button"
                      onClick={() => {
                        let next = currentStep + 1;
                        while (next <= totalSteps) {
                          const nextStepFields = fields.filter(f => 
                            getFieldStep(f) === next && 
                            (!f.childTemplateId || selectedChildIds.includes(f.childTemplateId))
                          );
                          if (nextStepFields.length > 0 || next === totalSteps) {
                            break;
                          }
                          next++;
                        }
                        setCurrentStep(Math.min(next, totalSteps));
                      }}
                      className="btn-premium bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-5 py-2.5 text-xs rounded-xl transition-all shadow-sm flex items-center gap-1 cursor-pointer"
                    >
                      Tiếp theo →
                    </button>
                  ) : (
                    <button
                      type="submit"
                      disabled={isSubmitting}
                      className="btn-premium bg-zinc-900 hover:bg-zinc-800 text-white font-bold px-6 py-2.5 text-xs rounded-xl transition-all shadow-sm flex items-center justify-center gap-2 uppercase tracking-wide cursor-pointer disabled:opacity-50"
                    >
                      {isSubmitting ? (
                        <>
                          <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin"></div>
                          Đang tạo...
                        </>
                      ) : (
                        'Xác nhận & Xuất File'
                      )}
                    </button>
                  )}
                </div>
              </form>

              {/* Progress Modal Overlay */}
              {isSubmitting && selectedTemplate && (() => {
                const childCount = selectedTemplate.children_count || 0;
                const { current, total, message } = submissionProgress;
                const percentage = total > 0 ? Math.round((current / total) * 100) : 0;
                
                const getStepState = (step) => {
                  if (step === 1) { // Chuẩn bị
                    if (current > 1) return 'completed';
                    if (current === 1) return 'processing';
                    return 'pending';
                  }
                  if (step === 2) { // File gốc
                    if (current > 2) return 'completed';
                    if (current === 2) return 'processing';
                    return 'pending';
                  }
                  if (step === 3) { // File con
                    if (childCount === 0) return 'hidden';
                    if (current > 2 + childCount) return 'completed';
                    if (current >= 3 && current <= 2 + childCount) return 'processing';
                    return 'pending';
                  }
                  if (step === 4) { // Đóng gói
                    const finalizeStep = childCount > 0 ? 3 + childCount : 3;
                    if (current >= total && total > 0) return 'completed';
                    if (current >= finalizeStep) return 'processing';
                    return 'pending';
                  }
                  return 'pending';
                };

                const renderStepIcon = (state, num) => {
                  if (state === 'completed') {
                    return (
                      <span className="w-5 h-5 rounded-full bg-emerald-100 text-emerald-700 border border-emerald-300 flex items-center justify-center text-[10px] font-bold shrink-0">
                        ✓
                      </span>
                    );
                  }
                  if (state === 'processing') {
                    return (
                      <span className="w-5 h-5 rounded-full bg-emerald-50 text-emerald-600 border border-emerald-500 ring-2 ring-emerald-100 flex items-center justify-center text-[10px] font-bold shrink-0 animate-spin">
                        ⟳
                      </span>
                    );
                  }
                  return (
                    <span className="w-5 h-5 rounded-full bg-zinc-100 text-zinc-400 border border-zinc-200 flex items-center justify-center text-[10px] font-bold shrink-0">
                      {num}
                    </span>
                  );
                };

                const step1State = getStepState(1);
                const step2State = getStepState(2);
                const step3State = getStepState(3);
                const step4State = getStepState(4);

                return (
                  <div className="fixed inset-0 z-50 flex items-center justify-center bg-zinc-900/60 backdrop-blur-sm p-4">
                    <div className="bg-white border border-zinc-200 w-full max-w-sm rounded-lg shadow-2xl p-8 space-y-6 animate-fade-up">
                      <div className="text-center">
                        <h3 className="text-sm font-bold text-zinc-800 uppercase tracking-wider font-display mb-1">Đang xử lý hồ sơ</h3>
                        <p className="text-xs text-zinc-400">{message || 'Vui lòng đợi trong giây lát...'}</p>
                      </div>

                      {/* Progress Bar */}
                      <div className="space-y-2">
                        <div className="w-full h-2 bg-zinc-100 rounded-full overflow-hidden">
                          <div
                            className="h-full bg-emerald-600 transition-all duration-500 ease-out rounded-full"
                            style={{ width: `${percentage}%` }}
                          />
                        </div>
                        <div className="text-xs text-zinc-500 text-center font-semibold">{percentage}%</div>
                      </div>

                      {/* Step-by-step list */}
                      <div className="space-y-4">
                        {/* Step 1 */}
                        <div className={`flex items-start gap-3 transition-opacity duration-300 ${step1State === 'pending' ? 'opacity-50' : 'opacity-100'}`}>
                          {renderStepIcon(step1State, '1')}
                          <div>
                            <p className={`text-xs font-semibold ${step1State === 'processing' ? 'text-emerald-700' : 'text-zinc-700'}`}>Chuẩn bị dữ liệu</p>
                            <p className="text-[10px] text-zinc-400">Kiểm tra thông tin & biểu mẫu hợp lệ</p>
                          </div>
                        </div>

                        {/* Step 2 */}
                        <div className={`flex items-start gap-3 transition-opacity duration-300 ${step2State === 'pending' ? 'opacity-50' : 'opacity-100'}`}>
                          {renderStepIcon(step2State, '2')}
                          <div>
                            <p className={`text-xs font-semibold ${step2State === 'processing' ? 'text-emerald-700' : 'text-zinc-700'}`}>Điền dữ liệu file gốc</p>
                            <p className="text-[10px] text-zinc-400">{selectedTemplate.name}.docx</p>
                          </div>
                        </div>

                        {/* Step 3 (Only show if childCount > 0) */}
                        {step3State !== 'hidden' && (
                          <div className={`flex items-start gap-3 transition-opacity duration-300 ${step3State === 'pending' ? 'opacity-50' : 'opacity-100'}`}>
                            {renderStepIcon(step3State, '3')}
                            <div>
                              <p className={`text-xs font-semibold ${step3State === 'processing' ? 'text-emerald-700' : 'text-zinc-700'}`}>Điền dữ liệu {childCount} file con</p>
                              <p className="text-[10px] text-zinc-400">Các biểu mẫu con đã liên kết</p>
                            </div>
                          </div>
                        )}

                        {/* Step 4 */}
                        <div className={`flex items-start gap-3 transition-opacity duration-300 ${step4State === 'pending' ? 'opacity-50' : 'opacity-100'}`}>
                          {renderStepIcon(step4State, childCount > 0 ? '4' : '3')}
                          <div>
                            <p className={`text-xs font-semibold ${step4State === 'processing' ? 'text-emerald-700' : 'text-zinc-700'}`}>Đóng gói & Hoàn tất</p>
                            <p className="text-[10px] text-zinc-400">{childCount > 0 ? 'Tạo tệp lưu trữ .zip' : 'Tạo tệp Word .docx'}</p>
                          </div>
                        </div>
                      </div>

                      <p className="text-[9px] text-zinc-450 text-center italic">
                        Vui lòng không tắt trình duyệt hoặc làm mất kết nối mạng.
                      </p>
                    </div>
                  </div>
                );
              })()}
            </div>

            {/* Thanh kéo giãn (Resizer Handle) */}
            {!isMobile && (
              <div 
                onMouseDown={startResizing}
                className="w-2 hover:w-3 bg-zinc-200/80 hover:bg-emerald-600 cursor-col-resize self-stretch transition-all rounded-md flex items-center justify-center select-none shadow-inner group mx-1"
                title="Kéo chuột sang trái/phải để thay đổi kích thước"
              >
                <div className="w-0.5 h-8 bg-zinc-400 group-hover:bg-white rounded-full transition-colors" />
              </div>
            )}

            {/* Khung xem trước tài liệu Word cập nhật thời gian thực bên phải */}
            <div 
              style={{ width: isMobile ? '100%' : `${100 - leftWidth}%` }}
              className={`flex-col flex-1 ${isMobile && mobileFillTab !== 'preview' ? 'hidden' : 'flex'}`}
            >
              {linkedChildren && linkedChildren.length > 0 && (
                <div className="flex bg-zinc-900 border-x border-t border-zinc-800 p-1.5 rounded-t-lg gap-1.5 overflow-x-auto select-none shrink-0">
                  <button
                    type="button"
                    onClick={() => handleSwitchPreview(selectedTemplate.id)}
                    className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer truncate max-w-[200px] ${previewTemplateId === selectedTemplate.id ? 'bg-zinc-800 text-white border border-zinc-700' : 'text-zinc-400 hover:text-zinc-200'}`}
                    title={`${selectedTemplate.name} (File Gốc)`}
                  >
                    👑 {selectedTemplate.name}
                  </button>
                  {linkedChildren.filter(child => selectedChildIds.includes(child.id)).map(child => (
                    <button
                      key={child.id}
                      type="button"
                      onClick={() => handleSwitchPreview(child.id)}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer truncate max-w-[200px] ${previewTemplateId === child.id ? 'bg-zinc-800 text-white border border-zinc-700' : 'text-zinc-400 hover:text-zinc-200'}`}
                      title={`${child.name} (File Con)`}
                    >
                      📄 {child.name}
                    </button>
                  ))}
                </div>
              )}
              <DocxPreview 
                fileUrl={`http://localhost:5000/api/templates/${previewTemplateId || selectedTemplate.id}/download-original?t=${previewKey}`} 
                title={templates.find(t => t.id === previewTemplateId)?.name || selectedTemplate.name}
                liveData={getResolvedPreviewData()}
                fields={previewFields && previewFields.length > 0 ? previewFields : fields}
              />
            </div>
          </div>
        </div>
      )}

      {/* Success View */}
      {activeView === 'success' && submissionResult && (
        <div className="w-full flex-1 flex flex-col pl-4 lg:pl-16 xl:pl-24 pr-4 gap-4 animate-fade-up">
          {/* Mobile Tab Switcher */}
          {isMobile && (
            <div className="flex border border-slate-200 bg-white p-1.5 rounded-2xl shadow-sm">
              <button
                type="button"
                onClick={() => setMobileSuccessTab('success')}
                className={`flex-1 text-center py-2.5 text-xs font-bold uppercase tracking-wider rounded-xl transition-all ${mobileSuccessTab === 'success' ? 'bg-slate-900 text-white shadow' : 'text-slate-500 hover:text-slate-800'}`}
              >
                🎉 Biên nhận
              </button>
              <button
                type="button"
                onClick={() => setMobileSuccessTab('preview')}
                className={`flex-1 text-center py-2.5 text-xs font-bold uppercase tracking-wider rounded-xl transition-all ${mobileSuccessTab === 'preview' ? 'bg-slate-900 text-white shadow' : 'text-slate-500 hover:text-slate-800'}`}
              >
                👁 Xem kết quả
              </button>
            </div>
          )}

          <div ref={splitContainerRef} className={`flex ${isMobile ? 'flex-col' : 'flex-row'} gap-4 items-stretch relative min-h-[600px] w-full`}>
            {/* Cột thông tin thành công bên trái */}
            <div
              style={{ width: isMobile ? '100%' : `${leftWidth}%` }}
              className={`glass-panel p-8 text-center space-y-6 h-fit flex-col ${isMobile && mobileSuccessTab !== 'success' ? 'hidden' : 'flex'}`}
            >
              <div className="w-16 h-16 bg-emerald-50 border border-emerald-100 text-emerald-600 flex items-center justify-center rounded-2xl mx-auto shadow-sm ring-4 ring-emerald-50">
                <svg className="w-7 h-7 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="3">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
                </svg>
              </div>

              <div>
                <h2 className="text-lg font-bold text-slate-900 font-display tracking-tight">Xuất hồ sơ thành công</h2>
                <p className="text-xs text-slate-500 mt-2 font-sans">
                  Hệ thống đã nạp và trộn dữ liệu thành công vào các file mẫu.
                </p>
              </div>

              <div className="bg-slate-50 p-6 rounded-2xl border border-slate-200 text-left space-y-3.5 text-xs text-slate-600 shadow-inner font-sans">
                <h3 className="text-[10px] font-bold text-zinc-550 uppercase tracking-wider border-b border-zinc-200/60 pb-2 mb-1 font-display">Chi tiết hồ sơ giao dịch</h3>
                <div className="flex items-center gap-2">
                  <svg className="w-3.5 h-3.5 text-zinc-450 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                  </svg>
                  <span><strong className="text-zinc-800">Biểu mẫu:</strong> {submissionResult.templateName}</span>
                </div>
                <div className="flex items-center gap-2">
                  <svg className="w-3.5 h-3.5 text-zinc-450 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" />
                  </svg>
                  <span><strong className="text-zinc-800">Khách hàng:</strong> {submissionResult.customerName}</span>
                </div>
                <div className="flex items-center gap-2">
                  <svg className="w-3.5 h-3.5 text-zinc-450 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M3 5a2 2 0 012-2h3.28a1 1 0 01.94.725l.548 2.2a1 1 0 01-.321.988l-1.305.98a10.582 10.582 0 004.872 4.872l.98-1.305a1 1 0 01.988-.321l2.2.548a1 1 0 01.725.94V19a2 2 0 01-2 2h-1C9.716 21 3 14.284 3 6V5z" />
                  </svg>
                  <span><strong className="text-zinc-800">Số điện thoại:</strong> {submissionResult.customerPhone}</span>
                </div>
              </div>

              {/* Danh sách các biểu mẫu tải lẻ */}
              <div className="bg-slate-50 p-6 rounded-2xl border border-slate-200 text-left space-y-4 text-xs text-slate-600 shadow-inner font-sans">
                <div className="flex justify-between items-center border-b border-zinc-200/60 pb-2 mb-1">
                  <h3 className="text-[10px] font-bold text-zinc-550 uppercase tracking-wider font-display">Danh sách tài liệu đã hoàn thành</h3>
                  {submissionResult.files && submissionResult.files.length > 1 && (
                    <button
                      type="button"
                      onClick={() => {
                        if (successSelectedFiles.length === submissionResult.files.length) {
                          setSuccessSelectedFiles([]);
                        } else {
                          setSuccessSelectedFiles([...submissionResult.files]);
                        }
                      }}
                      className="text-[10px] text-emerald-650 hover:text-emerald-700 font-bold transition-colors cursor-pointer"
                    >
                      {successSelectedFiles.length === submissionResult.files.length ? 'Bỏ chọn hết' : 'Chọn tất cả'}
                    </button>
                  )}
                </div>

                <div className="space-y-2 max-h-[220px] overflow-y-auto pr-1">
                  {submissionResult.files && submissionResult.files.map((filename) => {
                    const isChecked = successSelectedFiles.includes(filename);
                    const isMaster = filename === `${selectedTemplate.name}.docx`;
                    const isRepeated = filename.includes('_Căn_');
                    
                    let badgeLabel = 'Đồng bộ';
                    let badgeClass = 'bg-blue-50 text-blue-700 border-blue-500/20';
                    if (isMaster) {
                      badgeLabel = 'Gốc';
                      badgeClass = 'bg-emerald-50 text-emerald-700 border-emerald-500/20';
                    } else if (isRepeated) {
                      badgeLabel = 'Lặp';
                      badgeClass = 'bg-amber-50 text-amber-700 border-amber-500/20';
                    }

                    return (
                      <div 
                        key={filename} 
                        className={`flex items-center gap-2.5 p-3 rounded-xl border transition-all select-none ${
                          successPreviewFilename === filename 
                            ? 'bg-emerald-50/30 border-emerald-500 shadow-sm ring-1 ring-emerald-100/50' 
                            : 'bg-white border-zinc-200 hover:border-zinc-350'
                        }`}
                      >
                        <input
                          type="checkbox"
                          checked={isChecked}
                          onChange={(e) => {
                            if (e.target.checked) {
                              setSuccessSelectedFiles(prev => [...prev, filename]);
                            } else {
                              setSuccessSelectedFiles(prev => prev.filter(f => f !== filename));
                            }
                          }}
                          className="w-4 h-4 text-emerald-600 bg-white border-zinc-350 rounded accent-emerald-600 cursor-pointer"
                        />
                        <div 
                          onClick={() => setSuccessPreviewFilename(filename)}
                          className="flex-1 min-w-0 cursor-pointer text-left font-sans"
                        >
                          <span className="block text-xs font-bold text-zinc-800 truncate" title={filename}>
                            {isMaster ? '👑' : '📄'} {filename}
                          </span>
                        </div>
                        <span className={`text-[9px] border px-1.5 py-0.5 rounded font-medium select-none ${badgeClass}`}>{badgeLabel}</span>
                        <a
                          href={`http://localhost:5000/api/submissions/${submissionResult.submissionId}/download-file?filename=${encodeURIComponent(filename)}`}
                          download
                          className="text-zinc-400 hover:text-emerald-600 transition-colors p-1"
                          title="Tải nhanh file lẻ"
                          onClick={(e) => e.stopPropagation()}
                        >
                          <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2.5">
                            <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                          </svg>
                        </a>
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="space-y-3">
                <button
                  type="button"
                  onClick={handleDownloadSelectedFiles}
                  className="btn-premium block w-full bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-3.5 text-xs tracking-wider uppercase text-center shadow-md hover:shadow-emerald-500/20 hover:scale-[1.01] transition-all flex items-center justify-center gap-2 font-display cursor-pointer"
                >
                  <svg className="w-4 h-4 text-emerald-100 shrink-0 animate-bounce" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                  </svg>
                  Tải các tài liệu đã chọn (.docx)
                </button>

                <button
                  type="button"
                  onClick={() => setActiveView('dashboard')}
                  className="btn-premium block w-full bg-white hover:bg-slate-50 text-slate-700 font-semibold py-3.5 text-xs tracking-wider uppercase border border-slate-200 hover:border-slate-300 transition-colors text-center cursor-pointer font-display"
                >
                  Quay lại Bảng điều khiển
                </button>
              </div>
            </div>

            {/* Thanh kéo giãn (Resizer Handle) */}
            {!isMobile && (
              <div
                onMouseDown={startResizing}
                className="w-2 hover:w-3 bg-slate-200/80 hover:bg-emerald-600 cursor-col-resize self-stretch transition-all rounded-md flex items-center justify-center select-none shadow-inner group mx-1"
                title="Kéo chuột sang trái/phải để thay đổi kích thước"
              >
                <div className="w-0.5 h-8 bg-slate-400 group-hover:bg-white rounded-full transition-colors" />
              </div>
            )}

            {/* Cột xem trước tài liệu kết quả bên phải */}
            <div
              style={{ width: isMobile ? '100%' : `${100 - leftWidth}%` }}
              className={`flex-col flex-1 ${isMobile && mobileSuccessTab !== 'preview' ? 'hidden' : 'flex'}`}
            >
              {submissionResult.files && submissionResult.files.length > 1 && (
                <div className="flex bg-zinc-900 border-x border-t border-zinc-800 p-1.5 rounded-t-lg gap-1.5 overflow-x-auto select-none shrink-0">
                  {submissionResult.files.map(filename => (
                    <button
                      key={filename}
                      type="button"
                      onClick={() => setSuccessPreviewFilename(filename)}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer truncate max-w-[200px] ${successPreviewFilename === filename ? 'bg-zinc-800 text-white border border-zinc-700' : 'text-zinc-400 hover:text-zinc-200'}`}
                      title={filename}
                    >
                      {filename === `${selectedTemplate.name}.docx` ? '👑' : '📄'} {filename}
                    </button>
                  ))}
                </div>
              )}
              <DocxPreview
                fileUrl={`http://localhost:5000/api/submissions/${submissionResult.submissionId}/download-file?filename=${encodeURIComponent(successPreviewFilename || `${selectedTemplate.name}.docx`)}`}
                title={`${successPreviewFilename || selectedTemplate.name} (Bản hoàn thiện)`}
                fields={[]}
                liveData={{}}
              />
            </div>
          </div>
        </div>
      )}
      {/* PDF PREVIEW MODAL */}
      {showPdfModal && selectedTemplate && (() => {
        let totalPages = 1;
        const activeChildren = linkedChildren.filter(child => selectedChildIds.includes(child.id));
        activeChildren.forEach(child => {
          if (child.is_repeated) {
            totalPages += (formData[child.id] || []).length;
          } else {
            totalPages += 1;
          }
        });

        return (
          <div className="fixed inset-0 z-50 bg-zinc-900/90 flex flex-col font-sans overflow-hidden print:bg-white print:static print:inset-auto print:h-auto print:overflow-visible print-container-modal animate-fade-in">
            {/* Topbar of Modal */}
            <div className="bg-zinc-950 text-white px-6 py-4 flex justify-between items-center border-b border-zinc-800 shrink-0 print:hidden select-none">
              <div>
                <h3 className="font-bold text-sm text-emerald-400 uppercase tracking-wider font-display flex items-center gap-2">
                  <svg className="w-5 h-5 text-emerald-500 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2">
                    <path strokeLinecap="round" strokeLinejoin="round" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z" />
                  </svg>
                  XEM TRƯỚC BỘ TÀI LIỆU ({totalPages} BẢN)
                </h3>
                <p className="text-[10px] text-zinc-400 mt-0.5">Tự động điền dữ liệu & sẵn sàng in ấn</p>
              </div>
              <div className="flex items-center gap-2.5">
                <button
                  type="button"
                  onClick={() => window.print()}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold px-4 py-2 rounded-xl transition-all cursor-pointer shadow-md shadow-emerald-600/10 flex items-center gap-1.5"
                >
                  🖨 In / Xuất PDF
                </button>
                <button
                  type="button"
                  onClick={() => setShowPdfModal(false)}
                  className="bg-zinc-850 hover:bg-zinc-750 text-zinc-300 text-xs font-bold px-4 py-2 rounded-xl transition-all cursor-pointer border border-zinc-700"
                >
                  ✕ Đóng
                </button>
              </div>
            </div>

            {/* Modal Body */}
            <div className="flex-1 overflow-y-auto p-8 flex flex-col items-center gap-8 bg-zinc-900/40 print:bg-white print:p-0 print:overflow-visible print:block">
              {/* Sheet 1: Master Document */}
              <div className="pdf-page-wrapper w-full max-w-4xl print:max-w-none print:w-full print:block print:p-0">
                <div className="text-center text-[10px] text-zinc-400 font-bold uppercase tracking-widest mb-2.5 print:hidden select-none">
                  Tài liệu 1/{totalPages} — {selectedTemplate.name} (Gốc)
                </div>
                <div className="bg-white text-black p-12 shadow-2xl rounded border border-slate-200/50 min-h-[1120px] print:shadow-none print:border-none print:p-0 print:min-h-0 print:block doc-print-page">
                  <DocxPreviewInModal
                    fileUrl={`http://localhost:5000/api/templates/${selectedTemplate.id}/download-original?t=${previewKey}`}
                    liveData={formData}
                    fields={fields.filter(f => !f.groupName)}
                  />
                </div>
              </div>

              {/* Sheets for Linked Children */}
              {(() => {
                const pages = [];
                let currentPage = 2;

                activeChildren.forEach(child => {
                  if (child.is_repeated) {
                    const records = formData[child.id] || [];
                    records.forEach((record, rIdx) => {
                      const pageNum = currentPage++;
                      pages.push(
                        <div key={`${child.id}-${rIdx}`} className="pdf-page-wrapper w-full max-w-4xl print:max-w-none print:w-full print:block print:p-0 print:mt-10">
                          <div className="text-center text-[10px] text-zinc-400 font-bold uppercase tracking-widest mb-2.5 print:hidden select-none font-display">
                            Tài liệu {pageNum}/{totalPages} — {child.name} (Bản ghi {rIdx + 1})
                          </div>
                          <div className="bg-white text-black p-12 shadow-2xl rounded border border-slate-200/50 min-h-[1120px] print:shadow-none print:border-none print:p-0 print:min-h-0 print:block doc-print-page">
                            <DocxPreviewInModalChild
                              child={child}
                              previewKey={previewKey}
                              formData={formData}
                              recordData={record}
                            />
                          </div>
                        </div>
                      );
                    });
                  } else {
                    const pageNum = currentPage++;
                    pages.push(
                      <div key={child.id} className="pdf-page-wrapper w-full max-w-4xl print:max-w-none print:w-full print:block print:p-0 print:mt-10">
                        <div className="text-center text-[10px] text-zinc-400 font-bold uppercase tracking-widest mb-2.5 print:hidden select-none font-display">
                          Tài liệu {pageNum}/{totalPages} — {child.name}
                        </div>
                        <div className="bg-white text-black p-12 shadow-2xl rounded border border-slate-200/50 min-h-[1120px] print:shadow-none print:border-none print:p-0 print:min-h-0 print:block doc-print-page">
                          <DocxPreviewInModalChild
                            child={child}
                            previewKey={previewKey}
                            formData={formData}
                          />
                        </div>
                      </div>
                    );
                  }
                });

                return pages;
              })()}
            </div>
          </div>
        );
      })()}
        </main>
      </div>
    </div>
  );
}

