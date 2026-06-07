import { useState, useEffect, useRef, Fragment } from 'react';
import { Link } from 'react-router-dom';
import {
  DocxPreview,
  DocxPreviewInModal,
  DocxPreviewInModalChild
} from '../components/DocxPreview';
import Notification from '../components/Notification';
import {
  API_BASE,
  getGroupedFields,
  generateId
} from '../utils/helpers';

export default function CustomerView() {
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [sidebarActiveMenu, setSidebarActiveMenu] = useState('templates'); // templates | guide
  const [activeView, setActiveView] = useState('dashboard'); // dashboard | fill | success

  // Data State
  const [templates, setTemplates] = useState([]);
  const [categories, setCategories] = useState([]);
  const [selectedTemplate, setSelectedTemplate] = useState(null);
  const [fields, setFields] = useState([]);
  const [linkedChildren, setLinkedChildren] = useState([]);
  const [selectedChildIds, setSelectedChildIds] = useState([]);

  // Fill Form State
  const [formData, setFormData] = useState({});
  const [customerName, setCustomerName] = useState('');
  const [customerPhone, setCustomerPhone] = useState('');
  const [currentStep, setCurrentStep] = useState(1);
  const [showPdfModal, setShowPdfModal] = useState(false);

  // Preview State
  const [previewTemplateId, setPreviewTemplateId] = useState(null);
  const [previewFields, setPreviewFields] = useState([]);
  const [previewKey, setPreviewKey] = useState(1);

  // Loading & Action State
  const [isLoading, setIsLoading] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submissionProgress, setSubmissionProgress] = useState({ current: 0, total: 0, message: '' });
  const [submissionResult, setSubmissionResult] = useState(null);

  // Preview View State (before submit)
  const [previewActiveDocIdx, setPreviewActiveDocIdx] = useState(0);

  // Success View State
  const [successSelectedFiles, setSuccessSelectedFiles] = useState([]);
  const [successPreviewFilename, setSuccessPreviewFilename] = useState('');

  // Search & Filter State
  const [dashboardSearch, setDashboardSearch] = useState('');
  const [selectedCategoryId, setSelectedCategoryId] = useState('all');
  const [expandedCategoryIds, setExpandedCategoryIds] = useState(new Set());
  const [sortBy, setSortBy] = useState('newest');

  // UI Responsive State
  const [leftWidth, setLeftWidth] = useState(30);
  const [isMobile, setIsMobile] = useState(false);
  const [mobileFillTab, setMobileFillTab] = useState('form'); // form | preview
  const [mobileSuccessTab, setMobileSuccessTab] = useState('success'); // success | preview
  const [notification, setNotification] = useState(null);

  const splitContainerRef = useRef(null);
  const resizingRef = useRef(false);

  // Notifications
  const showNotification = (message, type = 'success') => {
    setNotification({ message, type });
    setTimeout(() => setNotification(null), 3000);
  };

  // Resize column panel helper
  useEffect(() => {
    const checkMobile = () => {
      setIsMobile(window.innerWidth < 1024);
    };
    checkMobile();
    window.addEventListener('resize', checkMobile);
    return () => window.removeEventListener('resize', checkMobile);
  }, []);

  const startResizing = (e) => {
    e.preventDefault();
    resizingRef.current = true;
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
  };

  const handleMouseMove = (e) => {
    if (!resizingRef.current || !splitContainerRef.current) return;
    const containerRect = splitContainerRef.current.getBoundingClientRect();
    const newWidth = ((e.clientX - containerRect.left) / containerRect.width) * 100;
    if (newWidth > 25 && newWidth < 75) {
      setLeftWidth(newWidth);
    }
  };

  const handleMouseUp = () => {
    resizingRef.current = false;
    document.removeEventListener('mousemove', handleMouseMove);
    document.removeEventListener('mouseup', handleMouseUp);
  };

  // Fetch templates and categories on mount
  const fetchData = async () => {
    setIsLoading(true);
    try {
      const [tRes, cRes] = await Promise.all([
        fetch(`${API_BASE}/templates`),
        fetch(`${API_BASE}/categories`)
      ]);
      if (tRes.ok) {
        const tData = await tRes.json();
        setTemplates(tData);
      }
      if (cRes.ok) {
        const cData = await cRes.json();
        setCategories(cData);
      }
    } catch (err) {
      showNotification('Không thể kết nối đến máy chủ API', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  // Category selection helpers
  const getCategoryChildren = (parentId) => {
    return categories.filter(c => c.parent_id === parentId);
  };

  const getDescendantCategoryIds = (catId) => {
    let ids = [];
    const children = categories.filter(c => c.parent_id === catId);
    children.forEach(child => {
      ids.push(child.id);
      ids = [...ids, ...getDescendantCategoryIds(child.id)];
    });
    return ids;
  };

  const toggleCategoryExpanded = (catId) => {
    setExpandedCategoryIds(prev => {
      const next = new Set(prev);
      if (next.has(catId)) next.delete(catId);
      else next.add(catId);
      return next;
    });
  };

  const getFilteredTemplates = () => {
    let list = [...templates];
    // Filter active only for customer
    list = list.filter(t => t.status === 'active');

    // Filter by search
    if (dashboardSearch.trim()) {
      const query = dashboardSearch.toLowerCase();
      list = list.filter(t => t.name.toLowerCase().includes(query));
    }

    // Filter by category
    if (selectedCategoryId !== 'all') {
      if (selectedCategoryId === 'uncategorized') {
        list = list.filter(t => !t.category_id);
      } else {
        const filterIds = [selectedCategoryId, ...getDescendantCategoryIds(selectedCategoryId)];
        list = list.filter(t => filterIds.includes(t.category_id));
      }
    }

    // Sort
    if (sortBy === 'newest') {
      list.sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0));
    } else if (sortBy === 'oldest') {
      list.sort((a, b) => new Date(a.created_at || 0) - new Date(b.created_at || 0));
    } else if (sortBy === 'name') {
      list.sort((a, b) => a.name.localeCompare(b.name));
    }

    return list;
  };

  // Open Template Wizard
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
      setPreviewKey(prev => prev + 1);

      // Init form structure
      const initialForm = {};
      finalFields.forEach(f => {
        const isRepeatedChildField = f.childTemplateId && currentLinks.find(c => c.id === f.childTemplateId)?.is_repeated === 1;
        if (!isRepeatedChildField) {
          initialForm[f.key_name] = f.field_type === 'boolean' ? false : '';
        }
      });

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
      setMobileFillTab('form');
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
    }
  };

  // Repeated child records logic
  const handleAddRepeatedRecord = (childId) => {
    const childFields = fields.filter(f => f.childTemplateId === childId);
    const newRecord = { _id: `rec-${Date.now()}-${Math.random().toString(36).substring(2, 6)}` };
    childFields.forEach(f => {
      newRecord[f.key_name] = f.field_type === 'boolean' ? false : '';
    });
    setFormData(prev => ({
      ...prev,
      [childId]: [...(prev[childId] || []), newRecord]
    }));
  };

  const handleRemoveRepeatedRecord = (childId, index) => {
    setFormData(prev => {
      const list = [...(prev[childId] || [])];
      if (list.length <= 1) return prev;
      list.splice(index, 1);
      return { ...prev, [childId]: list };
    });
  };

  const handleUpdateRepeatedField = (childId, recordIndex, keyName, value) => {
    setFormData(prev => {
      const list = [...(prev[childId] || [])];
      if (list[recordIndex]) {
        list[recordIndex] = { ...list[recordIndex], [keyName]: value };
      }
      return { ...prev, [childId]: list };
    });
  };

  const handleResetForm = () => {
    if (window.confirm("Bạn có chắc chắn muốn làm sạch toàn bộ form?")) {
      const initialForm = {};
      fields.forEach(f => {
        const isRepeatedChildField = f.childTemplateId && linkedChildren.find(c => c.id === f.childTemplateId)?.is_repeated === 1;
        if (!isRepeatedChildField) {
          initialForm[f.key_name] = f.field_type === 'boolean' ? false : '';
        }
      });
      linkedChildren.forEach(child => {
        if (child.is_repeated === 1) {
          const firstRecord = { _id: `rec-${Date.now()}-0` };
          fields.forEach(f => {
            if (f.childTemplateId === child.id) {
              firstRecord[f.key_name] = f.field_type === 'boolean' ? false : '';
            }
          });
          initialForm[child.id] = [firstRecord];
        }
      });
      setFormData(initialForm);
      showNotification('Đã đặt lại dữ liệu form!');
    }
  };

  // Submit and package documents
  const handleFormSubmit = async (e) => {
    if (e && e.preventDefault) e.preventDefault();
    setIsSubmitting(true);
    const childCount = selectedChildIds.length;
    const totalSteps = 2 + childCount + 2;

    setSubmissionProgress({ current: 1, total: totalSteps, message: 'Chuẩn bị dữ liệu...' });
    await new Promise(r => setTimeout(r, 300));

    try {
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
      setMobileSuccessTab('success');
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

  // Preview values resolution helper
  const getResolvedPreviewData = () => {
    if (previewTemplateId === selectedTemplate?.id) {
      return formData;
    }
    const child = linkedChildren.find(c => c.id === previewTemplateId);
    if (!child) return formData;

    if (child.is_repeated) {
      // For real-time preview, show the first record in repeated list
      const records = formData[child.id] || [];
      const firstRecord = records[0] || {};
      const resolved = {};
      previewFields.forEach(f => {
        if (f.parent_field_key) {
          resolved[f.key_name] = formData[f.parent_field_key] || '';
        } else {
          resolved[f.key_name] = firstRecord[f.key_name] || '';
        }
      });
      return resolved;
    } else {
      const resolved = {};
      previewFields.forEach(f => {
        if (f.parent_field_key) {
          resolved[f.key_name] = formData[f.parent_field_key] || '';
        } else {
          resolved[f.key_name] = formData[f.key_name] || '';
        }
      });
      return resolved;
    }
  };

  // Auto-switch preview tab if selected child template is deselected
  useEffect(() => {
    if (selectedTemplate && previewTemplateId && previewTemplateId !== selectedTemplate.id) {
      if (!selectedChildIds.includes(previewTemplateId)) {
        handleSwitchPreview(selectedTemplate.id);
      }
    }
  }, [selectedChildIds, selectedTemplate, previewTemplateId]);

  // Build steps dynamically: step 1 = main template, step N = Nth selected child with fields
  const activeSteps = (() => {
    const steps = [{ step: 1, label: selectedTemplate?.name || 'Thông tin hợp đồng', childId: null }];
    const activeChildren = linkedChildren.filter(c => selectedChildIds.includes(c.id));
    let stepNum = 2;
    for (const child of activeChildren) {
      if (fields.some(f => f.childTemplateId === child.id)) {
        steps.push({ step: stepNum++, label: child.name, childId: child.id });
      }
    }
    return steps;
  })();
  const totalSteps = activeSteps[activeSteps.length - 1]?.step || 1;

  const renderCategoryTree = (parentId = null, depth = 0) => {
    return getCategoryChildren(parentId).map(category => {
      const childCategories = getCategoryChildren(category.id);
      const hasChildren = childCategories.length > 0;
      const isExpanded = expandedCategoryIds.has(category.id);
      const descendantIds = [category.id, ...getDescendantCategoryIds(category.id)];
      const templateCount = templates.filter(t => t.status === 'active' && descendantIds.includes(t.category_id)).length;
      const isActive = selectedCategoryId === category.id;

      return (
        <div key={category.id}>
          <button
            type="button"
            className={`lx-cat-btn ${isActive ? 'active' : ''}`}
            style={{ paddingLeft: `${12 + depth * 14}px` }}
            onClick={() => {
              setSelectedCategoryId(category.id);
              if (hasChildren && !isExpanded) {
                setExpandedCategoryIds(prev => new Set(prev).add(category.id));
              }
            }}
          >
            <span style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              {hasChildren && (
                <span
                  className="material-symbols-outlined"
                  style={{ fontSize: 14, cursor: 'pointer', color: '#76777d' }}
                  onClick={(e) => { e.stopPropagation(); toggleCategoryExpanded(category.id); }}
                >
                  {isExpanded ? 'expand_more' : 'chevron_right'}
                </span>
              )}
              {!hasChildren && <span style={{ width: 14, display: 'inline-block' }} />}
              {category.name}
            </span>
            <span className="lx-cat-count">{templateCount}</span>
          </button>
          {hasChildren && isExpanded && renderCategoryTree(category.id, depth + 1)}
        </div>
      );
    });
  };

  // ── PRE-SUBMIT PREVIEW FULL-SCREEN PAGE ─────────────────────────
  if (activeView === 'preview' && selectedTemplate) {
    const activeChildren = linkedChildren.filter(c => selectedChildIds.includes(c.id));
    // Build doc list: main + children
    const docList = [
      { label: selectedTemplate.name, type: 'main' },
      ...activeChildren.map(c => ({ label: c.name, type: 'child', child: c }))
    ];
    const activeDoc = docList[previewActiveDocIdx] || docList[0];

    return (
      <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', background: '#f7f9fb', fontFamily: "'Inter', sans-serif" }}>
        {notification && (
          <div className={`lx-toast ${notification.type}`}>
            <span className="material-symbols-outlined">{notification.type === 'error' ? 'error' : 'check_circle'}</span>
            {notification.message}
          </div>
        )}

        {/* ── Header ── */}
        <header style={{ height: 60, background: '#ffffff', borderBottom: '1px solid #c6c6cd', display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '0 24px', flexShrink: 0, zIndex: 10 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <button
              className="lx-icon-btn-sm"
              onClick={() => setActiveView('fill')}
              title="Quay lại chỉnh sửa"
            >
              <span className="material-symbols-outlined">arrow_back</span>
            </button>
            <div style={{ width: 1, height: 22, background: '#c6c6cd' }} />
            <div>
              <div style={{ fontSize: 16, fontWeight: 600, color: '#000000', letterSpacing: '-0.01em' }}>Xem trước hồ sơ</div>
              <div style={{ fontSize: 12, color: '#76777d' }}>{selectedTemplate.name}</div>
            </div>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <button
              type="button"
              className="lx-btn lx-btn-secondary lx-btn-sm"
              onClick={() => setActiveView('fill')}
            >
              <span className="material-symbols-outlined" style={{ fontSize: 15 }}>edit</span>
              Chỉnh sửa
            </button>
            <button
              type="button"
              className="lx-btn lx-btn-primary lx-btn-sm"
              disabled={isSubmitting}
              onClick={() => handleFormSubmit()}
            >
              {isSubmitting ? (
                <>
                  <span className="lx-spinner" style={{ width: 14, height: 14, borderWidth: 2 }} />
                  Đang tạo hồ sơ...
                </>
              ) : (
                <>
                  <span className="material-symbols-outlined" style={{ fontSize: 15 }}>send</span>
                  Gửi hồ sơ
                </>
              )}
            </button>
          </div>
        </header>

        {/* ── Main workspace ── */}
        <div style={{ flex: 1, display: 'flex', overflow: 'hidden' }}>

          {/* Left sidebar — document tabs */}
          <aside style={{ width: 284, background: '#ffffff', borderRight: '1px solid #c6c6cd', display: 'flex', flexDirection: 'column', flexShrink: 0, overflow: 'hidden' }}>
            <div style={{ padding: '12px 16px', borderBottom: '1px solid #e6e8ea', background: '#f7f9fb', flexShrink: 0 }}>
              <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.07em', color: '#45464d' }}>
                Danh mục tài liệu ({docList.length})
              </div>
            </div>
            <nav style={{ flex: 1, overflowY: 'auto' }}>
              {docList.map((doc, idx) => {
                const isActive = previewActiveDocIdx === idx;
                return (
                  <button
                    key={idx}
                    type="button"
                    onClick={() => setPreviewActiveDocIdx(idx)}
                    style={{
                      width: '100%', display: 'flex', alignItems: 'flex-start', gap: 12,
                      padding: '14px 16px', textAlign: 'left', cursor: 'pointer',
                      background: isActive ? 'rgba(219,226,253,0.25)' : 'transparent',
                      borderLeft: `4px solid ${isActive ? '#000000' : 'transparent'}`,
                      border: 'none', borderBottom: '1px solid #f2f4f6',
                      transition: 'background 0.12s',
                    }}
                  >
                    <span className={`material-symbols-outlined${isActive ? ' ms-fill' : ''}`} style={{ fontSize: 20, color: isActive ? '#000000' : '#45464d', marginTop: 1, flexShrink: 0 }}>description</span>
                    <div style={{ overflow: 'hidden', flex: 1 }}>
                      <div style={{ fontSize: 12, fontWeight: 600, color: isActive ? '#000000' : '#191c1e', lineHeight: 1.3, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        {idx + 1}. {doc.label.length > 30 ? doc.label.slice(0, 30) + '…' : doc.label}
                      </div>
                      <div style={{ fontSize: 11, color: '#76777d', marginTop: 2 }}>
                        {doc.type === 'main' ? 'Hợp đồng chính' : 'Phụ lục / Biểu mẫu kèm theo'}
                      </div>
                    </div>
                  </button>
                );
              })}
            </nav>
          </aside>

          {/* Right — live preview (DocxPreviewInModal with formData) */}
          <section style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', background: '#e2e5e9' }}>
            {/* Toolbar */}
            <div style={{ height: 50, background: '#ffffff', borderBottom: '1px solid #e6e8ea', display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '0 20px', flexShrink: 0 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <span className="material-symbols-outlined" style={{ fontSize: 16, color: '#45464d' }}>preview</span>
                <span style={{ fontSize: 13, fontWeight: 500, color: '#191c1e' }}>
                  {activeDoc.label} — Bản xem trước (dữ liệu thực tế)
                </span>
              </div>
              <span style={{ fontSize: 11, color: '#76777d', fontStyle: 'italic' }}>Dữ liệu sẽ tự động điền vào tài liệu</span>
            </div>

            {/* Document preview — gray desktop, white A4 pages */}
            <div style={{ flex: 1, overflowY: 'auto', overflowX: 'auto', background: '#e2e5e9' }}>
              <div style={{ padding: '28px 20px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 24, minHeight: '100%' }}>
                {activeDoc.type === 'main' ? (
                  <div style={{ width: '100%', maxWidth: 860, background: '#ffffff', boxShadow: '0 2px 8px rgba(0,0,0,0.18)', borderRadius: 2, padding: '48px 60px' }}>
                    <DocxPreviewInModal
                      key={`main-${previewKey}`}
                      fileUrl={`${API_BASE}/templates/${selectedTemplate.id}/download-original?t=${previewKey}`}
                      liveData={formData}
                      fields={fields}
                    />
                  </div>
                ) : (
                  <div style={{ width: '100%', maxWidth: 860, background: '#ffffff', boxShadow: '0 2px 8px rgba(0,0,0,0.18)', borderRadius: 2, padding: '48px 60px' }}>
                    <DocxPreviewInModalChild
                      key={`child-${activeDoc.child.id}-${previewKey}`}
                      child={activeDoc.child}
                      previewKey={previewKey}
                      formData={formData}
                      recordData={null}
                    />
                  </div>
                )}
              </div>
            </div>
          </section>
        </div>

        {/* Submitting overlay */}
        {isSubmitting && (
          <div style={{ position: 'fixed', inset: 0, zIndex: 50, background: 'rgba(15,23,42,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <div style={{ background: '#ffffff', borderRadius: 8, padding: '32px 40px', textAlign: 'center', boxShadow: '0 8px 32px rgba(0,0,0,0.2)', minWidth: 280 }}>
              <div className="lx-spinner" style={{ width: 36, height: 36, margin: '0 auto 16px' }} />
              <div style={{ fontSize: 15, fontWeight: 600, color: '#191c1e', marginBottom: 6 }}>Đang tạo hồ sơ...</div>
              <div style={{ fontSize: 12, color: '#76777d' }}>{submissionProgress.message}</div>
            </div>
          </div>
        )}
      </div>
    );
  }

  // ── SUCCESS / PREVIEW FULL-SCREEN PAGE ──────────────────────────
  if (activeView === 'success' && submissionResult) {
    const files = submissionResult.files || [];
    const activeFile = successPreviewFilename || files[0] || '';

    return (
      <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', background: '#f7f9fb', fontFamily: "'Inter', sans-serif" }}>
        {notification && (
          <div className={`lx-toast ${notification.type}`}>
            <span className="material-symbols-outlined">{notification.type === 'error' ? 'error' : 'check_circle'}</span>
            {notification.message}
          </div>
        )}

        {/* ── Top header ── */}
        <header style={{ height: 60, background: '#ffffff', borderBottom: '1px solid #c6c6cd', display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '0 24px', flexShrink: 0, zIndex: 10 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <button className="lx-icon-btn-sm" onClick={() => setActiveView('dashboard')} title="Quay lại">
              <span className="material-symbols-outlined">arrow_back</span>
            </button>
            <div style={{ width: 1, height: 22, background: '#c6c6cd' }} />
            <div>
              <div style={{ fontSize: 16, fontWeight: 600, color: '#000000', letterSpacing: '-0.01em' }}>Kiểm tra lại hồ sơ</div>
              <div style={{ fontSize: 12, color: '#76777d' }}>
                {submissionResult.templateName} · {submissionResult.customerName}
              </div>
            </div>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginRight: 8 }}>
              <div style={{ width: 22, height: 22, borderRadius: '50%', background: '#6ffbbe', border: '2px solid #4edea3', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                <span className="material-symbols-outlined ms-fill" style={{ fontSize: 13, color: '#002113' }}>check</span>
              </div>
              <span style={{ fontSize: 12, fontWeight: 600, color: '#005236' }}>Soạn thảo hoàn tất</span>
            </div>
            <button
              type="button"
              className="lx-btn lx-btn-secondary lx-btn-sm"
              onClick={handleDownloadSelectedFiles}
            >
              <span className="material-symbols-outlined" style={{ fontSize: 15 }}>download</span>
              Tải {successSelectedFiles.length > 0 ? `(${successSelectedFiles.length})` : 'tất cả'}
            </button>
            <button
              type="button"
              className="lx-btn lx-btn-primary lx-btn-sm"
              onClick={() => setShowPdfModal(true)}
            >
              <span className="material-symbols-outlined" style={{ fontSize: 15 }}>send</span>
              Gửi hồ sơ
            </button>
          </div>
        </header>

        {/* ── Main workspace ── */}
        <div style={{ flex: 1, display: 'flex', overflow: 'hidden' }}>

          {/* Left sidebar — document tabs */}
          <aside style={{ width: 284, background: '#ffffff', borderRight: '1px solid #c6c6cd', display: 'flex', flexDirection: 'column', flexShrink: 0, overflow: 'hidden' }}>
            <div style={{ padding: '12px 16px', borderBottom: '1px solid #e6e8ea', background: '#f7f9fb', flexShrink: 0, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.07em', color: '#45464d' }}>
                Danh mục tài liệu ({files.length})
              </div>
              {files.length > 1 && (
                <button
                  type="button"
                  className="lx-btn lx-btn-ghost lx-btn-sm"
                  style={{ fontSize: 10, padding: '3px 8px' }}
                  onClick={() => {
                    if (successSelectedFiles.length === files.length) {
                      setSuccessSelectedFiles([]);
                    } else {
                      setSuccessSelectedFiles([...files]);
                    }
                  }}
                >
                  {successSelectedFiles.length === files.length ? 'Bỏ chọn' : 'Chọn hết'}
                </button>
              )}
            </div>
            <nav style={{ flex: 1, overflowY: 'auto' }}>
              {files.map((filename, idx) => {
                const isActive = activeFile === filename;
                const isChecked = successSelectedFiles.includes(filename);
                const isMaster = filename === `${selectedTemplate?.name}.docx`;
                const isRepeated = filename.includes('_Căn_');
                const shortName = filename.replace(/\.docx$/i, '').replace(/_/g, ' ');
                let badgeClass = 'lx-badge-draft';
                let badgeLabel = 'Đồng bộ';
                if (isMaster) { badgeClass = 'lx-badge-published'; badgeLabel = 'Gốc'; }
                else if (isRepeated) { badgeClass = 'lx-badge-warn'; badgeLabel = 'Lặp'; }

                return (
                  <button
                    key={filename}
                    type="button"
                    onClick={() => setSuccessPreviewFilename(filename)}
                    style={{
                      width: '100%', display: 'flex', alignItems: 'flex-start', gap: 10,
                      padding: '12px 16px', textAlign: 'left', cursor: 'pointer',
                      background: isActive ? 'rgba(219,226,253,0.25)' : 'transparent',
                      borderLeft: `4px solid ${isActive ? '#000000' : 'transparent'}`,
                      border: 'none', borderBottom: '1px solid #f2f4f6',
                      transition: 'background 0.12s',
                    }}
                  >
                    {/* Checkbox */}
                    <input
                      type="checkbox"
                      checked={isChecked}
                      style={{ accentColor: '#000000', flexShrink: 0, marginTop: 3 }}
                      onChange={e => {
                        e.stopPropagation();
                        if (e.target.checked) setSuccessSelectedFiles(prev => [...prev, filename]);
                        else setSuccessSelectedFiles(prev => prev.filter(f => f !== filename));
                      }}
                      onClick={e => e.stopPropagation()}
                    />
                    <span className={`material-symbols-outlined${isActive ? ' ms-fill' : ''}`} style={{ fontSize: 18, color: isActive ? '#000000' : '#45464d', marginTop: 1, flexShrink: 0 }}>description</span>
                    <div style={{ overflow: 'hidden', flex: 1 }}>
                      <div style={{ fontSize: 12, fontWeight: 600, color: isActive ? '#000000' : '#191c1e', lineHeight: 1.3, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        {idx + 1}. {shortName.length > 28 ? shortName.slice(0, 28) + '…' : shortName}
                      </div>
                      <div style={{ marginTop: 4, display: 'flex', alignItems: 'center', gap: 6 }}>
                        <span className={`lx-badge ${badgeClass}`} style={{ fontSize: 10 }}>{badgeLabel}</span>
                        <a
                          href={`http://localhost:5000/api/submissions/${submissionResult.submissionId}/download-file?filename=${encodeURIComponent(filename)}`}
                          download
                          className="lx-icon-btn-sm"
                          style={{ padding: 2 }}
                          onClick={e => e.stopPropagation()}
                          title="Tải xuống"
                        >
                          <span className="material-symbols-outlined" style={{ fontSize: 14 }}>download</span>
                        </a>
                      </div>
                    </div>
                  </button>
                );
              })}
            </nav>

            {/* Receipt info at bottom */}
            <div style={{ padding: '14px 16px', borderTop: '1px solid #e6e8ea', background: '#f7f9fb', flexShrink: 0, fontSize: 12 }}>
              <div style={{ fontWeight: 600, color: '#191c1e', marginBottom: 8 }}>Chi tiết hồ sơ</div>
              {[
                { label: 'Biểu mẫu', value: submissionResult.templateName },
                { label: 'Khách hàng', value: submissionResult.customerName },
                { label: 'Điện thoại', value: submissionResult.customerPhone },
              ].map(row => (
                <div key={row.label} style={{ display: 'flex', justifyContent: 'space-between', gap: 8, marginBottom: 4, color: '#45464d' }}>
                  <span>{row.label}:</span>
                  <span style={{ fontWeight: 500, color: '#191c1e', textAlign: 'right', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 160 }}>{row.value}</span>
                </div>
              ))}
            </div>
          </aside>

          {/* Right — document viewer */}
          <section style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', background: '#e2e5e9' }}>
            {/* Viewer toolbar */}
            <div style={{ height: 50, background: '#ffffff', borderBottom: '1px solid #e6e8ea', display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '0 20px', flexShrink: 0 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <span className="material-symbols-outlined" style={{ fontSize: 16, color: '#45464d' }}>description</span>
                <span style={{ fontSize: 13, fontWeight: 500, color: '#191c1e', maxWidth: 460, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{activeFile}</span>
              </div>
              {activeFile && (
                <a
                  href={`http://localhost:5000/api/submissions/${submissionResult.submissionId}/download-file?filename=${encodeURIComponent(activeFile)}`}
                  download
                  className="lx-btn lx-btn-secondary lx-btn-sm"
                >
                  <span className="material-symbols-outlined" style={{ fontSize: 14 }}>download</span>
                  Tải xuống tệp này
                </a>
              )}
            </div>

            {/* DocxPreview */}
            <div style={{ flex: 1, overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
              {activeFile ? (
                <DocxPreview
                  key={activeFile}
                  fileUrl={`http://localhost:5000/api/submissions/${submissionResult.submissionId}/download-file?filename=${encodeURIComponent(activeFile)}`}
                  title={activeFile}
                  fields={[]}
                  liveData={{}}
                  hideToolbar={true}
                />
              ) : (
                <div className="lx-empty" style={{ background: '#e2e5e9' }}>Không có tài liệu nào để xem trước.</div>
              )}
            </div>
          </section>
        </div>

        {/* PDF modal — kept intact */}
        {showPdfModal && selectedTemplate && (() => {
          let totalPages = 1;
          const activeChildren = linkedChildren.filter(child => selectedChildIds.includes(child.id));
          activeChildren.forEach(child => {
            if (child.is_repeated) totalPages += (formData[child.id] || []).length;
            else totalPages += 1;
          });
          return (
            <div style={{ position: 'fixed', inset: 0, zIndex: 50, background: 'rgba(15,23,42,0.85)', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
              <div style={{ background: '#ffffff', borderBottom: '1px solid #c6c6cd', padding: '14px 24px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexShrink: 0 }}>
                <div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: '#000000', display: 'flex', alignItems: 'center', gap: 8 }}>
                    <span className="material-symbols-outlined" style={{ fontSize: 20 }}>visibility</span>
                    Xem trước bộ tài liệu ({totalPages} bản)
                  </div>
                  <div style={{ fontSize: 11, color: '#76777d', marginTop: 2 }}>Tự động điền dữ liệu &amp; sẵn sàng in ấn</div>
                </div>
                <div style={{ display: 'flex', gap: 8 }}>
                  <button type="button" className="lx-btn lx-btn-secondary lx-btn-sm" onClick={() => window.print()}>
                    <span className="material-symbols-outlined" style={{ fontSize: 15 }}>print</span>In / Xuất PDF
                  </button>
                  <button type="button" className="lx-btn lx-btn-ghost lx-btn-sm" onClick={() => setShowPdfModal(false)}>
                    <span className="material-symbols-outlined" style={{ fontSize: 15 }}>close</span>Đóng
                  </button>
                </div>
              </div>
              <div style={{ flex: 1, overflowY: 'auto', padding: '32px 24px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 32, background: 'rgba(15,23,42,0.3)' }}>
                <div style={{ width: '100%', maxWidth: 860 }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: '#c6c6cd', textTransform: 'uppercase', letterSpacing: '0.1em', textAlign: 'center', marginBottom: 8 }}>
                    Tài liệu 1/{totalPages} — {selectedTemplate.name} (Gốc)
                  </div>
                  <div style={{ background: '#ffffff', padding: '48px', boxShadow: '0 8px 32px rgba(0,0,0,0.2)', borderRadius: 4 }}>
                    <DocxPreviewInModal fileUrl={`http://localhost:5000/api/templates/${selectedTemplate.id}/download-original?t=${previewKey}`} liveData={formData} fields={fields.filter(f => !f.groupName)} />
                  </div>
                </div>
                {(() => {
                  const pages = [];
                  let currentPage = 2;
                  activeChildren.forEach(child => {
                    if (child.is_repeated) {
                      const records = formData[child.id] || [];
                      records.forEach((record, rIdx) => {
                        pages.push(
                          <div key={`${child.id}-${rIdx}`} style={{ width: '100%', maxWidth: 860 }}>
                            <div style={{ fontSize: 10, fontWeight: 700, color: '#c6c6cd', textTransform: 'uppercase', letterSpacing: '0.1em', textAlign: 'center', marginBottom: 8 }}>
                              Tài liệu {currentPage++}/{totalPages} — {child.name} (Căn {rIdx + 1})
                            </div>
                            <div style={{ background: '#ffffff', padding: '48px', boxShadow: '0 8px 32px rgba(0,0,0,0.2)', borderRadius: 4 }}>
                              <DocxPreviewInModalChild child={child} previewKey={previewKey} formData={formData} recordData={record} />
                            </div>
                          </div>
                        );
                      });
                    } else {
                      pages.push(
                        <div key={child.id} style={{ width: '100%', maxWidth: 860 }}>
                          <div style={{ fontSize: 10, fontWeight: 700, color: '#c6c6cd', textTransform: 'uppercase', letterSpacing: '0.1em', textAlign: 'center', marginBottom: 8 }}>
                            Tài liệu {currentPage++}/{totalPages} — {child.name}
                          </div>
                          <div style={{ background: '#ffffff', padding: '48px', boxShadow: '0 8px 32px rgba(0,0,0,0.2)', borderRadius: 4 }}>
                            <DocxPreviewInModalChild child={child} previewKey={previewKey} formData={formData} recordData={null} />
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
      </div>
    );
  }

  return (
    <div className="lx-app">
      {/* Toast Notification */}
      {notification && (
        <div className={`lx-toast ${notification.type}`}>
          <span className="material-symbols-outlined">
            {notification.type === 'error' ? 'error' : 'check_circle'}
          </span>
          {notification.message}
        </div>
      )}

      {/* Sidebar */}
      <nav className="lx-sidebar">
        <div className="lx-sidebar-logo">
          <div className="lx-logo-icon">
            <span className="material-symbols-outlined ms-fill" style={{ fontSize: '22px' }}>gavel</span>
          </div>
          <div className="lx-logo-text">
            <h1>LexNotary</h1>
            <p>Notary Portal</p>
          </div>
        </div>

        <div className="lx-sidebar-cta">
          {selectedTemplate ? (
            <button
              className="lx-sidebar-cta-btn"
              onClick={() => handleOpenFill(selectedTemplate.id)}
            >
              <span className="material-symbols-outlined">add</span>
              Tạo Hồ Sơ Mới
            </button>
          ) : (
            <div style={{ height: 38 }} />
          )}
        </div>

        <div className="lx-sidebar-nav">
          <button
            className={`lx-nav-item ${sidebarActiveMenu === 'templates' && activeView === 'dashboard' ? 'active' : ''}`}
            onClick={() => { setSidebarActiveMenu('templates'); setActiveView('dashboard'); }}
          >
            <span className="material-symbols-outlined">description</span>
            <span>Biểu Mẫu</span>
          </button>
          <button
            className={`lx-nav-item ${sidebarActiveMenu === 'guide' && activeView === 'dashboard' ? 'active' : ''}`}
            onClick={() => { setSidebarActiveMenu('guide'); setActiveView('dashboard'); }}
          >
            <span className="material-symbols-outlined">help_outline</span>
            <span>Hướng Dẫn</span>
          </button>
        </div>

        <div className="lx-sidebar-bottom">
          <Link to="/admin" className="lx-nav-item">
            <span className="material-symbols-outlined">admin_panel_settings</span>
            <span>Quản Trị</span>
          </Link>
        </div>
      </nav>

      {/* Main */}
      <div className="lx-main">
        <header className="lx-header">
          <div className="lx-search">
            <span className="material-symbols-outlined">search</span>
            <input
              value={dashboardSearch}
              onChange={e => setDashboardSearch(e.target.value)}
              placeholder="Tìm kiếm biểu mẫu..."
            />
          </div>
          <div className="lx-header-right">
            <button className="lx-icon-btn">
              <span className="material-symbols-outlined">notifications</span>
            </button>
            <div className="lx-divider-v" />
            <div className="lx-user">
              <div className="lx-user-info">
                <div className="lx-user-name">Khách hàng</div>
                <div className="lx-user-role">Nộp hồ sơ trực tuyến</div>
              </div>
              <div className="lx-avatar">
                <span className="material-symbols-outlined ms-fill" style={{ fontSize: '34px' }}>account_circle</span>
              </div>
            </div>
          </div>
        </header>

        {/* ── VIEW: TEMPLATE DASHBOARD ── */}
        {activeView === 'dashboard' && sidebarActiveMenu === 'templates' && (
          <div className="lx-content">
            <div className="lx-page-header">
              <div className="lx-page-title">Chọn biểu mẫu công chứng</div>
              <div className="lx-page-subtitle">Chọn loại hợp đồng phù hợp để bắt đầu điền hồ sơ trực tuyến.</div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '220px 1fr', gap: 24, alignItems: 'start' }}>
              {/* Category panel */}
              <div className="lx-card">
                <div className="lx-card-header">
                  <span className="lx-card-title" style={{ fontSize: 13 }}>Danh mục</span>
                  <span className="lx-cat-count">{categories.length}</span>
                </div>
                <div style={{ padding: '8px 8px' }}>
                  <button
                    type="button"
                    className={`lx-cat-btn ${selectedCategoryId === 'all' ? 'active' : ''}`}
                    onClick={() => setSelectedCategoryId('all')}
                  >
                    <span>Tất cả</span>
                    <span className="lx-cat-count">{templates.filter(t => t.status === 'active').length}</span>
                  </button>
                  <button
                    type="button"
                    className={`lx-cat-btn ${selectedCategoryId === 'uncategorized' ? 'active' : ''}`}
                    onClick={() => setSelectedCategoryId('uncategorized')}
                  >
                    <span>Chưa phân loại</span>
                    <span className="lx-cat-count">{templates.filter(t => t.status === 'active' && !t.category_id).length}</span>
                  </button>
                  {categories.length > 0 && (
                    <div style={{ borderTop: '1px solid #e6e8ea', marginTop: 6, paddingTop: 6 }}>
                      {renderCategoryTree()}
                    </div>
                  )}
                </div>
              </div>

              {/* Template grid */}
              <div>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16, gap: 12 }}>
                  <span style={{ fontSize: 13, fontWeight: 600, color: '#191c1e' }}>
                    Biểu mẫu sẵn có
                    <span style={{ marginLeft: 8, fontSize: 12, color: '#76777d', fontWeight: 500 }}>
                      ({getFilteredTemplates().length})
                    </span>
                  </span>
                  <select
                    className="lx-select"
                    style={{ width: 'auto' }}
                    value={sortBy}
                    onChange={e => setSortBy(e.target.value)}
                  >
                    <option value="newest">Mới nhất</option>
                    <option value="oldest">Cũ nhất</option>
                    <option value="name">Tên A → Z</option>
                  </select>
                </div>

                {isLoading && templates.length === 0 ? (
                  <div className="lx-empty">
                    <span className="lx-spinner" style={{ display: 'inline-block', marginBottom: 12 }} />
                    <div>Đang tải biểu mẫu...</div>
                  </div>
                ) : getFilteredTemplates().length === 0 ? (
                  <div className="lx-empty">
                    <span className="material-symbols-outlined" style={{ fontSize: 40, marginBottom: 8, color: '#c6c6cd' }}>folder_open</span>
                    <div>Không tìm thấy biểu mẫu phù hợp.</div>
                  </div>
                ) : (
                  <div className="lx-template-grid">
                    {getFilteredTemplates().map(temp => (
                      <div key={temp.id} className="lx-template-card">
                        <div className="lx-template-card-header">
                          <div className="lx-template-icon">
                            <span className="material-symbols-outlined">draft</span>
                          </div>
                          <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                            <span className="lx-badge lx-badge-pending" style={{ fontSize: 10 }}>
                              {temp.variables_count} trường
                            </span>
                            {temp.children_count > 0 && (
                              <span className="lx-badge lx-badge-draft" style={{ fontSize: 10 }}>
                                +{temp.children_count} phụ lục
                              </span>
                            )}
                          </div>
                        </div>
                        <div
                          className="lx-template-card-name"
                          style={{ cursor: 'pointer' }}
                          onClick={() => handleOpenFill(temp.id)}
                        >
                          {temp.name}
                        </div>
                        <div className="line-clamp-2" style={{ fontSize: 12, color: '#76777d', flexGrow: 1 }}>
                          {temp.description || 'Biểu mẫu công chứng — bấm để bắt đầu điền hồ sơ trực tuyến.'}
                        </div>
                        <div className="lx-template-card-actions">
                          <button
                            type="button"
                            className="lx-btn lx-btn-primary"
                            style={{ width: '100%', justifyContent: 'center' }}
                            onClick={() => handleOpenFill(temp.id)}
                          >
                            Bắt đầu điền hồ sơ
                            <span className="material-symbols-outlined" style={{ fontSize: 16 }}>arrow_forward</span>
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* ── VIEW: GUIDE ── */}
        {activeView === 'dashboard' && sidebarActiveMenu === 'guide' && (
          <div className="lx-content">
            <div className="lx-page-header">
              <div className="lx-page-title">Hướng dẫn sử dụng</div>
              <div className="lx-page-subtitle">Hoàn tất hồ sơ pháp lý chỉ với 4 bước trực tuyến.</div>
            </div>

            <div className="lx-card" style={{ maxWidth: 680 }}>
              <div className="lx-card-header">
                <span className="lx-card-title">Quy trình nộp hồ sơ</span>
              </div>
              <div style={{ padding: '20px 24px', display: 'flex', flexDirection: 'column', gap: 28 }}>
                {[
                  {
                    num: 1,
                    title: 'Chọn biểu mẫu phù hợp',
                    desc: 'Duyệt qua danh mục biểu mẫu giao dịch tại mục Biểu mẫu. Nhấn Bắt đầu điền hồ sơ để khởi tạo.'
                  },
                  {
                    num: 2,
                    title: 'Điền thông tin cá nhân và giao dịch',
                    desc: 'Nhập họ tên, số điện thoại để văn phòng liên hệ. Sau đó hoàn thiện nội dung theo trình tự 4 bước: Bên bán, Bên mua, Tài sản, Hợp đồng.'
                  },
                  {
                    num: 3,
                    title: 'Chọn biểu mẫu phụ lục kèm theo',
                    desc: 'Hệ thống tự động liên kết các biểu mẫu con (Tờ khai thuế TNCN, Lệ phí trước bạ). Dữ liệu từ file gốc sẽ tự động đồng bộ sang các phụ lục.'
                  },
                  {
                    num: 4,
                    title: 'Xuất tài liệu và nộp hồ sơ',
                    desc: 'Xem trước toàn bộ tài liệu đã điền. Nhấn Xác nhận & Xuất File để tải trọn bộ hồ sơ. Hồ sơ sẽ được gửi đến Văn phòng công chứng để chuẩn bị bản in.'
                  }
                ].map(step => (
                  <div key={step.num} style={{ display: 'flex', gap: 16, alignItems: 'flex-start' }}>
                    <div style={{
                      width: 32, height: 32, borderRadius: '50%',
                      background: '#131b2e', color: '#ffffff',
                      display: 'flex', alignItems: 'center', justifyContent: 'center',
                      fontSize: 13, fontWeight: 700, flexShrink: 0
                    }}>
                      {step.num}
                    </div>
                    <div>
                      <div style={{ fontSize: 14, fontWeight: 600, color: '#191c1e', marginBottom: 4 }}>{step.title}</div>
                      <div style={{ fontSize: 13, color: '#45464d', lineHeight: 1.6 }}>{step.desc}</div>
                    </div>
                  </div>
                ))}
              </div>
              <div style={{
                padding: '14px 24px',
                background: '#f2f4f6',
                borderTop: '1px solid #e6e8ea',
                display: 'flex', alignItems: 'center', gap: 10,
                fontSize: 12, color: '#45464d'
              }}>
                <span className="material-symbols-outlined" style={{ fontSize: 18, color: '#76777d' }}>info</span>
                Mọi thắc mắc, quý khách vui lòng liên hệ hotline <strong style={{ marginLeft: 4 }}>090.123.4567</strong>&nbsp;hoặc đến trực tiếp Văn phòng Công chứng Trung tâm.
              </div>
            </div>
          </div>
        )}

        {/* ── VIEW: FILL FORM ── */}
        {activeView === 'fill' && selectedTemplate && (
          <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', position: 'relative' }}>
            {/* Mobile tab switcher */}
            {isMobile && (
              <div className="lx-tabs" style={{ margin: '8px', borderRadius: 4 }}>
                <button
                  type="button"
                  className={`lx-tab ${mobileFillTab === 'form' ? 'active' : ''}`}
                  style={{ flex: 1 }}
                  onClick={() => setMobileFillTab('form')}
                >
                  Điền hồ sơ
                </button>
                <button
                  type="button"
                  className={`lx-tab ${mobileFillTab === 'preview' ? 'active' : ''}`}
                  style={{ flex: 1 }}
                  onClick={() => setMobileFillTab('preview')}
                >
                  Xem tài liệu
                </button>
              </div>
            )}

            <div ref={splitContainerRef} className="lx-split" style={{ flex: 1 }}>
              {/* Left panel: Form */}
              <div
                style={{
                  width: isMobile ? '100%' : `${leftWidth}%`,
                  display: isMobile && mobileFillTab !== 'form' ? 'none' : 'flex',
                  flexDirection: 'column',
                  flexShrink: 0,
                  overflow: 'hidden',
                  background: '#ffffff',
                  borderRight: isMobile ? 'none' : '1px solid #e6e8ea'
                }}
              >
                {/* Form header */}
                <div style={{ padding: '16px 20px', borderBottom: '1px solid #e6e8ea', flexShrink: 0, background: '#ffffff' }}>
                  <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 12 }}>
                    <div>
                      <div style={{ fontSize: 15, fontWeight: 700, color: '#000000', letterSpacing: '-0.01em' }}>
                        {selectedTemplate.name}
                      </div>
                      <div style={{ fontSize: 12, color: '#76777d', marginTop: 2 }}>
                        Điền đầy đủ thông tin hồ sơ công chứng theo hướng dẫn bên dưới.
                      </div>
                    </div>
                    <div style={{ display: 'flex', gap: 6, flexShrink: 0 }}>
                      <button
                        type="button"
                        className="lx-btn lx-btn-ghost lx-btn-sm"
                        onClick={handleResetForm}
                      >
                        Nhập lại
                      </button>
                      <button
                        type="button"
                        className="lx-btn lx-btn-secondary lx-btn-sm"
                        onClick={() => setShowPdfModal(true)}
                      >
                        <span className="material-symbols-outlined" style={{ fontSize: 15 }}>visibility</span>
                        Xem trước
                      </button>
                    </div>
                  </div>
                </div>

                {/* Stepper — dynamic from linked forms */}
                <div className="lx-stepper">
                  {activeSteps.map((item, index) => {
                    const isCurrent = currentStep === item.step;
                    const isCompleted = currentStep > item.step;
                    const shortLabel = item.label.length > 20 ? item.label.slice(0, 20) + '…' : item.label;
                    return (
                      <Fragment key={item.step}>
                        <div className={`lx-step ${isCurrent ? 'active' : ''} ${isCompleted ? 'done' : ''}`}>
                          <button
                            type="button"
                            className="lx-step-num"
                            onClick={() => setCurrentStep(item.step)}
                            style={{ cursor: 'pointer', border: 'none', background: 'inherit', color: 'inherit', padding: 0, width: 22, height: 22, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 10, fontWeight: 700 }}
                          >
                            {isCompleted
                              ? <span className="material-symbols-outlined" style={{ fontSize: 13 }}>check</span>
                              : item.step}
                          </button>
                          <span className="lx-step-label">{shortLabel}</span>
                        </div>
                        {index < activeSteps.length - 1 && <div className="lx-step-line" />}
                      </Fragment>
                    );
                  })}
                </div>

                {/* Form scroll area */}
                <div style={{ flex: 1, overflowY: 'auto' }}>
                  <form onSubmit={handleFormSubmit} style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: 16 }}>

                    {/* Linked child document selection (step 1) */}
                    {currentStep === 1 && linkedChildren && linkedChildren.length > 0 && (
                      <div className="lx-card">
                        <div className="lx-card-header">
                          <span className="lx-card-title" style={{ fontSize: 13 }}>Các tài liệu kèm theo bộ hồ sơ</span>
                        </div>
                        <div style={{ padding: '12px 16px', display: 'flex', flexDirection: 'column', gap: 8 }}>
                          {/* Master doc (always required) */}
                          <div style={{
                            display: 'flex', alignItems: 'center', gap: 10,
                            padding: '10px 12px', borderRadius: 4,
                            background: '#f7f9fb', border: '1px solid #e6e8ea', opacity: 0.85
                          }}>
                            <input type="checkbox" checked disabled style={{ accentColor: '#000000' }} />
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ fontSize: 13, fontWeight: 600, color: '#191c1e', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                {selectedTemplate.name}
                              </div>
                              <div style={{ fontSize: 11, color: '#76777d', marginTop: 2 }}>Tài liệu chính — bắt buộc</div>
                            </div>
                            <span className="lx-badge lx-badge-published" style={{ fontSize: 10 }}>Gốc</span>
                          </div>

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
                                style={{
                                  display: 'flex', alignItems: 'center', gap: 10,
                                  padding: '10px 12px', borderRadius: 4,
                                  background: isChecked ? '#f7f9fb' : '#ffffff',
                                  border: isChecked ? '1px solid #000000' : '1px solid #c6c6cd',
                                  cursor: 'pointer', transition: 'border-color 0.12s'
                                }}
                              >
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  readOnly
                                  style={{ accentColor: '#000000', pointerEvents: 'none' }}
                                />
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <div style={{ fontSize: 13, fontWeight: 600, color: '#191c1e', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                    {child.name}
                                  </div>
                                  <div style={{ fontSize: 11, color: '#76777d', marginTop: 2 }}>
                                    {isChecked ? 'Đang bật — nhấn để tắt' : 'Đang tắt — nhấn để bật'}
                                  </div>
                                </div>
                                <span className={`lx-badge ${isChecked ? 'lx-badge-success' : 'lx-badge-draft'}`} style={{ fontSize: 10 }}>
                                  {isChecked ? 'Bật' : 'Tắt'}
                                </span>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    )}

                    {/* Dynamic fields */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
                      {(() => {
                        const currentStepDef = activeSteps.find(s => s.step === currentStep);
                        const stepFields = fields.filter(f => {
                          if (!currentStepDef) return false;
                          if (currentStepDef.childId === null) return !f.childTemplateId;
                          return f.childTemplateId === currentStepDef.childId;
                        });
                        if (stepFields.length === 0) {
                          return (
                            <div className="lx-empty">
                              <span className="material-symbols-outlined" style={{ fontSize: 32, marginBottom: 8, color: '#c6c6cd' }}>info</span>
                              <div>Tài liệu liên quan tới bước này đã bị bỏ chọn.</div>
                              <div style={{ marginTop: 4, fontSize: 12, color: '#76777d' }}>Bạn có thể nhấn "Tiếp theo" để tiếp tục.</div>
                            </div>
                          );
                        }

                        // Shared field rendering helpers (used by both label-path and grouped-path)
                        const renderCustField = (field, value, onChange) => (
                          <div key={field.id} className="cust-field">
                            <label className="cust-label">
                              {field.label}{!!field.is_required && <span style={{ color: '#c0392b' }}> *</span>}
                            </label>
                            {field.field_type === 'boolean' ? (
                              <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', padding: '7px 10px', border: '1px solid #c6c6cd', borderRadius: 5, background: '#f7f9fb' }}>
                                <input type="checkbox" checked={!!value} onChange={e => onChange(e.target.checked)}
                                  style={{ accentColor: '#b8924a', width: 14, height: 14 }} />
                                <span style={{ fontSize: 12, color: '#45464d', fontWeight: 500 }}>Kích hoạt / Xác nhận</span>
                              </label>
                            ) : (
                              <input
                                className="cust-input"
                                type={field.field_type === 'date' ? 'date' : field.field_type === 'number' ? 'number' : 'text'}
                                required={!!field.is_required}
                                value={value || ''}
                                onChange={e => onChange(e.target.value)}
                              />
                            )}
                          </div>
                        );

                        const isFullWidthField = (f) => {
                          const k = (f.key_name || '').toLowerCase();
                          const l = (f.label || '').toLowerCase();
                          return ['ho_ten','hoten','full_name','dia_chi','diachi','noi_cu','noicu','noi_sinh','address','duong','xa_','phuong_','quan_','tinh_','huyen_'].some(x => k.includes(x))
                            || ['họ tên','họ và tên','địa chỉ','nơi cư','nơi sinh','tên đầy đủ','đường','phường','quận','tỉnh','huyện'].some(x => l.includes(x));
                        };

                        const renderSectionFields = (sectionFields, getValue, makeOnChange) => {
                          const rows = [];
                          let i = 0;
                          while (i < sectionFields.length) {
                            const f = sectionFields[i];
                            const next = sectionFields[i + 1];
                            const canPair = f.field_type !== 'boolean' && !isFullWidthField(f)
                              && next && next.field_type !== 'boolean' && !isFullWidthField(next);
                            if (canPair) {
                              rows.push(
                                <div key={`r${i}`} className="cust-row2">
                                  {renderCustField(f, getValue(f), makeOnChange(f))}
                                  {renderCustField(next, getValue(next), makeOnChange(next))}
                                </div>
                              );
                              i += 2;
                            } else {
                              rows.push(renderCustField(f, getValue(f), makeOnChange(f)));
                              i++;
                            }
                          }
                          return rows;
                        };

                        const hasLabelFields = stepFields.some(f => f.field_type === 'label');
                        if (hasLabelFields) {
                          const sections = [];
                          let curSec = { title: null, fields: [] };
                          for (const field of stepFields) {
                            if (field.field_type === 'label') {
                              if (curSec.title !== null || curSec.fields.length > 0) sections.push(curSec);
                              curSec = { title: field.label, fields: [] };
                            } else {
                              curSec.fields.push(field);
                            }
                          }
                          sections.push(curSec);

                          return sections.map((sec, si) => (
                            <div key={si} className="cust-section">
                              {sec.title && (
                                <div className="cust-section-header">
                                  <span style={{ width: 5, height: 5, background: '#b8924a', borderRadius: '50%', display: 'inline-block', flexShrink: 0 }} />
                                  {sec.title}
                                </div>
                              )}
                              <div className="cust-section-body">
                                {renderSectionFields(
                                  sec.fields,
                                  (f) => formData[f.key_name],
                                  (f) => (v) => setFormData({ ...formData, [f.key_name]: v })
                                )}
                              </div>
                            </div>
                          ));
                        }

                        return Object.entries(getGroupedFields(stepFields)).sort(([a], [b]) => a.localeCompare(b)).map(([groupName, groupFields]) => {
                          const firstField = groupFields[0];
                          const childTemp = firstField?.childTemplateId
                            ? linkedChildren.find(c => c.id === firstField.childTemplateId)
                            : null;
                          const isRepeated = childTemp?.is_repeated === 1;

                          const displayGroupName = groupName.startsWith('Biểu mẫu con:')
                            ? groupName.substring(13).trim()
                            : (groupName.match(/^\d+\./) ? groupName.substring(3) : groupName);

                          if (isRepeated) {
                            const childId = childTemp.id;
                            const records = formData[childId] || [];
                            return (
                              <div key={groupName} className="cust-section">
                                <div className="cust-section-body">
                                  {records.map((record, rIdx) => (
                                    <div key={record._id || rIdx} style={{ background: '#f7f9fb', border: '1px solid #e6e8ea', borderRadius: 6, padding: '10px 12px', display: 'flex', flexDirection: 'column', gap: 8 }}>
                                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 2 }}>
                                        <span style={{ fontSize: 10, fontWeight: 700, color: '#76777d', textTransform: 'uppercase', letterSpacing: '0.06em' }}>
                                          Bản ghi #{rIdx + 1}
                                        </span>
                                        {records.length > 1 && (
                                          <button type="button" className="lx-btn lx-btn-danger lx-btn-sm"
                                            onClick={() => handleRemoveRepeatedRecord(childId, rIdx)}>
                                            Xóa
                                          </button>
                                        )}
                                      </div>
                                      {renderSectionFields(
                                        groupFields,
                                        (f) => record[f.key_name],
                                        (f) => (v) => handleUpdateRepeatedField(childId, rIdx, f.key_name, v)
                                      )}
                                    </div>
                                  ))}
                                  <button type="button" className="lx-btn lx-btn-secondary"
                                    style={{ width: '100%', justifyContent: 'center', borderStyle: 'dashed' }}
                                    onClick={() => handleAddRepeatedRecord(childId)}>
                                    <span className="material-symbols-outlined" style={{ fontSize: 16 }}>add</span>
                                    Thêm bản ghi mới
                                  </button>
                                </div>
                              </div>
                            );
                          }

                          return (
                            <div key={groupName} className="cust-section">
                              <div className="cust-section-body">
                                {renderSectionFields(
                                  groupFields,
                                  (f) => formData[f.key_name],
                                  (f) => (v) => setFormData({ ...formData, [f.key_name]: v })
                                )}
                              </div>
                            </div>
                          );
                        });
                      })()}
                    </div>

                    {/* Navigation buttons */}
                    <div style={{ display: 'flex', justifyContent: 'space-between', gap: 10, paddingTop: 4, paddingBottom: 8 }}>
                      {(() => {
                        const currentIdx = activeSteps.findIndex(s => s.step === currentStep);
                        return currentIdx > 0 ? (
                          <button
                            type="button"
                            className="lx-btn lx-btn-secondary"
                            onClick={() => setCurrentStep(activeSteps[currentIdx - 1].step)}
                          >
                            <span className="material-symbols-outlined" style={{ fontSize: 16 }}>arrow_back</span>
                            Quay lại
                          </button>
                        ) : (
                          <button
                            type="button"
                            className="lx-btn lx-btn-ghost"
                            onClick={() => {
                              if (window.confirm("Bạn có chắc muốn hủy hồ sơ và quay lại? Dữ liệu đã nhập sẽ bị mất.")) {
                                setActiveView('dashboard');
                              }
                            }}
                          >
                            Hủy hồ sơ
                          </button>
                        );
                      })()}

                      {(() => {
                        const currentIdx = activeSteps.findIndex(s => s.step === currentStep);
                        const isLastStep = currentIdx === activeSteps.length - 1;
                        return isLastStep ? (
                          <button
                            type="button"
                            className="lx-btn lx-btn-primary"
                            onClick={() => { setPreviewActiveDocIdx(0); setActiveView('preview'); }}
                          >
                            <span className="material-symbols-outlined" style={{ fontSize: 16 }}>visibility</span>
                            Xem trước hồ sơ
                          </button>
                        ) : (
                          <button
                            type="button"
                            className="lx-btn lx-btn-primary"
                            onClick={() => setCurrentStep(activeSteps[currentIdx + 1].step)}
                          >
                            Tiếp theo
                            <span className="material-symbols-outlined" style={{ fontSize: 16 }}>arrow_forward</span>
                          </button>
                        );
                      })()}
                    </div>
                  </form>
                </div>

                {/* Submission progress overlay */}
                {isSubmitting && selectedTemplate && (() => {
                  const childCount = selectedTemplate.children_count || 0;
                  const { current, total, message } = submissionProgress;
                  const percentage = total > 0 ? Math.round((current / total) * 100) : 0;

                  const getStepState = (step) => {
                    if (step === 1) {
                      if (current > 1) return 'completed';
                      if (current === 1) return 'processing';
                      return 'pending';
                    }
                    if (step === 2) {
                      if (current > 2) return 'completed';
                      if (current === 2) return 'processing';
                      return 'pending';
                    }
                    if (step === 3) {
                      if (childCount === 0) return 'hidden';
                      if (current > 2 + childCount) return 'completed';
                      if (current >= 3 && current <= 2 + childCount) return 'processing';
                      return 'pending';
                    }
                    if (step === 4) {
                      const finalizeStep = childCount > 0 ? 3 + childCount : 3;
                      if (current >= total && total > 0) return 'completed';
                      if (current >= finalizeStep) return 'processing';
                      return 'pending';
                    }
                    return 'pending';
                  };

                  const renderProgressIcon = (state, num) => {
                    if (state === 'completed') {
                      return (
                        <span style={{
                          width: 22, height: 22, borderRadius: '50%',
                          background: '#6ffbbe', color: '#002113',
                          display: 'flex', alignItems: 'center', justifyContent: 'center',
                          flexShrink: 0, fontSize: 11, fontWeight: 700
                        }}>
                          <span className="material-symbols-outlined" style={{ fontSize: 13 }}>check</span>
                        </span>
                      );
                    }
                    if (state === 'processing') {
                      return (
                        <span style={{
                          width: 22, height: 22, borderRadius: '50%',
                          background: '#e6e8ea', color: '#191c1e',
                          display: 'flex', alignItems: 'center', justifyContent: 'center',
                          flexShrink: 0
                        }}>
                          <span className="lx-spinner" style={{ width: 12, height: 12, borderWidth: 2 }} />
                        </span>
                      );
                    }
                    return (
                      <span style={{
                        width: 22, height: 22, borderRadius: '50%',
                        background: '#eceef0', color: '#76777d', border: '1px solid #c6c6cd',
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                        flexShrink: 0, fontSize: 10, fontWeight: 700
                      }}>
                        {num}
                      </span>
                    );
                  };

                  const step1State = getStepState(1);
                  const step2State = getStepState(2);
                  const step3State = getStepState(3);
                  const step4State = getStepState(4);

                  return (
                    <div className="lx-modal-overlay">
                      <div className="lx-modal" style={{ width: 360 }}>
                        <div className="lx-modal-header">
                          <span className="lx-modal-title">Đang xử lý hồ sơ</span>
                        </div>
                        <div className="lx-modal-body" style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
                          <div style={{ textAlign: 'center', fontSize: 13, color: '#45464d' }}>
                            {message || 'Vui lòng đợi trong giây lát...'}
                          </div>

                          <div>
                            <div style={{ height: 6, background: '#e6e8ea', borderRadius: 9999, overflow: 'hidden' }}>
                              <div style={{ width: `${percentage}%`, height: '100%', background: '#000000', borderRadius: 9999, transition: 'width 0.5s ease' }} />
                            </div>
                            <div style={{ textAlign: 'center', fontSize: 12, fontWeight: 600, color: '#000000', marginTop: 6 }}>{percentage}%</div>
                          </div>

                          <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
                            {[
                              { state: step1State, num: '1', title: 'Chuẩn bị dữ liệu', desc: 'Kiểm tra thông tin & biểu mẫu hợp lệ' },
                              { state: step2State, num: '2', title: 'Điền dữ liệu file gốc', desc: `${selectedTemplate.name}.docx` },
                              ...(step3State !== 'hidden' ? [{ state: step3State, num: '3', title: `Điền dữ liệu ${childCount} file con`, desc: 'Các biểu mẫu con đã liên kết' }] : []),
                              { state: step4State, num: childCount > 0 ? '4' : '3', title: 'Đóng gói & Hoàn tất', desc: childCount > 0 ? 'Tạo tệp lưu trữ .zip' : 'Tạo tệp Word .docx' }
                            ].map((s, i) => (
                              <div key={i} style={{ display: 'flex', gap: 12, alignItems: 'flex-start', opacity: s.state === 'pending' ? 0.45 : 1, transition: 'opacity 0.3s' }}>
                                {renderProgressIcon(s.state, s.num)}
                                <div>
                                  <div style={{ fontSize: 13, fontWeight: 600, color: s.state === 'processing' ? '#000000' : '#191c1e' }}>{s.title}</div>
                                  <div style={{ fontSize: 11, color: '#76777d' }}>{s.desc}</div>
                                </div>
                              </div>
                            ))}
                          </div>

                          <div style={{ fontSize: 11, color: '#76777d', textAlign: 'center', fontStyle: 'italic' }}>
                            Vui lòng không tắt trình duyệt hoặc làm mất kết nối mạng.
                          </div>
                        </div>
                      </div>
                    </div>
                  );
                })()}
              </div>

              {/* Resizer */}
              {!isMobile && (
                <div
                  className="lx-split-divider"
                  onMouseDown={startResizing}
                  title="Kéo để thay đổi kích thước"
                />
              )}

              {/* Right panel: Live Preview */}
              <div
                style={{
                  width: isMobile ? '100%' : `${100 - leftWidth}%`,
                  display: isMobile && mobileFillTab !== 'preview' ? 'none' : 'flex',
                  flexDirection: 'column',
                  overflow: 'hidden',
                  background: '#f7f9fb'
                }}
              >
                {/* Document tabs */}
                {linkedChildren && linkedChildren.length > 0 && (
                  <div style={{ display: 'flex', gap: 0, background: '#ffffff', borderBottom: '1px solid #e6e8ea', overflowX: 'auto', flexShrink: 0 }}>
                    <button
                      type="button"
                      style={{
                        padding: '9px 16px', fontSize: 12, fontWeight: 600,
                        borderBottom: previewTemplateId === selectedTemplate.id ? '2px solid #000000' : '2px solid transparent',
                        color: previewTemplateId === selectedTemplate.id ? '#000000' : '#76777d',
                        background: 'transparent', border: 'none', cursor: 'pointer',
                        borderBottomWidth: 2, borderBottomStyle: 'solid',
                        borderBottomColor: previewTemplateId === selectedTemplate.id ? '#000000' : 'transparent',
                        whiteSpace: 'nowrap'
                      }}
                      onClick={() => handleSwitchPreview(selectedTemplate.id)}
                    >
                      G — {selectedTemplate.name}
                    </button>
                    {linkedChildren.filter(child => selectedChildIds.includes(child.id)).map((child, idx) => (
                      <button
                        key={child.id}
                        type="button"
                        style={{
                          padding: '9px 16px', fontSize: 12, fontWeight: 600,
                          borderBottomWidth: 2, borderBottomStyle: 'solid',
                          borderBottomColor: previewTemplateId === child.id ? '#000000' : 'transparent',
                          color: previewTemplateId === child.id ? '#000000' : '#76777d',
                          background: 'transparent', border: 'none', cursor: 'pointer',
                          whiteSpace: 'nowrap'
                        }}
                        onClick={() => handleSwitchPreview(child.id)}
                      >
                        {idx + 1} — {child.name}
                      </button>
                    ))}
                  </div>
                )}
                <div style={{ flex: 1, overflowY: 'auto' }}>
                  <DocxPreview
                    fileUrl={`http://localhost:5000/api/templates/${previewTemplateId || selectedTemplate.id}/download-original?t=${previewKey}`}
                    title={templates.find(t => t.id === previewTemplateId)?.name || selectedTemplate.name}
                    liveData={getResolvedPreviewData()}
                    fields={previewFields && previewFields.length > 0 ? previewFields : fields}
                  />
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ── VIEW: SUCCESS ── */}
        {activeView === 'success' && submissionResult && (
          <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', position: 'relative' }}>
            {/* Mobile tab switcher */}
            {isMobile && (
              <div className="lx-tabs" style={{ margin: '8px' }}>
                <button
                  type="button"
                  className={`lx-tab ${mobileSuccessTab === 'success' ? 'active' : ''}`}
                  style={{ flex: 1 }}
                  onClick={() => setMobileSuccessTab('success')}
                >
                  Biên nhận
                </button>
                <button
                  type="button"
                  className={`lx-tab ${mobileSuccessTab === 'preview' ? 'active' : ''}`}
                  style={{ flex: 1 }}
                  onClick={() => setMobileSuccessTab('preview')}
                >
                  Xem kết quả
                </button>
              </div>
            )}

            <div ref={splitContainerRef} className="lx-split" style={{ flex: 1 }}>
              {/* Left panel: Receipt */}
              <div
                style={{
                  width: isMobile ? '100%' : `${leftWidth}%`,
                  display: isMobile && mobileSuccessTab !== 'success' ? 'none' : 'flex',
                  flexDirection: 'column',
                  overflowY: 'auto',
                  flexShrink: 0,
                  background: '#ffffff',
                  borderRight: isMobile ? 'none' : '1px solid #e6e8ea'
                }}
              >
                <div style={{ padding: '28px 24px', display: 'flex', flexDirection: 'column', gap: 20 }}>
                  {/* Success header */}
                  <div style={{ textAlign: 'center' }}>
                    <div style={{
                      width: 56, height: 56, borderRadius: '50%',
                      background: '#6ffbbe', border: '2px solid #4edea3',
                      display: 'flex', alignItems: 'center', justifyContent: 'center',
                      margin: '0 auto 14px'
                    }}>
                      <span className="material-symbols-outlined ms-fill" style={{ fontSize: 28, color: '#002113' }}>check_circle</span>
                    </div>
                    <div style={{ fontSize: 18, fontWeight: 700, color: '#000000', letterSpacing: '-0.01em' }}>
                      Soạn thảo hồ sơ hoàn tất
                    </div>
                    <div style={{ fontSize: 13, color: '#45464d', marginTop: 6 }}>
                      Bộ tài liệu đã được tự động soạn thảo và đóng gói.
                    </div>
                  </div>

                  {/* Receipt details */}
                  <div className="lx-card">
                    <div className="lx-card-header">
                      <span className="lx-card-title" style={{ fontSize: 13 }}>Chi tiết hồ sơ</span>
                    </div>
                    <div style={{ padding: '12px 16px', display: 'flex', flexDirection: 'column', gap: 8 }}>
                      {[
                        { label: 'Biểu mẫu', value: submissionResult.templateName },
                        { label: 'Khách hàng', value: submissionResult.customerName },
                        { label: 'Số điện thoại', value: submissionResult.customerPhone }
                      ].map(row => (
                        <div key={row.label} style={{ display: 'flex', justifyContent: 'space-between', gap: 12, fontSize: 13 }}>
                          <span style={{ color: '#76777d' }}>{row.label}:</span>
                          <span style={{ fontWeight: 600, color: '#191c1e', textAlign: 'right', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 200 }}>{row.value}</span>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* File list */}
                  <div className="lx-card">
                    <div className="lx-card-header">
                      <span className="lx-card-title" style={{ fontSize: 13 }}>Các tệp tin đã tạo</span>
                      {submissionResult.files && submissionResult.files.length > 1 && (
                        <button
                          type="button"
                          className="lx-btn lx-btn-ghost lx-btn-sm"
                          onClick={() => {
                            if (successSelectedFiles.length === submissionResult.files.length) {
                              setSuccessSelectedFiles([]);
                            } else {
                              setSuccessSelectedFiles([...submissionResult.files]);
                            }
                          }}
                        >
                          {successSelectedFiles.length === submissionResult.files.length ? 'Bỏ chọn hết' : 'Chọn tất cả'}
                        </button>
                      )}
                    </div>
                    <div style={{ padding: '8px 8px', display: 'flex', flexDirection: 'column', gap: 4, maxHeight: 220, overflowY: 'auto' }}>
                      {submissionResult.files && submissionResult.files.map(filename => {
                        const isChecked = successSelectedFiles.includes(filename);
                        const isMaster = filename === `${selectedTemplate.name}.docx`;
                        const isRepeated = filename.includes('_Căn_');
                        const isPreview = successPreviewFilename === filename;

                        let badgeClass = 'lx-badge-draft';
                        let badgeLabel = 'Đồng bộ';
                        if (isMaster) { badgeClass = 'lx-badge-published'; badgeLabel = 'Gốc'; }
                        else if (isRepeated) { badgeClass = 'lx-badge-warn'; badgeLabel = 'Lặp'; }

                        return (
                          <div
                            key={filename}
                            style={{
                              display: 'flex', alignItems: 'center', gap: 8,
                              padding: '7px 10px', borderRadius: 4,
                              background: isPreview ? '#f2f4f6' : 'transparent',
                              border: isPreview ? '1px solid #c6c6cd' : '1px solid transparent',
                              cursor: 'pointer'
                            }}
                          >
                            <input
                              type="checkbox"
                              checked={isChecked}
                              style={{ accentColor: '#000000', flexShrink: 0 }}
                              onChange={e => {
                                if (e.target.checked) {
                                  setSuccessSelectedFiles(prev => [...prev, filename]);
                                } else {
                                  setSuccessSelectedFiles(prev => prev.filter(f => f !== filename));
                                }
                              }}
                            />
                            <div
                              style={{ flex: 1, minWidth: 0 }}
                              onClick={() => setSuccessPreviewFilename(filename)}
                            >
                              <span style={{ fontSize: 12, fontWeight: 600, color: '#191c1e', display: 'block', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={filename}>
                                {filename}
                              </span>
                            </div>
                            <span className={`lx-badge ${badgeClass}`} style={{ fontSize: 10, flexShrink: 0 }}>{badgeLabel}</span>
                            <a
                              href={`http://localhost:5000/api/submissions/${submissionResult.submissionId}/download-file?filename=${encodeURIComponent(filename)}`}
                              download
                              className="lx-icon-btn-sm"
                              title="Tải nhanh tệp"
                              onClick={e => e.stopPropagation()}
                            >
                              <span className="material-symbols-outlined" style={{ fontSize: 16 }}>download</span>
                            </a>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {/* Action buttons */}
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                    <button
                      type="button"
                      className="lx-btn lx-btn-primary"
                      style={{ width: '100%', justifyContent: 'center' }}
                      onClick={handleDownloadSelectedFiles}
                    >
                      <span className="material-symbols-outlined" style={{ fontSize: 16 }}>download</span>
                      Tải các tệp đã chọn
                    </button>
                    <button
                      type="button"
                      className="lx-btn lx-btn-secondary"
                      style={{ width: '100%', justifyContent: 'center' }}
                      onClick={() => setActiveView('dashboard')}
                    >
                      <span className="material-symbols-outlined" style={{ fontSize: 16 }}>arrow_back</span>
                      Quay lại Bảng điều khiển
                    </button>
                  </div>
                </div>
              </div>

              {/* Resizer */}
              {!isMobile && (
                <div
                  className="lx-split-divider"
                  onMouseDown={startResizing}
                  title="Kéo để thay đổi kích thước"
                />
              )}

              {/* Right panel: Final document preview */}
              <div
                style={{
                  width: isMobile ? '100%' : `${100 - leftWidth}%`,
                  display: isMobile && mobileSuccessTab !== 'preview' ? 'none' : 'flex',
                  flexDirection: 'column',
                  overflow: 'hidden',
                  background: '#f7f9fb'
                }}
              >
                {/* Document tabs */}
                {submissionResult.files && submissionResult.files.length > 1 && (
                  <div style={{ display: 'flex', gap: 0, background: '#ffffff', borderBottom: '1px solid #e6e8ea', overflowX: 'auto', flexShrink: 0 }}>
                    {submissionResult.files.map((filename, idx) => (
                      <button
                        key={filename}
                        type="button"
                        style={{
                          padding: '9px 16px', fontSize: 12, fontWeight: 600,
                          borderBottomWidth: 2, borderBottomStyle: 'solid',
                          borderBottomColor: successPreviewFilename === filename ? '#000000' : 'transparent',
                          color: successPreviewFilename === filename ? '#000000' : '#76777d',
                          background: 'transparent', border: 'none', cursor: 'pointer',
                          whiteSpace: 'nowrap'
                        }}
                        onClick={() => setSuccessPreviewFilename(filename)}
                      >
                        {idx + 1} — {filename}
                      </button>
                    ))}
                  </div>
                )}
                <div style={{ flex: 1, overflowY: 'auto' }}>
                  <DocxPreview
                    fileUrl={`http://localhost:5000/api/submissions/${submissionResult.submissionId}/download-file?filename=${encodeURIComponent(successPreviewFilename || `${selectedTemplate.name}.docx`)}`}
                    title={`${successPreviewFilename || selectedTemplate.name} (Bản hoàn thiện)`}
                    fields={[]}
                    liveData={{}}
                  />
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

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
          <div style={{
            position: 'fixed', inset: 0, zIndex: 50,
            background: 'rgba(15,23,42,0.85)',
            display: 'flex', flexDirection: 'column',
            overflow: 'hidden'
          }}>
            {/* Modal header */}
            <div style={{
              background: '#ffffff', borderBottom: '1px solid #c6c6cd',
              padding: '14px 24px',
              display: 'flex', alignItems: 'center', justifyContent: 'space-between',
              flexShrink: 0
            }}>
              <div>
                <div style={{ fontSize: 15, fontWeight: 700, color: '#000000', display: 'flex', alignItems: 'center', gap: 8 }}>
                  <span className="material-symbols-outlined" style={{ fontSize: 20 }}>visibility</span>
                  Xem trước bộ tài liệu ({totalPages} bản)
                </div>
                <div style={{ fontSize: 11, color: '#76777d', marginTop: 2 }}>Tự động điền dữ liệu &amp; sẵn sàng in ấn</div>
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                <button
                  type="button"
                  className="lx-btn lx-btn-secondary lx-btn-sm"
                  onClick={() => window.print()}
                >
                  <span className="material-symbols-outlined" style={{ fontSize: 15 }}>print</span>
                  In / Xuất PDF
                </button>
                <button
                  type="button"
                  className="lx-btn lx-btn-ghost lx-btn-sm"
                  onClick={() => setShowPdfModal(false)}
                >
                  <span className="material-symbols-outlined" style={{ fontSize: 15 }}>close</span>
                  Đóng
                </button>
              </div>
            </div>

            {/* Modal content */}
            <div style={{
              flex: 1, overflowY: 'auto',
              padding: '32px 24px',
              display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 32,
              background: 'rgba(15,23,42,0.3)'
            }}>
              {/* Master doc */}
              <div style={{ width: '100%', maxWidth: 860 }}>
                <div style={{ fontSize: 10, fontWeight: 700, color: '#c6c6cd', textTransform: 'uppercase', letterSpacing: '0.1em', textAlign: 'center', marginBottom: 8 }}>
                  Tài liệu 1/{totalPages} — {selectedTemplate.name} (Gốc)
                </div>
                <div style={{ background: '#ffffff', padding: '48px', boxShadow: '0 8px 32px rgba(0,0,0,0.2)', borderRadius: 4 }}>
                  <DocxPreviewInModal
                    fileUrl={`http://localhost:5000/api/templates/${selectedTemplate.id}/download-original?t=${previewKey}`}
                    liveData={formData}
                    fields={fields.filter(f => !f.groupName)}
                  />
                </div>
              </div>

              {/* Child docs */}
              {(() => {
                const pages = [];
                let currentPage = 2;

                activeChildren.forEach(child => {
                  if (child.is_repeated) {
                    const records = formData[child.id] || [];
                    records.forEach((record, rIdx) => {
                      const pageNum = currentPage++;
                      pages.push(
                        <div key={`${child.id}-${rIdx}`} style={{ width: '100%', maxWidth: 860 }}>
                          <div style={{ fontSize: 10, fontWeight: 700, color: '#c6c6cd', textTransform: 'uppercase', letterSpacing: '0.1em', textAlign: 'center', marginBottom: 8 }}>
                            Tài liệu {pageNum}/{totalPages} — {child.name} (Bản ghi {rIdx + 1})
                          </div>
                          <div style={{ background: '#ffffff', padding: '48px', boxShadow: '0 8px 32px rgba(0,0,0,0.2)', borderRadius: 4 }}>
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
                      <div key={child.id} style={{ width: '100%', maxWidth: 860 }}>
                        <div style={{ fontSize: 10, fontWeight: 700, color: '#c6c6cd', textTransform: 'uppercase', letterSpacing: '0.1em', textAlign: 'center', marginBottom: 8 }}>
                          Tài liệu {pageNum}/{totalPages} — {child.name}
                        </div>
                        <div style={{ background: '#ffffff', padding: '48px', boxShadow: '0 8px 32px rgba(0,0,0,0.2)', borderRadius: 4 }}>
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
    </div>
  );
}
