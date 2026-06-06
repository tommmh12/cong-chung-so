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
  getFieldStep,
  generateId,
  getLogicalGroupName
} from '../utils/helpers';

export default function AdminView() {
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [sidebarActiveMenu, setSidebarActiveMenu] = useState('dashboard'); // dashboard | templates | submissions
  const [activeView, setActiveView] = useState('dashboard'); // dashboard | config

  // Data State
  const [templates, setTemplates] = useState([]);
  const [categories, setCategories] = useState([]);
  const [submissions, setSubmissions] = useState([]);
  const [selectedTemplate, setSelectedTemplate] = useState(null);
  const [fields, setFields] = useState([]);
  const [parentFields, setParentFields] = useState([]);
  const [linkedChildren, setLinkedChildren] = useState([]);
  const [activeConfigTab, setActiveConfigTab] = useState(null); // null = parent tab, or child template object
  const [parentTabFields, setParentTabFields] = useState([]); // parent's fields saved when editing child tab
  const [tabFieldsCache, setTabFieldsCache] = useState({}); // { [templateId]: fields[] } — preserves unsaved changes across tab switches

  // Undo/History State
  const [fieldsHistory, setFieldsHistory] = useState([]);
  const [historyIndex, setHistoryIndex] = useState(-1);
  const [hasUnsavedManual, setHasUnsavedManual] = useState(false);

  // Search & Filter State
  const [dashboardSearch, setDashboardSearch] = useState('');
  const [selectedCategoryId, setSelectedCategoryId] = useState('all');
  const [expandedCategoryIds, setExpandedCategoryIds] = useState(new Set());
  const [sortBy, setSortBy] = useState('newest');
  const [templateFilter, setTemplateFilter] = useState('all'); // all | active | draft

  // Text selection tagging state
  const [selectedText, setSelectedText] = useState('');
  const [paragraphContext, setParagraphContext] = useState('');
  const [quickKey, setQuickKey] = useState('');
  const [quickLabel, setQuickLabel] = useState('');
  const [quickType, setQuickType] = useState('text');
  const [quickInheritMode, setQuickInheritMode] = useState(false);
  const [quickParentFieldKey, setQuickParentFieldKey] = useState('');
  const [showManualAdd, setShowManualAdd] = useState(false);
  const [showLabelAdd, setShowLabelAdd] = useState(false);
  const [labelTitle, setLabelTitle] = useState('');
  const [draggedIdx, setDraggedIdx] = useState(null);
  const [dragOverIdx, setDragOverIdx] = useState(null);

  // Modal Dialogs
  const [showUploadModal, setShowUploadModal] = useState(false);
  const [showLinkModal, setShowLinkModal] = useState(false);
  const [linkingTemplate, setLinkingTemplate] = useState(null);
  const [linkSearch, setLinkSearch] = useState('');

  const [viewingSubmission, setViewingSubmission] = useState(null);
  const [expandedSubmissionIds, setExpandedSubmissionIds] = useState(new Set());
  const [onlinePreviewFile, setOnlinePreviewFile] = useState(null);
  const [activeSubmissionPreview, setActiveSubmissionPreview] = useState({});

  // Upload fields
  const [uploadName, setUploadName] = useState('');
  const [uploadDesc, setUploadDesc] = useState('');
  const [uploadCategoryId, setUploadCategoryId] = useState('');
  const [uploadFile, setUploadFile] = useState(null);
  const [isUploading, setIsUploading] = useState(false);

  // Loading & Action State
  const [isLoading, setIsLoading] = useState(false);
  const [notification, setNotification] = useState(null);

  // Resizer state
  const [leftWidth, setLeftWidth] = useState(30);
  const [isMobile, setIsMobile] = useState(false);
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

  // Fetch initial stats
  const fetchData = async () => {
    setIsLoading(true);
    try {
      const [tRes, cRes, sRes] = await Promise.all([
        fetch(`${API_BASE}/templates`),
        fetch(`${API_BASE}/categories`),
        fetch(`${API_BASE}/submissions`)
      ]);
      if (tRes.ok) {
        const tData = await tRes.json();
        setTemplates(tData);
      }
      if (cRes.ok) {
        const cData = await cRes.json();
        setCategories(cData);
      }
      if (sRes.ok) {
        const sData = await sRes.json();
        setSubmissions(sData);
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

  // Text selection change listener for admin tagging
  useEffect(() => {
    if (activeView !== 'config') return;

    const handleSelection = () => {
      const selection = window.getSelection();
      const text = selection ? selection.toString().trim() : '';
      const container = document.querySelector('.docx-container');

      if (text && container && container.contains(selection.anchorNode)) {
        let parentElement = selection.anchorNode.parentElement;
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
  }, [activeView]);

  // History Helper functions
  const updateFieldsAndHistory = (newFields) => {
    const updatedHistory = fieldsHistory.slice(0, historyIndex + 1);
    updatedHistory.push(newFields);

    setFields(newFields);
    setFieldsHistory(updatedHistory);
    setHistoryIndex(updatedHistory.length - 1);
    setHasUnsavedManual(true);
  };

  const handleUndo = () => {
    if (historyIndex > 0) {
      const prevIndex = historyIndex - 1;
      setFields(fieldsHistory[prevIndex]);
      setHistoryIndex(prevIndex);
      showNotification('Đã hoàn tác thay đổi gần nhất!');
    }
  };

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

  const getFlattenedCategoryOptions = (parentId = null, depth = 0) => {
    let options = [];
    categories.filter(c => c.parent_id === parentId).forEach(cat => {
      options.push({
        id: cat.id,
        label: `${'—'.repeat(depth)} ${cat.name}`
      });
      options = [...options, ...getFlattenedCategoryOptions(cat.id, depth + 1)];
    });
    return options;
  };

  // Upload logic
  const handleUploadTemplate = async (e) => {
    e.preventDefault();
    if (!uploadFile || !uploadName) {
      showNotification('Vui lòng nhập tên biểu mẫu và chọn tệp', 'error');
      return;
    }

    setIsUploading(true);
    const formData = new FormData();
    formData.append('templateFile', uploadFile);
    formData.append('name', uploadName);
    formData.append('description', uploadDesc);
    if (uploadCategoryId) {
      formData.append('categoryId', uploadCategoryId);
    }

    try {
      const res = await fetch(`${API_BASE}/templates`, {
        method: 'POST',
        body: formData
      });

      const data = await res.json();
      if (!res.ok) {
        const failedResult = data.results?.find(r => r.status === 'failed');
        const errMsg = failedResult?.error || data.message || 'Không thể tải lên biểu mẫu';
        throw new Error(errMsg);
      }

      showNotification('Đã tải lên và phân tích biểu mẫu thành công!');
      setShowUploadModal(false);
      setUploadName('');
      setUploadDesc('');
      setUploadCategoryId('');
      setUploadFile(null);
      fetchData();
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsUploading(false);
    }
  };

  // Open configuration panel
  const handleOpenConfig = async (templateId) => {
    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${templateId}/form`);
      if (!res.ok) {
        throw new Error('Lỗi nạp cấu hình biểu mẫu');
      }
      const data = await res.json();

      setSelectedTemplate(data.template);
      setFields(data.fields);
      setFieldsHistory([data.fields]);
      setHistoryIndex(0);
      setHasUnsavedManual(false);

      if (data.template.parent_template_id) {
        const pRes = await fetch(`${API_BASE}/templates/${data.template.parent_template_id}/form`);
        if (pRes.ok) {
          const pData = await pRes.json();
          setParentFields(pData.fields);
        }
      } else {
        setParentFields([]);
      }

      setActiveView('config');
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleSwitchConfigTab = async (child) => {
    // Save current tab's fields to cache before switching
    const currentId = activeConfigTab ? activeConfigTab.id : selectedTemplate.id;
    const updatedCache = { ...tabFieldsCache, [currentId]: fields };
    setTabFieldsCache(updatedCache);

    setSelectedText('');
    setQuickInheritMode(false);
    setQuickParentFieldKey('');

    if (!child) {
      // Switch back to parent tab — restore from cache or fallback to parentTabFields
      const restoredFields = updatedCache[selectedTemplate.id] ?? parentTabFields;
      setFields(restoredFields);
      setFieldsHistory([restoredFields]);
      setHistoryIndex(0);
      setActiveConfigTab(null);
      setHasUnsavedManual(updatedCache[selectedTemplate.id] !== undefined);
      return;
    }

    // Already visited this child tab — restore from cache
    if (updatedCache[child.id]) {
      if (activeConfigTab === null) setParentTabFields(fields);
      setFields(updatedCache[child.id]);
      setFieldsHistory([updatedCache[child.id]]);
      setHistoryIndex(0);
      setActiveConfigTab(child);
      setHasUnsavedManual(true);
      return;
    }

    // First visit to this child tab — fetch from server
    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${child.id}/form`);
      if (!res.ok) throw new Error('Không thể tải cấu hình phụ lục');
      const data = await res.json();
      if (activeConfigTab === null) setParentTabFields(fields);
      setFields(data.fields);
      setFieldsHistory([data.fields]);
      setHistoryIndex(0);
      setActiveConfigTab(child);
      setHasUnsavedManual(false);
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleUpdateTemplateCategory = async (templateId, categoryId) => {
    try {
      await fetch(`${API_BASE}/templates/${templateId}/category`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ categoryId: categoryId || null })
      });
      showNotification('Đã cập nhật danh mục thành công!');
      fetchData();
    } catch (err) {
      showNotification('Lỗi cập nhật danh mục', 'error');
    }
  };

  const handleFieldChange = (idx, prop, val) => {
    const updated = fields.map((f, i) => {
      if (i === idx) {
        const cleanVal = prop === 'is_required' ? (val ? 1 : 0) : val;
        return { ...f, [prop]: cleanVal };
      }
      return f;
    });
    updateFieldsAndHistory(updated);
  };

  // Quick field generator from text tagging
  const handleAddQuickField = (e) => {
    e.preventDefault();

    // Inherit mode: map to a parent field
    if (quickInheritMode) {
      if (!quickParentFieldKey) {
        showNotification('Vui lòng chọn trường từ template cha', 'error');
        return;
      }
      const parentField = parentTabFields.find(f => f.key_name === quickParentFieldKey);
      if (!parentField) return;
      // Auto-suffix key_name if already used (allow multiple inherits of same parent field)
      let inheritKey = parentField.key_name;
      let suffix = 2;
      while (fields.some(f => f.key_name === inheritKey)) {
        inheritKey = `${parentField.key_name}_${suffix++}`;
      }
      const inheritedField = {
        id: generateId(),
        key_name: inheritKey,
        label: parentField.label,
        field_type: parentField.field_type,
        is_required: parentField.is_required,
        order_index: fields.length,
        replace_text: selectedText,
        paragraph_context: paragraphContext,
        parent_field_key: parentField.key_name
      };
      updateFieldsAndHistory([...fields, inheritedField]);
      setSelectedText('');
      setParagraphContext('');
      setQuickInheritMode(false);
      setQuickParentFieldKey('');
      showNotification(`Đã kế thừa trường {{${parentField.key_name}}} từ template cha!`);
      window.getSelection()?.removeAllRanges();
      return;
    }

    if (!quickKey || !quickLabel) {
      showNotification('Vui lòng điền đầy đủ mã biến và nhãn hiển thị', 'error');
      return;
    }

    const cleanKey = quickKey.trim().toLowerCase().replace(/[^a-z0-9_]/g, '_');
    if (fields.some(f => f.key_name === cleanKey)) {
      showNotification('Mã biến này đã tồn tại', 'error');
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
      paragraph_context: paragraphContext,
      parent_field_key: null
    };

    updateFieldsAndHistory([...fields, newField]);
    setSelectedText('');
    setParagraphContext('');
    setQuickKey('');
    setQuickLabel('');
    showNotification(`Đã tạo biến {{${cleanKey}}} từ chữ bôi đen!`);
    window.getSelection()?.removeAllRanges();
  };

  const handleOpenManualAdd = () => {
    setShowManualAdd(true);
    setShowLabelAdd(false);
    setSelectedText('');
    setParagraphContext('');
    setQuickKey('');
    setQuickLabel('');
    setQuickType('text');
  };

  const handleAddManualField = (e) => {
    e.preventDefault();
    if (!quickKey || !quickLabel) {
      showNotification('Vui lòng nhập đầy đủ mã biến và nhãn hiển thị', 'error');
      return;
    }

    const cleanKey = quickKey.trim().toLowerCase().replace(/[^a-z0-9_]/g, '_');
    if (fields.some(f => f.key_name === cleanKey)) {
      showNotification('Mã biến đã tồn tại trong biểu mẫu', 'error');
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
      paragraph_context: null,
      parent_field_key: null
    };

    updateFieldsAndHistory([...fields, newField]);
    setShowManualAdd(false);
    setQuickKey('');
    setQuickLabel('');
    showNotification(`Đã thêm thủ công biến {{${cleanKey}}}`);
  };

  const handleAddLabelField = (e) => {
    e.preventDefault();
    if (!labelTitle.trim()) {
      showNotification('Vui lòng nhập nội dung tiêu đề nhóm', 'error');
      return;
    }
    const key = `label_section_${Date.now()}`;
    const newField = {
      id: generateId(),
      key_name: key,
      label: labelTitle.trim(),
      field_type: 'label',
      is_required: 0,
      order_index: fields.length,
      replace_text: null,
      paragraph_context: null,
      parent_field_key: null
    };
    updateFieldsAndHistory([...fields, newField]);
    setShowLabelAdd(false);
    setLabelTitle('');
    showNotification('Đã thêm tiêu đề nhóm');
  };

  // Save fields configuration
  const handleSaveConfig = async () => {
    setIsLoading(true);
    const saveTargetId = activeConfigTab ? activeConfigTab.id : selectedTemplate.id;
    try {
      const res = await fetch(`${API_BASE}/templates/${saveTargetId}/fields`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields })
      });

      if (!res.ok) {
        throw new Error('Lỗi lưu cấu hình biến');
      }

      showNotification('Đã lưu cấu hình biến và kích hoạt thành công!');
      setHasUnsavedManual(false);
      setTabFieldsCache(prev => { const next = { ...prev }; delete next[saveTargetId]; return next; });
      fetchData();
    } catch (err) {
      showNotification(err.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  // Export JSON configuration
  const handleExportConfig = () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(fields, null, 2));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute("href", dataStr);
    downloadAnchor.setAttribute("download", `config_${selectedTemplate.name.replace(/\s+/g, '_')}.json`);
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();
    showNotification('Đã xuất cấu hình JSON thành công!');
  };

  // Import JSON configuration
  const handleImportConfig = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const parsed = JSON.parse(event.target.result);
        if (!Array.isArray(parsed)) {
          throw new Error('Định dạng cấu hình JSON phải là một mảng.');
        }
        updateFieldsAndHistory(parsed);
        showNotification('Đã nhập cấu hình JSON thành công!');
      } catch (err) {
        showNotification(err.message, 'error');
      }
    };
    reader.readAsText(file);
    e.target.value = null;
  };

  // Delete template
  const handleDeleteTemplate = async (templateId) => {
    if (!window.confirm("Bạn có chắc chắn muốn xóa biểu mẫu này? Tất cả các trường dữ liệu và liên kết cũng sẽ bị xóa.")) {
      return;
    }

    try {
      const res = await fetch(`${API_BASE}/templates/${templateId}`, {
        method: 'DELETE'
      });

      if (!res.ok) {
        throw new Error('Không thể xóa biểu mẫu');
      }

      showNotification('Đã xóa biểu mẫu thành công!');
      fetchData();
    } catch (err) {
      showNotification(err.message, 'error');
    }
  };

  // Toggle template active status
  const handleToggleTemplateStatus = async (template) => {
    const nextStatus = template.status === 'active' ? 'draft' : 'active';
    try {
      const res = await fetch(`${API_BASE}/templates/${template.id}/status`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status: nextStatus })
      });

      if (!res.ok) throw new Error();
      showNotification(`Đã chuyển trạng thái biểu mẫu sang ${nextStatus === 'active' ? 'Hoạt động' : 'Bản nháp'}`);
      fetchData();
    } catch (err) {
      showNotification('Không thể cập nhật trạng thái biểu mẫu', 'error');
    }
  };

  // Mapping conflicts finder
  const getMappingConflicts = () => {
    const mappedKeys = fields.map(f => f.parent_field_key).filter(Boolean);
    const duplicates = mappedKeys.filter((item, index) => mappedKeys.indexOf(item) !== index);
    return new Set(duplicates);
  };

  // Linking child template modal helper
  const handleOpenLinkModal = async (template) => {
    setLinkingTemplate(template);
    setLinkSearch('');
    setIsLoading(true);
    try {
      const res = await fetch(`${API_BASE}/templates/${template.id}/links`);
      if (res.ok) {
        const data = await res.json();
        setLinkedChildren(data);
      }
      setShowLinkModal(true);
    } catch (err) {
      showNotification('Không thể tải danh sách tài liệu liên kết', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleLinkChild = async (childId) => {
    if (!linkingTemplate) return;
    try {
      const res = await fetch(`${API_BASE}/templates/${linkingTemplate.id}/link`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ childTemplateId: childId })
      });

      if (!res.ok) {
        throw new Error('Không thể liên kết biểu mẫu phụ');
      }

      showNotification('Đã liên kết biểu mẫu thành công!');

      // Refresh list
      const r = await fetch(`${API_BASE}/templates/${linkingTemplate.id}/links`);
      if (r.ok) {
        const data = await r.json();
        setLinkedChildren(data);
      }
      fetchData();
    } catch (err) {
      showNotification(err.message, 'error');
    }
  };

  const handleUnlinkChild = async (childId) => {
    if (!linkingTemplate) return;
    try {
      const res = await fetch(`${API_BASE}/templates/${childId}/unlink`, {
        method: 'POST'
      });

      if (!res.ok) {
        throw new Error('Không thể hủy liên kết');
      }

      showNotification('Đã hủy liên kết biểu mẫu!');

      // Refresh list
      const r = await fetch(`${API_BASE}/templates/${linkingTemplate.id}/links`);
      if (r.ok) {
        const data = await r.json();
        setLinkedChildren(data);
      }
      fetchData();
    } catch (err) {
      showNotification(err.message, 'error');
    }
  };

  const handleToggleRepeatedLink = async (childId, currentRepeatedVal) => {
    if (!linkingTemplate) return;
    const nextVal = currentRepeatedVal === 1 ? 0 : 1;
    try {
      const res = await fetch(`${API_BASE}/templates/${childId}/repeated`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ isRepeated: nextVal })
      });

      if (!res.ok) {
        throw new Error('Không thể cập nhật cấu hình lặp');
      }

      showNotification(`Đã chuyển trạng thái cấu hình lặp sang ${nextVal === 1 ? 'Lặp nhiều lần' : 'Duy nhất'}`);

      const r = await fetch(`${API_BASE}/templates/${linkingTemplate.id}/links`);
      if (r.ok) {
        const data = await r.json();
        setLinkedChildren(data);
      }
      fetchData();
    } catch (err) {
      showNotification(err.message, 'error');
    }
  };

  // Filter templates list
  const getFilteredTemplates = () => {
    let list = [...templates];

    if (dashboardSearch.trim()) {
      const query = dashboardSearch.toLowerCase();
      list = list.filter(t => t.name.toLowerCase().includes(query));
    }

    if (selectedCategoryId !== 'all') {
      if (selectedCategoryId === 'uncategorized') {
        list = list.filter(t => !t.category_id);
      } else {
        const filterIds = [selectedCategoryId, ...getDescendantCategoryIds(selectedCategoryId)];
        list = list.filter(t => filterIds.includes(t.category_id));
      }
    }

    if (templateFilter === 'active') {
      list = list.filter(t => t.status === 'active');
    } else if (templateFilter === 'draft') {
      list = list.filter(t => t.status !== 'active');
    }

    if (sortBy === 'newest') {
      list.sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0));
    } else if (sortBy === 'oldest') {
      list.sort((a, b) => new Date(a.created_at || 0) - new Date(b.created_at || 0));
    } else if (sortBy === 'name') {
      list.sort((a, b) => a.name.localeCompare(b.name));
    }

    return list;
  };

  const toggleSubmissionRow = (id) => {
    setExpandedSubmissionIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
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
                  style={{ fontSize: 14 }}
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

  // ── SUBMISSION DETAIL FULL-SCREEN PAGE ──────────────────────────
  if (viewingSubmission) {
    let parsedFiles = [];
    try { parsedFiles = typeof viewingSubmission.files === 'string' ? JSON.parse(viewingSubmission.files) : (viewingSubmission.files || []); } catch (e) {}
    const activeFile = activeSubmissionPreview[viewingSubmission.id] || parsedFiles[0] || '';

    return (
      <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', background: '#f7f9fb', fontFamily: "'Inter', sans-serif" }}>
        {/* Toast */}
        {notification && (
          <div className={`lx-toast ${notification.type === 'error' ? 'error' : 'success'}`}>
            <span className="material-symbols-outlined">{notification.type === 'error' ? 'warning' : 'check_circle'}</span>
            {notification.message}
          </div>
        )}

        {/* ── Top header ── */}
        <header style={{ height: 60, background: '#ffffff', borderBottom: '1px solid #c6c6cd', display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '0 24px', flexShrink: 0, zIndex: 10 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <button className="lx-icon-btn-sm" onClick={() => setViewingSubmission(null)} title="Quay lại danh sách">
              <span className="material-symbols-outlined">arrow_back</span>
            </button>
            <div style={{ width: 1, height: 22, background: '#c6c6cd' }} />
            <div>
              <div style={{ fontSize: 16, fontWeight: 600, color: '#000000', letterSpacing: '-0.01em' }}>Kiểm tra lại hồ sơ</div>
              <div style={{ fontSize: 12, color: '#76777d' }}>
                {viewingSubmission.customer_name} · {viewingSubmission.template_name}
              </div>
            </div>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <span className="lx-badge lx-badge-pending" style={{ fontSize: 11 }}>
              {parsedFiles.length} tài liệu
            </span>
            <button className="lx-btn lx-btn-primary lx-btn-sm">
              <span className="material-symbols-outlined" style={{ fontSize: 15 }}>download</span>
              Tải tất cả
            </button>
          </div>
        </header>

        {/* ── Main workspace ── */}
        <div style={{ flex: 1, display: 'flex', overflow: 'hidden' }}>

          {/* Left sidebar — document tabs */}
          <aside style={{ width: 284, background: '#ffffff', borderRight: '1px solid #c6c6cd', display: 'flex', flexDirection: 'column', flexShrink: 0, overflow: 'hidden' }}>
            <div style={{ padding: '12px 16px', borderBottom: '1px solid #e6e8ea', background: '#f7f9fb', flexShrink: 0 }}>
              <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.07em', color: '#45464d' }}>
                Danh mục tài liệu ({parsedFiles.length})
              </div>
            </div>
            <nav style={{ flex: 1, overflowY: 'auto' }}>
              {parsedFiles.length === 0 && (
                <div className="lx-empty" style={{ padding: 24 }}>Không có tài liệu nào.</div>
              )}
              {parsedFiles.map((file, idx) => {
                const isActive = activeFile === file;
                const shortName = file.replace(/\.docx$/i, '').replace(/_/g, ' ');
                return (
                  <button
                    key={file}
                    onClick={() => setActiveSubmissionPreview({ ...activeSubmissionPreview, [viewingSubmission.id]: file })}
                    style={{
                      width: '100%', display: 'flex', alignItems: 'flex-start', gap: 12,
                      padding: '14px 16px', textAlign: 'left', cursor: 'pointer',
                      background: isActive ? 'rgba(219,226,253,0.25)' : 'transparent',
                      borderLeft: `4px solid ${isActive ? '#000000' : 'transparent'}`,
                      border: 'none', borderBottom: '1px solid #f2f4f6',
                      transition: 'background 0.12s',
                    }}
                    onMouseEnter={e => { if (!isActive) e.currentTarget.style.background = '#f7f9fb'; }}
                    onMouseLeave={e => { if (!isActive) e.currentTarget.style.background = 'transparent'; }}
                  >
                    <span className={`material-symbols-outlined${isActive ? ' ms-fill' : ''}`} style={{ fontSize: 20, color: isActive ? '#000000' : '#45464d', marginTop: 1, flexShrink: 0 }}>description</span>
                    <div style={{ overflow: 'hidden', flex: 1 }}>
                      <div style={{ fontSize: 12, fontWeight: 600, letterSpacing: '0.03em', color: isActive ? '#000000' : '#191c1e', lineHeight: 1.3 }}>
                        {idx + 1}. {shortName.length > 32 ? shortName.slice(0, 32) + '…' : shortName}
                      </div>
                      <div style={{ fontSize: 11, color: '#76777d', marginTop: 3, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        {file}
                      </div>
                    </div>
                  </button>
                );
              })}
            </nav>
          </aside>

          {/* Right — document viewer */}
          <section style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', background: '#e2e5e9' }}>
            {/* Viewer toolbar */}
            <div style={{ height: 50, background: '#ffffff', borderBottom: '1px solid #e6e8ea', display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '0 20px', flexShrink: 0 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <span className="material-symbols-outlined" style={{ fontSize: 16, color: '#45464d' }}>description</span>
                <span style={{ fontSize: 13, fontWeight: 500, color: '#191c1e', maxWidth: 420, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{activeFile}</span>
              </div>
              {activeFile && (
                <a
                  href={`http://localhost:5000/api/submissions/${viewingSubmission.id}/download-file?filename=${encodeURIComponent(activeFile)}`}
                  download
                  className="lx-btn lx-btn-secondary lx-btn-sm"
                >
                  <span className="material-symbols-outlined" style={{ fontSize: 14 }}>download</span>
                  Tải xuống
                </a>
              )}
            </div>

            {/* DocxPreview — fills remaining height */}
            <div style={{ flex: 1, overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
              {activeFile ? (
                <DocxPreview
                  key={activeFile}
                  fileUrl={`http://localhost:5000/api/submissions/${viewingSubmission.id}/download-file?filename=${encodeURIComponent(activeFile)}`}
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
      </div>
    );
  }

  return (
    <div className="lx-app">
      {/* Toast notification */}
      {notification && (
        <div className={`lx-toast ${notification.type === 'error' ? 'error' : 'success'}`}>
          <span className="material-symbols-outlined">
            {notification.type === 'error' ? 'warning' : 'check_circle'}
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
            <p>Management System</p>
          </div>
        </div>

        <div className="lx-sidebar-cta">
          <button className="lx-sidebar-cta-btn" onClick={() => setShowUploadModal(true)}>
            <span className="material-symbols-outlined">add</span>
            Tải lên Template
          </button>
        </div>

        <div className="lx-sidebar-nav">
          <button
            className={`lx-nav-item ${sidebarActiveMenu === 'dashboard' ? 'active' : ''}`}
            onClick={() => {
              setSidebarActiveMenu('dashboard');
              setActiveView('dashboard');
            }}
          >
            <span className="material-symbols-outlined">dashboard</span>
            <span>Dashboard</span>
          </button>
          <button
            className={`lx-nav-item ${sidebarActiveMenu === 'templates' && activeView !== 'config' ? 'active' : ''}`}
            onClick={() => {
              setSidebarActiveMenu('templates');
              setActiveView('dashboard');
            }}
          >
            <span className="material-symbols-outlined">description</span>
            <span>Templates</span>
          </button>
          <button
            className={`lx-nav-item ${sidebarActiveMenu === 'submissions' ? 'active' : ''}`}
            onClick={() => {
              setSidebarActiveMenu('submissions');
              setActiveView('dashboard');
            }}
          >
            <span className="material-symbols-outlined">send</span>
            <span>Submissions</span>
          </button>
        </div>

        <div className="lx-sidebar-bottom">
          <Link to="/" className="lx-nav-item">
            <span className="material-symbols-outlined">person</span>
            <span>Trang Khách Hàng</span>
          </Link>
          <div className="lx-nav-divider" />
          <button className="lx-nav-item">
            <span className="material-symbols-outlined">settings</span>
            <span>Settings</span>
          </button>
        </div>
      </nav>

      {/* Main area */}
      <div className="lx-main">
        {/* Header */}
        <header className="lx-header">
          <div className="lx-search">
            <span className="material-symbols-outlined">search</span>
            <input placeholder="Tìm kiếm hồ sơ, biểu mẫu..." />
          </div>
          <div className="lx-header-right">
            <button className="lx-icon-btn">
              <span className="material-symbols-outlined">notifications</span>
            </button>
            <button className="lx-icon-btn">
              <span className="material-symbols-outlined">business_center</span>
            </button>
            <div className="lx-divider-v" />
            <div className="lx-user">
              <div className="lx-user-info">
                <div className="lx-user-name">Cán bộ Notary</div>
                <div className="lx-user-role">Quản trị viên</div>
              </div>
              <div className="lx-avatar">
                <span className="material-symbols-outlined ms-fill" style={{ fontSize: '34px' }}>account_circle</span>
              </div>
            </div>
          </div>
        </header>

        {/* Content */}
        <div className="lx-content">

          {/* ── VIEW: DASHBOARD ── */}
          {sidebarActiveMenu === 'dashboard' && activeView !== 'config' && (
            <div>
              <div className="lx-page-header">
                <div className="lx-page-title">Bảng điều khiển quản lý</div>
                <div className="lx-page-subtitle">Tổng quan tình hình hoạt động văn phòng công chứng hôm nay.</div>
              </div>

              {/* Stat cards */}
              <div className="lx-stats-grid">
                {/* Card 1: Pending submissions */}
                <div className="lx-stat-card">
                  <div className="lx-stat-card-top">
                    <div className="lx-stat-icon">
                      <span className="material-symbols-outlined">pending_actions</span>
                    </div>
                    <span className="lx-stat-badge warn">Cấp bách</span>
                  </div>
                  <div className="lx-stat-label">Hồ sơ chờ xử lý</div>
                  <div className="lx-stat-value">
                    {submissions.filter(s => s.status === 'pending').length || submissions.length}
                  </div>
                </div>

                {/* Card 2: Active templates */}
                <div className="lx-stat-card">
                  <div className="lx-stat-card-top">
                    <div className="lx-stat-icon">
                      <span className="material-symbols-outlined">description</span>
                    </div>
                  </div>
                  <div className="lx-stat-label">Mẫu đang hoạt động</div>
                  <div className="lx-stat-value">
                    {templates.filter(t => t.status === 'active').length}
                  </div>
                </div>

                {/* Card 3: Completed today */}
                <div className="lx-stat-card">
                  <div className="lx-stat-card-top">
                    <div className="lx-stat-icon success">
                      <span className="material-symbols-outlined">task_alt</span>
                    </div>
                    <span className="lx-stat-badge success">+8% hôm qua</span>
                  </div>
                  <div className="lx-stat-label">Hoàn thành hôm nay</div>
                  <div className="lx-stat-value">0</div>
                </div>

                {/* Card 4: Attention needed */}
                <div className="lx-stat-card lx-stat-error">
                  <div className="lx-stat-card-top">
                    <div className="lx-stat-icon error">
                      <span className="material-symbols-outlined">warning</span>
                    </div>
                  </div>
                  <div className="lx-stat-label">Cần chú ý</div>
                  <div className="lx-stat-value error">
                    {templates.filter(t => t.status !== 'active').length}
                  </div>
                </div>
              </div>

              {/* Recent submissions table */}
              <div className="lx-card">
                <div className="lx-card-header">
                  <span className="lx-card-title">Hồ sơ gần đây</span>
                  <button
                    className="lx-btn lx-btn-ghost lx-btn-sm"
                    onClick={() => { setSidebarActiveMenu('submissions'); setActiveView('dashboard'); }}
                  >
                    Xem tất cả
                    <span className="material-symbols-outlined" style={{ fontSize: 16 }}>arrow_forward</span>
                  </button>
                </div>
                {submissions.length === 0 ? (
                  <div className="lx-empty">
                    {isLoading ? <span className="lx-spinner" /> : 'Chưa có hồ sơ nào được nộp.'}
                  </div>
                ) : (
                  <>
                    <table className="lx-table">
                      <thead>
                        <tr>
                          <th>ID</th>
                          <th>Loại hợp đồng</th>
                          <th>Khách hàng</th>
                          <th>Trạng thái</th>
                          <th>Ngày nộp</th>
                          <th className="right">Thao tác</th>
                        </tr>
                      </thead>
                      <tbody>
                        {submissions.slice(0, 5).map((sub, idx) => (
                          <tr key={sub.id}>
                            <td className="lx-mono">#{sub.id}</td>
                            <td>{sub.template_name}</td>
                            <td>
                              <div style={{ fontWeight: 600 }}>{sub.customer_name}</div>
                              <div style={{ fontSize: 12, color: '#76777d' }}>{sub.customer_phone}</div>
                            </td>
                            <td>
                              <span className={`lx-badge ${sub.status === 'pending' ? 'lx-badge-pending' : 'lx-badge-success'}`}>
                                {sub.status === 'pending' ? 'Chờ xử lý' : 'Hoàn thành'}
                              </span>
                            </td>
                            <td className="lx-mono">{new Date(sub.created_at).toLocaleDateString('vi-VN')}</td>
                            <td className="right">
                              <button
                                className="lx-btn lx-btn-ghost lx-btn-sm"
                                onClick={() => { setSidebarActiveMenu('submissions'); setActiveView('dashboard'); toggleSubmissionRow(sub.id); }}
                              >
                                <span className="material-symbols-outlined" style={{ fontSize: 16 }}>visibility</span>
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    <div className="lx-pagination">
                      <span>Hiển thị 1–{Math.min(5, submissions.length)} / {submissions.length} hồ sơ</span>
                      <div style={{ display: 'flex', gap: 6 }}>
                        <button className="lx-page-btn" disabled>
                          <span className="material-symbols-outlined">chevron_left</span>
                        </button>
                        <button className="lx-page-btn" disabled={submissions.length <= 5}>
                          <span className="material-symbols-outlined">chevron_right</span>
                        </button>
                      </div>
                    </div>
                  </>
                )}
              </div>
            </div>
          )}

          {/* ── VIEW: TEMPLATES ── */}
          {sidebarActiveMenu === 'templates' && activeView !== 'config' && (
            <div>
              <div className="lx-page-header" style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 16 }}>
                <div>
                  <div className="lx-page-title">Quản lý Template</div>
                  <div className="lx-page-subtitle">
                    Thư viện biểu mẫu — {templates.length} biểu mẫu, {templates.filter(t => t.status === 'active').length} đang hoạt động
                  </div>
                </div>
                <button
                  className="lx-btn lx-btn-primary"
                  onClick={() => setShowUploadModal(true)}
                >
                  <span className="material-symbols-outlined">upload_file</span>
                  Tải lên Template mới
                </button>
              </div>

              {/* Filter toolbar */}
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 16, marginBottom: 20, flexWrap: 'wrap' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                  <div className="lx-tabs">
                    <button
                      className={`lx-tab ${templateFilter === 'all' ? 'active' : ''}`}
                      onClick={() => setTemplateFilter('all')}
                    >
                      Tất cả
                    </button>
                    <button
                      className={`lx-tab ${templateFilter === 'active' ? 'active' : ''}`}
                      onClick={() => setTemplateFilter('active')}
                    >
                      Đang hoạt động
                    </button>
                    <button
                      className={`lx-tab ${templateFilter === 'draft' ? 'active' : ''}`}
                      onClick={() => setTemplateFilter('draft')}
                    >
                      Bản nháp
                    </button>
                  </div>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                  <div style={{ position: 'relative' }}>
                    <span className="material-symbols-outlined" style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', fontSize: 18, color: '#76777d', pointerEvents: 'none' }}>search</span>
                    <input
                      type="text"
                      value={dashboardSearch}
                      onChange={(e) => setDashboardSearch(e.target.value)}
                      placeholder="Tìm theo tên biểu mẫu..."
                      className="lx-input"
                      style={{ paddingLeft: 36, width: 260 }}
                    />
                  </div>
                  <select
                    value={sortBy}
                    onChange={(e) => setSortBy(e.target.value)}
                    className="lx-select"
                    style={{ width: 'auto' }}
                  >
                    <option value="newest">Mới nhất</option>
                    <option value="oldest">Cũ nhất</option>
                    <option value="name">Tên A → Z</option>
                  </select>
                </div>
              </div>

              {isLoading && templates.length === 0 ? (
                <div className="lx-empty"><span className="lx-spinner" /></div>
              ) : templates.length === 0 ? (
                <div className="lx-empty">
                  <span className="material-symbols-outlined" style={{ fontSize: 40, display: 'block', marginBottom: 8 }}>description</span>
                  Chưa có biểu mẫu nào. Nhấn "Tải lên Template mới" để bắt đầu.
                </div>
              ) : (
                <div style={{ display: 'grid', gridTemplateColumns: '220px 1fr', gap: 20, alignItems: 'start' }}>
                  {/* Category sidebar */}
                  <div className="lx-card" style={{ padding: '14px 0' }}>
                    <div style={{ padding: '0 14px 10px', borderBottom: '1px solid #e6e8ea', marginBottom: 6 }}>
                      <span style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.06em', color: '#45464d' }}>
                        Danh mục ({categories.length})
                      </span>
                    </div>
                    <div style={{ padding: '0 4px' }}>
                      <button
                        className={`lx-cat-btn ${selectedCategoryId === 'all' ? 'active' : ''}`}
                        onClick={() => setSelectedCategoryId('all')}
                      >
                        <span>Tất cả danh mục</span>
                        <span className="lx-cat-count">{templates.length}</span>
                      </button>
                      <button
                        className={`lx-cat-btn ${selectedCategoryId === 'uncategorized' ? 'active' : ''}`}
                        onClick={() => setSelectedCategoryId('uncategorized')}
                      >
                        <span>Chưa phân loại</span>
                        <span className="lx-cat-count">{templates.filter(t => !t.category_id).length}</span>
                      </button>
                      <div style={{ height: 1, background: '#e6e8ea', margin: '4px 8px' }} />
                      {renderCategoryTree()}
                    </div>
                  </div>

                  {/* Template grid */}
                  {getFilteredTemplates().length === 0 ? (
                    <div className="lx-empty">
                      <span className="material-symbols-outlined" style={{ fontSize: 36, display: 'block', marginBottom: 8 }}>search</span>
                      Không tìm thấy biểu mẫu nào phù hợp.
                    </div>
                  ) : (
                    <div className="lx-template-grid">
                      {getFilteredTemplates().map((temp) => (
                        <div key={temp.id} className="lx-template-card">
                          <div className="lx-template-card-header">
                            <div className="lx-template-icon">
                              <span className="material-symbols-outlined">folder</span>
                            </div>
                            <span className={`lx-badge ${temp.status === 'active' ? 'lx-badge-published' : 'lx-badge-draft'}`}>
                              {temp.status === 'active' ? 'Hoạt động' : 'Bản nháp'}
                            </span>
                          </div>
                          <div className="lx-template-card-name line-clamp-2">{temp.name}</div>
                          <div className="lx-template-card-meta">
                            <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                              <span className="material-symbols-outlined" style={{ fontSize: 14 }}>code</span>
                              {temp.variables_count} biến
                            </span>
                            {temp.children_count > 0 && (
                              <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                                <span className="material-symbols-outlined" style={{ fontSize: 14 }}>link</span>
                                {temp.children_count} file con
                              </span>
                            )}
                            {temp.parent_template_id && (
                              <span style={{ display: 'flex', alignItems: 'center', gap: 4, color: '#76777d' }}>
                                <span className="material-symbols-outlined" style={{ fontSize: 14 }}>account_tree</span>
                                Con của biểu mẫu khác
                              </span>
                            )}
                          </div>
                          <div className="lx-template-card-actions">
                            <button
                              className="lx-btn lx-btn-primary lx-btn-sm"
                              style={{ flex: 1 }}
                              onClick={() => handleOpenConfig(temp.id)}
                            >
                              <span className="material-symbols-outlined" style={{ fontSize: 16 }}>tune</span>
                              Sửa Mapping
                            </button>
                            <button
                              className="lx-icon-btn-sm"
                              title="Bật/Tắt trạng thái"
                              onClick={() => handleToggleTemplateStatus(temp)}
                            >
                              <span className="material-symbols-outlined">
                                {temp.status === 'active' ? 'visibility' : 'visibility_off'}
                              </span>
                            </button>
                            <button
                              className="lx-icon-btn-sm"
                              title="Xóa biểu mẫu"
                              style={{ color: '#ba1a1a' }}
                              onClick={() => handleDeleteTemplate(temp.id)}
                            >
                              <span className="material-symbols-outlined">delete</span>
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>
          )}

          {/* ── VIEW: SUBMISSIONS ── */}
          {sidebarActiveMenu === 'submissions' && activeView !== 'config' && (
            <div>
              <div className="lx-page-header" style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
                <div>
                  <div className="lx-page-title">Danh sách Hồ sơ</div>
                  <div className="lx-page-subtitle">
                    Lịch sử khách hàng gửi dữ liệu trực tuyến
                  </div>
                </div>
                <span className="lx-badge lx-badge-pending" style={{ marginLeft: 8 }}>
                  {submissions.length} bản ghi
                </span>
              </div>

              <div className="lx-card">
                {submissions.length === 0 ? (
                  <div className="lx-empty">
                    {isLoading ? <span className="lx-spinner" /> : 'Chưa có khách hàng nào nộp hồ sơ trực tuyến.'}
                  </div>
                ) : (
                  <table className="lx-table">
                    <thead>
                      <tr>
                        <th style={{ width: 48, textAlign: 'center' }}>#</th>
                        <th>Khách hàng</th>
                        <th>Biểu mẫu gốc</th>
                        <th>Ngày gửi</th>
                        <th>Số lượng tệp</th>
                        <th className="right">Thao tác</th>
                      </tr>
                    </thead>
                    <tbody>
                      {submissions.map((sub, idx) => {
                        const isExpanded = expandedSubmissionIds.has(sub.id);
                        let parsedFiles = [];
                        try {
                          parsedFiles = typeof sub.files === 'string' ? JSON.parse(sub.files) : (sub.files || []);
                        } catch (e) {
                          console.warn('Lỗi parse files:', e);
                        }

                        return (
                          <Fragment key={sub.id}>
                            <tr>
                              <td style={{ textAlign: 'center' }} className="lx-mono">{idx + 1}</td>
                              <td>
                                <div style={{ fontWeight: 600 }}>{sub.customer_name}</div>
                                <div style={{ fontSize: 12, color: '#76777d' }}>{sub.customer_phone}</div>
                              </td>
                              <td style={{ fontWeight: 500 }}>{sub.template_name}</td>
                              <td className="lx-mono">{new Date(sub.created_at).toLocaleString('vi-VN')}</td>
                              <td>
                                <span className="lx-badge lx-badge-success">{parsedFiles.length} tài liệu</span>
                              </td>
                              <td className="right">
                                <button
                                  className="lx-btn lx-btn-secondary lx-btn-sm"
                                  onClick={() => {
                                    setViewingSubmission(sub);
                                    setActiveSubmissionPreview({ ...activeSubmissionPreview, [sub.id]: parsedFiles[0] });
                                  }}
                                >
                                  <span className="material-symbols-outlined" style={{ fontSize: 15 }}>open_in_full</span>
                                  Xem chi tiết
                                </button>
                              </td>
                            </tr>

                            {isExpanded && (
                              <tr>
                                <td colSpan={6} style={{ padding: 0, background: '#f2f4f6', borderBottom: '2px solid #c6c6cd' }}>
                                  {/* Two-column layout: 65% doc preview | 35% sidebar */}
                                  <div style={{ display: 'grid', gridTemplateColumns: '65fr 35fr', height: 720 }}>

                                    {/* Left: document preview — full height */}
                                    <div style={{ display: 'flex', flexDirection: 'column', borderRight: '1px solid #c6c6cd', overflow: 'hidden' }}>
                                      {/* Toolbar */}
                                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 16px', background: '#ffffff', borderBottom: '1px solid #e6e8ea', flexShrink: 0 }}>
                                        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                          <span className="material-symbols-outlined" style={{ fontSize: 17, color: '#45464d' }}>description</span>
                                          <span style={{ fontSize: 12, fontWeight: 600, color: '#191c1e' }}>Xem trước tài liệu</span>
                                        </div>
                                        <span className="lx-badge lx-badge-pending" style={{ maxWidth: 260, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                          {activeSubmissionPreview[sub.id] || parsedFiles[0]}
                                        </span>
                                      </div>
                                      {/* DocxPreview takes remaining height */}
                                      <div style={{ flex: 1, overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
                                        {parsedFiles.length > 0 ? (
                                          <DocxPreview
                                            fileUrl={`http://localhost:5000/api/submissions/${sub.id}/download-file?filename=${encodeURIComponent(activeSubmissionPreview[sub.id] || parsedFiles[0])}`}
                                            title={activeSubmissionPreview[sub.id] || parsedFiles[0]}
                                            fields={[]}
                                            liveData={{}}
                                          />
                                        ) : (
                                          <div className="lx-empty">Không có tài liệu nào để xem trước.</div>
                                        )}
                                      </div>
                                    </div>

                                    {/* Right sidebar: file list + JSON — scrollable */}
                                    <div style={{ display: 'flex', flexDirection: 'column', overflowY: 'auto', background: '#ffffff' }}>
                                      {/* File list */}
                                      <div style={{ padding: '16px 16px 12px' }}>
                                        <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.06em', color: '#45464d', borderBottom: '1px solid #e6e8ea', paddingBottom: 8, marginBottom: 10 }}>
                                          Tài liệu trong hồ sơ ({parsedFiles.length})
                                        </div>
                                        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                                          {parsedFiles.map(file => {
                                            const isActive = (activeSubmissionPreview[sub.id] || parsedFiles[0]) === file;
                                            return (
                                              <div key={file} style={{
                                                padding: '9px 12px', borderRadius: 4,
                                                border: `1px solid ${isActive ? '#4edea3' : '#e6e8ea'}`,
                                                background: isActive ? '#f0fdf4' : '#f7f9fb',
                                                display: 'flex', flexDirection: 'column', gap: 8
                                              }}>
                                                <span style={{ fontSize: 12, fontWeight: isActive ? 600 : 400, color: isActive ? '#005236' : '#191c1e', wordBreak: 'break-all', lineHeight: 1.4 }}>
                                                  {file}
                                                </span>
                                                <div style={{ display: 'flex', gap: 6 }}>
                                                  <button
                                                    className={`lx-btn lx-btn-sm ${isActive ? 'lx-btn-primary' : 'lx-btn-secondary'}`}
                                                    style={{ flex: 1 }}
                                                    onClick={() => setActiveSubmissionPreview({ ...activeSubmissionPreview, [sub.id]: file })}
                                                  >
                                                    <span className="material-symbols-outlined" style={{ fontSize: 13 }}>{isActive ? 'visibility' : 'preview'}</span>
                                                    {isActive ? 'Đang xem' : 'Xem trước'}
                                                  </button>
                                                  <a
                                                    href={`http://localhost:5000/api/submissions/${sub.id}/download-file?filename=${encodeURIComponent(file)}`}
                                                    download
                                                    className="lx-btn lx-btn-secondary lx-btn-sm"
                                                  >
                                                    <span className="material-symbols-outlined" style={{ fontSize: 13 }}>download</span>
                                                  </a>
                                                </div>
                                              </div>
                                            );
                                          })}
                                        </div>
                                      </div>

                                      {/* JSON data */}
                                      <div style={{ padding: '0 16px 16px', flex: 1 }}>
                                        <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.06em', color: '#45464d', borderBottom: '1px solid #e6e8ea', paddingBottom: 8, marginBottom: 10 }}>
                                          Dữ liệu đã nộp (JSON)
                                        </div>
                                        <pre style={{ background: '#191c1e', color: '#4edea3', fontSize: 10, padding: '12px', borderRadius: 4, overflowX: 'auto', lineHeight: 1.6, fontFamily: 'ui-monospace, monospace', whiteSpace: 'pre-wrap', wordBreak: 'break-all' }}>
                                          {JSON.stringify(typeof sub.values_json === 'string' ? JSON.parse(sub.values_json) : sub.values_json, null, 2)}
                                        </pre>
                                      </div>
                                    </div>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </Fragment>
                        );
                      })}
                    </tbody>
                  </table>
                )}
              </div>
            </div>
          )}

          {/* ── VIEW: CONFIG ── */}
          {activeView === 'config' && selectedTemplate && (
            <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
              {/* Config header */}
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16, flexWrap: 'wrap' }}>
                <button
                  className="lx-btn lx-btn-secondary lx-btn-sm"
                  onClick={() => {
                    setSidebarActiveMenu('templates');
                    setActiveView('dashboard');
                    setActiveConfigTab(null);
                    setParentTabFields([]);
                    setTabFieldsCache({});
                  }}
                >
                  <span className="material-symbols-outlined" style={{ fontSize: 16 }}>arrow_back</span>
                  Templates
                </button>
                <span style={{ color: '#76777d', fontSize: 13 }}>/</span>
                <span style={{ fontWeight: 600, fontSize: 14, color: '#191c1e' }}>{selectedTemplate.name}</span>
                {hasUnsavedManual && (
                  <span className="lx-unsaved">
                    <span className="material-symbols-outlined" style={{ fontSize: 14 }}>warning</span>
                    Có thay đổi chưa lưu
                  </span>
                )}
                <div style={{ marginLeft: 'auto', display: 'flex', gap: 8, alignItems: 'center' }}>
                  <button
                    className="lx-btn lx-btn-primary lx-btn-sm"
                    onClick={handleSaveConfig}
                  >
                    <span className="material-symbols-outlined" style={{ fontSize: 16 }}>save</span>
                    Lưu & Kích hoạt
                  </button>
                  <button
                    className="lx-btn lx-btn-secondary lx-btn-sm"
                    onClick={handleUndo}
                    disabled={historyIndex <= 0}
                  >
                    <span className="material-symbols-outlined" style={{ fontSize: 16 }}>undo</span>
                    Hoàn tác
                  </button>
                  <button
                    className="lx-btn lx-btn-secondary lx-btn-sm"
                    onClick={handleExportConfig}
                    title="Xuất cấu hình JSON"
                  >
                    <span className="material-symbols-outlined" style={{ fontSize: 16 }}>upload</span>
                    Xuất
                  </button>
                  <label className="lx-btn lx-btn-secondary lx-btn-sm" title="Nhập cấu hình JSON" style={{ cursor: 'pointer' }}>
                    <span className="material-symbols-outlined" style={{ fontSize: 16 }}>download</span>
                    Nhập
                    <input type="file" accept=".json" onChange={handleImportConfig} style={{ display: 'none' }} />
                  </label>
                </div>
              </div>

              {/* Template tab switcher — only shown when there are linked children */}
              {linkedChildren.length > 0 && (
                <div style={{ display: 'flex', gap: 0, marginBottom: 0, borderBottom: '2px solid #e6e8ea', overflowX: 'auto' }}>
                  {/* Parent tab */}
                  <button
                    type="button"
                    onClick={() => handleSwitchConfigTab(null)}
                    style={{
                      padding: '8px 16px',
                      fontSize: 13,
                      fontWeight: activeConfigTab === null ? 700 : 500,
                      color: activeConfigTab === null ? '#009668' : '#45464d',
                      background: 'transparent',
                      border: 'none',
                      borderBottom: activeConfigTab === null ? '2px solid #009668' : '2px solid transparent',
                      cursor: 'pointer',
                      whiteSpace: 'nowrap',
                      marginBottom: -2,
                      display: 'flex',
                      alignItems: 'center',
                      gap: 6
                    }}
                  >
                    <span className="material-symbols-outlined" style={{ fontSize: 14 }}>description</span>
                    {selectedTemplate.name}
                  </button>
                  {/* Child tabs */}
                  {linkedChildren.map(child => (
                    <button
                      key={child.id}
                      type="button"
                      onClick={() => handleSwitchConfigTab(child)}
                      style={{
                        padding: '8px 16px',
                        fontSize: 13,
                        fontWeight: activeConfigTab?.id === child.id ? 700 : 500,
                        color: activeConfigTab?.id === child.id ? '#3b5bdb' : '#45464d',
                        background: 'transparent',
                        border: 'none',
                        borderBottom: activeConfigTab?.id === child.id ? '2px solid #3b5bdb' : '2px solid transparent',
                        cursor: 'pointer',
                        whiteSpace: 'nowrap',
                        marginBottom: -2,
                        display: 'flex',
                        alignItems: 'center',
                        gap: 6
                      }}
                    >
                      <span className="material-symbols-outlined" style={{ fontSize: 14 }}>attach_file</span>
                      {child.name}
                    </button>
                  ))}
                </div>
              )}

              {/* Split panel */}
              <div
                ref={splitContainerRef}
                className="lx-split"
                style={{ flex: 1, minHeight: 600 }}
              >
                {/* Left panel */}
                <div
                  style={{ width: isMobile ? '100%' : `${leftWidth}%`, display: 'flex', flexDirection: 'column', overflowY: 'auto', gap: 16, padding: '4px 0' }}
                >
                  {/* Template info + linked children combined */}
                  <div style={{ background: '#ffffff', border: '1px solid #c6c6cd', borderRadius: 8, padding: 14 }}>
                    {/* Header row */}
                    <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 8, marginBottom: 10 }}>
                      <div style={{ minWidth: 0 }}>
                        <span style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.06em', color: '#009668' }}>Chi tiết biểu mẫu</span>
                        <div style={{ fontWeight: 700, fontSize: 14, color: '#191c1e', marginTop: 3, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{selectedTemplate.name}</div>
                      </div>
                      {selectedTemplate.parent_template_id ? (
                        <span className="lx-badge" style={{ flexShrink: 0, marginTop: 2 }}>Phụ lục</span>
                      ) : null}
                    </div>

                    {/* Category */}
                    <div className="lx-form-group" style={{ marginBottom: 10 }}>
                      <label className="lx-label">Danh mục thư viện</label>
                      <select
                        className="lx-select"
                        value={selectedTemplate.category_id || ''}
                        onChange={(e) => {
                          handleUpdateTemplateCategory(selectedTemplate.id, e.target.value);
                          setSelectedTemplate(prev => prev ? { ...prev, category_id: e.target.value || null } : prev);
                        }}
                      >
                        <option value="">Chưa phân loại</option>
                        {getFlattenedCategoryOptions().map(category => (
                          <option key={category.id} value={category.id}>{category.label}</option>
                        ))}
                      </select>
                    </div>

                    {/* Linked children (only for parent templates) */}
                    {!selectedTemplate.parent_template_id && (
                      <div style={{ borderTop: '1px solid #e6e8ea', paddingTop: 10 }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: linkedChildren.length > 0 ? 8 : 6 }}>
                          <span className="material-symbols-outlined" style={{ fontSize: 14, color: '#45464d' }}>account_tree</span>
                          <span style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.06em', color: '#45464d', flex: 1 }}>
                            Phụ lục ({linkedChildren.length})
                          </span>
                          <button
                            type="button"
                            className="lx-btn lx-btn-secondary lx-btn-sm"
                            onClick={() => handleOpenLinkModal(selectedTemplate)}
                          >
                            <span className="material-symbols-outlined" style={{ fontSize: 13 }}>add_link</span>
                            Quản lý
                          </button>
                        </div>
                        {linkedChildren.length === 0 ? (
                          <div style={{ fontSize: 12, color: '#76777d', fontStyle: 'italic' }}>
                            Chưa có phụ lục. Nhấn "Quản lý" để liên kết.
                          </div>
                        ) : (
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
                            {linkedChildren.map(child => (
                              <div key={child.id} style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '5px 8px', background: '#f7f9fb', borderRadius: 4, border: '1px solid #e6e8ea', minWidth: 0 }}>
                                <span className="material-symbols-outlined" style={{ fontSize: 14, color: '#009668', flexShrink: 0 }}>description</span>
                                <span style={{ fontSize: 12, fontWeight: 600, flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{child.name}</span>
                                {child.is_repeated === 1 && (
                                  <span style={{ fontSize: 9, fontWeight: 700, background: '#e0faf0', color: '#009668', borderRadius: 3, padding: '2px 4px', flexShrink: 0 }}>LẶP</span>
                                )}
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    )}
                  </div>

                  {/* Quick-tag box when text is selected */}
                  {selectedText && (
                    <div className="lx-quick-tag-box">
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
                        <span style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.06em', color: '#191c1e' }}>
                          Tạo biến nhanh từ chữ chọn
                        </span>
                        <button
                          className="lx-btn lx-btn-ghost lx-btn-sm"
                          onClick={() => { setSelectedText(''); window.getSelection()?.removeAllRanges(); }}
                        >
                          Đóng
                        </button>
                      </div>
                      <div style={{ background: '#ffffff', border: '1px solid #e6e8ea', borderRadius: 4, padding: '8px 12px', fontSize: 12, marginBottom: 12, color: '#45464d' }}>
                        Đang chọn: <strong style={{ color: '#191c1e' }}>"{selectedText}"</strong>
                      </div>
                      <form onSubmit={handleAddQuickField} style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                        {/* Inherit toggle — only shown when editing a child tab with parent fields available */}
                        {activeConfigTab && parentTabFields.filter(f => f.field_type !== 'label').length > 0 && (
                          <div style={{ display: 'flex', gap: 4, background: '#f0f4ff', borderRadius: 6, padding: 3 }}>
                            <button
                              type="button"
                              onClick={() => setQuickInheritMode(false)}
                              style={{ flex: 1, padding: '5px 8px', fontSize: 12, fontWeight: quickInheritMode ? 400 : 700, background: quickInheritMode ? 'transparent' : '#ffffff', border: 'none', borderRadius: 4, cursor: 'pointer', color: quickInheritMode ? '#45464d' : '#3b5bdb', boxShadow: quickInheritMode ? 'none' : '0 1px 3px rgba(0,0,0,0.12)' }}
                            >
                              Tạo biến mới
                            </button>
                            <button
                              type="button"
                              onClick={() => setQuickInheritMode(true)}
                              style={{ flex: 1, padding: '5px 8px', fontSize: 12, fontWeight: quickInheritMode ? 700 : 400, background: quickInheritMode ? '#ffffff' : 'transparent', border: 'none', borderRadius: 4, cursor: 'pointer', color: quickInheritMode ? '#009668' : '#45464d', boxShadow: quickInheritMode ? '0 1px 3px rgba(0,0,0,0.12)' : 'none', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 4 }}
                            >
                              <span className="material-symbols-outlined" style={{ fontSize: 13 }}>link</span>
                              Kế thừa từ Template cha
                            </button>
                          </div>
                        )}

                        {quickInheritMode ? (
                          <div className="lx-form-group">
                            <label className="lx-label">Chọn trường từ template cha</label>
                            <select
                              className="lx-select"
                              value={quickParentFieldKey}
                              onChange={(e) => setQuickParentFieldKey(e.target.value)}
                            >
                              <option value="">-- Chọn trường --</option>
                              {parentTabFields.filter(f => f.field_type !== 'label').map(f => {
                                const inheritCount = fields.filter(cf => cf.parent_field_key === f.key_name).length;
                                return (
                                  <option key={f.key_name} value={f.key_name}>
                                    {f.label} ({f.key_name}){inheritCount > 0 ? ` — đã kế thừa ${inheritCount} lần` : ''}
                                  </option>
                                );
                              })}
                            </select>
                            <div style={{ fontSize: 11, color: '#76777d', marginTop: 4 }}>
                              Trường kế thừa sẽ tự động điền giá trị từ template cha khi người dùng nhập.
                            </div>
                          </div>
                        ) : (
                          <>
                            <div className="lx-form-group">
                              <label className="lx-label">Mã biến (không dấu/khoảng trắng)</label>
                              <input
                                type="text"
                                required={!quickInheritMode}
                                value={quickKey}
                                onChange={(e) => setQuickKey(e.target.value)}
                                placeholder="vd: ho_ten_chu_dat"
                                className="lx-input"
                              />
                            </div>
                            <div className="lx-form-row">
                              <div className="lx-form-group">
                                <label className="lx-label">Nhãn hiển thị</label>
                                <input
                                  type="text"
                                  required={!quickInheritMode}
                                  value={quickLabel}
                                  onChange={(e) => setQuickLabel(e.target.value)}
                                  placeholder="vd: Họ tên chủ đất"
                                  className="lx-input"
                                />
                              </div>
                              <div className="lx-form-group">
                                <label className="lx-label">Kiểu dữ liệu</label>
                                <select className="lx-select" value={quickType} onChange={(e) => setQuickType(e.target.value)}>
                                  <option value="text">Chữ (Text)</option>
                                  <option value="date">Ngày (Date)</option>
                                  <option value="number">Số (Number)</option>
                                  <option value="boolean">Đúng/Sai</option>
                                </select>
                              </div>
                            </div>
                          </>
                        )}
                        <button type="submit" className="lx-btn lx-btn-primary lx-btn-sm" style={{ width: '100%' }}>
                          <span className="material-symbols-outlined" style={{ fontSize: 16 }}>{quickInheritMode ? 'link' : 'add'}</span>
                          {quickInheritMode ? 'Kế thừa trường này' : 'Tạo biến động'}
                        </button>
                      </form>
                    </div>
                  )}

                  {/* Manual add */}
                  <div>
                    {!showManualAdd ? (
                      <button
                        type="button"
                        className="lx-btn lx-btn-secondary"
                        style={{ width: '100%', borderStyle: 'dashed' }}
                        onClick={handleOpenManualAdd}
                      >
                        <span className="material-symbols-outlined" style={{ fontSize: 18 }}>add</span>
                        Thêm biến thủ công
                      </button>
                    ) : (
                      <div className="lx-quick-tag-box">
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
                          <span style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.06em', color: '#191c1e' }}>Thêm biến thủ công</span>
                          <button className="lx-btn lx-btn-ghost lx-btn-sm" type="button" onClick={() => setShowManualAdd(false)}>Đóng</button>
                        </div>
                        <p style={{ fontSize: 12, color: '#45464d', marginBottom: 10, lineHeight: 1.6 }}>
                          Dùng khi DOCX có dòng như <em>"Nơi cấp:"</em> nhưng chưa có <code style={{ background: '#e0faf0', color: '#009668', padding: '1px 6px', borderRadius: 3, fontFamily: 'monospace', fontSize: 11 }}>{'{{bien}}'}</code>. Sau khi thêm, hãy gõ <code style={{ background: '#e0faf0', color: '#009668', padding: '1px 6px', borderRadius: 3, fontFamily: 'monospace', fontSize: 11 }}>{`{{ten_bien}}`}</code> vào file Word.
                        </p>
                        <form onSubmit={handleAddManualField} style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                          <div className="lx-form-row">
                            <div className="lx-form-group">
                              <label className="lx-label">Mã biến <span style={{ color: '#ba1a1a' }}>*</span></label>
                              <input
                                type="text"
                                required
                                value={quickKey}
                                onChange={(e) => setQuickKey(e.target.value)}
                                placeholder="vd: noi_cap"
                                className="lx-input"
                              />
                            </div>
                            <div className="lx-form-group">
                              <label className="lx-label">Nhãn hiển thị <span style={{ color: '#ba1a1a' }}>*</span></label>
                              <input
                                type="text"
                                required
                                value={quickLabel}
                                onChange={(e) => setQuickLabel(e.target.value)}
                                placeholder="vd: Nơi cấp"
                                className="lx-input"
                              />
                            </div>
                          </div>
                          <div className="lx-form-group">
                            <label className="lx-label">Kiểu dữ liệu</label>
                            <select className="lx-select" value={quickType} onChange={(e) => setQuickType(e.target.value)}>
                              <option value="text">Chữ (Text)</option>
                              <option value="date">Ngày (Date)</option>
                              <option value="number">Số (Number)</option>
                              <option value="boolean">Đúng/Sai</option>
                            </select>
                          </div>
                          <button type="submit" className="lx-btn lx-btn-primary lx-btn-sm" style={{ width: '100%' }}>
                            <span className="material-symbols-outlined" style={{ fontSize: 16 }}>add</span>
                            Tạo biến thủ công
                          </button>
                        </form>
                      </div>
                    )}
                  </div>

                  {/* Section label add */}
                  <div>
                    {!showLabelAdd ? (
                      <button
                        type="button"
                        className="lx-btn lx-btn-secondary"
                        style={{ width: '100%', borderStyle: 'dashed', borderColor: '#3b5bdb', color: '#3b5bdb' }}
                        onClick={() => { setShowLabelAdd(true); setShowManualAdd(false); setQuickKey(''); setQuickLabel(''); setLabelTitle(''); }}
                      >
                        <span className="material-symbols-outlined" style={{ fontSize: 18 }}>title</span>
                        Thêm tiêu đề nhóm
                      </button>
                    ) : (
                      <div className="lx-quick-tag-box" style={{ borderColor: '#3b5bdb' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
                          <span style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.06em', color: '#3b5bdb' }}>Thêm tiêu đề nhóm</span>
                          <button className="lx-btn lx-btn-ghost lx-btn-sm" type="button" onClick={() => setShowLabelAdd(false)}>Đóng</button>
                        </div>
                        <p style={{ fontSize: 12, color: '#45464d', marginBottom: 10, lineHeight: 1.6 }}>
                          Tạo tiêu đề phân nhóm hiển thị trong form điền — không phải biến dữ liệu, không xuất hiện trong file Word.
                        </p>
                        <form onSubmit={handleAddLabelField} style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                          <div className="lx-form-group">
                            <label className="lx-label">Nội dung tiêu đề <span style={{ color: '#ba1a1a' }}>*</span></label>
                            <input
                              type="text"
                              required
                              value={labelTitle}
                              onChange={(e) => setLabelTitle(e.target.value)}
                              placeholder="vd: BÊN A — Bên Chuyển Nhượng (Bán)"
                              className="lx-input"
                              autoFocus
                            />
                          </div>
                          <button type="submit" className="lx-btn lx-btn-primary lx-btn-sm" style={{ width: '100%' }}>
                            <span className="material-symbols-outlined" style={{ fontSize: 16 }}>add</span>
                            Tạo tiêu đề nhóm
                          </button>
                        </form>
                      </div>
                    )}
                  </div>

                  {/* Field list */}
                  {fields.length === 0 ? (
                    <div className="lx-empty">
                      Không quét được biến động nào dạng {`{{ten_bien}}`} từ file. Vui lòng kiểm tra lại file Word.
                    </div>
                  ) : (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                      {fields.map((field, idx) => {
                        const isDragging = draggedIdx === idx;
                        const isDropTarget = dragOverIdx === idx && draggedIdx !== idx;
                        const commonDragProps = {
                          draggable: true,
                          onDragStart: () => setDraggedIdx(idx),
                          onDragOver: (e) => { e.preventDefault(); setDragOverIdx(idx); },
                          onDrop: () => {
                            if (draggedIdx === null || draggedIdx === idx) return;
                            const arr = [...fields];
                            const [moved] = arr.splice(draggedIdx, 1);
                            arr.splice(idx, 0, moved);
                            updateFieldsAndHistory(arr.map((f, i) => ({ ...f, order_index: i })));
                            setDraggedIdx(null); setDragOverIdx(null);
                          },
                          onDragEnd: () => { setDraggedIdx(null); setDragOverIdx(null); }
                        };

                        if (field.field_type === 'label') {
                          return (
                            <div key={field.id} {...commonDragProps} style={{ background: '#1e3a5f', borderRadius: 6, padding: '8px 12px', display: 'flex', alignItems: 'center', gap: 10, opacity: isDragging ? 0.4 : 1, borderTop: isDropTarget ? '2px solid #fbbf24' : undefined, cursor: 'grab' }}>
                              <span className="material-symbols-outlined" style={{ color: '#4a6a90', fontSize: 16, flexShrink: 0, cursor: 'grab' }}>drag_indicator</span>
                              <span className="material-symbols-outlined" style={{ color: '#fbbf24', fontSize: 18, flexShrink: 0 }}>title</span>
                              <div style={{ flex: 1 }}>
                                <div style={{ color: '#94a3b8', fontSize: 10, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 3 }}>Tiêu đề nhóm</div>
                                <input
                                  type="text"
                                  value={field.label}
                                  onChange={(e) => handleFieldChange(idx, 'label', e.target.value)}
                                  onDragStart={(e) => e.stopPropagation()}
                                  style={{ background: 'transparent', border: '1px solid #2d4a6e', borderRadius: 4, padding: '4px 8px', color: '#ffffff', fontSize: 13, fontWeight: 600, width: '100%', outline: 'none' }}
                                  onFocus={e => { e.target.style.borderColor = '#fbbf24'; }}
                                  onBlur={e => { e.target.style.borderColor = '#2d4a6e'; }}
                                />
                              </div>
                              <button
                                type="button"
                                title="Xóa tiêu đề"
                                onClick={() => updateFieldsAndHistory(fields.filter((_, i) => i !== idx))}
                                style={{ background: 'transparent', border: 'none', color: '#94a3b8', cursor: 'pointer', padding: 4, display: 'flex', alignItems: 'center' }}
                              >
                                <span className="material-symbols-outlined" style={{ fontSize: 18 }}>delete</span>
                              </button>
                            </div>
                          );
                        }
                        return (
                        <div id={`field-config-${field.key_name}`} key={field.id} className="lx-field-card" {...commonDragProps}
                          style={{ opacity: isDragging ? 0.4 : 1, borderTop: isDropTarget ? '2px solid #3b5bdb' : undefined, cursor: 'default' }}
                        >
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: 6, minWidth: 0 }}>
                              <span className="material-symbols-outlined" style={{ fontSize: 16, color: '#c6c6cd', cursor: 'grab', flexShrink: 0 }}>drag_indicator</span>
                              <span className="lx-field-key" style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{`{{${field.key_name}}}`}</span>
                            </div>
                            <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexShrink: 0 }}>
                              <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: '#45464d', cursor: 'pointer' }}>
                                <input
                                  type="checkbox"
                                  id={`req-${field.id}`}
                                  checked={!!field.is_required}
                                  onChange={(e) => handleFieldChange(idx, 'is_required', e.target.checked)}
                                  style={{ accentColor: '#000000', cursor: 'pointer' }}
                                />
                                Bắt buộc
                              </label>
                              <button
                                type="button"
                                title="Xóa trường"
                                onClick={() => updateFieldsAndHistory(fields.filter((_, i) => i !== idx))}
                                style={{ background: 'transparent', border: 'none', color: '#c6c6cd', cursor: 'pointer', padding: 2, display: 'flex', alignItems: 'center' }}
                                onMouseEnter={e => e.currentTarget.style.color = '#ba1a1a'}
                                onMouseLeave={e => e.currentTarget.style.color = '#c6c6cd'}
                              >
                                <span className="material-symbols-outlined" style={{ fontSize: 16 }}>delete</span>
                              </button>
                            </div>
                          </div>
                          <div className="lx-form-row">
                            <div className="lx-form-group">
                              <label className="lx-label">Nhãn hiển thị</label>
                              <input
                                type="text"
                                value={field.label}
                                onChange={(e) => handleFieldChange(idx, 'label', e.target.value)}
                                className="lx-input"
                              />
                            </div>
                            <div className="lx-form-group">
                              <label className="lx-label">Kiểu dữ liệu</label>
                              <select
                                value={field.field_type}
                                onChange={(e) => handleFieldChange(idx, 'field_type', e.target.value)}
                                className="lx-select"
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
                              <div style={{ marginTop: 10, paddingTop: 10, borderTop: `1px dashed ${isMissingParent ? '#ba1a1a' : isConflict ? '#f59e0b' : '#e6e8ea'}` }}>
                                <label className="lx-label" style={{ color: '#009668', display: 'flex', alignItems: 'center', gap: 6, marginBottom: 6 }}>
                                  <span className="material-symbols-outlined" style={{ fontSize: 14 }}>link</span>
                                  Đồng bộ từ trường File Gốc
                                  {sameNameInParent && !field.parent_field_key && (
                                    <span className="lx-badge lx-badge-success" style={{ marginLeft: 'auto', fontSize: 10 }}>tự đồng bộ theo tên</span>
                                  )}
                                </label>
                                <select
                                  value={field.parent_field_key || ''}
                                  onChange={(e) => handleFieldChange(idx, 'parent_field_key', e.target.value || null)}
                                  className="lx-select"
                                  style={{
                                    borderColor: isMissingParent ? '#ba1a1a' : isConflict ? '#f59e0b' : undefined
                                  }}
                                >
                                  <option value="">-- Tự điền độc lập (Không đồng bộ) --</option>
                                  {parentFields.map(pf => (
                                    <option key={pf.id} value={pf.key_name}>
                                      {pf.label} ({`{{${pf.key_name}}}`})
                                    </option>
                                  ))}
                                  {isMissingParent && (
                                    <option value={field.parent_field_key} disabled>
                                      [Lỗi] Thiếu trường cha: {`{{${field.parent_field_key}}}`}
                                    </option>
                                  )}
                                </select>
                              </div>
                            );
                          })()}
                        </div>
                        );
                      })}
                    </div>
                  )}
                </div>

                {/* Resizer */}
                {!isMobile && (
                  <div
                    className="lx-split-divider"
                    onMouseDown={startResizing}
                    title="Kéo chuột để thay đổi kích thước"
                  />
                )}

                {/* Right panel: docx preview */}
                <div style={{ width: isMobile ? '100%' : `${100 - leftWidth}%`, display: 'flex', flexDirection: 'column', flex: 1 }}>
                  <DocxPreview
                    fileUrl={`${API_BASE}/templates/${activeConfigTab ? activeConfigTab.id : selectedTemplate.id}/download-original`}
                    title={`${activeConfigTab ? activeConfigTab.name : selectedTemplate.name} (Bản xem trước để chọn text)`}
                    liveData={{}}
                    fields={fields}
                    highlightField={null}
                    isEditable={true}
                  />
                </div>
              </div>
            </div>
          )}

        </div>
      </div>

      {/* ── UPLOAD MODAL ── */}
      {showUploadModal && (
        <div className="lx-modal-overlay">
          <div className="lx-modal">
            <div className="lx-modal-header">
              <span className="lx-modal-title">Tải lên biểu mẫu Word mới</span>
              <button
                type="button"
                className="lx-icon-btn-sm"
                onClick={() => setShowUploadModal(false)}
              >
                <span className="material-symbols-outlined">close</span>
              </button>
            </div>
            <form onSubmit={handleUploadTemplate}>
              <div className="lx-modal-body" style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
                <div className="lx-form-group">
                  <label className="lx-label">Tên biểu mẫu <span style={{ color: '#ba1a1a' }}>*</span></label>
                  <input
                    type="text"
                    required
                    value={uploadName}
                    onChange={(e) => setUploadName(e.target.value)}
                    placeholder="Ví dụ: Hợp đồng mua bán chuyển nhượng quyền sử dụng đất"
                    className="lx-input"
                  />
                </div>
                <div className="lx-form-group">
                  <label className="lx-label">Mô tả tóm tắt</label>
                  <textarea
                    value={uploadDesc}
                    onChange={(e) => setUploadDesc(e.target.value)}
                    placeholder="Ví dụ: Áp dụng cho các giao dịch chuyển nhượng bất động sản..."
                    rows={3}
                    className="lx-input"
                  />
                </div>
                <div className="lx-form-group">
                  <label className="lx-label">Danh mục thư viện</label>
                  <select
                    value={uploadCategoryId}
                    onChange={(e) => setUploadCategoryId(e.target.value)}
                    className="lx-select"
                  >
                    <option value="">Chưa phân loại</option>
                    {getFlattenedCategoryOptions().map(category => (
                      <option key={category.id} value={category.id}>{category.label}</option>
                    ))}
                  </select>
                </div>
                <div className="lx-form-group">
                  <label className="lx-label">Chọn file (.docx) <span style={{ color: '#ba1a1a' }}>*</span></label>
                  <label style={{
                    display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center',
                    width: '100%', height: 110, border: '2px dashed #c6c6cd', borderRadius: 4,
                    cursor: 'pointer', background: '#f7f9fb', transition: 'border-color 0.15s'
                  }}>
                    <span className="material-symbols-outlined" style={{ fontSize: 32, color: '#76777d', marginBottom: 6 }}>upload_file</span>
                    <span style={{ fontSize: 13, fontWeight: 600, color: '#45464d' }}>
                      {uploadFile ? uploadFile.name : 'Nhấn để chọn file Word'}
                    </span>
                    <span style={{ fontSize: 11, color: '#76777d', marginTop: 2 }}>Định dạng: .doc, .docx</span>
                    <input
                      type="file"
                      accept=".doc,.docx"
                      required
                      style={{ display: 'none' }}
                      onChange={(e) => setUploadFile(e.target.files[0])}
                    />
                  </label>
                </div>
              </div>
              <div className="lx-modal-footer">
                <button
                  type="button"
                  className="lx-btn lx-btn-secondary"
                  onClick={() => setShowUploadModal(false)}
                >
                  Hủy bỏ
                </button>
                <button
                  type="submit"
                  disabled={isUploading}
                  className="lx-btn lx-btn-primary"
                  style={{ opacity: isUploading ? 0.6 : 1 }}
                >
                  <span className="material-symbols-outlined" style={{ fontSize: 16 }}>upload_file</span>
                  {isUploading ? 'Đang phân tích...' : 'Tải lên & Quét biến'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ── LINK MODAL ── */}
      {showLinkModal && linkingTemplate && (
        <div className="lx-modal-overlay">
          <div className="lx-modal lx-modal-lg">
            <div className="lx-modal-header">
              <div>
                <div className="lx-modal-title">Thiết lập quan hệ & liên kết Phụ lục</div>
                <div style={{ fontSize: 12, color: '#76777d', marginTop: 2 }}>
                  Liên kết phụ lục con đi kèm biểu mẫu chính <strong>{linkingTemplate.name}</strong>
                </div>
              </div>
              <button
                type="button"
                className="lx-icon-btn-sm"
                onClick={() => { setShowLinkModal(false); setLinkingTemplate(null); }}
              >
                <span className="material-symbols-outlined">close</span>
              </button>
            </div>
            <div className="lx-modal-body" style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
              {/* Linked children */}
              <div>
                <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.06em', color: '#45464d', borderBottom: '1px solid #e6e8ea', paddingBottom: 8, marginBottom: 12 }}>
                  Các biểu mẫu phụ con đang liên kết
                </div>
                {linkedChildren.length === 0 ? (
                  <div className="lx-empty" style={{ padding: '24px 0' }}>
                    Chưa có phụ lục con nào được liên kết.
                  </div>
                ) : (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 8, maxHeight: 200, overflowY: 'auto' }}>
                    {linkedChildren.map(child => (
                      <div key={child.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', border: '1px solid #e6e8ea', padding: '10px 16px', borderRadius: 4, background: '#ffffff' }}>
                        <div style={{ minWidth: 0, flex: 1, paddingRight: 12 }}>
                          <div style={{ fontWeight: 600, fontSize: 13, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                            <span className="material-symbols-outlined" style={{ fontSize: 15, verticalAlign: 'middle', marginRight: 4 }}>description</span>
                            {child.name}
                          </div>
                          <div style={{ fontSize: 11, color: '#76777d', marginTop: 2 }}>Chứa {child.fields_count || 0} biến</div>
                        </div>
                        <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexShrink: 0 }}>
                          <button
                            type="button"
                            className={`lx-btn lx-btn-sm ${child.is_repeated === 1 ? 'lx-btn-secondary' : 'lx-btn-ghost'}`}
                            onClick={() => handleToggleRepeatedLink(child.id, child.is_repeated)}
                            title={child.is_repeated === 1 ? 'Lặp nhiều bản ghi' : 'Một bản ghi duy nhất'}
                          >
                            {child.is_repeated === 1 ? 'Lặp nhiều bản ghi' : 'Một bản ghi'}
                          </button>
                          <button
                            type="button"
                            className="lx-btn lx-btn-danger lx-btn-sm"
                            onClick={() => handleUnlinkChild(child.id)}
                          >
                            Hủy liên kết
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {/* Add link */}
              <div style={{ borderTop: '1px solid #e6e8ea', paddingTop: 16 }}>
                <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.06em', color: '#45464d', borderBottom: '1px solid #e6e8ea', paddingBottom: 8, marginBottom: 12 }}>
                  Thêm biểu mẫu phụ từ thư viện
                </div>
                <div style={{ position: 'relative', marginBottom: 12 }}>
                  <span className="material-symbols-outlined" style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', fontSize: 18, color: '#76777d' }}>search</span>
                  <input
                    type="text"
                    value={linkSearch}
                    onChange={(e) => setLinkSearch(e.target.value)}
                    placeholder="Tìm kiếm biểu mẫu con cần liên kết..."
                    className="lx-input"
                    style={{ paddingLeft: 36 }}
                  />
                </div>
                {(() => {
                  const filteredAvailable = templates.filter(t => {
                    if (t.id === linkingTemplate.id) return false;
                    if (t.status !== 'active') return false;
                    if (t.parent_template_id) return false;
                    if (linkedChildren.some(linked => linked.id === t.id)) return false;
                    if (linkSearch.trim()) {
                      return t.name.toLowerCase().includes(linkSearch.toLowerCase());
                    }
                    return true;
                  });
                  return filteredAvailable.length === 0 ? (
                    <div className="lx-empty" style={{ padding: '16px 0' }}>
                      Không có biểu mẫu nào hợp lệ để liên kết thêm.
                    </div>
                  ) : (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 6, maxHeight: 180, overflowY: 'auto' }}>
                      {filteredAvailable.map(avail => (
                        <div key={avail.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', border: '1px solid #e6e8ea', padding: '10px 14px', borderRadius: 4, background: '#ffffff' }}>
                          <div style={{ minWidth: 0, flex: 1, paddingRight: 12 }}>
                            <div style={{ fontWeight: 600, fontSize: 13, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{avail.name}</div>
                            <div style={{ fontSize: 11, color: '#76777d', marginTop: 2 }}>{avail.variables_count || 0} biến</div>
                          </div>
                          <button
                            className="lx-btn lx-btn-primary lx-btn-sm"
                            onClick={() => handleLinkChild(avail.id)}
                          >
                            <span className="material-symbols-outlined" style={{ fontSize: 15 }}>add</span>
                            Liên kết
                          </button>
                        </div>
                      ))}
                    </div>
                  );
                })()}
              </div>
            </div>
            <div className="lx-modal-footer">
              <button
                className="lx-btn lx-btn-primary"
                onClick={() => { setShowLinkModal(false); setLinkingTemplate(null); setLinkSearch(''); }}
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
