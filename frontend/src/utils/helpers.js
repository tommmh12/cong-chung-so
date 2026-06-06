export const API_BASE = 'http://localhost:5000/api';
export const SUPPORTED_WORD_EXTENSIONS = ['doc', 'docx'];
export const MAX_UPLOAD_FILES = 10;
export const MAX_UPLOAD_FILE_SIZE = 1024 * 1024;

export const getFileExtension = (fileName) => fileName.split('.').pop().toLowerCase();
export const isSupportedWordFile = (fileName) => SUPPORTED_WORD_EXTENSIONS.includes(getFileExtension(fileName));
export const getWordBaseName = (fileName) => fileName.replace(/\.(docx|doc)$/i, '');

export const generateId = () => {
  return `temp-${Date.now()}-${Math.random().toString(36).substring(2, 11)}`;
};

export const getFieldStep = (field) => {
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

export const getLogicalGroupName = (field) => {
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

export const getGroupedFields = (fieldsList) => {
  const groups = {};
  fieldsList.forEach(field => {
    const gName = getLogicalGroupName(field);
    if (!groups[gName]) groups[gName] = [];
    groups[gName].push(field);
  });
  return groups;
};
