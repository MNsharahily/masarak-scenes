'use strict';

/* =========================================================
   نظام حفظ/تحميل التقدم — saveload.js
   ========================================================= */

const STORAGE_PREFIX = 'masarak_';
const MAX_SAVES = 10;
const SAVE_VERSION = 1;

/**
 * حفظ حالة اللعبة الحالية
 * @param {string} name - اسم الحفظية (اختياري)
 * @returns {boolean} - true إذا نجح الحفظ
 */
function saveGame(name = null) {
  try {
    if (!state) {
      console.error('لا توجد حالة لعبة لحفظها');
      return false;
    }

    const saveData = {
      version: SAVE_VERSION,
      name: name || `جلسة ${new Date().toLocaleDateString('ar')}`,
      timestamp: new Date().toISOString(),
      sceneId: state.sceneId,
      stats: { ...state.stats },
      flags: Array.from(state.flags),
      flagCounts: { ...state.flagCounts },
      variables: { ...state.variables },
      history: [...state.history],
      endingsReached: [...endingsReached],
      currentTheme: currentTheme,
      playDuration: getPlayDuration(),
      growthIndex: calculateGrowthIndex(state.stats)
    };

    const saveId = generateSaveId();
    const key = `${STORAGE_PREFIX}save_${saveId}`;

    localStorage.setItem(key, JSON.stringify(saveData));
    
    // تحديث قائمة الحفظيات
    updateSavesList();
    
    return true;
  } catch (error) {
    console.error('خطأ في حفظ اللعبة:', error);
    return false;
  }
}

/**
 * تحميل حالة لعبة محفوظة
 * @param {string} saveId - معرّف الحفظية
 * @returns {boolean} - true إذا نجح التحميل
 */
function loadGame(saveId) {
  try {
    const key = `${STORAGE_PREFIX}save_${saveId}`;
    const saved = localStorage.getItem(key);

    if (!saved) {
      console.error('الحفظية غير موجودة');
      return false;
    }

    const saveData = JSON.parse(saved);

    // التحقق من نسخة الحفظية
    if (saveData.version !== SAVE_VERSION) {
      console.warn('إصدار الحفظية قديم، قد تحدث مشاكل توافقية');
    }

    // استعادة الحالة
    state.sceneId = saveData.sceneId;
    state.stats = { ...saveData.stats };
    state.flags = new Set(saveData.flags);
    state.flagCounts = { ...saveData.flagCounts };
    state.variables = { ...saveData.variables };
    state.history = [...saveData.history];
    endingsReached = [...saveData.endingsReached];
    currentTheme = saveData.currentTheme || 'light';

    // تطبيق المظهر المحفوظ
    applyTheme(currentTheme);

    // تحديث آخر وقت تحميل
    const newKey = `${STORAGE_PREFIX}save_${saveId}`;
    saveData.lastLoaded = new Date().toISOString();
    localStorage.setItem(newKey, JSON.stringify(saveData));

    return true;
  } catch (error) {
    console.error('خطأ في تحميل اللعبة:', error);
    return false;
  }
}

/**
 * حذف حفظية
 * @param {string} saveId - معرّف الحفظية
 */
function deleteSave(saveId) {
  try {
    const key = `${STORAGE_PREFIX}save_${saveId}`;
    localStorage.removeItem(key);
    updateSavesList();
  } catch (error) {
    console.error('خطأ في حذف الحفظية:', error);
  }
}

/**
 * الحصول على قائمة جميع الحفظيات
 * @returns {Array} - مصفوفة الحفظيات مرتبة بالتاريخ
 */
function getAllSaves() {
  const saves = [];
  
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key && key.startsWith(STORAGE_PREFIX + 'save_')) {
      try {
        const data = JSON.parse(localStorage.getItem(key));
        saves.push({
          id: key.replace(STORAGE_PREFIX + 'save_', ''),
          ...data
        });
      } catch (e) {
        console.warn('فشل في تحليل حفظية:', key);
      }
    }
  }

  // ترتيب حسب التاريخ (الأحدث أولاً)
  return saves.sort((a, b) => 
    new Date(b.timestamp) - new Date(a.timestamp)
  );
}

/**
 * تصدير الحفظية كملف JSON
 * @param {string} saveId - معرّف الحفظية
 */
function exportSave(saveId) {
  const key = `${STORAGE_PREFIX}save_${saveId}`;
  const saveData = localStorage.getItem(key);

  if (!saveData) {
    console.error('الحفظية غير موجودة');
    return;
  }

  const blob = new Blob([saveData], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `masarak_save_${saveId}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

/**
 * استيراد حفظية من ملف
 * @param {File} file - ملف JSON
 */
function importSave(file) {
  const reader = new FileReader();
  
  reader.onload = (e) => {
    try {
      const saveData = JSON.parse(e.target.result);
      
      // التحقق من صحة البيانات
      if (!validateSaveData(saveData)) {
        throw new Error('بيانات الحفظية غير صحيحة');
      }

      const saveId = generateSaveId();
      const key = `${STORAGE_PREFIX}save_${saveId}`;
      
      localStorage.setItem(key, JSON.stringify(saveData));
      updateSavesList();
      
      showNotification('تم استيراد الحفظية بنجاح', 'success');
    } catch (error) {
      console.error('خطأ في استيراد الحفظية:', error);
      showNotification('فشل استيراد الحفظية', 'error');
    }
  };

  reader.readAsText(file);
}

/**
 * التحقق من صحة بيانات الحفظية
 * @param {Object} data - بيانات الحفظية
 * @returns {boolean}
 */
function validateSaveData(data) {
  return (
    data.version === SAVE_VERSION &&
    data.sceneId &&
    data.stats &&
    data.flags &&
    Array.isArray(data.flags)
  );
}

/**
 * حساب معامل النمو من الإحصائيات
 * @param {Object} stats - الإحصائيات
 * @returns {number}
 */
function calculateGrowthIndex(stats) {
  const growthStats = ['K', 'D', 'C', 'B', 'R'];
  const total = growthStats.reduce((sum, code) => sum + (stats[code] || 0), 0);
  return Math.round(total / growthStats.length);
}

/**
 * حساب مدة اللعب
 * @returns {number} - المدة بالثواني
 */
function getPlayDuration() {
  if (!window.gameStartTime) {
    window.gameStartTime = Date.now();
  }
  return Math.floor((Date.now() - window.gameStartTime) / 1000);
}

/**
 * تنسيق مدة اللعب
 * @param {number} seconds - المدة بالثواني
 * @returns {string}
 */
function formatDuration(seconds) {
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const secs = seconds % 60;

  if (hours > 0) {
    return `${hours}س ${minutes}د`;
  }
  return `${minutes}د ${secs}ث`;
}

/**
 * توليد معرف فريد للحفظية
 * @returns {string}
 */
function generateSaveId() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

/**
 * تحديث قائمة الحفظيات المعروضة
 */
function updateSavesList() {
  const savesList = getAllSaves();
  const container = document.getElementById('saves-list-container');

  if (!container) return;

  container.innerHTML = '';

  if (savesList.length === 0) {
    container.innerHTML = '<p class="empty-note">لا توجد حفظيات محفوظة</p>';
    return;
  }

  savesList.forEach((save) => {
    const card = createSaveCard(save);
    container.appendChild(card);
  });
}

/**
 * إنشاء بطاقة حفظية
 * @param {Object} save - بيانات الحفظية
 * @returns {HTMLElement}
 */
function createSaveCard(save) {
  const card = document.createElement('div');
  card.className = 'save-card';

  const saveDate = new Date(save.timestamp);
  const duration = formatDuration(save.playDuration || 0);

  card.innerHTML = `
    <div class="save-card-content">
      <div class="save-card-header">
        <h3 class="save-card-name">${escapeHtml(save.name)}</h3>
        <span class="save-card-date">${saveDate.toLocaleDateString('ar')}</span>
      </div>
      
      <div class="save-card-meta">
        <span class="save-meta-item">
          <span class="save-meta-label">المشهد:</span>
          <span class="save-meta-value">${save.sceneId}</span>
        </span>
        <span class="save-meta-item">
          <span class="save-meta-label">الوقت:</span>
          <span class="save-meta-value">${duration}</span>
        </span>
        <span class="save-meta-item">
          <span class="save-meta-label">النمو:</span>
          <span class="save-meta-value">${save.growthIndex || 0}</span>
        </span>
      </div>

      <div class="save-card-stats">
        ${createStatsBars(save.stats)}
      </div>
    </div>

    <div class="save-card-actions">
      <button class="btn btn-sm btn-primary" onclick="loadGameAndClose('${save.id}')">
        تحميل
      </button>
      <button class="btn btn-sm btn-secondary" onclick="exportSave('${save.id}')">
        تصدير
      </button>
      <button class="btn btn-sm btn-danger" onclick="deleteAndRefresh('${save.id}')">
        حذف
      </button>
    </div>
  `;

  return card;
}

/**
 * إنشاء أشرطة الإحصائيات
 * @param {Object} stats - الإحصائيات
 * @returns {string}
 */
function createStatsBars(stats) {
  const shortOrder = ['K', 'D', 'C', 'B', 'R'];
  
  return shortOrder.map(code => {
    const meta = META.stats[code];
    const value = Math.round(stats[code] || 0);
    
    return `
      <div class="stat-mini-bar">
        <div class="stat-mini-label">${meta.nameAr.substring(0, 2)}</div>
        <div class="stat-mini-track">
          <div class="stat-mini-fill" style="width: ${value}%"></div>
        </div>
        <div class="stat-mini-value">${value}</div>
      </div>
    `;
  }).join('');
}

/**
 * دالة مساعدة: تحميل لعبة وإغلاق النافذة
 * @param {string} saveId - معرّف الحفظية
 */
function loadGameAndClose(saveId) {
  if (loadGame(saveId)) {
    renderScene(state.sceneId);
    closeModal();
    showNotification('تم تحميل الحفظية بنجاح', 'success');
  } else {
    showNotification('فشل تحميل الحفظية', 'error');
  }
}

/**
 * دالة مساعدة: حذف وتحديث القائمة
 * @param {string} saveId - معرّف الحفظية
 */
function deleteAndRefresh(saveId) {
  if (confirm('هل أنت متأكد من حذف هذه الحفظية؟')) {
    deleteSave(saveId);
    updateSavesList();
    showNotification('تم حذف الحفظية', 'info');
  }
}

/**
 * تنظيف النصوص من الأحرف الخطرة
 * @param {string} text - النص
 * @returns {string}
 */
function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

/**
 * عرض إخطار
 * @param {string} message - الرسالة
 * @param {string} type - نوع الإخطار (success, error, info)
 */
function showNotification(message, type = 'info') {
  const notification = document.createElement('div');
  notification.className = `notification notification-${type}`;
  notification.textContent = message;
  
  document.body.appendChild(notification);
  
  setTimeout(() => {
    notification.classList.add('show');
  }, 10);

  setTimeout(() => {
    notification.classList.remove('show');
    setTimeout(() => notification.remove(), 300);
  }, 3000);
}

/**
 * دعم الحفظ التلقائي
 */
function enableAutoSave(intervalSeconds = 300) {
  setInterval(() => {
    if (state && state.sceneId !== 'P0') {
      const saved = saveGame('حفظ تلقائي');
      console.log('تم الحفظ التلقائي:', saved);
    }
  }, intervalSeconds * 1000);
}
