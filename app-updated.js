'use strict';

/* =========================================================
   تحديثات app.js — دمج نظام الحفظ والتحميل
   ========================================================= */

let DATA = null;
let META = null;
let SCENES = null;

let state = null;
let endingsReached = [];
let currentTheme = 'light';
let gameStartTime = null;

const PHASES = [
  'P0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8',
  'C9', 'C10', 'C11', 'C12', 'EVAL', 'ENDING', 'GROWTH_REPORT',
];
const STAT_ORDER = ['K', 'D', 'C', 'B', 'R', 'M', 'P'];
const ENDING_IDS = ['E1', 'E2', 'E3', 'E4'];

/* ==================== التهيئة ==================== */

window.addEventListener('DOMContentLoaded', init);

async function init() {
  try {
    const res = await fetch('./scenes.json');
    DATA = await res.json();
  } catch (err) {
    showFatalError();
    return;
  }
  META = DATA.meta;
  SCENES = DATA.scenes;
  initState();

  currentTheme = window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  applyTheme(currentTheme);

  wireEvents();
  renderStatsStrip();
  
  // تفعيل الحفظ التلقائي كل 5 دقائق
  enableAutoSave(300);
  
  // تحديث قائمة الحفظيات عند فتح النافذة
  document.getElementById('btn-open-log').addEventListener('click', updateSavesList);
}

function showFatalError() {
  document.body.innerHTML =
    '<div style="padding:2rem;text-align:center;font-family:sans-serif;direction:rtl">' +
    'تعذّر تحميل بيانات القصة (scenes.json). حدّث الصفحة وحاول مرة أخرى.' +
    '</div>';
}

function initState() {
  const stats = {};
  Object.keys(META.stats).forEach((code) => {
    stats[code] = META.stats[code].start;
  });
  state = {
    sceneId: 'P0',
    stats,
    flags: new Set(),
    flagCounts: {},
    variables: {},
    history: [],
  };
  gameStartTime = Date.now();
}

/* ==================== ربط الأحداث ==================== */

function wireEvents() {
  document.getElementById('btn-start').addEventListener('click', () => {
    initState();
    renderScene('P0');
  });

  document.querySelectorAll('[data-theme-toggle]').forEach((btn) => {
    btn.addEventListener('click', () => applyTheme(currentTheme === 'dark' ? 'light' : 'dark'));
  });

  document.getElementById('btn-open-log').addEventListener('click', openModal);
  document.getElementById('btn-open-log-report').addEventListener('click', openModal);
  document.getElementById('btn-close-log').addEventListener('click', closeModal);

  const overlay = document.getElementById('modal-log');
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) closeModal();
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && overlay.classList.contains('is-open')) closeModal();
  });

  // أزرار الحفظ والتحميل
  const saveBtn = document.getElementById('btn-save-game');
  if (saveBtn) {
    saveBtn.addEventListener('click', () => {
      const name = prompt('اسم الحفظية (اختياري):');
      if (saveGame(name)) {
        showNotification('تم حفظ اللعبة بنجاح ✓', 'success');
      } else {
        showNotification('فشل حفظ اللعبة ✗', 'error');
      }
    });
  }

  // استيراد ملف حفظية
  const importInput = document.getElementById('import-save-file');
  if (importInput) {
    importInput.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) {
        importSave(file);
      }
    });
  }
}

/* ==================== المظهر ==================== */

const SUN_ICON =
  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>';
const MOON_ICON =
  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="5"/><path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42"/></svg>';

function applyTheme(theme) {
  currentTheme = theme;
  document.documentElement.setAttribute('data-theme', theme);
  document.querySelectorAll('[data-theme-toggle]').forEach((btn) => {
    btn.setAttribute('aria-label', theme === 'dark' ? 'تبديل إلى الوضع الفاتح' : 'تبديل إلى الوضع الداكن');
    btn.innerHTML = theme === 'dark' ? MOON_ICON : SUN_ICON;
  });
}

/* ==================== أدوات الحالة ==================== */

function clampVal(v) {
  const range = (META && META.clampRange) || [0, 100];
  return Math.max(range[0], Math.min(range[1], v));
}

function applyEffects(effects) {
  if (!effects) return;
  Object.keys(effects).forEach((code) => {
    if (state.stats[code] === undefined) return;
    state.stats[code] = clampVal(state.stats[code] + effects[code]);
  });
}

function setFlag(flag) {
  if (!flag) return;
  state.flags.add(flag);
  state.flagCounts[flag] = (state.flagCounts[flag] || 0) + 1;
}

function getCtx() {
  return {
    stats: state.stats,
    flags: state.flags,
    flagCounts: state.flagCounts,
    variables: state.variables,
  };
}

/* ==================== مقيّم الشروط العام ==================== */

function evalCondition(cond, ctx) {
  if (!cond) return false;

  if (Object.prototype.hasOwnProperty.call(cond, 'default')) return !!cond.default;

  if (Array.isArray(cond.all)) return cond.all.every((c) => evalCondition(c, ctx));
  if (Array.isArray(cond.any)) return cond.any.some((c) => evalCondition(c, ctx));

  if (cond.stat) {
    const v = ctx.stats[cond.stat];
    if (typeof v !== 'number') return false;
    if (Object.prototype.hasOwnProperty.call(cond, 'gte')) return v >= cond.gte;
    if (Object.prototype.hasOwnProperty.call(cond, 'gt')) return v > cond.gt;
    if (Object.prototype.hasOwnProperty.call(cond, 'lte')) return v <= cond.lte;
    if (Object.prototype.hasOwnProperty.call(cond, 'lt')) return v < cond.lt;
    return false;
  }

  if (cond.flag) return ctx.flags.has(cond.flag);
  if (cond.notFlag) return !ctx.flags.has(cond.notFlag);
  if (Array.isArray(cond.anyFlag)) return cond.anyFlag.some((f) => ctx.flags.has(f));
  if (Array.isArray(cond.notAnyFlag)) return !cond.notAnyFlag.some((f) => ctx.flags.has(f));

  if (cond.variable && Object.prototype.hasOwnProperty.call(cond, 'equals')) {
    return ctx.variables[cond.variable] === cond.equals;
  }

  if (cond.repeatedFlag) {
    const count = ctx.flagCounts[cond.repeatedFlag] || 0;
    const countOk = count >= (cond.countGte || 0);
    const missingOk = cond.missingFlag ? !ctx.flags.has(cond.missingFlag) : true;
    return countOk && missingOk;
  }

  return false;
}

function evalVariantCond(variant, stats) {
  const val = stats[variant.axis];
  if (typeof val !== 'number') return false;
  switch (variant.op) {
    case 'gte': return val >= variant.value;
    case 'gt': return val > variant.value;
    case 'lte': return val <= variant.value;
    case 'lt': return val < variant.value;
    default: return false;
  }
}

/* ==================== تنقّل الشاشات ==================== */

function showScreen(name) {
  document.getElementById('screen-title').classList.toggle('is-active', name === 'title');
  document.getElementById('screen-story').classList.toggle('is-active', name === 'story');
  document.getElementById('screen-report').classList.toggle('is-active', name === 'report');
}

function renderScene(sceneId) {
  state.sceneId = sceneId;
  updateProgress(sceneId);
  renderStatsStrip();

  const scene = SCENES[sceneId];
  if (sceneId === 'EVAL') {
    showScreen('story');
    renderEvalScene(scene);
  } else if (ENDING_IDS.includes(sceneId)) {
    showScreen('story');
    renderEndingScene(scene);
  } else if (sceneId === 'GROWTH_REPORT') {
    showScreen('report');
    renderReport();
  } else {
    showScreen('story');
    renderChapterScene(scene);
  }
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function goToNext(nextId) {
  renderScene(nextId);
}

function updateProgress(sceneId) {
  let idx = PHASES.indexOf(sceneId);
  if (idx === -1) {
    idx = ENDING_IDS.includes(sceneId) ? PHASES.indexOf('ENDING') : PHASES.length - 1;
  }
  const pct = Math.round((idx / (PHASES.length - 1)) * 100);
  const fill = document.getElementById('progress-fill');
  if (fill) fill.style.width = pct + '%';
}

function actLabelFor(sceneId) {
  if (META.acts) {
    const act = META.acts.find((a) => a.chapters.includes(sceneId));
    if (act) return act.titleAr;
  }
  if (sceneId === 'EVAL') return 'عقدة التقييم';
  if (ENDING_IDS.includes(sceneId)) return 'مشهد النهاية';
  return '';
}

function setHeader(scene) {
  document.getElementById('tag-act').textContent = actLabelFor(scene.id);
  document.getElementById('tag-chapter').textContent = scene.titleAr;
}

function setSceneMeta(scene) {
  document.getElementById('scene-title').textContent = scene.titleAr;
  document.getElementById('scene-location').textContent = scene.location || '';
}

/* ==================== عرض الوقائع (beats) ==================== */

function renderBeats(beats) {
  const feed = document.getElementById('beat-feed');
  feed.innerHTML = '';
  (beats || []).forEach((b, i) => {
    let el;
    if (b.type === 'stage') {
      el = document.createElement('p');
      el.className = 'beat beat-stage';
      el.textContent = b.text;
    } else if (b.type === 'line') {
      el = document.createElement('div');
      el.className = 'beat beat-line';
      const sp = document.createElement('span');
      sp.className = 'beat-line-speaker';
      sp.textContent = b.speaker || '';
      const txt = document.createElement('p');
      txt.className = 'beat-line-text';
      txt.textContent = b.text;
      el.appendChild(sp);
      el.appendChild(txt);
    } else if (b.type === 'insight') {
      el = document.createElement('p');
      el.className = 'beat beat-insight';
      el.textContent = b.text;
    } else {
      el = document.createElement('p');
      el.className = 'beat beat-narr';
      el.textContent = b.text || '';
    }
    el.style.animationDelay = Math.min(i * 70, 700) + 'ms';
    feed.appendChild(el);
  });
}

function collectOpeningBeats(scene, ctx) {
  const beats = [];
  (scene.conditionalOpenings || []).forEach((co) => {
    if (evalCondition(co.condition, ctx)) beats.push(...co.beats);
  });
  return beats;
}

function collectConditionalBeats(scene, ctx) {
  const beats = [];
  (scene.conditionalBeats || []).forEach((cb) => {
    if (evalCondition(cb.condition, ctx)) beats.push(...cb.beats);
  });
  return beats;
}

/* ==================== مشاهد الفصول ==================== */

function renderChapterScene(scene) {
  setHeader(scene);
  setSceneMeta(scene);
  const ctx = getCtx();
  const allBeats = [
    ...collectOpeningBeats(scene, ctx),
    ...(scene.beats || []),
    ...collectConditionalBeats(scene, ctx),
  ];
  renderBeats(allBeats);

  if (scene.decision) {
    renderDecision(scene);
  } else {
    renderContinuePrompt(scene.transitionText, 'متابعة', () => goToNext(scene.next));
  }
}

function renderDecision(scene) {
  const zone = document.getElementById('interaction-zone');
  zone.innerHTML = '';

  const card = document.createElement('div');
  card.className = 'decision-card';

  const title = document.createElement('h3');
  title.className = 'decision-title';
  title.textContent = scene.decision.promptTitle;

  const prompt = document.createElement('p');
  prompt.className = 'decision-prompt';
  prompt.textContent = scene.decision.promptText;

  const optsWrap = document.createElement('div');
  optsWrap.className = 'decision-options';
  optsWrap.setAttribute('role', 'group');
  optsWrap.setAttribute('aria-label', scene.decision.promptTitle);

  scene.decision.options.forEach((opt) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'option-btn';

    const txt = document.createElement('span');
    txt.className = 'option-btn-text';
    txt.textContent = opt.button;
    btn.appendChild(txt);

    if (opt.tagsAr) {
      const tags = document.createElement('span');
      tags.className = 'option-btn-tags';
      tags.textContent = opt.tagsAr;
      btn.appendChild(tags);
    }

    btn.addEventListener('click', () => handleOptionClick(scene, opt));
    optsWrap.appendChild(btn);
  });

  card.appendChild(title);
  card.appendChild(prompt);
  card.appendChild(optsWrap);
  zone.appendChild(card);
}

function handleOptionClick(scene, opt) {
  document.querySelectorAll('.option-btn').forEach((b) => { b.disabled = true; });

  applyEffects(opt.effects);
  if (opt.flag) setFlag(opt.flag);
  if (opt.setsVariable) Object.assign(state.variables, opt.setsVariable);

  if (scene.id === 'C10' && Array.isArray(scene.modifiers)) {
    scene.modifiers.forEach((mod) => {
      if (mod.appliesToAllOptions && Array.isArray(mod.appliesWhenAnyFlag)) {
        if (mod.appliesWhenAnyFlag.some((f) => state.flags.has(f))) {
          applyEffects(mod.extraEffects);
        }
      }
    });
  }

  state.history.push({
    sceneId: scene.id,
    sceneTitle: scene.titleAr,
    button: opt.button,
    immediateResponse: opt.immediateResponse,
  });

  renderStatsStrip();
  renderAfterChoice(scene, opt);
}

function renderAfterChoice(scene, opt) {
  const zone = document.getElementById('interaction-zone');
  zone.innerHTML = '';

  const card = document.createElement('div');
  card.className = 'decision-card';

  const resp = document.createElement('p');
  resp.className = 'immediate-response';
  resp.textContent = opt.immediateResponse;
  card.appendChild(resp);

  if (scene.transitionText) {
    const trans = document.createElement('p');
    trans.className = 'chapter-transition';
    trans.textContent = scene.transitionText;
    card.appendChild(trans);
  }

  const row = document.createElement('div');
  row.className = 'continue-row';
  row.style.marginTop = 'var(--space-4)';
  const btn = document.createElement('button');
  btn.className = 'btn btn-primary';
  btn.type = 'button';
  btn.textContent = 'متابعة';
  btn.addEventListener('click', () => goToNext(scene.next));
  row.appendChild(btn);
  card.appendChild(row);

  zone.appendChild(card);
}

function renderContinuePrompt(transitionText, label, onClick) {
  const zone = document.getElementById('interaction-zone');
  zone.innerHTML = '';

  if (transitionText) {
    const trans = document.createElement('p');
    trans.className = 'chapter-transition';
    trans.textContent = transitionText;
    zone.appendChild(trans);
  }

  const row = document.createElement('div');
  row.className = 'continue-row';
  const btn = document.createElement('button');
  btn.className = 'btn btn-primary';
  btn.type = 'button';
  btn.textContent = label;
  btn.addEventListener('click', onClick);
  row.appendChild(btn);
  zone.appendChild(row);
}

/* ==================== عقدة التقييم ==================== */

function renderEvalScene(scene) {
  setHeader(scene);
  setSceneMeta(scene);
  renderBeats(scene.beats);
  renderContinuePrompt(scene.transitionText, 'احتساب النهاية', resolveEnding);
}

function resolveEnding() {
  const ctx = getCtx();
  const resolution = SCENES.EVAL.resolution;
  let winner = null;
  for (const eid of resolution.priority) {
    const cond = resolution.conditions[eid];
    if (evalCondition(cond, ctx)) {
      winner = eid;
      break;
    }
  }
  if (!winner) winner = 'E4';
  endingsReached.push(winner);
  goToNext(winner);
}

/* ==================== مشاهد النهايات ==================== */

function renderEndingScene(scene) {
  setHeader(scene);
  setSceneMeta(scene);

  const ctx = getCtx();
  let beats = [...scene.beats];
  const mode = scene.variantsCombineMode;

  if (mode === 'independentInsertions') {
    (scene.variants || []).forEach((v) => {
      if (evalVariantCond(v, ctx.stats)) beats.push({ type: 'insight', text: v.textAr });
    });
  } else if (mode === 'exclusiveSceneReplace') {
    const match = (scene.variants || []).find((v) => evalVariantCond(v, ctx.stats));
    if (match && Array.isArray(match.beats)) beats = [...beats, ...match.beats];
  } else if (mode === 'reasonDependentInsertion') {
    (scene.variants || []).forEach((v) => {
      if (evalCondition(v.reason, ctx)) beats.push({ type: 'insight', text: v.textAr });
    });
  }

  renderBeats(beats);

  const zone = document.getElementById('interaction-zone');
  zone.innerHTML = '';

  if (scene.taglineAr) {
    const tagline = document.createElement('p');
    tagline.className = 'ending-tagline';
    tagline.textContent = scene.taglineAr;
    zone.appendChild(tagline);
  }

  if (scene.transitionText) {
    const trans = document.createElement('p');
    trans.className = 'chapter-transition';
    trans.textContent = scene.transitionText;
    zone.appendChild(trans);
  }

  const row = document.createElement('div');
  row.className = 'continue-row';
  const btn = document.createElement('button');
  btn.className = 'btn btn-primary';
  btn.type = 'button';
  btn.textContent = 'متابعة إلى تقرير النمو';
  btn.addEventListener('click', () => goToNext(scene.next || 'GROWTH_REPORT'));
  row.appendChild(btn);
  zone.appendChild(row);
}

/* ==================== تقرير النمو ==================== */

function renderReport() {
  const scene = SCENES.GROWTH_REPORT;
  const main = document.getElementById('report-main');
  main.innerHTML = '';

  const header = document.createElement('div');
  header.className = 'report-header';
  const h2 = document.createElement('h2');
  h2.className = 'report-title';
  h2.textContent = scene.titleAr;
  const disclaimer = document.createElement('p');
  disclaimer.className = 'report-disclaimer';
  disclaimer.textContent = scene.disclaimerAr;
  header.appendChild(h2);
  header.appendChild(disclaimer);
  main.appendChild(header);

  const growthVal = Math.round(
    (state.stats.K + state.stats.D + state.stats.C + state.stats.B + state.stats.R) / 5
  );
  const band =
    scene.bands.find((b) => growthVal >= b.range[0] && growthVal <= b.range[1]) ||
    scene.bands[scene.bands.length - 1];

  const giCard = document.createElement('div');
  giCard.className = 'growth-index-card';

  const giVal = document.createElement('div');
  giVal.className = 'growth-index-value';
  giVal.textContent = String(growthVal);

  const giLabel = document.createElement('div');
  giLabel.className = 'growth-index-label';
  giLabel.textContent = scene.growthIndex.labelAr;

  const bandLabel = document.createElement('div');
  bandLabel.className = 'growth-band-label';
  bandLabel.textContent = band.labelAr;

  const bandText = document.createElement('p');
  bandText.className = 'growth-band-text';
  bandText.textContent = band.textAr;

  const excludeNote = document.createElement('p');
  excludeNote.className = 'trait-context-note';
  excludeNote.style.marginTop = 'var(--space-2)';
  excludeNote.textContent = scene.growthIndex.excludeNoteAr;

  giCard.appendChild(giVal);
  giCard.appendChild(giLabel);
  giCard.appendChild(bandLabel);
  giCard.appendChild(bandText);
  giCard.appendChild(excludeNote);
  main.appendChild(giCard);

  const traitsGrid = document.createElement('div');
  traitsGrid.className = 'traits-grid';
  scene.traitDisplay.forEach((t) => {
    const isContext = t.code === 'M' || t.code === 'P';
    const card = document.createElement('div');
    card.className = 'trait-card' + (isContext ? ' is-context' : '');

    const top = document.createElement('div');
    top.className = 'trait-card-top';
    const name = document.createElement('span');
    name.className = 'trait-name';
    name.textContent = t.labelAr;
    const val = document.createElement('span');
    val.className = 'trait-value';
    val.textContent = String(Math.round(state.stats[t.code]));
    top.appendChild(name);
    top.appendChild(val);
    card.appendChild(top);

    if (isContext) {
      const note = document.createElement('p');
      note.className = 'trait-context-note';
      note.textContent = t.displayAr;
      card.appendChild(note);
    } else {
      const track = document.createElement('div');
      track.className = 'trait-bar-track';
      const fill = document.createElement('div');
      fill.className = 'trait-bar-fill';
      fill.style.width = clampVal(state.stats[t.code]) + '%';
      track.appendChild(fill);
      card.appendChild(track);
    }
    traitsGrid.appendChild(card);
  });
  main.appendChild(traitsGrid);

  const flagsSection = document.createElement('div');
  flagsSection.className = 'report-section';
  const flagsTitle = document.createElement('h3');
  flagsTitle.className = 'report-section-title';
  flagsTitle.textContent = 'القرارات التي شكّلت قصتك';
  flagsSection.appendChild(flagsTitle);
  const flagsList = document.createElement('div');
  flagsList.className = 'flags-list';
  if (state.flags.size === 0) {
    const empty = document.createElement('p');
    empty.className = 'empty-note';
    empty.textContent = 'لم تُسجَّل أعلام في هذا المسار.';
    flagsList.appendChild(empty);
  } else {
    [...state.flags].forEach((f) => {
      const chip = document.createElement('span');
      chip.className = 'flag-chip';
      chip.textContent = f;
      flagsList.appendChild(chip);
    });
  }
  flagsSection.appendChild(flagsList);
  main.appendChild(flagsSection);

  const noJudgment = (scene.extraElements || []).find((e) => e.id === 'noJudgmentRule');
  if (noJudgment) {
    const note = document.createElement('p');
    note.className = 'no-judgment-note';
    note.textContent = noJudgment.textAr;
    main.appendChild(note);
  }

  if (endingsReached.length > 0) {
    const histSection = document.createElement('div');
    histSection.className = 'report-section';
    histSection.style.marginTop = 'var(--space-6)';
    const histTitle = document.createElement('h3');
    histTitle.className = 'report-section-title';
    histTitle.textContent = 'النهايات التي وصلتَ إليها في هذه الجلسة';
    const histText = document.createElement('p');
    histText.className = 'endings-history';
    histText.textContent = endingsReached.map((e) => SCENES[e].titleAr).join(' · ');
    histSection.appendChild(histTitle);
    histSection.appendChild(histText);
    main.appendChild(histSection);
  }

  const actions = document.createElement('div');
  actions.className = 'report-actions';
  
  const replayBtn = document.createElement('button');
  replayBtn.className = 'btn btn-primary';
  replayBtn.type = 'button';
  replayBtn.textContent = scene.replayButton.labelAr;
  replayBtn.addEventListener('click', () => {
    initState();
    renderScene('P0');
  });
  actions.appendChild(replayBtn);
  
  const exportBtn = document.createElement('button');
  exportBtn.className = 'btn btn-secondary';
  exportBtn.type = 'button';
  exportBtn.textContent = '📥 تصدير الحفظية';
  exportBtn.addEventListener('click', () => {
    const saves = getAllSaves();
    if (saves.length > 0) {
      exportSave(saves[0].id);
      showNotification('تم تصدير الحفظية', 'success');
    }
  });
  actions.appendChild(exportBtn);
  
  main.appendChild(actions);
}

/* ==================== شريط تتبّع السمات ==================== */

function renderStatsStrip() {
  const wrap = document.getElementById('stats-strip-inner');
  if (!wrap || !state) return;
  wrap.innerHTML = '';

  STAT_ORDER.forEach((code) => {
    const meta = META.stats[code];
    if (!meta) return;
    const isContext = meta.visibleDuringStory === 'contextOnly' || meta.visibleDuringStory === false && (code === 'M' || code === 'P');
    const val = Math.round(state.stats[code]);

    const chip = document.createElement('div');
    chip.className = 'stat-chip' + (code === 'M' || code === 'P' ? ' is-context' : '');

    const top = document.createElement('div');
    top.className = 'stat-chip-top';
    const name = document.createElement('span');
    name.className = 'stat-chip-name';
    name.textContent = meta.nameAr;
    const value = document.createElement('span');
    value.className = 'stat-chip-value';
    value.textContent = String(val);
    top.appendChild(name);
    top.appendChild(value);
    chip.appendChild(top);

    const track = document.createElement('div');
    track.className = 'stat-bar-track';
    const fill = document.createElement('div');
    fill.className = 'stat-bar-fill';
    fill.style.width = clampVal(val) + '%';
    track.appendChild(fill);
    chip.appendChild(track);

    wrap.appendChild(chip);
  });
}

/* ==================== نافذة الأعلام/السجل ==================== */

function openModal() {
  renderModal();
  document.getElementById('modal-log').classList.add('is-open');
}

function closeModal() {
  document.getElementById('modal-log').classList.remove('is-open');
}

function renderModal() {
  const flagsWrap = document.getElementById('modal-flags');
  flagsWrap.innerHTML = '';
  if (state.flags.size === 0) {
    const empty = document.createElement('p');
    empty.className = 'empty-note';
    empty.textContent = 'لا توجد أعلام مفعّلة بعد.';
    flagsWrap.appendChild(empty);
  } else {
    [...state.flags].forEach((f) => {
      const chip = document.createElement('span');
      chip.className = 'flag-chip';
      chip.textContent = f + (state.flagCounts[f] > 1 ? ' ×' + state.flagCounts[f] : '');
      flagsWrap.appendChild(chip);
    });
  }

  const varsWrap = document.getElementById('modal-vars');
  varsWrap.innerHTML = '';
  const varKeys = Object.keys(state.variables);
  if (varKeys.length === 0) {
    const empty = document.createElement('p');
    empty.className = 'empty-note';
    empty.textContent = 'لا توجد متغيرات محفوظة بعد.';
    varsWrap.appendChild(empty);
  } else {
    varKeys.forEach((k) => {
      const row = document.createElement('div');
      row.className = 'vars-list-row';
      const kEl = document.createElement('span');
      kEl.textContent = k;
      const vEl = document.createElement('span');
      vEl.textContent = state.variables[k];
      row.appendChild(kEl);
      row.appendChild(vEl);
      varsWrap.appendChild(row);
    });
  }

  const histWrap = document.getElementById('modal-history');
  histWrap.innerHTML = '';
  if (state.history.length === 0) {
    const empty = document.createElement('p');
    empty.className = 'empty-note';
    empty.textContent = 'لم تُتَّخذ أي قرارات بعد.';
    histWrap.appendChild(empty);
  } else {
    state.history.forEach((h) => {
      const item = document.createElement('div');
      item.className = 'history-item';
      const sceneEl = document.createElement('div');
      sceneEl.className = 'history-item-scene';
      sceneEl.textContent = h.sceneTitle;
      const choiceEl = document.createElement('div');
      choiceEl.className = 'history-item-choice';
      choiceEl.textContent = h.button;
      item.appendChild(sceneEl);
      item.appendChild(choiceEl);
      histWrap.appendChild(item);
    });
  }

  // إضافة قسم الحفظيات
  const savesWrap = document.getElementById('modal-saves');
  if (savesWrap) {
    updateSavesList();
  }
}
