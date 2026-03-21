// ==========================================================
// script.js - スト6コンボ入力ツール
// ==========================================================
// 目的: 方向/技ボタンからコンボ文字列を構築し、localStorage/JSONで永続化
// 設定: config.jsonから読み込み、UIを動的に生成
// 注意: index.htmlのIDに依存するグローバルスコープ

// ==========================================================
// Supabase接続設定
// ==========================================================

const SUPABASE_URL = 'https://kctyakepitpfvtrpsnut.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_-toqPQ-hYYayFrUh8nLvxg_9Lnc3i3z';
const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// ==========================================================
// 設定読み込み
// ==========================================================

let config = null; // config.jsonから読み込んだ設定

// config.jsonを読み込む
async function loadConfig() {
  try {
    console.log('Loading config.json...');
    const response = await fetch('config.json');
    console.log('Fetch response:', response.status, response.statusText);
    if (!response.ok) throw new Error(`Failed to load config.json: ${response.status} ${response.statusText}`);
    config = await response.json();
    console.log('Config loaded successfully:', config);
    return config;
  } catch (error) {
    console.error('Error loading config:', error);
    console.error('Error details:', error.message, error.stack);
    alert(`設定ファイルの読み込みに失敗しました: ${error.message}\n\nデフォルト設定で起動します。\n\n※ローカルファイルから開いている場合は、ローカルサーバーを使用してください。`);
    // フォールバック: 最小限のデフォルト設定
    config = {
      arrowMap: { '1': '↙', '2': '↓', '3': '↘', '4': '←', '5': 'N', '6': '→', '7': '↖', '8': '↑', '9': '↗' },
      buttons: { specialDirections: [], specialDirectionsExtra: [], specialCommandExtra: [], specialCommands: [], classicAttacks: [], modernAttacks: [], specialInputs: [] },
      filters: { screenPosition: [], startupTech: [], condition: [] },
      localStorage: { combosKey: 'sf6_savedCombos_v1', codeCounterKey: 'sf6_comboCodeCounter_v1', damageKey: 'sf6_damage', sa1DamageKey: 'sf6_sa1Damage', sa2DamageKey: 'sf6_sa2Damage', sa3DamageKey: 'sf6_sa3Damage', sortKey: 'sf6_sortKey', filterKey: 'sf6_filterKey', screenPositionFilterKey: 'sf6_screenPositionFilter', startupTechFilterKey: 'sf6_startupTechFilter' },
      ui: { titlePrefix: 'C', titleDigits: 4, outputPlaceholder: 'ここにコンボが表示されます', memoPlaceholder: 'メモを入力', damageMaxLength: 5 }
    };
    console.log('Using default config:', config);
    return config;
  }
}

// ==========================================================
// グローバル変数
// ==========================================================

// DOM要素の参照
const directionGrid = document.getElementById('directionGrid');
const specialDirections = document.getElementById('specialDirections');
const classicButtons = document.getElementById('classicButtons');
const modernButtons = document.getElementById('modernButtons');
const saButtons = document.getElementById('saButtons');
const saButtonsModern = document.getElementById('saButtonsModern');
const specialInputs = document.getElementById('specialInputs');
const specialCommandButtons = document.getElementById('specialCommandButtons');
const specialCommandButtonsModern = document.getElementById('specialCommandButtonsModern');
const outputBox = document.getElementById('output');
const memoBox = document.getElementById('memo');
const comboList = document.getElementById('comboList');
const damageBox = document.getElementById('damage');
const sa1DamageBox = document.getElementById('sa1Damage');
const sa2DamageBox = document.getElementById('sa2Damage');
const sa3DamageBox = document.getElementById('sa3Damage');
const outputCodeEl = document.getElementById('outputCode');

// 設定値（config.jsonから初期化）
let LS_COMBOS_KEY, LS_CODE_COUNTER_KEY, LS_DAMAGE_KEY, LS_SA1_DAMAGE_KEY, LS_SA2_DAMAGE_KEY, LS_SA3_DAMAGE_KEY;
let LS_SORT_KEY, LS_FILTER_KEY, LS_SCREEN_POSITION_FILTER_KEY, LS_STARTUP_TECH_FILTER_KEY;
let ARROW_MAP;

// アプリケーション状態（config読み込み後に初期化）
let sortKey, filterKey, screenPositionFilter, startupTechFilter;

// エディタ状態
let output = '';
let history = [];
let lastWasDirection = false;
let savedCombos = [];
let isUpdatingOutputBox = false; // プログラム的なoutputBox更新時のイベント制御用

// ==========================================================
// localStorage永続化
// ==========================================================

// 保存
function saveCombosToLocal() {
  try {
    localStorage.setItem(LS_COMBOS_KEY, JSON.stringify(savedCombos || []));
  } catch (e) {
    alert('保存に失敗しました（容量不足の可能性）');
  }
}

// 連番コード生成（C0001, C0002, ...）
function nextComboCode() {
  if (!config) return 'C0001'; // config未ロード時のフォールバック
  
  const prefix = config.ui.titlePrefix;
  const digits = config.ui.titleDigits;
  
  // 連番の起点（localStorageから復元）
  let n = parseInt(localStorage.getItem(LS_CODE_COUNTER_KEY) || '1', 10);
  if (!Number.isFinite(n) || n < 1) n = 1;

  // 既存と被らないコードを探す（削除/復元でも被らないように）
  const maxNum = Math.pow(10, digits) - 1; // 桁数に応じた最大値
  while (true) {
    const code = prefix + String(n).padStart(digits, '0');
    const exists = savedCombos.some(x => x.title === code);
    n++;
    localStorage.setItem(LS_CODE_COUNTER_KEY, String(n));
    if (!exists) return code;

    // 念のため無限ループ回避（最大値を超えたら桁増やす）
    if (n > maxNum) {
      const fallback = prefix + String(Date.now()).slice(-digits);
      if (!savedCombos.some(x => x.title === fallback)) return fallback;
    }
  }
}

// 読み込み（Supabaseから、失敗時はlocalStorageにフォールバック）
async function loadCombosFromLocal() {
  try {
    const { data, error } = await supabaseClient.from('combos').select('*').order('created_at', { ascending: true });
    if (error) throw error;
    savedCombos = data.map(row => ({
      id: row.id,
      title: row.title,
      damage: row.damage ?? '',
      sa1Damage: row.sa1_damage ?? '',
      sa2Damage: row.sa2_damage ?? '',
      sa3Damage: row.sa3_damage ?? '',
      driveGauge: row.drive_gauge ?? 0,
      favorite: row.favorite ?? false,
      combo: row.combo ?? '',
      memo: row.memo ?? '',
      createdAt: row.created_at ?? '',
      updatedAt: row.updated_at ?? '',
    }));
    savedCombos.forEach(combo => {
      combo.driveGauge = calculateDriveGauge(combo.combo);
    });
  } catch (e) {
    console.error('Supabase読み込みエラー、localStorageにフォールバック:', e);
    try {
      const raw = localStorage.getItem(LS_COMBOS_KEY);
      savedCombos = raw ? JSON.parse(raw) : [];
      savedCombos.forEach(combo => {
        combo.driveGauge = calculateDriveGauge(combo.combo);
      });
    } catch (e2) {
      savedCombos = [];
    }
  }
}

// ==========================================================
// Supabase CRUD
// ==========================================================

// 1件保存/更新（upsert）
async function upsertComboToSupabase(combo) {
  const { error } = await supabaseClient.from('combos').upsert({
    id: combo.id,
    title: combo.title,
    damage: combo.damage ?? '',
    sa1_damage: combo.sa1Damage ?? '',
    sa2_damage: combo.sa2Damage ?? '',
    sa3_damage: combo.sa3Damage ?? '',
    drive_gauge: combo.driveGauge ?? 0,
    favorite: combo.favorite ?? false,
    combo: combo.combo ?? '',
    memo: combo.memo ?? '',
    created_at: combo.createdAt,
    updated_at: combo.updatedAt,
  });
  if (error) console.error('Supabase upsertエラー:', error);
}

// 1件削除
async function deleteComboFromSupabase(id) {
  const { error } = await supabaseClient.from('combos').delete().eq('id', id);
  if (error) console.error('Supabase deleteエラー:', error);
}

// ==========================================================
// ブラウザ機能検出
// ==========================================================

// File System Access API 利用可能か
function canUseFSAccess() {
  return 'showOpenFilePicker' in window;
}

// ==========================================================
// ユーティリティ関数
// ==========================================================

// 日時フォーマット(YYYY-MM-DDTHH:MM → YYYY-MM-DD HH:MM)
function fmt(ts) {
  if (!ts) return '';
  return ts.replace('T', ' ');
}

// HTMLエスケープ
function escapeHtml(str) {
  return String(str)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

// コンボのスマートソート用キー生成
function firstMoveKey(combo) {
  const s = String(combo || '').trim();

  // 先頭行だけ見る
  const firstLine = s.split('\n')[0] || '';

  // 先頭が (タグ) なら抽出。なければ空
  const tagMatch = firstLine.match(/^\(([^)]+)\)/);
  const tag = tagMatch ? tagMatch[1] : '';
  const hasTag = tag ? 1 : 0; // 0=通常が先, 1=(...)が後

  // 区切り前の最初の塊
  const head = firstLine.split(/→|>|、|,|\s+/)[0] || firstLine;

  // 例: 2弱P / 5中K などを拾う（全角数字も一応吸収）
  const m = head.match(/([1-9１-９])\s*(弱|中|強)?\s*(P|K)/);

  const dirMap = { '１':1,'２':2,'３':3,'４':4,'５':5,'６':6,'７':7,'８':8,'９':9 };
  const dir = m ? (dirMap[m[1]] || parseInt(m[1], 10)) : 99;

  // 弱→中→強（小さいほど先）
  const strengthOrder = { '弱': 1, '中': 2, '強': 3 };
  const strength = (m && m[2]) ? strengthOrder[m[2]] : 9;

  // P→K（Pが先）
  const buttonOrder = { 'P': 1, 'K': 2 };
  const button = (m && m[3]) ? buttonOrder[m[3]] : 9;

  // 並び順：
  // 通常(0) → タグあり(1)
  // タグは同じものをまとめる
  // その後に方向→弱中強→P/K
  return [hasTag, tag, strength, dir, button, s];
}

// コンボのスマートソート比較関数
function compareComboSmart(aCombo, bCombo) {
  const a = firstMoveKey(aCombo);
  const b = firstMoveKey(bCombo);

  // 比較
  if (a[0] !== b[0]) return a[0] - b[0];
  const tagCmp = String(a[1]).localeCompare(String(b[1]), 'ja');
  if (tagCmp !== 0) return tagCmp;

  for (let i = 2; i <= 4; i++) {
    if (a[i] !== b[i]) return a[i] - b[i];
  }
  return String(a[5]).localeCompare(String(b[5]), 'ja');
}

// 始動技を抽出する（「(○○)」を除いた先頭文字）
// 長い技名から優先的にマッチング（214強P → 214弱P の順）
function extractStartupTech(combo) {
  let c = String(combo || '').trim();
  // 「(○○)」を全て削除
  c = c.replace(/\([^)]*\)/g, '').trim();
  if (!c) return null;

  // 始動技リスト（長い順）
  const startupTechs = ['214中P', '2弱P', '5弱P', '2弱K', '2中P', '5中P', '2中K', '2強P', '5強P', 'DR'];

  for (const tech of startupTechs) {
    if (c.startsWith(tech)) {
      return tech;
    }
  }

  return null;
}

// 一覧描画: `savedCombos` をフィルタ/ソートして DOM に描画する。
// UI のインタラクション（編集/削除/読み込み/お気に入り切替）はここでバインドされる。
function renderComboList() {
  if (!comboList) return;

  // ★描画前に重複除去を実行
  removeDuplicateIds();

  if (!savedCombos || savedCombos.length === 0) {
    comboList.innerHTML = '<div class="empty-message">（保存されたコンボはありません）</div>';
    return;
  }

  // フィルター適用
  let view = [...savedCombos];
  const comboText = (x) => String(x.combo || '').trim();

  // スクリーン位置フィルター（一番目）
  if (screenPositionFilter === 'center') {
    view = view.filter(x => {
      const c = comboText(x);
      return c && !c.includes('(壁限定)');
    });
  } else if (screenPositionFilter === 'corner') {
    view = view.filter(x => {
      const c = comboText(x);
      return c && c.includes('(壁限定)');
    });
  }

  // 始動技フィルター（二番目）
  if (startupTechFilter !== 'all') {
    view = view.filter(x => {
      const tech = extractStartupTech(comboText(x));
      return tech === startupTechFilter;
    });
  }

  // フィルター条件（中央フィルター）
  if (filterKey === 'noSpecial') {
    view = view.filter(x => {
      const c = comboText(x);
      return c && !c.trim().startsWith('(');
    });
  } else if (filterKey === 'fav') {
    view = view.filter(x => !!x.favorite);
  } else if (filterKey === 'counter') {
    view = view.filter(x => comboText(x).includes('(カウンター)'));
  } else if (filterKey === 'punish') {
    view = view.filter(x => comboText(x).includes('(パニカン)'));
  } else if (filterKey === 'impact') {
    view = view.filter(x => comboText(x).includes('(インパクト)'));
  } else if (filterKey === 'okizeme') {
    view = view.filter(x => comboText(x).includes('(起き攻め)'));
  } else if (filterKey === 'noSA') {
    view = view.filter(x => {
      // SA1, SA2, SA3のダメージ入力が全て空の場合のみ表示
      const hasSa1 = x.sa1Damage && String(x.sa1Damage).trim() !== '';
      const hasSa2 = x.sa2Damage && String(x.sa2Damage).trim() !== '';
      const hasSa3 = x.sa3Damage && String(x.sa3Damage).trim() !== '';
      return !hasSa1 && !hasSa2 && !hasSa3;
    });
  } else if (filterKey === 'SA1') {
    view = view.filter(x => x.sa1Damage && String(x.sa1Damage).trim() !== '');
  } else if (filterKey === 'SA2') {
    view = view.filter(x => x.sa2Damage && String(x.sa2Damage).trim() !== '');
  } else if (filterKey === 'SA3') {
    view = view.filter(x => x.sa3Damage && String(x.sa3Damage).trim() !== '');
  }

  // ソート適用
  // ダメージが高い順でソート（SA1/SA2/SA3フィルター時は対応するSAダメージでソート）
  const damageNum = (v) => {
    const n = parseInt(String(v ?? '').replace(/[^\d]/g, ''), 10);
    return Number.isFinite(n) ? n : null; // 未入力はnull
  };

  // SA1/SA2/SA3フィルターの場合は対応するSAダメージでソート、それ以外は通常のDmgでソート
  view.sort((a, b) => {
    let ad, bd;
    
    if (filterKey === 'SA1') {
      ad = damageNum(a.sa1Damage);
      bd = damageNum(b.sa1Damage);
    } else if (filterKey === 'SA2') {
      ad = damageNum(a.sa2Damage);
      bd = damageNum(b.sa2Damage);
    } else if (filterKey === 'SA3') {
      ad = damageNum(a.sa3Damage);
      bd = damageNum(b.sa3Damage);
    } else {
      ad = damageNum(a.damage);
      bd = damageNum(b.damage);
    }

    // 未入力は最後へ
    if (ad == null && bd == null) return 0;
    if (ad == null) return 1;
    if (bd == null) return -1;

    // 数値で降順（ダメージ高い順）
    return bd - ad;
  });

  // 描画
  comboList.innerHTML = '';
  view.forEach(item => {
    const wrapper = document.createElement('div');
    wrapper.className = 'combo-item';
    wrapper.dataset.id = item.id;

    // お気に入りスター
    const star = item.favorite ? '★' : '☆';
    const favClass = item.favorite ? 'fav on' : 'fav';

    // innerHTMLで一気に描画。ドライブゲージ〜ボタン群は row-right で右詰めにする
    const calculatedDriveGauge = calculateDriveGauge(item.combo); // ドライブゲージを自動計算
    wrapper.innerHTML = `
      <div class="meta">
        <button class="${favClass}" title="お気に入り">${star}</button>
         <span class="combo-code">${escapeHtml(item.title || '')}</span>
         <span class="created-updated">作成:${fmt(item.createdAt)} / 更新:${fmt(item.updatedAt)}</span>
      </div>
      <div class="row1">
        <span class="drive-gauge-label">${calculatedDriveGauge} <span>Dg</span></span>
        
        <label class="sa-damage-label">
          <span class="sa1-label${item.sa1Damage ? ' has-value' : ''}">SA1</span>
          <input type="text" class="sa1-damage sa-damage-small"
            value="${escapeHtml(item.sa1Damage || '')}"
            inputmode="numeric" pattern="[0-9]*" maxlength="5" />
        </label>
        <label class="sa-damage-label">
          <span class="sa2-label${item.sa2Damage ? ' has-value' : ''}">SA2</span>
          <input type="text" class="sa2-damage sa-damage-small"
            value="${escapeHtml(item.sa2Damage || '')}"
            inputmode="numeric" pattern="[0-9]*" maxlength="5" />
        </label>
        <label class="sa-damage-label">
          <span class="sa3-label${item.sa3Damage ? ' has-value' : ''}">SA3</span>
          <input type="text" class="sa3-damage sa-damage-small"
            value="${escapeHtml(item.sa3Damage || '')}"
            inputmode="numeric" pattern="[0-9]*" maxlength="5" />
        </label>

        <div class="row-right">
          <label class="damage-label">
            <span>Dmg</span>
            <input type="text" class="damage small"
              value="${escapeHtml(item.damage || '')}"
              inputmode="numeric" pattern="[0-9]*" maxlength="5" />
          </label>

          <button class="load">出力</button>
          <button class="save">更新</button>
          <button class="del">削除</button>
        </div>
      </div>
      <textarea class="combo" rows="3" placeholder="コンボ">${escapeHtml(item.combo || '')}</textarea>
      <textarea class="memo" rows="2" placeholder="メモ">${escapeHtml(item.memo || '')}</textarea>
    `;

    // イベントバインド
    const favBtn = wrapper.querySelector('button.fav');
    if (favBtn) {
      favBtn.addEventListener('click', () => {
        toggleFavorite(item.id);
      });
    }

    // innerHTMLの後で取れる
    const dmgEl = wrapper.querySelector('.damage');
    if (dmgEl) {
      dmgEl.addEventListener('input', () => {
        const n = normalizeDamage(dmgEl.value);
        if (dmgEl.value !== n) dmgEl.value = n;
      });
    }

    // SAダメージ入力欄の正規化
    const sa1DmgEl = wrapper.querySelector('.sa1-damage');
    const sa2DmgEl = wrapper.querySelector('.sa2-damage');
    const sa3DmgEl = wrapper.querySelector('.sa3-damage');
    if (sa1DmgEl) {
      sa1DmgEl.addEventListener('input', () => {
        const n = normalizeDamage(sa1DmgEl.value);
        if (sa1DmgEl.value !== n) sa1DmgEl.value = n;
      });
    }
    if (sa2DmgEl) {
      sa2DmgEl.addEventListener('input', () => {
        const n = normalizeDamage(sa2DmgEl.value);
        if (sa2DmgEl.value !== n) sa2DmgEl.value = n;
      });
    }
    if (sa3DmgEl) {
      sa3DmgEl.addEventListener('input', () => {
        const n = normalizeDamage(sa3DmgEl.value);
        if (sa3DmgEl.value !== n) sa3DmgEl.value = n;
      });
    }

    // 出力へ読み込み
    wrapper.querySelector('.load').addEventListener('click', () => {
      loadComboToEditor(item.id);
    });

    // 更新
    wrapper.querySelector('.save').addEventListener('click', () => {
      const combo = wrapper.querySelector('.combo').value;
      const memo = wrapper.querySelector('.memo').value;
      const damage = normalizeDamage(wrapper.querySelector('.damage').value);
      const sa1Damage = normalizeDamage(wrapper.querySelector('.sa1-damage').value);
      const sa2Damage = normalizeDamage(wrapper.querySelector('.sa2-damage').value);
      const sa3Damage = normalizeDamage(wrapper.querySelector('.sa3-damage').value);
      wrapper.querySelector('.damage').value = damage; // 表示側も正規化
      wrapper.querySelector('.sa1-damage').value = sa1Damage;
      wrapper.querySelector('.sa2-damage').value = sa2Damage;
      wrapper.querySelector('.sa3-damage').value = sa3Damage;
      updateSavedCombo(item.id, item.title, combo, memo, damage, sa1Damage, sa2Damage, sa3Damage);
    });

    // 削除
    wrapper.querySelector('.del').addEventListener('click', () => {
      deleteSavedCombo(item.id);
    });
    
    // リストに追加
    comboList.appendChild(wrapper);
  });

  // 件数表示（フィルター後の表示件数）
  const countEl = document.getElementById('comboCount');
  if (countEl) {
    countEl.textContent = `${view.length}件`;
  }
}

// お気に入り切替
function toggleFavorite(id) {
  const target = savedCombos.find(x => x.id === id);
  if (!target) return;

  target.favorite = !target.favorite;
  target.updatedAt = formatDate();

  saveCombosToLocal();
  upsertComboToSupabase(target);
  renderComboList();
}

// ========= 一覧機能ここまで =========

// ========= ドライブゲージ計算 =========
// コンボ文字列からドライブゲージ消費量を計算
function calculateDriveGauge(comboStr) {
  const s = String(comboStr || '');
  let gauge = 0;

  // DR: 1つあたり1加算
  const drCount = (s.match(/DR/g) || []).length;
  gauge += drCount * 1;

  // CR: 1つあたり3加算
  const crCount = (s.match(/CR/g) || []).length;
  gauge += crCount * 3;

  // PP / KK の処理（矢印の有無で加算値が変わる）
  // マッチ対象: PPまたはKK（PKは増減なし）
  const ppkkRegex = /([→>])?(PP|KK)/g;
  let match;
  while ((match = ppkkRegex.exec(s)) !== null) {
    const hasArrow = match[1] ? true : false;
    // 矢印ありなら1、なければ2
    gauge += hasArrow ? 1 : 2;
  }

  // コンボパーツ（#パーツ名）のドライブゲージを加算
  if (config && config.comboParts) {
    const partRegex = /#([^→\s]+)/g;
    let partMatch;
    while ((partMatch = partRegex.exec(s)) !== null) {
      const partName = partMatch[1];
      const part = config.comboParts.find(p => p.name === partName);
      if (part && part.driveGauge) {
        gauge += part.driveGauge;
      }
    }
  }

  return Math.min(gauge, 7); // 0～7に制限
}

// ========= 一覧操作関数 =========

// 出力欄へ読み込み（既存のoutput/memoを置換。履歴に積んでUndo可能にする）
function loadComboToEditor(id) {
  const target = savedCombos.find(x => x.id === id);
  if (!target) return;

  history.push(output);

  output = target.combo || '';
  if (outputBox) {
    isUpdatingOutputBox = true;
    outputBox.value = output;
    isUpdatingOutputBox = false;
  }
  if (memoBox) memoBox.value = target.memo || '';
  if (damageBox) damageBox.value = normalizeDamage(target.damage || '');
  if (sa1DamageBox) sa1DamageBox.value = normalizeDamage(target.sa1Damage || '');
  if (sa2DamageBox) sa2DamageBox.value = normalizeDamage(target.sa2Damage || '');
  if (sa3DamageBox) sa3DamageBox.value = normalizeDamage(target.sa3Damage || '');
  if (outputCodeEl) outputCodeEl.textContent = target.title || '';
  // 重複ハイライトを即時更新
  updateDuplicateCheck();

  // フォールバック: 一部環境でテキスト比較が微妙にずれる場合があるため
  // ID による直接ハイライトを行う（確実に該当行を強調する）
  try {
    const comboListEl = document.getElementById('comboList');
    if (comboListEl) {
      // 既存ハイライトをクリア
      const allComboItems = comboListEl.querySelectorAll('.combo');
      allComboItems.forEach(textarea => textarea.classList.remove('combo-highlighted'));

      // 対応する行を見つけて強調
      const row = comboListEl.querySelector(`.combo-item[data-id="${id}"]`);
      if (row) {
        const ta = row.querySelector('.combo');
        if (ta) ta.classList.add('combo-highlighted');
      }
    }
  } catch (e) {
    // フォールバック失敗しても処理継続
  }
}

// 一覧から更新（タイトル/コンボ/メモ/ダメージ/SAダメージ/ドライブゲージ）
function updateSavedCombo(id, newTitle, newCombo, newMemo, newDamage, newSa1Damage, newSa2Damage, newSa3Damage) {
  const target = savedCombos.find(x => x.id === id);
  if (!target) return;

  if (!newTitle) { alert('タイトルは空にできません'); return; }

  const dup = savedCombos.find(x => x.id !== id && x.title === newTitle);
  if (dup) { alert('同じタイトルが既にあります。別のタイトルにしてください。'); return; }

  target.title = newTitle;
  target.combo = newCombo;
  target.memo = newMemo;
  target.damage = normalizeDamage(newDamage);
  target.sa1Damage = normalizeDamage(newSa1Damage || '');
  target.sa2Damage = normalizeDamage(newSa2Damage || '');
  target.sa3Damage = normalizeDamage(newSa3Damage || '');
  target.driveGauge = calculateDriveGauge(newCombo); // ★ドライブゲージを自動計算
  target.updatedAt = formatDate();

  // 重複除去: 同じIDが複数ある場合、最新のupdatedAtを持つもの以外を削除
  removeDuplicateIds();

  alert('更新しました');
  renderComboList();
  saveCombosToLocal();
  upsertComboToSupabase(target);
}

// ID重複を除去する（最新のupdatedAtを持つものを残す）
function removeDuplicateIds() {
  const idMap = new Map();
  
  // 同じIDのデータを収集
  savedCombos.forEach(item => {
    if (!idMap.has(item.id)) {
      idMap.set(item.id, []);
    }
    idMap.get(item.id).push(item);
  });
  
  // 各IDグループで最新のものだけを残す
  const uniqueCombos = [];
  idMap.forEach((items, id) => {
    if (items.length === 1) {
      uniqueCombos.push(items[0]);
    } else {
      // 複数ある場合は最新のupdatedAtを持つものを残す
      console.warn(`[removeDuplicateIds] ID ${id} に重複があります (${items.length}件)。最新のものを残します。`);
      const latest = items.reduce((prev, current) => {
        const prevTime = prev.updatedAt || prev.createdAt || '';
        const currentTime = current.updatedAt || current.createdAt || '';
        return currentTime > prevTime ? current : prev;
      });
      uniqueCombos.push(latest);
    }
  });
  
  savedCombos = uniqueCombos;
}

// 一覧から削除
function deleteSavedCombo(id) {
  const target = savedCombos.find(x => x.id === id);
  if (!target) return;

  if (!confirm(`「${target.title}」を削除しますか？`)) return;
  savedCombos = savedCombos.filter(x => x.id !== id);
  renderComboList();
  saveCombosToLocal();
  deleteComboFromSupabase(id);
}

// ダメージ正規化（数字以外除去）
function normalizeDamage(v) {
  const s = String(v ?? '').replace(/[^\d]/g, ''); // 数字以外除去
  return s;
}

// ダメージ入力欄の変化を監視して正規化＆保存
if (damageBox) {
  damageBox.addEventListener('input', () => {
    const n = normalizeDamage(damageBox.value);
    if (damageBox.value !== n) damageBox.value = n;
    saveState();
  });
}

// SAダメージ入力欄の正規化＆保存
if (sa1DamageBox) {
  sa1DamageBox.addEventListener('input', () => {
    const n = normalizeDamage(sa1DamageBox.value);
    if (sa1DamageBox.value !== n) sa1DamageBox.value = n;
    saveState();
  });
}
if (sa2DamageBox) {
  sa2DamageBox.addEventListener('input', () => {
    const n = normalizeDamage(sa2DamageBox.value);
    if (sa2DamageBox.value !== n) sa2DamageBox.value = n;
    saveState();
  });
}
if (sa3DamageBox) {
  sa3DamageBox.addEventListener('input', () => {
    const n = normalizeDamage(sa3DamageBox.value);
    if (sa3DamageBox.value !== n) sa3DamageBox.value = n;
    saveState();
  });
}

// ---------- 各種ボタン初期化 ----------

// 方向キーグリッド初期化
function initDirectionGrid() {
  if (!directionGrid) return;
  for (let i = 7; i <= 9; i++) createDirectionButton(i);
  for (let i = 4; i <= 6; i++) createDirectionButton(i);
  for (let i = 1; i <= 3; i++) createDirectionButton(i);
}

// 方向ボタン作成
function createDirectionButton(value) {
  if (!directionGrid) return;
  const button = document.createElement('button');
  button.className = 'direction-button';
  button.setAttribute('data-value', value);
  button.textContent = ARROW_MAP[value];
  button.onclick = () => addInput(value.toString());
  directionGrid.appendChild(button);
}

// 特殊方向入力初期化
function initSpecialDirections() {
  if (!specialDirections || !config) return;

  const commands = config.buttons.specialDirections;
  commands.forEach(cmd => {
    const button = document.createElement('button');
    button.textContent = cmd.text;
    button.onclick = () => addInput(cmd.value);
    specialDirections.appendChild(button);
  });
}

// 追加特殊方向入力初期化
function initSpecialDirectionsExtra() {
  const box = document.getElementById('specialDirectionsExtra');
  if (!box || !config) return;

  const inputs = config.buttons.specialDirectionsExtra;
  inputs.forEach(inp => {
    const button = document.createElement('button');
    button.textContent = inp.text;
    if (inp.longText) button.className = 'long-text-button';
    button.onclick = () => addInput(inp.value);
    box.appendChild(button);
  });
}

// 追加特殊コマンド入力初期化（クラシック・モダン共通）
function initSpecialCommandExtra() {
  if (!config) return;
  const inputs = config.buttons.specialCommandExtra;

  // クラシック用
  const boxClassic = document.getElementById('specialCommandExtra');
  if (boxClassic) {
    inputs.forEach(inp => {
      const button = document.createElement('button');
      button.textContent = inp.text;
      button.className = inp.class;
      button.onclick = () => addInput(inp.value);
      boxClassic.appendChild(button);
    });
  }

  // モダン用
  const boxModern = document.getElementById('specialCommandExtraModern');
  if (boxModern) {
    inputs.forEach(inp => {
      const button = document.createElement('button');
      button.textContent = inp.text;
      button.className = inp.class;
      button.onclick = () => addInput(inp.value);
      boxModern.appendChild(button);
    });
  }
}

// 特殊コマンドボタン初期化
function initSpecialCommandButtons() {
  if (!config) return;
  const commands = config.buttons.specialCommands;

  // 特殊コマンドボタン作成
  const container = document.getElementById('specialCommandButtons');
  if (container) {
    commands.forEach(cmd => {
      const button = document.createElement('button');
      button.textContent = cmd.text;
      button.onclick = () => addInput(cmd.value);
      container.appendChild(button);
    });
    const placeholderButton = document.createElement('button');
    placeholderButton.style.visibility = 'hidden';
    container.appendChild(placeholderButton);
  }

  // モダン用特殊コマンドボタン作成
  const containerModern = document.getElementById('specialCommandButtonsModern');
  if (containerModern) {
    commands.forEach(cmd => {
      const button = document.createElement('button');
      button.textContent = cmd.text;
      button.onclick = () => addInput(cmd.value);
      containerModern.appendChild(button);
    });
    const placeholderButtonModern = document.createElement('button');
    placeholderButtonModern.style.visibility = 'hidden';
    containerModern.appendChild(placeholderButtonModern);
  }
}

// 攻撃ボタン初期化
function initAttackButtons() {
  if (!classicButtons || !modernButtons || !saButtons || !saButtonsModern || !config) return;

  // クラシック攻撃ボタン作成
  const classicAttacks = config.buttons.classicAttacks;
  classicAttacks.forEach(atk => {
    const button = document.createElement('button');
    button.textContent = atk.text;
    button.className = `btn-attack ${atk.class}`;
    button.onclick = () => addInput(atk.value, true);
    classicButtons.appendChild(button);
  });

  // モダン攻撃ボタン作成
  const modernAttacks = config.buttons.modernAttacks;
  modernAttacks.forEach(atk => {
    const button = document.createElement('button');
    button.textContent = atk.text;
    if (atk.class) button.className = `btn-attack ${atk.class}`;
    button.onclick = () => addInput(atk.value, true);
    modernButtons.appendChild(button);
  });

  // スーパーアーツボタン作成
  const createSAButtons = (container) => {
    if (!container) return;
    for (let i = 1; i <= 3; i++) {
      const button = document.createElement('button');
      button.textContent = `SA${i}`;
      button.onclick = () => addInput(`SA${i}`);
      container.appendChild(button);
    }
  };

  createSAButtons(saButtons);
  createSAButtons(saButtonsModern);
}

// 特殊入力ボタン初期化
function initSpecialInputs() {
  if (!specialInputs || !config) return;

  const inputs = config.buttons.specialInputs;
  inputs.forEach(inp => {
    const button = document.createElement('button');
    button.textContent = inp.text;
    if (inp.longText) button.className = 'long-text-button';
    button.onclick = () => addInput(inp.value);
    specialInputs.appendChild(button);
  });
}

// フィルター初期化
function initFilters() {
  if (!config) return;

  // 位置フィルター
  const screenPositionFilterSel = document.getElementById('screenPositionFilter');
  if (screenPositionFilterSel && config.filters.screenPosition) {
    config.filters.screenPosition.forEach(opt => {
      const option = document.createElement('option');
      option.value = opt.value;
      option.textContent = opt.text;
      screenPositionFilterSel.appendChild(option);
    });
  }

  // 始動技フィルター
  const startupTechFilterSel = document.getElementById('startupTechFilter');
  if (startupTechFilterSel && config.filters.startupTech) {
    config.filters.startupTech.forEach(opt => {
      const option = document.createElement('option');
      option.value = opt.value;
      option.textContent = opt.text;
      startupTechFilterSel.appendChild(option);
    });
  }

  // 条件フィルター
  const filterKeySel = document.getElementById('filterKey');
  if (filterKeySel && config.filters.condition) {
    config.filters.condition.forEach(opt => {
      const option = document.createElement('option');
      option.value = opt.value;
      option.textContent = opt.text;
      filterKeySel.appendChild(option);
    });
  }
}

// ==========================================================
// コンボパーツ管理
// ==========================================================

// コンボパーツ一覧を描画
function renderComboParts() {
  const container = document.getElementById('comboPartsList');
  if (!container || !config) return;

  container.innerHTML = '';
  
  if (!config.comboParts || config.comboParts.length === 0) {
    container.innerHTML = '<div class="empty-message">登録されているパーツはありません</div>';
    return;
  }

  config.comboParts.forEach((part, index) => {
    const row = document.createElement('div');
    row.className = 'part-item';
    
    const nameSpan = document.createElement('span');
    nameSpan.className = 'part-name';
    nameSpan.textContent = part.name;
    
    const compSpan = document.createElement('span');
    compSpan.className = 'part-composition';
    compSpan.textContent = part.composition;
    
    const gaugeSpan = document.createElement('span');
    gaugeSpan.className = 'part-drive-gauge';
    gaugeSpan.textContent = `Dg:${part.driveGauge || 0}`;
    
    const actions = document.createElement('div');
    actions.className = 'part-actions';
    
    const outputBtn = document.createElement('button');
    outputBtn.textContent = '出力';
    outputBtn.className = 'tiny-button';
    outputBtn.onclick = () => addPartToOutput(part);
    
    const deleteBtn = document.createElement('button');
    deleteBtn.textContent = '削除';
    deleteBtn.className = 'tiny-button';
    deleteBtn.onclick = () => deleteComboPart(index);
    
    actions.appendChild(outputBtn);
    actions.appendChild(deleteBtn);
    
    row.appendChild(nameSpan);
    row.appendChild(compSpan);
    row.appendChild(gaugeSpan);
    row.appendChild(actions);
    
    container.appendChild(row);
  });
}

// コンボパーツを保存
function saveComboPart() {
  const nameInput = document.getElementById('partName');
  const compInput = document.getElementById('partComposition');
  const gaugeInput = document.getElementById('partDriveGauge');
  
  if (!nameInput || !compInput || !config) return;
  
  const name = nameInput.value.trim();
  const composition = compInput.value.trim();
  const driveGauge = gaugeInput ? parseInt(gaugeInput.value) || 0 : 0;
  
  if (!name || !composition) {
    alert('パーツ名と構成を入力してください');
    return;
  }
  
  // 同名パーツがあれば更新、なければ追加
  const existingIndex = config.comboParts.findIndex(p => p.name === name);
  
  if (existingIndex >= 0) {
    config.comboParts[existingIndex].composition = composition;
    config.comboParts[existingIndex].driveGauge = driveGauge;
  } else {
    config.comboParts.push({ name, composition, driveGauge });
  }
  
  // config.jsonを更新（fetch APIでは書き込めないので、将来的にはlocalStorageに移行推奨）
  saveConfigToLocalStorage();
  
  // 入力欄をクリア
  nameInput.value = '';
  compInput.value = '';
  if (gaugeInput) gaugeInput.value = '';
  
  renderComboParts();
}

// コンボパーツを削除
function deleteComboPart(index) {
  if (!config || !config.comboParts) return;
  
  const part = config.comboParts[index];
  if (confirm(`「${part.name}」を削除しますか？`)) {
    config.comboParts.splice(index, 1);
    saveConfigToLocalStorage();
    renderComboParts();
  }
}

// パーツを出力欄に追加
function addPartToOutput(part) {
  // #パーツ名を出力欄に追加
  history.push(output);
  output += `#${part.name}`;
  
  if (outputBox) {
    isUpdatingOutputBox = true;
    outputBox.value = output;
    isUpdatingOutputBox = false;
  }
  
  // メモ欄の先頭に #構成 を追加
  if (memoBox) {
    const currentMemo = memoBox.value;
    const newMemo = `#${part.composition}\n${currentMemo}`;
    memoBox.value = newMemo;
  }
  
  saveState();
  updateDuplicateCheck();
  
  // 検索窓をクリア
  const comboSearchInput = document.getElementById('comboSearch');
  if (comboSearchInput) {
    comboSearchInput.value = '';
    updateSearchResult();
  }
}

// configをlocalStorageに保存（comboParts用）
function saveConfigToLocalStorage() {
  try {
    localStorage.setItem('sf6_comboParts', JSON.stringify(config.comboParts || []));
  } catch (e) {
    console.error('コンボパーツの保存に失敗しました', e);
  }
}

// configをlocalStorageから読み込み（comboParts用）
function loadConfigFromLocalStorage() {
  try {
    const saved = localStorage.getItem('sf6_comboParts');
    if (saved && config) {
      config.comboParts = JSON.parse(saved);
    }
  } catch (e) {
    console.error('コンボパーツの読み込みに失敗しました', e);
  }
}

// コンボパーツをJSONファイルとしてエクスポート
function exportComboParts() {
  if (!config || !config.comboParts) return;
  
  const json = JSON.stringify(config.comboParts, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'comboParts.json';
  a.click();
  URL.revokeObjectURL(url);
}

// コンボパーツをJSONファイルからインポート
async function importComboParts() {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.json';
  
  input.onchange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    try {
      const text = await file.text();
      const parts = JSON.parse(text);
      
      // 配列であることを確認
      if (!Array.isArray(parts)) {
        alert('無効なファイル形式です');
        return;
      }
      
      // 各要素がnameとcompositionを持つことを確認
      const valid = parts.every(p => p.name && p.composition);
      if (!valid) {
        alert('無効なパーツデータです');
        return;
      }
      
      // 確認ダイアログ
      if (confirm(`${parts.length}個のパーツを読み込みます。既存のパーツは上書きされますがよろしいですか？`)) {
        config.comboParts = parts;
        saveConfigToLocalStorage();
        renderComboParts();
        alert('パーツを読み込みました');
      }
    } catch (e) {
      alert('ファイルの読み込みに失敗しました: ' + e.message);
    }
  };
  
  input.click();
}

// モード取得（arrow / num）
function getMode() {
  const el = document.querySelector('input[name="mode"]:checked');
  return el ? el.value : 'arrow';
}

// 方向キー変換（数字文字列を矢印に変換）
function convertToArrow(seq) {
  return seq.split('').map(c => ARROW_MAP[c] || c).join('');
}

// ==========================================================
// 入力処理
// ==========================================================

// 入力追加（ボタンまたはキーボードから）
function addInput(value, isAttack = false) {
  let mode = getMode();
  let insert = value;

  if (!isNaN(value)) {
    insert = (mode === 'arrow') ? convertToArrow(value) : value;
    lastWasDirection = true;
  } else if (isAttack) {
    if (lastWasDirection) {
      insert = (mode === 'arrow') ? '+' + value : value;
    } else {
      insert = value;
    }
    insert = `${insert}`;
    lastWasDirection = false;
  } else {
    lastWasDirection = false;
  }

  output += insert;
  if (outputBox) {
    isUpdatingOutputBox = true;
    outputBox.value = output;
    isUpdatingOutputBox = false;
  }
  saveState();
  
  // 重複検查を実行
  updateDuplicateCheck();
  
  // 検索窓をクリア
  const comboSearchInput = document.getElementById('comboSearch');
  if (comboSearchInput) {
    comboSearchInput.value = '';
    updateSearchResult();
  }
}

// コピー出力欄のみ
function copyOutput() {
  navigator.clipboard.writeText(output);
}

// クリア出力欄＋メモ欄＋ダメージ欄
function clearOutput() {
  output = '';
  if (outputBox) {
    isUpdatingOutputBox = true;
    outputBox.value = '';
    isUpdatingOutputBox = false;
  }
  if (memoBox) memoBox.value = '';
  if (damageBox) damageBox.value = '';
  if (sa1DamageBox) sa1DamageBox.value = '';
  if (sa2DamageBox) sa2DamageBox.value = '';
  if (sa3DamageBox) sa3DamageBox.value = '';
  if (outputCodeEl) outputCodeEl.textContent = '';  // ★コード表示もクリア

  lastWasDirection = false;
  history = [];

  localStorage.removeItem('sf6_output');
  localStorage.removeItem('sf6_memo');
  localStorage.removeItem(LS_DAMAGE_KEY);
  localStorage.removeItem(LS_SA1_DAMAGE_KEY);
  localStorage.removeItem(LS_SA2_DAMAGE_KEY);
  localStorage.removeItem(LS_SA3_DAMAGE_KEY);
}

// 状態保存
function saveState() {
  localStorage.setItem('sf6_output', output);
  localStorage.setItem('sf6_memo', memoBox ? memoBox.value : '');
  try {
    const d = damageBox ? normalizeDamage(damageBox.value) : '';
    localStorage.setItem(LS_DAMAGE_KEY, d);
    const sa1d = sa1DamageBox ? normalizeDamage(sa1DamageBox.value) : '';
    localStorage.setItem(LS_SA1_DAMAGE_KEY, sa1d);
    const sa2d = sa2DamageBox ? normalizeDamage(sa2DamageBox.value) : '';
    localStorage.setItem(LS_SA2_DAMAGE_KEY, sa2d);
    const sa3d = sa3DamageBox ? normalizeDamage(sa3DamageBox.value) : '';
    localStorage.setItem(LS_SA3_DAMAGE_KEY, sa3d);
  } catch (e) {
    // ignore
  }
}

// 出力モード変換
function convertOutputMode() {
  let mode = getMode();
  let newOutput = output;

  if (mode === 'arrow') {
    newOutput = newOutput.replace(/[123456789]/g, m => ARROW_MAP[m]);
    newOutput = newOutput.replace(/([↖↗↙↘←→↑↓N])([A-Za-z一-龥])/g, '$1+$2');
  } else {
    const reverseMap = Object.fromEntries(Object.entries(ARROW_MAP).map(([k, v]) => [v, k]));
    newOutput = newOutput.replace(/[↖↗↙↘←→↑↓N]/g, m => reverseMap[m] || m);
    newOutput = newOutput.replace(/([1-9])\+([A-Za-z一-龥])/g, '$1$2');
  }

  output = newOutput;
  if (outputBox) {
    isUpdatingOutputBox = true;
    outputBox.value = output;
    isUpdatingOutputBox = false;
  }
  saveState();
}

// 方向ボタン表示更新
function updateDirectionButtons() {
  const mode = getMode();
  const buttons = document.querySelectorAll('.direction-button');

  buttons.forEach(button => {
    const value = button.getAttribute('data-value');
    button.textContent = mode === 'arrow' ? (ARROW_MAP[value] || value) : value;
  });
}

// 日付フォーマット (YYYY-MM-DDTHH:MM)
function formatDate() {
  let now = new Date();
  let year = now.getFullYear();
  let month = String(now.getMonth() + 1).padStart(2, '0');
  let day = String(now.getDate()).padStart(2, '0');
  let hour = String(now.getHours()).padStart(2, '0');
  let minute = String(now.getMinutes()).padStart(2, '0');
  return `${year}-${month}-${day}T${hour}:${minute}`;
}

// 一意ID生成（YYYYMMDDHHMMを16進数化）
function generateID() {
  let now = new Date();
  let dateStr = now.getFullYear() +
    String(now.getMonth() + 1).padStart(2, '0') +
    String(now.getDate()).padStart(2, '0') +
    String(now.getHours()).padStart(2, '0') +
    String(now.getMinutes()).padStart(2, '0');
  return parseInt(dateStr, 10).toString(16);
}

// 新規コード生成（COMBO-001 形式）
function saveCurrentCombo() {
  const combo = output;
  const memoText = memoBox ? memoBox.value : '';
  const damage = damageBox ? normalizeDamage(damageBox.value) : '';
  const sa1Damage = sa1DamageBox ? normalizeDamage(sa1DamageBox.value) : '';
  const sa2Damage = sa2DamageBox ? normalizeDamage(sa2DamageBox.value) : '';
  const sa3Damage = sa3DamageBox ? normalizeDamage(sa3DamageBox.value) : '';

  if (combo.trim() === "") {
    alert("コンボが入力されていません。");
    return;
  }

  // 重複検查
  const hasDuplicate = savedCombos.some(item => item.combo === combo);
  if (hasDuplicate) {
    alert('このコンボはすでに登録されています。');
    return;
  }

  const title = nextComboCode();     // ★新規コード
  if (outputCodeEl) outputCodeEl.textContent = title; // ★出力側にも反映

  const nowFormatted = formatDate();
  const driveGauge = calculateDriveGauge(combo); // ★ドライブゲージを計算

  const newCombo = {
    id: generateID(),
    title,
    combo,
    memo: memoText,
    damage,
    sa1Damage,
    sa2Damage,
    sa3Damage,
    driveGauge, // ★ドライブゲージを追加
    favorite: false,
    createdAt: nowFormatted,
    updatedAt: nowFormatted
  };
  savedCombos.push(newCombo);

  renderComboList();
  saveCombosToLocal();
  upsertComboToSupabase(newCombo);

  // 保存後にエディタ内容＆コード表示をクリア
  clearOutput();
}

// コンボ一覧をJSON出力
function exportCombosToJSON() {
  if (!savedCombos || savedCombos.length === 0) {
    alert("保存されたコンボがありません。");
    return;
  }

  const jsonContent = JSON.stringify(savedCombos, null, 2);
  const blob = new Blob([jsonContent], { type: "application/json;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "combos.json";
  a.click();
  URL.revokeObjectURL(url);
}

// ==========================================================
// グローバルイベント
// ==========================================================

// キーボードショートカット
document.addEventListener("keydown", function (event) {
  if (document.activeElement === memoBox) return;

  if (event.key === "ArrowRight") {
    event.preventDefault();
    addInput('→');
  }
});

// モード変更イベント
document.querySelectorAll('input[name="mode"]').forEach(radio => {
  radio.addEventListener('change', () => {
    convertOutputMode();
    updateDirectionButtons();
  });
});

// スタイル変更イベント
document.querySelectorAll('input[name="style"]').forEach(radio => {
  radio.addEventListener('change', () => {
    const el = document.querySelector('input[name="style"]:checked');
    const style = el ? el.value : 'classic';
    const classic = document.getElementById('classic-buttons');
    const modern = document.getElementById('modern-buttons');
    if (classic) classic.style.display = (style === 'classic') ? 'block' : 'none';
    if (modern) modern.style.display = (style === 'modern') ? 'block' : 'none';
  });
});

// 折りたたみセクション初期化
function initCollapsibles() {
  const collapsibles = document.querySelectorAll('.collapsible');

  collapsibles.forEach(item => {
    item.addEventListener('click', function () {
      this.classList.toggle('collapsed');
      const id = this.id;
      if (id) localStorage.setItem(`sf6_collapsed_${id}`, this.classList.contains('collapsed'));
    });

    const id = item.id;
    if (id && localStorage.getItem(`sf6_collapsed_${id}`) === 'true') {
      item.classList.add('collapsed');
    }
  });
}

// 重複検出発出を更新
 function updateDuplicateCheck() {
  const duplicateWarning = document.getElementById('duplicateWarning');
  const comboListEl = document.getElementById('comboList');

  // 出力が空ならハイライト解除して終了
  if (!duplicateWarning || !String(output || '').trim()) {
    if (duplicateWarning) {
      duplicateWarning.textContent = '';
      duplicateWarning.className = 'duplicate-label';
    }
    if (comboListEl) {
      const allComboItems = comboListEl.querySelectorAll('.combo');
      allComboItems.forEach(textarea => textarea.classList.remove('combo-highlighted'));
    }
    return;
  }

  // 比較時に改行や末尾空白の差を吸収する正規化
  const normalizeCombo = (s) => String(s ?? '').replace(/\r\n/g, '\n').trim();
  const normOut = normalizeCombo(output);

  const hasDuplicate = savedCombos.some(combo => normalizeCombo(combo.combo) === normOut);

  // ハイライト処理（textareaの値も正規化して比較）
  if (comboListEl) {
    const allComboItems = comboListEl.querySelectorAll('.combo');
    allComboItems.forEach(textarea => {
      if (normalizeCombo(textarea.value) === normOut) textarea.classList.add('combo-highlighted');
      else textarea.classList.remove('combo-highlighted');
    });
  }

  if (hasDuplicate) {
    duplicateWarning.textContent = '重複あり';
    duplicateWarning.className = 'duplicate-label has-duplicate';
  } else {
    duplicateWarning.textContent = '';
    duplicateWarning.className = 'duplicate-label';
  }
}

 function updateSearchResult() {
  const searchInput = document.getElementById('comboSearch');
  const searchResultLabel = document.getElementById('searchResult');
  if (!searchInput || !searchResultLabel) return;

  const searchText = searchInput.value.trim();
  if (!searchText) {
    searchResultLabel.textContent = '';
    searchResultLabel.className = 'search-result-label';
    // 検索が空の時を全ハイライトを消去
    const comboList = document.getElementById('comboList');
    if (comboList) {
      const allComboItems = comboList.querySelectorAll('.combo');
      allComboItems.forEach(textarea => {
        textarea.classList.remove('combo-highlighted');
      });
    }
    return;
  }

  const found = savedCombos.some(combo => combo.combo === searchText);
  
  // ハイライト処理: コンボ一覧を反映
  const comboList = document.getElementById('comboList');
  if (comboList) {
    const allComboItems = comboList.querySelectorAll('.combo');
    allComboItems.forEach(textarea => {
      if (textarea.value === searchText) {
        textarea.classList.add('combo-highlighted');
      } else {
        textarea.classList.remove('combo-highlighted');
      }
    });
  }
  
  if (found) {
    searchResultLabel.textContent = '登録あり';
    searchResultLabel.className = 'search-result-label found';
  } else {
    searchResultLabel.textContent = '登録なし';
    searchResultLabel.className = 'search-result-label not-found';
  }
}

// 初期化: UI生成、イベントバインド、データ復元
async function init() {
  // config.jsonを読み込み
  await loadConfig();
  
  // 設定値を初期化
  LS_COMBOS_KEY = config.localStorage.combosKey;
  LS_CODE_COUNTER_KEY = config.localStorage.codeCounterKey;
  LS_DAMAGE_KEY = config.localStorage.damageKey;
  LS_SA1_DAMAGE_KEY = config.localStorage.sa1DamageKey;
  LS_SA2_DAMAGE_KEY = config.localStorage.sa2DamageKey;
  LS_SA3_DAMAGE_KEY = config.localStorage.sa3DamageKey;
  LS_SORT_KEY = config.localStorage.sortKey;
  LS_FILTER_KEY = config.localStorage.filterKey;
  LS_SCREEN_POSITION_FILTER_KEY = config.localStorage.screenPositionFilterKey;
  LS_STARTUP_TECH_FILTER_KEY = config.localStorage.startupTechFilterKey;
  ARROW_MAP = config.arrowMap;
  
  // アプリケーション状態の初期化
  sortKey = localStorage.getItem(LS_SORT_KEY) || 'updatedDesc';
  filterKey = localStorage.getItem(LS_FILTER_KEY) || 'all';
  screenPositionFilter = localStorage.getItem(LS_SCREEN_POSITION_FILTER_KEY) || 'all';
  startupTechFilter = localStorage.getItem(LS_STARTUP_TECH_FILTER_KEY) || 'all';

  // UI初期化
  initDirectionGrid();
  initSpecialDirections();
  initSpecialDirectionsExtra(); 
  initAttackButtons();
  initSpecialInputs();
  initSpecialCommandButtons();
  initSpecialCommandExtra();
  initFilters();
  initCollapsibles();
  updateDirectionButtons();

  // コンボパーツ初期化（config.jsonにcomboPartsがない場合は空配列）
  if (!config.comboParts) {
    config.comboParts = [];
  }
  loadConfigFromLocalStorage();
  renderComboParts();

  await loadCombosFromLocal();
  renderComboList();

  const screenPositionFilterSel = document.getElementById('screenPositionFilter');
  if (screenPositionFilterSel) {
    screenPositionFilterSel.value = screenPositionFilter;
    screenPositionFilterSel.addEventListener('change', () => {
      screenPositionFilter = screenPositionFilterSel.value;
      localStorage.setItem(LS_SCREEN_POSITION_FILTER_KEY, screenPositionFilter);
      renderComboList();
    });
  }

  const startupTechFilterSel = document.getElementById('startupTechFilter');
  if (startupTechFilterSel) {
    startupTechFilterSel.value = startupTechFilter;
    startupTechFilterSel.addEventListener('change', () => {
      startupTechFilter = startupTechFilterSel.value;
      localStorage.setItem(LS_STARTUP_TECH_FILTER_KEY, startupTechFilter);
      renderComboList();
    });
  }

  const filterSel = document.getElementById('filterKey');
  if (filterSel) {
    filterSel.value = filterKey;
    filterSel.addEventListener('change', () => {
      filterKey = filterSel.value;
      localStorage.setItem(LS_FILTER_KEY, filterKey);
      renderComboList();
    });
  }

  const savedOutput = localStorage.getItem('sf6_output');
  const savedMemo = localStorage.getItem('sf6_memo');

  const savedDamage = localStorage.getItem(LS_DAMAGE_KEY);
  if (savedDamage && damageBox) {
    damageBox.value = normalizeDamage(savedDamage);
  }

  const savedSa1Damage = localStorage.getItem(LS_SA1_DAMAGE_KEY);
  if (savedSa1Damage && sa1DamageBox) {
    sa1DamageBox.value = normalizeDamage(savedSa1Damage);
  }

  const savedSa2Damage = localStorage.getItem(LS_SA2_DAMAGE_KEY);
  if (savedSa2Damage && sa2DamageBox) {
    sa2DamageBox.value = normalizeDamage(savedSa2Damage);
  }

  const savedSa3Damage = localStorage.getItem(LS_SA3_DAMAGE_KEY);
  if (savedSa3Damage && sa3DamageBox) {
    sa3DamageBox.value = normalizeDamage(savedSa3Damage);
  }

  if (savedOutput) {
    output = savedOutput;
    if (outputBox) {
      isUpdatingOutputBox = true;
      outputBox.value = output;
      isUpdatingOutputBox = false;
    }
  }
  if (outputBox) {
    outputBox.addEventListener('input', () => {
      // プログラム的な更新の場合は処理をスキップ
      if (isUpdatingOutputBox) return;
      
      output = outputBox.value;
      saveState();
      // 重複検査を実行
      updateDuplicateCheck();
      // 出力が入力されたら検索窓をクリア
      const comboSearchInput = document.getElementById('comboSearch');
      if (comboSearchInput) {
        comboSearchInput.value = '';
        updateSearchResult();
      }
    });
  }
  if (savedMemo && memoBox) {
    memoBox.value = savedMemo;
  }

  const comboSearchInput = document.getElementById('comboSearch');
  if (comboSearchInput) {
    comboSearchInput.value = '';
    comboSearchInput.addEventListener('input', updateSearchResult);
  }

  const clearSearchBtn = document.getElementById('clearSearchBtn');
  if (clearSearchBtn) {
    clearSearchBtn.addEventListener('click', () => {
      if (comboSearchInput) {
        comboSearchInput.value = '';
        updateSearchResult();
      }
    });
  }
}

// ==========================================================
// DOM初期化後のイベントリスナー設定
// ==========================================================

window.addEventListener('DOMContentLoaded', () => {
  init();
  
  // コントロールボタンのイベントバインド
  const copyBtn = document.getElementById('copyBtn');
  const clearBtn = document.getElementById('clearBtn');
  const saveComboBtn = document.getElementById('saveComboBtn');
  
  if (copyBtn) copyBtn.addEventListener('click', copyOutput);
  if (clearBtn) clearBtn.addEventListener('click', clearOutput);
  if (saveComboBtn) saveComboBtn.addEventListener('click', saveCurrentCombo);
  
  // ヘルプモーダル関連
  const overlay = document.getElementById('overlay');
  const infoModal = document.getElementById('infoModal');
  const closeModalBtn = document.getElementById('closeModal');
  const helpBtn = document.getElementById('helpBtn');

  function showInfoModal() {
    // モーダル内容をconfigから動的に生成
    const modalTitle = document.getElementById('modalTitle');
    const modalBody = document.getElementById('modalBody');
    
    if (modalTitle && config.help) {
      modalTitle.textContent = config.help.title;
    }
    
    if (modalBody && config.help && config.help.items) {
      modalBody.innerHTML = '';
      config.help.items.forEach(item => {
        const p = document.createElement('p');
        const label = document.createElement('b');
        label.textContent = item.label;
        p.appendChild(label);
        p.appendChild(document.createTextNode(': ' + item.description));
        modalBody.appendChild(p);
      });
    }
    
    if (overlay) overlay.style.display = 'block';
    if (infoModal) infoModal.style.display = 'block';
  }

  function hideInfoModal() {
    if (overlay) overlay.style.display = 'none';
    if (infoModal) infoModal.style.display = 'none';
  }

  if (closeModalBtn) closeModalBtn.addEventListener('click', hideInfoModal);
  if (overlay) overlay.addEventListener('click', hideInfoModal);
  if (helpBtn) helpBtn.addEventListener('click', showInfoModal);

  // コンボパーツ関連
  const savePartBtn = document.getElementById('savePartBtn');
  if (savePartBtn) {
    savePartBtn.addEventListener('click', saveComboPart);
  }
});

// JSON読み込みボタン
const openJsonBtn = document.getElementById('openJsonBtn');
if (openJsonBtn) {
  openJsonBtn.addEventListener('click', async () => {
    try {
      await importJson();
    } catch (e) {
      console.error('JSONインポートエラー:', e);
    }
  });
}

// JSON書き出しボタン
const exportJsonBtn = document.getElementById('exportJsonBtn');
if (exportJsonBtn) {
  exportJsonBtn.addEventListener('click', () => {
    exportCombosToJSON();
  });
}

// コンボパーツエクスポートボタン
const exportPartsBtn = document.getElementById('exportPartsBtn');
if (exportPartsBtn) {
  exportPartsBtn.addEventListener('click', exportComboParts);
}

// コンボパーツインポートボタン
const importPartsBtn = document.getElementById('importPartsBtn');
if (importPartsBtn) {
  importPartsBtn.addEventListener('click', importComboParts);
}

// JSON形式でインポート
async function importJson() {
  if (!canUseFSAccess()) {
    alert('このブラウザはファイル直接読み込みに未対応です。');
    return;
  }

  try {
    const [handle] = await window.showOpenFilePicker({
      types: [{ description: 'JSON', accept: { 'application/json': ['.json'] } }],
      multiple: false
    });

    const file = await handle.getFile();
    const text = await file.text();
    const data = JSON.parse(text);
    
    if (Array.isArray(data)) {
      savedCombos = data;
      // ドライブゲージを再計算
      savedCombos.forEach(combo => {
        combo.driveGauge = calculateDriveGauge(combo.combo);
      });
      saveCombosToLocal();
      savedCombos.forEach(c => upsertComboToSupabase(c));
      renderComboList();
      alert('JSONファイルからコンボを読み込みました');
    } else {
      alert('不正なJSON形式です');
    }
  } catch (e) {
    alert('ファイルの読み込みに失敗しました: ' + e.message);
  }
}