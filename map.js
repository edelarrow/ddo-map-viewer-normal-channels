import mapParams from './resources/map_params.json' with {type: "json"};
import landmarkData from './resources/landmarks.json' with {type: "json"};
import connectionData from './resources/connections.json' with {type: "json"};
import specialConnectionRules from './resources/specialConnections.json' with {type: "json"};
import gatherPoints from './resources/gatherPoints.json' with {type: "json"};
import stageIds from './resources/stageIds.json' with {type: "json"};
import itemNames from './resources/itemNames.json' with {type: "json"};
import emNames from './resources/emNames.json' with {type: "json"};
import iconIds from './resources/iconIds.json' with {type: "json"};
import npcShops from './resources/npcShops.json' with {type: "json"};
import npcSpecialShops from './resources/npcSpecialShops.json' with {type: "json"};
import npcNames from './resources/npcNames.json' with {type: "json"};
import namedParamList from './resources/namedParams.json'  with {type: "json"};
import hmPresetList   from './resources/hmPresets.json'    with {type: "json"};
import emThinkInfo      from './resources/emThinkInfo.json'      with {type: "json"};
import thinkTableNotes from './resources/thinkTableNotes.json' with {type: "json"};
import emMontageInfo   from './resources/emMontageInfo.json'   with {type: "json"};
import montageNotes    from './resources/montageNotes.json'    with {type: "json"};
import stageGroups   from './resources/stageGroups.json'   with {type: "json"};
import worldFlags      from './resources/worldFlags.json'      with {type: "json"};
import worldFlagsExtra from './resources/worldFlagsExtra.json' with {type: "json"};
import worldQuestFlags from './resources/worldQuestFlags.json' with {type: "json"};
import emRadii        from './resources/emRadii.json'        with {type: "json"};
import stageList      from './resources/stage_list.slt.json' with {type: "json"};

// ── Enemy spawn positions (lazy-loaded per stage from resources/enemyPositions/) ──
const ENEMY_POSITIONS_DIR = './resources/enemyPositions';
const ENEMY_POS_BATCH_SIZE = 16;
const _enemyPosStageCache = new Map();
const _enemyPosStageLoads = new Map();
let _enemyPosStageList = null;
let _enemyPosStageSet  = null;

const fetchEnemyPositionsIndex = async () => {
    if (_enemyPosStageList) return _enemyPosStageList;
    const res = await fetch(`${ENEMY_POSITIONS_DIR}/index.json`);
    if (!res.ok) throw new Error(`enemy positions index HTTP ${res.status}`);
    const data = await res.json();
    _enemyPosStageList = data.stages ?? [];
    _enemyPosStageSet = new Set(_enemyPosStageList);
    return _enemyPosStageList;
};

const ensureEnemyPosStageSet = async () => {
    if (_enemyPosStageSet) return _enemyPosStageSet;
    await fetchEnemyPositionsIndex();
    return _enemyPosStageSet ?? new Set();
};

const fetchEnemyPositionsStage = async (stageNo) => {
    const key = String(stageNo);
    if (_enemyPosStageCache.has(key)) return _enemyPosStageCache.get(key);
    if (_enemyPosStageLoads.has(key)) return _enemyPosStageLoads.get(key);
    const load = ensureEnemyPosStageSet().then(async (set) => {
        if (!set.has(key)) {
            _enemyPosStageCache.set(key, null);
            return null;
        }
        const res = await fetch(`${ENEMY_POSITIONS_DIR}/${key}.json`);
        const data = res.ok ? await res.json() : null;
        _enemyPosStageCache.set(key, data);
        return data;
    }).catch(() => {
        _enemyPosStageCache.set(key, null);
        return null;
    }).finally(() => {
        _enemyPosStageLoads.delete(key);
    });
    _enemyPosStageLoads.set(key, load);
    return load;
};

const fetchEnemyPositionsStages = (stageNos) =>
    Promise.all([...new Set(stageNos.map((n) => String(n)))].map(fetchEnemyPositionsStage));

const enemyPositionsForStage = (stageNo) =>
    _enemyPosStageCache.get(String(stageNo)) ?? null;

const preloadAllEnemyPositionsStages = async (onProgress) => {
    const stages = await fetchEnemyPositionsIndex();
    for (let i = 0; i < stages.length; i += ENEMY_POS_BATCH_SIZE) {
        const batch = stages.slice(i, i + ENEMY_POS_BATCH_SIZE);
        await Promise.all(batch.map(fetchEnemyPositionsStage));
        onProgress?.(Math.min(i + batch.length, stages.length), stages.length);
    }
};

const _iconIdSet = new Set(iconIds);
const itemWikiHref = (itemId) => `build/i${String(itemId).padStart(8, '0')}.html`;
// Build lookup map: id → named param entry
const namedParamsById = new Map(namedParamList.map(p => [p.id, p]));
const hmPresetsByEmCode = new Map(hmPresetList.filter(p => p.emCode).map(p => [p.emCode, p]));
// Helper: given a named param entry, return display label for the picker button
function namedParamLabel(p) {
    if (!p) return '0 (None)';
    const trimName = p.name?.trim();
    const typeSuffix = p.type === 'NAMED_TYPE_PREFIX' ? ' [Pfx]'
        : p.type === 'NAMED_TYPE_REPLACE'             ? ' [Rep]'
        : p.type === 'NAMED_TYPE_SUFFIX'              ? ' [Sfx]'
        : '';
    return trimName ? `${p.id}: ${trimName}${typeSuffix}` : `#${p.id}${typeSuffix}`;
}
// Helper: build inner stats table for a named param entry (used in edit panel)
function buildNamedStatsInner(p) {
    if (!p) return '';
    const stat = (label, val) => {
        if (val == null) return '';
        const cls = val > 100 ? 'nsi-high' : val < 100 ? 'nsi-low' : 'nsi-norm';
        return `<tr><td class="nsi-lbl">${label}</td><td class="${cls}">${val}%</td></tr>`;
    };
    return `<table class="nsi-table">` +
        stat('HP', p.hp) + stat('Atk P', p.atkP) + stat('Atk M', p.atkM) +
        stat('Def P', p.defP) + stat('Def M', p.defM) +
        stat('Exp', p.exp) + stat('Power', p.power) +
        `</table>`;
}
// Returns a persistent wrapper div (always present; hidden when namedId is 0)
function buildNamedStatsHtml(namedId) {
    const p = namedId ? namedParamsById.get(namedId) : null;
    const inner = p ? buildNamedStatsInner(p) : '';
    return `<div class="se-named-stats"${!inner ? ' style="display:none"' : ''}>${inner}</div>`;
}
// ── Named-stats companion panel (floats beside the enemy popup in edit mode) ──
function buildNamedParamPanelHtml(p, baseEmName) {
    if (!p || p.id === 0) return '';
    const typeName = p.type.replace('NAMED_TYPE_', '');
    const trimName = p.name?.trim();
    let combined = null;
    if (trimName) {
        const em = baseEmName ?? '…';
        const hi = `<span class="np-name-hi">${trimName}</span>`;
        if      (p.type === 'NAMED_TYPE_PREFIX')  combined = `${hi} ${em}`;
        else if (p.type === 'NAMED_TYPE_SUFFIX')  combined = `${em} ${hi}`;
        else if (p.type === 'NAMED_TYPE_REPLACE') combined = hi;
    }
    const pct = v => v != null ? `${v}%` : '—';
    const row = (label, val) => {
        const v = pct(val);
        const n = parseFloat(v);
        const cls = n > 100 ? 'np-high' : n < 100 ? 'np-low' : '';
        return `<tr><td>${label}</td><td${cls ? ` class="${cls}"` : ''}>${v}</td></tr>`;
    };
    const sec = (title, ...rows) =>
        `<tr class="np-stat-sec"><td colspan="2">${title}</td></tr>` + rows.join('');
    return (combined ? `<div class="np-preview-combined">${combined}</div>` : '') +
        `<div class="np-preview-type">${typeName} · ID ${p.id}</div>` +
        `<table class="np-stat-table">` +
        sec('HP',
            row('HP Rate',  p.hp),
            row('HP Sub',   p.hpSub)) +
        sec('Attack',
            row('Base Phys',  p.atkP),
            row('Base Magic', p.atkM),
            row('Wep Phys',   p.atkWepP),
            row('Wep Magic',  p.atkWepM)) +
        sec('Defence',
            row('Base Phys',  p.defP),
            row('Base Magic', p.defM),
            row('Wep Phys',   p.defWepP),
            row('Wep Magic',  p.defWepM),
            row('Guard Base', p.guardBase),
            row('Guard Wep',  p.guardWep)) +
        sec('Other',
            row('Ailment Dmg', p.ailment),
            row('Experience',  p.exp),
            row('Power',       p.power)) +
        sec('Endurance',
            row('Blow Main',   p.blowMain),
            row('Blow Sub',    p.blowSub),
            row('Down Main',   p.downMain),
            row('OCD',         p.ocd),
            row('Shake Main',  p.shakeMain),
            row('Shrink Main', p.shrinkMain),
            row('Shrink Sub',  p.shrinkSub)) +
        `</table>`;
}
let _namedStatsPanelAnchor = null;
function _repositionNamedStatsPanel() {
    const panel = document.getElementById('named-stats-panel');
    if (!panel || panel.style.display === 'none' || !_namedStatsPanelAnchor) return;
    const r  = _namedStatsPanelAnchor.getBoundingClientRect();
    const pw = panel.offsetWidth || 150;
    let left = r.right + 8;
    if (left + pw > window.innerWidth - 8) left = r.left - pw - 8;
    panel.style.left = Math.max(8, left) + 'px';
    panel.style.top  = r.top + 'px';
}
function showNamedStatsPanel(namedId, anchorEl, baseEmName = null) {
    const panel = document.getElementById('named-stats-panel');
    if (!panel) return;
    const p    = namedId ? namedParamsById.get(namedId) : null;
    const html = p ? buildNamedParamPanelHtml(p, baseEmName) : '';
    if (!html) { panel.style.display = 'none'; _namedStatsPanelAnchor = null; return; }
    panel.innerHTML = html;
    panel.style.display = 'block';
    _namedStatsPanelAnchor = anchorEl ?? null;
    _repositionNamedStatsPanel();
}
function hideNamedStatsPanel() {
    const panel = document.getElementById('named-stats-panel');
    if (panel) panel.style.display = 'none';
    _namedStatsPanelAnchor = null;
}

// ── Leaflet map setup ──────────────────────────────────────────────────────────
const leafletMap = L.map('map', {
    crs: L.CRS.Simple,
    maxZoom: 6,
    minZoom: -3,
    zoomSnap: 0.5,
});
// Dedicated pane for the map background image — z-index 201 keeps it below
// the overlayPane (400) so polylines (pd boundaries etc.) always render on top
// even after swapMapImage recreates the imageOverlay.
leafletMap.createPane('mapImagePane');
leafletMap.getPane('mapImagePane').style.zIndex = 201;

function xy(x, y) { return L.latLng(y, x); }

// ── Reset-view control (appears below zoom +/−) ────────────────────────────────
L.Control.ResetView = L.Control.extend({
    options: { position: 'topleft' },
    onAdd() {
        const container = L.DomUtil.create('div', 'leaflet-bar leaflet-control');
        const btn = L.DomUtil.create('a', 'leaflet-control-reset-view', container);
        btn.innerHTML = '&#8962;';  // ⌂ home symbol
        btn.title = 'Reset view';
        btn.href = '#';
        btn.setAttribute('role', 'button');
        btn.setAttribute('aria-label', 'Reset view');
        L.DomEvent.on(btn, 'click', (e) => { L.DomEvent.preventDefault(e); resetView(); });
        return container;
    },
});
new L.Control.ResetView().addTo(leafletMap);

// ── Day / night spawn filter (below zoom +/− and reset view) ─────────────────
const SPAWN_TIME_FILTER_KEY = 'ddon-spawn-time-filter';
let _activeSpawnTimeFilter  = null;   // null = all, 'day' | 'night'

const loadSpawnTimeFilter = () => {
    try {
        const v = localStorage.getItem(SPAWN_TIME_FILTER_KEY);
        return v === 'day' || v === 'night' ? v : null;
    } catch { return null; }
};

const saveSpawnTimeFilter = () => {
    try {
        if (_activeSpawnTimeFilter) localStorage.setItem(SPAWN_TIME_FILTER_KEY, _activeSpawnTimeFilter);
        else localStorage.removeItem(SPAWN_TIME_FILTER_KEY);
    } catch { /* ignore */ }
};

_activeSpawnTimeFilter = loadSpawnTimeFilter();

L.Control.SpawnTimeFilter = L.Control.extend({
    options: { position: 'topleft' },
    onAdd() {
        this._container = L.DomUtil.create('div', 'leaflet-bar leaflet-control leaflet-control-spawn-time');
        this._buttons = [];
        const modes = [
            { filter: null,    cls: 'spawn-time-all',   html: 'All', title: 'All spawn times' },
            { filter: 'day',   cls: 'spawn-time-day',   html: '☀',  title: 'Day spawns only' },
            { filter: 'night', cls: 'spawn-time-night', html: '🌙',  title: 'Night spawns only' },
        ];
        for (const mode of modes) {
            const btn = L.DomUtil.create('a', mode.cls, this._container);
            btn.innerHTML = mode.html;
            btn.href = '#';
            btn.title = mode.title;
            btn.setAttribute('role', 'button');
            L.DomEvent.on(btn, 'click', (e) => {
                L.DomEvent.preventDefault(e);
                L.DomEvent.stopPropagation(e);
                _activeSpawnTimeFilter = mode.filter;
                saveSpawnTimeFilter();
                this._updateActive();
                applySubGroupFilter();
            });
            this._buttons.push({ btn, filter: mode.filter });
        }
        L.DomEvent.disableClickPropagation(this._container);
        L.DomEvent.disableScrollPropagation(this._container);
        this._updateActive();
        return this._container;
    },
    _updateActive() {
        for (const { btn, filter } of this._buttons) {
            btn.classList.toggle('active', _activeSpawnTimeFilter === filter);
        }
    },
});
new L.Control.SpawnTimeFilter().addTo(leafletMap);

// ── Mob type filter (blood orb, high orb, manual, boss, key bearer, regular) ───
const MOB_TYPE_FILTER_KEY = 'ddon-mob-type-filters';
const MOB_TYPES_SPOT_KEY  = 'ddon-mob-types-apply-spot-search';
const MOB_TYPE_DEFAULTS   = {
    bloodOrb: true, highOrb: true, manual: true, boss: true, keyBearer: true, regular: true, dynamic: true,
};
const MOB_TYPE_MODES = [
    { id: 'bloodOrb',   html: '🩸', title: 'Blood orb' },
    { id: 'highOrb',    html: '⭐', title: 'High orb' },
    { id: 'manual',     html: '😴', title: 'Dormant' },
    { id: 'boss',       html: '☠', title: 'Boss' },
    { id: 'keyBearer',  html: '🗝', title: 'Key Mobs' },
    { id: 'regular',    html: '',   title: 'Regular' },
    { id: 'dynamic',    html: '⚡', title: 'Dynamic' },
];

const loadMobTypeFilters = () => {
    try {
        const raw = localStorage.getItem(MOB_TYPE_FILTER_KEY);
        return raw ? { ...MOB_TYPE_DEFAULTS, ...JSON.parse(raw) } : { ...MOB_TYPE_DEFAULTS };
    } catch { return { ...MOB_TYPE_DEFAULTS }; }
};

const saveMobTypeFilters = () => {
    try { localStorage.setItem(MOB_TYPE_FILTER_KEY, JSON.stringify(_mobTypeFilters)); } catch { /* ignore */ }
};

let _mobTypeFilters = loadMobTypeFilters();
let _applyMobTypesToSpotSearch = (() => {
    try { return localStorage.getItem(MOB_TYPES_SPOT_KEY) === '1'; }
    catch { return false; }
})();

const saveApplyMobTypesToSpotSearch = () => {
    try { localStorage.setItem(MOB_TYPES_SPOT_KEY, _applyMobTypesToSpotSearch ? '1' : '0'); }
    catch { /* ignore */ }
};

function anyMobTypeEnabled() {
    return MOB_TYPE_MODES.some(m => _mobTypeFilters[m.id]);
}

/** Refresh open spot panel when Spawns filters (or the apply-to-search option) change. */
function refreshSpotSearchFromMobTypes() {
    if (document.getElementById('spot-panel')?.classList.contains('open')) _runSpotSearch();
}

function syncMobTypesShowAllCheckbox() {
    const el = document.getElementById('mob-types-show-all');
    if (!el) return;
    const all = MOB_TYPE_MODES.every(m => _mobTypeFilters[m.id]);
    const none = MOB_TYPE_MODES.every(m => !_mobTypeFilters[m.id]);
    el.checked = all;
    el.indeterminate = !all && !none;
}

function setAllMobTypeFilters(on) {
    for (const m of MOB_TYPE_MODES) _mobTypeFilters[m.id] = on;
    saveMobTypeFilters();
    for (const m of MOB_TYPE_MODES) {
        const inp = document.getElementById(`mob-type-filter-${m.id}`);
        if (inp) inp.checked = on;
    }
    syncMobTypesShowAllCheckbox();
    updateEnemyVisibility();
    applySubGroupFilter();
    if (_applyMobTypesToSpotSearch) refreshSpotSearchFromMobTypes();
}

function initMobTypeFilters() {
    // Migrate legacy “Enemy Spawns” layer off → all mob types off (once).
    try {
        if (!localStorage.getItem('ddon-mob-type-migrated-enemies')) {
            let legacyOff = false;
            try {
                const layerRaw = localStorage.getItem(LAYER_PREFS_KEY);
                if (layerRaw && JSON.parse(layerRaw).enemies === false) legacyOff = true;
            } catch { /* ignore */ }
            const { layers: urlLayers } = parseHash();
            if (urlLayers?.enemies === false) legacyOff = true;
            if (legacyOff) {
                for (const k of Object.keys(MOB_TYPE_DEFAULTS)) _mobTypeFilters[k] = false;
                saveMobTypeFilters();
                const layerEnemies = document.getElementById('layer-enemies');
                if (layerEnemies) layerEnemies.checked = false;
            }
            localStorage.setItem('ddon-mob-type-migrated-enemies', '1');
        }
    } catch { /* ignore */ }

    const list = document.getElementById('mob-type-filter-list');
    if (!list) return;
    list.innerHTML = '';
    for (const mode of MOB_TYPE_MODES) {
        const label = document.createElement('label');
        label.className = 'layer-toggle';
        label.title = mode.title;
        const input = document.createElement('input');
        input.type = 'checkbox';
        input.id = `mob-type-filter-${mode.id}`;
        input.checked = !!_mobTypeFilters[mode.id];
        input.addEventListener('change', () => {
            _mobTypeFilters[mode.id] = input.checked;
            saveMobTypeFilters();
            syncMobTypesShowAllCheckbox();
            updateEnemyVisibility();
            applySubGroupFilter();
            if (_applyMobTypesToSpotSearch) refreshSpotSearchFromMobTypes();
        });
        const icon = document.createElement('span');
        icon.className = 'mob-type-icon';
        if (mode.html) {
            icon.textContent = mode.html;
            label.append(input, icon, document.createTextNode(` ${mode.title}`));
        } else {
            label.append(input, document.createTextNode(` ${mode.title}`));
        }
        list.appendChild(label);
    }
    const showAll = document.getElementById('mob-types-show-all');
    showAll?.addEventListener('change', () => setAllMobTypeFilters(showAll.checked));

    const applySpot = document.getElementById('mob-types-apply-spot-search');
    if (applySpot) {
        applySpot.checked = _applyMobTypesToSpotSearch;
        applySpot.addEventListener('change', () => {
            _applyMobTypesToSpotSearch = applySpot.checked;
            saveApplyMobTypesToSpotSearch();
            refreshSpotSearchFromMobTypes();
        });
    }

    syncMobTypesShowAllCheckbox();
    updateEnemyVisibility();
}


// ── World → pixel conversion ───────────────────────────────────────────────────
// Formula from GMP data + engine scale constant (derived from PS4 disassembly):
//   pixelX = worldX * scale   + center_x
//   pixelY = (imgHeight - center_y) - worldZ * scale_z
// scale_z may differ from scale when empirically calibrated (see map_params.json).
//
// For pd (parts-dungeon) maps, pd_pieces stores per-piece pixel boundaries so
// each piece can have its own scale (piecewise-linear mapping).
function worldToPixel(worldX, worldZ, info) {
    let png_y;
    if (info.pd_pieces?.length) {
        // Piecewise-linear. Image stacked deepest-first: top=deepest, bottom=entrance.
        // Each piece slot: entrance at pixel_y_entrance (bottom), deepest at pixel_y_start (top).
        const pieces = info.pd_pieces;
        let piece = pieces[0]; // default: shallowest (entrance)
        for (const p of pieces) {
            // Use full_size (original Z range) for lookup so there are no Z gaps.
            const rangeSize = p.full_size ?? p.size;
            if (worldZ >= p.connect_z + rangeSize && worldZ <= p.connect_z) {
                piece = p;
                break;
            }
        }
        const localZ  = worldZ - piece.connect_z;                          // ≤ 0
        // Use pixel_y_entrance_v (virtual entrance, accounts for bottom trim) and
        // info.scale (DUNGEON_MAP_SCALE, same as X axis) — the correct rendering scale.
        png_y = piece.pixel_y_entrance_v + localZ * info.scale;
        png_y = Math.max(piece.pixel_y_start, Math.min(piece.pixel_y_entrance, png_y));
    } else {
        const scaleZ = info.scale_z ?? info.scale;
        png_y = (info.img_height - info.center_y) - worldZ * scaleZ;
        // above is lat; convert to png_y
        png_y = info.img_height - png_y;
    }
    const py = info.img_height - png_y;   // lat (Leaflet y from bottom)
    const px = worldX * info.scale + info.center_x;
    return xy(px, py);
}

// ── Layer groups ───────────────────────────────────────────────────────────────
let imageOverlay    = null;
let enemyLayer      = L.layerGroup().addTo(leafletMap);  // group chip labels
let landmarkLayer   = L.layerGroup().addTo(leafletMap);
let connectionLayer = L.layerGroup().addTo(leafletMap);
let gridLayer        = L.layerGroup();   // off by default
let territoryLayer   = L.layerGroup();   // off by default; territory rects when groups expand
let stageLabelsLayer = L.layerGroup().addTo(leafletMap);  // area name text labels
let gatherLayer       = L.layerGroup().addTo(leafletMap);
let _gatherMarkerByKey   = new Map();    // "${stageNo}:${groupId}:${posId}" → L.marker
let _gatherGroupMarkers  = new Map();    // "${stageNo}:${groupId}" → L.marker[]
let _shopMarkerByNpcId = new Map();    // "${stageNo}:${npcId}" → L.marker
let _specialShopMarkerByNpcId = new Map();
let npcShopLayer        = L.layerGroup().addTo(leafletMap);
let specialShopLayer    = L.layerGroup().addTo(leafletMap);
let pdBoundaryLayer = L.layerGroup().addTo(leafletMap);
let spawnRadiiLayer   = L.layerGroup().addTo(leafletMap);  // aggro/link radius circles
let _spreadOverlay    = L.layerGroup().addTo(leafletMap);  // cross-group spoke lines + anchor dots

// Canvas renderer — all spawn circleMarkers share one <canvas> element (huge perf win).
const spawnRenderer = L.canvas({ padding: 0.5 });

// ── Group expand/collapse state ───────────────────────────────────────────────
// One entry per group; detailsLayer is lazily created on first expand.
// Must be `let` — map scene cache swaps the Map reference on restore.
let _groupStore = new Map(); // groupId string → { groupId, color, territory, items, pts,
                               //   centroid, labelMarker, detailsLayer, isExpanded }
let _currentMapInfo  = null;   // stored at loadEnemySpawns time; used by lazy expand

// ── Edit mode state ────────────────────────────────────────────────────────────
let _editMode        = false;
let _spawnSetMode    = false;   // when true, edits apply to all positions sharing the same SpawnGroup (sg) in the group
let _activeSubGroupId   = null; // map-wide filter: null = show all, number = show only that splitId
let _availableSubGroups = [0]; // distinct subGroupId values found in the loaded spawn data
let _editDirty       = false;       // any unsaved changes this session
let _dirtySet        = new Set();   // which source keys have unsaved changes
let _rawEnemyData    = null;        // full EnemySpawn.json object kept for write-back
let _rawEnemySchemas = null;        // field-index shortcuts (same as iLv etc. but accessible globally)
let _copiedEnemyConfig = null;      // cross-marker copy/paste clipboard
let _rawGatheringRows    = null;    // mutable array of CSV row-objects for write-back
let _rawGatheringHeaders = null;    // original CSV header string (including leading '#')
let _rawShopData          = null;   // full Shop.json array kept for write-back
let _rawSpecialShopData   = null;   // full SpecialShops.json object kept for write-back
let _currentFloorObbs = null;
let _dragItemId        = null;        // item ID being dragged from the Items panel
let _gatherPopupDropFn = null;        // fn(itemId) for the currently open gather popup
let _shopPopupDropFn   = null;        // fn(itemId) for the currently open shop popup
let _dragEmCode        = null;        // emCode being dragged from the Enemies panel
let _spawnPopupDropFn  = null;        // fn(emCode) for the currently open spawn popup
let _dropsTablesMap    = new Map();   // id → {id, name, mdlType, items[]} — populated on spawn parse
let _markDirty         = null;        // set by edit block; callable from gather/shop code
let _attachDragReorder = null;        // set by edit block; callable from gather/shop code
let _renderEditPanel   = null;        // set by edit block; callable from spawn popup code
let _rebuildOpenPopup  = null;        // set on enemy popupopen; rebuilds active popup HTML
let _dtEditorReadAndSave = null;      // set by openDropTableEditor; saves + closes the editor

function updateEnemyVisibility() {
    const layerOn = document.getElementById('layer-enemies')?.checked ?? true;
    const show = layerOn && anyMobTypeEnabled();
    if (show) {
        leafletMap.addLayer(enemyLayer);
        leafletMap.addLayer(_spreadOverlay);
        for (const g of _groupStore.values())
            if (g.isExpanded && g.detailsLayer) g.detailsLayer.addTo(leafletMap);
    } else {
        leafletMap.removeLayer(enemyLayer);
        leafletMap.removeLayer(_spreadOverlay);
        for (const g of _groupStore.values())
            if (g.isExpanded && g.detailsLayer) leafletMap.removeLayer(g.detailsLayer);
    }
}

// ── Layer preference persistence ───────────────────────────────────────────────
// ── Layer state — URL hash + localStorage ─────────────────────────────────────
// Hash format extension: …@zoom/y/x!elcgt
//   Each letter present = that layer is ON: e=enemies l=landmarks c=connections
//   g=grid t=territory.  Absent = OFF.
// URL state takes priority over localStorage (enables sharing exact views).

const LAYER_PREFS_KEY = 'ddon-maps-layers';
const DEV_PREFS_KEY     = 'ddon-dev-prefs';
const DEV_PANEL_KEY     = 'ddon-dev-panel-open';

let _devMobSpawnLabels   = false;
const DYNAMIC_ENEMY_LABEL = 'Dynamic Enemy';
let _enemySpawnDataLoaded  = false;

const useLegacyGroupChips = () => _devMobSpawnLabels;
const useMobSpawnDevLabels = () => _devMobSpawnLabels;

const loadDevPrefs = () => {
    try {
        const raw = localStorage.getItem(DEV_PREFS_KEY);
        return raw ? JSON.parse(raw) : {};
    } catch { return {}; }
};

const readMobSpawnLabelsPref = (devPrefs) => {
    if (devPrefs.mobSpawnLabels != null) return !!devPrefs.mobSpawnLabels;
    return !!(devPrefs.showSpawnIds || devPrefs.showGroupIds || devPrefs.showGroupCounts
        || devPrefs.legacyGroupChips);
};

const saveDevPrefs = () => {
    const prefs = { mobSpawnLabels: _devMobSpawnLabels };
    try { localStorage.setItem(DEV_PREFS_KEY, JSON.stringify(prefs)); } catch { /* ignore */ }
};

const applyDevDisplayPrefs = () => {
    document.body.classList.add('dev-coords-on');
};

// Territory bounds exist in spawn position data but are not drawn — same as upstream
// pacampbell.github.io. Most groups store stage-sized defaults (~130 km), not a tight
// trigger zone, so rectangles were misleading when expand-all was used in dev mode.
function removeGroupTerritoryRect(g) {
    g.territoryRect = null;
}

function removeAllGroupTerritoryRects() {
    for (const g of _groupStore.values()) removeGroupTerritoryRect(g);
}

const syncTerritoryLayer = () => {
    territoryLayer.clearLayers();
    if (leafletMap.hasLayer(territoryLayer)) leafletMap.removeLayer(territoryLayer);
};

const refreshGroupChipIcons = () => {
    for (const g of _groupStore.values()) {
        if (!g.labelMarker) continue;
        g.labelMarker.setIcon(makeChipIcon(
            g.groupId, g.color, g.items.length, g.isExpanded, g.yOffset,
            g.isKeyBearerGroup, _groupHasBoss(g),
        ));
    }
};

const refreshMobTooltips = () => {
    for (const g of _groupStore.values()) {
        if (!g.detailsLayer) continue;
        for (const layer of g.detailsLayer.getLayers()) {
            if (!layer._rebuildTooltip) continue;
            const tt = layer._rebuildTooltip();
            layer._label = tt;
            layer._naturalTooltip = tt;
            const display = layer._spreadCount
                ? `${tt} <span style="opacity:0.7">[×${layer._spreadCount} stacked]</span>`
                : tt;
            if (layer.isTooltipOpen()) layer.setTooltipContent(display);
            else layer.bindTooltip(display, { direction: 'top', offset: [0, -8] });
        }
    }
};

let _mobDisplayModeJob = 0;
const MOB_EXPAND_CHUNK = 12;

const applyMobDisplayModeFinish = (legacy) => {
    for (const g of _groupStore.values()) {
        if (legacy) ensureGroupStructureLayers(g);
        else removeGroupStructureLayers(g);
        if (!g.labelMarker) continue;
        const showChip = legacy;
        g.labelMarker.setOpacity(showChip ? 1 : 0);
        g.labelMarker.options.interactive = showChip;
        const chipEl = g.labelMarker.getElement();
        if (chipEl) chipEl.style.pointerEvents = showChip ? '' : 'none';
        syncGroupStructureVisibility(g, true);
    }
    refreshGroupChipIcons();
    reapplySpread();
    _updateExpandCollapseBtn();
    if (!legacy) removeAllGroupTerritoryRects();
    syncTerritoryLayer();
};

const applyMobDisplayMode = () => {
    const jobId = ++_mobDisplayModeJob;
    const legacy = useLegacyGroupChips();
    const expandBtn = document.getElementById('btn-expand-collapse');
    if (expandBtn) expandBtn.style.display = legacy ? '' : 'none';

    if (legacy) {
        applyMobDisplayModeFinish(true);
        return;
    }

    // Player mode — expand any collapsed groups in chunks so the UI stays responsive.
    const pending = [..._groupStore.values()].filter(g => !g.isExpanded);

    for (const g of _groupStore.values()) {
        removeGroupStructureLayers(g);
        if (!g.labelMarker) continue;
        g.labelMarker.setOpacity(0);
        g.labelMarker.options.interactive = false;
        const chipEl = g.labelMarker.getElement();
        if (chipEl) chipEl.style.pointerEvents = 'none';
    }

    if (!pending.length) {
        applyMobDisplayModeFinish(false);
        return;
    }

    let idx = 0;
    const expandChunk = () => {
        if (jobId !== _mobDisplayModeJob) return;
        const end = Math.min(idx + MOB_EXPAND_CHUNK, pending.length);
        for (; idx < end; idx++) _expandGroupCore(pending[idx], { skipFilter: true });
        if (idx < pending.length) {
            setTimeout(expandChunk, 0);
            return;
        }
        applySubGroupFilter();
        applyMobDisplayModeFinish(false);
    };
    expandChunk();
};

// Returns the !-suffix string: layer flags + optional ;groupId,groupId,...
// Format: !elcgt;0,3,80
function getLayersHash() {
    let s = '';
    if (document.getElementById('layer-enemies')?.checked)          s += 'e';
    if (document.getElementById('layer-grid')?.checked)             s += 'g';
    if (useMobSpawnDevLabels())                                     s += 't';
    if (document.getElementById('layer-radii')?.checked)            s += 'i';
    if (document.getElementById('layer-stage-labels')?.checked)     s += 'a';
    if (document.getElementById('sidebar')?.classList.contains('collapsed')) s += 's';
    const openIds = [..._groupStore.values()]
        .filter(g => g.isExpanded)
        .map(g => g.groupId)
        .sort((a, b) => parseInt(a) - parseInt(b));
    if (openIds.length) s += ';' + openIds.join(',');
    return s;
}

// Update only the layer portion of the current hash without triggering hashchange.
function updateLayersInHash() {
    const { name, stid } = parseHash();
    const mapName = name || _loadedMapName;
    if (!mapName) return;
    const z = leafletMap.getZoom().toFixed(2);
    const c = leafletMap.getCenter();
    const hasFloors = !!(mapParams[_loadedMapName]?.floor_obbs);
    const floorSuffix = hasFloors ? `/${currentLayer}` : '';
    const frag = (stid ? `${mapName}:${stid}` : mapName)
               + `@${z}/${c.lat.toFixed(1)}/${c.lng.toFixed(1)}${floorSuffix}`
               + `!${getLayersHash()}`;
    history.replaceState(null, '', '#' + frag);
}

function saveLayerPrefs() {
    const prefs = {
        enemies: document.getElementById('layer-enemies')?.checked ?? true,
        grid:  document.getElementById('layer-grid')?.checked ?? false,
        radii: document.getElementById('layer-radii')?.checked ?? false,
        stageLabels: document.getElementById('layer-stage-labels')?.checked ?? true,
    };
    try { localStorage.setItem(LAYER_PREFS_KEY, JSON.stringify(prefs)); } catch (_) {}
    updateLayersInHash();
}

function loadLayerPrefs() {
    try {
        const raw = localStorage.getItem(LAYER_PREFS_KEY);
        return raw ? JSON.parse(raw) : null;
    } catch (_) { return null; }
}

// Apply layer state from URL hash (if present) or localStorage, then fall back to defaults.
// Must run after layers are declared but before first loadMap.
(function applyLayerPrefs() {
    const { layers: urlLayers } = parseHash();
    const stored = loadLayerPrefs();
    // URL hash wins over localStorage; localStorage wins over hardcoded defaults.
    const prefs = urlLayers ?? stored ?? {};
    const isOn = (key, defaultOn) => key in prefs ? prefs[key] : defaultOn;

    const enemiesEl = document.getElementById('layer-enemies');
    const gridEl = document.getElementById('layer-grid');
    const radiiEl = document.getElementById('layer-radii');
    const labelsEl = document.getElementById('layer-stage-labels');
    if (enemiesEl) enemiesEl.checked = isOn('enemies', true);
    if (gridEl) gridEl.checked = isOn('grid', false);
    if (radiiEl) radiiEl.checked = isOn('radii', false);
    if (labelsEl) labelsEl.checked = isOn('stageLabels', true);

    if (gridEl?.checked) leafletMap.addLayer(gridLayer);
    if (labelsEl && !labelsEl.checked) leafletMap.removeLayer(stageLabelsLayer);
    if (!enemiesEl?.checked || !anyMobTypeEnabled())
        updateEnemyVisibility();
    if (isOn('sidebarHidden', false))
        document.getElementById('sidebar')?.classList.add('collapsed');
})();

// ── Layer toggles ──────────────────────────────────────────────────────────────
document.getElementById('layer-enemies')?.addEventListener('change', () => {
    updateEnemyVisibility();
    saveLayerPrefs();
});
document.getElementById('layer-grid')?.addEventListener('change', e => {
    e.target.checked ? leafletMap.addLayer(gridLayer) : leafletMap.removeLayer(gridLayer);
    saveLayerPrefs();
});
document.getElementById('layer-radii')?.addEventListener('change', e => {
    if (!e.target.checked) clearSpawnRadii();
    saveLayerPrefs();
});
document.getElementById('layer-stage-labels')?.addEventListener('change', e => {
    if (e.target.checked) leafletMap.addLayer(stageLabelsLayer);
    else leafletMap.removeLayer(stageLabelsLayer);
    saveLayerPrefs();
});
// ── Sidebar collapse / expand ──────────────────────────────────────────────────
function setSidebarCollapsed(collapsed) {
    document.getElementById('sidebar').classList.toggle('collapsed', collapsed);
    document.getElementById('sidebar-toggle').style.display = collapsed ? 'block' : 'none';
    leafletMap.invalidateSize();
    updateLayersInHash();
}
document.getElementById('sidebar-collapse').addEventListener('click', () => setSidebarCollapsed(true));
document.getElementById('sidebar-toggle').addEventListener('click',   () => setSidebarCollapsed(false));

document.getElementById('btn-expand-collapse')?.addEventListener('click', () => {
    const anyCollapsed = [..._groupStore.values()].some(g => !g.isExpanded);
    if (anyCollapsed) _expandAllGroups(); else _collapseAllGroups();
});

// ── Sidebar map list ───────────────────────────────────────────────────────────
function splitPascalCase(s) {
    // Pre-split pass: handle "to" preposition glued to the next PascalCase word
    // (e.g. "PathtoMorrow" → "Path to Morrow"). Must be done before basic split
    // so "to" at word-end (e.g. "Grotto") is not incorrectly split.
    // Handle "tothe" compound first ("PathtotheX" → "Path to the X").
    let result = s.replace(/([a-z])(to)(the)(?=[A-Z])/g, '$1 $2 $3 ');
    result = result.replace(/([a-z])(to)(?=[A-Z])/g, '$1 $2 ');
    // Insert space before each uppercase letter that follows a lowercase letter,
    // e.g. "TheWhiteDragonTemple" → "The White Dragon Temple"
    result = result.replace(/([a-z])([A-Z])/g, '$1 $2');
    // Insert space before "of"/"the"/"by" when glued to the preceding word (lowercase
    // prepositions embedded in PascalCase enum names), e.g. "Forestof Mist" → "Forest of Mist",
    // "Altarofthe Black Curse" → "Altar of the Black Curse",
    // "Wildernessby Castle" → "Wilderness by Castle".
    result = result.replace(/([a-z])(of)(?=the\b|[A-Z\s]|$)/g, '$1 $2');
    result = result.replace(/([a-z])(the)(?=[A-Z\s]|$)/g, '$1 $2');
    result = result.replace(/([ac-z])(by)(?=\s|$)/g, '$1 $2');
    // Insert space before a digit sequence, e.g. "Netherworld1" → "Netherworld 1"
    result = result.replace(/([a-zA-Z])(\d+)/g, '$1 $2');
    return result.replace(/  +/g, ' ').trim();
}

function displayName(mapName, info) {
    if (info.name_en) return splitPascalCase(info.name_en);
    return mapName;
}

function appendMapEntry(listEl, name, info, label, stid, currentMap, currentStage) {
    const isActive = name === currentMap && (stid === null ? !currentStage : stid === currentStage);
    const el = document.createElement('div');
    el.className = 'map-entry' + (isActive ? ' active' : '');
    el.dataset.map = name;

    const dot = document.createElement('span');
    dot.className = 'img-dot ' + (info.img_exists ? 'has-img' : 'no-img');
    el.appendChild(dot);

    const text = document.createElement('span');
    text.textContent = label + (info.name_en ? '' : ` (${name})`);
    el.appendChild(text);

    el.addEventListener('click', () => navigateTo(name, stid));
    listEl.appendChild(el);
}

// Render a collapsible row for multiple entries sharing the same display label.
// group: Array<{name, info, stid}> — each entry navigates independently.
function appendCollapsibleGroup(listEl, label, group, currentMap, currentStage) {
    const anyActive = group.some(e =>
        e.name === currentMap && (e.stid === null ? !currentStage : e.stid === currentStage)
    );
    const startOpen = anyActive;

    const el = document.createElement('div');
    el.className = 'map-entry' + (anyActive ? ' active' : '') + (startOpen ? ' expanded' : '');

    const arrow = document.createElement('span');
    arrow.className = 'expand-arrow';
    arrow.textContent = '▶';
    el.appendChild(arrow);

    const text = document.createElement('span');
    text.textContent = label;
    el.appendChild(text);

    const subList = document.createElement('div');
    subList.className = 'map-sublist' + (startOpen ? ' open' : '');

    for (const e of group) {
        const sub = document.createElement('div');
        const isActiveSub = e.name === currentMap &&
            (e.stid === null ? !currentStage : e.stid === currentStage);
        sub.className = 'map-subentry' + (isActiveSub ? ' active' : '');

        const subDot = document.createElement('span');
        subDot.className = 'img-dot ' + (e.info.img_exists ? 'has-img' : 'no-img');
        sub.appendChild(subDot);

        sub.appendChild(document.createTextNode(e.stid ?? e.name));
        sub.addEventListener('click', ev => { ev.stopPropagation(); navigateTo(e.name, e.stid); });
        subList.appendChild(sub);
    }

    el.addEventListener('click', () => {
        const open = subList.classList.toggle('open');
        el.classList.toggle('expanded', open);
    });

    listEl.appendChild(el);
    listEl.appendChild(subList);
}

// Sidebar world hierarchy — matches in-game minimap region menu.
const SIDEBAR_WORLD_GROUPS = [
    {
        key: 'world:lestania',
        name: 'Lestania',
        areas: [1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 10, 12],
        defaultCollapsed: true,
    },
    { key: 'area:13', areas: [13], flat: true, displayName: 'Bloodbane Isle' },
    {
        key: 'world:phindym',
        name: 'Phindym',
        areas: [14, 15, 16, 17],
    },
    {
        key: 'world:acre-selund',
        name: 'Acre Selund',
        areas: [18, 19, 20, 21],
    },
];
const SIDEBAR_GROUPED_AREA_IDS = new Set(SIDEBAR_WORLD_GROUPS.flatMap(g => g.areas));

/** Major overworld jumps — continent → region (matches in-game world menu). */
const MAJOR_LOCATION_CONTINENTS = [
    {
        id: 'lestania',
        label: 'Lestania',
        regions: [
            { map: 'field000_m00', stid: 'st0100', label: 'Lestania', sid: 1 },
            { map: 'rm110_m00',    stid: 'st0410', label: 'Mergoda Ruins', sid: 76 },
        ],
    },
    {
        id: 'bloodbane',
        label: 'Bloodbane Isle',
        regions: [
            { map: 'field003_m00', stid: 'st0110', label: 'Bloodbane Isle', sid: 335 },
            { map: 'field003_m00', stid: 'st0111', label: 'Precipice', sid: 336 },
            { map: 'field003_m00', stid: 'st0112', label: 'Summit', sid: 337 },
        ],
    },
    {
        id: 'phindym',
        label: 'Phindym',
        regions: [
            { map: 'field004_m00', stid: 'st0121', label: 'Elan Water Grove', sid: 372 },
            { map: 'field004_m00', stid: 'st0120', label: 'Farana Plains', sid: 371 },
            { map: 'field004_m00', stid: 'st0123', label: 'Morrow Forest', sid: 374 },
            { map: 'field004_m00', stid: 'st0122', label: 'Kingal Canyon', sid: 373 },
        ],
    },
    {
        id: 'acre-selund',
        label: 'Acre Selund',
        regions: [
            { map: 'field005_m00', stid: 'st0130', label: 'Rathnite Foothills', sid: 461 },
            { map: 'field005_m00', stid: 'st0131', label: 'Rathnite Foothills Lakeside', sid: 462 },
            { map: 'field005_m00', stid: 'st0132', label: 'Feryana Wilderness', sid: 463 },
            { map: 'field005_m00', stid: 'st0137', label: 'Feryana Wilderness (Lookout Castle)', sid: 579 },
            { map: 'field005_m00', stid: 'st0133', label: 'Megadosys Plateau', sid: 464 },
            { map: 'field005_m00', stid: 'st0134', label: 'Urteca Mountains', sid: 465 },
        ],
    },
];

const majorLocKey = (map, stid) => `${map}:${stid}`;

let _majorLocSyncing = false;

function populateMajorRegionSelect(continentId, { preserveRegion = '' } = {}) {
    const regionEl = document.getElementById('major-region-select');
    if (!regionEl) return;
    regionEl.innerHTML = '';
    regionEl.appendChild(new Option('— Region —', ''));
    if (!continentId) {
        regionEl.disabled = true;
        return;
    }
    const continent = MAJOR_LOCATION_CONTINENTS.find((c) => c.id === continentId);
    if (!continent) {
        regionEl.disabled = true;
        return;
    }
    regionEl.disabled = false;
    for (const region of continent.regions) {
        const val = majorLocKey(region.map, region.stid);
        const opt = new Option(region.label, val);
        opt.title = `${region.stid} · sid ${String(region.sid).padStart(4, '0')}`;
        regionEl.appendChild(opt);
    }
    if (preserveRegion && [...regionEl.options].some((o) => o.value === preserveRegion)) {
        regionEl.value = preserveRegion;
    }
}

function syncMajorLocSelectors() {
    const continentEl = document.getElementById('major-continent-select');
    const regionEl = document.getElementById('major-region-select');
    if (!continentEl || !regionEl) return;

    const map = currentMapName();
    const stid = currentStageName();
    let matchedContinent = '';
    let matchedRegion = '';

    for (const continent of MAJOR_LOCATION_CONTINENTS) {
        for (const region of continent.regions) {
            if (region.map === map && region.stid === stid) {
                matchedContinent = continent.id;
                matchedRegion = majorLocKey(region.map, region.stid);
                break;
            }
        }
        if (matchedContinent) break;
    }

    _majorLocSyncing = true;
    continentEl.value = matchedContinent;
    populateMajorRegionSelect(matchedContinent, { preserveRegion: matchedRegion });
    _majorLocSyncing = false;
}

function initMajorLocSelectors() {
    const continentEl = document.getElementById('major-continent-select');
    const regionEl = document.getElementById('major-region-select');
    if (!continentEl || !regionEl) return;

    for (const continent of MAJOR_LOCATION_CONTINENTS) {
        continentEl.appendChild(new Option(continent.label, continent.id));
    }

    continentEl.addEventListener('change', () => {
        if (_majorLocSyncing) return;
        populateMajorRegionSelect(continentEl.value);
        regionEl.value = '';
    });

    regionEl.addEventListener('change', () => {
        if (_majorLocSyncing || !regionEl.value) return;
        const sep = regionEl.value.indexOf(':');
        navigateTo(regionEl.value.slice(0, sep), regionEl.value.slice(sep + 1));
    });

    populateMajorRegionSelect('');
    syncMajorLocSelectors();
}

const MAP_AREA_OPEN_KEY = 'ddon-map-area-open';

const loadAreaOpenState = () => {
    try {
        const raw = localStorage.getItem(MAP_AREA_OPEN_KEY);
        return raw ? JSON.parse(raw) : {};
    } catch { return {}; }
};

const saveAreaOpenState = (state) => {
    try { localStorage.setItem(MAP_AREA_OPEN_KEY, JSON.stringify(state)); } catch { /* ignore */ }
};

const isEntryActive = (name, stid, currentMap, currentStage) =>
    name === currentMap && (stid === null ? !currentStage : stid === currentStage);

function renderMapEntries(container, entries, currentMap, currentStage) {
    const byLabel = new Map();
    for (const e of entries) {
        if (!byLabel.has(e.label)) byLabel.set(e.label, []);
        byLabel.get(e.label).push(e);
    }
    const sortedLabels = [...byLabel.keys()].sort((a, b) => a.localeCompare(b));
    for (const label of sortedLabels) {
        const group = byLabel.get(label);
        if (group.length === 1) {
            const e = group[0];
            appendMapEntry(container, e.name, e.info, label, e.stid, currentMap, currentStage);
        } else {
            appendCollapsibleGroup(container, label, group, currentMap, currentStage);
        }
    }
}

function appendAreaSection(listEl, areaId, areaName, entries, currentMap, currentStage, openState) {
    const anyActive = entries.some(e => isEntryActive(e.name, e.stid, currentMap, currentStage));
    const areaKey   = String(areaId);
    const startOpen = anyActive || openState[areaKey] === true;

    const section = document.createElement('div');
    section.className = 'map-area-section' + (startOpen ? ' expanded' : '');
    section.dataset.areaId = areaKey;

    const header = document.createElement('div');
    header.className = 'map-area-header';
    header.innerHTML = `<span class="expand-arrow">▶</span><span class="map-area-name">${areaName}</span><span class="map-area-count">${entries.length}</span>`;

    const body = document.createElement('div');
    body.className = 'map-area-body';
    renderMapEntries(body, entries, currentMap, currentStage);

    header.addEventListener('click', () => {
        const open = section.classList.toggle('expanded');
        openState[areaKey] = open;
        saveAreaOpenState(openState);
    });

    section.appendChild(header);
    section.appendChild(body);
    listEl.appendChild(section);
}

function appendWorldRegionSection(listEl, regionKey, regionName, subAreas, currentMap, currentStage, openState, { defaultCollapsed = false } = {}) {
    const allEntries = subAreas.flatMap(([, area]) => area.entries);
    const anyActive  = allEntries.some(e => isEntryActive(e.name, e.stid, currentMap, currentStage));
    let startOpen;
    if (regionKey in openState) {
        startOpen = openState[regionKey] === true;
    } else if (defaultCollapsed) {
        startOpen = false;
    } else {
        startOpen = anyActive;
    }

    const section = document.createElement('div');
    section.className = 'map-world-section' + (startOpen ? ' expanded' : '');
    section.dataset.worldRegion = regionKey;

    const header = document.createElement('div');
    header.className = 'map-world-header';
    header.innerHTML =
        `<span class="expand-arrow">▶</span>` +
        `<span class="map-world-name">${regionName}</span>` +
        `<span class="map-area-count">${allEntries.length}</span>`;

    const body = document.createElement('div');
    body.className = 'map-world-body';

    for (const [areaId, area] of subAreas) {
        appendAreaSection(body, areaId, area.name, area.entries, currentMap, currentStage, openState);
    }

    header.addEventListener('click', () => {
        const open = section.classList.toggle('expanded');
        openState[regionKey] = open;
        saveAreaOpenState(openState);
    });

    section.appendChild(header);
    section.appendChild(body);
    listEl.appendChild(section);
}

function stageLabel(info, stid) {
    // Display name for a specific stage variant of a map.
    const raw = info.stage_names?.[stid] || info.name_en || '';
    return raw ? splitPascalCase(raw) : stid;
}

// Parse special key=value tokens out of a search string.
// Recognised keys: stageid, stageno, area (case-insensitive).
// Returns { conditions: [{key, value}], text: remainingLowercase }.
function parseSearchQuery(raw) {
    const tokens = raw.trim().toLowerCase().split(/\s+/).filter(Boolean);
    const conditions = [];
    const textParts = [];
    for (const tok of tokens) {
        const m = tok.match(/^(stageid|stageno|area)=(.+)$/);
        if (m) conditions.push({ key: m[1], value: m[2] });
        else    textParts.push(tok);
    }
    return { conditions, text: textParts.join(' ') };
}

function matchesQuery(name, info, label, stid, { conditions, text }) {
    // Structured conditions (all must match)
    for (const { key, value } of conditions) {
        if (key === 'stageid') {
            const sid = stid ? info.stage_ids?.[stid] : undefined;
            if (sid === undefined || String(sid) !== value) return false;
        } else if (key === 'stageno') {
            if (!stid) return false;
            // Accept "100" or "0100" — compare numerically
            if (parseInt(stid.slice(2), 10) !== parseInt(value, 10)) return false;
        } else if (key === 'area') {
            const aname = (info.quest_area_name ?? '').toLowerCase();
            if (!aname.includes(value)) return false;
        }
    }
    // Free-text substring (against name, label, stid)
    if (text && !name.includes(text) && !label.toLowerCase().includes(text) && !(stid && stid.includes(text))) return false;
    return true;
}

function buildSidebar(filter = '') {
    const listEl = document.getElementById('map-list');
    listEl.innerHTML = '';
    const currentMap = currentMapName();
    const currentStage = currentStageName();
    const query = parseSearchQuery(filter);
    const hasFilter = query.conditions.length > 0 || query.text.length > 0;

    if (hasFilter && query.text) {
        appendNamedLocationSearchResults(listEl, filter);
    }

    // Build one entry per (name, stid) pair — stid suffix intentionally omitted from label.
    // Skip pd piece models (pd###_m##) — internal tileset pieces, not navigable locations.
    const pdPieceRe = /^pd\d+_m\d+$/;
    const entries = [];
    for (const [name, info] of Object.entries(mapParams)) {
        if (pdPieceRe.test(name)) continue;
        const stages = info.stages?.length ? info.stages : [null];
        for (const stid of stages) {
            // Label never includes the stid suffix — multi-stage maps are collapsed below
            const label = stid ? stageLabel(info, stid) : displayName(name, info);
            if (hasFilter && !matchesQuery(name, info, label, stid, query)) continue;
            entries.push({ name, info, label, stid });
        }
    }

    if (hasFilter) {
        entries.sort((a, b) => a.label.localeCompare(b.label));
        if (listEl.querySelector('.map-loc-section')) {
            const mapsHeader = document.createElement('div');
            mapsHeader.className = 'map-loc-header map-list-maps-header';
            mapsHeader.textContent = 'Maps';
            listEl.appendChild(mapsHeader);
        }
        renderMapEntries(listEl, entries, currentMap, currentStage);
        return;
    }

    // No search: group by quest area (sorted by quest_area_id), collapsible sections
    const areaMap = new Map(); // area_id -> { name, entries[] }
    for (const e of entries) {
        let aid   = e.info.quest_area_id  ?? 0;
        let aname = e.info.quest_area_name ?? 'Unknown';
        if (e.label.toLowerCase().includes('bitterblack')) {
            aid   = 24;
            aname = 'Bitterblack Maze';
        }
        if (!areaMap.has(aid)) areaMap.set(aid, { name: aname, entries: [] });
        areaMap.get(aid).entries.push(e);
    }

    const sortedAreas = [...areaMap.entries()].sort(([a], [b]) => {
        if (a === 0) return 1;
        if (b === 0) return -1;
        return a - b;
    });

    const areaById = new Map(sortedAreas);

    const openState = loadAreaOpenState();
    for (const group of SIDEBAR_WORLD_GROUPS) {
        const subAreas = group.areas
            .map(id => {
                const area = areaById.get(id);
                return area ? [id, area] : null;
            })
            .filter(Boolean);
        if (!subAreas.length) continue;

        if (group.flat && subAreas.length === 1) {
            const [areaId, area] = subAreas[0];
            appendAreaSection(listEl, areaId, group.displayName ?? area.name, area.entries,
                currentMap, currentStage, openState);
        } else {
            appendWorldRegionSection(listEl, group.key, group.name, subAreas,
                currentMap, currentStage, openState, { defaultCollapsed: !!group.defaultCollapsed });
        }
    }
    for (const [areaId, area] of sortedAreas) {
        if (SIDEBAR_GROUPED_AREA_IDS.has(areaId)) continue;
        appendAreaSection(listEl, areaId, area.name, area.entries,
            currentMap, currentStage, openState);
    }
}

function debounce(fn, ms) {
    let timer = null;
    return (...args) => {
        clearTimeout(timer);
        timer = setTimeout(() => fn(...args), ms);
    };
}

const _mapSearchInput = document.getElementById('map-search');
const _mapSearchClear = document.getElementById('map-search-clear');
const _buildSidebarDebounced = debounce(value => buildSidebar(value), 150);
_mapSearchInput.addEventListener('input', e => {
    _mapSearchClear.style.display = e.target.value ? 'block' : 'none';
    _buildSidebarDebounced(e.target.value);
});
_mapSearchClear.addEventListener('click', () => {
    _mapSearchInput.value = '';
    _mapSearchClear.style.display = 'none';
    _mapSearchInput.focus();
    buildSidebar('');
});

initMajorLocSelectors();

// ── URL hash navigation ────────────────────────────────────────────────────────
// Hash format: #mapname  or  #mapname:stid  or either suffixed with @zoom/y/x
// e.g. #rm000_m02:st0301@2.50/1024.0/800.0
function parseHash() {
    const raw = window.location.hash.slice(1);
    // Split off optional !layers suffix before everything else
    const [beforeLayers, layersPart = null] = raw.split('!');
    const [nameStid, viewPart = null] = beforeLayers.split('@');
    const [name, stid = null] = nameStid.split(':');
    let view = null;
    if (viewPart) {
        const [z, y, x, f] = viewPart.split('/').map(Number);
        if (!isNaN(z) && !isNaN(y) && !isNaN(x))
            view = { zoom: z, center: L.latLng(y, x), floor: !isNaN(f) ? f : null };
    }
    let layers = null, openGroups = null;
    if (layersPart !== null) {
        const [flagStr, groupsStr = ''] = layersPart.split(';');
        layers = {
            enemies:      flagStr.includes('e'),
            landmarks:    flagStr.includes('l'),
            connections:  flagStr.includes('c'),
            grid:         flagStr.includes('g'),
            territory:    flagStr.includes('t'),
            stageLabels:  flagStr.includes('a'),
            gather:       flagStr.includes('r'),
            radii:         flagStr.includes('i'),
            shops:         flagStr.includes('n') || flagStr.includes('p'),
            sidebarHidden: flagStr.includes('s'),
        };
        openGroups = groupsStr ? groupsStr.split(',').filter(Boolean) : [];
    }
    return { name, stid, view, layers, openGroups };
}

function currentMapName() {
    const { name } = parseHash();
    return (name && mapParams[name]) ? name : 'field000_m00';
}

function currentStageName() {
    return parseHash().stid;
}

function navigateTo(mapName, stid = null, view = null) {
    let hash = stid ? `${mapName}:${stid}` : mapName;
    if (view) {
        hash += `@${view.zoom.toFixed(2)}/${view.center.lat.toFixed(1)}/${view.center.lng.toFixed(1)}`;
    }
    window.location.hash = hash;
}

/** Ensure hash names a valid map + stage so the first load is never blank. */
function ensureInitialHash() {
    const raw = window.location.hash.slice(1);
    const parsed = parseHash();
    const mapName = (parsed.name && mapParams[parsed.name]) ? parsed.name : 'field000_m00';
    const info = mapParams[mapName];
    const stid = (parsed.stid && info?.stages?.includes(parsed.stid))
        ? parsed.stid
        : defaultFieldStage(mapName);

    const targetBase = stid ? `${mapName}:${stid}` : mapName;
    const [beforeLayers = ''] = raw.split('!');
    const [basePart = ''] = beforeLayers.split('@');
    if (!raw || basePart !== targetBase) {
        let hash = targetBase;
        if (parsed.view) {
            hash += `@${parsed.view.zoom.toFixed(2)}/${parsed.view.center.lat.toFixed(1)}/${parsed.view.center.lng.toFixed(1)}`;
            if (parsed.view.floor != null) hash += `/${parsed.view.floor}`;
        }
        if (raw.includes('!')) hash += raw.slice(raw.indexOf('!'));
        history.replaceState(null, '', `#${hash}`);
    }
}

// Track last-loaded map+stid so hashchange can skip reloads on view-only updates.
let _loadedMapName = null;
let _loadedStid = null;

window.addEventListener('hashchange', () => {
    const newMap  = currentMapName();
    const newStid = currentStageName();
    if (newMap !== _loadedMapName || newStid !== _loadedStid) {
        loadMap(newMap);
        buildSidebar(document.getElementById('map-search').value);
    }
    syncMajorLocSelectors();
});

// Persist zoom+pan in the hash via replaceState (no extra history entries).
let _viewUpdateTimer = null;
leafletMap.on('moveend zoomend', () => {
    clearTimeout(_viewUpdateTimer);
    _viewUpdateTimer = setTimeout(() => {
        const { name, stid } = parseHash();
        const mapName = name || _loadedMapName;
        if (!mapName) return;
        const z = leafletMap.getZoom().toFixed(2);
        const c = leafletMap.getCenter();
        const frag = (stid ? `${mapName}:${stid}` : mapName)
                   + `@${z}/${c.lat.toFixed(1)}/${c.lng.toFixed(1)}`
                   + `!${getLayersHash()}`;
        history.replaceState(null, '', '#' + frag);
    }, 200);
});

// ── Overlapping marker spread ──────────────────────────────────────────────────
// When multiple markers share the exact same pixel position they stack invisibly.
// This detects such groups and fans them out into a small ring.
// Spokes and anchor dots are written to `overlayLayer` so they can be cleared
// independently (cross-group recompute on expand/collapse).
const OVERLAP_SPREAD_R = 9;   // ring radius in world-pixels (visible at zoom ≥ 0)

// Reset a spread marker back to its natural position and style.
function _resetMarkerSpread(m) {
    if (m._naturalLatLng) m.setLatLng(m._naturalLatLng);
    if (m._origStyle)     m.setStyle(m._origStyle);
    if (m._naturalTooltip) m.bindTooltip(m._naturalTooltip, { direction: 'top', offset: [0, -8] });
}

function _doSpread(markers, overlayLayer) {
    const byPos = new Map();
    for (const m of markers) {
        const ll  = m.getLatLng();
        const key = `${ll.lat.toFixed(2)}:${ll.lng.toFixed(2)}`;
        if (!byPos.has(key)) byPos.set(key, []);
        byPos.get(key).push(m);
    }

    for (const group of byPos.values()) {
        if (group.length < 2) continue;
        const origin = group[0].getLatLng();
        const N = group.length;

        const stackLines = group.map(m => {
            const c = m.options.color;
            return `<span style="color:${c};font-weight:bold">&#x25CF;</span> ${m._label ?? '?'}`;
        });
        // Create anchor first so spread markers can reference it in their hover handlers.
        const anchor = L.circleMarker(origin, {
            radius: 4, color: '#fff', fillColor: '#fff',
            fillOpacity: 0.85, weight: 1.5, opacity: 1.0, className: 'enemy-marker',
        })
            .bindTooltip(`<b>${N} stacked here:</b><br>${stackLines.join('<br>')}`,
                         { direction: 'top', offset: [0, -8] })
            .addTo(overlayLayer);
        anchor.on('mouseover', function() { _applyHighlight(group); });
        anchor.on('mouseout',  _unhighlightSG);

        group.forEach((m, i) => {
            const angle = (i / N) * 2 * Math.PI - Math.PI / 2;
            m.setLatLng(L.latLng(
                origin.lat + OVERLAP_SPREAD_R * Math.sin(angle),
                origin.lng + OVERLAP_SPREAD_R * Math.cos(angle),
            ));
            m.setStyle({ dashArray: '4 3' });
            m._origStyle = { ...m._origStyle, dashArray: '4 3' };
            m._spreadCount = N;
            m.bindTooltip(
                `${m._naturalTooltip} <span style="opacity:0.7">[×${N} stacked]</span>`,
                { direction: 'top', offset: [0, -8] },
            );
            // Store anchor ref so SG-highlight can also enlarge it (task 1).
            m._spreadAnchor = anchor;
            // Hovering a spread node enlarges the anchor so the real position is easy to find.
            m.on('mouseover', () => { anchor.setRadius(8); anchor.setStyle({ weight: 2.5, fillOpacity: 1.0 }); });
            m.on('mouseout',  () => { anchor.setRadius(4); anchor.setStyle({ weight: 1.5, fillOpacity: 0.85 }); });
        });

        // Spokes: thin dashed lines that brighten when their spread node or the anchor is hovered.
        group.forEach(m => {
            const spoke = L.polyline([m.getLatLng(), origin], {
                color: m.options.color, weight: 1, opacity: 0.4,
                dashArray: '3 3', interactive: false,
            }).addTo(overlayLayer);
            m._spokeLine = spoke;

            m.on('mouseover', () => spoke.setStyle({ weight: 2.5, opacity: 1.0, dashArray: null }));
            m.on('mouseout',  () => spoke.setStyle({ weight: 1,   opacity: 0.4,  dashArray: '3 3' }));
        });

        // Hovering the anchor highlights all spokes in the group.
        anchor.on('mouseover', () => {
            for (const m of group) if (m._spokeLine)
                m._spokeLine.setStyle({ weight: 2.5, opacity: 1.0, dashArray: null });
        });
        anchor.on('mouseout', () => {
            for (const m of group) if (m._spokeLine)
                m._spokeLine.setStyle({ weight: 1, opacity: 0.4, dashArray: '3 3' });
        });
    }
}

// Recompute cross-group spread across all currently-expanded groups.
// Called after any expand or collapse so collisions between groups are always current.
function reapplySpread() {
    _spreadOverlay.clearLayers();
    const allMarkers = [];
    for (const g of _groupStore.values()) {
        if (!g.isExpanded || !g.detailsLayer) continue;
        for (const m of g.detailsLayer.getLayers()) {
            if (!m._spawn || m._hidden) continue;
            _resetMarkerSpread(m);
            allMarkers.push(m);
        }
    }
    _doSpread(allMarkers, _spreadOverlay);
}

// ── Group hull helpers ─────────────────────────────────────────────────────────
// Andrew's monotone chain — returns the convex hull of pts as [x,y] pairs.
function convexHull(pts) {
    if (pts.length < 3) return pts.slice();
    const s = [...pts].sort((a, b) => a[0] !== b[0] ? a[0] - b[0] : a[1] - b[1]);
    const cross = (o, a, b) => (a[0]-o[0])*(b[1]-o[1]) - (a[1]-o[1])*(b[0]-o[0]);
    const lo = [], hi = [];
    for (const p of s) {
        while (lo.length >= 2 && cross(lo.at(-2), lo.at(-1), p) <= 0) lo.pop();
        lo.push(p);
    }
    for (let i = s.length - 1; i >= 0; i--) {
        const p = s[i];
        while (hi.length >= 2 && cross(hi.at(-2), hi.at(-1), p) <= 0) hi.pop();
        hi.push(p);
    }
    hi.pop(); lo.pop();
    return [...lo, ...hi];
}

// Ray-casting point-in-polygon test.  pts: [[x,y], ...] (Leaflet [lng,lat] pairs).
function pointInPolygon(px, py, pts) {
    let inside = false;
    for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) {
        const [xi, yi] = pts[i], [xj, yj] = pts[j];
        if ((yi > py) !== (yj > py) && px < (xj - xi) * (py - yi) / (yj - yi) + xi)
            inside = !inside;
    }
    return inside;
}

// Middle-click anywhere inside an expanded group's hull collapses it.
// Hull polygons are non-interactive (pointer-events:none) so this map-level
// handler is the only way to catch middle-clicks in the hull area.
leafletMap.on('mousedown', (e) => {
    if (e.originalEvent.button !== 1) return;
    e.originalEvent.preventDefault();
    for (const g of _groupStore.values()) {
        if (!g.isExpanded || !g.hullPts || g.hullPts.length < 3) continue;
        if (pointInPolygon(e.latlng.lng, e.latlng.lat, g.hullPts)) {
            collapseGroup(g.groupId);
            return;
        }
    }
});

// ── Group chip / expand-collapse helpers ──────────────────────────────────────

function makeChipIcon(groupId, _color, count, expanded, yOffset = 10, isKeyBearerGroup = false, isBossGroup = false) {
    // Use a brighter variant of the same hue for chip text (dark chip background needs L~0.78).
    const chipColor = `oklch(0.78 0.13 ${(parseInt(groupId, 10) * 137) % 360})`;
    // Boss: red glow. Key bearer: gold glow + icon on the left. Both: layered glows.
    const glows = [];
    if (isBossGroup)      glows.push('0 0 7px 2px rgba(255,60,60,0.85)');
    if (isKeyBearerGroup) glows.push('0 0 7px 2px rgba(255,210,0,0.85)');
    const shadowStyle = glows.length ? `box-shadow:0 0 4px rgba(0,0,0,0.7),${glows.join(',')};` : '';
    const titleAttr = [isBossGroup ? 'Contains boss enemy' : '', isKeyBearerGroup ? 'Key bearer group' : ''].filter(Boolean).join(' · ');
    const groupPart = useMobSpawnDevLabels() ? `G${groupId} ` : '';
    const countPart = useMobSpawnDevLabels() ? `<span class="chip-count">${count}</span>` : '';
    return L.divIcon({
        className: '',
        html: `<div class="group-chip${expanded ? ' chip-open' : ''}" style="color:${chipColor};${shadowStyle}"${titleAttr ? ` title="${titleAttr}"` : ''}><span class="chip-arrow${expanded ? ' open' : ''}">&#9654;</span>${groupPart}${countPart}</div>`,
        iconSize:   null,
        // When expanded: anchor at bottom of chip so the chip floats above the marker position.
        // When collapsed: anchor near top (yOffset) so chip hangs below the centroid.
        iconAnchor: expanded ? [0, 22] : [0, yOffset],
    });
}

// Returns true if any spawn position in the group has a boss-type enemy in the cache.
function _groupHasBoss(g) {
    if (!_enemySpawnCache) return false;
    for (const { spawn, idx, stageNo } of g.items) {
        const sid = stageIds[stageNo];
        if (sid == null) continue;
        const key = `${sid},${g.groupId},${spawn.posIdx ?? idx}`;
        const entries = _enemySpawnCache.get(key) ?? [];
        if (entries.some(e => e.isBossGauge || e.isAreaBoss || (e.raidBossId > 0))) return true;
    }
    return false;
}

const _SPAWN_INFECTION_PREFIX = [null, 'Infected', 'Severely Infected', 'War-Ready'];

// Known SpawnTime windows (Lestania clock, "start,end" — night wraps midnight).
const SPAWN_TIME_ALWAYS = '00:00,23:59';
const SPAWN_TIME_DAY    = new Set(['07:00,17:59', '06:00,17:59']);   // Arrowgene / Rising·Revival
const SPAWN_TIME_NIGHT  = new Set(['18:00,06:59', '18:00,05:59']);   // Arrowgene / Rising·Revival

const spawnTimeKind = (t) => {
    if (!t || t === SPAWN_TIME_ALWAYS) return 'always';
    if (SPAWN_TIME_DAY.has(t))   return 'day';
    if (SPAWN_TIME_NIGHT.has(t)) return 'night';
    const [start, end] = t.split(',');
    if (!end) return 'other';
    // Heuristic for minor variants of the standard day/night windows
    if ((start.startsWith('06:') || start.startsWith('07:')) && end.startsWith('17:')) return 'day';
    if (start.startsWith('18:') && (end.startsWith('05:') || end.startsWith('06:'))) return 'night';
    return 'other';
};

const spawnTimeLabel = (t) => {
    const kind = spawnTimeKind(t);
    if (kind === 'day')    return '☀ Day';
    if (kind === 'night')  return '🌙 Night';
    if (kind === 'always') return '';
    return t;
};

const filterEntriesBySpawnTime = (entries, filter = _activeSpawnTimeFilter) => {
    if (!filter) return entries;
    return entries.filter(e => {
        const kind = spawnTimeKind(e.spawnTime);
        if (kind === 'always') return true;
        if (kind === 'other') return false;
        return kind === filter;
    });
};

// Layout slot with no row in EnemySpawn.json — quest/event/dynamic spawn.
const isDynamicSpawnSlot = (spawnCache, spawnKey) => {
    if (!_enemySpawnDataLoaded || spawnKey == null || !spawnCache) return false;
    const allEntries = spawnCache.get(spawnKey) ?? [];
    return !allEntries.some(e => !!e.lv);
};

const spawnMobTags = (entries, spawn, spawnCache, spawnKey) => {
    const tags = new Set();
    if (isDynamicSpawnSlot(spawnCache, spawnKey)) {
        tags.add('dynamic');
        return tags;
    }
    if (spawn?.KeyBearer === true) tags.add('keyBearer');
    for (const e of entries) {
        if (!e.lv) continue;
        if (e.isBloodOrbEnemy && e.bloodOrbs) tags.add('bloodOrb');
        if (e.isHighOrbEnemy && e.highOrbs)  tags.add('highOrb');
        if (e.isManualSet) tags.add('manual');
        if (e.isBossGauge || e.isAreaBoss || e.isBossBGM || e.raidBossId > 0) tags.add('boss');
    }
    if (entries.some(e => !!e.lv) && !tags.size) tags.add('regular');
    return tags;
};

const spawnMatchesTypeFilter = (cache, spawnKey, spawn) => {
    const enabledIds = MOB_TYPE_MODES.filter(m => _mobTypeFilters[m.id]).map(m => m.id);
    if (!enabledIds.length) return false;

    if (!_enemySpawnDataLoaded) {
        if (spawn?.KeyBearer && !_mobTypeFilters.keyBearer) return false;
        return _mobTypeFilters.regular;
    }

    const allEntries = spawnKey && cache ? (cache.get(spawnKey) ?? []) : [];
    const entries    = filterEntriesBySpawnTime(allEntries);
    const tags = spawnMobTags(entries, spawn, cache, spawnKey);
    if (!tags.size) return false;
    return [...tags].some(t => _mobTypeFilters[t]);
};

const resolveDisplayNameFromEntry = (e) => {
    const baseName = e.emCode ? (emNames[e.emCode]?.name ?? e.emCode) : null;
    if (!baseName) return null;
    const np = e.namedId ? namedParamsById.get(e.namedId) : null;
    const npName = np?.name?.trim();
    let name;
    if (!npName || np.type === 'NAMED_TYPE_NONE') name = baseName;
    else if (np.type === 'NAMED_TYPE_REPLACE') name = npName;
    else if (np.type === 'NAMED_TYPE_PREFIX')  name = `${npName} ${baseName}`;
    else if (np.type === 'NAMED_TYPE_SUFFIX')  name = `${baseName} ${npName}`;
    else name = baseName;
    const prefix = e.infection ? _SPAWN_INFECTION_PREFIX[e.infection] : null;
    return prefix ? `${prefix} ${name}` : name;
};

const replaceOriginSuffix = (e) => {
    if (!e?.namedId) return '';
    const np = namedParamsById.get(e.namedId);
    if (np?.type !== 'NAMED_TYPE_REPLACE') return '';
    const base = e.emCode ? (emNames[e.emCode]?.name ?? null) : null;
    if (!base) return '';
    return ` <span style="font-size:10px;color:#aaa;font-style:italic">(${base})</span>`;
};

const resolveSpawnMobLabelFromEntries = (entries, spawn, spawnCache) => {
    const hasEnemy = !spawnCache || entries.some(e => !!e.lv);

    if (entries.length > 1) {
        const parts = entries
            .filter(e => !!e.lv)
            .map(e => {
                const n = resolveDisplayNameFromEntry(e);
                const t = spawnTimeLabel(e.spawnTime);
                return n ? `${n} Lv${e.lv}${t ? [...t][0] : ''}${replaceOriginSuffix(e)}` : null;
            })
            .filter(Boolean);
        if (parts.length) return parts.join(' / ');
    } else if (entries.length === 1 && hasEnemy) {
        const e0 = entries[0];
        const n  = resolveDisplayNameFromEntry(e0) ?? (spawn?.EmName ? (emNames[spawn.EmName]?.name ?? null) : null);
        if (n) return `${n}${e0.lv ? ` Lv${e0.lv}` : ''}${replaceOriginSuffix(e0)}`;
    } else if (!spawnCache && spawn?.EmName) {
        const n = emNames[spawn.EmName]?.name ?? null;
        if (n) return n;
    }
    return null;
};

// Returns a player-facing mob label, or null when the slot should stay hidden.
const resolveSpawnMobLabel = (spawnCache, spawnKey, spawn) => {
    const allEntries = spawnKey && spawnCache ? (spawnCache.get(spawnKey) ?? []) : [];
    if (isDynamicSpawnSlot(spawnCache, spawnKey)) return DYNAMIC_ENEMY_LABEL;
    const entries = filterEntriesBySpawnTime(allEntries);
    if (_activeSpawnTimeFilter && allEntries.length && !entries.length) return null;
    return resolveSpawnMobLabelFromEntries(entries, spawn, spawnCache);
};

const spawnHasMobLabel = (spawnCache, spawnKey, spawn) => {
    if (!_enemySpawnDataLoaded) {
        if (_activeSpawnTimeFilter) return false;
        return !!(spawn?.EmName && emNames[spawn.EmName]?.name);
    }
    return resolveSpawnMobLabel(spawnCache, spawnKey, spawn) != null;
};

// Add convex hull / link polyline for a group (legacy group-chip mode only).
function addGroupStructureLayers(g, layer) {
    g.hullPts = null;
    if (g.pts.length >= 3) {
        const hull = convexHull(g.pts);
        if (hull.length >= 3) {
            const poly = L.polygon(hull.map(([px, py]) => xy(px, py)), {
                color:       g.color,
                weight:      1.5,
                opacity:     0.75,
                fillColor:   g.color,
                fillOpacity: 0.10,
                dashArray:   '6 4',
                interactive: false,  // pointer-events:none so canvas markers stay clickable
            });
            poly._isGroupStructure = true;
            layer.addLayer(poly);
            g.hullPts = hull;  // stored for map-level middle-click collapse
        }
    } else if (g.pts.length === 2) {
        const line = L.polyline(g.pts.map(([px, py]) => xy(px, py)), {
            color: g.color, weight: 1.5, opacity: 0.65, dashArray: '4 3', interactive: false,
        });
        line._isGroupStructure = true;
        layer.addLayer(line);
    }
}

const groupHasStructureLayers = (g) =>
    g.detailsLayer?.getLayers().some(l => l._isGroupStructure) ?? false;

// Lazy-build hulls when legacy chips are enabled (e.g. toggled on after initial expand).
function ensureGroupStructureLayers(g) {
    if (!useLegacyGroupChips() || !g.detailsLayer || groupHasStructureLayers(g)) return;
    addGroupStructureLayers(g, g.detailsLayer);
}

function removeGroupStructureLayers(g) {
    if (!g.detailsLayer) return;
    for (const layer of [...g.detailsLayer.getLayers()]) {
        if (layer._isGroupStructure) g.detailsLayer.removeLayer(layer);
    }
    g.hullPts = null;
}

function syncGroupStructureVisibility(g, groupVisible) {
    if (!g.detailsLayer) return;
    const show = groupVisible && useLegacyGroupChips();
    for (const layer of g.detailsLayer.getLayers()) {
        if (!layer._isGroupStructure) continue;
        if (show) {
            layer.setStyle(layer._origStyle ?? {});
        } else {
            if (!layer._origStyle) {
                layer._origStyle = {
                    opacity: layer.options.opacity ?? 0.75,
                    fillOpacity: layer.options.fillOpacity ?? 0.1,
                };
            }
            layer.setStyle({ opacity: 0, fillOpacity: 0 });
        }
    }
}

// Build the details layer (hull + territory + spawn dots) for a group entry.
// Does NOT add the layer to the map — that is done by expandGroup.
function buildGroupDetails(g) {
    const info  = _currentMapInfo;
    const layer = L.layerGroup();

    if (useLegacyGroupChips()) addGroupStructureLayers(g, layer);

    g.territoryRect = null;

    // Spawn circleMarkers
    g.sgMarkers = {};
    for (const { spawn, idx, sg, latlng, stageNo } of g.items) {
        const fillColor = spawnGroupColor(sg);
        const sgKey     = `${sg}:${g.groupId}`;

        const badge = `<span style="display:inline-block;padding:1px 6px;border-radius:3px;background:${fillColor};color:#111;font-weight:bold;font-size:11px;">Spawn Set: ${sg}</span>`;
        const subLine = spawn.SubGroupNo != null ? (() => {
            const subId    = spawn.SubGroupNo + 1;
            const subColor = spawnGroupColor(spawn.SubGroupNo);
            const subBadge = `<span style="display:inline-block;padding:1px 6px;border-radius:3px;background:${subColor};color:#111;font-weight:bold;font-size:11px;">${spawn.SubGroupNo}</span>`;
            return `<br><span style="font-size:11px;color:#aaa">SubGroup ${subId} on enemy event (lot SubGroupNo=${subBadge})</span>`;
        })() : '';
        const triggerLine = g.areaSpawn ? (() => {
            if (g.priority > 0)
                return `<br><span style="font-size:11px;color:#8cf">&#9889; SubGroup 1: player enters zone</span>`;
            else
                return `<br><span style="font-size:11px;color:#fa8">&#9889; SubGroup 1: condition trigger (fight event / kill)</span>`;
        })() : '';
        const groupLabel = `<span style="color:${g.color};font-weight:bold;">Group: ${g.groupId}</span>`;
        const isKeyBearer = spawn.KeyBearer === true;
        const keyLine = isKeyBearer ? '<br><span style="font-size:11px">🗝 Key Bearer</span>' : '';
        // ManualSet is a server-side spawn property — only shown when server data is loaded
        const buildManualSetLine = (spawnInfo) =>
            spawnInfo?.isManualSet
                ? `<br><span style="font-size:11px;color:#b0c4ff" title="Spawns dormant (mIsWaitting=true). Activated by boss SummonSet FSM action.">&#128564; Dormant until summoned</span>`
                : '';
        const buildBossLine = (spawnInfo) => {
            if (!spawnInfo) return '';
            const parts = [];
            if (spawnInfo.isBossGauge) parts.push('Boss Gauge');
            if (spawnInfo.isBossBGM)   parts.push('Boss BGM');
            if (spawnInfo.isAreaBoss)  parts.push('Area Boss');
            if (spawnInfo.raidBossId > 0) parts.push(`Raid Boss ID: ${spawnInfo.raidBossId}`);
            return parts.length ? `<br><span style="font-size:11px;color:#ff6666">☠ ${parts.join(' · ')}</span>` : '';
        };

        // Build popup HTML optionally enriched with server EnemySpawn data
        const serverStageId = stageIds[stageNo];
        const spawnKey = serverStageId != null
            ? `${serverStageId},${g.groupId},${spawn.posIdx ?? idx}` : null;

        // Per-marker display index for cycling through day/night/always variants
        let displayIdx = 0;
        // Tracks whether the user dismissed the mixed-content conflict notice for this marker
        let _spawnSetConflictDismissed = false;

        // Returns all OTHER positions in the same group that share the same SpawnGroup (sg) value.
        // Each result: { key, entries, idx }
        const getPeers = (cache) => {
            if (!spawnKey || !cache) return [];
            const [sid, gid] = spawnKey.split(',');
            const myPosIdx = String(spawn.posIdx ?? idx);
            return g.items
                .filter(item => item.sg === sg && String(item.spawn.posIdx ?? item.idx) !== myPosIdx)
                .flatMap(item => {
                    const pIdx = item.spawn.posIdx ?? item.idx;
                    const key  = `${sid},${gid},${pIdx}`;
                    const entries = cache.get(key) ?? [];
                    return [{ key, entries, idx: item.idx }];
                });
        };

        const buildDropsHtml = (spawnInfo) => {
            if (!spawnInfo?.drops?.length) return '';
            return '<br><table style="font-size:13px;margin-top:6px;border-collapse:collapse;line-height:1.8">' +
                spawnInfo.drops.map(row => {
                    // row = [itemId, minQty, maxQty, unknown, isHidden, dropRate]
                    const itemId   = row[0];
                    const minQty   = row[1] ?? 1;
                    const maxQty   = row[2] ?? 1;
                    const dropRate = row[5];
                    const entry    = itemNames[String(itemId)];
                    const name     = entry?.name ?? `Item #${itemId}`;
                    const iconNo   = entry?.iconNo;
                    const iconFile = iconNo != null ? `ii${String(iconNo).padStart(6, '0')}.png` : null;
                    const icon     = iconFile && _iconIdSet.has(iconNo)
                        ? `<img src="images/icons/small/${iconFile}" width="28" height="28" style="vertical-align:middle;margin-right:6px;image-rendering:pixelated">`
                        : `<span style="display:inline-block;width:28px;margin-right:6px"></span>`;
                    const href     = itemWikiHref(itemId);
                    const nameLink = `<a href="${href}" target="_blank" style="color:inherit;text-decoration:none" onmouseover="this.style.textDecoration='underline'" onmouseout="this.style.textDecoration='none'">${name}</a>`;
                    const qty      = maxQty > minQty ? ` ×${minQty}–${maxQty}` : ` ×${minQty}`;
                    const pct      = dropRate > 0
                        ? ` <span style="color:#777">(${(dropRate * 100).toFixed(0)}%)</span>` : '';
                    return `<tr><td style="color:#222;padding-right:8px">${icon}${nameLink}</td><td style="color:#555;white-space:nowrap;vertical-align:top;padding-top:4px">${qty}${pct}</td></tr>`;
                }).join('') +
                '</table>';
        };

        const buildEnemyPopup = (spawnCache) => {
            const entries   = spawnKey && spawnCache ? (spawnCache.get(spawnKey) ?? []) : [];
            const spawnInfo = entries[displayIdx] ?? null;
            // An enemy is "present" if: server data isn't loaded yet (unknown),
            // or any entry has a non-zero level.
            const hasEnemy  = !spawnCache || entries.some(e => !!e.lv);
            // Prefer em code from server data; fall back to lot file only while loading
            const emCode    = spawnInfo?.emCode ?? (hasEnemy ? spawn.EmName : null);
            const emEntry   = emCode ? emNames[emCode] : null;
            const dispName  = emEntry?.name ?? null;
            const lvText    = spawnInfo?.lv ? ` Lv${spawnInfo.lv}` : '';
            // Named param prefix/suffix/replace display
            const namedParam = spawnInfo?.namedId ? namedParamsById.get(spawnInfo.namedId) : null;
            const namedTrimmed = namedParam?.name?.trim();
            const namedDisplayName = namedTrimmed && namedParam.type !== 'NAMED_TYPE_NONE'
                ? namedTrimmed : null;
            const namedFull = namedDisplayName ? (() => {
                if (namedParam.type === 'NAMED_TYPE_REPLACE') return namedDisplayName;
                if (namedParam.type === 'NAMED_TYPE_PREFIX') return dispName ? `${namedDisplayName} ${dispName}` : namedDisplayName;
                if (namedParam.type === 'NAMED_TYPE_SUFFIX') return dispName ? `${dispName} ${namedDisplayName}` : namedDisplayName;
                return null;
            })() : null;
            const _infectionPrefix = [null, 'Infected', 'Severely Infected', 'War-Ready'];
            const _infPrefix = spawnInfo?.infection ? _infectionPrefix[spawnInfo.infection] : null;
            const shownName = _infPrefix
                ? `${_infPrefix} ${namedFull ?? dispName ?? ''}`
                : (namedFull ?? dispName);
            const isReplaced = namedParam?.type === 'NAMED_TYPE_REPLACE' && dispName;
            const emCodeLine = shownName && emCode
                ? `<span style="color:#888;font-size:10px"> (${isReplaced ? `${dispName} · ` : ''}${emCode})</span>` : '';
            const emLine = shownName
                ? `<br><span class="se-enemy-name" style="color:#333;font-size:12px">${shownName}${lvText}</span>${emCodeLine}` : '';

            // Cycle controls when multiple day/night variants exist
            const cycleHtml = entries.length > 1 ? (() => {
                const btnStyle = 'background:#e8e8e8;border:1px solid #bbb;border-radius:3px;padding:0 5px;cursor:pointer;font-size:11px;line-height:16px;';
                const timeLabel = spawnTimeLabel(spawnInfo?.spawnTime);
                const timePart  = timeLabel ? ` &nbsp;${timeLabel}&nbsp; ` : ` &nbsp;`;
                return `<br><span style="font-size:11px;color:#444">` +
                    `<button class="spawn-prev" style="${btnStyle}">◀</button>` +
                    `${timePart}<span style="color:#666">${displayIdx + 1}/${entries.length}</span>&nbsp;` +
                    `<button class="spawn-next" style="${btnStyle}">▶</button></span>`;
            })() : '';

            const radiiLine = hasEnemy ? (() => {
                const _ec    = spawnInfo?.emCode ?? spawn.EmName ?? null;
                const _radii = _ec ? (emRadii[_ec] ?? null) : null;
                const aggroR = _radii?.aggroRadius ?? spawn.AggroRadius;
                const linkR  = _radii?.linkRadius  ?? spawn.LinkRadius;
                if (!aggroR && !linkR) return '';
                const ag = aggroR ? `<span style="color:#ffd700">&#9679;</span> Aggro: ${aggroR}` : '';
                const lk = linkR  ? `<span style="color:#ff7700">&#9675;</span> Link: ${linkR}`  : '';
                return `<br><span style="font-size:11px">${[ag, lk].filter(Boolean).join(' &nbsp; ')}</span>`;
            })() : '';
            const orbsLine = hasEnemy && ((spawnInfo?.isBloodOrbEnemy && spawnInfo?.bloodOrbs) || (spawnInfo?.isHighOrbEnemy && spawnInfo?.highOrbs)) ? (() => {
                const b = (spawnInfo.isBloodOrbEnemy && spawnInfo.bloodOrbs) ? `<span title="Blood Orbs">🩸</span> ${spawnInfo.bloodOrbs}` : '';
                const h = (spawnInfo.isHighOrbEnemy  && spawnInfo.highOrbs)  ? `<span title="High Orbs">⭐</span> ${spawnInfo.highOrbs}`   : '';
                return `<br><span style="font-size:12px">${[b, h].filter(Boolean).join(' &nbsp;&nbsp; ')}</span>`;
            })() : '';
            // ── Spawn-set prev/next nav (visible in both viewer and edit mode) ──
            const buildSetNavRow = () => {
                if (!spawnKey || sg == null) return '';
                const [sid, gid] = spawnKey.split(',');
                const allPos = g.items
                    .filter(item => item.sg === sg)
                    .map(item => ({
                        key: `${sid},${gid},${item.spawn.posIdx ?? item.idx}`,
                        idx: item.idx,
                    }))
                    .sort((a, b) => a.idx - b.idx);
                if (allPos.length < 2) return '';
                const curIdx  = allPos.findIndex(p => p.key === spawnKey);
                const nb      = 'font-size:10px;padding:1px 6px;border-radius:3px;cursor:pointer;border:1px solid #ccc;background:#f0f0f0;color:#555';
                const prevPos = allPos[(curIdx - 1 + allPos.length) % allPos.length];
                const nextPos = allPos[(curIdx + 1) % allPos.length];
                const posLabel = curIdx >= 0 ? `${curIdx + 1}/${allPos.length}` : `?/${allPos.length}`;
                return `<div style="display:flex;gap:4px;align-items:center;margin-bottom:4px">`
                    + `<span style="font-size:9px;color:#aaa;text-transform:uppercase;letter-spacing:0.4px;margin-right:2px">Set ${sg}</span>`
                    + `<button class="se-set-nav-btn" data-key="${prevPos.key}" style="${nb}" title="Previous in spawn set">◀</button>`
                    + `<span style="font-size:10px;color:#666;min-width:28px;text-align:center">${posLabel}</span>`
                    + `<button class="se-set-nav-btn" data-key="${nextPos.key}" style="${nb}" title="Next in spawn set">▶</button>`
                    + `</div>`;
            };

            const editSection = _editMode ? (() => {
                if (!spawnKey) return '';
                if (!spawnInfo) {
                    // Empty node — show set nav + drop zone + paste button if clipboard has data
                    return `<div class="popup-edit-section">` +
                        buildSetNavRow() +
                        `<div class="se-spawn-view se-spawn-empty" style="min-height:40px;display:flex;align-items:center;justify-content:center;border:1px dashed rgba(120,120,120,0.4);border-radius:3px;margin:4px 0">` +
                        `<span style="color:#666;font-size:11px;pointer-events:none">Drop an enemy here to add a spawn</span>` +
                        `</div>` +
                        (_copiedEnemyConfig
                            ? `<div style="text-align:center;margin-top:4px">` +
                              `<button class="popup-edit-btn accent" data-edit-action="paste-config">📋 Paste ${emNames[_copiedEnemyConfig.emCode]?.name ?? _copiedEnemyConfig.emCode ?? 'enemy'}</button>` +
                              `</div>`
                            : '') +
                        `</div>`;
                }
                const rawIdx = spawnInfo._rawIdx ?? '';
                const inp = (key, val, w='48px', type='number') =>
                    `<input class="popup-edit-input" data-edit="${key}" type="${type}" value="${val ?? ''}" style="width:${w}">`;
                const chk = (key, val, label, title='') =>
                    `<label style="display:inline-flex;align-items:center;gap:3px;white-space:nowrap;cursor:pointer"${title ? ` title="${title}"` : ''}><input type="checkbox" data-edit="${key}"${val ? ' checked' : ''}> <span style="font-size:11px">${label}</span></label>`;
                const lbl = (text, content, title='') =>
                    `<label style="display:inline-flex;flex-direction:column;gap:1px"${title ? ` title="${title}"` : ''}>` +
                    `<span style="color:#888;font-size:9px;text-transform:uppercase;letter-spacing:0.4px">${text}</span>${content}</label>`;
                const grp = (title, ...rows) =>
                    `<div style="margin-top:4px">` +
                    `<div style="font-size:9px;color:#aaa;text-transform:uppercase;letter-spacing:0.6px;border-bottom:1px solid rgba(128,128,128,0.2);padding-bottom:1px;margin-bottom:2px">${title}</div>` +
                    rows.map(r => `<div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:2px">${r}</div>`).join('') +
                    `</div>`;
                const PRESETS = [
                    { label: 'Always', value: '00:00,23:59' },
                    { label: '☀ Day',  value: '07:00,17:59' },
                    { label: '🌙 Night',value: '18:00,06:59' },
                ];
                const curTime  = spawnInfo.spawnTime ?? '00:00,23:59';
                // Dropdown: Always / Day / Night (falls back to first preset if custom value)
                const timePicker =
                    `<select class="popup-edit-input" data-edit="spawnTime" style="font-size:11px;min-width:100px">` +
                    PRESETS.map(p =>
                        `<option value="${p.value}"${p.value === curTime ? ' selected' : ''}>${p.label}</option>`
                    ).join('') +
                    (PRESETS.some(p => p.value === curTime) ? '' :
                        `<option value="${curTime}" selected>${curTime} (custom)</option>`) +
                    `</select>`;
                // HmPreset + ThinkTbl controls (reused in both column and full-width contexts)
                const hmPresetCtrl = (() => {
                    const ep  = spawnInfo.emCode ? hmPresetsByEmCode.get(spawnInfo.emCode) : null;
                    const txt = ep ? `${ep.id}${ep.name ? ' — ' + ep.name : ''}` : '—';
                    return `<span style="font-size:11px;color:#666;padding:2px 4px;align-self:center" title="Derived from enemy type">${txt}</span>`;
                })();
                const thinkTblCtrl = (() => {
                    const ti    = spawnInfo.emCode ? emThinkInfo[spawnInfo.emCode] : null;
                    const cur   = spawnInfo.startThink ?? 0;
                    const notes = ti ? (thinkTableNotes[ti.res] ?? {}) : {};
                    const title = ti
                        ? `Think table index — ${ti.res}, observed range 0–${ti.max} in spawn data`
                        : 'AI behaviour table index';
                    const control = ti
                        ? `<select class="popup-edit-input" data-edit="startThink" style="width:auto;font-size:11px" title="${title}">` +
                          Array.from({length: ti.max + 1}, (_,i) => {
                              const note = notes[i] ? ` — ${notes[i]}` : '';
                              return `<option value="${i}"${cur===i?' selected':''}>${i}${note}</option>`;
                          }).join('') + `</select>`
                        : inp('startThink', cur, '56px');
                    const resLabel = ti ? ` [${ti.res}]` : '';
                    return lbl(`Think Tbl${resLabel}`, control, title);
                })();
                // ── Spawn-set mode bar ────────────────────────────────────────
                const setModeBar = (() => {
                    const btnBase = 'font-size:10px;padding:1px 8px;border-radius:10px;cursor:pointer;border:1px solid;';
                    const singleOn = !_spawnSetMode;
                    const singleBtn = `<button class="se-set-single-btn" style="${btnBase}${singleOn ? 'background:#4a90d9;color:#fff;border-color:#357abd' : 'background:#f0f0f0;color:#555;border-color:#ccc'}">Single</button>`;
                    const setBtn    = `<button class="se-set-mode-btn"   style="${btnBase}${_spawnSetMode ? 'background:#4a90d9;color:#fff;border-color:#357abd' : 'background:#f0f0f0;color:#555;border-color:#ccc'}">Set ${sg}</button>`;
                    const toggleRow = `<div style="display:flex;gap:4px;align-items:center;margin-bottom:4px"><span style="font-size:9px;color:#aaa;text-transform:uppercase;letter-spacing:0.4px;margin-right:2px">Edit</span>${singleBtn}${setBtn}</div>`;
                    if (!_spawnSetMode) {
                        return toggleRow + buildSetNavRow();
                    }
                    // In set mode — determine peers (same SpawnGroup) and check for mixed content
                    const peers = getPeers(spawnCache);
                    const total = peers.length + 1;
                    const navRow = buildSetNavRow();
                    if (!peers.length) {
                        const info = `<div style="background:#f0f4ff;border:1px solid #b0c4de;border-radius:3px;padding:3px 7px;font-size:11px;color:#555;margin-bottom:4px">⚡ Spawn Set ${sg} — only position in this group with that value</div>`;
                        return toggleRow + navRow + info;
                    }
                    const curEmCode = spawnInfo?.emCode;
                    const hasMixed  = !_spawnSetConflictDismissed && peers.some(p => p.entries.some(e => e.emCode && e.emCode !== curEmCode));
                    if (hasMixed) {
                        const nb = 'font-size:10px;padding:1px 7px;border-radius:3px;cursor:pointer;border:1px solid;';
                        const conflict = `<div class="se-set-conflict" style="background:#fff8e1;border:1px solid #f5c518;border-radius:3px;padding:4px 7px;font-size:11px;color:#555;margin-bottom:4px">` +
                            `⚠ Spawn Set ${sg} has mixed enemies across ${total} positions.&nbsp;` +
                            `<button class="se-set-use-template-btn" style="${nb}background:#4a90d9;color:#fff;border-color:#357abd">Use this as template</button>&nbsp;` +
                            `<button class="se-set-keep-diffs-btn"   style="${nb}background:#f0f0f0;color:#555;border-color:#ccc">Keep differences</button>` +
                            `</div>`;
                        return toggleRow + navRow + conflict;
                    }
                    const emptyCount   = peers.filter(p => !p.entries.length).length;
                    const nb           = 'font-size:10px;padding:1px 7px;border-radius:3px;cursor:pointer;border:1px solid;';
                    const fillBtn      = emptyCount > 0
                        ? `<button class="se-set-fill-btn" data-fill-mode="empty" style="${nb}background:#4a90d9;color:#fff;border-color:#357abd" title="Copy this position's enemy and all values to the ${emptyCount} empty position${emptyCount > 1 ? 's' : ''} in this spawn set">📋 Fill ${emptyCount} empty</button>` : '';
                    const copyAllBtn   = `<button class="se-set-fill-btn" data-fill-mode="all" style="${nb}background:#d97b4a;color:#fff;border-color:#b85e2e" title="Overwrite all ${total} positions in this spawn set with this position's enemy and values">📋 Copy to all ${total}</button>`;
                    const removeAllBtn = `<button class="se-set-remove-all-btn" style="${nb}background:#c0392b;color:#fff;border-color:#96281b" title="Remove enemy data from all ${total} positions in this spawn set">🗑 Remove all</button>`;
                    const banner = `<div style="background:#f0f7f0;border:1px solid #7ab87a;border-radius:3px;padding:3px 7px;font-size:11px;color:#3a6b3a;margin-bottom:4px">`
                        + `<div>⚡ Spawn Set ${sg} — ${total} positions</div>`
                        + `<div style="display:flex;gap:4px;margin-top:3px">${fillBtn}${copyAllBtn}${removeAllBtn}</div>`
                        + `</div>`;
                    return toggleRow + navRow + banner;
                })();
                return `<div class="popup-edit-section"><div class="se-spawn-view">${setModeBar}` +
                    grp('Drops',
                        ...(() => {
                            const dtId  = spawnInfo.dropsTableId ?? -1;
                            const dt    = dtId >= 0 ? _dropsTablesMap.get(dtId) : null;
                            const label = dt ? dt.name : 'None';
                            const idBadge = dtId >= 0 ? `<span style="color:#999;font-size:10px"> (id:${dtId}, ${dt?.items?.length ?? 0} items)</span>` : '';
                            const row1 = `<input type="hidden" data-edit="dropsTableId" value="${dtId}">` +
                                `<div class="se-drops-row1" style="display:flex;align-items:center;gap:5px;flex-wrap:wrap">` +
                                `<span class="se-drops-label" style="font-size:11px;color:#444;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${label} (id:${dtId})">${label}${idBadge}</span>` +
                                `<button class="popup-edit-btn se-drops-picker-btn" style="font-size:10px;padding:1px 6px">Change</button>` +
                                (dtId >= 0 ? `<button class="popup-edit-btn se-drops-edit-btn" data-dt="${dtId}" style="font-size:10px;padding:1px 6px">Edit</button>` : '') +
                                `</div>`;
                            const items = dt?.items ?? [];
                            if (!items.length) return [row1];
                            const chips = items.map(row => {
                                const itemId   = row[0] ?? 0;
                                const minQty   = row[1] ?? 1;
                                const maxQty   = row[2] ?? 1;
                                const dropRate = row[5] ?? 0;
                                const iconNo   = itemNames[String(itemId)]?.iconNo;
                                const iconFile = iconNo != null ? `ii${String(iconNo).padStart(6,'0')}.png` : null;
                                const name     = itemNames[String(itemId)]?.name ?? `#${itemId}`;
                                const qty      = maxQty > minQty ? `×${minQty}–${maxQty}` : `×${minQty}`;
                                const pct      = dropRate > 0 ? ` ${(dropRate * 100).toFixed(0)}%` : '';
                                const imgEl    = iconFile && _iconIdSet.has(iconNo)
                                    ? `<img src="images/icons/small/${iconFile}" width="20" height="20" style="image-rendering:pixelated;vertical-align:middle" title="${name}">`
                                    : `<span style="display:inline-block;width:20px;height:20px;background:#ddd;border-radius:2px;font-size:8px;text-align:center;line-height:20px;vertical-align:middle" title="${name}">${itemId}</span>`;
                                return `<span style="display:inline-flex;align-items:center;gap:2px;white-space:nowrap">` +
                                    imgEl +
                                    `<span style="font-size:9px;color:#666">${qty}${pct}</span>` +
                                    `</span>`;
                            }).join('');
                            return [row1, `<div class="se-drops-chips" style="display:flex;flex-wrap:wrap;gap:4px 6px">${chips}</div>`];
                        })()
                    ) +
                    `<div class="se-grp-cols">` +
                    // ── Left column ──────────────────────────────────────────
                    `<div>` +
                    grp('Stats',
                        lbl('Level',      inp('lv',  spawnInfo.lv  ?? 1, '52px')) +
                        lbl('Experience', inp('exp', spawnInfo.exp ?? 0, '60px')) +
                        lbl('Play Pts',   inp('ppDrop', spawnInfo.ppDrop ?? 0, '56px'), 'Play Points — post-cap experience gained after reaching max EXP'),
                        lbl('Named ID',
                            `<input type="hidden" data-edit="namedId" value="${spawnInfo.namedId ?? 0}">` +
                            `<button class="popup-edit-btn se-named-picker-btn" data-named-id="${spawnInfo.namedId ?? 0}" style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:left" title="Click to pick a named enemy param">${namedParamLabel(namedParamsById.get(spawnInfo.namedId ?? 0))}</button>`
                        ),
                    ) +
                    grp('Spawn',
                        lbl('Time', timePicker),
                        lbl('Scale %',   inp('scale',      spawnInfo.scale      ?? 100, '56px')) +
                        lbl('Set Type', (() => {
                            const cur = spawnInfo.setType ?? 0;
                            const opts = [
                                [0, '0 — Normal'],
                                [1, '1 — (Unknown)'],
                                [2, '2 — Gather Spawn'],
                                [3, '3 — Network Spawn'],
                            ];
                            return `<select class="popup-edit-input" data-edit="setType" style="width:auto;font-size:11px" title="0: Normal position from layout&#10;2: Spawns at linked gather node (mOmUID)&#10;3: Spawns at live network/player position">` +
                                opts.map(([v, l]) => `<option value="${v}"${cur===v?' selected':''}>${l}</option>`).join('') +
                                `</select>`;
                        })()),
                        lbl('Repop Num', inp('repopNum',   spawnInfo.repopNum   ?? 0,   '56px')) +
                        lbl('Repop Cnt', inp('repopCount', spawnInfo.repopCount ?? 0,   '56px')),
                    ) +
                    `</div>` +
                    // ── Right column ─────────────────────────────────────────
                    `<div>` +
                    grp('Combat',
                        lbl('Variant', (() => {
                            const cur = spawnInfo.infection ?? 0;
                            const opts = [
                                [0, '0 — None'],
                                [1, '1 — Infected'],
                                [2, '2 — Severely Infected'],
                                [3, '3 — War-Ready'],
                            ];
                            return `<select class="popup-edit-input" data-edit="infection" style="width:auto;font-size:11px">` +
                                opts.map(([v, l]) => `<option value="${v}"${cur===v?' selected':''}>${l}</option>`).join('') +
                                `</select>`;
                        })()) +
                        lbl('Target Type', (() => {
                            const cur = spawnInfo.targetTypeId ?? 1;
                            const opts = [
                                [1, '1 — None'],
                                [2, '2 — #2'],
                                [3, '3 — #3'],
                                [4, '4 — #4'],
                                [6, '6 — Area Boss'],
                                [7, '7 — Stage Boss'],
                            ];
                            return `<select class="popup-edit-input" data-edit="targetTypeId" style="width:auto;font-size:11px">` +
                                opts.map(([v, l]) => `<option value="${v}"${cur===v?' selected':''}>${l}</option>`).join('') +
                                `</select>`;
                        })()) +
                        (() => {
                            const mi    = spawnInfo.emCode ? (emMontageInfo[spawnInfo.emCode]?.length > 0 ? emMontageInfo[spawnInfo.emCode] : null) : null;
                            const cur   = spawnInfo.montage ?? 0;
                            const notes = mi ? (montageNotes[spawnInfo.emCode] ?? {}) : {};
                            const title = mi
                                ? `Montage Fix — controls the enemy's appearance (model parts, colors, attachments). 0 typically randomizes the look each spawn; other values lock a specific variant. Valid indices extracted from the enemy's .dme file: ${mi.join(', ')}`
                                : 'Montage Fix — controls the enemy\'s appearance (model parts, colors, attachments). 0 typically randomizes the look each spawn; other values lock a specific variant.';
                            const control = mi
                                ? `<select class="popup-edit-input" data-edit="montage" style="width:auto;font-size:11px" title="${title}">` +
                                  mi.map(v => {
                                      const note = notes[String(v)] ? ` — ${notes[String(v)]}` : '';
                                      return `<option value="${v}"${cur===v?' selected':''}>${v}${note}</option>`;
                                  }).join('') +
                                  (mi.includes(cur) ? '' : `<option value="${cur}" selected>${cur} (custom)</option>`) +
                                  `</select>`
                                : inp('montage', cur, '56px');
                            return lbl('Montage Fix', control, title);
                        })(),
                    ) +
                    grp('Rewards',
                        `<div style="display:flex;flex-direction:column;gap:4px">` +
                        [['🩸', 'Blood Orbs', 'bloodOrbs', spawnInfo.bloodOrbs ?? 0, 'isBloodOrbEnemy', spawnInfo.isBloodOrbEnemy],
                         ['⭐', 'High Orbs',  'highOrbs',  spawnInfo.highOrbs  ?? 0, 'isHighOrbEnemy',  spawnInfo.isHighOrbEnemy]].map(([icon, name, key, val, markKey, marked]) =>
                            `<div style="display:flex;align-items:center;gap:5px">` +
                            `<span style="width:14px;text-align:center;font-size:13px">${icon}</span>` +
                            `<span style="color:#888;font-size:9px;text-transform:uppercase;letter-spacing:0.4px;width:54px">${name}</span>` +
                            inp(key, val, '56px') +
                            `<label class="orb-map-toggle" title="Mark as ${name} spawn on map">` +
                            `<input type="checkbox" class="popup-edit-input" data-edit="${markKey}"${marked ? ' checked' : ''}>📍</label>` +
                            `</div>`
                        ).join('') +
                        `</div>`,
                    ) +
                    grp('Boss',
                        lbl('Raid Boss ID', inp('raidBossId', spawnInfo.raidBossId ?? 0, '64px')),
                        chk('isBossGauge', spawnInfo.isBossGauge, 'Boss Gauge') + '&nbsp;&nbsp;' +
                        chk('isBossBGM',   spawnInfo.isBossBGM,   'Boss BGM')   + '&nbsp;&nbsp;' +
                        chk('isAreaBoss',  spawnInfo.isAreaBoss,   'Area Boss'),
                    ) +
                    grp('Behaviour',
                        lbl('Hm Preset', hmPresetCtrl) + thinkTblCtrl,
                        chk('isManualSet', spawnInfo.isManualSet, 'Manual Set', 'Enemy spawns dormant at its exact position (mIsWaitting=true). Activated when a boss fires a SummonSet FSM action (cEmActAtkSummonSet). Position index is significant — the client uses the exact spawn slot.'),
                    ) +
                    `</div>` +
                    `</div>` +
                    `</div>` +
                    `<div style="display:flex;gap:6px;margin-top:8px">` +
                    `<button class="popup-edit-btn" data-edit-action="apply" data-raw="${rawIdx}" style="flex:1;opacity:0.45;cursor:not-allowed" disabled>✔ Apply</button>` +
                    `<button class="popup-edit-btn" data-edit-action="copy-config" title="Copy this enemy's config to clipboard">📋 Copy</button>` +
                    (_copiedEnemyConfig ? `<button class="popup-edit-btn accent" data-edit-action="paste-config" title="Paste clipboard config onto this enemy">📋 Paste</button>` : '') +
                    `<button class="popup-edit-btn danger" data-edit-action="remove-spawn" data-raw="${rawIdx}">🗑 Remove</button>` +
                    `</div></div>`;
            })() : '';
            return `${badge}<br>${groupLabel}, Index: <b>${idx}</b>${subLine}${triggerLine}${cycleHtml}${emLine}${keyLine}${buildBossLine(spawnInfo)}${buildManualSetLine(spawnInfo)}${radiiLine}${orbsLine}${_editMode ? '' : buildDropsHtml(spawnInfo)}${editSection}`;
        };

        const buildTooltip = (spawnCache) => {
            const entries = spawnKey && spawnCache ? (spawnCache.get(spawnKey) ?? []) : [];
            const e0 = entries[0] ?? null;
            const mobLabel = resolveSpawnMobLabel(spawnCache, spawnKey, spawn);
            const orbBadge  = (e0?.isBloodOrbEnemy && e0?.bloodOrbs ? ' 🩸' : '') + (e0?.isHighOrbEnemy && e0?.highOrbs ? ' ⭐' : '');
            const manualBadge = e0?.isManualSet ? ' 😴' : '';
            const bossBadge = (e0?.isBossGauge || e0?.isAreaBoss || e0?.raidBossId > 0) ? ' <span style="color:#ff4444" title="Boss enemy">☠</span>' : '';
            const idPart = useMobSpawnDevLabels() ? `${g.groupId}.${idx} [SS:${sg}]` : '';
            const core = [mobLabel, idPart].filter(Boolean).join(' — ');
            return `${core || DYNAMIC_ENEMY_LABEL}${orbBadge}${manualBadge}${bossBadge}${isKeyBearer ? ' <span style="color:#c8a000;font-size:16px;">🗝</span>' : ''}`;
        };

        const marker = L.circleMarker(latlng, {
            renderer:    spawnRenderer,
            className:   `enemy-marker${isKeyBearer ? ' key-bearer-spawn' : ''}`,
            color:       isKeyBearer ? '#c8a000' : g.color,
            fillColor,
            fillOpacity: 0.85,
            weight:      isKeyBearer ? 3.5 : 2.5,
            radius:      5,
        })
            .bindPopup(buildEnemyPopup(_enemySpawnCache), { minWidth: 320, maxWidth: 500 })
            .bindTooltip('', { direction: 'top', offset: [0, -8] });

        // Rebuild tooltip fresh on every hover so level/name is always current
        marker.on('tooltipopen', function() {
            const tt = buildTooltip(_enemySpawnCache);
            this.setTooltipContent(tt);
            this._label = tt;
            this._naturalTooltip = tt;
            if (!_enemySpawnCache && spawnKey) {
                _enemySpawnPromise.then(cache => {
                    const updated = buildTooltip(cache);
                    this._label = updated;
                    this._naturalTooltip = updated;
                    if (this.isTooltipOpen()) this.setTooltipContent(updated);
                });
            }
        });

        // Rebuild popup content and attach cycle-button handlers on open.
        // Uses direct innerHTML update on the content div to avoid Leaflet's
        // setContent/update reflow, which closes tooltips and repositions the popup.
        const watchEditChanges = (cd) => {
            const applyBtn = cd?.querySelector('[data-edit-action="apply"]');
            if (!applyBtn) return;
            const enable = () => {
                applyBtn.disabled = false;
                applyBtn.style.opacity = '';
                applyBtn.style.cursor = '';
            };
            const section = cd.querySelector('.popup-edit-section');
            if (!section) return;
            section.addEventListener('input', enable, { once: true });
            section.addEventListener('change', enable, { once: true });
        };
        let _popupClickHandler = null;
        marker.on('popupopen', function() {
            leafletMap.on('move', _repositionNamedStatsPanel);
            const popup = this.getPopup();
            _rebuildOpenPopup = () => {
                const el = popup.getElement();
                if (!el) return;
                const cd = el.querySelector('.leaflet-popup-content');
                if (cd) {
                    cd.innerHTML = buildEnemyPopup(_enemySpawnCache);
                    cd.style.width = '';
                    popup._updateLayout?.();
                    popup._updatePosition?.();
                    const ents0  = _enemySpawnCache?.get(spawnKey) ?? [];
                    const si0    = ents0[displayIdx] ?? null;
                    const baseEm0 = si0?.emCode ? (emNames[si0.emCode]?.name ?? null) : null;
                    const ni = cd.querySelector('[data-edit="namedId"]');
                    if (_editMode && ni) showNamedStatsPanel(parseInt(ni.value) || 0, el, baseEm0);
                    else hideNamedStatsPanel();
                    watchEditChanges(cd);
                }
            };
            const bind = (cache) => {
                requestAnimationFrame(() => {
                    const el = popup.getElement();
                    if (!el) return;
                    // Update content without triggering Leaflet reflow
                    const contentDiv = el.querySelector('.leaflet-popup-content');
                    if (contentDiv) {
                        contentDiv.innerHTML = buildEnemyPopup(cache);
                        // Reposition popup to account for new content size (edit mode
                        // popup is taller than the initial view-mode content).
                        // _updateLayout/_updatePosition don't trigger tooltip teardown.
                        popup._updateLayout?.();
                        popup._updatePosition?.();
                        const ents   = cache?.get(spawnKey) ?? [];
                        const si     = ents[displayIdx] ?? null;
                        const baseEm = si?.emCode ? (emNames[si.emCode]?.name ?? null) : null;
                        const ni = contentDiv.querySelector('[data-edit="namedId"]');
                        if (_editMode && ni) showNamedStatsPanel(parseInt(ni.value) || 0, el, baseEm);
                        else hideNamedStatsPanel();
                        watchEditChanges(contentDiv);
                    }
                    // Replace click handler (event delegation — survives innerHTML swaps)
                    if (_popupClickHandler) el.removeEventListener('click', _popupClickHandler);
                    // Always read entries fresh from cache — the cache may gain entries after
                    // this handler is registered (e.g. drag-drop onto a previously empty spot).
                    const getEntries = () => cache?.get(spawnKey) ?? [];
                    const rebuildPopup = () => {
                        const cd = el.querySelector('.leaflet-popup-content');
                        if (!cd) return;
                        cd.innerHTML = buildEnemyPopup(cache);
                        popup._updateLayout?.();
                        popup._updatePosition?.();
                        const si     = getEntries()[displayIdx] ?? null;
                        const baseEm = si?.emCode ? (emNames[si.emCode]?.name ?? null) : null;
                        const ni = cd.querySelector('[data-edit="namedId"]');
                        if (_editMode && ni) showNamedStatsPanel(parseInt(ni.value) || 0, el, baseEm);
                        else hideNamedStatsPanel();
                        watchEditChanges(cd);
                    };
                    _popupClickHandler = (e) => {
                        // ── Spawn-set set navigation ───────────────────────────
                        if (e.target.closest('.se-set-nav-btn')) {
                            e.stopPropagation();
                            const targetKey = e.target.closest('.se-set-nav-btn').dataset.key;
                            let targetMarker = null;
                            for (const markers of Object.values(g.sgMarkers)) {
                                for (const m of markers) {
                                    if (m._spawnKey === targetKey) { targetMarker = m; break; }
                                }
                                if (targetMarker) break;
                            }
                            if (targetMarker) {
                                marker.closePopup();
                                setTimeout(() => targetMarker.openPopup(), 0);
                            }
                            return;
                        }
                        // ── Spawn-set mode toggles ─────────────────────────────
                        if (e.target.closest('.se-set-single-btn')) {
                            e.stopPropagation();
                            _spawnSetMode = false;
                            _spawnSetConflictDismissed = false;
                            rebuildPopup();
                            return;
                        }
                        if (e.target.closest('.se-set-mode-btn')) {
                            e.stopPropagation();
                            _spawnSetMode = true;
                            _spawnSetConflictDismissed = false;
                            rebuildPopup();
                            return;
                        }
                        // ── Conflict resolution ────────────────────────────────
                        if (e.target.closest('.se-set-keep-diffs-btn')) {
                            e.stopPropagation();
                            _spawnSetConflictDismissed = true;
                            rebuildPopup();
                            return;
                        }
                        if (e.target.closest('.se-set-fill-btn')) {
                            e.stopPropagation();
                            const fillMode = e.target.closest('.se-set-fill-btn').dataset.fillMode; // 'empty' | 'all'
                            const si0 = getEntries()[displayIdx];
                            if (!si0) return;
                            const peers = getPeers(cache);
                            const fields = {
                                emCode: si0.emCode, lv: si0.lv, bloodOrbs: si0.bloodOrbs,
                                highOrbs: si0.highOrbs, scale: si0.scale, exp: si0.exp,
                                repopNum: si0.repopNum, repopCount: si0.repopCount,
                                setType: si0.setType, infection: si0.infection,
                                targetTypeId: si0.targetTypeId, spawnTime: si0.spawnTime,
                                namedId: si0.namedId, raidBossId: si0.raidBossId,
                                isBossGauge: si0.isBossGauge, isBossBGM: si0.isBossBGM,
                                isAreaBoss: si0.isAreaBoss, isManualSet: si0.isManualSet,
                                isBloodOrbEnemy: si0.isBloodOrbEnemy, isHighOrbEnemy: si0.isHighOrbEnemy,
                                startThink: si0.startThink, montage: si0.montage,
                                ppDrop: si0.ppDrop, dropsTableId: si0.dropsTableId, drops: si0.drops,
                                subGroupId: si0.subGroupId,
                            };
                            const hexId = si0.emCode ? ('0x' + si0.emCode.slice(2).toUpperCase()) : null;
                            for (const peer of peers) {
                                if (peer.entries.length > 0 && fillMode === 'all') {
                                    // Overwrite existing entries (copy-to-all mode only)
                                    for (const pEntry of peer.entries) {
                                        Object.assign(pEntry, fields);
                                        if (_rawEnemyData && pEntry._rawIdx >= 0) {
                                            const pRow = _rawEnemyData.enemies[pEntry._rawIdx];
                                            if (pRow) {
                                                const { iEnemyId, iLv, iBlood, iHigh, iScale, iNamed, iRaidBoss,
                                                        iSetType, iInfection, iIsBossG, iIsBossBGM, iIsAreaBoss,
                                                        iExp, iRepopNum, iRepopCount, iTargetType, iSpawnTime,
                                                        iStartThink, iMontage, iIsManualSet, iPPDrop,
                                                        iIsBloodOrbEnemy, iIsHighOrbEnemy, iDrops, iSubGroup } = _rawEnemySchemas;
                                                if (iEnemyId >= 0 && hexId) pRow[iEnemyId] = hexId;
                                                pRow[iLv] = fields.lv;
                                                pRow[iBlood] = fields.bloodOrbs;
                                                pRow[iHigh]  = fields.highOrbs;
                                                if (iScale      >= 0) pRow[iScale]      = fields.scale;
                                                if (iExp        >= 0) pRow[iExp]        = fields.exp;
                                                if (iRepopNum   >= 0) pRow[iRepopNum]   = fields.repopNum;
                                                if (iRepopCount >= 0) pRow[iRepopCount] = fields.repopCount;
                                                if (iSetType    >= 0) pRow[iSetType]    = fields.setType;
                                                if (iInfection  >= 0) pRow[iInfection]  = fields.infection;
                                                if (iTargetType >= 0) pRow[iTargetType] = fields.targetTypeId;
                                                if (iSpawnTime  >= 0) pRow[iSpawnTime]  = fields.spawnTime;
                                                if (iNamed      >= 0) pRow[iNamed]      = fields.namedId;
                                                if (iRaidBoss   >= 0) pRow[iRaidBoss]   = fields.raidBossId;
                                                if (iIsBossG    >= 0) pRow[iIsBossG]    = fields.isBossGauge;
                                                if (iIsBossBGM  >= 0) pRow[iIsBossBGM]  = fields.isBossBGM;
                                                if (iIsAreaBoss >= 0) pRow[iIsAreaBoss] = fields.isAreaBoss;
                                                if (iIsManualSet    >= 0) pRow[iIsManualSet]    = fields.isManualSet;
                                                if (iIsBloodOrbEnemy >= 0) pRow[iIsBloodOrbEnemy] = fields.isBloodOrbEnemy;
                                                if (iIsHighOrbEnemy  >= 0) pRow[iIsHighOrbEnemy]  = fields.isHighOrbEnemy;
                                                if (iStartThink >= 0) pRow[iStartThink] = fields.startThink;
                                                if (iMontage    >= 0) pRow[iMontage]    = fields.montage;
                                                if (iPPDrop     >= 0) pRow[iPPDrop]     = fields.ppDrop;
                                                if (iSubGroup   >= 0) pRow[iSubGroup]   = fields.subGroupId;
                                                pRow[iDrops] = fields.dropsTableId;
                                            }
                                        }
                                    }
                                } else {
                                    // Create a brand-new entry for this empty position
                                    const [pSid, pGid, pPosIdx] = peer.key.split(',');
                                    const newEntry = { ...fields, hmPreset: si0.hmPreset ?? 0, _rawIdx: -1 };
                                    if (_rawEnemyData && hexId) {
                                        const { iStage, iGroup, iPosIdx: iPI, iEnemyId, iLv, iBlood, iHigh,
                                                iSpawnTime, iDrops, iScale, iSubGroup, iNamed, iRaidBoss,
                                                iSetType, iInfection, iIsBossG, iIsBossBGM, iIsAreaBoss,
                                                iIsBloodOrbEnemy, iIsHighOrbEnemy,
                                                iExp, iRepopNum, iRepopCount, iTargetType,
                                                iHmPreset, iStartThink, iMontage, iIsManualSet, iPPDrop,
                                                iLayerNo } = _rawEnemySchemas;
                                        const newRaw = new Array(_rawEnemyData.schemas.enemies.length).fill(null);
                                        newRaw[iStage]  = Number(pSid);
                                        newRaw[iGroup]  = Number(pGid);
                                        newRaw[iPI]     = Number(pPosIdx);
                                        if (iLayerNo >= 0) newRaw[iLayerNo] = 0;
                                        if (iEnemyId >= 0) newRaw[iEnemyId] = hexId;
                                        newRaw[iLv]     = fields.lv;
                                        newRaw[iBlood]  = fields.bloodOrbs;
                                        newRaw[iHigh]   = fields.highOrbs;
                                        newRaw[iSpawnTime] = fields.spawnTime ?? '00:00,23:59';
                                        newRaw[iDrops]  = fields.dropsTableId;
                                        if (iScale      >= 0) newRaw[iScale]      = fields.scale;
                                        if (iSubGroup   >= 0) newRaw[iSubGroup]   = fields.subGroupId;
                                        if (iNamed      >= 0) newRaw[iNamed]      = fields.namedId;
                                        if (iRaidBoss   >= 0) newRaw[iRaidBoss]   = fields.raidBossId;
                                        if (iSetType    >= 0) newRaw[iSetType]    = fields.setType;
                                        if (iInfection  >= 0) newRaw[iInfection]  = fields.infection;
                                        if (iIsBossG    >= 0) newRaw[iIsBossG]    = fields.isBossGauge;
                                        if (iIsBossBGM  >= 0) newRaw[iIsBossBGM]  = fields.isBossBGM;
                                        if (iIsAreaBoss >= 0) newRaw[iIsAreaBoss] = fields.isAreaBoss;
                                        if (iIsManualSet    >= 0) newRaw[iIsManualSet]    = fields.isManualSet;
                                        if (iIsBloodOrbEnemy >= 0) newRaw[iIsBloodOrbEnemy] = fields.isBloodOrbEnemy;
                                        if (iIsHighOrbEnemy  >= 0) newRaw[iIsHighOrbEnemy]  = fields.isHighOrbEnemy;
                                        if (iExp        >= 0) newRaw[iExp]        = fields.exp;
                                        if (iRepopNum   >= 0) newRaw[iRepopNum]   = fields.repopNum;
                                        if (iRepopCount >= 0) newRaw[iRepopCount] = fields.repopCount;
                                        if (iTargetType >= 0) newRaw[iTargetType] = fields.targetTypeId;
                                        if (iHmPreset   >= 0) newRaw[iHmPreset]   = hmPresetsByEmCode.get(fields.emCode)?.id ?? 0;
                                        if (iStartThink >= 0) newRaw[iStartThink] = fields.startThink;
                                        if (iMontage    >= 0) newRaw[iMontage]    = fields.montage;
                                        if (iIsManualSet >= 0) newRaw[iIsManualSet] = fields.isManualSet;
                                        if (iPPDrop     >= 0) newRaw[iPPDrop]     = fields.ppDrop;
                                        _rawEnemyData.enemies.push(newRaw);
                                        newEntry._rawIdx = _rawEnemyData.enemies.length - 1;
                                    }
                                    cache.set(peer.key, [newEntry]);
                                }
                            }
                            _spawnSetConflictDismissed = true;
                            if (_markDirty) _markDirty('ddon-src-spawns');
                            rebuildPopup();
                            return;
                        }
                        if (e.target.closest('.se-set-remove-all-btn')) {
                            e.stopPropagation();
                            const peers = getPeers(cache);
                            // Remove current position's entry too
                            if (spawnKey) {
                                const arr = cache.get(spawnKey);
                                if (arr) {
                                    if (_rawEnemyData) {
                                        for (const ent of arr) {
                                            if (ent._rawIdx >= 0) {
                                                _rawEnemyData.enemies.splice(ent._rawIdx, 1);
                                                for (const entryArr of cache.values())
                                                    for (const e2 of entryArr)
                                                        if (e2._rawIdx > ent._rawIdx) e2._rawIdx--;
                                            }
                                        }
                                    }
                                    cache.delete(spawnKey);
                                }
                            }
                            // Remove all peer entries
                            for (const peer of peers) {
                                if (!peer.entries.length) continue;
                                if (_rawEnemyData) {
                                    for (const ent of peer.entries) {
                                        if (ent._rawIdx >= 0) {
                                            _rawEnemyData.enemies.splice(ent._rawIdx, 1);
                                            for (const entryArr of cache.values())
                                                for (const e2 of entryArr)
                                                    if (e2._rawIdx > ent._rawIdx) e2._rawIdx--;
                                        }
                                    }
                                }
                                cache.delete(peer.key);
                            }
                            if (_markDirty) _markDirty('ddon-src-spawns');
                            marker.closePopup();
                            return;
                        }
                        if (e.target.closest('.se-set-use-template-btn')) {
                            e.stopPropagation();
                            const si0 = getEntries()[displayIdx];
                            if (si0) {
                                const peers = getPeers(cache);
                                const fields = {
                                    emCode: si0.emCode, lv: si0.lv, bloodOrbs: si0.bloodOrbs,
                                    highOrbs: si0.highOrbs, scale: si0.scale, exp: si0.exp,
                                    repopNum: si0.repopNum, repopCount: si0.repopCount,
                                    setType: si0.setType, infection: si0.infection,
                                    targetTypeId: si0.targetTypeId, spawnTime: si0.spawnTime,
                                    namedId: si0.namedId, raidBossId: si0.raidBossId,
                                    isBossGauge: si0.isBossGauge, isBossBGM: si0.isBossBGM,
                                    isAreaBoss: si0.isAreaBoss, isManualSet: si0.isManualSet,
                                    isBloodOrbEnemy: si0.isBloodOrbEnemy, isHighOrbEnemy: si0.isHighOrbEnemy,
                                    startThink: si0.startThink, montage: si0.montage,
                                    ppDrop: si0.ppDrop, dropsTableId: si0.dropsTableId, drops: si0.drops,
                                };
                                for (const peer of peers) {
                                    for (const pEntry of peer.entries) {
                                        Object.assign(pEntry, fields);
                                        if (_rawEnemyData && pEntry._rawIdx >= 0) {
                                            const row = _rawEnemyData.enemies[pEntry._rawIdx];
                                            if (row) {
                                                const { iEnemyId, iLv, iBlood, iHigh, iScale, iNamed, iRaidBoss,
                                                        iSetType, iInfection, iIsBossG, iIsBossBGM, iIsAreaBoss,
                                                        iExp, iRepopNum, iRepopCount, iTargetType, iSpawnTime,
                                                        iStartThink, iMontage, iIsManualSet, iPPDrop,
                                                        iIsBloodOrbEnemy, iIsHighOrbEnemy, iDrops } = _rawEnemySchemas;
                                                if (iEnemyId >= 0 && fields.emCode) row[iEnemyId] = '0x' + fields.emCode.slice(2).toUpperCase();
                                                row[iLv]    = fields.lv;
                                                row[iBlood] = fields.bloodOrbs;
                                                row[iHigh]  = fields.highOrbs;
                                                if (iScale    >= 0) row[iScale]    = fields.scale;
                                                if (iExp      >= 0) row[iExp]      = fields.exp;
                                                if (iRepopNum >= 0) row[iRepopNum] = fields.repopNum;
                                                if (iRepopCount >= 0) row[iRepopCount] = fields.repopCount;
                                                if (iSetType  >= 0) row[iSetType]  = fields.setType;
                                                if (iInfection >= 0) row[iInfection] = fields.infection;
                                                if (iTargetType >= 0) row[iTargetType] = fields.targetTypeId;
                                                if (iSpawnTime >= 0) row[iSpawnTime] = fields.spawnTime;
                                                if (iNamed    >= 0) row[iNamed]    = fields.namedId;
                                                if (iRaidBoss >= 0) row[iRaidBoss] = fields.raidBossId;
                                                if (iIsBossG  >= 0) row[iIsBossG]  = fields.isBossGauge;
                                                if (iIsBossBGM >= 0) row[iIsBossBGM] = fields.isBossBGM;
                                                if (iIsAreaBoss >= 0) row[iIsAreaBoss] = fields.isAreaBoss;
                                                if (iIsManualSet >= 0) row[iIsManualSet] = fields.isManualSet;
                                                if (iIsBloodOrbEnemy >= 0) row[iIsBloodOrbEnemy] = fields.isBloodOrbEnemy;
                                                if (iIsHighOrbEnemy >= 0) row[iIsHighOrbEnemy] = fields.isHighOrbEnemy;
                                                if (iStartThink >= 0) row[iStartThink] = fields.startThink;
                                                if (iMontage  >= 0) row[iMontage]  = fields.montage;
                                                if (iPPDrop   >= 0) row[iPPDrop]   = fields.ppDrop;
                                                row[iDrops] = fields.dropsTableId;
                                            }
                                        }
                                    }
                                }
                                _spawnSetConflictDismissed = true;
                                if (_markDirty) _markDirty('ddon-src-spawns');
                            }
                            rebuildPopup();
                            return;
                        }
                        // ── Cycle prev/next ───────────────────────────────────
                        const cycleBtn = e.target.closest('.spawn-prev, .spawn-next');
                        if (cycleBtn) {
                            const entries = getEntries();
                            if (entries.length <= 1) return;
                            e.stopPropagation();
                            displayIdx = (displayIdx + (cycleBtn.classList.contains('spawn-prev') ? -1 : 1) + entries.length) % entries.length;
                            rebuildPopup();
                            return;
                        }
                        // ── Named param picker button ─────────────────────────
                        const namedPickerBtn = e.target.closest('.se-named-picker-btn');
                        if (namedPickerBtn && _editMode) {
                            e.stopPropagation();
                            const section = el.querySelector('.popup-edit-section');
                            const spawnInfo0 = getEntries()[displayIdx] ?? null;
                            const emCode0    = spawnInfo0?.emCode ?? spawn.EmName ?? null;
                            const baseEmName = emCode0 ? (emNames[emCode0]?.name ?? null) : null;
                            openNamedParamPicker(section, baseEmName);
                            return;
                        }
                        // ── Drop table picker ─────────────────────────────────
                        const dtPickerBtn = e.target.closest('.se-drops-picker-btn');
                        if (dtPickerBtn && _editMode) {
                            e.stopPropagation();
                            const section = el.querySelector('.popup-edit-section');
                            openDropTablePicker(section);
                            return;
                        }
                        // ── Drop table editor (from spawn) ────────────────────
                        const dtEditBtn = e.target.closest('.se-drops-edit-btn');
                        if (dtEditBtn && _editMode) {
                            e.stopPropagation();
                            openDropTableEditor(parseInt(dtEditBtn.dataset.dt));
                            return;
                        }
                        // ── Spawn time presets ────────────────────────────────
                        // ── Edit actions ──────────────────────────────────────
                        if (!_editMode) return;
                        const actionBtn = e.target.closest('[data-edit-action]');
                        if (!actionBtn) return;
                        e.stopPropagation();
                        const rawIdx  = parseInt(actionBtn.dataset.raw);
                        const action  = actionBtn.dataset.editAction;
                        const section = el.querySelector('.popup-edit-section');
                        if (action === 'apply' && section) {
                            const g = (key) => section.querySelector(`[data-edit="${key}"]`);
                            const iv = (key, def=0) => parseInt(g(key)?.value) || def;
                            const bv = (key) => g(key)?.checked ?? false;
                            const sv = (key) => g(key)?.value?.trim() || null;
                            const lv         = iv('lv', 1) || null;
                            const blood      = iv('bloodOrbs');
                            const high       = iv('highOrbs');
                            const scale      = iv('scale', 100);
                            const exp        = iv('exp');
                            const repopNum   = iv('repopNum');
                            const setType    = iv('setType');
                            const infection  = iv('infection');
                            const targetTypeId = iv('targetTypeId');
                            const spawnTime = g('spawnTime')?.value || '00:00,23:59';
                            const namedId     = (() => {
                                const btn = section.querySelector('.se-named-picker-btn');
                                if (btn?.dataset.namedId != null) return parseInt(btn.dataset.namedId) || 0;
                                return iv('namedId');
                            })();
                            const raidBossId  = iv('raidBossId');
                            const isBossGauge = bv('isBossGauge');
                            const isBossBGM   = bv('isBossBGM');
                            const isAreaBoss  = bv('isAreaBoss');
                            const isManualSet      = bv('isManualSet');
                            const isBloodOrbEnemy  = bv('isBloodOrbEnemy');
                            const isHighOrbEnemy   = bv('isHighOrbEnemy');
                            const repopCount  = iv('repopCount');
                            const startThink  = iv('startThink');
                            const montage     = iv('montage');
                            const ppDrop      = iv('ppDrop');
                            const dropsTableId = iv('dropsTableId', -1);
                            // Update cache entry (read fresh — entries may have been added after handler registration)
                            const entry = getEntries()[displayIdx];
                            if (entry) {
                                Object.assign(entry, {
                                    lv, bloodOrbs: blood, highOrbs: high,
                                    scale, exp, repopNum, repopCount, setType,
                                    infection, targetTypeId, spawnTime, namedId,
                                    raidBossId, isBossGauge, isBossBGM, isAreaBoss,
                                    isManualSet, isBloodOrbEnemy, isHighOrbEnemy,
                                    startThink, montage, ppDrop,
                                    dropsTableId,
                                    drops: dropsTableId >= 0 ? (_dropsTablesMap.get(dropsTableId)?.items ?? []) : [],
                                });
                            }
                            // Update raw JSON row
                            if (_rawEnemyData && !isNaN(rawIdx)) {
                                const row = _rawEnemyData.enemies[rawIdx];
                                if (row) {
                                    const { iLv, iBlood, iHigh, iScale, iNamed, iRaidBoss,
                                            iSetType, iInfection, iIsBossG, iIsBossBGM, iIsAreaBoss,
                                            iExp, iRepopNum, iRepopCount, iTargetType, iSpawnTime,
                                            iHmPreset, iStartThink, iMontage, iIsManualSet, iPPDrop,
                                            iIsBloodOrbEnemy, iIsHighOrbEnemy, iDrops } = _rawEnemySchemas;
                                    row[iLv]         = lv;
                                    row[iBlood]      = blood;
                                    row[iHigh]       = high;
                                    if (iScale       >= 0) row[iScale]       = scale;
                                    if (iExp         >= 0) row[iExp]         = exp;
                                    if (iRepopNum    >= 0) row[iRepopNum]    = repopNum;
                                    if (iRepopCount  >= 0) row[iRepopCount]  = repopCount;
                                    if (iSetType     >= 0) row[iSetType]     = setType;
                                    if (iInfection   >= 0) row[iInfection]   = infection;
                                    if (iTargetType  >= 0) row[iTargetType]  = targetTypeId;
                                    if (iSpawnTime   >= 0) row[iSpawnTime]   = spawnTime;
                                    if (iNamed       >= 0) row[iNamed]       = namedId;
                                    if (iRaidBoss    >= 0) row[iRaidBoss]    = raidBossId;
                                    if (iIsBossG     >= 0) row[iIsBossG]     = isBossGauge;
                                    if (iIsBossBGM   >= 0) row[iIsBossBGM]   = isBossBGM;
                                    if (iIsAreaBoss  >= 0) row[iIsAreaBoss]  = isAreaBoss;
                                    if (iIsManualSet     >= 0) row[iIsManualSet]     = isManualSet;
                                    if (iIsBloodOrbEnemy >= 0) row[iIsBloodOrbEnemy] = isBloodOrbEnemy;
                                    if (iIsHighOrbEnemy  >= 0) row[iIsHighOrbEnemy]  = isHighOrbEnemy;
                                    // HmPresetNo is derived from em code — not written back from UI
                                    if (iStartThink  >= 0) row[iStartThink]  = startThink;
                                    if (iMontage     >= 0) row[iMontage]     = montage;
                                    if (iPPDrop      >= 0) row[iPPDrop]      = ppDrop;
                                    row[iDrops] = dropsTableId;
                                }
                            }
                            // ── Propagate to spawn-set peers ──────────────────
                            if (_spawnSetMode) {
                                const peers = getPeers(cache);
                                const peerFields = {
                                    lv, bloodOrbs: blood, highOrbs: high, scale, exp,
                                    repopNum, repopCount, setType, infection, targetTypeId,
                                    spawnTime, namedId, raidBossId, isBossGauge, isBossBGM,
                                    isAreaBoss, isManualSet, isBloodOrbEnemy, isHighOrbEnemy,
                                    startThink, montage, ppDrop, dropsTableId,
                                    drops: dropsTableId >= 0 ? (_dropsTablesMap.get(dropsTableId)?.items ?? []) : [],
                                    // Note: subGroupId intentionally excluded — it defines the set membership
                                };
                                for (const peer of peers) {
                                    for (const pEntry of peer.entries) {
                                        Object.assign(pEntry, peerFields);
                                        if (_rawEnemyData && pEntry._rawIdx >= 0) {
                                            const pRow = _rawEnemyData.enemies[pEntry._rawIdx];
                                            if (pRow) {
                                                const { iLv, iBlood, iHigh, iScale, iNamed, iRaidBoss,
                                                        iSetType, iInfection, iIsBossG, iIsBossBGM, iIsAreaBoss,
                                                        iExp, iRepopNum, iRepopCount, iTargetType, iSpawnTime,
                                                        iStartThink, iMontage, iIsManualSet, iPPDrop,
                                                        iIsBloodOrbEnemy, iIsHighOrbEnemy, iDrops } = _rawEnemySchemas;
                                                pRow[iLv]    = lv;
                                                pRow[iBlood] = blood;
                                                pRow[iHigh]  = high;
                                                if (iScale      >= 0) pRow[iScale]      = scale;
                                                if (iExp        >= 0) pRow[iExp]        = exp;
                                                if (iRepopNum   >= 0) pRow[iRepopNum]   = repopNum;
                                                if (iRepopCount >= 0) pRow[iRepopCount] = repopCount;
                                                if (iSetType    >= 0) pRow[iSetType]    = setType;
                                                if (iInfection  >= 0) pRow[iInfection]  = infection;
                                                if (iTargetType >= 0) pRow[iTargetType] = targetTypeId;
                                                if (iSpawnTime  >= 0) pRow[iSpawnTime]  = spawnTime;
                                                if (iNamed      >= 0) pRow[iNamed]      = namedId;
                                                if (iRaidBoss   >= 0) pRow[iRaidBoss]   = raidBossId;
                                                if (iIsBossG    >= 0) pRow[iIsBossG]    = isBossGauge;
                                                if (iIsBossBGM  >= 0) pRow[iIsBossBGM]  = isBossBGM;
                                                if (iIsAreaBoss >= 0) pRow[iIsAreaBoss] = isAreaBoss;
                                                if (iIsManualSet    >= 0) pRow[iIsManualSet]    = isManualSet;
                                                if (iIsBloodOrbEnemy >= 0) pRow[iIsBloodOrbEnemy] = isBloodOrbEnemy;
                                                if (iIsHighOrbEnemy  >= 0) pRow[iIsHighOrbEnemy]  = isHighOrbEnemy;
                                                if (iStartThink >= 0) pRow[iStartThink] = startThink;
                                                if (iMontage    >= 0) pRow[iMontage]    = montage;
                                                if (iPPDrop     >= 0) pRow[iPPDrop]     = ppDrop;
                                                pRow[iDrops] = dropsTableId;
                                            }
                                        }
                                    }
                                }
                            }
                            if (_markDirty) _markDirty('ddon-src-spawns');
                            const cd = el.querySelector('.leaflet-popup-content');
                            if (cd) {
                                cd.innerHTML = buildEnemyPopup(cache);
                                popup._updateLayout?.();
                                popup._updatePosition?.();
                                watchEditChanges(cd);
                                const view = cd.querySelector('.se-spawn-view');
                                if (view) {
                                    view.style.transition = 'background 0.08s';
                                    view.style.background = 'rgba(80,200,120,0.35)';
                                    setTimeout(() => {
                                        view.style.transition = 'background 0.5s';
                                        view.style.background = '';
                                    }, 150);
                                }
                            }
                        } else if (action === 'copy-config') {
                            const src = getEntries()[displayIdx];
                            if (!src) return;
                            _copiedEnemyConfig = {
                                emCode: src.emCode, lv: src.lv,
                                bloodOrbs: src.bloodOrbs, highOrbs: src.highOrbs,
                                scale: src.scale, exp: src.exp,
                                repopNum: src.repopNum, repopCount: src.repopCount,
                                setType: src.setType, infection: src.infection,
                                targetTypeId: src.targetTypeId, spawnTime: src.spawnTime,
                                namedId: src.namedId, raidBossId: src.raidBossId,
                                isBossGauge: src.isBossGauge, isBossBGM: src.isBossBGM,
                                isAreaBoss: src.isAreaBoss, isManualSet: src.isManualSet,
                                isBloodOrbEnemy: src.isBloodOrbEnemy, isHighOrbEnemy: src.isHighOrbEnemy,
                                startThink: src.startThink, montage: src.montage, ppDrop: src.ppDrop,
                                dropsTableId: src.dropsTableId, drops: [...(src.drops ?? [])],
                            };
                            _updateClipboardBar();
                            actionBtn.textContent = '✓ Copied!';
                            setTimeout(() => { if (_rebuildOpenPopup) _rebuildOpenPopup(); }, 1000);
                        } else if (action === 'paste-config') {
                            if (!_copiedEnemyConfig) return;
                            const cfg = _copiedEnemyConfig;
                            const entry = getEntries()[displayIdx];
                            const hexId = cfg.emCode ? ('0x' + cfg.emCode.slice(2).toUpperCase()) : null;
                            if (entry) {
                                // ── Overwrite existing entry ──────────────────
                                Object.assign(entry, cfg, { drops: [...(cfg.drops ?? [])] });
                                if (_rawEnemyData && entry._rawIdx >= 0) {
                                    const row = _rawEnemyData.enemies[entry._rawIdx];
                                    if (row) {
                                        const { iEnemyId, iLv, iBlood, iHigh, iScale, iNamed, iRaidBoss,
                                                iSetType, iInfection, iIsBossG, iIsBossBGM, iIsAreaBoss,
                                                iExp, iRepopNum, iRepopCount, iTargetType, iSpawnTime,
                                                iStartThink, iMontage, iIsManualSet, iPPDrop,
                                                iIsBloodOrbEnemy, iIsHighOrbEnemy, iDrops } = _rawEnemySchemas;
                                        if (iEnemyId >= 0 && hexId) row[iEnemyId] = hexId;
                                        row[iLv]    = cfg.lv;
                                        row[iBlood] = cfg.bloodOrbs;
                                        row[iHigh]  = cfg.highOrbs;
                                        if (iScale      >= 0) row[iScale]      = cfg.scale;
                                        if (iExp        >= 0) row[iExp]        = cfg.exp;
                                        if (iRepopNum   >= 0) row[iRepopNum]   = cfg.repopNum;
                                        if (iRepopCount >= 0) row[iRepopCount] = cfg.repopCount;
                                        if (iSetType    >= 0) row[iSetType]    = cfg.setType;
                                        if (iInfection  >= 0) row[iInfection]  = cfg.infection;
                                        if (iTargetType >= 0) row[iTargetType] = cfg.targetTypeId;
                                        if (iSpawnTime  >= 0) row[iSpawnTime]  = cfg.spawnTime;
                                        if (iNamed      >= 0) row[iNamed]      = cfg.namedId;
                                        if (iRaidBoss   >= 0) row[iRaidBoss]   = cfg.raidBossId;
                                        if (iIsBossG    >= 0) row[iIsBossG]    = cfg.isBossGauge;
                                        if (iIsBossBGM  >= 0) row[iIsBossBGM]  = cfg.isBossBGM;
                                        if (iIsAreaBoss >= 0) row[iIsAreaBoss] = cfg.isAreaBoss;
                                        if (iIsManualSet     >= 0) row[iIsManualSet]     = cfg.isManualSet;
                                        if (iIsBloodOrbEnemy >= 0) row[iIsBloodOrbEnemy] = cfg.isBloodOrbEnemy;
                                        if (iIsHighOrbEnemy  >= 0) row[iIsHighOrbEnemy]  = cfg.isHighOrbEnemy;
                                        if (iStartThink >= 0) row[iStartThink] = cfg.startThink;
                                        if (iMontage    >= 0) row[iMontage]    = cfg.montage;
                                        if (iPPDrop     >= 0) row[iPPDrop]     = cfg.ppDrop;
                                        row[iDrops] = cfg.dropsTableId;
                                    }
                                }
                            } else {
                                // ── Create new entry on empty node ────────────
                                if (!spawnKey) return;
                                const newEntry = {
                                    ...cfg, drops: [...(cfg.drops ?? [])],
                                    subGroupId: _activeSubGroupId ?? 0,
                                    hmPreset: cfg.emCode ? (hmPresetsByEmCode.get(cfg.emCode)?.id ?? 0) : 0,
                                    _rawIdx: -1,
                                };
                                if (_rawEnemyData && hexId) {
                                    const { iStage, iGroup, iPosIdx, iEnemyId, iLv, iBlood, iHigh,
                                            iSpawnTime, iDrops, iScale, iSubGroup, iNamed, iRaidBoss,
                                            iSetType, iInfection, iIsBossG, iIsBossBGM, iIsAreaBoss,
                                            iIsBloodOrbEnemy, iIsHighOrbEnemy,
                                            iExp, iRepopNum, iRepopCount, iTargetType,
                                            iHmPreset, iStartThink, iMontage, iIsManualSet, iPPDrop,
                                            iLayerNo } = _rawEnemySchemas;
                                    const [sid, gid, pidx] = spawnKey.split(',');
                                    const newRaw = new Array(_rawEnemyData.schemas.enemies.length).fill(null);
                                    newRaw[iStage]     = Number(sid);
                                    newRaw[iGroup]     = Number(gid);
                                    newRaw[iPosIdx]    = Number(pidx);
                                    if (iLayerNo >= 0) newRaw[iLayerNo] = 0;
                                    if (iEnemyId >= 0) newRaw[iEnemyId] = hexId;
                                    newRaw[iLv]        = cfg.lv;
                                    newRaw[iBlood]     = cfg.bloodOrbs;
                                    newRaw[iHigh]      = cfg.highOrbs;
                                    newRaw[iSpawnTime] = cfg.spawnTime ?? '00:00,23:59';
                                    newRaw[iDrops]     = cfg.dropsTableId;
                                    if (iScale      >= 0) newRaw[iScale]      = cfg.scale;
                                    if (iSubGroup   >= 0) newRaw[iSubGroup]   = _activeSubGroupId ?? 0;
                                    if (iNamed      >= 0) newRaw[iNamed]      = cfg.namedId;
                                    if (iRaidBoss   >= 0) newRaw[iRaidBoss]   = cfg.raidBossId;
                                    if (iSetType    >= 0) newRaw[iSetType]    = cfg.setType;
                                    if (iInfection  >= 0) newRaw[iInfection]  = cfg.infection;
                                    if (iIsBossG    >= 0) newRaw[iIsBossG]    = cfg.isBossGauge;
                                    if (iIsBossBGM  >= 0) newRaw[iIsBossBGM]  = cfg.isBossBGM;
                                    if (iIsAreaBoss >= 0) newRaw[iIsAreaBoss] = cfg.isAreaBoss;
                                    if (iIsManualSet     >= 0) newRaw[iIsManualSet]     = cfg.isManualSet;
                                    if (iIsBloodOrbEnemy >= 0) newRaw[iIsBloodOrbEnemy] = cfg.isBloodOrbEnemy;
                                    if (iIsHighOrbEnemy  >= 0) newRaw[iIsHighOrbEnemy]  = cfg.isHighOrbEnemy;
                                    if (iExp        >= 0) newRaw[iExp]        = cfg.exp;
                                    if (iRepopNum   >= 0) newRaw[iRepopNum]   = cfg.repopNum;
                                    if (iRepopCount >= 0) newRaw[iRepopCount] = cfg.repopCount;
                                    if (iTargetType >= 0) newRaw[iTargetType] = cfg.targetTypeId;
                                    if (iHmPreset   >= 0) newRaw[iHmPreset]   = newEntry.hmPreset;
                                    if (iStartThink >= 0) newRaw[iStartThink] = cfg.startThink;
                                    if (iMontage    >= 0) newRaw[iMontage]    = cfg.montage;
                                    if (iPPDrop     >= 0) newRaw[iPPDrop]     = cfg.ppDrop;
                                    _rawEnemyData.enemies.push(newRaw);
                                    newEntry._rawIdx = _rawEnemyData.enemies.length - 1;
                                }
                                let arr = cache.get(spawnKey);
                                if (!arr) { arr = []; cache.set(spawnKey, arr); }
                                arr.push(newEntry);
                                displayIdx = arr.length - 1;
                                if (_editMode && _renderEditPanel) _renderEditPanel();
                            }
                            if (_markDirty) _markDirty('ddon-src-spawns');
                            if (_rebuildOpenPopup) _rebuildOpenPopup();
                        } else if (action === 'remove-spawn') {
                            // Remove from cache
                            const arr = cache.get(spawnKey);
                            if (arr) {
                                arr.splice(displayIdx, 1);
                                if (!arr.length) cache.delete(spawnKey);
                                else displayIdx = Math.min(displayIdx, arr.length - 1);
                            }
                            // Remove from raw data
                            if (_rawEnemyData && !isNaN(rawIdx)) {
                                _rawEnemyData.enemies.splice(rawIdx, 1);
                                // Shift _rawIdx on all remaining entries
                                for (const entryArr of cache.values())
                                    for (const ent of entryArr)
                                        if (ent._rawIdx > rawIdx) ent._rawIdx--;
                            }
                            if (_markDirty) _markDirty('ddon-src-spawns');
                            marker.closePopup();
                            if (_editMode && _renderEditPanel) _renderEditPanel();
                        }
                    };
                    el.addEventListener('click', _popupClickHandler);

                    // ── Drag-drop: enemy panel → spawn popup ──────────────────
                    if (_editMode && spawnKey) {
                        _spawnPopupDropFn = (emCode) => {
                            // Convert emCode → raw hex EnemyId (e.g. 'em011200' → '0x011200')
                            const hexId = '0x' + emCode.slice(2).toUpperCase();
                            const newEntry = {
                                emCode, lv: 1, bloodOrbs: 0, highOrbs: 0,
                                spawnTime: null, drops: [], scale: 100,
                                subGroupId: _activeSubGroupId ?? 0, namedId: 2298, raidBossId: 0,
                                setType: 0, infection: 0, isBossGauge: false,
                                isBossBGM: false, isAreaBoss: false, isManualSet: false,
                                isBloodOrbEnemy: false, isHighOrbEnemy: false,
                                dropsTableId: -1,
                                exp: 0, repopNum: 0, repopCount: 0, targetTypeId: 1,
                                hmPreset: 0, startThink: 0, montage: 0, ppDrop: 0,
                                _rawIdx: -1,
                            };
                            if (_rawEnemyData) {
                                const { iStage, iGroup, iPosIdx, iEnemyId, iLv, iBlood, iHigh,
                                        iSpawnTime, iDrops, iScale, iSubGroup, iNamed, iRaidBoss,
                                        iSetType, iInfection, iIsBossG, iIsBossBGM, iIsAreaBoss,
                                        iIsBloodOrbEnemy, iIsHighOrbEnemy,
                                        iExp, iRepopNum, iRepopCount, iTargetType,
                                        iHmPreset, iStartThink, iMontage, iIsManualSet, iPPDrop,
                                        iLayerNo } = _rawEnemySchemas;
                                const [sid, gid, pidx] = spawnKey.split(',');
                                const newRaw = new Array(_rawEnemyData.schemas.enemies.length).fill(null);
                                newRaw[iStage]    = Number(sid);
                                newRaw[iGroup]    = Number(gid);
                                newRaw[iPosIdx]   = Number(pidx);
                                if (iLayerNo >= 0) newRaw[iLayerNo] = 0;
                                newRaw[iEnemyId]  = hexId;
                                newRaw[iLv]       = 1;
                                newRaw[iBlood]    = 0;
                                newRaw[iHigh]     = 0;
                                newRaw[iSpawnTime]= '00:00,23:59';
                                newRaw[iDrops]    = -1;
                                if (iScale       >= 0) newRaw[iScale]       = 100;
                                if (iSubGroup    >= 0) newRaw[iSubGroup]    = _activeSubGroupId ?? 0;
                                if (iNamed       >= 0) newRaw[iNamed]       = 2298;
                                if (iRaidBoss    >= 0) newRaw[iRaidBoss]    = 0;
                                if (iSetType     >= 0) newRaw[iSetType]     = 0;
                                if (iInfection   >= 0) newRaw[iInfection]   = 0;
                                if (iIsBossG     >= 0) newRaw[iIsBossG]     = false;
                                if (iIsBossBGM   >= 0) newRaw[iIsBossBGM]   = false;
                                if (iIsAreaBoss  >= 0) newRaw[iIsAreaBoss]  = false;
                                if (iIsManualSet     >= 0) newRaw[iIsManualSet]     = false;
                                if (iIsBloodOrbEnemy >= 0) newRaw[iIsBloodOrbEnemy] = false;
                                if (iIsHighOrbEnemy  >= 0) newRaw[iIsHighOrbEnemy]  = false;
                                if (iExp         >= 0) newRaw[iExp]         = 0;
                                if (iRepopNum    >= 0) newRaw[iRepopNum]    = 0;
                                if (iRepopCount  >= 0) newRaw[iRepopCount]  = 0;
                                if (iTargetType  >= 0) newRaw[iTargetType]  = 1;
                                if (iHmPreset    >= 0) newRaw[iHmPreset]    = hmPresetsByEmCode.get(emCode)?.id ?? 0;
                                if (iStartThink  >= 0) newRaw[iStartThink]  = 0;
                                if (iMontage     >= 0) newRaw[iMontage]     = 0;
                                if (iPPDrop      >= 0) newRaw[iPPDrop]      = 0;
                                _rawEnemyData.enemies.push(newRaw);
                                newEntry._rawIdx = _rawEnemyData.enemies.length - 1;
                            }
                            let arr = cache.get(spawnKey);
                            if (!arr) { arr = []; cache.set(spawnKey, arr); }
                            arr.push(newEntry);
                            displayIdx = arr.length - 1;
                            // ── Propagate emCode to spawn-set peers ───────────
                            if (_spawnSetMode) {
                                const peers = getPeers(cache);
                                for (const peer of peers) {
                                    for (const pEntry of peer.entries) {
                                        pEntry.emCode = emCode;
                                        if (_rawEnemyData && pEntry._rawIdx >= 0) {
                                            const pRow = _rawEnemyData.enemies[pEntry._rawIdx];
                                            const { iEnemyId } = _rawEnemySchemas;
                                            if (pRow && iEnemyId >= 0) pRow[iEnemyId] = hexId;
                                        }
                                    }
                                }
                            }
                            if (_markDirty) _markDirty('ddon-src-spawns');
                            const cd = el.querySelector('.leaflet-popup-content');
                            if (cd) {
                                cd.innerHTML = buildEnemyPopup(cache);
                                watchEditChanges(cd);
                                const view = cd.querySelector('.se-spawn-view');
                                if (view) {
                                    view.style.transition = 'background 0.1s';
                                    view.style.background = 'rgba(80,200,120,0.25)';
                                    setTimeout(() => { view.style.background = ''; }, 700);
                                }
                            }
                        };
                    } else {
                        _spawnPopupDropFn = null;
                    }
                });
            };
            if (_enemySpawnCache) { bind(_enemySpawnCache); return; }
            _enemySpawnPromise.then(cache => { if (this.isPopupOpen()) bind(cache); });
        });
        marker.on('popupclose', () => { _spawnPopupDropFn = null; _rebuildOpenPopup = null; _spawnSetConflictDismissed = false; hideNamedStatsPanel(); leafletMap.off('move', _repositionNamedStatsPanel); if (_activeRadiiMarker === marker) clearSpawnRadii(); });

        marker._sgKey          = sgKey;
        marker._label          = buildTooltip(_enemySpawnCache);
        marker._origStyle      = { color: isKeyBearer ? '#c8a000' : g.color, weight: isKeyBearer ? 3.5 : 2.5, opacity: 1, fillOpacity: 0.85 };
        marker._spawn          = spawn;
        marker._info           = info;
        marker._spawnKey       = spawnKey;
        marker._naturalLatLng  = latlng;    // saved for spread reset
        marker._rebuildTooltip = () => buildTooltip(_enemySpawnCache);
        marker._naturalTooltip = buildTooltip(_enemySpawnCache);

        if (!g.sgMarkers[sgKey]) g.sgMarkers[sgKey] = [];
        g.sgMarkers[sgKey].push(marker);

        // Style boss markers distinctly once cache is available
        const applyBossStyle = (cache) => {
            if (!spawnKey) return;
            const ents = cache.get(spawnKey) ?? [];
            if (ents.some(e => e.isBossGauge || e.isAreaBoss || e.raidBossId > 0)) {
                marker.setStyle({ color: '#ff3333', fillColor: '#cc0000', weight: 3, radius: 8 });
                marker._origStyle = { ...marker._origStyle, color: '#ff3333', fillColor: '#cc0000', weight: 3, radius: 8 };
            }
        };
        if (_enemySpawnCache) applyBossStyle(_enemySpawnCache);
        else _enemySpawnPromise.then(applyBossStyle).catch(() => {});

        marker
            .on('mouseover', function() { _highlightSG(this._sgKey); })
            .on('mouseout',  _unhighlightSG)
            .on('click',     function() { _radiiClickConsumed = true; showSpawnRadii(this); });
        layer.addLayer(marker);
    }

    g.detailsLayer = layer;
}

// _expandGroupCore / _collapseGroupCore do the state change without triggering
// reapplySpread or updateLayersInHash — used by bulk operations.
function _expandGroupCore(g, { skipFilter = false } = {}) {
    // Ignore stale expand jobs from a previous map (chunked applyMobDisplayMode).
    if (!_groupStore.has(g.groupId) || _groupStore.get(g.groupId) !== g) return;
    if (!g.detailsLayer) buildGroupDetails(g);
    else if (useLegacyGroupChips()) ensureGroupStructureLayers(g);
    const enemiesOn = (document.getElementById('layer-enemies')?.checked ?? true) && anyMobTypeEnabled();
    if (enemiesOn) g.detailsLayer.addTo(leafletMap);
    g.isExpanded = true;
    // Move chip to just above the topmost spawn so it doesn't cover any enemies.
    // Use the topmost spawn's own X so the chip stays directly above the hull, not the centroid.
    let topPx = g.pts[0][0], topPy = g.pts[0][1];
    for (const [px, py] of g.pts) { if (py > topPy) { topPy = py; topPx = px; } }
    g.labelMarker.setLatLng(xy(topPx, topPy + 10));
    g.labelMarker.setIcon(makeChipIcon(g.groupId, g.color, g.items.length, true, g.yOffset, g.isKeyBearerGroup, _groupHasBoss(g)));
    for (const [sgKey, markers] of Object.entries(g.sgMarkers)) {
        if (!_sgMarkers[sgKey]) _sgMarkers[sgKey] = [];
        _sgMarkers[sgKey].push(...markers);
    }
    if (!skipFilter) applySubGroupFilter();  // after setIcon so opacity isn't reset by icon replacement
}

function _collapseGroupCore(g) {
    if (g.detailsLayer) leafletMap.removeLayer(g.detailsLayer);
    removeGroupTerritoryRect(g);
    g.isExpanded = false;
    g.labelMarker.setLatLng(xy(g.centroid.px, g.centroid.py));
    g.labelMarker.setIcon(makeChipIcon(g.groupId, g.color, g.items.length, false, g.yOffset, g.isKeyBearerGroup, _groupHasBoss(g)));
    for (const sgKey of Object.keys(g.sgMarkers)) delete _sgMarkers[sgKey];
    if (_activeRadiiMarker && g.items.some(it => it.spawn === _activeRadiiMarker._spawn))
        clearSpawnRadii();
    // Re-apply chip visibility in case setIcon reset it
    _setGroupVisible(g, _activeSubGroupId === null ||
        g.items.some(({ spawn }) => _spawnSubGroupId(spawn) === _activeSubGroupId) ||
        (_activeSubGroupId === 1 && g.areaSpawn));
}

// ── SubGroup filter ────────────────────────────────────────────────────────────
// Network SubGroupId is derived from two mechanisms:
//
// 1. Kill-triggered (lot SubGroupNo):
//    SubGroupNo = -1  → SubGroupId = 0  (initial/always-present spawn)
//    SubGroupNo =  N  → SubGroupId = N+1 (triggered spawn; client requests N+1 after trigger fires)
//    Confirmed by PS4 disassembly: moveSetUnitSubGroupSever sends v5 = flagNo+1.
//
// 2. Area-triggered (GPB SetCondition1=1 + AreaHit=True, stored as g.areaSpawn):
//    When the player enters the KillArea zone, the client sends SubGroupId=1 directly.
//    The group's spawns all have SubGroupNo=-1, but are re-requested under SubGroupId=1.
//    Confirmed by server logs: Lestania G27 (areaSpawn=true) sends SubGroupId=1 on area entry.
//
// g.splitId (from GPB) is NOT used for filtering — it identifies spatial sections for PD dungeons.

const _spawnSubGroupId = (spawn) =>
    (spawn?.SubGroupNo == null || spawn.SubGroupNo === -1) ? 0 : spawn.SubGroupNo + 1;

function _computeAvailableSubGroups() {
    const counts = new Map();
    for (const g of _groupStore.values()) {
        for (const { spawn } of g.items) {
            const id = _spawnSubGroupId(spawn);
            counts.set(id, (counts.get(id) ?? 0) + 1);
        }
        if (g.areaSpawn) {
            const n = g.items.filter(({ spawn }) =>
                spawn.SubGroupNo == null || spawn.SubGroupNo === -1,
            ).length;
            if (n > 0) counts.set(1, (counts.get(1) ?? 0) + n);
        }
    }
    _availableSubGroups = [...counts.keys()].sort((a, b) => a - b);
    if (_activeSubGroupId != null && !_availableSubGroups.includes(_activeSubGroupId)) {
        _activeSubGroupId = null;
    }
    _renderSubGroupSelector();
}

function stageHasSpawnWaves() {
    return _availableSubGroups.length > 1;
}

function _setMarkerVisible(m, visible) {
    m._hidden = !visible;
    if (visible) {
        m.setStyle(m._origStyle);
        m.options.interactive = true;
        if (m._spokeLine) m._spokeLine.setStyle({ weight: 1, opacity: 0.4, dashArray: '3 3' });
    } else {
        m.setStyle({ opacity: 0, fillOpacity: 0 });
        m.options.interactive = false;
        if (m._spokeLine) m._spokeLine.setStyle({ opacity: 0 });
    }
}

function _setGroupVisible(g, visible) {
    const showChip = visible && useLegacyGroupChips();
    g.labelMarker.setOpacity(showChip ? 1 : 0);
    g.labelMarker.options.interactive = showChip;
    const chipEl = g.labelMarker.getElement();
    if (chipEl) chipEl.style.pointerEvents = showChip ? '' : 'none';
    // Hull / link polyline visibility (legacy group-chip mode only)
    syncGroupStructureVisibility(g, visible);
}

function applySubGroupFilter() {
    for (const g of _groupStore.values()) {
        const groupVisible = _activeSubGroupId === null ||
            g.items.some(({ spawn }) => _spawnSubGroupId(spawn) === _activeSubGroupId) ||
            (_activeSubGroupId === 1 && g.areaSpawn);
        _setGroupVisible(g, groupVisible);
        if (!g.isExpanded || !g.detailsLayer) continue;
        for (const m of g.detailsLayer.getLayers()) {
            if (!m._spawn) continue;
            // Area-spawn groups (g.areaSpawn) re-request their SubGroupNo=-1 spawns under SubGroupId=1
            const spawnIsAreaSpawnInitial = g.areaSpawn && (m._spawn?.SubGroupNo == null || m._spawn.SubGroupNo === -1);
            const subgroupVisible = _activeSubGroupId === null ||
                _spawnSubGroupId(m._spawn) === _activeSubGroupId ||
                (_activeSubGroupId === 1 && spawnIsAreaSpawnInitial);
            const labelVisible = spawnHasMobLabel(_enemySpawnCache, m._spawnKey, m._spawn);
            const typeVisible  = spawnMatchesTypeFilter(_enemySpawnCache, m._spawnKey, m._spawn);
            _setMarkerVisible(m, subgroupVisible && labelVisible && typeVisible);
        }
    }
    refreshMobTooltips();
    reapplySpread();
}

function _renderSubGroupSelector() {
    const bar = document.getElementById('subgroup-bar');
    if (!bar) return;
    if (!stageHasSpawnWaves()) {
        bar.style.display = 'none';
        bar.innerHTML = '';
        return;
    }
    bar.style.display = 'flex';
    const pill = (sg, label) => {
        const on = sg === _activeSubGroupId;
        return `<button type="button" class="sg-filter-btn${on ? ' active' : ''}" data-sg="${sg === null ? '' : sg}">${label}</button>`;
    };
    bar.innerHTML = `<span class="sg-filter-label">SubGroup</span>`
        + pill(null, 'All') + _availableSubGroups.map(sg => pill(sg, sg)).join('');
    bar.querySelectorAll('.sg-filter-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            _activeSubGroupId = btn.dataset.sg === '' ? null : Number(btn.dataset.sg);
            _renderSubGroupSelector();
            applySubGroupFilter();
        });
    });
}

function expandGroup(groupId) {
    const g = _groupStore.get(groupId);
    if (!g || g.isExpanded) return;
    _expandGroupCore(g);
    _updateExpandCollapseBtn();
    reapplySpread();
}

function collapseGroup(groupId) {
    const g = _groupStore.get(groupId);
    if (!g || !g.isExpanded) return;
    _collapseGroupCore(g);
    _updateExpandCollapseBtn();
    reapplySpread();
}

function toggleGroup(groupId) {
    const g = _groupStore.get(groupId);
    if (!g) return;
    if (g.isExpanded) collapseGroup(groupId); else expandGroup(groupId);
}

function _expandAllGroups() {
    for (const g of _groupStore.values()) if (!g.isExpanded) _expandGroupCore(g);
    _updateExpandCollapseBtn();
    reapplySpread();
}

function _collapseAllGroups() {
    for (const g of _groupStore.values()) if (g.isExpanded) _collapseGroupCore(g);
    _updateExpandCollapseBtn();
    reapplySpread();
}

function _updateExpandCollapseBtn() {
    const btn = document.getElementById('btn-expand-collapse');
    if (!btn) return;
    const anyCollapsed = [..._groupStore.values()].some(g => !g.isExpanded);
    btn.textContent = anyCollapsed ? 'Expand All' : 'Collapse All';
    updateLayersInHash();
}

// ── Enemy spawn markers ────────────────────────────────────────────────────────

// Registry of circleMarkers keyed by "sg:groupId", rebuilt on each loadEnemySpawns call.
// Used to highlight all markers sharing the same SG within the same group on hover.
let _sgMarkers = {};
let _unhighlightTimer = null;
let _highlightedSet   = new Set();  // all markers currently in highlighted state

function _restoreMarkerStyle(m) {
    if (m._hidden) {
        m.setStyle({ opacity: 0, fillOpacity: 0 });
        return;
    }
    m.setStyle(m._origStyle);
    m.setRadius(m._origStyle?.radius ?? 5);
}

function _clearHighlight() {
    for (const m of _highlightedSet) {
        _restoreMarkerStyle(m);
        m.closeTooltip();
        if (m._spreadAnchor) {
            m._spreadAnchor.setRadius(4);
            m._spreadAnchor.setStyle({ weight: 1.5, fillOpacity: 0.85 });
        }
        if (m._spokeLine) m._spokeLine.setStyle({ weight: 1, opacity: 0.4, dashArray: '3 3' });
    }
    _highlightedSet.clear();
}

function _applyHighlight(markers) {
    clearTimeout(_unhighlightTimer);
    _clearHighlight();                  // synchronously reset any previously lit markers
    for (const m of markers) {
        if (m._hidden) continue;
        m.setStyle({ weight: 4, fillOpacity: 1.0, color: '#ffffff' });
        m.setRadius(9);
        m.openTooltip();
        _highlightedSet.add(m);
        if (m._spreadAnchor) {
            m._spreadAnchor.setRadius(8);
            m._spreadAnchor.setStyle({ weight: 2.5, fillOpacity: 1.0 });
        }
        if (m._spokeLine) m._spokeLine.setStyle({ weight: 2.5, opacity: 1.0, dashArray: null });
    }
}

function _highlightSG(sgKey) {
    _applyHighlight((_sgMarkers[sgKey] || []).filter(m => !m._hidden));
}

function _unhighlightSG() {
    _unhighlightTimer = setTimeout(_clearHighlight, 160);
}

// ── Gathering group hover highlight ──────────────────────────────────────────
let _gatherHighlightTimer = null;
let _gatherHighlightedSet = new Set();   // markers currently highlighted

function _clearGatherHighlight() {
    for (const m of _gatherHighlightedSet) {
        const el = m.getElement()?.firstElementChild;
        if (!el) continue;
        el.style.outline = '';
        el.style.outlineOffset = '';
        el.style.boxShadow = '';
        if (el.tagName === 'IMG') el.style.filter = POI_IMG_FILTER;
    }
    _gatherHighlightedSet.clear();
}

function _applyGatherHighlight(markers) {
    clearTimeout(_gatherHighlightTimer);
    _clearGatherHighlight();
    for (const m of markers) {
        if (m._poiHidden) continue;
        const el = m.getElement()?.firstElementChild;
        if (!el) continue;
        if (el.tagName === 'IMG') {
            el.style.filter = POI_IMG_FILTER_HOVER;
        } else {
            el.style.outline = '2px solid #fff';
            el.style.outlineOffset = '2px';
            el.style.boxShadow = '0 0 6px 2px rgba(255,255,255,0.8)';
        }
        _gatherHighlightedSet.add(m);
    }
}

function _unhighlightGather() {
    _gatherHighlightTimer = setTimeout(_clearGatherHighlight, 160);
}

// ── Spawn aggro/link radius circles ───────────────────────────────────────────
let _activeRadiiMarker = null;   // the marker whose circles are currently shown

function clearSpawnRadii() {
    spawnRadiiLayer.clearLayers();
    _activeRadiiMarker = null;
}

function showSpawnRadii(marker) {
    if (!document.getElementById('layer-radii').checked) return;
    // Toggle off if clicking the same marker again
    if (_activeRadiiMarker === marker) {
        clearSpawnRadii();
        return;
    }
    spawnRadiiLayer.clearLayers();
    _activeRadiiMarker = marker;

    const spawn = marker._spawn;
    const info  = marker._info;
    if (!spawn || !info) return;

    // Resolve emCode: prefer the first EnemySpawn.json entry for this spawn,
    // fall back to the lot-file placeholder (EmName).
    const spawnEntries = marker._spawnKey && _enemySpawnCache
        ? (_enemySpawnCache.get(marker._spawnKey) ?? []) : [];

    // If the cache is loaded and no enemy is assigned to this slot, don't draw.
    const hasEnemy = !_enemySpawnCache || spawnEntries.some(e => !!e.lv);
    if (!hasEnemy) { clearSpawnRadii(); return; }

    const emCode = spawnEntries[0]?.emCode ?? spawn.EmName ?? null;
    const radii  = emCode ? (emRadii[emCode] ?? null) : null;

    // Convert world-unit radius → map CRS units (image pixels).
    // info.scale is pixels-per-world-unit (used for the X axis on all map types).
    const scale = info.scale;
    const latlng = marker.getLatLng();

    const aggroR = radii?.aggroRadius ?? spawn.AggroRadius;
    const linkR  = radii?.linkRadius  ?? spawn.LinkRadius;

    if (aggroR) {
        L.circle(latlng, {
            radius:      aggroR * scale,
            color:       '#ffd700',
            weight:      1.5,
            opacity:     0.9,
            fillColor:   '#ffd700',
            fillOpacity: 0.07,
            dashArray:   null,
            interactive: false,
        }).addTo(spawnRadiiLayer);
    }

    if (linkR) {
        L.circle(latlng, {
            radius:      linkR * scale,
            color:       '#ff7700',
            weight:      2,
            opacity:     0.9,
            fillColor:   '#ff7700',
            fillOpacity: 0.05,
            dashArray:   '6 4',
            interactive: false,
        }).addTo(spawnRadiiLayer);
    }
}

// Dismiss circles when clicking the map background (not a marker).
// _radiiClickConsumed is set by the marker click handler so the map-level
// click that Leaflet always fires afterwards doesn't immediately clear them.
let _radiiClickConsumed = false;
leafletMap.on('click', () => {
    if (_radiiClickConsumed) { _radiiClickConsumed = false; return; }
    clearSpawnRadii();
});

// Each distinct SpawnGroup value (0–255) gets its own deterministic fill colour.
// Same SpawnGroup value = same spawn condition = same colour everywhere on the map.
// Uses OKLCH so perceived brightness is uniform across all hues (unlike HSL where
// yellow looks nearly white and blue looks dark at the same L value).
function spawnGroupColor(sg) {
    const hue = (sg * 137) % 360;
    const L = sg < 50 ? 0.80 : 0.72;    // perceptual lightness (0–1)
    const C = sg < 50 ? 0.10 : 0.17;    // chroma: softer for common spawns
    return `oklch(${L} ${C} ${hue})`;
}

// Deterministic OKLCH colour from a file group ID (GroupNo).
// L=0.55 is darker than the tan dungeon backgrounds (~0.62) so outlines stand out,
// and still readable against the dark chip background.  C=0.13 is muted but distinct.
function groupBorderColor(groupId) {
    const hue = (groupId * 137) % 360;
    return `oklch(0.55 0.13 ${hue})`;
}

// ── Floor OBB test ────────────────────────────────────────────────────────────
// Returns the FloorId (= layer index) for a world position using the GMP OBBs,
// or null if no OBB contains the point.
// OBB test: translate to local space via dot products with the two XZ axes,
// then test each local coordinate against the half-extents.
// The Y axis is always world-vertical (M2=[0,1,0]), so |dy| <= ey directly.
function getEnemyFloor(worldX, worldY, worldZ, floorObbs) {
    for (const o of floorObbs) {
        const dx = worldX - o.cx;
        const dy = worldY - o.cy;
        const dz = worldZ - o.cz;
        const lx = dx * o.ax + dz * o.az;   // project onto local X axis
        const lz = dx * o.bx + dz * o.bz;   // project onto local Z axis
        if (Math.abs(lx) <= o.ex && Math.abs(dy) <= o.ey && Math.abs(lz) <= o.ez)
            return o.floor_id;
    }
    return null;
}


async function loadEnemySpawns(info, stid = null) {
    ++_mobDisplayModeJob; // cancel any in-flight group expansion from a prior map
    // Tear down all previous group state
    enemyLayer.clearLayers();
    for (const g of _groupStore.values()) {
        if (g.detailsLayer) leafletMap.removeLayer(g.detailsLayer);
    }
    _groupStore.clear();
    _sgMarkers = {};
    _spreadOverlay.clearLayers();
    territoryLayer.clearLayers();
    clearSpawnRadii();
    _activeSubGroupId = null;
    _availableSubGroups = [];
    _renderSubGroupSelector();
    _currentMapInfo   = info;
    _currentFloorObbs = info.floor_obbs ?? null;
    if (!info.stages?.length) return;

    const floorObbs     = _currentFloorObbs;
    const filterByFloor = floorObbs !== null;

    const stagesToLoad = (stid && info.stages.includes(stid)) ? [stid] : info.stages;
    await fetchEnemyPositionsStages(stagesToLoad.map((id) => parseInt(id.slice(2), 10)));

    // Collect all groups, merging across stages if multiple are loaded
    const byGroupId = new Map(); // groupId string → { territory, items:[{spawn,idx,sg,latlng}], pts:[] }
    for (const stageId of stagesToLoad) {
        const stageNo   = String(parseInt(stageId.slice(2), 10));
        const stageData = enemyPositionsForStage(stageNo);
        if (!stageData) continue;
        for (const [groupId, groupData] of Object.entries(stageData)) {
            const spawns         = groupData.spawns         ?? groupData;  // back-compat
            const territory      = groupData.territory      ?? null;
            const keyBearerGroup = groupData.keyBearerGroup ?? false;
            const splitId        = groupData.splitId        ?? 0;
            const areaSpawn      = groupData.areaSpawn      ?? false;
            const priority       = groupData.priority       ?? 0;
            if (!byGroupId.has(groupId)) byGroupId.set(groupId, { territory, keyBearerGroup, splitId, areaSpawn, priority, items: [], pts: [] });
            else if (keyBearerGroup) byGroupId.get(groupId).keyBearerGroup = true;
            const entry = byGroupId.get(groupId);
            for (let i = 0; i < spawns.length; i++) {
                const spawn = spawns[i];
                const pos   = spawn.Position;
                if (filterByFloor) {
                    const floor = getEnemyFloor(pos.x, pos.y, pos.z, floorObbs);
                    if (floor !== null && floor !== currentLayer) continue;
                }
                const latlng = worldToPixel(pos.x, pos.z, info);
                entry.pts.push([latlng.lng, latlng.lat]);
                entry.items.push({ spawn, idx: i, sg: spawn.SpawnGroup ?? 0, latlng, stageNo });
            }
        }
    }

    // Detect centroid collisions so stacked chips can be offset
    const centroidBuckets = new Map();
    for (const [groupId, { pts }] of byGroupId) {
        if (!pts.length) continue;
        const cx = pts.reduce((s, p) => s + p[0], 0) / pts.length;
        const cy = pts.reduce((s, p) => s + p[1], 0) / pts.length;
        const key = `${Math.round(cx)}:${Math.round(cy)}`;
        if (!centroidBuckets.has(key)) centroidBuckets.set(key, []);
        centroidBuckets.get(key).push(groupId);
    }

    // Create one chip label marker per group (collapsed by default)
    for (const [groupId, { territory, keyBearerGroup, splitId, areaSpawn, priority, items, pts }] of byGroupId) {
        if (!pts.length) continue;
        const color = groupBorderColor(parseInt(groupId, 10));
        const cx = pts.reduce((s, p) => s + p[0], 0) / pts.length;
        const cy = pts.reduce((s, p) => s + p[1], 0) / pts.length;

        // Offset chip vertically if multiple groups share the same centroid
        const bucketKey = `${Math.round(cx)}:${Math.round(cy)}`;
        const bucket    = centroidBuckets.get(bucketKey);
        const slotIdx   = bucket.indexOf(groupId);
        const yOffset   = 10 + slotIdx * 20;  // pixels below anchor

        const g = { groupId, color, territory, isKeyBearerGroup: keyBearerGroup, splitId, areaSpawn, priority, items, pts,
                    centroid: { px: cx, py: cy }, yOffset,
                    labelMarker: null, detailsLayer: null, isExpanded: false, sgMarkers: {} };
        _groupStore.set(groupId, g);

        // Check cache immediately (may already be loaded on map re-navigation); defer only if not yet available.
        const chipIcon    = makeChipIcon(groupId, color, items.length, false, yOffset, keyBearerGroup, _groupHasBoss(g));
        const labelMarker = L.marker(xy(cx, cy), { icon: chipIcon, zIndexOffset: 100 });
        g.labelMarker = labelMarker;

        labelMarker.on('click', (e) => { L.DomEvent.stopPropagation(e); toggleGroup(groupId); });
        labelMarker.addTo(enemyLayer);
    }

    _updateExpandCollapseBtn();
    _computeAvailableSubGroups();
    applySubGroupFilter();
    // Expanding hundreds of groups blocks the main thread — defer so the map image can paint.
    setTimeout(() => applyMobDisplayMode(), 0);
}

// ── Landmark markers ──────────────────────────────────────────────────────────
const LANDMARK_COLORS = {
    TYPE_DOOR:       '#ffd700',
    TYPE_CAVE:       '#ff8c00',
    TYPE_BASEMENT:   '#cd853f',
    TYPE_CATACOMB:   '#9b59b6',
    TYPE_ELF_RUIN:   '#1abc9c',
    TYPE_AREA_WARP:  '#00bcd4',
    TYPE_SHRINE:     '#ffffff',
    TYPE_OUTPOST:    '#4caf50',
    TYPE_WATER_LINE: '#4fc3f7',
    TYPE_WELL:       '#81d4fa',
    TYPE_TEXT:       '#888888',
    TYPE_NONE:       '#444444',
};

// Types that clutter the map without being useful landmarks
const HIDDEN_LANDMARK_TYPES = new Set(['TYPE_TEXT', 'TYPE_WATER_LINE', 'TYPE_NONE']);

// ── Map point (POI) category filters ───────────────────────────────────────────
const POI_FILTER_KEY = 'ddon-poi-filters';
const POI_ICON_BASE  = 'resources/poi-icons';
const MMAPICON_SLICE_BASE = 'resources/mmapicon_slices';
/** Confirmed mmapicon slice IDs (see resources/mmapicon_manifest.json). */
const LANDMARK_MMAPICON = {
    door:         8,
    house:        14,
    outpost:      13,
    cave:         15,
    well:         17,
    areaWarp:     48,
    pathway:      44,
    inn:          21,
    ark:          19,
    itemShop:     11,
    materialShop: 11,
    shop:         11,
};

/** poi-icons filename when a filter category has no mmapicon slice. */
const POI_ICON_FALLBACK = {
    gatherNode: 'sparkNode',
};

function poiIconSrc(categoryId, fallbackIcon) {
    const mmapId = LANDMARK_MMAPICON[categoryId];
    if (mmapId) return `${MMAPICON_SLICE_BASE}/${mmapId}.png`;
    const iconKey = fallbackIcon ?? POI_ICON_FALLBACK[categoryId] ?? categoryId;
    if (iconKey) return `${POI_ICON_BASE}/${iconKey}.png`;
    return null;
}

const POI_IMG_FILTER      = 'drop-shadow(0 0 1px #000) drop-shadow(0 1px 2px #000)';
const POI_IMG_FILTER_HOVER = `${POI_IMG_FILTER} brightness(1.3)`;
/** Purple tint for exchange / appraisal NPCs (SpecialShops.json). */
const POI_EXCHANGE_SHOP_IMG_FILTER = 'hue-rotate(285deg) saturate(1.45) brightness(1.1)';

function makePoiMapIcon(src, size = 24) {
    const half = size / 2;
    return L.divIcon({
        className: 'poi-map-marker',
        html: `<img src="${src}" width="${size}" height="${size}" alt="" `
            + `style="display:block;image-rendering:pixelated;filter:${POI_IMG_FILTER};">`,
        iconSize:    [size, size],
        iconAnchor:  [half, half],
        tooltipAnchor: [0, -half],
    });
}

/** Same shop mmapicon with purple ring + tint — still uses the shop POI filter. */
function makeExchangeShopMapIcon(src, size = 22) {
    const pad = 4;
    const outer = size + pad;
    const half = outer / 2;
    const imgFilter = `${POI_IMG_FILTER} ${POI_EXCHANGE_SHOP_IMG_FILTER}`;
    return L.divIcon({
        className: 'poi-map-marker poi-exchange-shop-marker',
        html: `<div style="position:relative;width:${outer}px;height:${outer}px;display:flex;align-items:center;justify-content:center;">`
            + `<div style="position:absolute;inset:0;border:2px solid #a855f7;border-radius:50%;`
            + `box-shadow:0 0 4px rgba(168,85,247,0.7);pointer-events:none;"></div>`
            + `<img src="${src}" width="${size}" height="${size}" alt="" `
            + `style="display:block;image-rendering:pixelated;filter:${imgFilter};">`
            + `</div>`,
        iconSize:    [outer, outer],
        iconAnchor:  [half, half],
        tooltipAnchor: [0, -half],
    });
}

// Connections that share a map point with a landmark use POI icons instead of diamonds.
const CONNECTION_POI_MATCH_DIST = { _default: 1200, door: 2500, cave: 1500, house: 1500 };
const connectionMatchDist = (poiCat) =>
    CONNECTION_POI_MATCH_DIST[poiCat] ?? CONNECTION_POI_MATCH_DIST._default;

const LANDMARK_ENTRANCE_TYPES = new Set([
    'TYPE_WELL', 'TYPE_CATACOMB', 'TYPE_CAVE', 'TYPE_SHRINE',
    'TYPE_BASEMENT', 'TYPE_DOOR', 'TYPE_ELF_RUIN',
]);
/** POI categories that represent the same map entrance as a landmark type. */
const POI_COVERS_LANDMARK = {
    cave:     ['TYPE_CAVE', 'TYPE_CATACOMB', 'TYPE_BASEMENT', 'TYPE_SHRINE', 'TYPE_ELF_RUIN'],
    basement: ['TYPE_BASEMENT', 'TYPE_CATACOMB', 'TYPE_CAVE'],
    well:     ['TYPE_WELL'],
    door:     ['TYPE_DOOR'],
    house:    ['TYPE_DOOR'],
};

const isGenericLandmark = (lm) => !lm.spot_name_en?.trim();

const landmarkMatchDist = (lm) => {
    const poiCat = LANDMARK_TYPE_TO_POI[lm.type];
    return poiCat ? connectionMatchDist(poiCat) : CONNECTION_POI_MATCH_DIST._default;
};

const poiCoversLandmarkType = (poiCat, lmType) =>
    (POI_COVERS_LANDMARK[_normalizePoiFilterCategory(poiCat) ?? poiCat] ?? []).includes(lmType);
const LANDMARK_TYPE_TO_POI = {
    TYPE_CAVE:     'cave',
    TYPE_DOOR:     'door',
    TYPE_BASEMENT: 'basement',
    TYPE_CATACOMB: 'catacomb',
    TYPE_SHRINE:   'shrine',
    TYPE_WELL:     'well',
    TYPE_ELF_RUIN: 'elfRuin',
};

const SPECIAL_CONNECTION_RULES = specialConnectionRules.connections ?? [];
const SPECIAL_CONNECTION_NAME_PATTERNS = (specialConnectionRules.name_patterns ?? [])
    .map((pattern) => new RegExp(pattern, 'i'));

// Quest-specific Ark variants; st0573 / st0574 / st0575 are the usual entrances.
const SUPPRESSED_ARK_DEST_STAGES = new Set([571, 572, 576]);
const isSuppressedArkConnection = (conn) =>
    conn.to_stage != null && SUPPRESSED_ARK_DEST_STAGES.has(conn.to_stage);

const matchSpecialConnection = (mapName, conn) => {
    for (const rule of SPECIAL_CONNECTION_RULES) {
        if (rule.map !== mapName) continue;
        if (rule.from_stage != null && conn.from_stage !== rule.from_stage) continue;
        if (rule.to_stage != null && conn.to_stage !== rule.to_stage) continue;
        if (rule.to_map != null && conn.to_map !== rule.to_map) continue;
        return rule;
    }
    const name = conn.name_en?.trim();
    if (name && SPECIAL_CONNECTION_NAME_PATTERNS.some((re) => re.test(name))) {
        return { id: 'name-pattern' };
    }
    return null;
};

const connectionDisplayName = (conn) => {
    const name = conn.name_en?.trim();
    return name || `Stage ${conn.to_stage}`;
};

const connectionExactCoordKey = (conn) =>
    (conn.x != null && conn.z != null) ? `${conn.x}|${conn.z}` : null;

const connectionDestKey = (conn) => `${conn.to_map ?? ''}|${conn.to_stage ?? ''}`;

function connectionChoiceVisible(choice) {
    return _isConnectionVisible(choice.poiCat, !!choice.specialRule);
}

function navigateConnectionChoice(sourceMap, choice) {
    if (!choice.hasMap) return;
    navigateTo(
        choice.navMap,
        choice.stageId,
        arrivalView(sourceMap, choice.navMap, choice.fromStage ?? null),
    );
}

function buildConnectionChoicePopup(sourceMap, choices, marker) {
    const wrap = document.createElement('div');
    wrap.className = 'conn-choice-popup';
    const title = document.createElement('div');
    title.className = 'conn-choice-title';
    title.textContent = 'Choose destination';
    wrap.appendChild(title);
    for (const choice of choices) {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'conn-choice-btn';
        const label = document.createElement('span');
        label.className = 'conn-choice-label';
        label.textContent = connectionDisplayName(choice.conn);
        btn.appendChild(label);
        const sub = document.createElement('span');
        sub.className = 'conn-choice-sub';
        sub.textContent = choice.hasMap ? choice.stageId : `No map (${choice.stageId})`;
        btn.appendChild(sub);
        btn.addEventListener('click', (e) => {
            L.DomEvent.stopPropagation(e);
            if (choice.hasMap) navigateConnectionChoice(sourceMap, choice);
            marker.closePopup();
        });
        wrap.appendChild(btn);
    }
    return wrap;
}

function addConnectionMapMarker(sourceMap, latlng, choices) {
    const visibleChoices = () => choices.filter(connectionChoiceVisible);
    const labels = [...new Set(choices.map((c) => connectionDisplayName(c.conn)))];
    const primary = choices[0];
    const poiCat = choices.length > 1 ? 'outpost' : (primary.poiCat ?? null);
    const icon = makeConnectionMarkerIcon(poiCat);
    const tooltip = labels.length === 1 ? labels[0] : labels.join(' · ');

    const marker = L.marker(latlng, { icon });
    marker.bindTooltip(tooltip, { permanent: false, direction: 'top', offset: [0, -10] });
    marker.on('click', (e) => {
        L.DomEvent.stopPropagation(e);
        const vis = visibleChoices();
        if (!vis.length) return;
        if (vis.length === 1 && !vis[0].hasMap) {
            marker.bindPopup(`No map data for Stage ${vis[0].conn.to_stage}<br>${vis[0].destName}`).openPopup();
            return;
        }
        if (vis.length === 1) {
            navigateConnectionChoice(sourceMap, vis[0]);
            return;
        }
        marker.bindPopup(buildConnectionChoicePopup(sourceMap, vis, marker), { maxWidth: 300 }).openPopup();
    });
    marker.addTo(connectionLayer);
    marker._poiCategory = poiCat;
    marker._poiIsConnection = true;
    marker._poiConnectionChoices = choices.map((c) => ({
        poiCat: c.poiCat,
        isSpecial: !!c.specialRule,
    }));
    _applyPoiMarkerVisibility(marker);
    return marker;
}

const isCaveLikeConnectionName = (name) => {
    if (!name) return false;
    const lower = name.toLowerCase();
    if (/\bterrace\b|\bsecret lodge\b/i.test(lower)) return true;
    if (/\b(inn|shrine|chapel|well|house|fort|castle|temple|palace|cellar|basement|drain)\b/.test(lower)) return false;
    if (/\blodge\b/.test(lower)) return false;
    return /\b(caves?|cavern|grotto|tunnels?|labyrinth|crypt|\bdens?\b|hollow|quarry|mines?|barrows?|\bold route\b|\bsecret route\b|\bsecret passage\b|guardhouse|storehouse|floodway|waterway|ruins|laboratory|mausoleum|columbarium|cemetery|penitentiary|burial chamber)\b/.test(lower);
};

function finalizeConnectionPoi(poiCat) {
    if (poiCat === 'catacomb' || poiCat === 'shrine' || poiCat === 'elfRuin') return 'cave';
    return poiCat;
}

function normalizeShopName(name) {
    return String(name ?? '').toLowerCase().replace(/['']/g, '').replace(/\s+/g, ' ').trim();
}

let _interiorShopNamesByStage = null;
function interiorShopNamesByStage() {
    if (_interiorShopNamesByStage) return _interiorShopNamesByStage;
    _interiorShopNamesByStage = new Map();
    const add = (stageNo, label) => {
        const norm = normalizeShopName(label);
        if (!norm) return;
        if (!_interiorShopNamesByStage.has(stageNo)) _interiorShopNamesByStage.set(stageNo, new Set());
        _interiorShopNamesByStage.get(stageNo).add(norm);
    };
    for (const [stageNo, npcs] of Object.entries(npcShops)) {
        if (!isAccessibleShopStage(stageNo)) continue;
        for (const npc of npcs) {
            const npcName = npcShopDisplayName(npc.NpcId);
            const funcLabel = NPC_FUNC_LABELS[npc.InstitutionFunctionId] ?? '';
            add(stageNo, resolveShopLabel(stageNo, npcName, funcLabel));
        }
    }
    for (const [stageNo, npcs] of Object.entries(npcSpecialShops)) {
        if (!isAccessibleShopStage(stageNo)) continue;
        for (const npc of npcs) {
            const npcName = npcShopDisplayName(npc.NpcId, npc.NpcName);
            const shopType = npc.ShopTypeName
                ? splitPascalCase(String(npc.ShopTypeName).replace(/Exchange/g, ' Exchange '))
                : 'Appraisals';
            add(stageNo, resolveShopLabel(stageNo, npcName, shopType));
        }
    }
    return _interiorShopNamesByStage;
}

/** Field entrance to a building whose shop NPC inside uses the same name. */
function connectionHasInteriorShopMatch(conn) {
    if (conn.to_stage == null) return false;
    const label = conn.name_en?.trim();
    if (!label) return false;
    const names = interiorShopNamesByStage().get(String(conn.to_stage));
    return names?.has(normalizeShopName(label)) ?? false;
}

function applyInteriorShopEntranceOverride(conn, poiCat) {
    if (poiCat === 'shop' && connectionHasInteriorShopMatch(conn)) return 'house';
    return poiCat;
}

function classifyConnectionPoi(conn, mapName) {
    const name = conn.name_en ?? '';
    const lower = name.toLowerCase();

    if (/\bbowmaster's secret lodge\b/i.test(lower)) return applyInteriorShopEntranceOverride(conn, 'cave');
    if (/\bterrace\b/i.test(lower)) return applyInteriorShopEntranceOverride(conn, 'cave');
    if (/\binn\b/i.test(name)) return applyInteriorShopEntranceOverride(conn, 'inn');
    if (/\bwell\b/i.test(lower)) return applyInteriorShopEntranceOverride(conn, 'well');
    if (/\bshrine\b|\bchapel\b/i.test(lower)) return applyInteriorShopEntranceOverride(conn, 'cave');
    if (/\bbasement\b|\bcellar\b|\bdrain(s|age)?\b|\bflood control\b/i.test(lower)) return applyInteriorShopEntranceOverride(conn, 'basement');
    if (/\b(underground|channel)\b/i.test(lower) && /\b(waterway|passage|drain|flood)\b/i.test(lower)) return applyInteriorShopEntranceOverride(conn, 'basement');
    if (/\bcatacombs?\b/i.test(lower)) return applyInteriorShopEntranceOverride(conn, 'cave');
    if (/\bthe second ark\b|\bthe third ark\b|\bthe ark\b/i.test(lower)) return applyInteriorShopEntranceOverride(conn, 'ark');
    if (/\bbitterblack\b|\bmaze cove\b/i.test(lower)) return applyInteriorShopEntranceOverride(conn, 'areaWarp');
    if (/^lestania$/i.test(name.trim())) return applyInteriorShopEntranceOverride(conn, 'door');
    if (/^lb\d+_m\d+$/.test(mapName) && conn.to_map && /^field\d+_m\d+$/.test(conn.to_map)) return applyInteriorShopEntranceOverride(conn, 'outpost');
    if (/\bshop\b|\bstore(?!house)\b|\bbazaar\b|\bmarket\b|\boutfitter\b|\bworkshop\b|\btrading post\b|\btrading company\b/i.test(lower)) return applyInteriorShopEntranceOverride(conn, 'shop');
    if (/^house$/i.test(name.trim()) || /\bhouse in the\b|\bhermit's house\b|\bkeeper's house\b|\bchief's home\b|\balchemist's home\b|\bivan's lodge\b/i.test(lower)) return applyInteriorShopEntranceOverride(conn, 'house');
    if (/\bfort\b|\bcastle\b|\bgate\b|\bdoor\b/i.test(lower)) return applyInteriorShopEntranceOverride(conn, 'door');
    if (isCaveLikeConnectionName(name)) return applyInteriorShopEntranceOverride(conn, 'cave');

    const landmarks = landmarkData[mapName];
    if (!landmarks || conn.x == null || conn.z == null) return applyInteriorShopEntranceOverride(conn, null);
    for (const lm of landmarks) {
        const poiCat = LANDMARK_TYPE_TO_POI[lm.type];
        if (!poiCat) continue;
        if (Math.hypot(lm.x - conn.x, lm.z - conn.z) >= connectionMatchDist(poiCat)) continue;
        return applyInteriorShopEntranceOverride(conn, finalizeConnectionPoi(poiCat));
    }
    return applyInteriorShopEntranceOverride(conn, null);
}

// Per-map cache — avoids re-classifying every connection for each landmark (O(n²) on load).
let _connectionPoiCacheMap = null;
let _connectionPoiCache    = null;

function resetConnectionPoiCache(mapName) {
    _connectionPoiCacheMap = mapName;
    _connectionPoiCache = new Map();
    const conns = connectionData[mapName];
    if (!conns) return;
    for (const conn of conns) {
        if (conn.x == null || conn.z == null) continue;
        _connectionPoiCache.set(conn, classifyConnectionPoi(conn, mapName));
    }
}

function getConnectionPoi(conn, mapName) {
    if (_connectionPoiCacheMap !== mapName || !_connectionPoiCache) resetConnectionPoiCache(mapName);
    return _connectionPoiCache.get(conn) ?? null;
}

function isLandmarkCoveredByConnection(mapName, lm) {
    const conns = connectionData[mapName];
    if (!conns) return false;
    if (_connectionPoiCacheMap !== mapName || !_connectionPoiCache) resetConnectionPoiCache(mapName);
    for (const conn of conns) {
        if (conn.x == null || conn.z == null) continue;
        const dist = Math.hypot(lm.x - conn.x, lm.z - conn.z);
        const poiCat = _connectionPoiCache.get(conn);

        if (poiCat) {
            if (dist < connectionMatchDist(poiCat) && poiCoversLandmarkType(poiCat, lm.type))
                return true;
        }

        // Unlabeled landmark layer (tooltip shows "WELL", "CATACOMB", …) — hide when a
        // named stage connection sits on the same entrance (even if orange diamond).
        if (isGenericLandmark(lm) && LANDMARK_ENTRANCE_TYPES.has(lm.type) &&
            conn.name_en?.trim() && dist < landmarkMatchDist(lm)) {
            return true;
        }

        // Decorative water-line overlay at the same entrance (e.g. Kinoza Drains).
        if (lm.type === 'TYPE_WATER_LINE' && dist < 100) return true;
    }
    return false;
}

const POI_LOCATION_CATEGORIES = [
    { id: 'cave',       label: 'Cave',               icon: 'cave' },
    { id: 'basement',   label: 'Basement',           icon: 'basement' },
    { id: 'house',      label: 'House',              icon: 'house' },
    { id: 'well',       label: 'Well',               icon: 'well' },
    { id: 'areaWarp',   label: 'Area warp',          icon: 'areaWarp' },
    { id: 'outpost',    label: 'Outpost',            icon: 'outpost' },
    { id: 'door',       label: 'Door',               icon: 'door' },
    { id: 'inn',        label: 'Inn',                icon: 'inn' },
    { id: 'shop',         label: 'Shops & appraisals', icon: 'shop' },
    { id: 'special',      label: 'Special' },
];
const POI_GATHER_CATEGORIES = [
    { id: 'mushroom',   label: 'Mushrooms',                   icon: 'mushroom' },
    { id: 'treasure',   label: 'Treasure chests',             icon: 'treasure' },
    { id: 'box',        label: 'Boxes',                       icon: 'box' },
    { id: 'antique',    label: 'Antiques',                    icon: 'antique' },
    { id: 'grassHerb',  label: 'Grass / herb',                icon: 'grassHerb' },
    { id: 'flower',     label: 'Flower',                      icon: 'flower' },
    { id: 'sand',       label: 'Sand',                        icon: 'sand' },
    { id: 'shell',      label: 'Shell',                       icon: 'shell' },
    { id: 'crystal',    label: 'Crystal',                     icon: 'gemstone' },
    { id: 'gemstone',   label: 'Gemstone',                    icon: 'crystal' },
    { id: 'gatherNode', label: 'Spark',                       icon: 'sparkNode' },
    { id: 'water',      label: 'Water',                       icon: 'water' },
    { id: 'lumber',     label: 'Lumber',                      icon: 'lumber' },
    { id: 'oneOff',     label: 'Off nodes',                   icon: 'oneOff' },
];
const POI_CATEGORIES = [...POI_LOCATION_CATEGORIES, ...POI_GATHER_CATEGORIES];
const LANDMARK_POI_CATEGORIES = new Set([
    'areaWarp', 'outpost', 'door', 'house', 'cave', 'basement', 'catacomb', 'elfRuin', 'shrine', 'well', 'ark', 'landmarkOther',
]);
const GATHER_POI_CATEGORIES = new Set([
    'mushroom', 'treasure', 'box', 'antique', 'grassHerb', 'flower', 'sand', 'shell', 'crystal', 'gemstone', 'gatherNode', 'water', 'lumber', 'oneOff', 'gatherOther',
]);
const POI_FILTER_DEFAULTS = Object.fromEntries(POI_CATEGORIES.map(c => [c.id, true]));
POI_FILTER_DEFAULTS.special = false;

const loadPoiFilters = () => {
    try {
        const raw = localStorage.getItem(POI_FILTER_KEY);
        const parsed = raw ? JSON.parse(raw) : {};
        // Migrate renamed gather category (chest → treasure)
        if ('chest' in parsed && !('treasure' in parsed)) parsed.treasure = parsed.chest;
        delete parsed.chest;
        // Merge split shop toggles into one
        if ('itemShop' in parsed || 'materialShop' in parsed) {
            parsed.shop = parsed.shop !== false &&
                parsed.itemShop !== false && parsed.materialShop !== false;
            delete parsed.itemShop;
            delete parsed.materialShop;
        }
        // Merge spark + dragon into gathering nodes
        if ('sparkNode' in parsed || 'dragon' in parsed) {
            parsed.gatherNode = parsed.gatherNode !== false &&
                parsed.sparkNode !== false && parsed.dragon !== false;
            delete parsed.sparkNode;
            delete parsed.dragon;
        }
        try {
            if (!localStorage.getItem('ddon-poi-split-crystal-gemstone')) {
                // JWL nodes previously shared the combined `crystal` filter toggle.
                if ('crystal' in parsed) parsed.gemstone = parsed.crystal;
                localStorage.setItem('ddon-poi-split-crystal-gemstone', '1');
            }
        } catch { /* ignore */ }
        // Old pathways / connections layer → all location filters off
        if (parsed.pathway === false || parsed.entrance === false) {
            for (const c of POI_LOCATION_CATEGORIES) parsed[c.id] = false;
        }
        delete parsed.pathway;
        if ('entrance' in parsed) delete parsed.entrance;
        try {
            if (!localStorage.getItem('ddon-poi-migrated-connections')) {
                let connOff = false;
                const layerRaw = localStorage.getItem(LAYER_PREFS_KEY);
                if (layerRaw && JSON.parse(layerRaw).connections === false) connOff = true;
                const { layers: urlLayers } = parseHash();
                if (urlLayers?.connections === false) connOff = true;
                if (connOff) {
                    for (const c of POI_LOCATION_CATEGORIES) parsed[c.id] = false;
                }
                localStorage.setItem('ddon-poi-migrated-connections', '1');
            }
        } catch { /* ignore */ }
        if ('eventDungeon' in parsed && !('special' in parsed)) parsed.special = parsed.eventDungeon;
        delete parsed.eventDungeon;
        if (parsed.appraisal === false) parsed.shop = false;
        delete parsed.appraisal;
        // Removed categories fold into cave
        if (parsed.catacomb === false || parsed.shrine === false || parsed.elfRuin === false)
            parsed.cave = false;
        delete parsed.catacomb;
        delete parsed.shrine;
        delete parsed.elfRuin;
        delete parsed.ark;
        // inn is now a first-class location filter
        return raw || Object.keys(parsed).length
            ? { ...POI_FILTER_DEFAULTS, ...parsed }
            : { ...POI_FILTER_DEFAULTS };
    } catch { return { ...POI_FILTER_DEFAULTS }; }
};

const savePoiFilters = () => {
    try { localStorage.setItem(POI_FILTER_KEY, JSON.stringify(_poiFilters)); } catch { /* ignore */ }
};

let _poiFilters = loadPoiFilters();

function classifyLandmarkPoiCategory(type, lm = null) {
    const spot = lm?.spot_name_en?.trim().toLowerCase();
    if (spot === 'the ark' || spot === 'the second ark' || spot === 'the third ark') return 'areaWarp';
    switch (type) {
        case 'TYPE_AREA_WARP':  return 'areaWarp';
        case 'TYPE_OUTPOST':    return 'outpost';
        case 'TYPE_DOOR':       return 'door';
        case 'TYPE_CAVE':       return 'cave';
        case 'TYPE_BASEMENT':   return 'basement';
        case 'TYPE_CATACOMB':
        case 'TYPE_ELF_RUIN':
        case 'TYPE_SHRINE':     return 'cave';
        case 'TYPE_WELL':       return 'well';
        default:                return 'landmarkOther';
    }
}

function classifyGatherPoiCategory(type) {
    if (type === 'OM_GATHER_MUSHROOM') return 'mushroom';
    if (type === 'OM_GATHER_GRASS') return 'grassHerb';
    if (type === 'OM_GATHER_FLOWER') return 'flower';
    if (type === 'OM_GATHER_SAND') return 'sand';
    if (type === 'OM_GATHER_SHELL') return 'shell';
    if (/^OM_GATHER_CRST_LV/.test(type)) return 'crystal';
    if (/^OM_GATHER_JWL_LV/.test(type)) return 'gemstone';
    if (type === 'OM_GATHER_TWINKLE' || type === 'OM_GATHER_DRAGON' || type === 'OM_GATHER_CORPSE') return 'gatherNode';
    if (type === 'OM_GATHER_WATER') return 'water';
    if (/^OM_GATHER_TREE_LV/.test(type)) return 'lumber';
    if (type === 'OM_GATHER_BOX') return 'box';
    if (type === 'OM_GATHER_ANTIQUE') return 'antique';
    if (type === 'OM_GATHER_ONE_OFF') return 'oneOff';
    if (type.startsWith('CHEST_') || type.startsWith('OM_GATHER_TREA_') ||
        type.startsWith('OM_GATHER_KEY_')) return 'treasure';
    return 'gatherOther';
}

function classifyShopPoiCategory(_funcId) {
    return 'shop';
}

/** Map legacy / internal POI ids to sidebar filter keys. */
function _normalizePoiFilterCategory(category) {
    if (!category) return category;
    if (category === 'itemShop' || category === 'materialShop' || category === 'appraisal') return 'shop';
    if (category === 'sparkNode' || category === 'dragon' || category === 'corpse') return 'gatherNode';
    if (category === 'ark') return 'areaWarp';
    if (category === 'catacomb' || category === 'shrine' || category === 'elfRuin') return 'cave';
    return category;
}

function poiFilterIconSrc(filterCategoryId) {
    const catDef = POI_CATEGORIES.find(c => c.id === filterCategoryId);
    return poiIconSrc(filterCategoryId, catDef?.icon);
}

function poiMapIconSrc(internalCategory, { uncategorizedConnection = false } = {}) {
    if (uncategorizedConnection || internalCategory == null) return poiFilterIconSrc('door');
    if (internalCategory === 'ark') return poiIconSrc('ark');
    const filterCat = _normalizePoiFilterCategory(internalCategory) ?? internalCategory;
    return poiFilterIconSrc(filterCat);
}

function gatherMapIconSrc(gatherType) {
    const cat = classifyGatherPoiCategory(gatherType);
    const catDef = POI_GATHER_CATEGORIES.find(c => c.id === cat);
    return poiIconSrc(cat, catDef?.icon ?? cat);
}

function makeConnectionMarkerIcon(poiCat) {
    return makePoiMapIcon(poiMapIconSrc(poiCat, { uncategorizedConnection: !poiCat }), 22);
}

// ── Named location search (map sidebar — portcrystals, caves, inns, etc.) ───────
/** Main overworld field maps only — excludes sfield* spot slices (wrong projection). */
const isMainWorldFieldMap = (mapName) => /^field\d+_m\d+$/.test(mapName);

const NAMED_LOCATION_CATEGORIES = new Set(POI_LOCATION_CATEGORIES.map(c => c.id));
const NAMED_LOCATION_RESULT_CAP = 30;

/** Prefer portcrystal outpost over area-warp duplicate sharing the same spot_id. */
const LANDMARK_SEARCH_TYPE_RANK = {
    TYPE_OUTPOST:    0,
    TYPE_AREA_WARP:  1,
    TYPE_DOOR:       2,
    TYPE_CAVE:       3,
    TYPE_WELL:       4,
    TYPE_BASEMENT:   5,
    TYPE_SHRINE:     6,
    TYPE_CATACOMB:   7,
    TYPE_ELF_RUIN:   8,
};
const landmarkSearchTypeRank = (type) => LANDMARK_SEARCH_TYPE_RANK[type] ?? 99;

function stageIdFromNo(stageNo) {
    if (stageNo == null) return null;
    return `st${String(stageNo).padStart(4, '0')}`;
}

function defaultFieldStage(mapName) {
    const info = mapParams[mapName];
    if (!info?.stages?.length) return null;
    if (info.stages.includes('st0100')) return 'st0100';
    return info.stages[0];
}

function mapAreaLabel(mapName) {
    const info = mapParams[mapName];
    const area = info?.quest_area_name;
    if (area && area !== 'Unknown') return area;
    if (info?.name_en) return splitPascalCase(info.name_en);
    return mapName;
}

function locationTypeLabel(categoryId) {
    const norm = _normalizePoiFilterCategory(categoryId) ?? categoryId;
    return POI_LOCATION_CATEGORIES.find(c => c.id === norm)?.label ?? 'Location';
}

let _namedLocationIndex = null;
let _namedLocationIndexVersion = 0;
const NAMED_LOCATION_INDEX_VERSION = 17;
let _pendingNamedLocNav = null;

function namedLocationMatchesEntry(entry, term, exact) {
    const label = entry.label.toLowerCase();
    if (exact) return label.startsWith(term);
    if (label.startsWith(term)) return true;
    const esc = term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const wordRe = new RegExp(`\\b${esc}\\b`, 'i');
    if (wordRe.test(entry.label)) return true;
    if (entry.labelAlt && wordRe.test(entry.labelAlt)) return true;
    return false;
}

function namedLocationEntryIconSrc(entry) {
    return poiMapIconSrc(entry.poiCat, { uncategorizedConnection: entry.isConnection && !entry.poiCat });
}

let _stageNoToMap = null;
function stageNoToMapLookup() {
    if (_stageNoToMap) return _stageNoToMap;
    _stageNoToMap = new Map();
    for (const [mapName, info] of Object.entries(mapParams)) {
        if (!info.img_exists || !info.stages?.length) continue;
        for (const stageId of info.stages) {
            _stageNoToMap.set(String(parseInt(stageId.slice(2), 10)), { mapName, info, stageId });
        }
    }
    return _stageNoToMap;
}

function npcShopDisplayName(npcId, fallback = '') {
    const raw = npcNames[String(npcId)];
    if (typeof raw === 'string') return raw;
    return raw?.name ?? fallback;
}

let _stageOfficialNames = null;
function stageOfficialNameLookup() {
    if (_stageOfficialNames) return _stageOfficialNames;
    _stageOfficialNames = new Map();
    for (const entry of stageList.StageListInfoList ?? []) {
        const name = entry.StageName?.En?.trim();
        if (name) _stageOfficialNames.set(String(entry.StageNo), name);
    }
    return _stageOfficialNames;
}

let _shopCountByStage = null;
function shopCountByStage() {
    if (_shopCountByStage) return _shopCountByStage;
    _shopCountByStage = new Map();
    for (const [stageNo, npcs] of Object.entries(npcShops)) {
        if (!isAccessibleShopStage(stageNo)) continue;
        _shopCountByStage.set(stageNo, (_shopCountByStage.get(stageNo) ?? 0) + npcs.length);
    }
    for (const [stageNo, npcs] of Object.entries(npcSpecialShops)) {
        if (!isAccessibleShopStage(stageNo)) continue;
        _shopCountByStage.set(stageNo, (_shopCountByStage.get(stageNo) ?? 0) + npcs.length);
    }
    return _shopCountByStage;
}

const GENERIC_SHOP_PLACE_NAMES = new Set([
    'house', 'lestania', 'cave harbor', 'hidden area', 'breya coast',
]);

function isDedicatedShopStageName(name) {
    if (!name || GENERIC_SHOP_PLACE_NAMES.has(name.toLowerCase())) return false;
    return /\b(shop|store|workshop|bazaar|market|outfitter|trading|merchant|appraisal|inn|lodge|general store|arms and armor)\b/i.test(name);
}

function possessiveNpcName(npcName) {
    const trimmed = npcName?.trim();
    if (!trimmed) return '';
    return /s$/i.test(trimmed) ? `${trimmed}'` : `${trimmed}'s`;
}

function npcShopLabel(npcName, funcLabel) {
    if (!npcName) return funcLabel || 'Shop';
    const pos = possessiveNpcName(npcName);
    if (!funcLabel || funcLabel === 'Shop') return `${pos} Shop`;
    return `${pos} ${funcLabel}`;
}

function resolveShopLabel(stageNo, npcName, funcLabel) {
    const official = stageOfficialNameLookup().get(stageNo);
    const shopCount = shopCountByStage().get(stageNo) ?? 1;
    if (shopCount === 1 && isDedicatedShopStageName(official)) return official;
    return npcShopLabel(npcName, funcLabel);
}

function shopNamedLocationAlt(shopLabel, npcName, funcLabel) {
    const parts = [];
    if (npcName && !shopLabel.toLowerCase().includes(npcName.toLowerCase())) parts.push(npcName);
    if (funcLabel && !shopLabel.toLowerCase().includes(funcLabel.toLowerCase())) parts.push(funcLabel);
    return parts.join(' · ');
}

function namedLocationAreaLabel(entry) {
    const area = mapAreaLabel(entry.mapName);
    if (entry.poiCat === 'shop' && !entry.isConnection) return area;
    if (!entry.stid) return area;
    const info = mapParams[entry.mapName];
    if (!info) return area;
    const sLabel = stageLabel(info, entry.stid);
    const mapLabel = info.name_en ? splitPascalCase(info.name_en) : entry.mapName;
    if (sLabel && sLabel !== mapLabel) return `${area} · ${sLabel}`;
    return area;
}

function indexShopNamedLocations(push) {
    const stageLookup = stageNoToMapLookup();
    for (const [stageNo, npcs] of Object.entries(npcShops)) {
        if (!isAccessibleShopStage(stageNo)) continue;
        const loc = stageLookup.get(stageNo);
        if (!loc) continue;
        const { mapName, stageId } = loc;
        for (const npc of npcs) {
            const npcName = npcShopDisplayName(npc.NpcId);
            const funcLabel = NPC_FUNC_LABELS[npc.InstitutionFunctionId] ?? '';
            const shopLabel = resolveShopLabel(stageNo, npcName, funcLabel);
            push({
                label: shopLabel,
                labelAlt: shopNamedLocationAlt(shopLabel, npcName, funcLabel),
                typeLabel: locationTypeLabel('shop'),
                poiCat: 'shop',
                isConnection: false,
                mapName,
                stid: stageId,
                worldX: npc.Position.x,
                worldZ: npc.Position.z,
            });
        }
    }
    for (const [stageNo, npcs] of Object.entries(npcSpecialShops)) {
        if (!isAccessibleShopStage(stageNo)) continue;
        const loc = stageLookup.get(stageNo);
        if (!loc) continue;
        const { mapName, stageId } = loc;
        for (const npc of npcs) {
            const npcName = npcShopDisplayName(npc.NpcId, npc.NpcName);
            const shopType = npc.ShopTypeName
                ? splitPascalCase(String(npc.ShopTypeName).replace(/Exchange/g, ' Exchange '))
                : 'Appraisals';
            const shopLabel = resolveShopLabel(stageNo, npcName, shopType);
            push({
                label: shopLabel,
                labelAlt: shopNamedLocationAlt(shopLabel, npcName, shopType),
                typeLabel: locationTypeLabel('shop'),
                poiCat: 'shop',
                isConnection: false,
                mapName,
                stid: stageId,
                worldX: npc.Position.x,
                worldZ: npc.Position.z,
            });
        }
    }
}

function buildNamedLocationIndex() {
    if (_namedLocationIndex && _namedLocationIndexVersion === NAMED_LOCATION_INDEX_VERSION) return;
    const entries = [];
    const seen = new Set();
    const spotBest = new Map();

    const push = (entry) => {
        const key = `${entry.mapName}:${Math.round(entry.worldX / 50)}:${Math.round(entry.worldZ / 50)}:${entry.label.toLowerCase()}`;
        if (seen.has(key)) return;
        seen.add(key);
        entries.push(entry);
    };

    for (const [mapName, landmarks] of Object.entries(landmarkData)) {
        if (!isMainWorldFieldMap(mapName)) continue;
        for (const lm of landmarks) {
            if (HIDDEN_LANDMARK_TYPES.has(lm.type)) continue;
            const label = lm.spot_name_en?.trim();
            if (!label) continue;
            if (/^Darkness\b/i.test(label)) continue;
            const category = classifyLandmarkPoiCategory(lm.type, lm);
            const filterCat = _normalizePoiFilterCategory(category) ?? category;
            if (!NAMED_LOCATION_CATEGORIES.has(filterCat)) continue;
            const entry = {
                label,
                labelAlt: lm.spot_name_jp ?? '',
                typeLabel: locationTypeLabel(filterCat),
                poiCat: category,
                isConnection: false,
                mapName,
                stid: defaultFieldStage(mapName),
                worldX: lm.x,
                worldZ: lm.z,
            };
            if (lm.spot_id != null) {
                const spotKey = `${mapName}:${lm.spot_id}`;
                const rank = landmarkSearchTypeRank(lm.type);
                const prev = spotBest.get(spotKey);
                if (prev && rank >= prev.rank) continue;
                spotBest.set(spotKey, { entry, rank });
            } else {
                push(entry);
            }
        }
    }
    for (const { entry } of spotBest.values()) push(entry);

    for (const [sourceMap, conns] of Object.entries(connectionData)) {
        if (!isMainWorldFieldMap(sourceMap)) continue;
        for (const conn of conns) {
            const label = conn.name_en?.trim();
            if (!label || conn.x == null || conn.z == null) continue;
            if (matchSpecialConnection(sourceMap, conn)) continue;
            if (isSuppressedArkConnection(conn)) continue;
            // Overworld entrances are indexed from the field map only. Hub/dungeon exits
            // back to fields (e.g. "Farana Plains" from rm011) duplicate field-side hits.
            if (!conn.to_map || isMainWorldFieldMap(conn.to_map)) continue;
            if (/^house$/i.test(label)) continue;
            const poiCat = classifyConnectionPoi(conn, sourceMap) ?? 'door';
            const filterCat = _normalizePoiFilterCategory(poiCat) ?? poiCat;
            push({
                label,
                labelAlt: '',
                typeLabel: locationTypeLabel(filterCat),
                poiCat,
                isConnection: true,
                mapName: sourceMap,
                stid: stageIdFromNo(conn.from_stage) ?? defaultFieldStage(sourceMap),
                worldX: conn.x,
                worldZ: conn.z,
            });
        }
    }

    indexShopNamedLocations(push);

    _namedLocationIndex = entries;
    _namedLocationIndexVersion = NAMED_LOCATION_INDEX_VERSION;
}

function filterNamedLocationEntries(rawFilter) {
    buildNamedLocationIndex();
    const query = parseSearchQuery(rawFilter);
    const { term, exact } = _parseSpotQuery(query.text);
    if (!term) return [];

    const matches = [];
    for (const entry of _namedLocationIndex) {
        if (!namedLocationMatchesEntry(entry, term, exact)) continue;
        matches.push(entry);
    }

    matches.sort((a, b) => {
        const rank = (entry) => {
            const ll = entry.label.toLowerCase();
            if (ll === term) return 0;
            if (ll.startsWith(term)) return 1;
            return 2;
        };
        const rDiff = rank(a) - rank(b);
        if (rDiff !== 0) return rDiff;
        const areaDiff = mapAreaLabel(a.mapName).localeCompare(mapAreaLabel(b.mapName));
        if (areaDiff !== 0) return areaDiff;
        return a.label.localeCompare(b.label, undefined, { numeric: true, sensitivity: 'base' });
    });
    return matches;
}

function navigateToNamedLocation(entry) {
    const info = mapParams[entry.mapName];
    if (!info) return;
    const latlng = worldToPixel(entry.worldX, entry.worldZ, info);
    const zoom = 1.75;
    const stid = entry.stid;
    const sameMap = entry.mapName === _loadedMapName;
    const sameStage = !stid || stid === currentStageName();

    const persistViewInHash = () => {
        const mapPart = stid ? `${entry.mapName}:${stid}` : entry.mapName;
        history.replaceState(null, '', `#${mapPart}@${zoom.toFixed(2)}/${latlng.lat.toFixed(1)}/${latlng.lng.toFixed(1)}!${getLayersHash()}`);
    };

    if (sameMap && sameStage) {
        leafletMap.flyTo(latlng, zoom, { duration: 0.45 });
        _clearSpotHighlights();
        _addSpotHighlight(latlng);
        persistViewInHash();
        return;
    }

    _pendingNamedLocNav = { latlng };
    navigateTo(entry.mapName, stid, { zoom, center: latlng });
}

function appendNamedLocationSearchResults(listEl, rawFilter) {
    const matches = filterNamedLocationEntries(rawFilter);
    if (!matches.length) return;

    const section = document.createElement('div');
    section.className = 'map-loc-section';

    const header = document.createElement('div');
    header.className = 'map-loc-header';
    header.textContent = `Named locations (${matches.length})`;
    section.appendChild(header);

    for (const entry of matches.slice(0, NAMED_LOCATION_RESULT_CAP)) {
        const el = document.createElement('div');
        el.className = 'map-loc-entry';
        el.title = `${entry.typeLabel} · ${namedLocationAreaLabel(entry)}`;

        const icon = document.createElement('span');
        icon.className = 'map-loc-icon poi-icon';
        const iconSrc = namedLocationEntryIconSrc(entry);
        if (iconSrc) {
            const img = document.createElement('img');
            img.className = 'poi-icon-img';
            img.src = iconSrc;
            img.alt = '';
            img.loading = 'lazy';
            icon.appendChild(img);
        } else {
            icon.textContent = '◆';
        }

        const body = document.createElement('div');
        body.className = 'map-loc-body';

        const labelEl = document.createElement('div');
        labelEl.className = 'map-loc-label';
        labelEl.textContent = entry.label;

        const subEl = document.createElement('div');
        subEl.className = 'map-loc-sub';
        subEl.textContent = `${entry.typeLabel} · ${namedLocationAreaLabel(entry)}`;

        body.append(labelEl, subEl);
        el.append(icon, body);
        el.addEventListener('click', () => navigateToNamedLocation(entry));
        section.appendChild(el);
    }

    if (matches.length > NAMED_LOCATION_RESULT_CAP) {
        const more = document.createElement('div');
        more.className = 'map-loc-more';
        more.textContent = `${matches.length - NAMED_LOCATION_RESULT_CAP} more — refine search`;
        section.appendChild(more);
    }

    listEl.appendChild(section);
}

function anyLocationFilterEnabled() {
    return POI_LOCATION_CATEGORIES.some(c => _poiFilters[c.id] !== false);
}

function syncLocationsShowAllCheckbox() {
    const el = document.getElementById('poi-locations-show-all');
    if (!el) return;
    const all = POI_LOCATION_CATEGORIES.every(c => _poiFilters[c.id] !== false);
    const none = POI_LOCATION_CATEGORIES.every(c => _poiFilters[c.id] === false);
    el.checked = all;
    el.indeterminate = !all && !none;
}

function setAllLocationFilters(on) {
    for (const c of POI_LOCATION_CATEGORIES) _poiFilters[c.id] = on;
    savePoiFilters();
    for (const c of POI_LOCATION_CATEGORIES) {
        const inp = document.getElementById(`poi-filter-${c.id}`);
        if (inp) inp.checked = on;
    }
    syncLocationsShowAllCheckbox();
    applyAllPoiVisibility();
}

function syncGatherablesShowAllCheckbox() {
    const el = document.getElementById('poi-gatherables-show-all');
    if (!el) return;
    const all = POI_GATHER_CATEGORIES.every(c => _poiFilters[c.id] !== false);
    const none = POI_GATHER_CATEGORIES.every(c => _poiFilters[c.id] === false);
    el.checked = all;
    el.indeterminate = !all && !none;
}

function setAllGatherableFilters(on) {
    for (const c of POI_GATHER_CATEGORIES) _poiFilters[c.id] = on;
    savePoiFilters();
    for (const c of POI_GATHER_CATEGORIES) {
        const inp = document.getElementById(`poi-filter-${c.id}`);
        if (inp) inp.checked = on;
    }
    syncGatherablesShowAllCheckbox();
    applyAllPoiVisibility();
}

function _isConnectionVisible(poiCat, isSpecialConnection = false) {
    if (!anyLocationFilterEnabled()) return false;
    if (isSpecialConnection && !_isPoiCategoryEnabled('special')) return false;
    if (!poiCat) return true;
    const norm = _normalizePoiFilterCategory(poiCat);
    return _isPoiCategoryEnabled(norm);
}

function _isPoiCategoryEnabled(category) {
    if (!category || category.endsWith('Other')) return true;
    const norm = _normalizePoiFilterCategory(category);
    if (!(norm in POI_FILTER_DEFAULTS)) return true;
    return _poiFilters[norm] !== false;
}

function _setPoiMarkerVisible(m, visible) {
    m._poiHidden = !visible;
    if (m.setOpacity) m.setOpacity(visible ? 1 : 0);
    if (m.setStyle) {
        if (visible) {
            m.setStyle({ opacity: 1, fillOpacity: m._poiFillOpacity ?? 0.85 });
        } else {
            if (m._poiFillOpacity == null && m.options?.fillOpacity != null)
                m._poiFillOpacity = m.options.fillOpacity;
            m.setStyle({ opacity: 0, fillOpacity: 0 });
        }
    }
    m.options.interactive = visible;
    const el = m.getElement?.();
    if (el) el.style.pointerEvents = visible ? '' : 'none';
}

function _applyPoiMarkerVisibility(m) {
    const cat = m._poiCategory;
    if (!cat && !m._poiIsConnection) return;
    let visible;
    if (m._poiConnectionChoices?.length) {
        visible = m._poiConnectionChoices.some((c) => _isConnectionVisible(c.poiCat, c.isSpecial));
    } else if (m._poiIsConnection) {
        visible = _isConnectionVisible(cat, m._poiSpecialConnection);
    } else {
        visible = _isPoiCategoryEnabled(cat);
    }
    _setPoiMarkerVisible(m, visible);
}

function applyAllPoiVisibility() {
    for (const m of _gatherMarkerByKey.values()) _applyPoiMarkerVisibility(m);
    for (const m of _shopMarkerByNpcId.values()) _applyPoiMarkerVisibility(m);
    for (const m of _specialShopMarkerByNpcId.values()) _applyPoiMarkerVisibility(m);
    for (const m of landmarkLayer.getLayers()) _applyPoiMarkerVisibility(m);
    for (const m of connectionLayer.getLayers()) _applyPoiMarkerVisibility(m);
}

function _buildPoiFilterList(listEl, categories, { onItemChange } = {}) {
    if (!listEl) return;
    listEl.innerHTML = '';
    for (const cat of categories) {
        const label = document.createElement('label');
        label.className = 'layer-toggle';
        label.title = `Show/hide ${cat.label.toLowerCase()}`;
        const input = document.createElement('input');
        input.type = 'checkbox';
        input.id = `poi-filter-${cat.id}`;
        input.checked = _poiFilters[cat.id] !== false;
        input.addEventListener('change', () => {
            _poiFilters[cat.id] = input.checked;
            savePoiFilters();
            onItemChange?.();
            applyAllPoiVisibility();
        });
        const icon = document.createElement('span');
        icon.className = 'poi-icon';
        icon.dataset.poi = cat.id;
        if (cat.icon) {
            const img = document.createElement('img');
            img.className = 'poi-icon-img';
            const src = poiFilterIconSrc(cat.id);
            if (src) img.src = src;
            img.alt = '';
            img.loading = 'lazy';
            icon.appendChild(img);
            label.append(input, icon, document.createTextNode(` ${cat.label}`));
        } else {
            label.append(input, document.createTextNode(` ${cat.label}`));
        }
        listEl.appendChild(label);
    }
}

const POI_SECTION_OPEN_KEY = 'ddon-poi-section-open';
const POI_SECTION_DEFAULTS = { mobTypes: true, locations: true, gatherables: true };

function initPoiSectionToggles() {
    let open = { ...POI_SECTION_DEFAULTS };
    try {
        const raw = localStorage.getItem(POI_SECTION_OPEN_KEY);
        if (raw) open = { ...open, ...JSON.parse(raw) };
    } catch { /* ignore */ }

    document.querySelectorAll('.poi-section[data-poi-section]').forEach(section => {
        const key = section.dataset.poiSection;
        const btn = section.querySelector('.poi-section-toggle');
        const expanded = open[key] !== false;
        section.classList.toggle('expanded', expanded);
        btn?.setAttribute('aria-expanded', String(expanded));
        btn?.addEventListener('click', () => {
            const nowExpanded = section.classList.toggle('expanded');
            btn.setAttribute('aria-expanded', String(nowExpanded));
            open[key] = nowExpanded;
            try { localStorage.setItem(POI_SECTION_OPEN_KEY, JSON.stringify(open)); } catch { /* ignore */ }
        });
    });
}

function initPoiFilters() {
    document.getElementById('poi-locations-show-all')?.addEventListener('change', (e) => {
        setAllLocationFilters(e.target.checked);
    });
    document.getElementById('poi-gatherables-show-all')?.addEventListener('change', (e) => {
        setAllGatherableFilters(e.target.checked);
    });
    _buildPoiFilterList(
        document.getElementById('poi-filter-locations'),
        POI_LOCATION_CATEGORIES,
        { onItemChange: syncLocationsShowAllCheckbox },
    );
    _buildPoiFilterList(
        document.getElementById('poi-filter-gatherables'),
        POI_GATHER_CATEGORIES,
        { onItemChange: syncGatherablesShowAllCheckbox },
    );
    initPoiSectionToggles();
    syncLocationsShowAllCheckbox();
    syncGatherablesShowAllCheckbox();
}

// ── Live server data (fetched from GitHub at runtime) ─────────────────────────
// File content for local overrides is stored in IndexedDB (higher capacity than
// localStorage). localStorage only holds the sentinel '__local__' or a URL string.
const _idb = new Promise(resolve => {
    try {
        const req = indexedDB.open('ddon-maps-src', 1);
        req.onupgradeneeded = e => e.target.result.createObjectStore('files');
        req.onsuccess = e => resolve(e.target.result);
        req.onerror = () => resolve(null);
    } catch { resolve(null); }
});
function _idbGet(key) {
    return _idb.then(db => db ? new Promise(res => {
        const r = db.transaction('files').objectStore('files').get(key);
        r.onsuccess = () => res(r.result ?? null); r.onerror = () => res(null);
    }) : null);
}
function _idbSet(key, val) {
    return _idb.then(db => db ? new Promise((res, rej) => {
        const tx = db.transaction('files', 'readwrite');
        tx.objectStore('files').put(val, key);
        tx.oncomplete = res; tx.onerror = rej;
    }) : null);
}
function _idbDel(key) {
    return _idb.then(db => db ? new Promise((res, rej) => {
        const tx = db.transaction('files', 'readwrite');
        tx.objectStore('files').delete(key);
        tx.oncomplete = res; tx.onerror = rej;
    }) : null);
}

// Returns a URL to use for fetching (custom URL, blob from IDB/FSA handle, or default).
// FSA handles (stored under lsKey+'-handle') always re-read the live file from disk.
// Stored content (IDB under lsKey) is the fallback when FSA isn't available or lacks permission.
async function getSrcUrl(lsKey, defaultUrl) {
    try {
        const val = localStorage.getItem(lsKey);
        if (!val) return defaultUrl;
        if (val === '__local__') {
            // Try FSA handle first — always reads fresh file from disk
            const handle = await _idbGet(lsKey + '-handle');
            if (handle) {
                try {
                    if (typeof handle.queryPermission === 'function') {
                        // Chrome / Brave: full permission API available
                        let perm = await handle.queryPermission({ mode: 'read' });
                        if (perm === 'prompt') {
                            // Attempt to re-request permission. Requires a user gesture;
                            // may succeed silently within the same browser session.
                            try { perm = await handle.requestPermission({ mode: 'read' }); } catch { /* blocked */ }
                        }
                        if (perm !== 'granted') {
                            // Not granted — fall through to stored content.
                            // User can open ⚙ Settings to grant it (opening the modal = a user gesture).
                            throw new Error('permission not granted');
                        }
                    }
                    // Read via stream() → Response to bypass Chromium's FSA write-through cache.
                    // file.text() can return buffered content from the browser's last write;
                    // consuming the stream through a Response uses a different code path.
                    const file = await handle.getFile();
                    const text = await new Response(file.stream()).text();
                    return URL.createObjectURL(new Blob([text], { type: file.type || 'text/plain' }));
                } catch { /* handle stale or permission denied — fall through */ }
            }
            // Fall back to stored content (IDB or legacy localStorage)
            const data = await _idbGet(lsKey);
            if (data) return URL.createObjectURL(new Blob([data]));
            const lsData = localStorage.getItem(lsKey + '-data');
            if (lsData) return URL.createObjectURL(new Blob([lsData]));
            return defaultUrl;
        }
        return val;
    } catch { return defaultUrl; }
}

// On page load: for every __local__ source, show a sidebar indicator.
// If permission needs re-granting (Chromium after restart), show a Re-grant button instead.
async function checkLocalSources() {
    const ua = navigator.userAgent;
    const isBrave   = navigator.brave?.isBrave != null;
    const isFirefox = ua.includes('Firefox/');
    // Chrome and Edge handle FSA permissions natively with their own dialogs — no indicator needed.
    if (!isBrave && !isFirefox) return;

    if (sessionStorage.getItem('ddon-src-reloaded')) {
        sessionStorage.removeItem('ddon-src-reloaded');
        return;
    }
    const KEYS = ['ddon-src-gathering', 'ddon-src-spawns', 'ddon-src-shop', 'ddon-src-special-shop'];
    const LABELS = { 'ddon-src-gathering': 'Gathering Items', 'ddon-src-spawns': 'Enemy Spawns',
                     'ddon-src-shop': 'Shop Data', 'ddon-src-special-shop': 'Appraisal Data' };
    for (const lsKey of KEYS) {
        if (localStorage.getItem(lsKey) !== '__local__') continue;
        const label  = LABELS[lsKey];
        const handle = await _idbGet(lsKey + '-handle');
        if (handle && typeof handle.queryPermission === 'function') {
            try {
                const perm = await handle.queryPermission({ mode: 'read' });
                if (perm !== 'granted') { showSrcPermissionWarning(label, handle); continue; }
            } catch { /* ignore */ }
        }
        showSrcLocalIndicator(label);
    }
}
function _srcErrorBox() {
    let box = document.getElementById('src-errors');
    if (!box) {
        box = document.createElement('div');
        box.id = 'src-errors';
        box.style.cssText = 'font-size:0.75rem;padding:2px 8px;';
        const sidebar = document.getElementById('sidebar');
        const anchor  = sidebar.querySelector('#search-box') ?? sidebar.children[1];
        sidebar.insertBefore(box, anchor);
    }
    return box;
}
function showSrcLocalIndicator(label) {
    const box = _srcErrorBox();
    if ([...box.children].some(c => c.dataset.localLabel === label)) return;
    const item = document.createElement('div');
    item.dataset.localLabel = label;
    item.style.cssText = 'display:flex;align-items:center;gap:5px;margin-bottom:3px;'
        + 'color:#aad4ff;background:#0d1f2d;border-left:3px solid #42a5f5;border-radius:2px;padding:3px 6px;';
    item.innerHTML = `&#128196; <span style="flex:1"><b>${label}</b> from local file</span>`
        + `<button data-action="reload" style="font-size:0.7rem;padding:1px 5px;cursor:pointer;`
        + `background:#42a5f5;color:#000;border:none;border-radius:2px" title="Reload to re-read file from disk">&#8635; Reload</button>`
        + `<button data-action="dismiss" style="font-size:0.7rem;padding:1px 4px;cursor:pointer;`
        + `background:none;color:#aaa;border:none">&#10005;</button>`;
    item.querySelector('[data-action="reload"]').addEventListener('click', () => {
        sessionStorage.setItem('ddon-src-reloaded', '1');
        location.reload();
    });
    item.querySelector('[data-action="dismiss"]').addEventListener('click', () => item.remove());
    box.appendChild(item);
}
function showSrcPermissionWarning(label, handle) {
    const box = _srcErrorBox();
    if ([...box.children].some(c => c.dataset.permLabel === label)) return;
    const item = document.createElement('div');
    item.dataset.permLabel = label;
    item.style.cssText = 'display:flex;align-items:center;gap:5px;margin-bottom:3px;'
        + 'color:#ffd27f;background:#2a1f00;border-left:3px solid #ffa726;border-radius:2px;padding:3px 6px;';
    item.innerHTML = `&#128274; <span style="flex:1"><b>${label}</b> needs file access</span>`
        + `<button data-action="grant" style="font-size:0.7rem;padding:1px 5px;cursor:pointer;`
        + `background:#ffa726;color:#111;border:none;border-radius:2px">Re-grant &amp; Reload</button>`
        + `<button data-action="dismiss" style="font-size:0.7rem;padding:1px 4px;cursor:pointer;`
        + `background:none;color:#aaa;border:none">&#10005;</button>`;
    item.querySelector('[data-action="grant"]').addEventListener('click', async () => {
        try {
            const perm = await handle.requestPermission({ mode: 'read' });
            if (perm === 'granted') location.reload();
            else alert('Permission was not granted.');
        } catch (err) { alert('Could not request permission: ' + err.message); }
    });
    item.querySelector('[data-action="dismiss"]').addEventListener('click', () => item.remove());
    box.appendChild(item);
}
function showSrcError(label) {
    const box = _srcErrorBox();
    const item = document.createElement('div');
    item.style.cssText = 'display:flex;align-items:center;gap:5px;margin-bottom:3px;'
        + 'color:#f99;background:#2a0f0f;border-left:3px solid #c0392b;border-radius:2px;padding:3px 6px;';
    item.innerHTML = `⚠ <span style="flex:1"><b>${label}</b> failed to load</span>`
        + `<button data-action="fix" style="font-size:0.7rem;padding:1px 5px;cursor:pointer;`
        + `background:none;color:#ffd700;border:1px solid #666;border-radius:2px">⚙ Fix</button>`
        + `<button data-action="dismiss" style="font-size:0.7rem;padding:1px 4px;cursor:pointer;`
        + `background:none;color:#aaa;border:none">✕</button>`;
    item.querySelector('[data-action="fix"]').addEventListener('click',
        () => {});
    item.querySelector('[data-action="dismiss"]').addEventListener('click', () => item.remove());
    box.appendChild(item);
}

function _updateClipboardBar() {
    const bar = document.getElementById('edit-clipboard-bar');
    if (!bar) return;
    if (_copiedEnemyConfig) {
        const name = emNames[_copiedEnemyConfig.emCode]?.name ?? _copiedEnemyConfig.emCode ?? '?';
        bar.querySelector('.edit-clipboard-label').textContent = `📋 ${name} — click Paste on any marker`;
        bar.style.display = '';
    } else {
        bar.style.display = 'none';
    }
}

// Rising Maps: Normal + Endgame (CH4) only.
const _EDELARROW_SPAWNS_BASE = 'https://raw.githubusercontent.com/edelarrow/map-spawns/refs/heads/main';
const _DEFAULT_GATHERING_URL    = `${_EDELARROW_SPAWNS_BASE}/Normal%20Channels/GatheringItem.csv`;
const _DEFAULT_SPAWNS_URL       = `${_EDELARROW_SPAWNS_BASE}/Normal%20Channels/EnemySpawn.json`;
const _DEFAULT_SHOP_URL         = `${_EDELARROW_SPAWNS_BASE}/Normal%20Channels/Shop.json`;
const _DEFAULT_SPECIAL_SHOP_URL = `${_EDELARROW_SPAWNS_BASE}/Normal%20Channels/SpecialShops.json`;

const _PRESET_LS_KEY = 'ddon-rising-channel';
const _SRC_KEYS = ['ddon-src-gathering', 'ddon-src-spawns', 'ddon-src-shop', 'ddon-src-special-shop'];

const _SERVER_PRESETS = {
    rising: {
        label: 'Normal Channels',
        urls: {
            'ddon-src-gathering':    _DEFAULT_GATHERING_URL,
            'ddon-src-spawns':       _DEFAULT_SPAWNS_URL,
            'ddon-src-shop':         _DEFAULT_SHOP_URL,
            'ddon-src-special-shop': _DEFAULT_SPECIAL_SHOP_URL,
        },
    },
    'rising-endgame': {
        label: 'Endgame (CH4)',
        urls: {
            'ddon-src-gathering':    `${_EDELARROW_SPAWNS_BASE}/Channel%204/GatheringItem.csv`,
            'ddon-src-spawns':       `${_EDELARROW_SPAWNS_BASE}/Channel%204/EnemySpawn.json`,
            'ddon-src-shop':         `${_EDELARROW_SPAWNS_BASE}/Channel%204/Shop.json`,
            'ddon-src-special-shop': `${_EDELARROW_SPAWNS_BASE}/Channel%204/SpecialShops.json`,
        },
    },
};

(() => {
    try {
        const oldPreset = localStorage.getItem('ddon-server-preset');
        if (oldPreset && oldPreset !== 'rising' && oldPreset !== 'rising-endgame') {
            localStorage.removeItem('ddon-server-preset');
        }
        if (!localStorage.getItem(_PRESET_LS_KEY) && (oldPreset === 'rising' || oldPreset === 'rising-endgame')) {
            localStorage.setItem(_PRESET_LS_KEY, oldPreset);
        }
        for (const key of _SRC_KEYS) {
            const val = localStorage.getItem(key);
            if (!val || val === '__local__') continue;
            if (!val.includes('edelarrow/map-spawns')) {
                localStorage.removeItem(key);
                localStorage.removeItem(key + '-name');
                localStorage.removeItem(key + '-data');
            }
        }
    } catch { /* ignore */ }
})();

const presetMatchesStored = (presetId) => {
    const preset = _SERVER_PRESETS[presetId];
    if (!preset) return false;
    return _SRC_KEYS.every(key => {
        const stored = localStorage.getItem(key);
        if (stored === '__local__') return false;
        const expected = preset.urls[key];
        const effective = stored ?? _SERVER_PRESETS.rising.urls[key];
        return effective === expected;
    });
};

const detectActivePreset = () => {
    const storedPreset = localStorage.getItem(_PRESET_LS_KEY)
        || localStorage.getItem('ddon-server-preset');
    if (storedPreset === 'rising' || storedPreset === 'rising-endgame') return storedPreset;
    for (const presetId of Object.keys(_SERVER_PRESETS)) {
        if (presetMatchesStored(presetId)) return presetId;
    }
    return 'rising';
};

function updateSidebarPresetLabel() {
    const el = document.getElementById('sidebar-preset-label');
    if (el) el.textContent = '';
    const presetId = detectActivePreset();
    const titleEl = document.getElementById('rising-title-text');
    if (titleEl) titleEl.textContent = 'Rising Maps';
    document.title = presetId === 'rising-endgame'
        ? 'Rising Maps · Endgame (CH4)'
        : 'Rising Maps · Normal';
}

function applyServerPresetAndReload(presetId) {
    const preset = _SERVER_PRESETS[presetId];
    if (!preset) return;
    for (const key of _SRC_KEYS) {
        const url = preset.urls[key];
        if (url === _SERVER_PRESETS.rising.urls[key]) localStorage.removeItem(key);
        else localStorage.setItem(key, url);
        localStorage.removeItem(key + '-name');
        localStorage.removeItem(key + '-data');
    }
    localStorage.setItem(_PRESET_LS_KEY, presetId);
    localStorage.removeItem('ddon-server-preset');
    location.reload();
}

function updateRisingChannelToggle() {
    const bar = document.getElementById('rising-channel-bar');
    if (!bar) return;
    const presetId = detectActivePreset();
    for (const btn of bar.querySelectorAll('button[data-preset]')) {
        btn.classList.toggle('active', btn.dataset.preset === presetId);
    }
}

function initRisingChannelToggle() {
    const bar = document.getElementById('rising-channel-bar');
    if (!bar || bar.dataset.wired) return;
    bar.dataset.wired = '1';
    bar.addEventListener('click', e => {
        const btn = e.target.closest('button[data-preset]');
        if (!btn) return;
        const presetId = btn.dataset.preset;
        if (!_SERVER_PRESETS[presetId]) return;
        if (btn.classList.contains('active') && detectActivePreset() === presetId) return;
        applyServerPresetAndReload(presetId);
    });
    updateRisingChannelToggle();
}

const presetUrlFor = (presetId, lsKey) =>
    _SERVER_PRESETS[presetId]?.urls[lsKey] ?? _SERVER_PRESETS.rising.urls[lsKey];

// Shared metadata for all user-editable data sources.
// Used by both the settings dialog and the edit panel footer.
const _SOURCE_META = {
    'ddon-src-spawns':       { label: 'Enemy Spawns',   name: 'EnemySpawn.json',   types: [{ description: 'JSON', accept: { 'application/json': ['.json'] } }], defaultUrl: _DEFAULT_SPAWNS_URL },
    'ddon-src-gathering':    { label: 'Gathering Items', name: 'GatherItem.csv',    types: [{ description: 'CSV',  accept: { 'text/csv':            ['.csv'] } }], defaultUrl: _DEFAULT_GATHERING_URL },
    'ddon-src-shop':         { label: 'Shop Data',       name: 'Shop.json',         types: [{ description: 'JSON', accept: { 'application/json': ['.json'] } }], defaultUrl: _DEFAULT_SHOP_URL },
    'ddon-src-special-shop': { label: 'Appraisal Data', name: 'SpecialShops.json', types: [{ description: 'JSON', accept: { 'application/json': ['.json'] } }], defaultUrl: _DEFAULT_SPECIAL_SHOP_URL },
};

// Download the current source URL (or default) to disk via FSA, then immediately
// assign it as the active local source.  Returns the saved filename, or null if
// the user cancelled the save-file picker.
async function downloadAndAssignLocal(lsKey, overrideUrl = null) {
    const meta    = _SOURCE_META[lsKey];
    const stored  = localStorage.getItem(lsKey);
    const url     = overrideUrl ?? ((!stored || stored === '__local__') ? meta.defaultUrl : stored);
    const r       = await fetch(url, { cache: 'no-store' });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const text = await r.text();
    if (typeof showSaveFilePicker === 'function') {
        let handle;
        try {
            handle = await showSaveFilePicker({ suggestedName: meta.name, types: meta.types });
        } catch (e) {
            if (e.name === 'AbortError') return null;   // user cancelled picker
            throw e;
        }
        const writable = await handle.createWritable();
        await writable.write(text);
        await writable.close();
        await _idbSet(lsKey + '-handle', handle);
        localStorage.setItem(lsKey + '-name', handle.name);
        localStorage.setItem(lsKey, '__local__');
        return handle.name;
    } else {
        // FSA unavailable — trigger browser download and cache in IDB
        const blob = new Blob([text], { type: 'text/plain' });
        const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: meta.name });
        document.body.appendChild(a); a.click(); document.body.removeChild(a);
        URL.revokeObjectURL(a.href);
        await _idbSet(lsKey, text);
        localStorage.setItem(lsKey + '-name', meta.name);
        localStorage.setItem(lsKey, '__local__');
        return meta.name;
    }
}

// Cached promises — fetch starts once, result shared by all callers.
// Map key: "stageId,groupId,posId" → [{itemId, itemNum, maxItemNum, quality, dropChance, isHidden}]
let _gatherItemsCache = null;

function parseGatheringCsv(text) {
    const lines = text.trim().split(/\r?\n/);
    _rawGatheringHeaders = lines[0];
    lines[0] = lines[0].replace(/^#/, '');
    const result = new Map();
    const headers = lines[0].split(',');
    const idx = name => headers.indexOf(name);
    const iStage = idx('StageId'), iLayer = idx('LayerNo'), iGroup = idx('GroupId'),
          iPos = idx('PosId'), iItem = idx('ItemId'), iNum = idx('ItemNum'),
          iMax = idx('MaxItemNum'), iQual = idx('Quality'),
          iHidden = idx('IsHidden'), iChance = idx('DropChance');
    _rawGatheringRows = [];
    for (let i = 1; i < lines.length; i++) {
        const cols = lines[i].split(',');
        if (cols.length < 5) continue;
        const row = {
            stageId:    cols[iStage],
            layerNo:    iLayer >= 0 ? cols[iLayer] : 0,
            groupId:    cols[iGroup],
            posId:      cols[iPos],
            itemId:     parseInt(cols[iItem]),
            itemNum:    parseInt(cols[iNum]),
            maxItemNum: parseInt(cols[iMax]),
            quality:    parseInt(cols[iQual]),
            isHidden:   cols[iHidden] === 'true' || cols[iHidden] === '1',
            dropChance: iChance >= 0 ? parseFloat(cols[iChance]) : 1,
        };
        _rawGatheringRows.push(row);
        const key = `${row.stageId},${row.groupId},${row.posId}`;
        if (!result.has(key)) result.set(key, []);
        result.get(key).push({
            itemId: row.itemId, itemNum: row.itemNum, maxItemNum: row.maxItemNum,
            quality: row.quality, isHidden: row.isHidden, dropChance: row.dropChance,
        });
    }
    return result;
}

async function loadGatheringFromUrl(url) {
    const res = await fetch(url, { cache: 'no-store' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return parseGatheringCsv(await res.text());
}

const _gatherItemsPromise = getSrcUrl('ddon-src-gathering', _DEFAULT_GATHERING_URL)
    .then(loadGatheringFromUrl)
    .then(result => { _gatherItemsCache = result; return result; })
    .catch(() => { showSrcError('Gathering Items'); _gatherItemsCache = new Map(); return _gatherItemsCache; });

// Map key for enemy spawns: "stageId,groupId,posIdx" → [{emCode, lv, bloodOrbs, highOrbs, spawnTime, drops}]
let _enemySpawnCache = null;
const _enemySpawnPromise = getSrcUrl('ddon-src-spawns', _DEFAULT_SPAWNS_URL)
    .then(url => fetch(url, { cache: 'no-store' }).then(r => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.json(); }))
    .then(data => {
        const schemas = data.schemas?.enemies ?? [];
        const iStage      = schemas.indexOf('StageId'),
              iGroup      = schemas.indexOf('GroupId'),
              iPosIdx     = schemas.indexOf('PositionIndex'),
              iEnemyId    = schemas.indexOf('EnemyId'),
              iLv         = schemas.indexOf('Lv'),
              iBlood      = schemas.indexOf('BloodOrbs'),
              iHigh       = schemas.indexOf('HighOrbs'),
              iSpawnTime  = schemas.indexOf('SpawnTime'),
              iDrops      = schemas.indexOf('DropsTableId'),
              iScale      = schemas.indexOf('Scale'),
              iSubGroup   = schemas.indexOf('SubGroupId'),
              iNamed      = schemas.indexOf('NamedEnemyParamsId'),
              iRaidBoss   = schemas.indexOf('RaidBossId'),
              iSetType    = schemas.indexOf('SetType'),
              iInfection  = schemas.indexOf('InfectionType'),
              iIsBossG    = schemas.indexOf('IsBossGauge'),
              iIsBossBGM  = schemas.indexOf('IsBossBGM'),
              iIsAreaBoss = schemas.indexOf('IsAreaBoss'),
              iExp        = schemas.indexOf('Experience'),
              iRepopNum   = schemas.indexOf('RepopNum'),
              iRepopCount = schemas.indexOf('RepopCount'),
              iTargetType = schemas.indexOf('EnemyTargetTypesId'),
              iHmPreset   = schemas.indexOf('HmPresetNo'),
              iStartThink = schemas.indexOf('StartThinkTblNo'),
              iMontage    = schemas.indexOf('MontageFixNo'),
              iIsManualSet     = schemas.indexOf('IsManualSet'),
              iPPDrop          = schemas.indexOf('PPDrop'),
              iIsBloodOrbEnemy = schemas.indexOf('IsBloodOrbEnemy'),
              iIsHighOrbEnemy  = schemas.indexOf('IsHighOrbEnemy'),
              iLayerNo         = schemas.indexOf('LayerNo');
        // Store raw data and schema indices for write-back
        _rawEnemyData    = data;
        _rawEnemySchemas = {
            iStage, iGroup, iPosIdx, iEnemyId, iLv, iBlood, iHigh, iSpawnTime, iDrops,
            iScale, iSubGroup, iNamed, iRaidBoss, iSetType, iInfection,
            iIsBossG, iIsBossBGM, iIsAreaBoss, iExp, iRepopNum, iRepopCount,
            iTargetType, iHmPreset, iStartThink, iMontage, iIsManualSet, iPPDrop,
            iIsBloodOrbEnemy, iIsHighOrbEnemy, iLayerNo,
        };
        const dropsTables = {};
        _dropsTablesMap = new Map();
        for (const dt of (data.dropsTables ?? [])) { dropsTables[dt.id] = dt; _dropsTablesMap.set(dt.id, dt); }
        const result = new Map();
        for (let rawIdx = 0; rawIdx < (data.enemies?.length ?? 0); rawIdx++) {
            const e    = data.enemies[rawIdx];
            const key  = `${e[iStage]},${e[iGroup]},${e[iPosIdx]}`;
            const dtId = e[iDrops];
            const dt   = dtId != null && dtId >= 0 ? dropsTables[dtId] : null;
            // Convert '0x011200' → 'em011200'
            const hexStr = e[iEnemyId];
            const emCode = hexStr ? `em${hexStr.slice(2).toLowerCase().padStart(6, '0')}` : null;
            const bloodOrbs = e[iBlood] ?? 0;
            const highOrbs  = e[iHigh]  ?? 0;
            const entry = {
                emCode,
                lv:          e[iLv]         ?? null,
                bloodOrbs,
                highOrbs,
                spawnTime:     e[iSpawnTime]  ?? null,
                dropsTableId:  dtId ?? -1,
                drops:         dt ? dt.items : [],
                scale:       e[iScale]      ?? 100,
                subGroupId:  e[iSubGroup]   ?? 0,
                namedId:     e[iNamed]      ?? 0,
                raidBossId:  e[iRaidBoss]   ?? 0,
                setType:     e[iSetType]    ?? 0,
                infection:   e[iInfection]  ?? 0,
                isBossGauge: e[iIsBossG]     ?? false,
                isBossBGM:   e[iIsBossBGM]  ?? false,
                isAreaBoss:  e[iIsAreaBoss]  ?? false,
                isManualSet:      e[iIsManualSet] ?? false,
                // Fallback mirrors server logic: old schema derives from orb value > 0
                isBloodOrbEnemy:  iIsBloodOrbEnemy >= 0 ? (e[iIsBloodOrbEnemy] ?? false) : bloodOrbs > 0,
                isHighOrbEnemy:   iIsHighOrbEnemy  >= 0 ? (e[iIsHighOrbEnemy]  ?? false) : highOrbs  > 0,
                exp:         e[iExp]         ?? 0,
                repopNum:    e[iRepopNum]    ?? 0,
                repopCount:  e[iRepopCount]  ?? 0,
                targetTypeId:e[iTargetType]  ?? 0,
                hmPreset:    e[iHmPreset]    ?? 0,
                startThink:  e[iStartThink]  ?? 0,
                montage:     e[iMontage]     ?? 0,
                ppDrop:      e[iPPDrop]      ?? 0,
                _rawIdx:     rawIdx,
            };
            if (result.has(key)) result.get(key).push(entry);
            else                 result.set(key, [entry]);
        }
        _enemySpawnCache = result;
        return result;
    })
    .catch(() => { showSrcError('Enemy Spawns'); _enemySpawnCache = new Map(); return _enemySpawnCache; })
    .finally(() => {
        _enemySpawnDataLoaded = true;
        applySubGroupFilter();
        refreshMobTooltips();
    });

// Map: ShopId → {walletType, items:[{itemId, price, stock}]}
let _shopCache = null;
const _shopPromise = getSrcUrl('ddon-src-shop', _DEFAULT_SHOP_URL)
    .then(url => fetch(url, { cache: 'no-store' }).then(r => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.json(); }))
    .then(data => {
        _rawShopData = data; // store for write-back
        const result = new Map();
        for (const shop of data) {
            result.set(shop.ShopId, {
                walletType: shop.Data.WalletType,
                items:      shop.Data.GoodsParamList ?? [],
            });
        }
        _shopCache = result;
        return result;
    })
    .catch(() => { showSrcError('Shop Data'); _shopCache = new Map(); return _shopCache; });

// Map: ShopType string → { shopTypeId, categories:[{label, appraisals:[...]}] }
// appraisal: { label, comment?, base_items:[{item_id,name,amount}], pool:[{item_id,name,amount,crests?}] }
let _specialShopCache = null;
const _specialShopPromise = getSrcUrl('ddon-src-special-shop', _DEFAULT_SPECIAL_SHOP_URL)
    .then(url => fetch(url, { cache: 'no-store' }).then(r => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.json(); }))
    .then(data => {
        _rawSpecialShopData = data;
        const result = new Map();
        for (const shop of (data.shops ?? [])) {
            const typeStr = shop.shop_type;
            const typeId  = SHOP_TYPE_IDS[typeStr] ?? 0;
            result.set(typeStr, { shopTypeId: typeId, categories: shop.categories ?? [], rawShop: shop });
        }
        _specialShopCache = result;
        return result;
    })
    .catch(() => { showSrcError('Appraisal Data'); _specialShopCache = new Map(); return _specialShopCache; });

// ── NPC shop constants ─────────────────────────────────────────────────────────
const NPC_FUNC_LABELS = {
    3: 'Shop',               4: 'Item Shop',        5: 'Equipment Shop',
    6: 'Material Shop',      8: 'Weapon Shop',       9: 'Armor Shop',
    19: 'Orb Exchange (Crests)',                     20: 'Orb Exchange (Materials)',
    57: 'Play Point Shop',   74: 'Adventure Pass Shop',  97: 'Bitterblack Shop',
};
/** Unreachable shop instances — duplicate WDT (3602) and Megado clones (471/472). */
const INACCESSIBLE_SHOP_STAGE_NOS = new Set(['3602', '471', '472']);
function isAccessibleShopStage(stageNo) {
    return !INACCESSIBLE_SHOP_STAGE_NOS.has(String(stageNo));
}
const NPC_FUNC_COLORS = {
    3: '#ffd700',  4: '#4caf50',  5: '#2196f3',  6: '#ff9800',
    8: '#e91e63',  9: '#9c27b0', 19: '#00bcd4', 20: '#00bcd4',
    57: '#ffeb3b', 74: '#607d8b', 97: '#6a0020',
};
const WALLET_LABELS = {
    1: 'G', 2: 'R', 3: 'BO', 4: 'Tickets', 5: 'Gems',
    6: 'RP', 9: 'HO', 10: 'DP', 11: 'BP', 15: 'Dragon Marks',
};
// ShopType enum (server-side), used for Special Shops (DispelGetDispelItemSettings).
const SHOP_TYPE_IDS = {
    Unknown: 0, Trinkets: 1,
    EmblemMedalExchangeRathniteFoothills: 2, EmblemMedalExchangeFeryanaWilderness: 3,
    EmblemMedalExchangeMegadosysPlateau:  4, EmblemMedalExchangeUrtecaMountains:    5,
    BitterBlackMaze: 6,
    MedalExchangeHidellPlains: 7,  MedalExchangeBriaCoast: 8,
    MedalExchangeVerda: 9,         MedalExchangeHarspuds: 10,
    MedalExchangeWideMoor: 11,     MedalExchangeLimgom: 12,
    MedalExchangeCrumblingStones: 13, MedalExchangeGarlevMountain: 14,
    MedalExchangeOldSorMountain: 15, MedalExchangeCazhetteForest: 16,
    MedalExchangeTataru: 17,       MedalExchangeTamaSpa: 18,
    MedalExchangeGransysOrchard: 19, MedalExchangeGustFront: 20,
    MedalExchangeVolatileVolcano: 21, MedalExchangeCanyon: 22,
    MedalExchangeWestFrontier: 23, ExtremeMission: 27,
};
const SHOP_TYPE_NAMES = Object.fromEntries(Object.entries(SHOP_TYPE_IDS).map(([k, v]) => [v, k]));


// ── Gathering node colors ──────────────────────────────────────────────────────
// Colours and labels for all known GatheringType enum values + treasure-chest UnitID types.
const GATHER_COLORS = {
    // ── Plants ────────────────────────────────────────────────────────────────
    OM_GATHER_GRASS:       '#4CAF50',  // green
    OM_GATHER_FLOWER:      '#E91E63',  // pink
    OM_GATHER_MUSHROOM:    '#9C27B0',  // purple
    OM_GATHER_CLOTH:       '#CE93D8',  // light purple — cloth/fibre
    // ── Ore / Crystal ─────────────────────────────────────────────────────────
    OM_GATHER_CRST_LV1:    '#42A5F5',  // blue
    OM_GATHER_CRST_LV2:    '#1E88E5',  // medium blue
    OM_GATHER_CRST_LV3:    '#1565C0',  // dark blue
    OM_GATHER_CRST_LV4:    '#0D47A1',  // very dark blue
    // ── Gemstone ──────────────────────────────────────────────────────────────
    OM_GATHER_JWL_LV1:     '#FFEE58',  // yellow
    OM_GATHER_JWL_LV2:     '#FDD835',  // deeper yellow
    OM_GATHER_JWL_LV3:     '#F9A825',  // amber
    OM_GATHER_TWINKLE:     '#FFF9C4',  // pale yellow — sparkle/twinkle node
    // ── Lumber ────────────────────────────────────────────────────────────────
    OM_GATHER_TREE_LV1:    '#A1887F',  // light brown
    OM_GATHER_TREE_LV2:    '#795548',  // medium brown
    OM_GATHER_TREE_LV3:    '#5D4037',  // dark brown
    OM_GATHER_TREE_LV4:    '#3E2723',  // very dark brown
    // ── Ground / Water ────────────────────────────────────────────────────────
    OM_GATHER_SAND:        '#FF9800',  // orange
    OM_GATHER_SHELL:       '#FFCC80',  // light amber
    OM_GATHER_WATER:       '#00BCD4',  // cyan
    // ── Items / Misc ──────────────────────────────────────────────────────────
    OM_GATHER_ANTIQUE:     '#FF5722',  // deep orange
    OM_GATHER_BOX:         '#8D6E63',  // brownish
    OM_GATHER_ALCHEMY:     '#00BFA5',  // teal
    OM_GATHER_BOOK:        '#78909C',  // blue-grey
    OM_GATHER_ONE_OFF:     '#B0BEC5',  // light grey — one-off / event node
    OM_GATHER_SHIP:        '#29B6F6',  // light blue — ship/maritime gather
    // ── Key items ─────────────────────────────────────────────────────────────
    OM_GATHER_KEY_LV1:     '#EF9A9A',  // light red
    OM_GATHER_KEY_LV2:     '#EF5350',  // red
    OM_GATHER_KEY_LV3:     '#C62828',  // dark red
    OM_GATHER_KEY_LV4:     '#B71C1C',  // very dark red
    // ── Treasure (gather type on chest OM) ────────────────────────────────────
    OM_GATHER_TREA_IRON:   '#E0E0E0',  // light grey — common chest
    OM_GATHER_TREA_OLD:    '#BCAAA4',  // warm grey  — old/common chest
    OM_GATHER_TREA_TREE:   '#A5D6A7',  // light green — wooden chest
    OM_GATHER_TREA_SILVER: '#CFD8DC',  // silver-grey
    OM_GATHER_TREA_GOLD:   '#FFD54F',  // gold
    // ── Special ───────────────────────────────────────────────────────────────
    OM_GATHER_CORPSE:      '#546E7A',  // dark grey
    OM_GATHER_DRAGON:      '#EF5350',  // bright red
    // ── Treasure chest types (UnitID-based, from SetInfoOmTreasureBox/G) ─────
    CHEST_IRON:            '#90A4AE',  // iron = grey-blue
    CHEST_BROWN:           '#A1887F',  // brown
    CHEST_TREASURE:        '#80CBC4',  // teal-ish
    CHEST_BRONZE:          '#FFAB40',  // bronze/amber
    CHEST_SILVER:          '#E0E0E0',  // silver
    CHEST_GOLD:            '#FFD700',  // gold
    CHEST_PURPLE:          '#CE93D8',  // purple
    CHEST_PEARL:           '#B2EBF2',  // pearlescent — EXM reward
    CHEST_ROUND:           '#FFF59D',  // pale yellow — small round chest
    CHEST_SEALED_BROWN:    '#A1887F',  // Chain Dungeon (EXM)
    CHEST_SEALED_TEAL:     '#60807E',  // Chain Dungeon (EXM)
    CHEST_SEALED_ORANGE:   '#FF6F00',  // orange — BBM sealed
    CHEST_SEALED_PURPLE:   '#7B1FA2',  // purple — BBM sealed
    CHEST_SEALED_PEARL:    '#B2EBF2',  // pearlescent — EXM reward
    CHEST_UNKNOWN:         '#607D8B',  // fallback grey
};

const GATHER_LABELS = {
    OM_GATHER_GRASS:       'Grass / Herb',
    OM_GATHER_FLOWER:      'Flower',
    OM_GATHER_MUSHROOM:    'Mushroom',
    OM_GATHER_CLOTH:       'Cloth / Fibre',
    OM_GATHER_CRST_LV1:    'Crystal (Lv1)',
    OM_GATHER_CRST_LV2:    'Crystal (Lv2)',
    OM_GATHER_CRST_LV3:    'Crystal (Lv3)',
    OM_GATHER_CRST_LV4:    'Crystal (Lv4)',
    OM_GATHER_JWL_LV1:     'Gemstone (Lv1)',
    OM_GATHER_JWL_LV2:     'Gemstone (Lv2)',
    OM_GATHER_JWL_LV3:     'Gemstone (Lv3)',
    OM_GATHER_TWINKLE:     'Sparkle Node',
    OM_GATHER_TREE_LV1:    'Lumber (Lv1)',
    OM_GATHER_TREE_LV2:    'Lumber (Lv2)',
    OM_GATHER_TREE_LV3:    'Lumber (Lv3)',
    OM_GATHER_TREE_LV4:    'Lumber (Lv4)',
    OM_GATHER_SAND:        'Sand',
    OM_GATHER_SHELL:       'Shell',
    OM_GATHER_WATER:       'Water',
    OM_GATHER_ANTIQUE:     'Antique',
    OM_GATHER_BOX:         'Box',
    OM_GATHER_ALCHEMY:     'Alchemy Node',
    OM_GATHER_BOOK:        'Book',
    OM_GATHER_ONE_OFF:     'One-off Node',
    OM_GATHER_SHIP:        'Maritime Gather',
    OM_GATHER_KEY_LV1:     'Locked Chest (Lv1)',
    OM_GATHER_KEY_LV2:     'Locked Chest (Lv2)',
    OM_GATHER_KEY_LV3:     'Locked Chest (Lv3)',
    OM_GATHER_KEY_LV4:     'Locked Chest (Lv4)',
    OM_GATHER_TREA_IRON:   'Treasure (Iron)',
    OM_GATHER_TREA_OLD:    'Treasure (Old)',
    OM_GATHER_TREA_TREE:   'Treasure (Wood)',
    OM_GATHER_TREA_SILVER: 'Treasure (Silver)',
    OM_GATHER_TREA_GOLD:   'Treasure (Gold)',
    OM_GATHER_CORPSE:      'Examine (Corpse)',
    OM_GATHER_DRAGON:      'Dragon Node',
    CHEST_IRON:            'Iron Chest',
    CHEST_BROWN:           'Brown Chest',
    CHEST_TREASURE:        'Treasure Chest',
    CHEST_BRONZE:          'Bronze Chest',
    CHEST_SILVER:          'Silver Chest',
    CHEST_GOLD:            'Gold Chest',
    CHEST_PURPLE:          'Purple Chest',
    CHEST_PEARL:           'Pearlescent Chest',
    CHEST_ROUND:           'Small Round Chest',
    CHEST_SEALED_BROWN:    'Sealed Chest (Brown)',
    CHEST_SEALED_TEAL:     'Sealed Chest (Teal)',
    CHEST_SEALED_ORANGE:   'Sealed Chest (Orange)',
    CHEST_SEALED_PURPLE:   'Sealed Chest (Purple)',
    CHEST_SEALED_PEARL:    'Sealed Chest (Pearlescent)',
    CHEST_UNKNOWN:         'Chest',
};

function loadGatherPoints(info, stid = null) {
    gatherLayer.clearLayers();
    _gatherMarkerByKey.clear();
    _gatherGroupMarkers.clear();
    if (!info.stages?.length) return;

    const floorObbs     = info.floor_obbs ?? null;
    const filterByFloor = floorObbs !== null;
    const stagesToLoad  = (stid && info.stages.includes(stid)) ? [stid] : info.stages;

    for (const stageId of stagesToLoad) {
        const stageNo    = String(parseInt(stageId.slice(2), 10));
        const nodes      = gatherPoints[stageNo];
        if (!nodes) continue;
        const serverStid = stageIds[stageNo];  // server stage_id for CSV lookup

        for (const node of nodes) {
            if (filterByFloor) {
                const floor = getEnemyFloor(node.x, node.y, node.z, floorObbs);
                if (floor !== null && floor !== currentLayer) continue;
            }
            const latlng = worldToPixel(node.x, node.z, info);
            const color  = GATHER_COLORS[node.type] ?? '#aaaaaa';
            const label  = GATHER_LABELS[node.type] ?? node.type.replace(/^(OM_GATHER_|CHEST_)/, '').replace(/_/g, ' ');
            const darkBadge = ['CHEST_SEALED_PURPLE', 'OM_GATHER_CRST_LV4', 'OM_GATHER_TREE_LV3',
                               'OM_GATHER_TREE_LV4', 'OM_GATHER_KEY_LV3', 'OM_GATHER_KEY_LV4',
                               'OM_GATHER_CORPSE'].includes(node.type);
            const badgeText = darkBadge ? '#eee' : '#111';
            const badge     = `<span style="display:inline-block;padding:1px 6px;border-radius:3px;background:${color};color:${badgeText};font-weight:bold;font-size:11px;">${label}</span>`;
            const listIdPart = node.itemListId != null ? ` <span style="color:#888;font-size:10px">· List ${node.itemListId}</span>` : '';
            const typeLine  = `<br><span style="color:#666;font-size:11px">${node.type} <span style="color:#888;font-size:10px">(${node.groupId}.${node.posId})</span>${listIdPart}</span>`;
            const coordLine = `<br><span style="font-size:11px;color:#555">X:&nbsp;${node.x.toFixed(0)}&nbsp; Y:&nbsp;${node.y.toFixed(0)}&nbsp; Z:&nbsp;${node.z.toFixed(0)}</span>`;

            const csvKey = serverStid != null ? `${serverStid},${node.groupId},${node.posId}` : null;

            const buildGatherPopup = (gatherMap) => {
                let itemsHtml = '';
                if (csvKey && gatherMap) {
                    const items = gatherMap.get(csvKey) ?? [];
                    const displayItems = _editMode ? items : items.filter(it => !it.isHidden);
                    if (displayItems.length) {
                        itemsHtml =
                            '<div class="ge-items-view" style="margin-top:6px;display:flex;flex-direction:column;gap:4px">' +
                            displayItems.map(it => {
                                const entry    = itemNames[String(it.itemId)];
                                const name     = entry?.name ?? `Item #${it.itemId}`;
                                const iconNo   = entry?.iconNo;
                                const iconFile = iconNo != null ? `ii${String(iconNo).padStart(6, '0')}.png` : null;
                                const icon     = iconFile && _iconIdSet.has(iconNo)
                                    ? `<img src="images/icons/small/${iconFile}" width="28" height="28" style="vertical-align:middle;image-rendering:pixelated;flex-shrink:0">`
                                    : `<span style="display:inline-block;width:28px;flex-shrink:0"></span>`;
                                const href     = itemWikiHref(it.itemId);
                                const nameLink = `<a href="${href}" target="_blank" style="color:#222;text-decoration:none;font-size:12px;line-height:1.3;word-break:break-word" onmouseover="this.style.textDecoration='underline'" onmouseout="this.style.textDecoration='none'">${name}</a>`;
                                const qty      = it.maxItemNum > it.itemNum ? `×${it.itemNum}–${it.maxItemNum}` : `×${it.itemNum}`;
                                const pct      = it.dropChance > 0 ? ` · ${(it.dropChance * 100).toFixed(0)}%` : '';
                                const meta     = `<span style="font-size:12px;color:#777">${qty}${pct}</span>`;
                                const hiddenStyle = it.isHidden ? 'opacity:0.35;font-style:italic;' : '';
                                const hiddenTitle = it.isHidden ? ' title="Hidden"' : '';
                                return `<div style="${hiddenStyle}display:flex;align-items:flex-start;gap:7px"${hiddenTitle}>`
                                    + icon
                                    + `<div style="display:flex;flex-direction:column;gap:1px;min-width:0">${nameLink}${meta}</div>`
                                    + `</div>`;
                            }).join('') +
                            '</div>';
                    } else if (_editMode && !items.length) {
                        // Empty drop zone placeholder shown only in edit mode when no items at all
                        itemsHtml = '<div class="ge-items-view ge-items-empty" style="min-height:44px;display:flex;align-items:center;justify-content:center;margin-top:8px;border-radius:4px"><span style="color:#888;font-size:11px;font-style:italic;pointer-events:none">Drop items here to add</span></div>';
                    }
                }
                // Edit section — shown only in edit mode when cache is loaded
                let editSection = '';
                if (_editMode && gatherMap && csvKey) {
                    const items = gatherMap.get(csvKey) ?? [];
                    const rows  = items.map((it, i) => {
                        const nm     = itemNames[String(it.itemId)]?.name ?? `Item #${it.itemId}`;
                        const chance = Math.round(((it.dropChance || 0)) * 100);
                        return `<tr data-idx="${i}">` +
                            `<td style="padding:2px 4px;color:#222;min-width:70px">${nm}</td>` +
                            `<td style="padding:2px 1px"><input class="popup-edit-input ge-min" type="number" min="1" value="${it.itemNum || 1}" style="width:36px" title="Min qty"></td>` +
                            `<td style="padding:2px 1px">–</td>` +
                            `<td style="padding:2px 1px"><input class="popup-edit-input ge-max" type="number" min="1" value="${it.maxItemNum || 1}" style="width:36px" title="Max qty"></td>` +
                            `<td style="padding:2px 1px;white-space:nowrap" title="Drop chance %"><input class="popup-edit-input ge-chance" type="number" min="0" max="100" step="1" value="${chance}" style="width:44px">%</td>` +
                            `<td style="padding:2px 1px" title="Rarity (quality)"><input class="popup-edit-input ge-qual" type="number" min="0" max="9" value="${it.quality || 0}" style="width:38px"></td>` +
                            `<td style="padding:2px 3px" title="Hidden"><input type="checkbox" class="ge-hidden"${it.isHidden ? ' checked' : ''}></td>` +
                            `<td style="padding:2px 3px"><button class="popup-edit-btn danger ge-remove" data-idx="${i}" style="padding:1px 4px">✕</button></td>` +
                            `</tr>`;
                    }).join('');
                    editSection = `<div class="popup-edit-section">` +
                        `<table class="popup-edit-items-table" id="ge-table-${csvKey.replace(/,/g,'-')}">${rows}</table>` +
                        `<div class="popup-edit-row" style="margin-top:5px">` +
                        `<input class="popup-edit-input" id="ge-add-id" type="number" placeholder="Item ID" style="width:70px">` +
                        `<button class="popup-edit-btn" data-edit-action="ge-add-item">+ Add</button>` +
                        `<button class="popup-edit-btn ge-apply-btn" data-edit-action="ge-apply" disabled>✔ Apply</button>` +
                        `</div></div>`;
                }
                return `${badge}${typeLine}${coordLine}${itemsHtml}${editSection}`;
            };

            const tooltipText = `${label} — ${node.groupId}.${node.posId}`;
            const poiCategory = classifyGatherPoiCategory(node.type);
            const iconSrc     = gatherMapIconSrc(node.type);
            const gatherIcon  = iconSrc
                ? makePoiMapIcon(iconSrc, 22)
                : L.divIcon({
                    className: '',
                    html: `<div style="width:9px;height:9px;background:${color};border:2px solid rgba(255,255,255,0.7);box-shadow:0 0 3px rgba(0,0,0,0.7);"></div>`,
                    iconSize:    [9, 9],
                    iconAnchor:  [4, 4],
                    popupAnchor: [0, -8],
                });
            const marker = L.marker(latlng, { icon: gatherIcon })
            .bindPopup(buildGatherPopup(_gatherItemsCache))
            .bindTooltip(tooltipText, { permanent: false, direction: 'top', offset: [0, iconSrc ? -12 : -8] })
            .addTo(gatherLayer);
            marker._poiCategory = poiCategory;
            _gatherMarkerByKey.set(`${stageNo}:${node.groupId}:${node.posId}`, marker);

            // Group key kept for potential future group actions; hover highlights this marker only.
            const groupKey = `${stageNo}:${node.groupId}`;
            if (!_gatherGroupMarkers.has(groupKey)) _gatherGroupMarkers.set(groupKey, []);
            _gatherGroupMarkers.get(groupKey).push(marker);
            marker._gatherGroupKey = groupKey;
            marker.on('mouseover', function() { _applyGatherHighlight([this]); });
            marker.on('mouseout',  _unhighlightGather);

            // Drop target: drag items from the Items panel onto a gather node marker.
            // Guard against multiple 'add' events (layer toggled off/on repeatedly).
            if (csvKey) {
                marker.on('add', function() {
                    const el = this.getElement();
                    if (!el || el._ddonDropBound) return;
                    el._ddonDropBound = true;
                    const self = this;
                    el.addEventListener('dragover', (e) => {
                        if (_dragItemId == null || !_editMode) return;
                        e.preventDefault();
                        e.dataTransfer.dropEffect = 'copy';
                        el.style.outline = '2px solid #fff';
                        el.style.outlineOffset = '3px';
                    });
                    el.addEventListener('dragleave', () => {
                        el.style.outline = '';
                        el.style.outlineOffset = '';
                    });
                    el.addEventListener('drop', (e) => {
                        el.style.outline = '';
                        el.style.outlineOffset = '';
                        e.preventDefault();
                        e.stopPropagation();
                        if (_dragItemId == null || !_editMode) return;
                        const itemId = _dragItemId;
                        const doAdd = (gatherMap) => {
                            let items = gatherMap.get(csvKey);
                            if (!items) { items = []; gatherMap.set(csvKey, items); }
                            const [rs, rg, rp] = csvKey.split(',');
                            const newItem = { itemId, itemNum: 1, maxItemNum: 1, quality: 0, isHidden: false, dropChance: 1 };
                            items.push(newItem);
                            if (_rawGatheringRows)
                                _rawGatheringRows.push({ stageId: rs, groupId: rg, posId: rp, ...newItem });
                            if (_markDirty) _markDirty('ddon-src-gathering');
                            if (self.isPopupOpen()) {
                                self.getPopup().setContent(buildGatherPopup(gatherMap));
                                self.getPopup().update();
                            }
                        };
                        if (_gatherItemsCache) doAdd(_gatherItemsCache);
                        else _gatherItemsPromise.then(doAdd);
                    });
                });
            }

            // Always set up popupopen: handles async cache load + edit-mode interactions.
            if (csvKey) {
                const rebuildGatherPopup = (m, gatherMap) => {
                    m.getPopup().setContent(buildGatherPopup(gatherMap));
                    m.getPopup().update();
                    if (_editMode) {
                        const popupEl = m.getPopup().getElement();
                        const tbl = popupEl?.querySelector('.popup-edit-items-table');
                        const applyBtn = popupEl?.querySelector('.ge-apply-btn');
                        // Enable Apply when any field in the table changes
                        if (tbl && applyBtn) {
                            tbl.addEventListener('input', () => { applyBtn.disabled = false; });
                            tbl.addEventListener('change', () => { applyBtn.disabled = false; });
                        }
                        if (_attachDragReorder) {
                            const items = gatherMap.get(csvKey) ?? [];
                            _attachDragReorder(tbl, items, (reorderedItems) => {
                                // Sync raw rows to match new order
                                if (_rawGatheringRows) {
                                    const [rs, rg, rp] = csvKey.split(',');
                                    const others = _rawGatheringRows.filter(
                                        r => !(r.stageId === rs && r.groupId === rg && r.posId === rp));
                                    const reordered = reorderedItems.map(it => ({
                                        stageId: rs, groupId: rg, posId: rp,
                                        itemId: it.itemId, itemNum: it.itemNum, maxItemNum: it.maxItemNum,
                                        quality: it.quality, isHidden: it.isHidden, dropChance: it.dropChance,
                                    }));
                                    _rawGatheringRows = [...others, ...reordered];
                                }
                                if (_markDirty) _markDirty('ddon-src-gathering');
                                rebuildGatherPopup(m, gatherMap);
                            });
                        }
                    }
                };
                let _gatherClickHandler = null;
                marker.on('popupclose', () => {
                    // Clear the drop target so document-level handlers ignore stale popups.
                    if (_gatherPopupDropFn?._marker === marker) _gatherPopupDropFn = null;
                });
                marker.on('popupopen', function() {
                    const self = this;
                    const attach = (gatherMap) => {
                        rebuildGatherPopup(self, gatherMap);
                        // Attach edit action handler
                        const popupEl = self.getPopup().getElement();
                        if (!popupEl) return;
                        if (_gatherClickHandler) popupEl.removeEventListener('click', _gatherClickHandler);
                        if (!_editMode) return;
                        _gatherClickHandler = (e) => {
                            const actionBtn = e.target.closest('[data-edit-action]');
                            const removeBtn = e.target.closest('.ge-remove');
                            if (!actionBtn && !removeBtn) return;
                            e.stopPropagation();
                            let items = gatherMap.get(csvKey);
                            if (!items) { items = []; gatherMap.set(csvKey, items); }
                            if (removeBtn) {
                                const idx = parseInt(removeBtn.dataset.idx);
                                items.splice(idx, 1);
                                // Sync to raw rows
                                if (_rawGatheringRows) {
                                    const [rs, rg, rp] = csvKey.split(',');
                                    const before = _rawGatheringRows.length;
                                    // Remove the idx-th matching row
                                    let matchCount = 0;
                                    _rawGatheringRows = _rawGatheringRows.filter(row => {
                                        if (row.stageId === rs && row.groupId === rg && row.posId === rp) {
                                            return matchCount++ !== idx;
                                        }
                                        return true;
                                    });
                                }
                                if (_markDirty) _markDirty('ddon-src-gathering');
                                rebuildGatherPopup(self, gatherMap);
                            } else if (actionBtn.dataset.editAction === 'ge-apply') {
                                const tbl = popupEl.querySelector('.popup-edit-items-table');
                                if (!tbl) return;
                                const [rs, rg, rp] = csvKey.split(',');
                                tbl.querySelectorAll('tr[data-idx]').forEach(tr => {
                                    const idx     = parseInt(tr.dataset.idx);
                                    const minV    = parseInt(tr.querySelector('.ge-min').value) || 1;
                                    const maxV    = parseInt(tr.querySelector('.ge-max').value) || minV;
                                    const chance  = Math.min(100, Math.max(0, parseFloat(tr.querySelector('.ge-chance').value) || 0)) / 100;
                                    const qualV   = parseInt(tr.querySelector('.ge-qual').value) || 0;
                                    const hiddenV = tr.querySelector('.ge-hidden').checked;
                                    if (items[idx]) {
                                        items[idx].itemNum    = minV;
                                        items[idx].maxItemNum = Math.max(minV, maxV);
                                        items[idx].dropChance = chance;
                                        items[idx].quality    = qualV;
                                        items[idx].isHidden   = hiddenV;
                                    }
                                    if (_rawGatheringRows) {
                                        let matchCount = 0;
                                        for (const row of _rawGatheringRows) {
                                            if (row.stageId === rs && row.groupId === rg && row.posId === rp) {
                                                if (matchCount === idx) {
                                                    row.itemNum    = minV;
                                                    row.maxItemNum = Math.max(minV, maxV);
                                                    row.dropChance = chance;
                                                    row.quality    = qualV;
                                                    row.isHidden   = hiddenV;
                                                    break;
                                                }
                                                matchCount++;
                                            }
                                        }
                                    }
                                });
                                if (_markDirty) _markDirty('ddon-src-gathering');
                                actionBtn.textContent = '✔ Saved';
                                actionBtn.style.background = '#4a9';
                                actionBtn.style.color = '#fff';
                                setTimeout(() => rebuildGatherPopup(self, gatherMap), 500);
                            } else if (actionBtn.dataset.editAction === 'ge-add-item') {
                                const idInput = popupEl.querySelector('#ge-add-id');
                                const newId   = parseInt(idInput?.value);
                                if (!newId || isNaN(newId)) return;
                                const newItem = { itemId: newId, itemNum: 1, maxItemNum: 1, quality: 0, isHidden: false, dropChance: 1 };
                                items.push(newItem);
                                if (_rawGatheringRows) {
                                    const [rs, rg, rp] = csvKey.split(',');
                                    _rawGatheringRows.push({ stageId: rs, groupId: rg, posId: rp, ...newItem });
                                }
                                if (idInput) idInput.value = '';
                                if (_markDirty) _markDirty('ddon-src-gathering');
                                rebuildGatherPopup(self, gatherMap);
                            }
                        };
                        popupEl.addEventListener('click', _gatherClickHandler);

                        // Register this node's add-item function for the document-level
                        // drag-drop handler (see edit block).  Tagged with marker ref so
                        // the popupclose handler can safely clear it.
                        const doAddDraggedItem = (itemId) => {
                            const [rs, rg, rp] = csvKey.split(',');
                            let items = gatherMap.get(csvKey);
                            if (!items) { items = []; gatherMap.set(csvKey, items); }
                            const newItem = { itemId, itemNum: 1, maxItemNum: 1, quality: 0, isHidden: false, dropChance: 1 };
                            items.push(newItem);
                            if (_rawGatheringRows)
                                _rawGatheringRows.push({ stageId: rs, groupId: rg, posId: rp, ...newItem });
                            if (_markDirty) _markDirty('ddon-src-gathering');
                            rebuildGatherPopup(self, gatherMap);
                            // Flash the edit section green to confirm the drop.
                            const flashEl = self.getPopup()?.getElement()?.querySelector('.ge-items-view');
                            if (flashEl) {
                                flashEl.style.transition = 'background 0.1s';
                                flashEl.style.background = 'rgba(80,200,120,0.25)';
                                setTimeout(() => { flashEl.style.background = ''; }, 700);
                            }
                        };
                        doAddDraggedItem._marker = marker;
                        _gatherPopupDropFn = doAddDraggedItem;
                    };
                    if (_gatherItemsCache) { attach(_gatherItemsCache); return; }
                    _gatherItemsPromise.then(gm => { if (self.isPopupOpen()) attach(gm); });
                });
            }
        }
    }

    applyAllPoiVisibility();
}

function loadNpcShops(info, stid = null) {
    npcShopLayer.clearLayers();
    _shopMarkerByNpcId.clear();
    if (!info.stages?.length) return;

    const floorObbs     = info.floor_obbs ?? null;
    const filterByFloor = floorObbs !== null;
    const stagesToLoad  = (stid && info.stages.includes(stid)) ? [stid] : info.stages;

    for (const stageId of stagesToLoad) {
        const stageNo = String(parseInt(stageId.slice(2), 10));
        if (!isAccessibleShopStage(stageNo)) continue;
        const npcs    = npcShops[stageNo];
        if (!npcs) continue;

        for (const npc of npcs) {
            if (filterByFloor) {
                const floor = getEnemyFloor(npc.Position.x, npc.Position.y, npc.Position.z, floorObbs);
                if (floor !== null && floor !== currentLayer) continue;
            }
            const latlng   = worldToPixel(npc.Position.x, npc.Position.z, info);
            const funcId   = npc.InstitutionFunctionId;
            const color    = NPC_FUNC_COLORS[funcId] ?? '#aaaaaa';
            const funcLabel = NPC_FUNC_LABELS[funcId] ?? `Function ${funcId}`;
            const npcName  = npcNames[String(npc.NpcId)] ?? `NPC #${npc.NpcId}`;

            const badge = `<span style="display:inline-block;padding:1px 6px;border-radius:3px;background:${color};color:#111;font-weight:bold;font-size:11px;">${funcLabel}</span>`;

            const buildShopPopup = (shopCache) => {
                const shop     = shopCache?.get(npc.ShopId);
                const currency = shop ? (WALLET_LABELS[shop.walletType] ?? `Type ${shop.walletType}`) : '?';
                const header   = `${badge}<br><span style="color:#333;font-size:12px;font-weight:bold">${npcName}</span>` +
                                 `<br><span style="color:#666;font-size:11px">Shop ID: ${npc.ShopId}</span>`;

                if (!shopCache) return header + '<br><span style="color:#888;font-size:11px">Loading...</span>';

                if (!shop) return header + '<br><span style="color:#888;font-size:11px">No inventory data</span>';

                const currencyLine = `<br><div style="font-size:11px;color:#666;margin-top:2px">Currency: ${currency}</div>`;

                const itemIcon = (it) => {
                    const entry    = itemNames[String(it.ItemId)];
                    const iconNo   = entry?.iconNo;
                    const iconFile = iconNo != null ? `ii${String(iconNo).padStart(6, '0')}.png` : null;
                    return iconFile && _iconIdSet.has(iconNo)
                        ? `<img src="images/icons/small/${iconFile}" width="28" height="28" style="vertical-align:middle;margin-right:6px;image-rendering:pixelated">`
                        : `<span style="display:inline-block;width:28px;margin-right:6px"></span>`;
                };

                if (_editMode) {
                    // ── Edit mode: editable rows with price, stock, remove ──────────
                    const rows = shop.items.map((it, idx) => {
                        const name = itemNames[String(it.ItemId)]?.name ?? `Item #${it.ItemId}`;
                        return `<tr data-idx="${idx}">` +
                            `<td style="color:#888;padding-right:4px;cursor:grab;user-select:none;white-space:nowrap;width:1px" title="Drag to reorder">⠿</td>` +
                            `<td style="padding-right:6px;font-size:12px">${itemIcon(it)}${name}</td>` +
                            `<td style="padding-right:4px;white-space:nowrap;width:1px"><input class="popup-edit-input sh-price" type="number" min="0" value="${it.Price}" style="width:68px" title="Price"></td>` +
                            `<td style="white-space:nowrap;width:1px"><input class="popup-edit-input sh-stock" type="number" min="1" max="255" value="${it.Stock}" style="width:48px" title="Stock (255 = unlimited)"> ` +
                            `<button class="popup-edit-btn danger sh-remove" data-idx="${idx}" style="padding:1px 5px" title="Remove item">✕</button></td>` +
                            `</tr>`;
                    }).join('');
                    const emptyZone = shop.items.length === 0
                        ? `<div class="npc-shop-view npc-shop-empty" style="min-height:44px;display:flex;align-items:center;justify-content:center;margin-top:6px;border-radius:4px"><span style="color:#888;font-size:11px;font-style:italic;pointer-events:none">Drop items here to add</span></div>`
                        : '';
                    const tblHtml = shop.items.length
                        ? `<div class="npc-shop-view" style="margin-top:4px;max-height:260px;overflow-y:auto;overflow-x:hidden">` +
                          `<table class="popup-edit-items-table sh-items-table" style="font-size:12px;border-collapse:collapse;line-height:1.9;width:100%">${rows}</table>` +
                          `</div>`
                        : '';
                    const footer = `<div class="popup-edit-row" style="margin-top:5px">` +
                        `<button class="popup-edit-btn" data-edit-action="sh-apply" title="Save price/stock changes">✔ Apply</button>` +
                        `</div>`;
                    return header + currencyLine + tblHtml + emptyZone + footer;
                }

                // ── View mode ────────────────────────────────────────────────────
                if (!shop.items.length) return header + '<br><span style="color:#888;font-size:11px">No inventory data</span>';
                const viewRows = shop.items.map(it => {
                    const entry    = itemNames[String(it.ItemId)];
                    const name     = entry?.name ?? `Item #${it.ItemId}`;
                    const href     = itemWikiHref(it.ItemId);
                    const nameLink = `<a href="${href}" target="_blank" style="color:inherit;text-decoration:none" onmouseover="this.style.textDecoration='underline'" onmouseout="this.style.textDecoration='none'">${name}</a>`;
                    const stock    = it.Stock === 255 ? '' : ` <span style="color:#888">(×${it.Stock})</span>`;
                    return `<tr><td style="color:#222;padding-right:8px;white-space:nowrap">${itemIcon(it)}${nameLink}</td>` +
                           `<td style="color:#555;white-space:nowrap">${it.Price} ${currency}${stock}</td></tr>`;
                }).join('');
                return header + currencyLine +
                    `<br><div style="max-height:280px;overflow-y:auto;overflow-x:hidden;margin-top:4px;min-width:280px">` +
                    `<table style="font-size:13px;border-collapse:collapse;line-height:1.8;width:100%">${viewRows}</table></div>`;
            };

            const shopCategory = classifyShopPoiCategory(funcId);
            const shopIconSrc  = poiMapIconSrc(shopCategory);
            const icon = shopIconSrc
                ? makePoiMapIcon(shopIconSrc, 22)
                : L.divIcon({
                    className: '',
                    html: `<div style="width:12px;height:12px;background:${color};border:2px solid #111;transform:rotate(45deg);box-shadow:0 0 3px rgba(0,0,0,0.6);"></div>`,
                    iconSize:    [12, 12],
                    iconAnchor:  [6, 6],
                    popupAnchor: [0, -10],
                });
            const marker = L.marker(latlng, { icon })
            .bindPopup(buildShopPopup(_shopCache), { minWidth: 340 })
            .bindTooltip(`${npcName} — ${funcLabel}`, { direction: 'top', offset: [0, -10] })
            .addTo(npcShopLayer);
            marker._poiCategory = shopCategory;
            _shopMarkerByNpcId.set(`${stageNo}:${npc.NpcId}`, marker);

            let _shopClickHandler = null;
            marker.on('popupclose', () => {
                if (_shopPopupDropFn?._marker === marker) _shopPopupDropFn = null;
            });
            marker.on('popupopen', function() {
                const popup = this.getPopup();
                const self  = this;

                const rebuildShopPopup = (cache) => {
                    const el = popup.getElement()?.querySelector('.leaflet-popup-content');
                    if (el) el.innerHTML = buildShopPopup(cache);
                    else { popup.setContent(buildShopPopup(cache)); popup.update(); }
                    wireShopEdit(cache);
                };

                const wireShopEdit = (cache) => {
                    const popupEl = popup.getElement();
                    if (!popupEl) return;

                    // Drag-to-reorder
                    if (_editMode && _attachDragReorder) {
                        const shop = cache?.get(npc.ShopId);
                        const tbl  = popupEl.querySelector('.sh-items-table');
                        if (tbl && shop?.items?.length) {
                            _attachDragReorder(tbl, shop.items, () => {
                                if (_markDirty) _markDirty('ddon-src-shop');
                                rebuildShopPopup(cache);
                            });
                        }
                    }

                    if (!_editMode) return;

                    // Click handler: apply / remove
                    if (_shopClickHandler) popupEl.removeEventListener('click', _shopClickHandler);
                    _shopClickHandler = (e) => {
                        const actionBtn = e.target.closest('[data-edit-action]');
                        const removeBtn = e.target.closest('.sh-remove');
                        if (!actionBtn && !removeBtn) return;
                        e.stopPropagation();
                        const shop = cache?.get(npc.ShopId);
                        if (!shop) return;
                        if (removeBtn) {
                            const idx = parseInt(removeBtn.dataset.idx);
                            shop.items.splice(idx, 1);
                            if (_markDirty) _markDirty('ddon-src-shop');
                            rebuildShopPopup(cache);
                        } else if (actionBtn.dataset.editAction === 'sh-apply') {
                            const tbl = popupEl.querySelector('.sh-items-table');
                            if (!tbl) return;
                            tbl.querySelectorAll('tr[data-idx]').forEach(tr => {
                                const idx   = parseInt(tr.dataset.idx);
                                const price = Math.max(0, parseInt(tr.querySelector('.sh-price').value) || 0);
                                const stock = Math.min(255, Math.max(1, parseInt(tr.querySelector('.sh-stock').value) || 255));
                                if (shop.items[idx]) {
                                    shop.items[idx].Price = price;
                                    shop.items[idx].Stock = stock;
                                }
                            });
                            if (_markDirty) _markDirty('ddon-src-shop');
                            rebuildShopPopup(cache);
                        }
                    };
                    popupEl.addEventListener('click', _shopClickHandler);

                    // Register drop function for Items panel drag-drop
                    const doAddDraggedItem = (itemId) => {
                        const shop = cache?.get(npc.ShopId);
                        if (!shop) return;
                        const defaultPrice = itemNames[String(itemId)]?.sellPrice ?? 0;
                        shop.items.push({ ItemId: itemId, Price: defaultPrice, Stock: 255 });
                        if (_markDirty) _markDirty('ddon-src-shop');
                        rebuildShopPopup(cache);
                        const flashEl = self.getPopup()?.getElement()?.querySelector('.npc-shop-view');
                        if (flashEl) {
                            flashEl.style.transition = 'background 0.1s';
                            flashEl.style.background = 'rgba(80,200,120,0.25)';
                            setTimeout(() => { flashEl.style.background = ''; }, 700);
                        }
                    };
                    doAddDraggedItem._marker = marker;
                    _shopPopupDropFn = doAddDraggedItem;
                };

                if (_shopCache) { rebuildShopPopup(_shopCache); return; }
                _shopPromise.then(cache => { if (self.isPopupOpen()) rebuildShopPopup(cache); });
            });
        }
    }
    applyAllPoiVisibility();
}

function loadSpecialShops(info, stid = null) {
    specialShopLayer.clearLayers();
    _specialShopMarkerByNpcId.clear();
    if (!info.stages?.length) return;

    const floorObbs     = info.floor_obbs ?? null;
    const filterByFloor = floorObbs !== null;
    const stagesToLoad  = (stid && info.stages.includes(stid)) ? [stid] : info.stages;

    for (const stageId of stagesToLoad) {
        const stageNo = String(parseInt(stageId.slice(2), 10));
        if (!isAccessibleShopStage(stageNo)) continue;
        const npcs    = npcSpecialShops[stageNo];
        if (!npcs) continue;

        for (const npc of npcs) {
            if (filterByFloor) {
                const floor = getEnemyFloor(npc.Position.x, npc.Position.y, npc.Position.z, floorObbs);
                if (floor !== null && floor !== currentLayer) continue;
            }
            const latlng   = worldToPixel(npc.Position.x, npc.Position.z, info);
            const shopType = npc.ShopTypeName ?? npc.ShopType;   // string for cache lookup
            const typeId   = npc.ShopType ?? SHOP_TYPE_IDS[shopType] ?? 0;
            const npcName  = npcNames[String(npc.NpcId)]?.name ?? npc.NpcName ?? `NPC #${npc.NpcId}`;
            const color    = '#c084fc';

            // ── Per-marker popup state ────────────────────────────────────────
            let _sspClickHandler = null;
            let _selCat = 0, _selAp = 0;
            let _collapsedCats   = new Set();
            let _searchQuery     = '';
            let _popupAbort      = null;   // AbortController for per-popup listeners

            // ── Shared helpers ────────────────────────────────────────────────
            const itemIcon = (itemId) => {
                const entry    = itemNames[String(itemId)];
                const iconNo   = entry?.iconNo;
                const iconFile = iconNo != null ? `ii${String(iconNo).padStart(6, '0')}.png` : null;
                return iconFile && _iconIdSet.has(iconNo)
                    ? `<img src="images/icons/small/${iconFile}" width="18" height="18" style="vertical-align:middle;margin-right:3px;image-rendering:pixelated;flex-shrink:0">`
                    : `<span style="display:inline-block;width:18px;flex-shrink:0;margin-right:3px"></span>`;
            };

            const itemRef = (itemId, name) => {
                const href = itemWikiHref(itemId);
                return `<a href="${href}" target="_blank" style="color:inherit;text-decoration:none" onmouseover="this.style.textDecoration='underline'" onmouseout="this.style.textDecoration='none'">${name}</a>`;
            };

            // ── Crest badge rendering ─────────────────────────────────────────
            const crestRef = (crestId) => {
                const name = itemNames[String(crestId)]?.name ?? `#${crestId}`;
                const href = itemWikiHref(crestId);
                return `<a href="${href}" target="_blank" style="color:inherit;text-decoration:underline dotted">${name}</a>`;
            };

            const buildCrestBadges = (crests) => {
                if (!crests?.length) return '';
                return crests.map(cr => {
                    switch (cr.type) {
                        case 'Imbued':
                            return `<span class="ssp-crest-badge ssp-crest-imbued">✦ ${crestRef(cr.crest_id)} lv.${cr.amount ?? '?'}</span>`;
                        case 'CrestLottery': {
                            const ids  = cr.values ?? [];
                            const pills = ids.map(id =>
                                `<span class="ssp-crest-pill">${crestRef(id)}</span>`
                            ).join('');
                            return `<details class="ssp-crest-badge ssp-crest-lottery ssp-crest-expand">` +
                                `<summary>🎲 <span class="ssp-more-hint">Random (${ids.length})</span></summary>` +
                                `<div class="ssp-crest-rest">${pills}</div>` +
                                `</details>`;
                        }
                        case 'DragonTrinketAlpha':
                            return `<span class="ssp-crest-badge ssp-crest-dragon">α slot — ${cr.job_id ?? '?'}</span>`;
                        case 'DragonTrinketBeta':
                            return `<span class="ssp-crest-badge ssp-crest-dragon">β slot — ${cr.job_id ?? '?'}</span>`;
                        case 'BitterblackBracelet':
                            return `<span class="ssp-crest-badge ssp-crest-bbm">BBM Bracelet</span>`;
                        case 'BitterBlackEarring':
                            return `<span class="ssp-crest-badge ssp-crest-bbm">BBM Earring</span>`;
                        default:
                            return `<span class="ssp-crest-badge">${cr.type}</span>`;
                    }
                }).join(' ');
            };

            // ── Left panel: appraisal list ────────────────────────────────────
            const buildListPanel = (categories, sCat, sAp) => {
                const catBlocks = categories.map((cat, catIdx) => {
                    const isCollapsed = _collapsedCats.has(catIdx);
                    const arrow = isCollapsed ? '▶' : '▼';
                    const catHeader = _editMode
                        ? `<div class="ssp-list-cat-header" data-cat="${catIdx}"><span class="ssp-cat-toggle">${arrow}</span>` +
                          `<input class="popup-edit-input ssp-cat-label" value="${cat.label.replace(/"/g, '&quot;')}" style="flex:1;font-size:11px;font-weight:bold;min-width:0" placeholder="Category" data-cat="${catIdx}">` +
                          `<button class="popup-edit-btn danger ssp-remove-cat" data-cat="${catIdx}" style="padding:0 4px;font-size:10px;flex-shrink:0">✕</button></div>`
                        : `<div class="ssp-list-cat-header" data-cat="${catIdx}"><span class="ssp-cat-toggle">${arrow}</span>${cat.label}</div>`;

                    const apRows = cat.appraisals.map((ap, apIdx) => {
                        const isSelected = catIdx === sCat && apIdx === sAp;
                        const costIcons  = ap.base_items.slice(0, 2).map(bi =>
                            `<span style="display:inline-flex;align-items:center;white-space:nowrap">${itemIcon(bi.item_id)}<span style="font-size:10px;color:#888">×${bi.amount}</span></span>`
                        ).join('<span style="color:#bbb;font-size:9px;margin:0 1px">+</span>');
                        const costMore = ap.base_items.length > 2
                            ? `<span style="font-size:9px;color:#999"> +${ap.base_items.length - 2}</span>` : '';
                        const rewardSummary = ap.pool.length === 1
                            ? `<span style="font-size:10px;color:#333;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0" title="${ap.pool[0].name.replace(/"/g, '&quot;')}">${ap.pool[0].name}</span>`
                            : `<span class="ssp-lottery-badge">${ap.pool.length} possible</span>`;
                        const editBtn = _editMode
                            ? `<button class="popup-edit-btn danger ssp-remove-ap" data-cat="${catIdx}" data-ap="${apIdx}" style="padding:0 3px;font-size:10px;flex-shrink:0;margin-left:2px" title="Remove">✕</button>`
                            : '';
                        // Build searchable text: ap label + all item names (lowercased)
                        const searchText = [
                            ap.label,
                            ...ap.base_items.map(bi => bi.name),
                            ...ap.pool.map(pi => pi.name),
                        ].join(' ').toLowerCase();
                        return `<div class="ssp-ap-row${isSelected ? ' selected' : ''}" data-cat="${catIdx}" data-ap="${apIdx}" data-search="${searchText.replace(/"/g, '&quot;')}">` +
                            `<div class="ssp-row-cost">${costIcons}${costMore}</div>` +
                            `<span class="ssp-row-arrow">→</span>` +
                            `<div class="ssp-row-reward">${rewardSummary}</div>` +
                            editBtn +
                            `</div>`;
                    }).join('');

                    const addApBtn = _editMode
                        ? `<button class="popup-edit-btn ssp-add-ap" data-cat="${catIdx}" style="width:100%;font-size:10px;padding:2px;margin:2px 0">+ Appraisal</button>`
                        : '';
                    const rowsHtml = `<div class="ssp-cat-rows"${isCollapsed ? ' style="display:none"' : ''}>${apRows}${addApBtn}</div>`;
                    return `<div class="ssp-cat-block" data-cat="${catIdx}">${catHeader}${rowsHtml}</div>`;
                }).join('');

                const addCatBtn = _editMode
                    ? `<div style="padding:3px 4px"><button class="popup-edit-btn ssp-add-cat" style="width:100%;font-size:10px;padding:3px">+ Category</button></div>`
                    : '';
                return `<div class="ssp-list-panel">` +
                    `<div class="ssp-search-wrap"><input class="ssp-search-input" type="search" placeholder="Search items…" value="${_searchQuery.replace(/"/g, '&quot;')}"></div>` +
                    `<div class="ssp-list-scroll">${catBlocks}${addCatBtn}</div>` +
                    `</div>`;
            };

            // ── Right panel: appraisal detail ─────────────────────────────────
            const buildDetailPanel = (ap, catIdx, apIdx) => {
                if (!ap) return `<div class="ssp-detail-panel"><span style="color:#999;font-size:11px;padding:8px;display:block">← Select an appraisal</span></div>`;

                // Header: label (editable in edit mode)
                const headerHtml = _editMode
                    ? `<div style="display:flex;align-items:center;gap:4px;margin-bottom:8px">` +
                      `<input class="popup-edit-input ssp-ap-label" value="${ap.label.replace(/"/g, '&quot;')}" style="flex:1;font-size:11px" placeholder="Appraisal label" data-cat="${catIdx}" data-ap="${apIdx}">` +
                      `<button class="popup-edit-btn" data-edit-action="ssp-apply" style="padding:2px 6px;font-size:10px;flex-shrink:0">✔ Apply</button>` +
                      `</div>`
                    : `<div class="ssp-detail-label">${ap.label}</div>`;

                // Cost rows
                const costRows = ap.base_items.map((bi, biIdx) => _editMode
                    ? `<div style="display:flex;align-items:center;gap:4px;padding:2px 0">` +
                      `${itemIcon(bi.item_id)}<span style="font-size:11px;flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${bi.name.replace(/"/g, '&quot;')}">${bi.name}</span>` +
                      `<input class="popup-edit-input ssp-amount" type="number" min="1" value="${bi.amount}" style="width:44px" data-side="base" data-cat="${catIdx}" data-ap="${apIdx}" data-idx="${biIdx}">` +
                      `<button class="popup-edit-btn danger ssp-remove-item" data-side="base" data-cat="${catIdx}" data-ap="${apIdx}" data-idx="${biIdx}" style="padding:0 4px">✕</button>` +
                      `</div>`
                    : `<div style="display:flex;align-items:center;gap:4px;padding:2px 0">` +
                      `${itemIcon(bi.item_id)}<span style="font-size:11px">${itemRef(bi.item_id, bi.name)}</span>` +
                      `<span style="font-size:11px;color:#888;margin-left:auto;white-space:nowrap">×${bi.amount}</span>` +
                      `</div>`
                ).join('');
                const dropBase = _editMode
                    ? `<div class="ssp-drop-zone" data-side="base" data-cat="${catIdx}" data-ap="${apIdx}">Drop item here</div>` : '';

                // Pool label
                const isLottery  = ap.pool.length > 1;
                const poolHeader = isLottery
                    ? `<div class="ssp-section-title">Reward <span class="ssp-lottery-badge" style="font-size:10px;font-weight:normal">1 of ${ap.pool.length}</span></div>`
                    : `<div class="ssp-section-title">Reward</div>`;

                // Pool rows
                const poolRows = ap.pool.map((pi, piIdx) => {
                    if (_editMode) {
                        const crests = pi.crests ?? [];
                        const crestRows = crests.map((cr, ciIdx) => {
                            const ds = `data-cat="${catIdx}" data-ap="${apIdx}" data-pi="${piIdx}" data-ci="${ciIdx}"`;
                            const rmBtn = `<button class="popup-edit-btn danger ssp-remove-crest" ${ds} style="padding:0 4px;font-size:10px;flex-shrink:0">✕</button>`;
                            if (cr.type === 'Imbued') {
                                const cName = itemNames[String(cr.crest_id)]?.name ?? `#${cr.crest_id}`;
                                return `<div style="display:flex;align-items:center;gap:3px;padding:1px 0 1px 21px;font-size:10px">` +
                                    `<span style="color:#a78bfa;flex-shrink:0">✦</span>` +
                                    `<span class="ssp-crest-drop" ${ds} title="Drop item to set crest" style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;padding:1px 4px;border:1px dashed rgba(167,139,250,0.35);border-radius:3px;cursor:default">${cName}</span>` +
                                    `<input class="popup-edit-input ssp-crest-id" type="number" min="1" value="${cr.crest_id ?? ''}" style="width:58px" placeholder="id" ${ds}>` +
                                    `<span style="color:#888;flex-shrink:0">lv.</span>` +
                                    `<input class="popup-edit-input ssp-crest-amount" type="number" min="1" value="${cr.amount ?? 1}" style="width:44px" ${ds}>` +
                                    rmBtn + `</div>`;
                            }
                            if (cr.type === 'CrestLottery') {
                                return `<div style="display:flex;align-items:center;gap:3px;padding:1px 0 1px 21px;font-size:10px">` +
                                    `<span style="flex-shrink:0">🎲</span>` +
                                    `<input class="popup-edit-input ssp-crest-values" value="${(cr.values ?? []).join(',')}" style="flex:1;min-width:0" placeholder="id1,id2,..." ${ds}>` +
                                    rmBtn + `</div>`;
                            }
                            return `<div style="display:flex;align-items:center;gap:3px;padding:1px 0 1px 21px;font-size:10px">` +
                                `<span style="color:#888;flex:1">${cr.type}</span>` + rmBtn + `</div>`;
                        }).join('');
                        const addCrestBtn = `<div style="padding:2px 0 2px 21px">` +
                            `<button class="popup-edit-btn ssp-add-crest" data-cat="${catIdx}" data-ap="${apIdx}" data-pi="${piIdx}" style="font-size:10px;padding:1px 6px">+ Imbued</button>` +
                            `</div>`;
                        return `<div style="border-top:1px solid rgba(0,0,0,0.07)">` +
                            `<div style="display:flex;align-items:center;gap:4px;padding:3px 0">` +
                            `${itemIcon(pi.item_id)}<span style="font-size:11px;flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${pi.name.replace(/"/g, '&quot;')}">${pi.name}</span>` +
                            `<input class="popup-edit-input ssp-amount" type="number" min="1" value="${pi.amount}" style="width:44px" data-side="pool" data-cat="${catIdx}" data-ap="${apIdx}" data-idx="${piIdx}">` +
                            `<button class="popup-edit-btn danger ssp-remove-item" data-side="pool" data-cat="${catIdx}" data-ap="${apIdx}" data-idx="${piIdx}" style="padding:0 4px">✕</button>` +
                            `</div>` +
                            crestRows + addCrestBtn +
                            `</div>`;
                    }
                    const crestHtml = buildCrestBadges(pi.crests);
                    return `<div class="ssp-pool-item">` +
                        `<div style="display:flex;align-items:center;gap:4px">` +
                        `${itemIcon(pi.item_id)}<span style="font-size:11px;flex:1;min-width:0" title="${pi.name.replace(/"/g, '&quot;')}">${itemRef(pi.item_id, pi.name)}</span>` +
                        `<span style="font-size:11px;color:#888;white-space:nowrap;margin-left:4px">×${pi.amount}</span>` +
                        `</div>` +
                        (crestHtml ? `<div style="padding-left:21px;margin-top:3px">${crestHtml}</div>` : '') +
                        `</div>`;
                }).join('');
                const dropPool = _editMode
                    ? `<div class="ssp-drop-zone" data-side="pool" data-cat="${catIdx}" data-ap="${apIdx}">Drop item here</div>` : '';

                return `<div class="ssp-detail-panel">` +
                    headerHtml +
                    `<div class="ssp-section-title">Cost</div>` +
                    `<div style="margin-bottom:10px">${costRows || '<span style="color:#999;font-size:11px">Nothing</span>'}${dropBase}</div>` +
                    poolHeader +
                    `<div>${poolRows || '<span style="color:#999;font-size:11px">No rewards</span>'}${dropPool}</div>` +
                    `</div>`;
            };

            // ── Full popup HTML ───────────────────────────────────────────────
            const buildSpecialShopPopup = (cache, sCat, sAp) => {
                const header = `<div class="ssp-popup-header">` +
                    `<span class="ssp-badge">Appraisals</span>` +
                    `<strong style="font-size:12px;margin-left:6px">${npcName}</strong>` +
                    `<span style="color:#888;font-size:11px;margin-left:6px">${shopType} (ID: ${typeId})</span>` +
                    `</div>`;

                if (!cache) return header + `<div style="padding:10px;color:#888;font-size:11px">Loading…</div>`;

                let shopData = cache.get(shopType);
                if (!shopData) {
                    if (!_editMode) return header + `<div style="padding:10px;color:#888;font-size:11px">No data for "${shopType}"</div>`;
                    // Bootstrap an empty shop so the editor can build from scratch
                    const rawShop = { shop_type: shopType, categories: [] };
                    shopData = { shopTypeId: typeId, categories: rawShop.categories, rawShop };
                    cache.set(shopType, shopData);
                    if (_rawSpecialShopData) {
                        if (!_rawSpecialShopData.shops) _rawSpecialShopData.shops = [];
                        _rawSpecialShopData.shops.push(rawShop);
                    }
                }

                const { categories } = shopData;
                if (!categories.length && !_editMode) return header + `<div style="padding:10px;color:#888;font-size:11px">No categories</div>`;

                // Clamp selection to valid range
                const cat0 = Math.min(sCat, categories.length - 1);
                const ap0  = Math.min(sAp,  Math.max(0, (categories[cat0]?.appraisals.length ?? 1) - 1));
                const selAp = categories[cat0]?.appraisals[ap0];

                return header +
                    `<div class="ssp-split-panel">` +
                    buildListPanel(categories, cat0, ap0) +
                    buildDetailPanel(selAp, cat0, ap0) +
                    `</div>`;
            };

            const shopIconSrc = poiMapIconSrc('shop');
            const icon = shopIconSrc
                ? makeExchangeShopMapIcon(shopIconSrc, 22)
                : L.divIcon({
                    className: '',
                    html: `<div style="width:12px;height:12px;background:${color};border:2px solid #111;transform:rotate(45deg);box-shadow:0 0 4px rgba(192,132,252,0.7);"></div>`,
                    iconSize:    [12, 12],
                    iconAnchor:  [6, 6],
                    popupAnchor: [0, -10],
                });
            const marker = L.marker(latlng, { icon })
                .bindPopup(buildSpecialShopPopup(_specialShopCache, 0, 0), { minWidth: 600 })
                .bindTooltip(`${npcName} — Appraisals (${shopType})`, { direction: 'top', offset: [0, -10] })
                .addTo(specialShopLayer);

            marker._poiCategory = 'shop';
            marker._poiExchangeShop = true;
            _specialShopMarkerByNpcId.set(`${stageNo}:${npc.NpcId}`, marker);

            marker.on('popupclose', () => {
                if (_shopPopupDropFn?._marker === marker) _shopPopupDropFn = null;
                _popupAbort?.abort();
                _popupAbort = null;
            });

            marker.on('popupopen', function() {
                const popup = this.getPopup();
                const self  = this;

                const rebuildPopup = (cache) => {
                    const el = popup.getElement()?.querySelector('.leaflet-popup-content');
                    if (el) el.innerHTML = buildSpecialShopPopup(cache, _selCat, _selAp);
                    else { popup.setContent(buildSpecialShopPopup(cache, _selCat, _selAp)); popup.update(); }
                    wireInteractions(cache);
                };

                // Swap only the detail panel — preserves list scroll position
                const swapDetail = (cache, catIdx, apIdx) => {
                    const popupEl = popup.getElement();
                    const old = popupEl?.querySelector('.ssp-detail-panel');
                    if (!old) return;
                    const shopD = cache?.get(shopType);
                    const ap    = shopD?.categories[catIdx]?.appraisals[apIdx];
                    const tmp   = document.createElement('template');
                    tmp.innerHTML = buildDetailPanel(ap, catIdx, apIdx);
                    old.replaceWith(tmp.content.firstElementChild);
                    wireDropZones(cache);
                };

                const wireDropZones = (cache) => {
                    if (!_editMode) return;
                    const popupEl = popup.getElement();
                    popupEl?.querySelectorAll('.ssp-drop-zone').forEach(zone => {
                        zone.addEventListener('dragover',  e => { e.preventDefault(); zone.style.background = 'rgba(192,132,252,0.15)'; });
                        zone.addEventListener('dragleave', ()  => { zone.style.background = ''; });
                        zone.addEventListener('drop', e => {
                            e.preventDefault();
                            zone.style.background = '';
                            if (_dragItemId == null) return;
                            const shopD  = cache?.get(shopType);
                            if (!shopD) return;
                            const ap = shopD.categories[parseInt(zone.dataset.cat)]?.appraisals[parseInt(zone.dataset.ap)];
                            if (!ap) return;
                            const arr   = zone.dataset.side === 'base' ? ap.base_items : ap.pool;
                            const entry = itemNames[String(_dragItemId)];
                            arr.push({ item_id: _dragItemId, name: entry?.name ?? `Item #${_dragItemId}`, amount: 1 });
                            if (_markDirty) _markDirty('ddon-src-special-shop');
                            rebuildPopup(cache);
                        });
                    });
                    popupEl?.querySelectorAll('.ssp-crest-drop').forEach(zone => {
                        zone.addEventListener('dragover',  e => { e.preventDefault(); zone.style.background = 'rgba(192,132,252,0.15)'; });
                        zone.addEventListener('dragleave', ()  => { zone.style.background = ''; });
                        zone.addEventListener('drop', e => {
                            e.preventDefault();
                            zone.style.background = '';
                            if (_dragItemId == null) return;
                            const shopD = cache?.get(shopType);
                            if (!shopD) return;
                            const { cat: cI, ap: aI, pi: pI, ci } = zone.dataset;
                            const cr = shopD.categories[cI]?.appraisals[aI]?.pool[pI]?.crests?.[ci];
                            if (!cr) return;
                            cr.crest_id = _dragItemId;
                            if (_markDirty) _markDirty('ddon-src-special-shop');
                            rebuildPopup(cache);
                        });
                    });
                };

                // Apply search filter to visible list rows (no rebuild needed)
                const applySearch = (popupEl) => {
                    const q = _searchQuery.trim().toLowerCase();
                    popupEl.querySelectorAll('.ssp-cat-block').forEach(block => {
                        let visibleRows = 0;
                        block.querySelectorAll('.ssp-ap-row').forEach(row => {
                            const match = !q || row.dataset.search.includes(q);
                            row.style.display = match ? '' : 'none';
                            if (match) visibleRows++;
                        });
                        // When searching, override collapse; when not, respect it
                        const catIdx  = parseInt(block.dataset.cat);
                        const catRows = block.querySelector('.ssp-cat-rows');
                        if (catRows) catRows.style.display = (q || !_collapsedCats.has(catIdx)) ? '' : 'none';
                        block.style.display = (q && visibleRows === 0) ? 'none' : '';
                    });
                };

                // Single delegated click handler for the whole popup
                const wireInteractions = (cache) => {
                    const popupEl = popup.getElement();
                    if (!popupEl) return;

                    // ── Prevent popup closing on text-selection drag out ──────
                    // When a drag starts inside the popup (mousedown) and the
                    // mouseup fires outside it, the browser synthesises a click
                    // on the map element.  Leaflet's bubble-phase map-click
                    // handler then closes the popup.  Fix: intercept that click
                    // in the capture phase — before Leaflet ever sees it — and
                    // swallow it when we know the drag originated inside.
                    if (!_popupAbort) {
                        _popupAbort = new AbortController();
                        const { signal } = _popupAbort;
                        let _dragFromPopup = false;
                        popupEl.addEventListener('mousedown', () => { _dragFromPopup = true; }, { signal });
                        // Mouseup inside the popup = normal click; clear immediately so the
                        // click event is not intercepted below.
                        popupEl.addEventListener('mouseup', () => { _dragFromPopup = false; }, { signal });
                        // Mouseup outside the popup = drag released outside; the deferred
                        // clear fires after click, so the capture handler can swallow it.
                        document.addEventListener('mouseup', () => { setTimeout(() => { _dragFromPopup = false; }, 0); }, { signal });
                        document.addEventListener('click', (e) => {
                            if (_dragFromPopup) { e.stopPropagation(); _dragFromPopup = false; }
                        }, { capture: true, signal });
                    }

                    // ── Search input ──────────────────────────────────────────
                    const searchInput = popupEl.querySelector('.ssp-search-input');
                    if (searchInput) {
                        searchInput.addEventListener('input', () => {
                            _searchQuery = searchInput.value;
                            applySearch(popupEl);
                        });
                        // Restore filter state after rebuild
                        if (_searchQuery) applySearch(popupEl);
                    }

                    if (_sspClickHandler) popupEl.removeEventListener('click', _sspClickHandler);
                    _sspClickHandler = (e) => {
                        const shopD = cache?.get(shopType);
                        if (!shopD) return;

                        // ── Category collapse toggle ──────────────────────────
                        const catHeader = e.target.closest('.ssp-list-cat-header');
                        if (catHeader && !e.target.closest('input') && !e.target.closest('button')) {
                            const catIdx  = parseInt(catHeader.dataset.cat);
                            const block   = catHeader.closest('.ssp-cat-block');
                            const catRows = block?.querySelector('.ssp-cat-rows');
                            if (_collapsedCats.has(catIdx)) {
                                _collapsedCats.delete(catIdx);
                                if (catRows) catRows.style.display = '';
                                catHeader.querySelector('.ssp-cat-toggle').textContent = '▼';
                            } else {
                                _collapsedCats.add(catIdx);
                                if (catRows) catRows.style.display = 'none';
                                catHeader.querySelector('.ssp-cat-toggle').textContent = '▶';
                            }
                            return;
                        }

                        // ── Row selection (view + edit) ───────────────────────
                        const apRow = e.target.closest('.ssp-ap-row');
                        if (apRow && !e.target.closest('button')) {
                            const catIdx = parseInt(apRow.dataset.cat);
                            const apIdx  = parseInt(apRow.dataset.ap);
                            if (catIdx === _selCat && apIdx === _selAp) return;
                            _selCat = catIdx;
                            _selAp  = apIdx;
                            popupEl.querySelectorAll('.ssp-ap-row').forEach(r =>
                                r.classList.toggle('selected', r === apRow));
                            swapDetail(cache, catIdx, apIdx);
                            return;
                        }

                        if (!_editMode) return;

                        // ── Edit actions ──────────────────────────────────────
                        const removeItem  = e.target.closest('.ssp-remove-item');
                        const removeAp    = e.target.closest('.ssp-remove-ap');
                        const removeCat   = e.target.closest('.ssp-remove-cat');
                        const addAp       = e.target.closest('.ssp-add-ap');
                        const addCat      = e.target.closest('.ssp-add-cat');
                        const applyBtn    = e.target.closest('[data-edit-action="ssp-apply"]');
                        const removeCrest = e.target.closest('.ssp-remove-crest');
                        const addCrest    = e.target.closest('.ssp-add-crest');
                        if (!removeItem && !removeAp && !removeCat && !addAp && !addCat && !applyBtn && !removeCrest && !addCrest) return;
                        e.stopPropagation();

                        if (removeItem) {
                            const { side, cat: cI, ap: aI, idx } = removeItem.dataset;
                            const ap  = shopD.categories[cI]?.appraisals[aI];
                            if (!ap) return;
                            (side === 'base' ? ap.base_items : ap.pool).splice(parseInt(idx), 1);
                        } else if (removeAp) {
                            const cI = parseInt(removeAp.dataset.cat), aI = parseInt(removeAp.dataset.ap);
                            shopD.categories[cI]?.appraisals.splice(aI, 1);
                            if (cI === _selCat && aI === _selAp) _selAp = Math.max(0, _selAp - 1);
                        } else if (removeCat) {
                            shopD.categories.splice(parseInt(removeCat.dataset.cat), 1);
                            _selCat = Math.max(0, _selCat - 1);
                            _selAp  = 0;
                        } else if (addAp) {
                            const cI = parseInt(addAp.dataset.cat);
                            shopD.categories[cI]?.appraisals.push({ label: 'New Appraisal', base_items: [], pool: [] });
                            _selCat = cI;
                            _selAp  = shopD.categories[cI].appraisals.length - 1;
                        } else if (addCat) {
                            shopD.categories.push({ label: 'New Category', appraisals: [] });
                            _selCat = shopD.categories.length - 1;
                            _selAp  = 0;
                        } else if (removeCrest) {
                            const { cat: cI, ap: aI, pi: pI, ci } = removeCrest.dataset;
                            const pool = shopD.categories[cI]?.appraisals[aI]?.pool;
                            if (pool?.[pI]?.crests) pool[pI].crests.splice(parseInt(ci), 1);
                        } else if (addCrest) {
                            const { cat: cI, ap: aI, pi: pI } = addCrest.dataset;
                            const pool = shopD.categories[cI]?.appraisals[aI]?.pool;
                            if (pool?.[pI]) {
                                if (!pool[pI].crests) pool[pI].crests = [];
                                pool[pI].crests.push({ type: 'Imbued', crest_id: 1, amount: 1 });
                            }
                        } else if (applyBtn) {
                            popupEl.querySelectorAll('.ssp-cat-label').forEach(inp => {
                                const cI = parseInt(inp.dataset.cat);
                                if (shopD.categories[cI]) shopD.categories[cI].label = inp.value;
                            });
                            popupEl.querySelectorAll('.ssp-ap-label').forEach(inp => {
                                const ap = shopD.categories[inp.dataset.cat]?.appraisals[inp.dataset.ap];
                                if (ap) ap.label = inp.value;
                            });
                            popupEl.querySelectorAll('.ssp-amount').forEach(inp => {
                                const { side, cat: cI, ap: aI, idx } = inp.dataset;
                                const ap  = shopD.categories[cI]?.appraisals[aI];
                                if (!ap) return;
                                const arr = side === 'base' ? ap.base_items : ap.pool;
                                if (arr[idx]) arr[idx].amount = Math.max(1, parseInt(inp.value) || 1);
                            });
                            popupEl.querySelectorAll('.ssp-crest-id').forEach(inp => {
                                const { cat: cI, ap: aI, pi: pI, ci } = inp.dataset;
                                const cr = shopD.categories[cI]?.appraisals[aI]?.pool[pI]?.crests?.[ci];
                                const val = parseInt(inp.value);
                                if (cr && !isNaN(val) && val > 0) cr.crest_id = val;
                            });
                            popupEl.querySelectorAll('.ssp-crest-amount').forEach(inp => {
                                const { cat: cI, ap: aI, pi: pI, ci } = inp.dataset;
                                const cr = shopD.categories[cI]?.appraisals[aI]?.pool[pI]?.crests?.[ci];
                                const val = parseInt(inp.value);
                                if (cr && !isNaN(val) && val >= 1) cr.amount = val;
                            });
                            popupEl.querySelectorAll('.ssp-crest-values').forEach(inp => {
                                const { cat: cI, ap: aI, pi: pI, ci } = inp.dataset;
                                const cr = shopD.categories[cI]?.appraisals[aI]?.pool[pI]?.crests?.[ci];
                                if (cr) cr.values = inp.value.split(',').map(s => parseInt(s.trim())).filter(n => !isNaN(n));
                            });
                        }

                        if (_markDirty) _markDirty('ddon-src-special-shop');
                        rebuildPopup(cache);
                    };
                    popupEl.addEventListener('click', _sspClickHandler);

                    wireDropZones(cache);

                    if (_editMode) {
                        const doAddDraggedItem = (itemId) => {};
                        doAddDraggedItem._marker = marker;
                        _shopPopupDropFn = doAddDraggedItem;
                    }
                };

                if (_specialShopCache) { rebuildPopup(_specialShopCache); return; }
                _specialShopPromise.then(cache => { if (self.isPopupOpen()) rebuildPopup(cache); });
            });
        }
    }
    applyAllPoiVisibility();
}

function loadStageLabels(info) {
    stageLabelsLayer.clearLayers();
    const labels = info.stage_labels;
    if (!labels?.length) return;
    for (const lbl of labels) {
        const latlng = worldToPixel(lbl.x, lbl.z, info);
        const r = lbl.radius ?? 0;
        const fs = r >= 50000 ? 18 : r >= 20000 ? 13 : 10;
        const opacity = r >= 50000 ? 0.85 : r >= 20000 ? 0.70 : 0.55;
        L.marker(latlng, {
            icon: L.divIcon({
                className: '',
                html: `<div class="stage-label" style="font-size:${fs}px;opacity:${opacity}">${lbl.name}</div>`,
                iconAnchor: [0, 0],
            }),
            interactive: false,
            zIndexOffset: -1000,
        }).addTo(stageLabelsLayer);
    }
}

function landmarkDisplayName(lm) {
    if (lm.spot_name_en) return lm.spot_name_en;
    return lm.type.replace('TYPE_', '').replace(/_/g, ' ');
}

function loadLandmarks(mapName, info) {
    landmarkLayer.clearLayers();
    const entries = landmarkData[mapName];
    if (!entries) return;

    for (const lm of entries) {
        if (HIDDEN_LANDMARK_TYPES.has(lm.type)) continue;
        if (isLandmarkCoveredByConnection(mapName, lm)) continue;
        const latlng = worldToPixel(lm.x, lm.z, info);
        const label = landmarkDisplayName(lm);
        const category = classifyLandmarkPoiCategory(lm.type, lm);
        const iconSrc  = poiMapIconSrc(category);
        let marker;
        if (iconSrc) {
            marker = L.marker(latlng, { icon: makePoiMapIcon(iconSrc) })
                .bindTooltip(label, { permanent: false, direction: 'top', offset: [0, -12] })
                .addTo(landmarkLayer);
        } else {
            const color = LANDMARK_COLORS[lm.type] ?? '#aaaaaa';
            marker = L.circleMarker(latlng, {
                color,
                fillColor: color,
                fillOpacity: 0.85,
                radius: 6,
                weight: 1.5,
            })
            .bindTooltip(label, { permanent: false, direction: 'top', offset: [0, -6] })
            .addTo(landmarkLayer);
        }
        marker._poiCategory = category;
    }
    applyAllPoiVisibility();
}

// ── Stage connection markers ───────────────────────────────────────────────────

/**
 * Return a { zoom, center } view for arriving at destMap coming from srcMap,
 * or null if no positioned reverse connection exists.
 * Uses the reverse connection entry in connectionData[destMap] that points
 * back to srcMap — its x/z is where the door sits on the destination map.
 */
function arrivalView(srcMap, destMap, srcStageNo = null, zoom = 1.5) {
    const destInfo = mapParams[destMap];
    if (!destInfo) return null;
    const destConns = connectionData[destMap] || [];
    // Prefer a match by both map and originating stage (handles multi-stage maps like
    // rm000_m03 where each stage is a different building at a different field location).
    const rev = destConns.find(c =>
        c.to_map === srcMap && c.x != null && c.z != null &&
        (srcStageNo == null || c.to_stage === srcStageNo)
    ) ?? destConns.find(c => c.to_map === srcMap && c.x != null && c.z != null);
    if (!rev) return null;
    return { zoom, center: worldToPixel(rev.x, rev.z, destInfo) };
}

function loadConnections(mapName, info) {
    connectionLayer.clearLayers();

    // Clear any previous unpositioned exits list
    const exitsPanel = document.getElementById('exits-panel');
    const exitsList  = document.getElementById('exits-list');
    if (exitsList) exitsList.innerHTML = '';

    const allEntries = connectionData[mapName];
    if (!allEntries) { if (exitsPanel) exitsPanel.style.display = 'none'; return; }

    // Filter by active stage: connections with a from_stage only show when that
    // stage is active (or when no stid is set, e.g. navigating directly by map name).
    const stid = currentStageName();
    const activeStageNo = stid ? parseInt(stid.slice(2), 10) : null;
    const stageFiltered = allEntries.filter(c =>
        (c.from_stage == null || activeStageNo == null || c.from_stage === activeStageNo)
        && !isSuppressedArkConnection(c)
    );

    // Proximity deduplication: drop duplicate markers at the same map spot.
    // Merges same-stage copies, repeated FAA rows, and variant stages (e.g. st0406
    // + st0437 both labeled "Volden Large Tunnels" ~180 units apart).
    const DEDUP_DIST = 500;
    const connectionLabelKey = (conn) => (conn.name_en ?? '').trim().toLowerCase();
    const isDuplicateConnection = (a, b) => {
        if (a.x == null || b.x == null) return false;
        if (Math.hypot(a.x - b.x, a.z - b.z) >= DEDUP_DIST) return false;
        if ((a.to_map ?? null) !== (b.to_map ?? null)) return false;
        if (a.to_stage === b.to_stage) return true;
        const keyA = connectionLabelKey(a);
        const keyB = connectionLabelKey(b);
        return !!(keyA && keyA === keyB);
    };
    const connectionKeepScore = (conn) => [
        conn.stp_approx ? 0 : 1,
        conn.to_map && mapParams[conn.to_map] ? 1 : 0,
        /_m00(?:_|$)/.test(conn.to_map ?? '') ? 1 : 0,
        -(conn.to_stage ?? 99999),
    ];
    const scoreGt = (a, b) => {
        for (let i = 0; i < a.length; i++) {
            if (a[i] > b[i]) return true;
            if (a[i] < b[i]) return false;
        }
        return false;
    };

    const entries = [];
    for (const conn of stageFiltered) {
        if (conn.x == null) { entries.push(conn); continue; }
        const dupeIdx = entries.findIndex(c => isDuplicateConnection(c, conn));
        if (dupeIdx < 0) {
            entries.push(conn);
        } else if (scoreGt(connectionKeepScore(conn), connectionKeepScore(entries[dupeIdx]))) {
            entries[dupeIdx] = conn;
        }
    }

    const unpositioned = [];
    const positioned = [];

    for (const conn of entries) {
        const navMap  = (conn.to_map && mapParams[conn.to_map]) ? conn.to_map : null;
        const hasMap  = !!navMap;
        const stageId = `st${String(conn.to_stage).padStart(4, '0')}`;
        const destName = connectionDisplayName(conn) + ` (${stageId})`;

        // Unpositioned connections (pd stage exits, etc.) go in the sidebar list
        if (conn.x == null || conn.z == null) {
            unpositioned.push({ navMap, hasMap, stageId, destName });
            continue;
        }

        // Floor filter: on multi-floor maps only show connections on the active floor
        if (info.floor_obbs) {
            const floor = getEnemyFloor(conn.x, conn.y ?? 0, conn.z, info.floor_obbs);
            if (floor !== null && floor !== currentLayer) {
                unpositioned.push({
                    navMap,
                    hasMap,
                    stageId,
                    destName: `${destName} (floor ${floor})`,
                    fromStage: conn.from_stage ?? null,
                    targetFloor: floor,
                    latlng: worldToPixel(conn.x, conn.z, info),
                });
                continue;
            }
        }

        positioned.push({
            conn,
            navMap,
            hasMap,
            stageId,
            destName,
            fromStage: conn.from_stage ?? null,
            latlng: worldToPixel(conn.x, conn.z, info),
            specialRule: matchSpecialConnection(mapName, conn),
            poiCat: getConnectionPoi(conn, mapName),
        });
    }

    const coordGroups = new Map();
    for (const item of positioned) {
        const key = connectionExactCoordKey(item.conn);
        if (!coordGroups.has(key)) coordGroups.set(key, []);
        coordGroups.get(key).push(item);
    }

    for (const group of coordGroups.values()) {
        const byDest = new Map();
        for (const item of group) {
            const dk = connectionDestKey(item.conn);
            if (!byDest.has(dk)) byDest.set(dk, item);
        }
        const choices = [...byDest.values()].sort((a, b) =>
            connectionDisplayName(a.conn).localeCompare(connectionDisplayName(b.conn))
            || a.stageId.localeCompare(b.stageId),
        );
        addConnectionMapMarker(mapName, choices[0].latlng, choices);
    }

    // Render unpositioned exits in the sidebar
    if (unpositioned.length) {
        for (const exit of unpositioned) {
            const { navMap, hasMap, stageId, destName, fromStage, targetFloor, latlng } = exit;
            const li = document.createElement('div');
            li.style.cssText = 'padding:2px 0;font-size:0.78rem;';
            if (hasMap) {
                const a = document.createElement('span');
                a.textContent = destName;
                a.style.cssText = 'color:#ff6b35;cursor:pointer;text-decoration:underline dotted;';
                a.addEventListener('click', () => {
                    if (targetFloor != null && targetFloor !== currentLayer) {
                        _switchToFloor(targetFloor, info);
                        if (latlng) leafletMap.setView(latlng, Math.max(leafletMap.getZoom(), 1.5));
                        return;
                    }
                    navigateTo(navMap, stageId, arrivalView(mapName, navMap, fromStage ?? null));
                });
                li.appendChild(a);
            } else {
                li.textContent = destName;
                li.style.color = '#666';
            }
            exitsList?.appendChild(li);
        }
        if (exitsPanel) exitsPanel.style.display = '';
    } else if (exitsPanel) {
        exitsPanel.style.display = 'none';
    }
    applyAllPoiVisibility();
}

// ── Grid overlay ──────────────────────────────────────────────────────────────
// Grid unit = 10 pixels.  Major lines every GRID_MAJOR units.
// Coordinate displayed: gx = floor(pixel_x / 10), gy = floor(pixel_y_from_top / 10)
const GRID_UNIT  = 10;
const GRID_MAJOR = 50;   // major grid line every 50 units = 500 px

function pixelToGrid(px, py, imgH) {
    return [Math.floor(px / GRID_UNIT), Math.floor((imgH - py) / GRID_UNIT)];
}

function loadGrid(info) {
    gridLayer.clearLayers();
    const lineStyle   = { color: '#ffffff', weight: 0.5, opacity: 0.2, interactive: false };
    const maxGX = Math.floor(info.img_width  / GRID_UNIT);
    const maxGY = Math.floor(info.img_height / GRID_UNIT);

    // Vertical lines + X-axis labels (along top edge)
    for (let gx = 0; gx <= maxGX; gx += GRID_MAJOR) {
        const px = gx * GRID_UNIT;
        L.polyline([xy(px, 0), xy(px, info.img_height)], lineStyle).addTo(gridLayer);
        L.marker(xy(px, info.img_height), {
            icon: L.divIcon({
                className: 'grid-label',
                html: `${gx}`,
                iconSize: [40, 14],
                iconAnchor: [-2, 7],
            }),
            interactive: false,
        }).addTo(gridLayer);
    }

    // Horizontal lines + Y-axis labels (along left edge)
    for (let gy = 0; gy <= maxGY; gy += GRID_MAJOR) {
        const py = info.img_height - gy * GRID_UNIT;   // Leaflet y from bottom
        L.polyline([xy(0, py), xy(info.img_width, py)], lineStyle).addTo(gridLayer);
        L.marker(xy(0, py), {
            icon: L.divIcon({
                className: 'grid-label',
                html: `${gy}`,
                iconSize: [40, 14],
                iconAnchor: [42, 7],
            }),
            interactive: false,
        }).addTo(gridLayer);
    }
}

// ── Floor selector ────────────────────────────────────────────────────────────
let currentLayer = 0;

// ── Stage groups panel (editor mode) ──────────────────────────────────────────
let _sgpCollapsed     = true;
let _wfCollapsed      = true;
let _wfQuestCollapsed = true;

function buildStageGroupsPanel(info, stid = null) {
    const panel = document.getElementById('stage-groups-panel');
    if (!panel) return;

    // Collect stage numbers to show
    const stagesToShow = (stid && info.stages?.includes(stid))
        ? [stid]
        : (info.stages ?? []);

    const sgpRows   = [];
    const wfRows    = [];   // combined: qst entries tagged _src:'qst', named entries tagged _src:'named'
    const stageNoSet = new Set();
    for (const sid of stagesToShow) {
        const sno = String(parseInt(sid.slice(2), 10));
        stageNoSet.add(parseInt(sno, 10));
        const groups = stageGroups[sno];
        if (groups) sgpRows.push(...groups.map(g => ({ ...g, stageNo: sno })));
        const flags = worldFlags[sno];
        if (flags) wfRows.push(...flags.map(f => ({ ...f, stageNo: sno, _src: 'qst' })));
        const extra = worldFlagsExtra[sno];
        if (extra) wfRows.push(...extra.map(f => ({ ...f, stageNo: sno, _src: 'named' })));
    }
    // Sort combined layout rows: by questNo then flagNo
    wfRows.sort((a, b) => (a.questNo - b.questNo) || (a.flagNo - b.flagNo));
    // Quest flags: include any whose stageNo matches one of the shown stages
    const wfQuestRows = worldQuestFlags.filter(f => f.stageNo != null && stageNoSet.has(f.stageNo));

    // Hide panel when not in edit mode or no data in any section
    if (!_editMode || (sgpRows.length === 0 && wfRows.length === 0 && wfQuestRows.length === 0)) {
        panel.style.display = 'none';
        return;
    }

    panel.style.display = 'block';

    // ── Stage Groups section ──────────────────────────────────────────────────
    const sgpBodyHtml = sgpRows.map(r => {
        const name  = r.questName ? `<span class="sgp-name" title="${r.questName.replace(/"/g,'&quot;')}">${r.questName}</span>` : `<span style="color:#445">—</span>`;
        const quest = r.questNo    ? `<span class="sgp-quest">${r.questNo}</span>`    : `<span style="color:#445">—</span>`;
        const flag  = r.layoutFlagNo ? `<span class="sgp-flag">${r.layoutFlagNo}</span>` : `<span style="color:#445">—</span>`;
        const npcLine = (r.npcs?.length)
            ? `<div class="sgp-npc-list">${r.npcs.map(n =>
                `<span class="sgp-npc">${n.npcName || ''}${n.npcName ? ' ' : ''}<span class="sgp-npc-id">#${n.npcId}</span></span>`
              ).join('')}</div>` : '';
        return `<div class="sgp-row${npcLine ? ' sgp-row-has-npc' : ''}">
            <span class="sgp-group">Grp ${r.group}${r.gplType && r.gplType !== 'p' ? ` <span style="color:#556;font-size:0.65rem">${r.gplType}</span>` : ''}</span>${name}${quest}${flag}
            ${npcLine}
        </div>`;
    }).join('');

    // ── WM Layout section (qst + named, merged) ──────────────────────────────
    const seasonColors = {
        'S1': '#7a5c9a', 'S2': '#5c7a9a', 'S3': '#5c9a7a',
    };
    const wfBodyHtml = wfRows.map(r => {
        const seasonKey = r.season?.slice(0, 2) ?? '';
        const chipColor = seasonColors[seasonKey] ?? '#556';
        const chip = r.season
            ? `<span class="wf-season-chip" style="background:${chipColor}">${r.season}</span>`
            : '';

        // Col 3: Clear flag (qst entries only)
        let clearCol;
        if (r._src === 'qst') {
            const permanent = r.eraseFlagNo === 0
                ? `<span class="wf-permanent" title="Cannot be cleared">🔒</span>` : '';
            clearCol = `<span class="wf-erase">${r.eraseFlagNo || '—'}${permanent}</span>`;
        } else {
            clearCol = `<span class="wf-erase" style="color:#445">—</span>`;
        }

        // Col 4: Name (named entries only)
        const nameCol = r._src === 'named' && r.name
            ? `<span class="wf-name" title="${(r.className ? r.className + '.' : '') + r.name}">${r.name}</span>`
            : `<span class="wf-name" style="color:#445">—</span>`;

        // Col 6: Description
        let desc;
        if (r._src === 'named') {
            desc = r.description
                ? `<span class="wf-comment" title="${r.description.replace(/"/g,'&quot;')}">${r.description}</span>`
                : `<span style="color:#445">—</span>`;
        } else {
            const depStyle = r.deprecated ? 'text-decoration:line-through;opacity:0.5;' : '';
            const displayText = r.commentEn || r.comment || '';
            const tooltipText = (r.commentEn && r.comment)
                ? `${r.commentEn}\n\u30FB JP: ${r.comment}` : (r.comment || r.commentEn || '');
            desc = displayText
                ? `<span class="wf-comment" style="${depStyle}" title="${tooltipText.replace(/"/g,'&quot;')}">${displayText}</span>`
                : `<span style="color:#445">—</span>`;
        }

        return `<div class="wf-row${r.deprecated ? ' wf-deprecated' : ''}">
            ${chip}
            <span class="wf-flag">${r.flagNo}</span>
            ${clearCol}
            ${nameCol}
            <span class="wf-questno">${r.questNo}</span>
            ${desc}
        </div>`;
    }).join('');

    // ── WM Quest flags — worldQuestFlags ─────────────────────────────────────
    const wfQuestBodyHtml = wfQuestRows.map(r => {
        const seasonKey = r.season?.slice(0, 2) ?? '';
        const chipColor = seasonColors[seasonKey] ?? '#556';
        const chip = r.season
            ? `<span class="wf-season-chip" style="background:${chipColor}">${r.season}</span>` : '';
        const desc = r.description
            ? `<span class="wf-comment" title="${r.description.replace(/"/g,'&quot;')}">${r.description}</span>`
            : `<span style="color:#445">—</span>`;
        const nameStr = r.name
            ? `<span class="wf-name" title="${(r.className ? r.className + '.' : '') + r.name}">${r.name}</span>`
            : `<span class="wf-name" style="color:#445">—</span>`;
        return `<div class="wf-row">
            ${chip}
            <span class="wf-flag">${r.flagNo}</span>
            <span class="wf-erase" style="color:#445">—</span>
            ${nameStr}
            <span class="wf-questno">${r.questNo}</span>
            ${desc}
        </div>`;
    }).join('');

    panel.innerHTML =
        // ── Stage groups sub-panel ──
        (sgpRows.length ? `
        <div class="sgp-header" id="sgp-header-btn">
            <span class="sgp-title">Stage Groups (${sgpRows.length})</span>
            <span class="sgp-toggle">${_sgpCollapsed ? '▼ show' : '▲ hide'}</span>
        </div>
        <div class="sgp-col-head">
            <span>Group</span><span>Quest Name</span><span>QuestNo</span><span>LayoutFlag</span>
        </div>
        <div class="sgp-body">${sgpBodyHtml}</div>` : '') +
        // ── WM Layout flags (qst + named, merged) ──
        (wfRows.length ? `
        <div class="sgp-header wf-header" id="wf-header-btn">
            <span class="sgp-title">WM Layout Flags (${wfRows.length})</span>
            <span class="sgp-toggle">${_wfCollapsed ? '▼ show' : '▲ hide'}</span>
        </div>
        <div class="wf-col-head">
            <span>Season</span><span>Set Flag</span><span>Clear Flag</span><span>Name</span><span>QuestNo</span><span>Description</span>
        </div>
        <div class="wf-body">${wfBodyHtml}</div>` : '') +
        // ── WM Quest Flags ──
        (wfQuestRows.length ? `
        <div class="sgp-header wf-header" id="wf-quest-header-btn">
            <span class="sgp-title">WM Quest Flags (${wfQuestRows.length})</span>
            <span class="sgp-toggle">${_wfQuestCollapsed ? '▼ show' : '▲ hide'}</span>
        </div>
        <div class="wf-quest-col-head wf-col-head">
            <span>Season</span><span>FlagNo</span><span>—</span><span>Name</span><span>QuestNo</span><span>Description</span>
        </div>
        <div class="wf-quest-body wf-body">${wfQuestBodyHtml}</div>` : '');

    // ── Apply collapse state and wire click handlers ───────────────────────
    function wireSection(headerId, colHeadSel, bodySel, collapsed, setter) {
        const colHead = panel.querySelector(colHeadSel);
        const body    = panel.querySelector(bodySel);
        if (collapsed) {
            colHead?.classList.add('wf-hidden');
            body?.classList.add('wf-hidden');
        }
        document.getElementById(headerId)?.addEventListener('click', () => {
            const next = !collapsed;
            setter(next);
            collapsed = next;
            panel.querySelector(colHeadSel)?.classList.toggle('wf-hidden', collapsed);
            panel.querySelector(bodySel)?.classList.toggle('wf-hidden', collapsed);
            document.getElementById(headerId).querySelector('.sgp-toggle').textContent =
                collapsed ? '▼ show' : '▲ hide';
        });
    }

    if (sgpRows.length)     wireSection('sgp-header-btn',      '.sgp-col-head',       '.sgp-body',       _sgpCollapsed,     v => { _sgpCollapsed     = v; });
    if (wfRows.length)      wireSection('wf-header-btn',       '.wf-col-head',        '.wf-body',        _wfCollapsed,      v => { _wfCollapsed      = v; });
    if (wfQuestRows.length) wireSection('wf-quest-header-btn', '.wf-quest-col-head',  '.wf-quest-body',  _wfQuestCollapsed, v => { _wfQuestCollapsed = v; });
}

function buildFloorSelector(info) {
    const el = document.getElementById('floor-selector');
    el.innerHTML = '';
    const layers = (info.layers || []).filter(l => l.img_exists);
    if (layers.length <= 1) return;

    for (const { layer, img_file } of layers) {
        const btn = document.createElement('button');
        btn.textContent = `Floor ${layer}`;
        if (layer === currentLayer) btn.classList.add('active');
        btn.addEventListener('click', () => _switchToFloor(layer, info));
        el.appendChild(btn);
    }
}

function _switchToFloor(layer, info = _currentMapInfo) {
    if (!info || layer === currentLayer) return;
    const layers = (info.layers || []).filter(l => l.img_exists);
    const target = layers.find(l => l.layer === layer);
    if (!target) return;
    currentLayer = layer;
    swapMapImage(info, target.img_file);
    const el = document.getElementById('floor-selector');
    el.querySelectorAll('button').forEach(b => b.classList.toggle('active', b.textContent === `Floor ${layer}`));
    if (info.floor_obbs) {
        loadConnections(currentMapName(), info);
        loadEnemySpawns(info, currentStageName());
        loadGatherPoints(info, currentStageName());
        loadNpcShops(info, currentStageName());
        loadSpecialShops(info, currentStageName());
        if (typeof loadBreakTargets === "function") loadBreakTargets(info, currentStageName());
    }
    updateLayersInHash();
}

function swapMapImage(info, imgFile) {
    if (imageOverlay && _activeMapImageFile === imgFile) {
        fitMapToImage(info);
        return;
    }
    if (imageOverlay) imageOverlay.remove();
    const bounds = [xy(0, 0), xy(info.img_width, info.img_height)];
    imageOverlay = L.imageOverlay('images/maps/' + imgFile, bounds, { pane: 'mapImagePane' }).addTo(leafletMap);
    _activeMapImageFile = imgFile;
    fitMapToImage(info);
}

// ── Tile-layer selector (pd maps with multi-layer pieces) ─────────────────────
let _tileLayerSel = null;   // id of currently shown tile-layer image, null = merged default

function buildTileLayerSelector(info, { activeKey = null } = {}) {
    const el = document.getElementById('tile-layer-selector');
    if (!el) return;
    el.innerHTML = '';
    const tlImages = info.tile_layer_images;
    if (!tlImages || Object.keys(tlImages).length === 0) {
        _tileLayerSel = null;
        return;
    }

    const initialKey = activeKey !== undefined ? activeKey : _tileLayerSel;
    _tileLayerSel = initialKey;

    const makeBtn = (label, key, imgFile) => {
        const btn = document.createElement('button');
        btn.textContent = label;
        if (key === initialKey) btn.classList.add('active');
        btn.addEventListener('click', () => {
            _tileLayerSel = key;
            swapMapImage(info, imgFile);
            el.querySelectorAll('button').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
        });
        el.appendChild(btn);
    };

    makeBtn('Merged', null, info.img_file);
    for (const [tl, imgFile] of Object.entries(tlImages).sort((a, b) => a[0] - b[0])) {
        makeBtn(`Tile L${tl}`, Number(tl), imgFile);
    }
}

// ── Piece image lightbox ──────────────────────────────────────────────────────
const _lb = (() => {
    const overlay = document.createElement('div');
    overlay.style.cssText = [
        'position:fixed;inset:0;z-index:9999',
        'background:rgba(0,0,0,0.82)',
        'display:none;align-items:center;justify-content:center;flex-direction:column',
        'gap:10px;cursor:zoom-out',
    ].join(';');

    const title = document.createElement('div');
    title.style.cssText = 'color:#ffcc44;font-family:monospace;font-size:0.85rem;font-weight:700;user-select:none';

    const DISPLAY_CSS = 'max-width:90vw;max-height:78vh;object-fit:contain;image-rendering:pixelated;border:1px solid #0f3460';
    const img = document.createElement('img');
    img.style.cssText = DISPLAY_CSS;

    const canvas = document.createElement('canvas');
    canvas.style.cssText = DISPLAY_CSS + ';display:none';

    const nav = document.createElement('div');
    nav.style.cssText = 'display:flex;gap:6px;flex-wrap:wrap;justify-content:center;cursor:default';

    overlay.append(title, img, canvas, nav);
    document.body.appendChild(overlay);
    overlay.addEventListener('click', e => {
        if (e.target === overlay) overlay.style.display = 'none';
    });
    document.addEventListener('keydown', e => { if (e.key === 'Escape') overlay.style.display = 'none'; });

    let _activeLayer = null;  // currently shown layer index (null = primary/merged)
    let _compositeGen = 0;    // incremented on every showLayer call; stale async results are discarded

    const layerNav = document.createElement('div');
    layerNav.style.cssText = 'display:none;gap:4px;align-items:center;cursor:default';
    overlay.insertBefore(layerNav, nav);

    function showLayer(piece, pieceIdx, layer) {
        // layer === null means "composite / merged view"
        _activeLayer = layer;
        const gen = ++_compositeGen;
        layerNav.querySelectorAll('button').forEach(btn => {
            const active = String(btn.dataset.layer) === String(layer);
            btn.style.background  = active ? '#4a90d9' : '#0f3460';
            btn.style.borderColor = active ? '#4a90d9' : '#1a4a7a';
            btn.style.color       = active ? '#fff'    : '#ccd';
        });

        if (layer !== null) {
            // Single layer — just show the image
            canvas.style.display = 'none';
            img.style.display    = '';
            img.src = `images/maps/${piece.model}_l${layer}.png`;
            title.textContent = `${pieceIdx}: ${piece.model}_l${layer}.png`;
            return;
        }

        // Composite view
        if (piece.has_merged) {
            // Pre-generated merged PNG exists — use it directly
            canvas.style.display = 'none';
            img.style.display    = '';
            img.src = `images/maps/${piece.model}_merged.png`;
            title.textContent = `${pieceIdx}: ${piece.model}_merged.png`;
            return;
        }

        // Composite on-the-fly: draw each layer in order onto a canvas
        const allLayers = piece.layers ?? [piece.layer ?? 0];
        title.textContent = `${pieceIdx}: ${piece.model} (composite)`;
        Promise.all(allLayers.map(lyr => new Promise((res, rej) => {
            const i = new Image();
            i.onload = () => res(i);
            i.onerror = () => rej(new Error(`missing l${lyr}`));
            i.src = `images/maps/${piece.model}_l${lyr}.png`;
        }))).then(imgs => {
            if (_compositeGen !== gen) return; // user switched layer before images loaded
            canvas.width  = imgs[0].naturalWidth;
            canvas.height = imgs[0].naturalHeight;
            const ctx = canvas.getContext('2d');
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            for (const i of imgs) ctx.drawImage(i, 0, 0);
            img.style.display    = 'none';
            canvas.style.display = '';
        }).catch(() => {
            if (_compositeGen !== gen) return;
            // Fallback: show primary layer if any image fails to load
            canvas.style.display = 'none';
            img.style.display    = '';
            img.src = `images/maps/${piece.model}_l${piece.layer ?? 0}.png`;
        });
    }

    function showPiece(piece, idx) {
        // Build layer sub-buttons when the piece has multiple layers
        const allLayers = piece.layers ?? [piece.layer ?? 0];
        if (allLayers.length > 1) {
            layerNav.innerHTML = '';
            layerNav.style.display = 'flex';
            const label = document.createElement('span');
            label.textContent = 'Layers:';
            label.style.cssText = 'color:#888;font-family:monospace;font-size:0.75rem';
            layerNav.appendChild(label);
            // "Composite" button = primary layer image (what's used in the map)
            const btnC = document.createElement('button');
            btnC.textContent = 'composite';
            btnC.dataset.layer = 'null';
            btnC.style.cssText = 'background:#0f3460;color:#ccd;border:1px solid #1a4a7a;border-radius:4px;padding:2px 7px;font-size:0.72rem;font-family:monospace;cursor:pointer';
            btnC.addEventListener('click', e => { e.stopPropagation(); showLayer(piece, idx, null); });
            layerNav.appendChild(btnC);
            // One button per individual layer
            for (const lyr of allLayers) {
                const btn = document.createElement('button');
                btn.textContent = `l${lyr}`;
                btn.dataset.layer = lyr;
                btn.style.cssText = 'background:#0f3460;color:#ccd;border:1px solid #1a4a7a;border-radius:4px;padding:2px 7px;font-size:0.72rem;font-family:monospace;cursor:pointer';
                btn.addEventListener('click', e => { e.stopPropagation(); showLayer(piece, idx, lyr); });
                layerNav.appendChild(btn);
            }
        } else {
            layerNav.style.display = 'none';
        }
        // Update active state on piece nav chips
        nav.querySelectorAll('button').forEach((btn, i) => {
            btn.style.background  = i === idx ? '#4a90d9' : '#0f3460';
            btn.style.borderColor = i === idx ? '#4a90d9' : '#1a4a7a';
            btn.style.color       = i === idx ? '#fff'    : '#ccd';
        });
        showLayer(piece, idx, null);  // default: show composite (primary layer)
    }

    return {
        show(pieces, idx) {
            // Build piece nav buttons
            nav.innerHTML = '';
            pieces.forEach((p, i) => {
                const btn = document.createElement('button');
                btn.textContent = `${i}: ${p.model.replace(/^pd\d+_/, '')}`;
                btn.style.cssText = [
                    'background:#0f3460;color:#ccd;border:1px solid #1a4a7a',
                    'border-radius:4px;padding:3px 8px;font-size:0.75rem',
                    'font-family:monospace;cursor:pointer',
                ].join(';');
                btn.addEventListener('click', e => { e.stopPropagation(); showPiece(p, i); });
                nav.appendChild(btn);
            });
            showPiece(pieces[idx], idx);
            overlay.style.display = 'flex';
        },
    };
})();

// ── Pd piece boundaries ───────────────────────────────────────────────────────
function loadPdBoundaries(info) {
    pdBoundaryLayer.clearLayers();
    if (!info.pd_pieces) return;
    const pieces = info.pd_pieces;
    pieces.forEach((piece, i) => {
        // Seam: entrance of this piece at the bottom of its slot
        const png_y = piece.pixel_y_entrance;
        const py = info.img_height - png_y;   // lat (Leaflet y from bottom)
        const label = piece.model.replace(/^pd\d+_/, '');
        const line = L.polyline(
            [xy(0, py), xy(info.img_width, py)],
            { color: '#ffcc44', weight: 1, opacity: 0.7, dashArray: '4 4', interactive: false }
        );
        const marker = L.marker(xy(info.img_width, py), {
            icon: L.divIcon({
                className: 'pd-label pd-label-btn',
                html: `${i}: ${label}`,
                iconSize: [80, 14],
                iconAnchor: [-4, 7],
            }),
            interactive: true,
        });
        marker.on('click', () => _lb.show(pieces, i));
        pdBoundaryLayer.addLayer(line);
        pdBoundaryLayer.addLayer(marker);
    });
}

// ── Map loader ────────────────────────────────────────────────────────────────
const nextFrame = () => new Promise(resolve => {
    requestAnimationFrame(() => requestAnimationFrame(resolve));
});

const waitForMapContainer = (timeoutMs = 8000) => new Promise(resolve => {
    const mapEl = document.getElementById('map');
    const t0 = performance.now();
    const check = () => {
        if (mapEl && mapEl.offsetWidth > 0 && mapEl.offsetHeight > 0) return resolve(mapEl);
        if (performance.now() - t0 >= timeoutMs) return resolve(mapEl);
        requestAnimationFrame(check);
    };
    check();
});

const waitForImageOverlay = (overlay, timeoutMs = 15000) => new Promise(resolve => {
    if (!overlay) return resolve();
    const img = overlay.getElement?.();
    if (img?.complete && img.naturalWidth > 0) return resolve();
    let settled = false;
    const done = () => { if (!settled) { settled = true; resolve(); } };
    overlay.once('load', done);
    overlay.once('error', done);
    setTimeout(done, timeoutMs);
});

function fitMapToImage(info) {
    if (!info?.img_exists) {
        leafletMap.invalidateSize({ animate: false });
        return;
    }
    leafletMap.invalidateSize({ animate: false });
    const bounds = [xy(0, 0), xy(info.img_width, info.img_height)];
    const savedView = parseHash().view;
    if (savedView?.center
        && Number.isFinite(savedView.center.lat)
        && Number.isFinite(savedView.center.lng)
        && Number.isFinite(savedView.zoom)) {
        leafletMap.setView(savedView.center, savedView.zoom, { animate: false });
        return;
    }
    const mapEl = document.getElementById('map');
    if (mapEl?.offsetWidth > 0 && mapEl?.offsetHeight > 0) {
        leafletMap.fitBounds(bounds, { animate: false, padding: [0, 0] });
    } else {
        leafletMap.setView(xy(info.img_width / 2, info.img_height / 2), 0, { animate: false });
    }
}

function refitActiveMapView() {
    fitMapToImage(mapParams[_loadedMapName]);
}

function setMapLoading(on) {
    document.getElementById('map-container')?.classList.toggle('map-loading', !!on);
}

function resetView() {
    const info = mapParams[_loadedMapName];
    if (!info) return;
    fitMapToImage(info);
}

function watchMapContainerSize() {
    const mapEl = document.getElementById('map');
    if (!mapEl || typeof ResizeObserver === 'undefined') return;
    let lastW = 0;
    let lastH = 0;
    new ResizeObserver(() => {
        const w = mapEl.offsetWidth;
        const h = mapEl.offsetHeight;
        if (w > 0 && h > 0 && (lastW === 0 || lastH === 0)) fitMapToImage(mapParams[_loadedMapName]);
        else if (w > 0 && h > 0 && (Math.abs(w - lastW) > 24 || Math.abs(h - lastH) > 24)) {
            leafletMap.invalidateSize({ animate: false });
        }
        lastW = w;
        lastH = h;
    }).observe(mapEl);
}

async function loadMapImagePhase(info, imgFile = null) {
    const file = imgFile ?? resolveDisplayedMapImageFile(info, currentLayer, _tileLayerSel);
    if (!info.img_exists) {
        if (imageOverlay) imageOverlay.remove();
        imageOverlay = null;
        _activeMapImageFile = null;
        leafletMap.setView(xy(info.img_width / 2, info.img_height / 2), 0);
        leafletMap.invalidateSize({ animate: false });
        return;
    }
    if (imageOverlay && _activeMapImageFile === file) {
        await waitForMapContainer();
        leafletMap.invalidateSize({ animate: false });
        fitMapToImage(info);
        return;
    }
    if (imageOverlay) imageOverlay.remove();
    await waitForMapContainer();
    leafletMap.invalidateSize({ animate: false });
    const bounds = [xy(0, 0), xy(info.img_width, info.img_height)];
    imageOverlay = L.imageOverlay('images/maps/' + file, bounds, { pane: 'mapImagePane' }).addTo(leafletMap);
    _activeMapImageFile = file;
    fitMapToImage(info);
    await waitForImageOverlay(imageOverlay);
    await nextFrame();
    fitMapToImage(info);
}

async function loadMapOverlaysPhase(mapName, info, stid, openGroups) {
    resetConnectionPoiCache(mapName);
    loadGrid(info);
    loadPdBoundaries(info);
    loadStageLabels(info);
    loadLandmarks(mapName, info);
    loadConnections(mapName, info);
    loadGatherPoints(info, stid);
    loadNpcShops(info, stid);
    loadSpecialShops(info, stid);
    await nextFrame();

    await loadEnemySpawns(info, stid);

    if (openGroups?.length) {
        for (const id of openGroups) if (_groupStore.has(id)) _expandGroupCore(_groupStore.get(id));
        _updateExpandCollapseBtn();
        reapplySpread();
    }

    _spotOpenedGroup = null;
    await buildSpotIndex(info);
    if (document.getElementById('spot-panel')?.classList.contains('open')) _runSpotSearch();

    if (_pendingGlobalNavTarget) {
        const t = _pendingGlobalNavTarget;
        _pendingGlobalNavTarget = null;
        setTimeout(() => _navigateToSpot(t), 300);
    }

    if (_pendingNamedLocNav) {
        const { latlng } = _pendingNamedLocNav;
        _pendingNamedLocNav = null;
        setTimeout(() => {
            _clearSpotHighlights();
            _addSpotHighlight(latlng);
        }, 400);
    }

    fitMapToImage(info);
}

// ── Map scene cache (fast browser back / revisit) ─────────────────────────────
const MAP_SCENE_CACHE_MAX = 10;
const _mapSceneCache = new Map();
let _activeMapImageFile = null;

function mapSceneCacheEnabled() {
    return !_editMode && _dirtySet.size === 0;
}

function normalizeSceneStid(stid) {
    return stid ?? null;
}

function mapSceneCacheKey(mapName, stid, floor, tileSel) {
    return `${mapName}|${normalizeSceneStid(stid)}|${floor}|${tileSel ?? 'merged'}`;
}

function resolveDisplayedMapImageFile(info, floor = currentLayer, tileSel = _tileLayerSel) {
    if (tileSel != null && info.tile_layer_images?.[tileSel]) {
        return info.tile_layer_images[tileSel];
    }
    if (floor !== 0 && info.layers) {
        const floorLayer = info.layers.find(l => l.layer === floor && l.img_exists);
        if (floorLayer) return floorLayer.img_file;
    }
    return info.img_file;
}

function resolveInitialFloor(mapName, stid, info) {
    if (!info.floor_obbs) return 0;
    const savedFloor = parseHash().view?.floor;
    if (savedFloor != null) return savedFloor;
    if (stid) {
        const arrivalStageNo = parseInt(stid.slice(2), 10);
        const conns = connectionData[mapName] ?? [];
        const arrConn = conns.find(c => c.from_stage === arrivalStageNo && c.x != null && c.z != null);
        if (arrConn) {
            const arrFloor = getEnemyFloor(arrConn.x, arrConn.y ?? 0, arrConn.z, info.floor_obbs);
            if (arrFloor !== null) return arrFloor;
        }
    }
    return 0;
}

function mapSceneIdentity(mapName, stid, floor, tileSel = _tileLayerSel) {
    return mapSceneCacheKey(mapName, stid, floor, tileSel);
}

function buildMapDocumentTitle(mapName, info, stid) {
    const baseName = stid ? stageLabel(info, stid) : (info.name_en ? splitPascalCase(info.name_en) : mapName);
    const sidNum   = stid ? info.stage_ids?.[stid] : null;
    const sidStr   = sidNum != null ? ` - sid${String(sidNum).padStart(4, '0')}` : '';
    return baseName + ` (${stid ?? mapName}${sidStr})`;
}

function detachSceneFromMap(scene) {
    const layers = [
        scene.enemyLayer, scene.landmarkLayer, scene.connectionLayer, scene.gridLayer,
        scene.stageLabelsLayer, scene.gatherLayer, scene.npcShopLayer, scene.specialShopLayer,
        scene.pdBoundaryLayer, scene.spawnRadiiLayer, scene._spreadOverlay,
    ];
    for (const lg of layers) {
        if (lg && leafletMap.hasLayer(lg)) leafletMap.removeLayer(lg);
    }
    if (scene.territoryLayer && leafletMap.hasLayer(scene.territoryLayer)) {
        leafletMap.removeLayer(scene.territoryLayer);
    }
    for (const g of scene.groupStore.values()) {
        if (g.detailsLayer && leafletMap.hasLayer(g.detailsLayer)) {
            leafletMap.removeLayer(g.detailsLayer);
        }
    }
    if (scene.imageOverlay && leafletMap.hasLayer(scene.imageOverlay)) {
        leafletMap.removeLayer(scene.imageOverlay);
    }
}

function attachSceneToMap(scene) {
    if (scene.imageOverlay) scene.imageOverlay.addTo(leafletMap);
    const layers = [
        scene.enemyLayer, scene.landmarkLayer, scene.connectionLayer,
        scene.stageLabelsLayer, scene.gatherLayer, scene.npcShopLayer, scene.specialShopLayer,
        scene.pdBoundaryLayer, scene.spawnRadiiLayer, scene._spreadOverlay,
    ];
    for (const lg of layers) if (lg) lg.addTo(leafletMap);
    if (scene.gridLayer && document.getElementById('layer-grid')?.checked) {
        scene.gridLayer.addTo(leafletMap);
    }
    updateEnemyVisibility();
    syncTerritoryLayer();
    applyAllPoiVisibility();
}

function createFreshSceneLayers() {
    // Cancel any in-flight chunked group expansion from the previous map before
    // awaiting image loads — otherwise stale _expandGroupCore calls re-add
    // orphaned spawn dots that are outside _groupStore (mob toggles won't hide them).
    ++_mobDisplayModeJob;

    const oldLayers = [
        enemyLayer, landmarkLayer, connectionLayer, gridLayer, territoryLayer,
        stageLabelsLayer, gatherLayer, npcShopLayer, specialShopLayer,
        pdBoundaryLayer, spawnRadiiLayer, _spreadOverlay,
    ];
    for (const lg of oldLayers) {
        if (lg && leafletMap.hasLayer(lg)) leafletMap.removeLayer(lg);
    }
    for (const g of _groupStore.values()) {
        if (g.detailsLayer && leafletMap.hasLayer(g.detailsLayer)) {
            leafletMap.removeLayer(g.detailsLayer);
        }
    }

    enemyLayer = L.layerGroup().addTo(leafletMap);
    landmarkLayer = L.layerGroup().addTo(leafletMap);
    connectionLayer = L.layerGroup().addTo(leafletMap);
    gridLayer = L.layerGroup();
    territoryLayer = L.layerGroup();
    stageLabelsLayer = L.layerGroup().addTo(leafletMap);
    gatherLayer = L.layerGroup().addTo(leafletMap);
    npcShopLayer = L.layerGroup().addTo(leafletMap);
    specialShopLayer = L.layerGroup().addTo(leafletMap);
    pdBoundaryLayer = L.layerGroup().addTo(leafletMap);
    spawnRadiiLayer = L.layerGroup().addTo(leafletMap);
    _spreadOverlay = L.layerGroup().addTo(leafletMap);

    // New Maps — do not clear() the previous ones; scene cache still holds them.
    _groupStore = new Map();
    _gatherMarkerByKey = new Map();
    _gatherGroupMarkers = new Map();
    _shopMarkerByNpcId = new Map();
    _specialShopMarkerByNpcId = new Map();
    _connectionPoiCacheMap = null;
    _connectionPoiCache = null;
    _spotIndex = [];
    _activeSubGroupId = null;
    _availableSubGroups = [0];
    _spotOpenedGroup = null;
    _sgMarkers = {};
    clearSpawnRadii();
    _activeRadiiMarker = null;
}

function captureCurrentScene() {
    const cacheKey = mapSceneCacheKey(_loadedMapName, _loadedStid, currentLayer, _tileLayerSel);
    return {
        cacheKey,
        mapName: _loadedMapName,
        stid: normalizeSceneStid(_loadedStid),
        floor: currentLayer,
        tileLayerSel: _tileLayerSel,
        imgFile: _activeMapImageFile,
        enemyLayer, landmarkLayer, connectionLayer, gridLayer, territoryLayer,
        stageLabelsLayer, gatherLayer, npcShopLayer, specialShopLayer,
        pdBoundaryLayer, spawnRadiiLayer, _spreadOverlay,
        imageOverlay,
        groupStore: _groupStore,
        gatherMarkerByKey: _gatherMarkerByKey,
        gatherGroupMarkers: _gatherGroupMarkers,
        shopMarkerByNpcId: _shopMarkerByNpcId,
        specialShopMarkerByNpcId: _specialShopMarkerByNpcId,
        connectionPoiCacheMap: _connectionPoiCacheMap,
        connectionPoiCache: _connectionPoiCache,
        spotIndex: _spotIndex,
        currentMapInfo: _currentMapInfo,
        currentFloorObbs: _currentFloorObbs,
        activeSubGroupId: _activeSubGroupId,
        availableSubGroups: _availableSubGroups,
        spotOpenedGroup: _spotOpenedGroup,
        sgMarkers: _sgMarkers,
    };
}

function applySceneToGlobals(scene) {
    enemyLayer = scene.enemyLayer;
    landmarkLayer = scene.landmarkLayer;
    connectionLayer = scene.connectionLayer;
    gridLayer = scene.gridLayer;
    territoryLayer = scene.territoryLayer;
    stageLabelsLayer = scene.stageLabelsLayer;
    gatherLayer = scene.gatherLayer;
    npcShopLayer = scene.npcShopLayer;
    specialShopLayer = scene.specialShopLayer;
    pdBoundaryLayer = scene.pdBoundaryLayer;
    spawnRadiiLayer = scene.spawnRadiiLayer;
    _spreadOverlay = scene._spreadOverlay;
    imageOverlay = scene.imageOverlay;
    _groupStore = scene.groupStore;
    _gatherMarkerByKey = scene.gatherMarkerByKey;
    _gatherGroupMarkers = scene.gatherGroupMarkers;
    _shopMarkerByNpcId = scene.shopMarkerByNpcId;
    _specialShopMarkerByNpcId = scene.specialShopMarkerByNpcId;
    _connectionPoiCacheMap = scene.connectionPoiCacheMap;
    _connectionPoiCache = scene.connectionPoiCache;
    _spotIndex = scene.spotIndex;
    _currentMapInfo = scene.currentMapInfo;
    _currentFloorObbs = scene.currentFloorObbs;
    _activeSubGroupId = scene.activeSubGroupId;
    _availableSubGroups = scene.availableSubGroups;
    _spotOpenedGroup = scene.spotOpenedGroup;
    _sgMarkers = scene.sgMarkers;
    currentLayer = scene.floor;
    _tileLayerSel = scene.tileLayerSel;
    _activeMapImageFile = scene.imgFile;
}

function destroyCachedScene(scene) {
    if (!scene) return;
    detachSceneFromMap(scene);
    for (const g of scene.groupStore.values()) {
        if (g.detailsLayer) {
            g.detailsLayer.clearLayers();
            g.detailsLayer = null;
        }
    }
    scene.enemyLayer?.clearLayers();
    scene.landmarkLayer?.clearLayers();
    scene.connectionLayer?.clearLayers();
    scene.gatherLayer?.clearLayers();
    scene.npcShopLayer?.clearLayers();
    scene.specialShopLayer?.clearLayers();
}

function pruneMapSceneCache() {
    while (_mapSceneCache.size > MAP_SCENE_CACHE_MAX) {
        const oldestKey = _mapSceneCache.keys().next().value;
        destroyCachedScene(_mapSceneCache.get(oldestKey));
        _mapSceneCache.delete(oldestKey);
    }
}

function putSceneInCache(scene) {
    if (!scene) return;
    _mapSceneCache.delete(scene.cacheKey);
    _mapSceneCache.set(scene.cacheKey, scene);
    pruneMapSceneCache();
}

function stashCurrentSceneInCache() {
    if (!_loadedMapName || !_currentMapInfo || !mapSceneCacheEnabled()) return;
    ++_mobDisplayModeJob; // stop chunked expand before detach/await gaps
    const scene = captureCurrentScene();
    detachSceneFromMap(scene);
    putSceneInCache(scene);
    imageOverlay = null;
    _activeMapImageFile = null;
    _currentMapInfo = null;
}

function takeCachedScene(mapName, stid, floor) {
    const wantStid = normalizeSceneStid(stid);
    for (const [key, scene] of _mapSceneCache) {
        if (scene.mapName === mapName && scene.stid === wantStid && scene.floor === floor) {
            _mapSceneCache.delete(key);
            return scene;
        }
    }
    return null;
}

function invalidateMapSceneCache() {
    for (const scene of _mapSceneCache.values()) destroyCachedScene(scene);
    _mapSceneCache.clear();
}

async function restoreCachedMapScene(scene, mapName, info, stid, isCurrent) {
    applySceneToGlobals(scene);
    attachSceneToMap(scene);
    if (!isCurrent()) return false;

    buildFloorSelector(info);
    buildTileLayerSelector(info, { activeKey: scene.tileLayerSel });
    buildStageGroupsPanel(info, stid);

    await nextFrame();
    fitMapToImage(info);
    reapplySpread();

    if (document.getElementById('spot-panel')?.classList.contains('open')) _runSpotSearch();
    return true;
}

let _loadMapGen = 0;

async function loadMap(mapName) {
    const info = mapParams[mapName];
    if (!info) return;

    // Cancel pending mob expansion immediately — loadMap awaits image I/O and a
    // stale expandChunk can paint previous-map spawns onto the new view.
    ++_mobDisplayModeJob;

    const loadGen = ++_loadMapGen;
    const isCurrent = () => loadGen === _loadMapGen;
    let stid = currentStageName();
    if (info.stages?.length && (!stid || !info.stages.includes(stid))) {
        stid = defaultFieldStage(mapName);
        const parsed = parseHash();
        let hash = stid ? `${mapName}:${stid}` : mapName;
        if (parsed.view) {
            hash += `@${parsed.view.zoom.toFixed(2)}/${parsed.view.center.lat.toFixed(1)}/${parsed.view.center.lng.toFixed(1)}`;
            if (parsed.view.floor != null) hash += `/${parsed.view.floor}`;
        }
        const layersSuffix = window.location.hash.includes('!')
            ? window.location.hash.slice(window.location.hash.indexOf('!')) : '';
        history.replaceState(null, '', `#${hash}${layersSuffix}`);
        _loadedStid = stid;
    }
    const targetFloor = resolveInitialFloor(mapName, stid, info);

    const incomingIdentity = mapSceneIdentity(mapName, stid, targetFloor);
    const outgoingIdentity = _loadedMapName && _currentMapInfo
        ? mapSceneIdentity(_loadedMapName, _loadedStid, currentLayer) : null;

    if (outgoingIdentity && outgoingIdentity !== incomingIdentity && mapSceneCacheEnabled()) {
        stashCurrentSceneInCache();
    }

    _loadedMapName = mapName;
    _loadedStid = stid;

    const title = buildMapDocumentTitle(mapName, info, stid);
    document.getElementById('map-title').textContent = title;
    document.title = `${title} — DDON Maps`;

    if (mapSceneCacheEnabled()) {
        const cached = takeCachedScene(mapName, stid, targetFloor);
        if (cached) {
            setMapLoading(true);
            try {
                currentLayer = targetFloor;
                const restored = await restoreCachedMapScene(cached, mapName, info, stid, isCurrent);
                if (restored && isCurrent()) return;
                if (!isCurrent()) return;
                putSceneInCache(cached);
            } catch (err) {
                if (!isCurrent()) return;
                console.error('restore cached map failed:', err);
                putSceneInCache(cached);
            } finally {
                if (isCurrent()) setMapLoading(false);
            }
        }
    }

    createFreshSceneLayers();
    currentLayer = targetFloor;
    setMapLoading(true);

    try {
        const imgFile = resolveDisplayedMapImageFile(info, currentLayer, null);
        await loadMapImagePhase(info, imgFile);
        if (!isCurrent()) return;

        if (info.floor_obbs && currentLayer !== 0) {
            const floorLayer = (info.layers || []).find(l => l.layer === currentLayer && l.img_exists);
            if (floorLayer && _activeMapImageFile !== floorLayer.img_file) {
                swapMapImage(info, floorLayer.img_file);
                await waitForImageOverlay(imageOverlay);
                await nextFrame();
                fitMapToImage(info);
            }
        }

        buildFloorSelector(info);
        buildTileLayerSelector(info);
        buildStageGroupsPanel(info, stid);

        const { openGroups } = parseHash();
        await loadMapOverlaysPhase(mapName, info, stid, openGroups);
        if (!isCurrent()) return;
    } catch (err) {
        if (!isCurrent()) return;
        console.error('loadMap failed:', err);
        document.getElementById('map-title').textContent = `${title} — load error`;
    } finally {
        if (isCurrent()) setMapLoading(false);
    }
}

// ── Spot search panel ─────────────────────────────────────────────────────────

let _spotIndex        = [];    // searchable entries for the current map
let _spotHighlights   = [];    // active pulsing ring markers
let _spotOpenedGroup  = null;  // groupId of the enemy group last opened by spot search
const _spotHlLayer  = L.layerGroup().addTo(leafletMap);
let _spotGlobal       = false; // true = global (all stages) search mode
let _globalSpotIndex  = [];    // searchable entries across all maps/stages
let _globalSpotIndexReady = false;
let _globalSpotIndexPromise = null;
let _pendingGlobalNavTarget = null; // deferred navigation target after stage switch
const SPOT_RESULT_CAP = 80;     // max grouped rows rendered (keeps typing responsive)
const _spotOriginLayer = L.layerGroup().addTo(leafletMap);

function _clearSpotHighlights() {
    for (const m of _spotHighlights) _spotHlLayer.removeLayer(m);
    _spotHighlights = [];
    for (const el of _spotHlChipEls) el.classList.remove('spot-hl-chip');
    _spotHlChipEls = [];
}

let _spotHlChipEls = [];   // chip DOM elements that have the highlight class

function _addSpotHighlight(latlng) {
    // className='spot-hl-outer' keeps Leaflet's translate3d intact; inner div carries the scale animation
    const icon = L.divIcon({ className: 'spot-hl-outer', html: '<div class="spot-hl"></div>', iconSize: [22, 22], iconAnchor: [11, 11] });
    _spotHighlights.push(L.marker(latlng, { icon, interactive: false }).addTo(_spotHlLayer));
}

function _addChipHighlight(g) {
    const el = g.labelMarker.getElement()?.querySelector('.group-chip');
    if (!el) return;
    el.classList.add('spot-hl-chip');
    _spotHlChipEls.push(el);
}

// Resolve the best highlight/flyTo position for a spot index entry at the moment it's needed.
// For enemies: chip position when collapsed, live marker position when expanded.
// sgMarkers is only populated after the group's first expand (buildGroupDetails), so this
// must be called at interaction time, not at index-build time.
function _resolveSpotLatLng(item) {
    const g = item.groupId ? _groupStore.get(item.groupId) : null;
    if (!g) return item.latlng;
    if (!g.isExpanded) return g.labelMarker.getLatLng();
    // Group is expanded — prefer exact spawnKey match, fall back to first emCode match
    if (_enemySpawnCache) {
        for (const markers of Object.values(g.sgMarkers)) {
            for (const m of markers) {
                if (item.spawnKey && m._spawnKey === item.spawnKey) {
                    const entries = _enemySpawnCache.get(m._spawnKey) ?? [];
                    if (!item.emCode || entries.some(e => e.emCode === item.emCode)) return m.getLatLng();
                }
            }
        }
        if (item.emCode) {
            for (const markers of Object.values(g.sgMarkers)) {
                for (const m of markers) {
                    const entries = m._spawnKey ? (_enemySpawnCache.get(m._spawnKey) ?? []) : [];
                    if (entries.some(e => e.emCode === item.emCode)) return m.getLatLng();
                }
            }
        }
    }
    return item.latlng;
}

function _navigateToSpot(target) {
    // Switch floor if the target lives on a different floor
    if (_currentFloorObbs && target.worldPos) {
        const targetFloor = getEnemyFloor(target.worldPos.x, target.worldPos.y, target.worldPos.z, _currentFloorObbs);
        if (targetFloor !== null && targetFloor !== currentLayer) _switchToFloor(targetFloor);
    }
    if (target.type === 'enemy' && target.groupId) {
        if (_spotOpenedGroup && _spotOpenedGroup !== target.groupId) collapseGroup(_spotOpenedGroup);
        _spotOpenedGroup = target.groupId;
        expandGroup(target.groupId);
        const g = _groupStore.get(target.groupId);
        let targetMarker = null;
        if (g) {
            // Prefer exact spawnKey match, fall back to first marker with matching emCode
            outer: for (const markers of Object.values(g.sgMarkers)) {
                for (const m of markers) {
                    if (target.spawnKey && m._spawnKey === target.spawnKey) {
                        const entries = m._spawnKey ? (_enemySpawnCache?.get(m._spawnKey) ?? []) : [];
                        if (!target.emCode || entries.some(e => e.emCode === target.emCode)) {
                            targetMarker = m; break outer;
                        }
                    }
                }
            }
            if (!targetMarker && target.emCode && _enemySpawnCache) {
                outer2: for (const markers of Object.values(g.sgMarkers)) {
                    for (const m of markers) {
                        const entries = m._spawnKey ? (_enemySpawnCache.get(m._spawnKey) ?? []) : [];
                        if (entries.some(e => e.emCode === target.emCode)) { targetMarker = m; break outer2; }
                    }
                }
            }
            if (!targetMarker) {
                const firstMarkers = Object.values(g.sgMarkers)[0];
                targetMarker = firstMarkers?.[0] ?? null;
            }
        }
        // Fly to and highlight the specific marker position, not just the group centroid
        const focusLatLng = targetMarker?.getLatLng() ?? target.latlng;
        leafletMap.flyTo(focusLatLng, Math.max(leafletMap.getZoom(), 2), { duration: 0.4 });
        _clearSpotHighlights();
        _addSpotHighlight(focusLatLng);
        if (targetMarker) setTimeout(() => targetMarker.openPopup(), 450);
    } else if (target.type === 'item' && target.source === 'enemy' && target.groupId) {
        // Item — enemy drop: reuse full enemy navigation
        if (_spotOpenedGroup && _spotOpenedGroup !== target.groupId) collapseGroup(_spotOpenedGroup);
        _spotOpenedGroup = target.groupId;
        expandGroup(target.groupId);
        const g = _groupStore.get(target.groupId);
        let targetMarker = null;
        if (g) {
            outer: for (const markers of Object.values(g.sgMarkers)) {
                for (const m of markers) {
                    if (target.spawnKey && m._spawnKey === target.spawnKey) {
                        const entries = m._spawnKey ? (_enemySpawnCache?.get(m._spawnKey) ?? []) : [];
                        if (!target.emCode || entries.some(e => e.emCode === target.emCode)) {
                            targetMarker = m; break outer;
                        }
                    }
                }
            }
            if (!targetMarker && target.emCode && _enemySpawnCache) {
                outer2: for (const markers of Object.values(g.sgMarkers)) {
                    for (const m of markers) {
                        const entries = m._spawnKey ? (_enemySpawnCache.get(m._spawnKey) ?? []) : [];
                        if (entries.some(e => e.emCode === target.emCode)) { targetMarker = m; break outer2; }
                    }
                }
            }
        }
        const focusLatLng = targetMarker?.getLatLng() ?? target.latlng;
        leafletMap.flyTo(focusLatLng, Math.max(leafletMap.getZoom(), 2), { duration: 0.4 });
        _clearSpotHighlights();
        _addSpotHighlight(focusLatLng);
        if (targetMarker) setTimeout(() => targetMarker.openPopup(), 450);
    } else {
        // Gather, shop item, or generic fallback
        leafletMap.flyTo(target.latlng, Math.max(leafletMap.getZoom(), 2), { duration: 0.4 });
        _clearSpotHighlights();
        _addSpotHighlight(target.latlng);
        if ((target.type === 'gather' || target.source === 'gather') && target.nodeKey) {
            const m = _gatherMarkerByKey.get(target.nodeKey);
            if (m) setTimeout(() => m.openPopup(), 450);
        } else if (target.source === 'shop' && target.shopKey != null) {
            const m = _shopMarkerByNpcId.get(target.shopKey);
            if (m) setTimeout(() => m.openPopup(), 450);
        }
    }
}

function _navigateToSpotGlobal(item) {
    const info = mapParams[item.mapName];
    if (!info) return;
    const latlng = worldToPixel(item.worldPos.x, item.worldPos.z, info);
    const target = { ...item, latlng };

    const isSameStage = item.mapName === _loadedMapName &&
        (item.stageId === currentStageName() || (!item.stageId && !currentStageName()));

    if (isSameStage) {
        _navigateToSpot(target);
    } else {
        _pendingGlobalNavTarget = target;
        navigateTo(item.mapName, item.stageId);
    }
}

async function buildSpotIndex(info) {
    _spotIndex = [];
    if (!info.stages?.length) return;

    const stid = currentStageName();
    const stagesToLoad = (stid && info.stages.includes(stid)) ? [stid] : info.stages;
    await fetchEnemyPositionsStages(stagesToLoad.map((id) => parseInt(id.slice(2), 10)));

    for (const stageId of stagesToLoad) {
        const stageNo = String(parseInt(stageId.slice(2), 10));
        const serverStageId = stageIds[stageNo];
        const cache = _enemySpawnCache;  // may be null if promise not yet resolved

        // ── Enemies: one entry per emCode per spawn position ────────────────
        const groups = enemyPositionsForStage(stageNo);
        if (groups) {
            for (const [groupId, groupData] of Object.entries(groups)) {
                const spawns = groupData.spawns ?? groupData;  // back-compat: array may be direct
                if (!Array.isArray(spawns) || !spawns.length) continue;
                for (let i = 0; i < spawns.length; i++) {
                    const s = spawns[i];
                    const posLatlng = worldToPixel(s.Position.x, s.Position.z, info);
                    const spawnKey  = serverStageId != null ? `${serverStageId},${groupId},${s.posIdx ?? i}` : null;
                    if (cache && spawnKey) {
                        for (const [emCode, lvSet] of groupSpawnLevelsByEmCode(cache.get(spawnKey) ?? [])) {
                            const baseName = emNames[emCode]?.name;
                            if (!baseName) continue;
                            _spotIndex.push(makeIndexedEnemySpot({
                                emCode, lvSet, baseName, groupId,
                                latlng: posLatlng, spawnKey,
                                worldPos: { x: s.Position.x, y: s.Position.y, z: s.Position.z },
                                stageId,
                            }));
                        }
                    } else if (s.EmName) {
                        _spotIndex.push({
                            type: 'enemy', name: s.EmName,
                            searchText: s.EmName.toLowerCase(),
                            latlng: posLatlng, groupId, emCode: null, spawnKey: null,
                            worldPos: { x: s.Position.x, y: s.Position.y, z: s.Position.z },
                            previewLines: [`<b>${s.EmName}</b>`, `Group ${groupId}`],
                            stageId,
                        });
                    }
                }
            }
        }

        // ── Gathering: one entry per node ───────────────────────────────────
        const nodes = gatherPoints[stageNo];
        if (nodes) {
            for (const node of nodes) {
                const label = GATHER_LABELS[node.type]
                    ?? node.type.replace(/^(OM_GATHER_|CHEST_)/, '').replace(/_/g, ' ');
                _spotIndex.push({
                    type:       'gather',
                    name:       label,
                    gatherType: node.type,
                    searchText: label.toLowerCase(),
                    latlng:     worldToPixel(node.x, node.z, info),
                    nodeKey:    `${stageNo}:${node.groupId}:${node.posId}`,
                    worldPos:   { x: node.x, y: node.y, z: node.z },
                    previewLines: [
                        `<b>${label}</b>`,
                        `X: ${node.x.toFixed(0)}  Z: ${node.z.toFixed(0)}`,
                    ],
                    stageId,
                });
            }
        }

        // ── Items: enemy drops ───────────────────────────────────────────────
        if (cache && serverStageId != null) {
            for (const [groupId, groupData] of Object.entries(groups ?? {})) {
                const spawns = groupData.spawns ?? groupData;
                if (!Array.isArray(spawns) || !spawns.length) continue;
                for (let i = 0; i < spawns.length; i++) {
                    const s = spawns[i];
                    const spawnKey = `${serverStageId},${groupId},${s.posIdx ?? i}`;
                    const posLatlng = worldToPixel(s.Position.x, s.Position.z, info);
                    const seen = new Set(); // dedup (itemId, emCode) within this position
                    for (const e of (cache.get(spawnKey) ?? [])) {
                        if (!e.emCode || !e.drops?.length) continue;
                        for (const row of e.drops) {
                            const itemId = row[0];
                            const dedup = `${itemId}\0${e.emCode}`;
                            if (seen.has(dedup)) continue;
                            seen.add(dedup);
                            const itemName = itemNames[String(itemId)]?.name ?? `Item #${itemId}`;
                            const emName   = emNames[e.emCode]?.name ?? e.emCode;
                            const qty = row[2] > row[1] ? `×${row[1]}–${row[2]}` : `×${row[1] ?? 1}`;
                            const pct = row[5] > 0 && row[5] < 1 ? ` (${Math.round(row[5] * 100)}%)` : '';
                            _spotIndex.push({
                                type: 'item', source: 'enemy',
                                name: itemName,
                                searchText: `${itemName} ${itemId}`.toLowerCase(),
                                itemId, latlng: posLatlng, groupId, emCode: e.emCode, spawnKey,
                                worldPos: { x: s.Position.x, y: s.Position.y, z: s.Position.z },
                                previewLines: [`<b>${itemName}</b>`, `Dropped by: ${emName}`, `${qty}${pct}`],
                                stageId,
                            });
                        }
                    }
                }
            }
        }

        // ── Items: gathering nodes ───────────────────────────────────────────
        if (_gatherItemsCache && serverStageId != null) {
            for (const node of (gatherPoints[stageNo] ?? [])) {
                const csvKey  = `${serverStageId},${node.groupId},${node.posId}`;
                const nodeItems = _gatherItemsCache.get(csvKey) ?? [];
                const latlng  = worldToPixel(node.x, node.z, info);
                const nodeLabel = GATHER_LABELS[node.type] ?? node.type.replace(/^(OM_GATHER_|CHEST_)/, '').replace(/_/g, ' ');
                for (const it of nodeItems) {
                    const itemName = itemNames[String(it.itemId)]?.name ?? `Item #${it.itemId}`;
                    const qty = it.maxItemNum > it.itemNum ? `×${it.itemNum}–${it.maxItemNum}` : `×${it.itemNum}`;
                    const pct = it.dropChance < 1 ? ` (${Math.round(it.dropChance * 100)}%)` : '';
                    _spotIndex.push({
                        type: 'item', source: 'gather',
                        name: itemName,
                        searchText: `${itemName} ${it.itemId}`.toLowerCase(),
                        itemId: it.itemId, latlng,
                        nodeKey: `${stageNo}:${node.groupId}:${node.posId}`,
                        worldPos: { x: node.x, y: node.y, z: node.z },
                        previewLines: [`<b>${itemName}</b>`, `From: ${nodeLabel}`, `${qty}${pct}`],
                        stageId,
                    });
                }
            }
        }

        // ── Items: shop NPCs ─────────────────────────────────────────────────
        if (_shopCache && isAccessibleShopStage(stageNo)) {
            for (const npc of (npcShops[stageNo] ?? [])) {
                if (npc.ShopId == null) continue;
                const shop = _shopCache.get(npc.ShopId);
                if (!shop?.items?.length) continue;
                const latlng  = worldToPixel(npc.Position.x, npc.Position.z, info);
                const npcName = npcNames[String(npc.NpcId)] ?? `NPC #${npc.NpcId}`;
                for (const it of shop.items) {
                    if (it.ItemId == null) continue;
                    const itemName = itemNames[String(it.ItemId)]?.name ?? `Item #${it.ItemId}`;
                    _spotIndex.push({
                        type: 'item', source: 'shop',
                        name: itemName,
                        searchText: `${itemName} ${it.ItemId}`.toLowerCase(),
                        itemId: it.ItemId, latlng, shopKey: `${stageNo}:${npc.NpcId}`,
                        worldPos: { x: npc.Position.x, y: npc.Position.y, z: npc.Position.z },
                        previewLines: [`<b>${itemName}</b>`, `Sold by: ${npcName}`,
                                       it.Price != null ? `${it.Price.toLocaleString()} gold` : ''].filter(Boolean),
                        stageId,
                    });
                }
            }
        }
    }
}

// Rebuild spot index as async data caches load (enemy spawns, gather items, shop data)
const _rebuildSpotIndex = () => {
    if (!_currentMapInfo) return;
    buildSpotIndex(_currentMapInfo).then(() => {
        if (document.getElementById('spot-panel')?.classList.contains('open')) _runSpotSearch();
    });
};
_enemySpawnPromise  .then(() => {
    _rebuildSpotIndex();
    _rebuildGlobalSpotIndex();
    // Refresh all chip icons and mark boss markers now that spawn data is available
    for (const g of _groupStore.values()) {
        if (_groupHasBoss(g)) {
            g.labelMarker.setIcon(makeChipIcon(g.groupId, g.color, g.items.length, g.isExpanded, g.yOffset, g.isKeyBearerGroup, true));
        }
        // Re-apply boss styling for any already-expanded groups
        for (const markers of Object.values(g.sgMarkers)) {
            for (const m of markers) {
                const entries = m._spawnKey ? (_enemySpawnCache.get(m._spawnKey) ?? []) : [];
                if (entries.some(e => e.isBossGauge || e.isAreaBoss || e.raidBossId > 0)) {
                    m.setStyle({ color: '#ff3333', fillColor: '#cc0000', weight: 3, radius: 8 });
                    m._origStyle = { ...m._origStyle, color: '#ff3333', fillColor: '#cc0000', weight: 3, radius: 8 };
                }
            }
        }
    }
}).catch(() => {});
_gatherItemsPromise .then(() => { _rebuildSpotIndex(); _rebuildGlobalSpotIndex(); }).catch(() => {});
_shopPromise        .then(() => { _rebuildSpotIndex(); _rebuildGlobalSpotIndex(); }).catch(() => {});

async function _buildGlobalSpotIndex() {
    if (_globalSpotIndexReady) return;
    if (_globalSpotIndexPromise) return _globalSpotIndexPromise;

    _globalSpotIndexPromise = (async () => {
        const resultsEl = document.getElementById('spot-results');
        const showProgress = _spotGlobal && resultsEl;
        try {
            await preloadAllEnemyPositionsStages((done, total) => {
                if (showProgress) {
                    resultsEl.innerHTML =
                        `<div class="spot-empty">Loading world spawn index… ${done}/${total}</div>`;
                }
            });

            _globalSpotIndex = [];
            for (const [mapName, info] of Object.entries(mapParams)) {
                if (!info.stages?.length) continue;
                const mapDisplayName = info.name_en ? splitPascalCase(info.name_en) : mapName;
                for (const stageId of info.stages) {
                    const stageNo = String(parseInt(stageId.slice(2), 10));
                    const serverStageId = stageIds[stageNo];
                    const sLabel = stageLabel(info, stageId);
                    const locationTag = `${mapDisplayName} · ${sLabel}`;

                    // Enemies — same logic as local buildSpotIndex, using global spawn cache
                    const groups = enemyPositionsForStage(stageNo);
                    if (groups) {
                        for (const [groupId, groupData] of Object.entries(groups)) {
                            const spawns = groupData.spawns ?? groupData;
                            if (!Array.isArray(spawns) || !spawns.length) continue;
                            for (let i = 0; i < spawns.length; i++) {
                                const s = spawns[i];
                                const spawnKey = serverStageId != null ? `${serverStageId},${groupId},${s.posIdx ?? i}` : null;
                                if (_enemySpawnCache && spawnKey) {
                                    for (const [emCode, lvSet] of groupSpawnLevelsByEmCode(_enemySpawnCache.get(spawnKey) ?? [])) {
                                        const baseName = emNames[emCode]?.name;
                                        if (!baseName) continue;
                                        _globalSpotIndex.push(makeGlobalEnemySpot({
                                            emCode, lvSet, baseName, groupId, spawnKey,
                                            worldPos: { x: s.Position.x, y: s.Position.y, z: s.Position.z },
                                            mapName, stageId, stageNo, locationTag,
                                        }));
                                    }
                                } else if (s.EmName) {
                                    const emCode = s.EmName;
                                    const displayName = emNames[emCode]?.name ?? emCode;
                                    _globalSpotIndex.push({
                                        type: 'enemy', name: displayName,
                                        searchText: `${displayName} ${emCode}`.toLowerCase(),
                                        worldPos: { x: s.Position.x, y: s.Position.y, z: s.Position.z },
                                        groupId, emCode: null, spawnKey: null,
                                        mapName, stageId, stageNo, locationTag,
                                    });
                                }
                            }
                        }
                    }

                    // Enemy drops
                    if (_enemySpawnCache && serverStageId != null && groups) {
                        for (const [groupId, groupData] of Object.entries(groups)) {
                            const spawns = groupData.spawns ?? groupData;
                            if (!Array.isArray(spawns) || !spawns.length) continue;
                            for (let i = 0; i < spawns.length; i++) {
                                const s = spawns[i];
                                const spawnKey = `${serverStageId},${groupId},${s.posIdx ?? i}`;
                                const seen = new Set();
                                for (const e of (_enemySpawnCache.get(spawnKey) ?? [])) {
                                    if (!e.emCode || !e.drops?.length) continue;
                                    for (const row of e.drops) {
                                        const itemId = row[0];
                                        const dedup = `${itemId}\0${e.emCode}`;
                                        if (seen.has(dedup)) continue;
                                        seen.add(dedup);
                                        const itemName = itemNames[String(itemId)]?.name ?? `Item #${itemId}`;
                                        const emName   = emNames[e.emCode]?.name ?? e.emCode;
                                        const qty = row[2] > row[1] ? `×${row[1]}–${row[2]}` : `×${row[1] ?? 1}`;
                                        const pct = row[5] > 0 && row[5] < 1 ? ` (${Math.round(row[5] * 100)}%)` : '';
                                        _globalSpotIndex.push({
                                            type: 'item', source: 'enemy',
                                            name: itemName,
                                            searchText: `${itemName} ${itemId}`.toLowerCase(),
                                            itemId, groupId, emCode: e.emCode, spawnKey,
                                            dropDesc: `${emName} ${qty}${pct}`,
                                            worldPos: { x: s.Position.x, y: s.Position.y, z: s.Position.z },
                                            mapName, stageId, stageNo, locationTag,
                                        });
                                    }
                                }
                            }
                        }
                    }

                    // Gathering spots
                    const nodes = gatherPoints[stageNo];
                    if (nodes) {
                        for (const node of nodes) {
                            const label = GATHER_LABELS[node.type]
                                ?? node.type.replace(/^(OM_GATHER_|CHEST_)/, '').replace(/_/g, ' ');
                            _globalSpotIndex.push({
                                type: 'gather', name: label,
                                gatherType: node.type,
                                searchText: label.toLowerCase(),
                                worldPos: { x: node.x, y: node.y, z: node.z },
                                nodeKey: `${stageNo}:${node.groupId}:${node.posId}`,
                                mapName, stageId, stageNo, locationTag,
                            });
                        }
                    }

                    // Shop items (uses globally-loaded shop cache)
                    if (_shopCache && isAccessibleShopStage(stageNo)) {
                        for (const npc of (npcShops[stageNo] ?? [])) {
                            if (npc.ShopId == null) continue;
                            const shop = _shopCache.get(npc.ShopId);
                            if (!shop?.items?.length) continue;
                            const npcName = npcNames[String(npc.NpcId)] ?? `NPC #${npc.NpcId}`;
                            for (const it of shop.items) {
                                if (it.ItemId == null) continue;
                                const itemName = itemNames[String(it.ItemId)]?.name ?? `Item #${it.ItemId}`;
                                _globalSpotIndex.push({
                                    type: 'item', source: 'shop',
                                    name: itemName,
                                    searchText: `${itemName} ${it.ItemId}`.toLowerCase(),
                                    itemId: it.ItemId,
                                    worldPos: { x: npc.Position.x, y: npc.Position.y, z: npc.Position.z },
                                    shopKey: `${stageNo}:${npc.NpcId}`,
                                    mapName, stageId, stageNo, locationTag,
                                });
                            }
                        }
                    }

                    // Gather items (uses globally-loaded gather items cache)
                    if (_gatherItemsCache && serverStageId != null) {
                        for (const node of (gatherPoints[stageNo] ?? [])) {
                            const csvKey = `${serverStageId},${node.groupId},${node.posId}`;
                            const nodeItems = _gatherItemsCache.get(csvKey) ?? [];
                            for (const it of nodeItems) {
                                const itemName = itemNames[String(it.itemId)]?.name ?? `Item #${it.itemId}`;
                                _globalSpotIndex.push({
                                    type: 'item', source: 'gather',
                                    name: itemName,
                                    searchText: `${itemName} ${it.itemId}`.toLowerCase(),
                                    itemId: it.itemId,
                                    worldPos: { x: node.x, y: node.y, z: node.z },
                                    nodeKey: `${stageNo}:${node.groupId}:${node.posId}`,
                                    mapName, stageId, stageNo, locationTag,
                                });
                            }
                        }
                    }
                }
            }
            _globalSpotIndexReady = true;
        } finally {
            _globalSpotIndexPromise = null;
        }
    })();

    return _globalSpotIndexPromise;
}

const _rebuildGlobalSpotIndex = () => {
    if (!_globalSpotIndexReady) return;
    _globalSpotIndexReady = false;
    _globalSpotIndex = [];
    _globalSpotIndexPromise = null;
    _buildGlobalSpotIndex().then(() => {
        if (_spotGlobal && document.getElementById('spot-panel')?.classList.contains('open')) _runSpotSearch();
    });
};

// ── Spot search — pure helpers ───────────────────────────────────────────────

const SPOT_LEVEL_FILTER_KEY = 'ddon-spot-level-filter';
const SPOT_MIN_LEVEL_KEY_LEGACY = 'ddon-spot-min-level';

/** Game stage order from stage_list.slt.json (StageNo → list index). */
const _stageOrderByStageId = new Map(
    stageList.StageListInfoList.map((entry, index) => [
        `st${String(entry.StageNo).padStart(4, '0')}`,
        index,
    ]),
);

const parsePositiveInt = (value) => {
    const n = parseInt(String(value ?? '').trim(), 10);
    return Number.isFinite(n) && n > 0 ? n : 0;
};

const normalizeLevelBounds = ({ minLevel, maxLevel }) => {
    if (minLevel > 0 && maxLevel > 0 && minLevel > maxLevel) {
        return { minLevel: maxLevel, maxLevel: minLevel };
    }
    return { minLevel, maxLevel };
};

const readSpotLevelInputs = () => ({
    min: document.getElementById('spot-min-level')?.value ?? '',
    max: document.getElementById('spot-max-level')?.value ?? '',
});

const readSpotLevelBounds = () =>
    normalizeLevelBounds({
        minLevel: parsePositiveInt(readSpotLevelInputs().min),
        maxLevel: parsePositiveInt(readSpotLevelInputs().max),
    });

const hasActiveSpotLevelFilter = ({ minLevel, maxLevel }) =>
    minLevel > 0 || maxLevel > 0;

const enemyLevelSortKey = (entry) =>
    entry.minLv ?? entry.maxLv ?? Number.MAX_SAFE_INTEGER;

const stageOrderKey = (entry) =>
    _stageOrderByStageId.get(entry.stageId) ?? Number.MAX_SAFE_INTEGER;

const compareSpotResultNames = (a, b) =>
    a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' });

/** Hidell Plains overworld — distance sort anchor (WDT entrance on st0100). */
const HIDELL_PLAINS_MAP = 'field000_m00';

const readWhiteDragonTempleAnchor = () => {
    const entrance = (connectionData[HIDELL_PLAINS_MAP] ?? []).find(c =>
        c.to_map === 'lb000_m00'
        && c.x != null
        && c.z != null
        && (c.from_stage == null || c.from_stage === 100));
    return entrance ? { x: entrance.x, z: entrance.z } : null;
};

const _whiteDragonTempleAnchor = readWhiteDragonTempleAnchor();

const worldDistSqXZ = (pos, anchor) => {
    if (!pos || !anchor) return Number.MAX_SAFE_INTEGER;
    const dx = pos.x - anchor.x;
    const dz = pos.z - anchor.z;
    return dx * dx + dz * dz;
};

const SPOT_NO_PATH = Number.MAX_SAFE_INTEGER;
const _spotPathDistCache = new Map();

/** null = default WDT anchor; otherwise a custom map position. */
let _spotSortOrigin = null;
let _spotSetOriginMode = false;

const getDefaultWdtOrigin = () =>
    _whiteDragonTempleAnchor
        ? {
            mapName: HIDELL_PLAINS_MAP,
            x: _whiteDragonTempleAnchor.x,
            z: _whiteDragonTempleAnchor.z,
            label: 'White Dragon Temple',
        }
        : null;

const getSpotSortOrigin = () => _spotSortOrigin ?? getDefaultWdtOrigin();

const worldPosFromLatLng = (latlng, info) => {
    if (!info) return null;
    const px = latlng.lng;
    const py = latlng.lat;
    const png_y = info.img_height - py;
    const wx = (px - info.center_x) / info.scale;
    let wz;
    if (info.pd_pieces?.length) {
        const pieces = info.pd_pieces;
        let piece = pieces[0];
        for (const p of pieces) {
            if (png_y >= p.pixel_y_start && png_y <= p.pixel_y_entrance) { piece = p; break; }
        }
        wz = piece.connect_z + (png_y - piece.pixel_y_entrance_v) / info.scale;
    } else {
        const scaleZ = info.scale_z ?? info.scale;
        wz = ((info.img_height - info.center_y) - py) / scaleZ;
    }
    return { x: wx, z: wz };
};

/** Where you arrive on the destination map after taking conn from fromMap. */
const findConnectionArrivalPos = (fromMap, conn) => {
    const destMap = conn.to_map;
    if (!destMap) return null;
    const destConns = connectionData[destMap] || [];
    const rev = destConns.find(c =>
        c.to_map === fromMap && c.x != null && c.z != null
        && (conn.from_stage == null || c.to_stage === conn.from_stage)
    ) ?? destConns.find(c => c.to_map === fromMap && c.x != null && c.z != null);
    if (!rev) return null;
    return { mapName: destMap, x: rev.x, z: rev.z };
};

/** Shortest path distance (squared XZ) via positioned connection markers; SPOT_NO_PATH if unreachable. */
const pathDistSqFromOrigin = (origin, targetMap, targetPos) => {
    if (!origin || !targetPos || !targetMap) return SPOT_NO_PATH;

    const cacheKey = `${origin.mapName}:${origin.x},${origin.z}|${targetMap}:${targetPos.x},${targetPos.z}`;
    if (_spotPathDistCache.has(cacheKey)) return _spotPathDistCache.get(cacheKey);

    if (origin.mapName === targetMap) {
        const d = worldDistSqXZ(targetPos, origin);
        _spotPathDistCache.set(cacheKey, d);
        return d;
    }

    const best = new Map();
    best.set(origin.mapName, { cost: 0, x: origin.x, z: origin.z });
    const queue = [{ mapName: origin.mapName, cost: 0, x: origin.x, z: origin.z }];

    while (queue.length) {
        queue.sort((a, b) => a.cost - b.cost);
        const cur = queue.shift();
        const prev = best.get(cur.mapName);
        if (!prev || cur.cost > prev.cost) continue;

        for (const conn of (connectionData[cur.mapName] ?? [])) {
            if (!conn.to_map || conn.x == null || conn.z == null) continue;
            const arrival = findConnectionArrivalPos(cur.mapName, conn);
            if (!arrival) continue;
            const leg = worldDistSqXZ({ x: conn.x, z: conn.z }, { x: cur.x, z: cur.z });
            const newCost = cur.cost + leg;
            const existing = best.get(arrival.mapName);
            if (!existing || newCost < existing.cost) {
                best.set(arrival.mapName, { cost: newCost, x: arrival.x, z: arrival.z });
                queue.push({ mapName: arrival.mapName, cost: newCost, x: arrival.x, z: arrival.z });
            }
        }
    }

    const onTarget = best.get(targetMap);
    const result = onTarget
        ? onTarget.cost + worldDistSqXZ(targetPos, { x: onTarget.x, z: onTarget.z })
        : SPOT_NO_PATH;
    _spotPathDistCache.set(cacheKey, result);
    return result;
};

const spotEntryMapName = (entry, scope) =>
    entry.mapName ?? (scope === 'local' ? _loadedMapName : null);

const spotEntryPathDistSq = (entry, scope = 'local') => {
    if (!entry?.worldPos) return SPOT_NO_PATH;
    const origin = getSpotSortOrigin();
    if (!origin) return SPOT_NO_PATH;
    const targetMap = spotEntryMapName(entry, scope);
    if (!targetMap) return SPOT_NO_PATH;
    return pathDistSqFromOrigin(origin, targetMap, entry.worldPos);
};

const spotGroupPathDistSq = (items, scope = 'local') =>
    Math.min(...items.map(e => spotEntryPathDistSq(e, scope)));

const compareSpotEnemyGroups = (itemsA, itemsB, { scope = 'local', criteria = null } = {}) => {
    const a = itemsA[0];
    const b = itemsB[0];
    if (a.type !== b.type) return compareSpotResultNames(a, b);

    const distA = spotGroupPathDistSq(itemsA, scope);
    const distB = spotGroupPathDistSq(itemsB, scope);
    const aPath = distA < SPOT_NO_PATH;
    const bPath = distB < SPOT_NO_PATH;

    // 1. Path distance — nearer reachable spawns first
    if (aPath && bPath) {
        const distDiff = distA - distB;
        if (distDiff !== 0) return distDiff;
    } else if (aPath !== bPath) {
        return aPath ? -1 : 1;
    }

    // 2. Stage order (global results)
    if (scope === 'global') {
        const stDiff = stageOrderKey(a) - stageOrderKey(b);
        if (stDiff !== 0) return stDiff;
    }

    // 3. Level — tie-breaker after distance (applies with or without level filter)
    if (a.type === 'enemy' && b.type === 'enemy') {
        const lvDiff = enemyLevelSortKey(a) - enemyLevelSortKey(b);
        if (lvDiff !== 0) return lvDiff;
    }

    return compareSpotResultNames(a, b);
};

const compareSpotLocalGroups = (itemsA, itemsB, criteria = null) =>
    compareSpotEnemyGroups(itemsA, itemsB, { scope: 'local', criteria });

const compareSpotGlobalGroups = (itemsA, itemsB, criteria = null) =>
    compareSpotEnemyGroups(itemsA, itemsB, { scope: 'global', criteria });

const orderSpotGroupItems = (items, scope, criteria = null) => {
    if (!items[0]?.worldPos) return items;
    return [...items].sort((a, b) => {
        const distA = spotEntryPathDistSq(a, scope);
        const distB = spotEntryPathDistSq(b, scope);
        const aPath = distA < SPOT_NO_PATH;
        const bPath = distB < SPOT_NO_PATH;

        if (aPath && bPath) {
            const distDiff = distA - distB;
            if (distDiff !== 0) return distDiff;
        } else if (aPath !== bPath) {
            return aPath ? -1 : 1;
        }

        if (a.type === 'enemy' && b.type === 'enemy') {
            const lvDiff = enemyLevelSortKey(a) - enemyLevelSortKey(b);
            if (lvDiff !== 0) return lvDiff;
        }

        return 0;
    });
};

const sortSpotGroupedResults = (grouped, scope, criteria = null) => {
    const compare = scope === 'global' ? compareSpotGlobalGroups : compareSpotLocalGroups;
    return [...grouped.values()].sort((a, b) => compare(a, b, criteria));
};

const formatSpotSortOriginLabel = (origin) => {
    if (!origin) return 'No connection anchor';
    if (origin.label) return origin.label;
    const info = mapParams[origin.mapName];
    return info?.name_en ?? origin.mapName;
};

const syncSpotOriginMarker = () => {
    _spotOriginLayer.clearLayers();
    const origin = getSpotSortOrigin();
    if (!origin || origin.mapName !== _loadedMapName) return;
    const info = mapParams[_loadedMapName];
    if (!info) return;
    const latlng = worldToPixel(origin.x, origin.z, info);
    const icon = L.divIcon({
        className: 'spot-hl-outer',
        html: '<div class="spot-hl-origin"></div>',
        iconSize: [14, 14],
        iconAnchor: [7, 7],
    });
    L.marker(latlng, { icon, interactive: false }).addTo(_spotOriginLayer);
};

const syncSpotSortUI = () => {
    const origin = getSpotSortOrigin();
    document.getElementById('spot-sort-label')?.replaceChildren(
        document.createTextNode(formatSpotSortOriginLabel(origin)),
    );
    const badge = document.getElementById('spot-sort-badge');
    if (badge) {
        badge.textContent = _spotSortOrigin ? 'Custom' : 'WDT';
        badge.hidden = false;
    }
    const pickBtn = document.getElementById('spot-sort-pick');
    if (pickBtn) {
        pickBtn.classList.toggle('active', _spotSetOriginMode);
        pickBtn.textContent = _spotSetOriginMode ? 'Click map to set origin…' : 'Set from map click…';
    }
};

const setSpotSortOrigin = (origin) => {
    _spotSortOrigin = origin;
    _spotSetOriginMode = false;
    syncSpotSortUI();
    syncSpotOriginMarker();
    _spotPathDistCache.clear();
    if (document.getElementById('spot-panel')?.classList.contains('open')) _runSpotSearch();
};

const setSpotSortSectionExpanded = (expanded) => {
    const section = document.getElementById('spot-sort-section');
    const toggle = document.getElementById('spot-sort-toggle');
    if (!section || !toggle) return;
    section.classList.toggle('expanded', expanded);
    toggle.setAttribute('aria-expanded', expanded ? 'true' : 'false');
};

const cancelSpotSetOriginMode = () => {
    if (!_spotSetOriginMode) return;
    _spotSetOriginMode = false;
    syncSpotSortUI();
};

const clearStoredSpotLevelFilter = () => {
    try {
        localStorage.removeItem(SPOT_LEVEL_FILTER_KEY);
        localStorage.removeItem(SPOT_MIN_LEVEL_KEY_LEGACY);
    } catch (_) {}
};

const spawnOverlapsLevelFilter = (entry, { minLevel, maxLevel }) => {
    if (!hasActiveSpotLevelFilter({ minLevel, maxLevel })) return true;
    if (entry.minLv == null && entry.maxLv == null) return true;
    const lo = entry.minLv ?? entry.maxLv;
    const hi = entry.maxLv ?? entry.minLv;
    if (minLevel > 0 && hi < minLevel) return false;
    if (maxLevel > 0 && lo > maxLevel) return false;
    return true;
};

const formatSpotLevelFilterLabel = ({ minLevel, maxLevel }) => {
    if (minLevel > 0 && maxLevel > 0) return `Lv${minLevel}–${maxLevel}`;
    if (minLevel > 0) return `≥ Lv${minLevel}`;
    if (maxLevel > 0) return `≤ Lv${maxLevel}`;
    return '';
};

const formatSpotLevelFilterNote = ({ filter, minLevel, maxLevel }) => {
    if (filter !== 'enemy') return '';
    const label = formatSpotLevelFilterLabel({ minLevel, maxLevel });
    return label ? ` (${label})` : '';
};

const formatSpotLevelEmptyMessage = ({ minLevel, maxLevel }) => {
    if (minLevel > 0 && maxLevel > 0) {
        return `No enemy spawns between Lv${minLevel} and Lv${maxLevel} on this stage.`;
    }
    if (minLevel > 0) return `No enemy spawns at Lv${minLevel} or above on this stage.`;
    if (maxLevel > 0) return `No enemy spawns at Lv${maxLevel} or below on this stage.`;
    return 'No matches.';
};

const syncSpotLevelFilterBadge = () => {
    const badge = document.getElementById('spot-level-badge');
    if (!badge) return;
    const { minLevel, maxLevel } = readSpotLevelBounds();
    const label = formatSpotLevelFilterLabel({ minLevel, maxLevel });
    badge.textContent = label;
    badge.hidden = !label;
};

const syncSpotEnemyPanelSections = () => {
    const filter = document.querySelector('.spot-tab.active')?.dataset.filter ?? 'enemy';
    document.getElementById('spot-level-section')?.classList.toggle('visible', filter === 'enemy');
    document.getElementById('spot-sort-section')?.classList.toggle('visible',
        filter === 'enemy' || filter === 'gather' || filter === 'item');
};

const setSpotLevelSectionExpanded = (expanded) => {
    const section = document.getElementById('spot-level-section');
    const toggle = document.getElementById('spot-level-toggle');
    if (!section || !toggle) return;
    section.classList.toggle('expanded', expanded);
    toggle.setAttribute('aria-expanded', expanded ? 'true' : 'false');
};

function _parseSpotQuery(raw) {
    if (!raw) return { term: '', terms: [], exact: false };
    if (raw.length > 2 && raw.startsWith('"') && raw.endsWith('"') && raw.indexOf('"', 1) === raw.length - 1) {
        const inner = raw.slice(1, -1);
        return { term: inner, terms: [inner], exact: true };
    }
    const parts = raw.split(/\s*[,+]\s*/).map((s) => s.trim()).filter(Boolean);
    const terms = parts.map((p) => {
        if (p.length > 2 && p.startsWith('"') && p.endsWith('"')) return p.slice(1, -1);
        return p;
    });
    return { term: terms[0] ?? '', terms, exact: false };
}

function _spotEntryMatches(e, term, exact) {
    if (!term) return true;
    if (exact) {
        const name = e.name.toLowerCase();
        return name.startsWith(term) || e.searchText.startsWith(term);
    }
    if (!e.searchText.includes(term)) return false;
    if (e.type === 'enemy') {
        const base = spotEnemyBaseName(e).toLowerCase();
        const esc = term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        if (new RegExp(`\\b${esc}\\b`, 'i').test(base)) return true;
        if (e.emCode && e.emCode.toLowerCase().includes(term)) return true;
    }
    return true;
}

const enemyLevelRange = (levels) => {
    const sorted = [...levels].sort((a, b) => a - b);
    if (!sorted.length) return { minLv: null, maxLv: null, label: '' };
    const minLv = sorted[0];
    const maxLv = sorted[sorted.length - 1];
    const label = minLv === maxLv ? `Lv${minLv}` : `Lv${minLv}-${maxLv}`;
    return { minLv, maxLv, label };
};

const enemyDisplayName = (baseName, levelLabel) =>
    levelLabel ? `${baseName} ${levelLabel}` : baseName;

const groupSpawnLevelsByEmCode = (spawnEntries) =>
    spawnEntries.reduce((byEmCode, entry) => {
        if (!entry.emCode) return byEmCode;
        if (!byEmCode.has(entry.emCode)) byEmCode.set(entry.emCode, new Set());
        if (entry.lv != null) byEmCode.get(entry.emCode).add(entry.lv);
        return byEmCode;
    }, new Map());

const makeIndexedEnemySpot = ({ emCode, lvSet, baseName, groupId, ...rest }) => {
    const { label, minLv, maxLv } = enemyLevelRange(lvSet);
    const name = enemyDisplayName(baseName, label);
    return {
        type: 'enemy',
        name,
        searchText: `${baseName} ${emCode} ${label}`.toLowerCase(),
        minLv,
        maxLv,
        emCode,
        groupId,
        previewLines: [`<b>${name}</b>`, `Code: ${emCode}`, `Group ${groupId}`],
        ...rest,
    };
};

const makeGlobalEnemySpot = ({ emCode, lvSet, baseName, ...rest }) => {
    const { label, minLv, maxLv } = enemyLevelRange(lvSet);
    const name = enemyDisplayName(baseName, label);
    return {
        type: 'enemy',
        name,
        searchText: `${baseName} ${emCode} ${label}`.toLowerCase(),
        minLv,
        maxLv,
        emCode,
        ...rest,
    };
};

const readSpotSearchCriteria = () => {
    const rawInput = (document.getElementById('spot-search-input')?.value ?? '').trim();
    const parsed = _parseSpotQuery(rawInput.toLowerCase());
    const terms = parsed.terms.map((t) => t.toLowerCase());
    const { minLevel, maxLevel } = readSpotLevelBounds();
    return {
        raw: rawInput.toLowerCase(),
        rawDisplay: rawInput,
        term: (terms[0] ?? '').toLowerCase(),
        terms,
        exact: parsed.exact,
        multiTerm: terms.length >= 2,
        filter: document.querySelector('.spot-tab.active')?.dataset.filter ?? 'enemy',
        minLevel,
        maxLevel,
    };
};

const meetsSpotType = (entry, filter) => entry.type === filter;

const meetsSpotLevelFilter = (entry, filter, bounds) =>
    filter !== 'enemy' || spawnOverlapsLevelFilter(entry, bounds);

/** When "Apply to spot search" is on, enemy rows must match Spawns type checkboxes. */
const meetsSpotMobTypeFilter = (entry) => {
    if (!_applyMobTypesToSpotSearch || entry.type !== 'enemy') return true;
    if (!anyMobTypeEnabled()) return false;
    const tags = spotItemMobTags(entry);
    if (!tags.size) return false;
    return [...tags].some(t => _mobTypeFilters[t]);
};

const matchesSpotCriteria = (entry, criteria) =>
    meetsSpotType(entry, criteria.filter)
    && meetsSpotLevelFilter(entry, criteria.filter, criteria)
    && meetsSpotMobTypeFilter(entry)
    && (criteria.multiTerm || _spotEntryMatches(entry, criteria.term, criteria.exact));

const spotEnemyBaseName = (entry) => {
    if (entry.emCode && emNames[entry.emCode]?.name) return emNames[entry.emCode].name;
    return entry.name.replace(/\s+Lv[\d.-]+$/, '');
};

function* iterStagesForMultiMob(scope) {
    if (scope === 'global') {
        for (const [mapName, info] of Object.entries(mapParams)) {
            if (!info.stages?.length) continue;
            const mapDisplayName = info.name_en ? splitPascalCase(info.name_en) : mapName;
            for (const stageId of info.stages) {
                const stageNo = String(parseInt(stageId.slice(2), 10));
                const sLabel = stageLabel(info, stageId);
                yield {
                    stageNo,
                    stageId,
                    mapName,
                    info,
                    locationTag: `${mapDisplayName} · ${sLabel}`,
                };
            }
        }
        return;
    }
    if (!_currentMapInfo?.stages?.length) return;
    const stid = currentStageName();
    const stages = (stid && _currentMapInfo.stages.includes(stid))
        ? [stid]
        : _currentMapInfo.stages;
    for (const stageId of stages) {
        const stageNo = String(parseInt(stageId.slice(2), 10));
        yield {
            stageNo,
            stageId,
            mapName: _loadedMapName,
            info: _currentMapInfo,
            locationTag: stageLabel(_currentMapInfo, stageId),
        };
    }
}

function collectMultiMobStageMatches(criteria, scope) {
    const { terms } = criteria;
    if (terms.length < 2) return [];

    const cache = _enemySpawnCache;
    const buckets = new Map();

    for (const { stageNo, stageId, mapName, info, locationTag } of iterStagesForMultiMob(scope)) {
        const groups = enemyPositionsForStage(stageNo);
        if (!groups) continue;
        const serverStageId = stageIds[stageNo];
        const key = scope === 'global' ? `${mapName}\0${stageId}` : stageId;

        for (const [groupId, groupData] of Object.entries(groups)) {
            const spawns = groupData.spawns ?? groupData;
            if (!Array.isArray(spawns)) continue;
            for (let i = 0; i < spawns.length; i++) {
                const s = spawns[i];
                const spawnKey = serverStageId != null
                    ? `${serverStageId},${groupId},${s.posIdx ?? i}`
                    : null;
                const latlng = worldToPixel(s.Position.x, s.Position.z, info);
                const worldPos = { x: s.Position.x, y: s.Position.y, z: s.Position.z };

                const rows = cache && spawnKey
                    ? [...groupSpawnLevelsByEmCode(cache.get(spawnKey) ?? [])]
                        .map(([emCode, lvSet]) => ({ emCode, lvSet, baseName: emNames[emCode]?.name }))
                    : s.EmName
                        ? [{ emCode: s.EmName, lvSet: new Set(), baseName: emNames[s.EmName]?.name ?? s.EmName }]
                        : [];

                for (const { emCode, lvSet, baseName } of rows) {
                    if (!baseName) continue;
                    const entry = scope === 'global'
                        ? {
                            ...makeGlobalEnemySpot({
                                emCode, lvSet, baseName, groupId, spawnKey, worldPos,
                                mapName, stageId, stageNo, locationTag,
                            }),
                            latlng,
                            groupId,
                        }
                        : makeIndexedEnemySpot({
                            emCode, lvSet, baseName, groupId,
                            latlng, spawnKey, worldPos, stageId,
                        });

                    if (!meetsSpotLevelFilter(entry, 'enemy', criteria)) continue;
                    if (!meetsSpotMobTypeFilter(entry)) continue;

                    let termHit = false;
                    for (const term of terms) {
                        if (_spotEntryMatches(entry, term, criteria.exact)) {
                            termHit = true;
                            break;
                        }
                    }
                    if (!termHit) continue;

                    if (!buckets.has(key)) {
                        buckets.set(key, {
                            key,
                            stageNo,
                            stageId,
                            mapName,
                            locationTag,
                            termsMatched: new Set(),
                            entries: [],
                        });
                    }
                    const bucket = buckets.get(key);
                    bucket.entries.push(entry);
                    for (const term of terms) {
                        if (_spotEntryMatches(entry, term, criteria.exact)) {
                            bucket.termsMatched.add(term);
                        }
                    }
                }
            }
        }
    }

    const compareGroups = scope === 'global' ? compareSpotGlobalGroups : compareSpotLocalGroups;
    const results = [];
    for (const bucket of buckets.values()) {
        if (!bucket.termsMatched.size) continue;
        bucket.matchCount = bucket.termsMatched.size;
        bucket.totalTerms = terms.length;
        results.push(bucket);
    }

    results.sort((a, b) => {
        if (b.matchCount !== a.matchCount) return b.matchCount - a.matchCount;
        const itemsA = orderSpotGroupItems(a.entries, scope, criteria);
        const itemsB = orderSpotGroupItems(b.entries, scope, criteria);
        return compareGroups(itemsA, itemsB, criteria);
    });
    return results;
}

function findRosterEnemySpots(row) {
    // Same name + same level range (not emCode) — Undead Lv70 from different IDs merge;
    // Lv71 stays separate. Exact base name so "Death" ≠ "Death Knight".
    const want = row.baseName.toLowerCase();
    return _spotIndex.filter((e) => {
        if (e.type !== 'enemy') return false;
        if (spotEnemyBaseName(e).toLowerCase() !== want) return false;
        if (spotEnemyLevelKey(e) !== row.levelKey) return false;
        return meetsSpotMobTypeFilter(e);
    });
}

function buildStageEnemyRoster() {
    // Group by name + level (Rising-accurate): same label & level merge across emCodes.
    const byKey = new Map();
    for (const entry of _spotIndex) {
        if (entry.type !== 'enemy') continue;
        const baseName = spotEnemyBaseName(entry);
        const levelKey = spotEnemyLevelKey(entry);
        const key = `${baseName.toLowerCase()}\0${levelKey}`;
        if (!byKey.has(key)) {
            byKey.set(key, {
                baseName,
                levelKey,
                emCodes: new Set(),
                levels: new Set(),
                groupIds: new Set(),
            });
        }
        const row = byKey.get(key);
        if (entry.emCode) row.emCodes.add(entry.emCode);
        if (entry.minLv != null) row.levels.add(entry.minLv);
        if (entry.maxLv != null) row.levels.add(entry.maxLv);
        if (entry.groupId != null) row.groupIds.add(entry.groupId);
    }
    return [...byKey.values()]
        .map((row) => {
            const { label, minLv } = enemyLevelRange(row.levels);
            const emCode = row.emCodes.size === 1 ? [...row.emCodes][0] : null;
            return {
                baseName: row.baseName,
                levelKey: row.levelKey,
                emCode,
                levels: row.levels,
                groupIds: row.groupIds,
                minLv: minLv ?? 0,
                displayName: enemyDisplayName(row.baseName, label),
                searchName: row.baseName,
                count: row.groupIds.size,
            };
        })
        .sort((a, b) => {
            const byName = a.baseName.localeCompare(b.baseName, undefined, { sensitivity: 'base' });
            if (byName) return byName;
            return (a.minLv ?? 0) - (b.minLv ?? 0);
        });
}

function findRosterGatherSpots(row) {
    return _spotIndex.filter((e) => e.type === 'gather' && e.gatherType === row.gatherType);
}

function buildStageGatherRoster() {
    const byKey = new Map();
    for (const entry of _spotIndex) {
        if (entry.type !== 'gather') continue;
        const key = entry.gatherType ?? entry.name;
        if (!byKey.has(key)) {
            byKey.set(key, {
                name: entry.name,
                gatherType: entry.gatherType,
                nodeKeys: new Set(),
            });
        }
        const row = byKey.get(key);
        if (entry.nodeKey) row.nodeKeys.add(entry.nodeKey);
    }
    return [...byKey.values()]
        .map((row) => ({ ...row, displayName: row.name, count: row.nodeKeys.size }))
        .sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
}

function findRosterItemSpots(row) {
    return _spotIndex.filter((e) =>
        e.type === 'item' && e.source === row.source && e.name === row.name);
}

function rosterItemLocationKey(entry) {
    if (entry.nodeKey) return `n:${entry.nodeKey}`;
    if (entry.shopKey) return `s:${entry.shopKey}`;
    return `e:${entry.groupId ?? ''}:${entry.spawnKey ?? ''}:${entry.emCode ?? ''}`;
}

function buildStageItemRoster() {
    const byKey = new Map();
    for (const entry of _spotIndex) {
        if (entry.type !== 'item') continue;
        const key = `item:${entry.source}\0${entry.name}`;
        if (!byKey.has(key)) {
            byKey.set(key, {
                name: entry.name,
                source: entry.source,
                itemId: entry.itemId,
                locations: new Set(),
            });
        }
        byKey.get(key).locations.add(rosterItemLocationKey(entry));
    }
    return [...byKey.values()]
        .map((row) => ({ ...row, displayName: row.name, count: row.locations.size }))
        .sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
}

const filterSpotEntries = (entries, criteria) =>
    entries.filter(entry => matchesSpotCriteria(entry, criteria));

/** Level identity for a spot entry — same name+level merge; different levels stay apart. */
const spotEnemyLevelKey = (entry) => {
    const lo = entry.minLv ?? entry.maxLv ?? '';
    const hi = entry.maxLv ?? entry.minLv ?? '';
    if (lo === '' && hi === '') return '';
    return lo === hi ? `Lv${lo}` : `Lv${lo}-${hi}`;
};

/** Same name + same level (emCode may differ). "Death" ≠ "Death Knight". */
const spotEnemyGroupKey = (entry) =>
    `${spotEnemyBaseName(entry).toLowerCase()}\0${spotEnemyLevelKey(entry)}`;

const spotGroupKey = (entry, { global = false } = {}) => {
    if (entry.type === 'item') {
        const base = `item:${entry.source}\0${entry.name}`;
        return global ? `${base}\0${entry.mapName}\0${entry.stageId}` : base;
    }
    if (entry.type === 'enemy') {
        const base = `enemy\0${spotEnemyGroupKey(entry)}`;
        return global ? `${base}\0${entry.mapName}\0${entry.stageId}` : base;
    }
    const base = `${entry.type}\0${entry.name}`;
    return global ? `${base}\0${entry.mapName}\0${entry.stageId}` : base;
};

/** Label for a grouped enemy row (items already share name+level). */
const spotEnemyGroupDisplayName = (items) => {
    const baseName = spotEnemyBaseName(items[0]);
    const levels = new Set();
    for (const it of items) {
        if (it.minLv != null) levels.add(it.minLv);
        if (it.maxLv != null) levels.add(it.maxLv);
    }
    const { label } = enemyLevelRange(levels);
    return enemyDisplayName(baseName, label);
};

const spotGroupDisplayName = (items) => {
    const first = items[0];
    if (first.type === 'enemy') return spotEnemyGroupDisplayName(items);
    return first.name;
};

const formatSpotEmptyMessage = (criteria, { global = false } = {}) => {
    const levelNote = formatSpotLevelFilterNote(criteria);
    const mobNote = (_applyMobTypesToSpotSearch && criteria.filter === 'enemy')
        ? ' (Spawns filters applied)'
        : '';
    const label = criteria.rawDisplay || criteria.raw;
    if (global) {
        return `<div class="spot-empty">No matches for <em>${label}</em>${levelNote}${mobNote} across all stages.</div>`;
    }
    if (criteria.raw) {
        return `<div class="spot-empty">No matches for <em>${label}</em>${levelNote}${mobNote}.</div>`;
    }
    if (criteria.filter === 'enemy' && hasActiveSpotLevelFilter(criteria)) {
        return `<div class="spot-empty">${formatSpotLevelEmptyMessage(criteria)}</div>`;
    }
    return `<div class="spot-empty">No matches${levelNote}${mobNote}.</div>`;
};

const formatSpotMultiEmptyMessage = (criteria, { global = false } = {}) => {
    const levelNote = formatSpotLevelFilterNote(criteria);
    const scope = global ? ' across all stages' : ' on this stage';
    const termsLabel = criteria.rawDisplay || criteria.terms.join(', ');
    return `<div class="spot-empty">No stages match <em>${termsLabel}</em>${levelNote}${scope}.</div>`;
};

const SPOT_LIST_ICON_STYLE = 'font-size:10px;line-height:1;flex-shrink:0';

const spotListIconSpan = (glyph, color, extraClass = '') =>
    `<span class="spot-list-icon ${extraClass}" style="${SPOT_LIST_ICON_STYLE};color:${color}">${glyph}</span>`;

const SPOT_LIST_POI_ICON_SIZE = 20;

function spotListPoiImgHtml(src, { size = SPOT_LIST_POI_ICON_SIZE, extraFilter = '', extraClass = '' } = {}) {
    if (!src) return null;
    const filter = extraFilter ? `${POI_IMG_FILTER} ${extraFilter}` : POI_IMG_FILTER;
    return `<img class="spot-list-icon spot-list-poi-icon ${extraClass}" src="${src}" width="${size}" height="${size}" alt="" `
        + `style="flex-shrink:0;display:block;image-rendering:pixelated;filter:${filter};">`;
}

function gatherTypeFromNodeKey(nodeKey) {
    if (!nodeKey) return null;
    const [stageNo, groupId, posId] = nodeKey.split(':');
    if (!stageNo || groupId == null || posId == null) return null;
    const node = (gatherPoints[stageNo] ?? []).find(
        (n) => String(n.groupId) === groupId && String(n.posId) === posId,
    );
    return node?.type ?? null;
}

const SPOT_LIST_ITEM_ICON_SIZE = 20;

/** Item sprite — same ii icon shown in gather/shop/enemy popups. */
function spotItemSpriteIconHtml(itemId) {
    const entry = itemNames[String(itemId)];
    const iconNo = entry?.iconNo;
    const iconFile = iconNo != null ? `ii${String(iconNo).padStart(6, '0')}.png` : null;
    if (!iconFile || !_iconIdSet.has(iconNo)) {
        return `<span class="spot-list-icon" style="display:inline-block;width:${SPOT_LIST_ITEM_ICON_SIZE}px;flex-shrink:0"></span>`;
    }
    return `<img class="spot-list-icon spot-list-item-icon" src="images/icons/small/${iconFile}" `
        + `width="${SPOT_LIST_ITEM_ICON_SIZE}" height="${SPOT_LIST_ITEM_ICON_SIZE}" alt="" `
        + `style="flex-shrink:0;display:block;image-rendering:pixelated;">`;
}

/** Gather rows — icon matches the gather-node POI marker on the map. */
function spotGatherIconHtml(first) {
    if (!first || first.type !== 'gather') return spotListIconSpan('🌿', '#8c8');
    const gatherType = first.gatherType ?? gatherTypeFromNodeKey(first.nodeKey);
    const iconSrc = gatherType ? gatherMapIconSrc(gatherType) : null;
    if (iconSrc) return spotListPoiImgHtml(iconSrc);
    return `<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${GATHER_COLORS[gatherType] ?? '#aaa'};flex-shrink:0"></span>`;
}

/** Empty-search stage list — same type icons as search (boss ☠, orbs, etc.). */
const spotStageListEnemyIconHtml = (items) => spotSearchEnemyIconHtml(items);

/** Tags for one spot-index enemy row — scoped to that mob's emCode, not the whole group. */
function spotItemMobTags(it) {
    const tags = new Set();
    if (it.type !== 'enemy') return tags;

    let spawn = null;
    const g = it.groupId ? _groupStore.get(String(it.groupId)) : null;
    if (g && it.spawnKey) {
        for (const item of g.items) {
            const sid = stageIds[item.stageNo];
            if (sid == null) continue;
            const key = `${sid},${g.groupId},${item.spawn.posIdx ?? item.idx}`;
            if (key === it.spawnKey) {
                spawn = item.spawn;
                break;
            }
        }
    }
    // Global / roster before groups expand: resolve KeyBearer etc. from layout positions.
    if (!spawn && it.spawnKey) {
        const parts = String(it.spawnKey).split(',');
        const groupId = parts[1];
        const posIdx = parts[2];
        const stageNo = it.stageNo != null
            ? String(it.stageNo)
            : (it.stageId ? String(parseInt(it.stageId.slice(2), 10)) : null);
        if (stageNo != null && groupId != null && posIdx != null) {
            const gd = enemyPositionsForStage(stageNo)?.[groupId];
            const spawns = gd?.spawns ?? (Array.isArray(gd) ? gd : null);
            spawn = spawns?.find(s => String(s.posIdx ?? '') === String(posIdx)) ?? null;
        }
    }

    let entries = it.spawnKey && _enemySpawnCache
        ? filterEntriesBySpawnTime(_enemySpawnCache.get(it.spawnKey) ?? [])
        : [];
    if (it.emCode) entries = entries.filter((e) => e.emCode === it.emCode);

    if (isDynamicSpawnSlot(_enemySpawnCache, it.spawnKey) && !entries.some((e) => e.lv)) {
        tags.add('dynamic');
        return tags;
    }
    for (const t of spawnMobTags(entries, spawn, _enemySpawnCache, it.spawnKey)) tags.add(t);
    return tags;
}

/** Enemy rows — icon reflects mob type (boss, orb, dormant, etc.). */
function spotSearchEnemyIconHtml(items) {
    if (!items?.length) return spotListIconSpan('⚔', '#c88', 'spot-list-icon-search');
    const tags = new Set();
    for (const it of items) {
        for (const t of spotItemMobTags(it)) tags.add(t);
    }
    if (tags.has('boss'))      return spotListIconSpan('☠', '#f44', 'spot-list-icon-search');
    if (tags.has('keyBearer')) return spotListIconSpan('🗝', '#c8a000', 'spot-list-icon-search');
    if (tags.has('bloodOrb'))  return spotListIconSpan('🩸', '#e88', 'spot-list-icon-search');
    if (tags.has('highOrb'))   return spotListIconSpan('⭐', '#ec4', 'spot-list-icon-search');
    if (tags.has('manual'))    return spotListIconSpan('😴', '#9ab', 'spot-list-icon-search');
    if (tags.has('dynamic'))   return spotListIconSpan('⚡', '#bdf', 'spot-list-icon-search');
    return spotListIconSpan('⚔', '#c88', 'spot-list-icon-search');
}

/** Item rows — show the item's own sprite (same as inside gather/shop containers). */
function spotItemIconHtml(first, items) {
    if (!first || first.type !== 'item') return spotListIconSpan('📦', '#aaa');
    const itemId = first.itemId ?? items?.[0]?.itemId;
    return spotItemSpriteIconHtml(itemId);
}

function spotResultIconHtml(first, items) {
    if (first.type === 'gather') return spotGatherIconHtml(first);
    if (first.type === 'item') return spotItemIconHtml(first, items);
    if (first.type === 'enemy') return spotSearchEnemyIconHtml(items);
    return spotListIconSpan('⚔', '#c88');
}

/** ◀ 1/N ▶ controls — shared by single, multi-mob, and roster spot rows. */
function wireSpotRowNavigation(row, items, navigate) {
    if (!items.length) return;
    if (items.length === 1) {
        row.addEventListener('click', () => navigate(items[0]));
        return;
    }

    let idx = -1;
    const nav = document.createElement('div');
    nav.className = 'spot-nav';
    nav.style.cssText = 'display:flex;align-items:center;gap:2px;flex-shrink:0';
    nav.innerHTML =
        `<button class="spot-nav-btn spot-prev" title="Previous">&#9664;</button>`
        + `<span class="spot-nav-pos" style="font-size:0.68rem;color:var(--text-dim);min-width:28px;text-align:center">1/${items.length}</span>`
        + `<button class="spot-nav-btn spot-next" title="Next">&#9654;</button>`;

    const preview = row.querySelector('.spot-preview');
    if (preview) row.insertBefore(nav, preview);
    else row.appendChild(nav);

    const posEl = nav.querySelector('.spot-nav-pos');
    const goTo = (i) => {
        idx = (i + items.length) % items.length;
        posEl.textContent = `${idx + 1}/${items.length}`;
        navigate(items[idx]);
    };

    nav.querySelector('.spot-prev').addEventListener('click', (e) => {
        e.stopPropagation();
        goTo(idx <= 0 ? items.length - 1 : idx - 1);
    });
    nav.querySelector('.spot-next').addEventListener('click', (e) => {
        e.stopPropagation();
        goTo(idx + 1);
    });
    row.addEventListener('click', (e) => {
        if (e.target.closest('.spot-nav')) return;
        goTo(idx < 0 ? 0 : idx + 1);
    });
}

function _highlightSpotItems(items) {
    const drawnGroups = new Set();
    for (const item of items) {
        if (_currentFloorObbs && item.worldPos) {
            const f = getEnemyFloor(item.worldPos.x, item.worldPos.y, item.worldPos.z, _currentFloorObbs);
            if (f !== null && f !== currentLayer) continue;
        }
        const grp = item.groupId ? _groupStore.get(item.groupId) : null;
        if (item.type === 'enemy' && grp && !grp.isExpanded) {
            if (!drawnGroups.has(item.groupId)) {
                drawnGroups.add(item.groupId);
                _addChipHighlight(grp);
            }
        } else {
            _addSpotHighlight(_resolveSpotLatLng(item));
        }
    }
}

function wireSpotRowHoverHighlight(row, items) {
    if (!items.length) return;
    row.addEventListener('mouseenter', () => {
        _clearSpotHighlights();
        _highlightSpotItems(items);
    });
    row.addEventListener('mouseleave', _clearSpotHighlights);
}

function _renderStageRoster(resultsEl, {
    roster,
    emptyMessage,
    summaryNoun,
    findMatches,
    iconForRow,
    labelForRow,
}) {
    if (!roster.length) {
        resultsEl.innerHTML = `<div class="spot-empty">${emptyMessage}</div>`;
        return;
    }

    const criteria = readSpotSearchCriteria();
    const rows = roster.map((row) => {
        const matches = findMatches(row);
        const items = orderSpotGroupItems(matches, 'local', criteria);
        return { row, items };
    });
    const totalResults = rows.reduce((n, r) => n + r.items.length, 0);
    const uniqueCount = rows.length;

    const frag = document.createDocumentFragment();
    const summary = document.createElement('div');
    summary.className = 'spot-summary';
    summary.textContent =
        `${totalResults} result${totalResults !== 1 ? 's' : ''} · ${uniqueCount} unique `
        + `${summaryNoun}${uniqueCount !== 1 ? 's' : ''} on this stage · click to go`;
    frag.appendChild(summary);

    const enableHoverHl = uniqueCount <= 50;

    for (const { row, items } of rows) {
        const el = document.createElement('div');
        el.className = 'spot-result-row spot-roster-row';
        const rowLabel = labelForRow(row);
        el.title = `Go to ${rowLabel}`;
        el.innerHTML =
            iconForRow(row, items)
            + `<span class="spot-result-name">${row.displayName}</span>`;
        wireSpotRowNavigation(el, items, _navigateToSpot);
        if (enableHoverHl) wireSpotRowHoverHighlight(el, items);
        frag.appendChild(el);
    }

    resultsEl.innerHTML = '';
    resultsEl.appendChild(frag);
}

function _renderStageEnemyRoster(resultsEl) {
    const fullRoster = buildStageEnemyRoster();
    const roster = fullRoster.filter((row) => findRosterEnemySpots(row).length > 0);
    _renderStageRoster(resultsEl, {
        roster,
        emptyMessage: _applyMobTypesToSpotSearch && fullRoster.length
            ? 'No enemies match the current Spawns filters.'
            : 'No enemy data for this stage yet.',
        summaryNoun: 'enemy type',
        findMatches: findRosterEnemySpots,
        iconForRow: (_row, items) => spotStageListEnemyIconHtml(items),
        labelForRow: (row) => row.baseName,
    });
}

function _renderStageGatherRoster(resultsEl) {
    _renderStageRoster(resultsEl, {
        roster: buildStageGatherRoster(),
        emptyMessage: 'No gathering nodes on this stage.',
        summaryNoun: 'gather type',
        findMatches: findRosterGatherSpots,
        iconForRow: (_row, items) => spotGatherIconHtml(items[0]),
        labelForRow: (row) => row.name,
    });
}

function _renderStageItemRoster(resultsEl) {
    _renderStageRoster(resultsEl, {
        roster: buildStageItemRoster(),
        emptyMessage: 'No item data for this stage yet.',
        summaryNoun: 'item',
        findMatches: findRosterItemSpots,
        iconForRow: (_row, items) => spotItemIconHtml(items[0], items),
        labelForRow: (row) => row.name,
    });
}

function _renderMultiMobResults(matches, resultsEl, criteria, scope) {
    const byCount = new Map();
    for (const g of matches) {
        if (!byCount.has(g.matchCount)) byCount.set(g.matchCount, []);
        byCount.get(g.matchCount).push(g);
    }

    const frag = document.createDocumentFragment();
    const summary = document.createElement('div');
    summary.className = 'spot-summary';
    summary.textContent = `${matches.length} stage${matches.length !== 1 ? 's' : ''} · ${criteria.terms.length} enemies searched`;
    frag.appendChild(summary);

    const navigate = scope === 'global' ? _navigateToSpotGlobal : _navigateToSpot;

    for (const count of [...byCount.keys()].sort((a, b) => b - a)) {
        const section = byCount.get(count);
        const hdr = document.createElement('div');
        hdr.className = 'spot-section-header';
        hdr.textContent = count === criteria.terms.length
            ? `All ${count} matched (${section.length})`
            : `${count} of ${criteria.terms.length} matched (${section.length})`;
        frag.appendChild(hdr);

        for (const g of section) {
            const items = orderSpotGroupItems(g.entries, scope, criteria);
            const matchedNames = [...g.termsMatched].map((term) => {
                const hit = items.find((e) => _spotEntryMatches(e, term, criteria.exact));
                return hit ? spotEnemyBaseName(hit) : term;
            }).sort((a, b) => a.localeCompare(b));

            const locLabel = g.locationTag;
            const badge = `${g.matchCount}/${g.totalTerms}`;

            const row = document.createElement('div');
            row.className = 'spot-result-row';
            row.title = matchedNames.join(', ');
            row.innerHTML =
                spotSearchEnemyIconHtml(items)
                + `<span class="spot-result-name">${locLabel}</span>`
                + `<span class="spot-multi-badge">${badge}</span>`
                + `<div class="spot-preview">${matchedNames.map((n) => `<b>${n}</b>`).join('<br>')}`
                + `<br><span style="color:#667">${locLabel}</span></div>`;

            wireSpotRowNavigation(row, items, navigate);
            if (scope !== 'global') wireSpotRowHoverHighlight(row, items);
            frag.appendChild(row);
        }
    }

    resultsEl.innerHTML = '';
    resultsEl.appendChild(frag);
}

function _runSpotSearch() {
    const resultsEl = document.getElementById('spot-results');
    if (!resultsEl) return;

    const criteria = readSpotSearchCriteria();
    _clearSpotHighlights();

    if (_spotGlobal) {
        if (_globalSpotIndexPromise && !_globalSpotIndexReady) {
            resultsEl.innerHTML = `<div class="spot-empty">Loading world spawn index…</div>`;
            return;
        }
        if (!criteria.raw) {
            resultsEl.innerHTML = `<div class="spot-empty">Enter a search term to search across all stages.</div>`;
            return;
        }
        if (criteria.multiTerm && criteria.filter === 'enemy') {
            const groupMatches = collectMultiMobStageMatches(criteria, 'global');
            if (!groupMatches.length) {
                resultsEl.innerHTML = formatSpotMultiEmptyMessage(criteria, { global: true });
                return;
            }
            _renderMultiMobResults(groupMatches, resultsEl, criteria, 'global');
            return;
        }
        const matches = filterSpotEntries(_globalSpotIndex, criteria);
        if (!matches.length) {
            resultsEl.innerHTML = formatSpotEmptyMessage(criteria, { global: true });
            return;
        }
        _renderGlobalResults(matches, resultsEl, criteria);
        return;
    }

    if (!criteria.raw || !criteria.terms.length) {
        if (criteria.filter === 'enemy') {
            _renderStageEnemyRoster(resultsEl);
            return;
        }
        if (criteria.filter === 'gather') {
            _renderStageGatherRoster(resultsEl);
            return;
        }
        if (criteria.filter === 'item') {
            _renderStageItemRoster(resultsEl);
            return;
        }
        resultsEl.innerHTML = `<div class="spot-empty">Enter a search term to search this stage.</div>`;
        return;
    }

    if (criteria.multiTerm && criteria.filter === 'enemy') {
        const groupMatches = collectMultiMobStageMatches(criteria, 'local');
        if (!groupMatches.length) {
            resultsEl.innerHTML = formatSpotMultiEmptyMessage(criteria);
            return;
        }
        _renderMultiMobResults(groupMatches, resultsEl, criteria, 'local');
        return;
    }

    const matches = filterSpotEntries(_spotIndex, criteria);
    if (!matches.length) {
        resultsEl.innerHTML = formatSpotEmptyMessage(criteria);
        return;
    }

    // Group by name+type (enemies: base name so emCode variants merge like Rising)
    const grouped = new Map();
    for (const m of matches) {
        const key = spotGroupKey(m);
        if (!grouped.has(key)) grouped.set(key, []);
        grouped.get(key).push(m);
    }

    const sortedGroups = sortSpotGroupedResults(grouped, 'local', criteria);
    const totalGroups = sortedGroups.length;
    const cappedGroups = sortedGroups.slice(0, SPOT_RESULT_CAP);

    const frag = document.createDocumentFragment();
    const summary = document.createElement('div');
    summary.className = 'spot-summary';
    let summaryText = `${matches.length} result${matches.length !== 1 ? 's' : ''} · ${totalGroups} unique`;
    if (totalGroups > SPOT_RESULT_CAP) {
        summaryText += ` · showing ${SPOT_RESULT_CAP} (refine search)`;
    }
    summary.textContent = summaryText;
    frag.appendChild(summary);

    const enableHoverHl = totalGroups <= 50;

    for (const rawItems of cappedGroups) {
        const items = orderSpotGroupItems(rawItems, 'local', criteria);
        const first = items[0];
        const multi = items.length > 1;
        const rowName = spotGroupDisplayName(items);

        const dotHtml = spotResultIconHtml(first, items);

        const previewHtml = first.previewLines.join('<br>');

        const row = document.createElement('div');
        row.className = 'spot-result-row';
        row.title = rowName;

        row.innerHTML =
            `${dotHtml}<span class="spot-result-name">${rowName}</span>`
            + (multi ? '' : `<div class="spot-preview">${previewHtml}</div>`);
        if (multi) {
            const preview = document.createElement('div');
            preview.className = 'spot-preview';
            preview.innerHTML = previewHtml;
            row.appendChild(preview);
        }
        wireSpotRowNavigation(row, items, _navigateToSpot);

        if (enableHoverHl) {
            wireSpotRowHoverHighlight(row, items);
        }

        frag.appendChild(row);
    }

    resultsEl.innerHTML = '';
    resultsEl.appendChild(frag);
}

function _renderGlobalResults(matches, resultsEl, criteria = null) {
    // Group by name + type + source + stage — enemies merge by base name (Rising-style)
    const grouped = new Map();
    for (const m of matches) {
        const key = spotGroupKey(m, { global: true });
        if (!grouped.has(key)) grouped.set(key, []);
        grouped.get(key).push(m);
    }

    const sortedGroups = sortSpotGroupedResults(grouped, 'global', criteria);
    const totalGroups = sortedGroups.length;
    const cappedGroups = sortedGroups.slice(0, SPOT_RESULT_CAP);

    const uniqueNames = new Set(
        matches.map(m => m.type === 'enemy' ? spotEnemyGroupKey(m) : m.name),
    ).size;
    const frag = document.createDocumentFragment();
    const summary = document.createElement('div');
    summary.className = 'spot-summary';
    let summaryText = `${matches.length} result${matches.length !== 1 ? 's' : ''} · ${uniqueNames} unique · ${totalGroups} stage entries`;
    if (totalGroups > SPOT_RESULT_CAP) {
        summaryText += ` · showing ${SPOT_RESULT_CAP} (refine search)`;
    }
    summary.textContent = summaryText;
    frag.appendChild(summary);

    for (const rawItems of cappedGroups) {
        const items = orderSpotGroupItems(rawItems, 'global', criteria);
        const first = items[0];
        const rowName = spotGroupDisplayName(items);

        const dotHtml = spotResultIconHtml(first, items);

        const row = document.createElement('div');
        row.className = 'spot-result-row';
        row.title = `${rowName} — ${first.locationTag}`;

        row.innerHTML =
            `${dotHtml}<div style="flex:1;min-width:0">`
            + `<div class="spot-result-name">${rowName}</div>`
            + `<div class="spot-stage-sub">${first.locationTag}</div>`
            + `</div>`;
        wireSpotRowNavigation(row, items, _navigateToSpotGlobal);
        // No hover highlight effects in global mode (results may be on other stages)

        frag.appendChild(row);
    }

    resultsEl.innerHTML = '';
    resultsEl.appendChild(frag);
}

// Panel wiring — called once on startup
(function initSpotPanel() {
    const panel  = document.getElementById('spot-panel');
    const toggle = document.getElementById('spot-panel-toggle');
    const close  = document.getElementById('spot-panel-close');
    const input  = document.getElementById('spot-search-input');
    const clearBtn = document.getElementById('spot-search-clear');
    if (!panel || !toggle || !close || !input) return;

    const openPanel = () => {
        panel.classList.add('open');
        toggle.style.display = 'none';
        input.focus();
        _runSpotSearch();
    };
    const closePanel = () => {
        panel.classList.remove('open');
        toggle.style.display = '';
        cancelSpotSetOriginMode();
        _clearSpotHighlights();
    };

    toggle.addEventListener('click', openPanel);
    close.addEventListener('click', closePanel);
    const _runSpotSearchDebounced = debounce(() => {
        requestAnimationFrame(() => _runSpotSearch());
    }, 300);
    input.addEventListener('input', () => {
        if (clearBtn) clearBtn.style.display = input.value ? 'block' : 'none';
        _runSpotSearchDebounced();
    });
    if (clearBtn) {
        clearBtn.addEventListener('click', () => {
            input.value = '';
            clearBtn.style.display = 'none';
            input.focus();
            _runSpotSearch();
        });
    }

    const onLevelFilterChange = debounce(() => {
        syncSpotLevelFilterBadge();
        _runSpotSearch();
    }, 250);

    const minLevelInput = document.getElementById('spot-min-level');
    const maxLevelInput = document.getElementById('spot-max-level');
    clearStoredSpotLevelFilter();
    minLevelInput?.addEventListener('input', onLevelFilterChange);
    maxLevelInput?.addEventListener('input', onLevelFilterChange);

    document.getElementById('spot-level-toggle')?.addEventListener('click', () => {
        const section = document.getElementById('spot-level-section');
        setSpotLevelSectionExpanded(!section?.classList.contains('expanded'));
    });

    document.getElementById('spot-sort-toggle')?.addEventListener('click', () => {
        const section = document.getElementById('spot-sort-section');
        setSpotSortSectionExpanded(!section?.classList.contains('expanded'));
    });

    document.getElementById('spot-sort-pick')?.addEventListener('click', () => {
        _spotSetOriginMode = !_spotSetOriginMode;
        syncSpotSortUI();
    });

    leafletMap.on('click', (e) => {
        if (!_spotSetOriginMode || !_currentMapInfo || !_loadedMapName) return;
        const pos = worldPosFromLatLng(e.latlng, _currentMapInfo);
        if (!pos) return;
        const mapLabel = mapParams[_loadedMapName]?.name_en ?? _loadedMapName;
        setSpotSortOrigin({
            mapName: _loadedMapName,
            x: pos.x,
            z: pos.z,
            label: `My location (${mapLabel})`,
        });
    });

    syncSpotLevelFilterBadge();
    syncSpotEnemyPanelSections();
    syncSpotSortUI();
    syncSpotOriginMarker();

    document.querySelectorAll('.spot-tab').forEach(btn =>
        btn.addEventListener('click', () => {
            document.querySelectorAll('.spot-tab').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            syncSpotEnemyPanelSections();
            _runSpotSearch();
        })
    );

    document.querySelectorAll('.spot-scope').forEach(btn =>
        btn.addEventListener('click', () => {
            document.querySelectorAll('.spot-scope').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            _spotGlobal = btn.dataset.scope === 'global';
            _clearSpotHighlights();
            if (_spotGlobal && !_globalSpotIndexReady) {
                _buildGlobalSpotIndex().then(() => _runSpotSearch());
            } else {
                _runSpotSearch();
            }
        })
    );

    document.addEventListener('keydown', e => {
        if ((e.ctrlKey || e.metaKey) && e.key === 'f' && !e.shiftKey && _loadedMapName) {
            e.preventDefault();
            panel.classList.contains('open') ? (input.focus(), input.select()) : openPanel();
        }
        if (e.key === 'Escape' && panel.classList.contains('open')) closePanel();
    });

    // Resize handled by shared _initPanelResize (called after this IIFE)
    openPanel();
})();

// ── Coordinate readout ────────────────────────────────────────────────────────
// Shows pixel and world coordinates under the cursor, useful for calibration.
(function () {
    const el = document.getElementById('coord-display');
    if (!el) return;

    let currentInfo = null;

    // Keep a reference to the current map info so we can invert the transform
    window._setCurrentInfo = (info) => { currentInfo = info; };

    leafletMap.on('mousemove', (e) => {
        if (!currentInfo) return;
        const px = e.latlng.lng;
        const py = e.latlng.lat;
        const wx = (px - currentInfo.center_x) / currentInfo.scale;
        let wz;
        if (currentInfo.pd_pieces?.length) {
            const png_y = currentInfo.img_height - py;
            const pieces = currentInfo.pd_pieces;
            let piece = pieces[0];
            for (const p of pieces) {
                if (png_y >= p.pixel_y_start && png_y <= p.pixel_y_entrance) {
                    piece = p; break;
                }
            }
            // Inverse of: png_y = pixel_y_entrance_v + localZ * scale
            wz = piece.connect_z + (png_y - piece.pixel_y_entrance_v) / currentInfo.scale;
        } else {
            const scaleZ = currentInfo.scale_z ?? currentInfo.scale;
            wz = ((currentInfo.img_height - currentInfo.center_y) - py) / scaleZ;
        }
        const [gx, gy] = pixelToGrid(px, py, currentInfo.img_height);
        el.textContent = `(${gx}, ${gy})   world (${wx.toFixed(0)}, ${wz.toFixed(0)})`;
    });

    leafletMap.on('mouseout', () => { el.textContent = ''; });

    // Alt+click → log coordinates to console and copy to clipboard.
    // Use this for calibration: Alt+click on a known map feature (bridge, door, etc.)
    // then compare the logged world coords with lot.json / connections.json values.
    leafletMap.on('click', (e) => {
        if (!e.originalEvent.altKey) return;
        if (!currentInfo) return;
        const px   = e.latlng.lng;
        const py   = e.latlng.lat;
        const png_y = currentInfo.img_height - py;
        const wx   = (px - currentInfo.center_x) / currentInfo.scale;
        let wz;
        if (currentInfo.pd_pieces?.length) {
            const pieces = currentInfo.pd_pieces;
            let piece = pieces[0];
            for (const p of pieces) {
                if (png_y >= p.pixel_y_start && png_y <= p.pixel_y_entrance) { piece = p; break; }
            }
            wz = piece.connect_z + (png_y - piece.pixel_y_entrance_v) / currentInfo.scale;
        } else {
            const scaleZ = currentInfo.scale_z ?? currentInfo.scale;
            wz = ((currentInfo.img_height - currentInfo.center_y) - py) / scaleZ;
        }
        const mapKey = location.hash.split(':')[0].replace('#', '');
        const msg = `CALIB [${currentInfo.name_en ?? mapKey}]  pixel=(${px.toFixed(1)}, ${png_y.toFixed(1)})  world X=${wx.toFixed(1)} Z=${wz.toFixed(1)}`;
        console.log(msg);
        navigator.clipboard?.writeText(msg).catch(() => {});
        // Flash the coord display to confirm
        el.style.color = '#ffd700';
        setTimeout(() => { el.style.color = ''; }, 600);
    });
})();

// ── Panel resize helpers ───────────────────────────────────────────────────────
function _initPanelResize({ handleId, panelId, lsKey, minW, maxW, dragDir }) {
    const handle = document.getElementById(handleId);
    const panel  = document.getElementById(panelId);
    if (!handle || !panel) return;

    const saved = parseInt(localStorage.getItem(lsKey), 10);
    if (saved >= minW && saved <= maxW) panel.style.width = saved + 'px';

    handle.addEventListener('mousedown', e => {
        e.preventDefault();
        const startX     = e.clientX;
        const startWidth = panel.offsetWidth;
        panel.classList.add('resizing');

        const onMove = (e) => {
            const delta  = dragDir === 'left' ? startX - e.clientX : e.clientX - startX;
            const newW   = Math.min(maxW, Math.max(minW, startWidth + delta));
            panel.style.width = newW + 'px';
        };
        const onUp = () => {
            panel.classList.remove('resizing');
            localStorage.setItem(lsKey, panel.offsetWidth);
            document.removeEventListener('mousemove', onMove);
            document.removeEventListener('mouseup', onUp);
            refitActiveMapView();
        };
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
    });
}

_initPanelResize({ handleId: 'sidebar-resize-handle', panelId: 'sidebar',    lsKey: 'ddon-sidebar-width',    minW: 180, maxW: 520, dragDir: 'right' });
_initPanelResize({ handleId: 'spot-resize-handle',    panelId: 'spot-panel', lsKey: 'ddon-spot-panel-width', minW: 200, maxW: 600, dragDir: 'left'  });

// Patch loadMap to keep currentInfo updated
const _origLoadMap = loadMap;
loadMap = async function (mapName) {
    await _origLoadMap(mapName);
    if (window._setCurrentInfo) window._setCurrentInfo(mapParams[mapName]);
    syncSpotOriginMarker();
};


// ── Edit mode (disabled on Rising — no edit panel) ─────────────────────────────
// Keep stubs so shared popup/render paths that check these remain safe.
_editMode = false;
_markDirty = null;
_attachDragReorder = null;
_renderEditPanel = null;

// ── Named param picker modal ──────────────────────────────────────────────────
function openNamedParamPicker(popupSection, baseEmName) {
    const modal = document.getElementById('named-param-modal');
    if (!modal) return;

    const searchInput  = modal.querySelector('#np-search');
    const list         = modal.querySelector('#np-list');
    const preview      = modal.querySelector('#np-preview');
    const toggleBtn    = modal.querySelector('#np-stat-toggle');

    // Current value in the popup section
    const hiddenInput  = popupSection?.querySelector('[data-edit="namedId"]');
    const currentId    = parseInt(hiddenInput?.value) || 0;

    // Build the combined enemy name HTML for a param, highlighting the param part
    function combinedNameHtml(p) {
        const trimName = p?.name?.trim();
        if (!trimName || !p || p.id === 0) return null;
        const em = baseEmName ?? '…';
        const hi = `<span class="np-name-hi">${trimName}</span>`;
        if (p.type === 'NAMED_TYPE_PREFIX')  return `${hi} ${em}`;
        if (p.type === 'NAMED_TYPE_SUFFIX')  return `${em} ${hi}`;
        if (p.type === 'NAMED_TYPE_REPLACE') return hi;
        return null; // NONE — no name change
    }

    function pct(v) { return v != null ? `${v}%` : '—'; }
    function renderPreview(p) {
        if (!p || p.id === 0) { preview.innerHTML = '<div class="np-preview-empty">Select a param to preview stats</div>'; return; }
        const typeName = p.type.replace('NAMED_TYPE_', '');
        const combined = combinedNameHtml(p);
        const row = (label, val) => {
            const v = pct(val);
            const n = parseFloat(v);
            const cls = n > 100 ? 'np-high' : n < 100 ? 'np-low' : '';
            return `<tr><td>${label}</td><td${cls ? ` class="${cls}"` : ''}>${v}</td></tr>`;
        };
        const sec = (title, ...rows) =>
            `<tr class="np-stat-sec"><td colspan="2">${title}</td></tr>` + rows.join('');
        preview.innerHTML =
            (combined ? `<div class="np-preview-combined">${combined}</div>` : '') +
            `<div class="np-preview-type">${typeName} · ID ${p.id}</div>` +
            `<table class="np-stat-table">` +
            sec('HP',
                row('HP Rate',  p.hp),
                row('HP Sub',   p.hpSub)) +
            sec('Attack',
                row('Base Phys',  p.atkP),
                row('Base Magic', p.atkM),
                row('Wep Phys',   p.atkWepP),
                row('Wep Magic',  p.atkWepM)) +
            sec('Defence',
                row('Base Phys',  p.defP),
                row('Base Magic', p.defM),
                row('Wep Phys',   p.defWepP),
                row('Wep Magic',  p.defWepM),
                row('Guard Base', p.guardBase),
                row('Guard Wep',  p.guardWep)) +
            sec('Other',
                row('Ailment Dmg', p.ailment),
                row('Experience',  p.exp),
                row('Power',       p.power)) +
            sec('Endurance',
                row('Blow Main',    p.blowMain),
                row('Blow Sub',     p.blowSub),
                row('Down Main',    p.downMain),
                row('OCD',          p.ocd),
                row('Shake Main',   p.shakeMain),
                row('Shrink Main',  p.shrinkMain),
                row('Shrink Sub',   p.shrinkSub)) +
            `</table>`;
    }

    function itemHtml(p, active) {
        const trimName   = p.name?.trim();
        const combined   = combinedNameHtml(p);
        const typeSuffix = p.type === 'NAMED_TYPE_PREFIX' ? 'Pfx'
            : p.type === 'NAMED_TYPE_REPLACE'             ? 'Rep'
            : p.type === 'NAMED_TYPE_SUFFIX'              ? 'Sfx'
            : '';
        const nameHtml = combined ?? (trimName
            ? `<span class="np-name-hi">${trimName}</span>`
            : `<span style="color:#555">(unnamed)</span>`);
        return `<div class="np-item${active ? ' np-active' : ''}" data-id="${p.id}">` +
            `<span class="np-item-name">${nameHtml}</span>` +
            `<span class="np-item-meta">${typeSuffix ? `<span class="np-type-badge">${typeSuffix}</span> ` : ''}#${p.id}</span>` +
            `</div>`;
    }

    function renderList(query) {
        const q           = query.trim().toLowerCase();
        const showStatOnly = toggleBtn.classList.contains('active');

        if (!q && showStatOnly) {
            // No search + stat-only on: render two labelled sections, no slice
            const named    = namedParamList.filter(p =>  p.name?.trim().length > 0);
            const statOnly = namedParamList.filter(p => !p.name?.trim().length);
            const noneEntry = { id: 0, name: '(None)', type: 'NAMED_TYPE_NONE' };
            const divider = (label, count) =>
                `<div class="np-section-divider">${label} <span class="np-section-count">${count}</span></div>`;
            list.innerHTML =
                itemHtml(noneEntry, currentId === 0) +
                divider('Named', named.length) +
                named.map(p => itemHtml(p, p.id === currentId)).join('') +
                divider('Stat-only (unnamed)', statOnly.length) +
                statOnly.map(p => itemHtml(p, p.id === currentId)).join('');
        } else {
            const base = namedParamList.filter(p => {
                const hasName = p.name?.trim().length > 0;
                if (!showStatOnly && !hasName) return false;
                if (!q) return true;
                const nameMatch    = p.name?.toLowerCase().includes(q);
                const idMatch      = String(p.id).includes(q);
                const combined     = combinedNameHtml(p)?.replace(/<[^>]*>/g, '') ?? '';
                const combinedMatch = combined.toLowerCase().includes(q);
                return nameMatch || idMatch || combinedMatch;
            }).slice(0, 200);
            const results = (!q || '0'.includes(q) || 'none'.includes(q))
                ? [{ id: 0, name: '(None)', type: 'NAMED_TYPE_NONE' }, ...base]
                : base;
            if (!results.length) {
                list.innerHTML = '<div class="np-empty">No params match.</div>';
                return;
            }
            list.innerHTML = results.map(p => itemHtml(p, p.id === currentId)).join('');
        }
        list.querySelectorAll('.np-item').forEach(item => {
            item.addEventListener('mouseenter', () => {
                const p = namedParamsById.get(parseInt(item.dataset.id));
                renderPreview(p);
            });
            item.addEventListener('click', () => {
                const id = parseInt(item.dataset.id);
                const p  = namedParamsById.get(id);
                if (hiddenInput) {
                    hiddenInput.value = id;
                    // Fire a change event so watchEditChanges enables the Apply button
                    hiddenInput.dispatchEvent(new Event('change', { bubbles: true }));
                    const btn = popupSection.querySelector('.se-named-picker-btn');
                    if (btn) { btn.textContent = namedParamLabel(p); btn.dataset.namedId = id; }
                    // Live-update the enemy name shown at the top of the popup
                    const nameSpan = popupSection.closest('.leaflet-popup-content')?.querySelector('.se-enemy-name');
                    if (nameSpan) {
                        const npName = p?.name?.trim();
                        const combined = (!npName || p.type === 'NAMED_TYPE_NONE') ? baseEmName
                            : p.type === 'NAMED_TYPE_REPLACE' ? npName
                            : p.type === 'NAMED_TYPE_PREFIX'  ? (baseEmName ? `${npName} ${baseEmName}` : npName)
                            : p.type === 'NAMED_TYPE_SUFFIX'  ? (baseEmName ? `${baseEmName} ${npName}` : npName)
                            : baseEmName;
                        if (combined) {
                            // Preserve the " LvN" suffix already in the span text
                            const lvMatch = nameSpan.textContent.match(/ Lv\d+$/);
                            nameSpan.textContent = combined + (lvMatch?.[0] ?? '');
                        }
                    }
                    // Update companion named-stats panel
                    const anchorEl = popupSection.closest('.leaflet-popup');
                    showNamedStatsPanel(id, anchorEl, baseEmName);
                }
                modal.classList.remove('open');
            });
        });
    }

    // Replace signal controller each open to avoid duplicate listeners
    if (modal._abortCtrl) modal._abortCtrl.abort();
    modal._abortCtrl = new AbortController();
    const sig = modal._abortCtrl.signal;
    searchInput.addEventListener('input', () => renderList(searchInput.value), { signal: sig });
    toggleBtn.addEventListener('click', () => {
        const on = toggleBtn.classList.toggle('active');
        toggleBtn.textContent = on ? 'Stat-only: On' : 'Stat-only: Off';
        renderList(searchInput.value);
    }, { signal: sig });

    searchInput.value = '';
    toggleBtn.classList.remove('active');
    toggleBtn.textContent = 'Stat-only: Off';
    renderPreview(namedParamsById.get(currentId) ?? null);
    renderList('');

    modal._popupSection = popupSection;
    modal.classList.add('open');
    setTimeout(() => searchInput.focus(), 50);
}

// ── Drop Table helpers ────────────────────────────────────────────────────────
function createDropTable() {
    const newId = (_rawEnemyData?.dropsTables?.length
        ? Math.max(..._rawEnemyData.dropsTables.map(t => t.id)) + 1 : 1);
    const dt = { id: newId, name: 'New Drop Table', mdlType: 0, items: [] };
    _dropsTablesMap.set(newId, dt);
    if (_rawEnemyData) {
        if (!_rawEnemyData.dropsTables) _rawEnemyData.dropsTables = [];
        _rawEnemyData.dropsTables.push(dt);
    }
    if (_markDirty) _markDirty('ddon-src-spawns');
    return dt;
}

function _dtItemRow(row, idx) {
    const itemId   = row[0] ?? 0;
    const nameHint = itemNames[String(itemId)]?.name ?? '';
    const iconNo   = itemNames[String(itemId)]?.iconNo;
    const iconFile = iconNo != null ? `ii${String(iconNo).padStart(6, '0')}.png` : null;
    const iconCell = iconFile && _iconIdSet.has(iconNo)
        ? `<img src="images/icons/small/${iconFile}" width="28" height="28" style="vertical-align:middle;image-rendering:pixelated">`
        : `<span style="display:inline-block;width:28px"></span>`;
    const href     = itemWikiHref(itemId);
    const nameCell = nameHint
        ? `<a href="${href}" target="_blank" class="dt-item-name-hint" data-name-for="${idx}" style="color:inherit;text-decoration:none" onmouseover="this.style.textDecoration='underline'" onmouseout="this.style.textDecoration='none'">${nameHint}</a>`
        : `<span class="dt-item-name-hint" data-name-for="${idx}"></span>`;
    return `<tr data-idx="${idx}">` +
        `<td><span class="dt-grab-handle" title="Drag to reorder">⠿</span></td>` +
        `<td style="text-align:center">${iconCell}</td>` +
        `<td><input type="number" data-col="0" value="${itemId}" style="width:58px" readonly tabindex="-1"></td>` +
        `<td>${nameCell}</td>` +
        `<td><input type="number" data-col="1" value="${row[1] ?? 1}" style="width:38px" min="1"></td>` +
        `<td><input type="number" data-col="2" value="${row[2] ?? 1}" style="width:38px" min="1"></td>` +
        `<td><input type="number" data-col="3" value="${row[3] ?? 0}" style="width:44px" min="0"></td>` +
        `<td style="text-align:center"><input type="checkbox" data-col="4"${row[4] ? ' checked' : ''}></td>` +
        `<td><input type="number" data-col="5" value="${((row[5] ?? 0) * 100).toFixed(1)}" style="width:58px" step="0.1" min="0" max="100"></td>` +
        `<td><button class="dt-del-row" data-row="${idx}">✕</button></td>` +
        `</tr>`;
}

function _buildDropChipsHtml(dt) {
    const items = dt?.items ?? [];
    if (!items.length) return '';
    return items.map(row => {
        const itemId   = row[0] ?? 0;
        const minQty   = row[1] ?? 1;
        const maxQty   = row[2] ?? 1;
        const dropRate = row[5] ?? 0;
        const iconNo   = itemNames[String(itemId)]?.iconNo;
        const iconFile = iconNo != null ? `ii${String(iconNo).padStart(6,'0')}.png` : null;
        const name     = itemNames[String(itemId)]?.name ?? `#${itemId}`;
        const qty      = maxQty > minQty ? `×${minQty}–${maxQty}` : `×${minQty}`;
        const pct      = dropRate > 0 ? ` ${(dropRate * 100).toFixed(0)}%` : '';
        const imgEl    = iconFile && _iconIdSet.has(iconNo)
            ? `<img src="images/icons/small/${iconFile}" width="20" height="20" style="image-rendering:pixelated;vertical-align:middle" title="${name}">`
            : `<span style="display:inline-block;width:20px;height:20px;background:#ddd;border-radius:2px;font-size:8px;text-align:center;line-height:20px;vertical-align:middle" title="${name}">${itemId}</span>`;
        return `<span style="display:inline-flex;align-items:center;gap:2px;white-space:nowrap">` +
            imgEl +
            `<span style="font-size:9px;color:#666">${qty}${pct}</span>` +
            `</span>`;
    }).join('');
}

function _updateDropsChips(popupSection, dt) {
    if (!popupSection) return;
    let chipsDiv = popupSection.querySelector('.se-drops-chips');
    const html = _buildDropChipsHtml(dt);
    if (html) {
        if (!chipsDiv) {
            // The chips div lives in a grp() row-wrapper sibling of the row1 row-wrapper.
            // grp() wraps each row in: <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:2px">
            // .se-drops-row1 is inside that row-wrapper, so parentElement is the row-wrapper.
            const row1Inner = popupSection.querySelector('.se-drops-row1');
            const rowWrapper = row1Inner?.parentElement;
            if (rowWrapper) {
                const newRowWrapper = document.createElement('div');
                newRowWrapper.style.cssText = 'display:flex;gap:6px;flex-wrap:wrap;margin-bottom:2px';
                chipsDiv = document.createElement('div');
                chipsDiv.className = 'se-drops-chips';
                chipsDiv.style.cssText = 'display:flex;flex-wrap:wrap;gap:4px 6px';
                newRowWrapper.appendChild(chipsDiv);
                rowWrapper.after(newRowWrapper);
            }
        }
        if (chipsDiv) chipsDiv.innerHTML = html;
    } else if (chipsDiv) {
        // Remove the row-wrapper too if it was dynamically inserted
        const parent = chipsDiv.parentElement;
        if (parent && !parent.classList.contains('se-drops-row1')) parent.remove();
        else chipsDiv.remove();
    }
}

function openDropTablePicker(popupSection) {
    const modal   = document.getElementById('dt-picker-modal');
    const search  = document.getElementById('dt-picker-search');
    const list    = document.getElementById('dt-picker-list');
    const preview = document.getElementById('dt-picker-preview');
    if (!modal) return;
    const hiddenInput = popupSection?.querySelector('[data-edit="dropsTableId"]');
    const currentId   = parseInt(hiddenInput?.value ?? -1);

    function renderPreview(dt) {
        if (!dt) { preview.innerHTML = '<div class="dt-preview-empty">Hover a table to preview</div>'; return; }
        if (!dt.items?.length) {
            preview.innerHTML = `<div class="dt-preview-title">${dt.name || '(unnamed)'}</div>` +
                `<div class="dt-preview-meta">id:${dt.id} · mdlType:${dt.mdlType ?? 0}</div>` +
                `<div class="dt-preview-empty">No items</div>`;
            return;
        }
        const rows = dt.items.map(row => {
            const itemId  = row[0] ?? 0;
            const name    = itemNames[String(itemId)]?.name ?? `#${itemId}`;
            const minQty  = row[1] ?? 1;
            const maxQty  = row[2] ?? 1;
            const chance  = row[5] ?? 0;
            const qty     = maxQty > minQty ? `×${minQty}–${maxQty}` : `×${minQty}`;
            const pct     = chance > 0 ? `${(chance * 100).toFixed(0)}%` : '—';
            return `<div class="dt-preview-item">` +
                `<span class="dt-preview-item-name" title="${name}">${name}</span>` +
                `<span class="dt-preview-item-qty">${qty}</span>` +
                `<span class="dt-preview-item-pct">${pct}</span>` +
                `</div>`;
        }).join('');
        preview.innerHTML =
            `<div class="dt-preview-title">${dt.name || '(unnamed)'}</div>` +
            `<div class="dt-preview-meta">id:${dt.id} · ${dt.items.length} item${dt.items.length !== 1 ? 's' : ''} · mdlType:${dt.mdlType ?? 0}</div>` +
            rows;
    }

    function renderList(q) {
        q = q.trim().toLowerCase();
        const noneHtml = `<div class="dt-picker-item${currentId < 0 ? ' dt-active' : ''}" data-id="-1">` +
            `<span class="dt-picker-item-name" style="color:#999">None (no drops)</span></div>`;
        const rows = [..._dropsTablesMap.values()]
            .filter(dt => !q || dt.name?.toLowerCase().includes(q) || String(dt.id).includes(q))
            .sort((a, b) => a.id - b.id)
            .map(dt =>
                `<div class="dt-picker-item${dt.id === currentId ? ' dt-active' : ''}" data-id="${dt.id}">` +
                `<span class="dt-picker-item-name">${dt.name || '(unnamed)'}</span>` +
                `<span class="dt-picker-item-meta">id:${dt.id} · ${dt.items?.length ?? 0} items</span>` +
                `</div>`
            );
        list.innerHTML = noneHtml + (rows.length ? rows.join('') : '<div class="np-empty">No tables match.</div>');
        list.querySelectorAll('.dt-picker-item').forEach(item => {
            item.addEventListener('mouseenter', () => {
                const id = parseInt(item.dataset.id);
                renderPreview(id >= 0 ? _dropsTablesMap.get(id) : null);
            });
            item.addEventListener('click', () => {
                const id = parseInt(item.dataset.id);
                if (hiddenInput) {
                    hiddenInput.value = id;
                    // Dispatch change event so watchEditChanges enables the Apply button
                    hiddenInput.dispatchEvent(new Event('change', { bubbles: true }));
                    const dt = id >= 0 ? _dropsTablesMap.get(id) : null;
                    const label = dt ? dt.name : 'None';
                    const idBadge = id >= 0 ? ` (id:${id}, ${dt?.items?.length ?? 0} items)` : '';
                    const labelEl = popupSection?.querySelector('.se-drops-label');
                    if (labelEl) { labelEl.textContent = label; labelEl.title = label + idBadge; }
                    // Show/hide Edit button
                    const editBtn = popupSection?.querySelector('.se-drops-edit-btn');
                    if (editBtn) { editBtn.dataset.dt = id; editBtn.style.display = id >= 0 ? '' : 'none'; }
                    else if (id >= 0) {
                        const changeBtn = popupSection?.querySelector('.se-drops-picker-btn');
                        if (changeBtn) {
                            const newEdit = document.createElement('button');
                            newEdit.className = 'popup-edit-btn se-drops-edit-btn';
                            newEdit.dataset.dt = id;
                            newEdit.style.cssText = 'font-size:10px;padding:1px 6px';
                            newEdit.textContent = 'Edit';
                            changeBtn.after(newEdit);
                        }
                    }
                    // Update chips area
                    _updateDropsChips(popupSection, dt);
                }
                modal.classList.remove('open');
            });
        });
    }

    if (modal._abortCtrl) modal._abortCtrl.abort();
    modal._abortCtrl = new AbortController();
    const sig = modal._abortCtrl.signal;
    search.addEventListener('input', () => renderList(search.value), { signal: sig });
    document.getElementById('dt-picker-none').addEventListener('click', () => {
        if (hiddenInput) {
            hiddenInput.value = -1;
            hiddenInput.dispatchEvent(new Event('change', { bubbles: true }));
        }
        const labelEl = popupSection?.querySelector('.se-drops-label');
        if (labelEl) { labelEl.textContent = 'None'; labelEl.title = 'None'; }
        const editBtn = popupSection?.querySelector('.se-drops-edit-btn');
        if (editBtn) editBtn.style.display = 'none';
        _updateDropsChips(popupSection, null);
        modal.classList.remove('open');
    }, { signal: sig });
    document.getElementById('dt-picker-new').addEventListener('click', () => {
        modal.classList.remove('open');
        const newDt = createDropTable();
        if (_renderEditPanel) _renderEditPanel();
        openDropTableEditor(newDt.id);
    }, { signal: sig });
    document.getElementById('dt-picker-cancel').addEventListener('click',
        () => modal.classList.remove('open'), { signal: sig });

    search.value = '';
    renderList('');
    renderPreview(currentId >= 0 ? _dropsTablesMap.get(currentId) : null);
    modal._popupSection = popupSection;
    modal.classList.add('open');
    setTimeout(() => search.focus(), 50);
}

function openDropTableEditor(tableId) {
    const modal  = document.getElementById('dt-editor-modal');
    const metaEl = document.getElementById('dt-editor-meta');
    const tbody  = document.getElementById('dt-editor-tbody');
    const title  = document.getElementById('dt-editor-title');
    if (!modal) return;

    const dt = _dropsTablesMap.get(tableId);
    if (!dt) return;

    title.textContent = `Edit Drop Table — id:${dt.id}`;
    metaEl.innerHTML =
        `<label>Name<input id="dt-ed-name" type="text" value="${dt.name ?? ''}" style="width:200px"></label>` +
        `<label>mdlType<input id="dt-ed-mdltype" type="number" value="${dt.mdlType ?? 0}" style="width:60px" min="0"></label>`;

    function rebuildRows() {
        tbody.innerHTML = dt.items.length
            ? dt.items.map((row, i) => _dtItemRow(row, i)).join('')
            : `<tr><td colspan="10" class="dt-drop-hint">Search for items above, or click + Add Blank</td></tr>`;
        // Live item name lookup
        tbody.querySelectorAll('input[data-col="0"]').forEach((inp, i) => {
            inp.addEventListener('input', () => {
                const hint = tbody.querySelector(`[data-name-for="${i}"]`);
                if (hint) hint.textContent = itemNames[inp.value]?.name ?? '';
            });
        });
        // Delete row
        tbody.querySelectorAll('.dt-del-row').forEach(btn => {
            btn.addEventListener('click', () => {
                const idx = parseInt(btn.dataset.row);
                dt.items.splice(idx, 1);
                rebuildRows();
            });
        });
        // Row reorder by drag
        if (_attachDragReorder) _attachDragReorder(tbody, dt.items, rebuildRows);
    }
    rebuildRows();

    function readAndSave() {
        const name    = document.getElementById('dt-ed-name')?.value.trim() || 'Unnamed';
        const mdlType = parseInt(document.getElementById('dt-ed-mdltype')?.value) || 0;
        const newItems = Array.from(tbody.querySelectorAll('tr[data-idx]')).map(row => {
            const cols = row.querySelectorAll('[data-col]');
            const item = [0, 1, 1, 0, false, 0];
            cols.forEach(f => {
                const col = parseInt(f.dataset.col);
                if (col === 4) item[col] = f.checked;
                else if (col === 5) item[col] = parseFloat(f.value) / 100 || 0;
                else item[col] = parseInt(f.value) || 0;
            });
            return item;
        });
        dt.name    = name;
        dt.mdlType = mdlType;
        dt.items   = newItems;
        if (_markDirty) _markDirty('ddon-src-spawns');
        if (_renderEditPanel) _renderEditPanel();
        if (_rebuildOpenPopup) _rebuildOpenPopup();
        modal.classList.remove('open');
    }
    _dtEditorReadAndSave = readAndSave;

    if (modal._abortCtrl) modal._abortCtrl.abort();
    modal._abortCtrl = new AbortController();
    const sig = modal._abortCtrl.signal;

    // Item search — find and click to add
    const searchInput   = document.getElementById('dt-item-search-input');
    const searchResults = document.getElementById('dt-item-search-results');
    if (searchInput)  searchInput.value = '';
    if (searchResults) searchResults.innerHTML = '';
    function renderItemSearch(q) {
        if (!searchResults) return;
        if (!q) { searchResults.innerHTML = ''; return; }
        const matches = Object.entries(itemNames)
            .filter(([id, e]) => e.name.toLowerCase().includes(q.toLowerCase()) || id.includes(q))
            .slice(0, 10);
        searchResults.innerHTML = matches.length
            ? matches.map(([id, e]) => {
                const iconNo   = e.iconNo;
                const iconFile = iconNo != null ? `ii${String(iconNo).padStart(6,'0')}.png` : null;
                const icon     = iconFile && _iconIdSet.has(iconNo)
                    ? `<img src="images/icons/small/${iconFile}" width="16" height="16" style="image-rendering:pixelated;flex-shrink:0">`
                    : `<span style="width:16px;flex-shrink:0"></span>`;
                return `<div class="dt-item-result" data-id="${id}">${icon}<span class="dt-item-result-name">${e.name}</span><span class="dt-item-result-id">#${id}</span></div>`;
              }).join('')
            : `<div style="color:#555;font-size:0.75rem;padding:4px 6px">No matches</div>`;
        searchResults.querySelectorAll('.dt-item-result[data-id]').forEach(el => {
            el.addEventListener('click', () => {
                dt.items.push([parseInt(el.dataset.id), 1, 1, 0, false, 1.0]);
                rebuildRows();
                searchInput.value = '';
                searchResults.innerHTML = '';
                searchInput.focus();
            });
        });
    }
    searchInput?.addEventListener('input', () => renderItemSearch(searchInput.value.trim()), { signal: sig });

    // Drag from Items panel → append new row
    tbody.addEventListener('dragover', (e) => {
        if (_dragItemId == null) return;
        e.preventDefault();
        e.dataTransfer.dropEffect = 'copy';
        tbody.classList.add('dt-item-drop-hover');
    }, { signal: sig });
    tbody.addEventListener('dragleave', (e) => {
        if (!tbody.contains(e.relatedTarget)) tbody.classList.remove('dt-item-drop-hover');
    }, { signal: sig });
    tbody.addEventListener('drop', (e) => {
        if (_dragItemId == null) return;
        e.preventDefault();
        tbody.classList.remove('dt-item-drop-hover');
        dt.items.push([_dragItemId, 1, 1, 0, false, 1.0]);
        rebuildRows();
    }, { signal: sig });

    document.getElementById('dt-editor-save').addEventListener('click', readAndSave, { signal: sig });
    document.getElementById('dt-editor-add-row').addEventListener('click', () => {
        dt.items.push([0, 1, 1, 0, false, 1.0]);
        rebuildRows();
    }, { signal: sig });
    document.getElementById('dt-editor-delete').addEventListener('click', () => {
        if (!confirm(`Delete drop table "${dt.name}" (id:${dt.id})?\nAny spawns using this table will lose their drops.`)) return;
        _dropsTablesMap.delete(dt.id);
        if (_rawEnemyData?.dropsTables) {
            const idx = _rawEnemyData.dropsTables.findIndex(t => t.id === dt.id);
            if (idx >= 0) _rawEnemyData.dropsTables.splice(idx, 1);
        }
        if (_markDirty) _markDirty('ddon-src-spawns');
        if (_renderEditPanel) _renderEditPanel();
        modal.classList.remove('open');
    }, { signal: sig });
    // close button is wired statically below — no signal needed

    modal.classList.add('open');
}

// Static close + backdrop handlers for drop table modals
const _dtPickerModal = document.getElementById('dt-picker-modal');
const _dtEditorModal = document.getElementById('dt-editor-modal');
if (_dtPickerModal) {
    let _bdDown = false;
    _dtPickerModal.addEventListener('mousedown', e => { _bdDown = e.target === _dtPickerModal; });
    _dtPickerModal.addEventListener('click', e => { if (_bdDown && e.target === _dtPickerModal) _dtPickerModal.classList.remove('open'); _bdDown = false; });
}
if (_dtEditorModal) {
    let _bdDown = false;
    _dtEditorModal.addEventListener('mousedown', e => { _bdDown = e.target === _dtEditorModal; });
    _dtEditorModal.addEventListener('click', e => {
        if (_bdDown && e.target === _dtEditorModal) { if (_dtEditorReadAndSave) _dtEditorReadAndSave(); else _dtEditorModal.classList.remove('open'); }
        _bdDown = false;
    });
}
document.getElementById('dt-picker-close')?.addEventListener('click',
    () => document.getElementById('dt-picker-modal').classList.remove('open'));
document.getElementById('dt-editor-close')?.addEventListener('click',
    () => { if (_dtEditorReadAndSave) _dtEditorReadAndSave(); else document.getElementById('dt-editor-modal').classList.remove('open'); });

// ── Unsaved-changes guard ─────────────────────────────────────────────────────
window.addEventListener('beforeunload', e => {
    if (_dirtySet.size > 0) {
        e.preventDefault();
        e.returnValue = ''; // required for Chrome to show the dialog
    }
});

// ── Init ──────────────────────────────────────────────────────────────────────
ensureInitialHash();

async function bootstrapMapApp() {
    initMobTypeFilters();
    initPoiFilters();
    initDevPanel();
    updateSidebarPresetLabel();
    initRisingChannelToggle();
    try {
        buildSidebar();
    } catch (err) {
        console.error('buildSidebar failed:', err);
        const listEl = document.getElementById('map-list');
        if (listEl) {
            listEl.innerHTML = '<div style="padding:8px 16px;color:#e94560;font-size:0.8rem">Map list failed to load — see console.</div>';
        }
    }
    await waitForMapContainer();
    await nextFrame();
    leafletMap.invalidateSize({ animate: false });
    const sz = leafletMap.getSize();
    if (sz.x === 0 || sz.y === 0) {
        await waitForMapContainer();
        await nextFrame();
        leafletMap.invalidateSize({ animate: false });
    }
    await loadMap(currentMapName());
    watchMapContainerSize();
    checkLocalSources();
    refitActiveMapView();
}

bootstrapMapApp();
window.addEventListener('resize', () => leafletMap.invalidateSize({ animate: false }));

// ── Developer panel ───────────────────────────────────────────────────────────
function initDevPanel() {
    const devPanel    = document.getElementById('dev-panel');
    const panelToggle = document.getElementById('dev-panel-toggle');
    const panelBody   = document.getElementById('dev-panel-body');
    if (!devPanel || !panelToggle || !panelBody) return;

    const devPrefs = loadDevPrefs();
    _devMobSpawnLabels = readMobSpawnLabelsPref(devPrefs);

    const mobLabelsEl = document.getElementById('dev-mob-spawn-labels');
    if (mobLabelsEl) mobLabelsEl.checked = _devMobSpawnLabels;
    applyDevDisplayPrefs();

    mobLabelsEl?.addEventListener('change', () => {
        _devMobSpawnLabels = mobLabelsEl.checked;
        saveDevPrefs();
        if (_devMobSpawnLabels) _collapseAllGroups();
        else removeAllGroupTerritoryRects();
        applyMobDisplayMode();
        syncTerritoryLayer();
        refreshMobTooltips();
        refreshGroupChipIcons();
        saveLayerPrefs();
    });

    const setPanelOpen = (open) => {
        devPanel.classList.toggle('expanded', open);
        panelToggle.setAttribute('aria-expanded', open ? 'true' : 'false');
        try { localStorage.setItem(DEV_PANEL_KEY, open ? '1' : '0'); } catch { /* ignore */ }
        if (open) devPanel.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    };
    setPanelOpen(localStorage.getItem(DEV_PANEL_KEY) === '1');
    panelToggle.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        setPanelOpen(!devPanel.classList.contains('expanded'));
    });
}
