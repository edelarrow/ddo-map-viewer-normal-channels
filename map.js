import enemyPositions from './resources/enemyPositions.json' with {type: "json"};
import mapParams from './resources/map_params.json' with {type: "json"};
import landmarkData from './resources/landmarks.json' with {type: "json"};
import connectionData from './resources/connections.json' with {type: "json"};
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
import breakTargets   from './resources/breakTargets.json'   with {type: "json"};
import stageGroups   from './resources/stageGroups.json'   with {type: "json"};
import worldFlags      from './resources/worldFlags.json'      with {type: "json"};
import worldFlagsExtra from './resources/worldFlagsExtra.json' with {type: "json"};
import worldQuestFlags from './resources/worldQuestFlags.json' with {type: "json"};
import emRadii        from './resources/emRadii.json'        with {type: "json"};
const _iconIdSet = new Set(iconIds);
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
let gatherLayer       = L.layerGroup();   // off by default
const _gatherMarkerByKey = new Map();    // "${stageNo}:${groupId}:${posId}" → L.marker
const _shopMarkerByNpcId = new Map();    // "${stageNo}:${npcId}" → L.marker
let npcShopLayer        = L.layerGroup();   // off by default
let specialShopLayer    = L.layerGroup();   // off by default
let breakTargetLayer  = L.layerGroup();   // off by default
let pdBoundaryLayer = L.layerGroup().addTo(leafletMap);
let spawnRadiiLayer   = L.layerGroup().addTo(leafletMap);  // aggro/link radius circles
let _spreadOverlay    = L.layerGroup().addTo(leafletMap);  // cross-group spoke lines + anchor dots

// Canvas renderer — all spawn circleMarkers share one <canvas> element (huge perf win).
const spawnRenderer = L.canvas({ padding: 0.5 });

// ── Group expand/collapse state ───────────────────────────────────────────────
// One entry per group; detailsLayer is lazily created on first expand.
const _groupStore = new Map(); // groupId string → { groupId, color, territory, items, pts,
                               //   centroid, labelMarker, detailsLayer, isExpanded }
let _currentMapInfo  = null;   // stored at loadEnemySpawns time; used by lazy expand


// ── Edit mode is permanently disabled in this viewer build ──────────────────
const _editMode = false;
let _activeSubGroupId   = null;
let _availableSubGroups = [0];
let _currentFloorObbs = null;
let _dropsTablesMap    = new Map();

function updateEnemyVisibility() {
    const checked = document.getElementById('layer-enemies').checked;
    if (checked) {
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

// Returns the !-suffix string: layer flags + optional ;groupId,groupId,...
// Format: !elcgt;0,3,80
function getLayersHash() {
    let s = '';
    if (document.getElementById('layer-enemies').checked)       s += 'e';
    if (document.getElementById('layer-landmarks').checked)     s += 'l';
    if (document.getElementById('layer-connections').checked)   s += 'c';
    if (document.getElementById('layer-grid').checked)          s += 'g';
    if (document.getElementById('layer-stage-labels').checked)  s += 'a';
    if (document.getElementById('layer-gather').checked)        s += 'r';
    if (document.getElementById('layer-radii').checked)         s += 'i';
    if (document.getElementById('layer-shops').checked)               s += 'n';
    if (document.getElementById('layer-break-targets').checked)       s += 'b';
    if (document.getElementById('sidebar').classList.contains('collapsed')) s += 's';
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
        enemies:      document.getElementById('layer-enemies').checked,
        landmarks:    document.getElementById('layer-landmarks').checked,
        connections:  document.getElementById('layer-connections').checked,
        grid:         document.getElementById('layer-grid').checked,
        stageLabels:  document.getElementById('layer-stage-labels').checked,
        gather:       document.getElementById('layer-gather').checked,
        radii:        document.getElementById('layer-radii').checked,
        shops:         document.getElementById('layer-shops').checked,
        breakTargets:  document.getElementById('layer-break-targets').checked,
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

    document.getElementById('layer-enemies').checked       = isOn('enemies',      true);
    document.getElementById('layer-landmarks').checked     = isOn('landmarks',    true);
    document.getElementById('layer-connections').checked   = isOn('connections',  true);
    document.getElementById('layer-grid').checked          = isOn('grid',         false);
    document.getElementById('layer-stage-labels').checked  = isOn('stageLabels',  true);
    document.getElementById('layer-gather').checked        = isOn('gather',        false);
    document.getElementById('layer-radii').checked         = isOn('radii',         false);
    document.getElementById('layer-shops').checked          = isOn('shops', false) || isOn('npcShops', true);
    document.getElementById('layer-break-targets').checked  = isOn('breakTargets',  false);

    if (!document.getElementById('layer-landmarks').checked)
        leafletMap.removeLayer(landmarkLayer);
    if (!document.getElementById('layer-connections').checked)
        leafletMap.removeLayer(connectionLayer);
    if (document.getElementById('layer-grid').checked)
        leafletMap.addLayer(gridLayer);
    if (!document.getElementById('layer-stage-labels').checked)
        leafletMap.removeLayer(stageLabelsLayer);
    if (document.getElementById('layer-gather').checked)
        leafletMap.addLayer(gatherLayer);
    if (document.getElementById('layer-shops').checked) {
        leafletMap.addLayer(npcShopLayer);
        leafletMap.addLayer(specialShopLayer);
    }
    if (document.getElementById('layer-break-targets').checked)
        leafletMap.addLayer(breakTargetLayer);
    if (!document.getElementById('layer-enemies').checked)
        updateEnemyVisibility();
    if (isOn('sidebarHidden', false))
        document.getElementById('sidebar').classList.add('collapsed');
})();

// ── Layer toggles ──────────────────────────────────────────────────────────────
document.getElementById('layer-enemies').addEventListener('change', () => {
    updateEnemyVisibility(); saveLayerPrefs();
});
document.getElementById('layer-landmarks').addEventListener('change', e => {
    e.target.checked ? leafletMap.addLayer(landmarkLayer) : leafletMap.removeLayer(landmarkLayer);
    saveLayerPrefs();
});
document.getElementById('layer-connections').addEventListener('change', e => {
    e.target.checked ? leafletMap.addLayer(connectionLayer) : leafletMap.removeLayer(connectionLayer);
    saveLayerPrefs();
});
document.getElementById('layer-grid').addEventListener('change', e => {
    e.target.checked ? leafletMap.addLayer(gridLayer) : leafletMap.removeLayer(gridLayer);
    saveLayerPrefs();
});
document.getElementById('layer-stage-labels').addEventListener('change', e => {
    e.target.checked ? leafletMap.addLayer(stageLabelsLayer) : leafletMap.removeLayer(stageLabelsLayer);
    saveLayerPrefs();
});
document.getElementById('layer-gather').addEventListener('change', e => {
    e.target.checked ? leafletMap.addLayer(gatherLayer) : leafletMap.removeLayer(gatherLayer);
    saveLayerPrefs();
});
document.getElementById('layer-radii').addEventListener('change', e => {
    if (!e.target.checked) clearSpawnRadii();
    saveLayerPrefs();
});
document.getElementById('layer-shops').addEventListener('change', e => {
    if (e.target.checked) { leafletMap.addLayer(npcShopLayer); leafletMap.addLayer(specialShopLayer); }
    else { leafletMap.removeLayer(npcShopLayer); leafletMap.removeLayer(specialShopLayer); }
    saveLayerPrefs();
});
document.getElementById('layer-break-targets').addEventListener('change', e => {
    e.target.checked ? leafletMap.addLayer(breakTargetLayer) : leafletMap.removeLayer(breakTargetLayer);
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

document.getElementById('btn-expand-collapse').addEventListener('click', () => {
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

function appendGroupHeader(listEl, text) {
    const header = document.createElement('div');
    header.className = 'map-group-header';
    header.textContent = text;
    listEl.appendChild(header);
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
        const byLabel = new Map();
        for (const e of entries) {
            if (!byLabel.has(e.label)) byLabel.set(e.label, []);
            byLabel.get(e.label).push(e);
        }
        for (const [label, group] of byLabel) {
            if (group.length === 1) {
                const e = group[0];
                appendMapEntry(listEl, e.name, e.info, label, e.stid, currentMap, currentStage);
            } else {
                appendCollapsibleGroup(listEl, label, group, currentMap, currentStage);
            }
        }
        return;
    }

    // No search: group by quest area (sorted by quest_area_id), then alphabetically within
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

    for (const [, area] of sortedAreas) {
        area.entries.sort((a, b) => a.label.localeCompare(b.label));
        appendGroupHeader(listEl, area.name);

        // Collapse entries that share the same display label into one expandable row.
        // This handles: same map model with multiple stages AND different models with identical names.
        const byLabel = new Map();
        for (const e of area.entries) {
            if (!byLabel.has(e.label)) byLabel.set(e.label, []);
            byLabel.get(e.label).push(e);
        }

        for (const [label, group] of byLabel) {
            if (group.length === 1) {
                const e = group[0];
                appendMapEntry(listEl, e.name, e.info, label, e.stid, currentMap, currentStage);
            } else {
                appendCollapsibleGroup(listEl, label, group, currentMap, currentStage);
            }
        }
    }
}

const _mapSearchInput = document.getElementById('map-search');
const _mapSearchClear = document.getElementById('map-search-clear');
_mapSearchInput.addEventListener('input', e => {
    _mapSearchClear.style.display = e.target.value ? 'block' : 'none';
    buildSidebar(e.target.value);
});
_mapSearchClear.addEventListener('click', () => {
    _mapSearchInput.value = '';
    _mapSearchClear.style.display = 'none';
    _mapSearchInput.focus();
    buildSidebar('');
});

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
            stageLabels:  flagStr.includes('a'),
            gather:       flagStr.includes('r'),
            radii:         flagStr.includes('i'),
            shops:         flagStr.includes('n') || flagStr.includes('p'),
            breakTargets:  flagStr.includes('b'),
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
    return L.divIcon({
        className: '',
        html: `<div class="group-chip${expanded ? ' chip-open' : ''}" style="color:${chipColor};${shadowStyle}"${titleAttr ? ` title="${titleAttr}"` : ''}><span class="chip-arrow${expanded ? ' open' : ''}">&#9654;</span>G${groupId} <span class="chip-count">${count}</span></div>`,
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

// Build the details layer (hull + territory + spawn dots) for a group entry.
// Does NOT add the layer to the map — that is done by expandGroup.
function buildGroupDetails(g) {
    const info  = _currentMapInfo;
    const layer = L.layerGroup();

    // Hull
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
            layer.addLayer(poly);
            g.hullPts = hull;  // stored for map-level middle-click collapse
        }
    } else if (g.pts.length === 2) {
        L.polyline(g.pts.map(([px, py]) => xy(px, py)), {
            color: g.color, weight: 1.5, opacity: 0.65, dashArray: '4 3', interactive: false,
        }).addTo(layer);
    }

    // Territory rectangle — stored separately so it respects the territory layer toggle
    g.territoryRect = null;
    if (g.territory) {
        const { xMin, xMax, zMin, zMax } = g.territory;
        const sw = worldToPixel(xMin, zMin, info);
        const ne = worldToPixel(xMax, zMax, info);
        g.territoryRect = L.rectangle([sw, ne], {
            color:       g.color,
            weight:      2,
            opacity:     0.85,
            fillColor:   g.color,
            fillOpacity: 0.08,
            dashArray:   '8 4',
            interactive: false,
        });
    }

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

        const spawnTimeLabel = (t) => {
            if (!t || t === '00:00,23:59') return '';
            if (t.startsWith('07:')) return '☀ Day';
            if (t.startsWith('18:')) return '🌙 Night';
            return t;
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
                    const href     = `https://reference.dd-on.com/build/i${String(itemId).padStart(8, '0')}.html`;
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

            const editSection = '';
            return `${badge}<br>${groupLabel}, Index: <b>${idx}</b>${subLine}${triggerLine}${cycleHtml}${emLine}${keyLine}${buildBossLine(spawnInfo)}${buildManualSetLine(spawnInfo)}${radiiLine}${orbsLine}${_editMode ? '' : buildDropsHtml(spawnInfo)}${editSection}`;
        };

        const buildTooltip = (spawnCache) => {
            const entries  = spawnKey && spawnCache ? (spawnCache.get(spawnKey) ?? []) : [];
            const hasEnemy = !spawnCache || entries.some(e => !!e.lv);
            // Helper: resolve display name with named param applied
            const infectionPrefix = [null, 'Infected', 'Severely Infected', 'War-Ready'];
            const resolveDisplayName = (e) => {
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
                const prefix = e.infection ? infectionPrefix[e.infection] : null;
                return prefix ? `${prefix} ${name}` : name;
            };
            // Returns a small styled "(OriginalName)" suffix when namedId is REPLACE type
            const replaceOriginSuffix = (e) => {
                if (!e?.namedId) return '';
                const np = namedParamsById.get(e.namedId);
                if (np?.type !== 'NAMED_TYPE_REPLACE') return '';
                const base = e.emCode ? (emNames[e.emCode]?.name ?? null) : null;
                if (!base) return '';
                return ` <span style="font-size:10px;color:#aaa;font-style:italic">(${base})</span>`;
            };
            let namePart = '';
            if (entries.length > 1) {
                // Show all variants: "Killer Bee Lv3☀ / Skeleton Lv3🌙"
                const parts = entries
                    .filter(e => !!e.lv)
                    .map(e => {
                        const n = resolveDisplayName(e);
                        const t = spawnTimeLabel(e.spawnTime);
                        return n ? `${n} Lv${e.lv}${t ? [...t][0] : ''}${replaceOriginSuffix(e)}` : null;
                    })
                    .filter(Boolean);
                if (parts.length) namePart = parts.join(' / ') + ' — ';
            } else if (entries.length === 1 && hasEnemy) {
                const e0 = entries[0];
                const n  = resolveDisplayName(e0) ?? (spawn.EmName ? (emNames[spawn.EmName]?.name ?? null) : null);
                if (n) namePart = `${n}${e0.lv ? ` Lv${e0.lv}` : ''}${replaceOriginSuffix(e0)} — `;
            } else if (!spawnCache && hasEnemy && spawn.EmName) {
                const n = emNames[spawn.EmName]?.name ?? null;
                if (n) namePart = `${n} — `;
            }
            const e0 = entries[0] ?? null;
            const orbBadge  = (e0?.isBloodOrbEnemy && e0?.bloodOrbs ? ' 🩸' : '') + (e0?.isHighOrbEnemy && e0?.highOrbs ? ' ⭐' : '');
            const manualBadge = e0?.isManualSet ? ' 😴' : '';
            const bossBadge = (e0?.isBossGauge || e0?.isAreaBoss || e0?.raidBossId > 0) ? ' <span style="color:#ff4444" title="Boss enemy">☠</span>' : '';
            return `${namePart}${g.groupId}.${idx} [SS:${sg}]${orbBadge}${manualBadge}${bossBadge}${isKeyBearer ? ' <span style="color:#c8a000;font-size:16px;">🗝</span>' : ''}`;
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

        // Rebuild popup and wire cycle-button handlers on open.
        let _popupClickHandler = null;
        marker.on('popupopen', function() {
            const popup = this.getPopup();
            const bind = (cache) => {
                requestAnimationFrame(() => {
                    const el = popup.getElement();
                    if (!el) return;
                    const contentDiv = el.querySelector('.leaflet-popup-content');
                    if (contentDiv) {
                        contentDiv.innerHTML = buildEnemyPopup(cache);
                        popup._updateLayout?.();
                        popup._updatePosition?.();
                    }
                    if (_popupClickHandler) el.removeEventListener('click', _popupClickHandler);
                    const getEntries = () => cache?.get(spawnKey) ?? [];
                    const rebuildPopup = () => {
                        const cd = el.querySelector('.leaflet-popup-content');
                        if (!cd) return;
                        cd.innerHTML = buildEnemyPopup(cache);
                        popup._updateLayout?.();
                        popup._updatePosition?.();
                    };
                    _popupClickHandler = (e) => {
                        // ── Spawn-set navigation (jump to sibling marker) ──────
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
                        // ── Cycle prev/next day/night variants ─────────────────
                        const cycleBtn = e.target.closest('.spawn-prev, .spawn-next');
                        if (cycleBtn) {
                            const entries = getEntries();
                            if (entries.length <= 1) return;
                            e.stopPropagation();
                            displayIdx = (displayIdx + (cycleBtn.classList.contains('spawn-prev') ? -1 : 1) + entries.length) % entries.length;
                            rebuildPopup();
                            return;
                        }
                    };
                    el.addEventListener('click', _popupClickHandler);

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
function _expandGroupCore(g) {
    if (!g.detailsLayer) buildGroupDetails(g);
    const enemiesOn = document.getElementById('layer-enemies').checked;
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
    applySubGroupFilter();  // after setIcon so opacity isn't reset by icon replacement
}

function _collapseGroupCore(g) {
    if (g.detailsLayer) leafletMap.removeLayer(g.detailsLayer);
    if (g.territoryRect) territoryLayer.removeLayer(g.territoryRect);
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
    const sgSet = new Set([0]);
    for (const g of _groupStore.values()) {
        if (g.areaSpawn) sgSet.add(1);
        for (const { spawn } of g.items) sgSet.add(_spawnSubGroupId(spawn));
    }
    _availableSubGroups = [...sgSet].sort((a, b) => a - b);
    _renderSubGroupSelector();
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
    // Chip (label marker)
    g.labelMarker.setOpacity(visible ? 1 : 0);
    const chipEl = g.labelMarker.getElement();
    if (chipEl) chipEl.style.pointerEvents = visible ? '' : 'none';
    // Hull and other structure layers in an expanded detailsLayer
    if (g.detailsLayer) {
        for (const layer of g.detailsLayer.getLayers()) {
            if (layer._spawn) continue;  // skip spawn markers — handled separately
            if (visible) {
                layer.setStyle(layer._origStyle ?? {});
            } else {
                if (!layer._origStyle) layer._origStyle = { opacity: layer.options.opacity ?? 0.75, fillOpacity: layer.options.fillOpacity ?? 0.1 };
                layer.setStyle({ opacity: 0, fillOpacity: 0 });
            }
        }
    }
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
            const visible = _activeSubGroupId === null ||
                _spawnSubGroupId(m._spawn) === _activeSubGroupId ||
                (_activeSubGroupId === 1 && spawnIsAreaSpawnInitial);
            _setMarkerVisible(m, visible);
        }
    }
    reapplySpread();
}

function _renderSubGroupSelector() {
    const bar = document.getElementById('subgroup-bar');
    if (!bar) return;
    if (_availableSubGroups.length <= 1) { bar.style.display = 'none'; return; }
    bar.style.display = 'flex';
    const pill = (sg, label) => {
        const on = sg === _activeSubGroupId;
        return `<button class="sg-filter-btn" data-sg="${sg === null ? '' : sg}" style="font-size:10px;padding:1px 8px;border-radius:10px;cursor:pointer;border:1px solid;${on ? 'background:#4a90d9;color:#fff;border-color:#357abd' : 'background:#2a3a5a;color:#9ab;border-color:#3a5a7a'}">${label}</button>`;
    };
    bar.innerHTML = `<span style="font-size:9px;color:#778;text-transform:uppercase;letter-spacing:0.4px;margin-right:4px;align-self:center">SubGroup</span>` +
        pill(null, 'All') + _availableSubGroups.map(sg => pill(sg, sg)).join('');
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

function _clearHighlight() {
    for (const m of _highlightedSet) {
        m.setStyle(m._origStyle);
        m.setRadius(5);
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
    _applyHighlight(_sgMarkers[sgKey] || []);
}

function _unhighlightSG() {
    _unhighlightTimer = setTimeout(_clearHighlight, 160);
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


function loadEnemySpawns(info, stid = null) {
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
    _renderSubGroupSelector();
    _currentMapInfo   = info;
    _currentFloorObbs = info.floor_obbs ?? null;
    if (!info.stages?.length) return;

    const floorObbs     = _currentFloorObbs;
    const filterByFloor = floorObbs !== null;

    const stagesToLoad = (stid && info.stages.includes(stid)) ? [stid] : info.stages;

    // Collect all groups, merging across stages if multiple are loaded
    const byGroupId = new Map(); // groupId string → { territory, items:[{spawn,idx,sg,latlng}], pts:[] }
    for (const stageId of stagesToLoad) {
        const stageNo   = String(parseInt(stageId.slice(2), 10));
        const stageData = enemyPositions[stageNo];
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


// ── Live server data (fetched at runtime) ─────────────────────────────────────
const _DEFAULT_GATHERING_URL = 'https://raw.githubusercontent.com/edelarrow/map-spawns/refs/heads/main/Normal%20Channels/GatheringItem.csv';
const _DEFAULT_SPAWNS_URL    = 'https://raw.githubusercontent.com/edelarrow/map-spawns/refs/heads/main/Normal%20Channels/EnemySpawn.json';
const _DEFAULT_SHOP_URL      = 'https://raw.githubusercontent.com/edelarrow/map-spawns/refs/heads/main/Normal%20Channels/Shop.json';
const _DEFAULT_SPECIAL_SHOP_URL = 'https://raw.githubusercontent.com/edelarrow/map-spawns/refs/heads/main/Normal%20Channels/SpecialShops.json';
const _DEFAULT_SPECIAL_SHOP_URL = 'https://raw.githubusercontent.com/edelarrow/map-spawns/refs/heads/main/Normal%20Channels/SpecialShops.json';

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
function showSrcError(label) {
    const box = _srcErrorBox();
    const item = document.createElement('div');
    item.style.cssText = 'display:flex;align-items:center;gap:5px;margin-bottom:3px;'
        + 'color:#f99;background:#2a0f0f;border-left:3px solid #c0392b;border-radius:2px;padding:3px 6px;';
    item.innerHTML = `⚠ <span style="flex:1"><b>${label}</b> failed to load</span>`
        + `<button data-action="dismiss" style="font-size:0.7rem;padding:1px 4px;cursor:pointer;`
        + `background:none;color:#aaa;border:none">✕</button>`;
    item.querySelector('[data-action="dismiss"]').addEventListener('click', () => item.remove());
    box.appendChild(item);
}

// Cached promises — fetch starts once, result shared by all callers.
// Map key: "stageId,groupId,posId" → [{itemId, itemNum, maxItemNum, quality, dropChance, isHidden}]
let _gatherItemsCache = null;
const _gatherItemsPromise = fetch(_DEFAULT_GATHERING_URL)
    .then(r => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.text(); })
    .then(text => {
        const lines = text.split('\n');
        // First column header is "#StageId" — strip the leading '#'
        lines[0] = lines[0].replace(/^#/, '');
        const result = new Map();
        const headers = lines[0].split(',');
        const idx = name => headers.indexOf(name);
        const iStage = idx('StageId'), iGroup = idx('GroupId'),
              iPos = idx('PosId'), iItem = idx('ItemId'), iNum = idx('ItemNum'),
              iMax = idx('MaxItemNum'), iQual = idx('Quality'),
              iHidden = idx('IsHidden'), iChance = idx('DropChance');
        for (let i = 1; i < lines.length; i++) {
            const cols = lines[i].split(',');
            if (cols.length < 5) continue;
            const row = {
                stageId:    cols[iStage],
                groupId:    cols[iGroup],
                posId:      cols[iPos],
                itemId:     parseInt(cols[iItem]),
                itemNum:    parseInt(cols[iNum]),
                maxItemNum: parseInt(cols[iMax]),
                quality:    parseInt(cols[iQual]),
                isHidden:   cols[iHidden] === 'true' || cols[iHidden] === '1',
                dropChance: iChance >= 0 ? parseFloat(cols[iChance]) : 1,
            };
            const key = `${row.stageId},${row.groupId},${row.posId}`;
            if (!result.has(key)) result.set(key, []);
            result.get(key).push({
                itemId: row.itemId, itemNum: row.itemNum, maxItemNum: row.maxItemNum,
                quality: row.quality, isHidden: row.isHidden, dropChance: row.dropChance,
            });
        }
        _gatherItemsCache = result;
        return result;
    })
    .catch(() => { showSrcError('Gathering Items'); _gatherItemsCache = new Map(); return _gatherItemsCache; });
    .catch(() => { showSrcError('Gathering Items'); _gatherItemsCache = new Map(); return _gatherItemsCache; });

// Map key for enemy spawns: "stageId,groupId,posIdx" → [{emCode, lv, bloodOrbs, highOrbs, spawnTime, drops}]
let _enemySpawnCache = null;
const _enemySpawnPromise = fetch(_DEFAULT_SPAWNS_URL)
    .then(r => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.json(); })
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
              iIsHighOrbEnemy  = schemas.indexOf('IsHighOrbEnemy');
        // Store raw data and schema indices for write-back
            iStage, iGroup, iPosIdx, iEnemyId, iLv, iBlood, iHigh, iSpawnTime, iDrops,
            iScale, iSubGroup, iNamed, iRaidBoss, iSetType, iInfection,
            iIsBossG, iIsBossBGM, iIsAreaBoss, iExp, iRepopNum, iRepopCount,
            iTargetType, iHmPreset, iStartThink, iMontage, iIsManualSet, iPPDrop,
            iIsBloodOrbEnemy, iIsHighOrbEnemy,
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
            };
            if (result.has(key)) result.get(key).push(entry);
            else                 result.set(key, [entry]);
        }
        _enemySpawnCache = result;
        return result;
    })
    .catch(() => { showSrcError('Enemy Spawns'); _enemySpawnCache = new Map(); return _enemySpawnCache; });

// Map: ShopId → {walletType, items:[{itemId, price, stock}]}
let _shopCache = null;
const _shopPromise = fetch(_DEFAULT_SHOP_URL)
    .then(r => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.json(); })
    .then(data => {
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
const _specialShopPromise = fetch(_DEFAULT_SPECIAL_SHOP_URL)
    .then(r => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.json(); })
    .then(data => {
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
    CHEST_ROUND:           '#FFF59D',  // pale yellow — small round chest
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
    CHEST_ROUND:           'Small Round Chest',
    CHEST_SEALED_ORANGE:   'Sealed Chest (Orange)',
    CHEST_SEALED_PURPLE:   'Sealed Chest (Purple)',
    CHEST_SEALED_PEARL:    'Pearlescent Chest',
    CHEST_UNKNOWN:         'Chest',
};

function loadGatherPoints(info, stid = null) {
    gatherLayer.clearLayers();
    _gatherMarkerByKey.clear();
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
                    const displayItems = items.filter(it => !it.isHidden);
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
                                const href     = `https://reference.dd-on.com/build/i${String(it.itemId).padStart(8, '0')}.html`;
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
                        `<button class="popup-edit-btn" data-edit-action="ge-apply">✔ Apply</button>` +
                        `</div></div>`;
                }
                return `${badge}${typeLine}${coordLine}${itemsHtml}${editSection}`;
            };

            const tooltipText = `${label} — ${node.groupId}.${node.posId}`;

            const gatherIcon = L.divIcon({
                className: '',
                html: `<div style="width:9px;height:9px;background:${color};border:2px solid rgba(255,255,255,0.7);box-shadow:0 0 3px rgba(0,0,0,0.7);"></div>`,
                iconSize:    [9, 9],
                iconAnchor:  [4, 4],
                popupAnchor: [0, -8],
            });
            const marker = L.marker(latlng, { icon: gatherIcon })
            .bindPopup(buildGatherPopup(_gatherItemsCache))
            .bindTooltip(tooltipText, { permanent: false, direction: 'top', offset: [0, -8] })
            .addTo(gatherLayer);
            _gatherMarkerByKey.set(`${stageNo}:${node.groupId}:${node.posId}`, marker);

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
                            const items = gatherMap.get(csvKey);
                            if (!items) return;
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
                    if (_editMode && _attachDragReorder) {
                        const popupEl = m.getPopup().getElement();
                        const tbl = popupEl?.querySelector('.popup-edit-items-table');
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
                            const items = gatherMap.get(csvKey) ?? [];
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
                                rebuildGatherPopup(self, gatherMap);
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
                            const items = gatherMap.get(csvKey) ?? [];
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


}

function loadBreakTargets(info, stid = null) {
    breakTargetLayer.clearLayers();
    if (!info.stages?.length) return;

    const floorObbs     = info.floor_obbs ?? null;
    const filterByFloor = floorObbs !== null;
    const stagesToLoad  = (stid && info.stages.includes(stid)) ? [stid] : info.stages;

    for (const stageId of stagesToLoad) {
        const stageNo = String(parseInt(stageId.slice(2), 10));
        const nodes   = breakTargets[stageNo];
        if (!nodes) continue;

        for (const node of nodes) {
            if (filterByFloor) {
                const floor = getEnemyFloor(node.x, node.y, node.z, floorObbs);
                if (floor !== null && floor !== currentLayer) continue;
            }
            const latlng = worldToPixel(node.x, node.z, info);

            const questLine  = node.questName
                ? `<br><span style="color:#c97a00;font-size:10px;font-style:italic">${node.questName.replace(/\n/g, ' ')}</span>`
                : '';
            const hitsLine   = node.hitNum != null
                ? `<br><span style="font-size:10px;color:#888">${node.hitNum} hit${node.hitNum !== 1 ? 's' : ''} to destroy</span>`
                : '';
            const condLine   = (node.questNo || node.layoutFlagNo)
                ? `<br><span style="font-size:10px;color:#888">` +
                  (node.questNo      ? `Quest: ${node.questNo}` : '') +
                  (node.questNo && node.layoutFlagNo ? ' &nbsp;·&nbsp; ' : '') +
                  (node.layoutFlagNo ? `LayoutFlag: ${node.layoutFlagNo}` : '') +
                  `</span>`
                : '';
            const omLine     = node.unitId != null
                ? `<br><span style="font-size:10px;color:#666">OMID: ${node.unitId}` +
                  (node.omName ? ` &nbsp;(${node.omName})` : '') +
                  `</span>`
                : '';
            const coordLine  = `<br><span style="font-size:11px;color:#555">X:&nbsp;${node.x.toFixed(0)}&nbsp; Y:&nbsp;${node.y.toFixed(0)}&nbsp; Z:&nbsp;${node.z.toFixed(0)}</span>`;
            const groupLine  = `<br><span style="color:#666;font-size:10px">Group ${node.groupId} · pos ${node.posId}</span>`;

            const popupHtml  =
                `<span style="font-weight:bold;color:#e65c00">Destroyable Object</span>` +
                questLine + hitsLine + condLine + omLine + groupLine + coordLine;

            const tooltipText = node.questName
                ? `Destroyable Object — ${node.questName}`
                : `Destroyable Object (group ${node.groupId})`;

            const icon = L.divIcon({
                className: '',
                html: `<div style="color:#ffb300;font-size:20px;line-height:1;text-shadow:0 0 4px #000,0 0 4px #000;margin:-2px 0 0 -2px">◈</div>`,
                iconSize:    [18, 18],
                iconAnchor:  [8, 11],
                popupAnchor: [0, -10],
            });

            L.marker(latlng, { icon })
                .bindPopup(popupHtml)
                .bindTooltip(tooltipText, { permanent: false, direction: 'top', offset: [0, -10] })
                .addTo(breakTargetLayer);
        }
    }
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

                
                // ── View mode ────────────────────────────────────────────────────
                if (!shop.items.length) return header + '<br><span style="color:#888;font-size:11px">No inventory data</span>';
                const viewRows = shop.items.map(it => {
                    const entry    = itemNames[String(it.ItemId)];
                    const name     = entry?.name ?? `Item #${it.ItemId}`;
                    const href     = `https://reference.dd-on.com/build/i${String(it.ItemId).padStart(8, '0')}.html`;
                    const nameLink = `<a href="${href}" target="_blank" style="color:inherit;text-decoration:none" onmouseover="this.style.textDecoration='underline'" onmouseout="this.style.textDecoration='none'">${name}</a>`;
                    const stock    = it.Stock === 255 ? '' : ` <span style="color:#888">(×${it.Stock})</span>`;
                    return `<tr><td style="color:#222;padding-right:8px;white-space:nowrap">${itemIcon(it)}${nameLink}</td>` +
                           `<td style="color:#555;white-space:nowrap">${it.Price} ${currency}${stock}</td></tr>`;
                }).join('');
                return header + currencyLine +
                    `<br><div style="max-height:280px;overflow-y:auto;overflow-x:hidden;margin-top:4px;min-width:280px">` +
                    `<table style="font-size:13px;border-collapse:collapse;line-height:1.8;width:100%">${viewRows}</table></div>`;
            };

            const icon = L.divIcon({
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
}

function loadSpecialShops(info, stid = null) {
    specialShopLayer.clearLayers();
    if (!info.stages?.length) return;

    const floorObbs     = info.floor_obbs ?? null;
    const filterByFloor = floorObbs !== null;
    const stagesToLoad  = (stid && info.stages.includes(stid)) ? [stid] : info.stages;

    for (const stageId of stagesToLoad) {
        const stageNo = String(parseInt(stageId.slice(2), 10));
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
                const href = `https://reference.dd-on.com/build/i${String(itemId).padStart(8, '0')}.html`;
                return `<a href="${href}" target="_blank" style="color:inherit;text-decoration:none" onmouseover="this.style.textDecoration='underline'" onmouseout="this.style.textDecoration='none'">${name}</a>`;
            };

            // ── Crest badge rendering ─────────────────────────────────────────
            const crestRef = (crestId) => {
                const name = itemNames[String(crestId)]?.name ?? `#${crestId}`;
                const href = `https://reference.dd-on.com/build/i${String(crestId).padStart(8, '0')}.html`;
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

            const icon = L.divIcon({
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

                                    };

                if (_specialShopCache) { rebuildPopup(_specialShopCache); return; }
                _specialShopPromise.then(cache => { if (self.isPopupOpen()) rebuildPopup(cache); });
            });
        }
    }
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

function loadLandmarks(mapName, info) {
    landmarkLayer.clearLayers();
    const entries = landmarkData[mapName];
    if (!entries) return;

    for (const lm of entries) {
        if (HIDDEN_LANDMARK_TYPES.has(lm.type)) continue;
        const latlng = worldToPixel(lm.x, lm.z, info);
        const color = LANDMARK_COLORS[lm.type] ?? '#aaaaaa';
        const label = lm.type.replace('TYPE_', '').replace(/_/g, ' ');
        L.circleMarker(latlng, {
            color,
            fillColor: color,
            fillOpacity: 0.85,
            radius: 6,
            weight: 1.5,
        })
        .bindTooltip(label, { permanent: false, direction: 'top', offset: [0, -6] })
        .addTo(landmarkLayer);
    }
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
    exitsList.innerHTML = '';

    const allEntries = connectionData[mapName];
    if (!allEntries) { exitsPanel.style.display = 'none'; return; }

    // Filter by active stage: connections with a from_stage only show when that
    // stage is active (or when no stid is set, e.g. navigating directly by map name).
    const stid = currentStageName();
    const activeStageNo = stid ? parseInt(stid.slice(2), 10) : null;
    const stageFiltered = allEntries.filter(c =>
        c.from_stage == null || activeStageNo == null || c.from_stage === activeStageNo
    );

    // Proximity deduplication: if two connections go to the same destination and
    // are within 500 world units of each other, keep only the first.  This handles
    // near-duplicate markers from FAA data and multi-stage maps without from_stage.
    const DEDUP_DIST = 500;
    const entries = [];
    for (const conn of stageFiltered) {
        if (conn.x == null) { entries.push(conn); continue; }
        const near = entries.some(c =>
            c.to_stage === conn.to_stage && c.x != null &&
            Math.hypot(c.x - conn.x, c.z - conn.z) < DEDUP_DIST
        );
        if (!near) entries.push(conn);
    }

    const unpositioned = [];

    for (const conn of entries) {
        const navMap  = (conn.to_map && mapParams[conn.to_map]) ? conn.to_map : null;
        const hasMap  = !!navMap;
        const stageId = `st${String(conn.to_stage).padStart(4, '0')}`;
        const destName = (conn.name_en || `Stage ${conn.to_stage}`) + ` (${stageId})`;

        // Unpositioned connections (pd stage exits, etc.) go in the sidebar list
        if (conn.x == null || conn.z == null) {
            unpositioned.push({ navMap, hasMap, stageId, destName });
            continue;
        }

        // Floor filter: on multi-floor maps only show connections on the active floor
        if (info.floor_obbs) {
            const floor = getEnemyFloor(conn.x, conn.y ?? 0, conn.z, info.floor_obbs);
            if (floor !== null && floor !== currentLayer) continue;
        }

        const latlng = worldToPixel(conn.x, conn.z, info);
        const color  = hasMap ? '#ff6b35' : '#666666';

        const icon = L.divIcon({
            className: '',
            html: `<div style="
                width:14px;height:14px;
                background:${color};
                border:2px solid #fff;
                border-radius:3px;
                transform:rotate(45deg);
                box-shadow:0 0 4px rgba(0,0,0,0.7);
            "></div>`,
            iconSize: [14, 14],
            iconAnchor: [7, 7],
        });

        const marker = L.marker(latlng, { icon });
        marker.bindTooltip(destName, { permanent: false, direction: 'top', offset: [0, -10] });
        if (hasMap) {
            marker.on('click', () => navigateTo(navMap, stageId, arrivalView(mapName, navMap, conn.from_stage ?? null)));
        } else {
            marker.bindPopup(`No map data for Stage ${conn.to_stage}<br>${destName}`);
        }
        marker.addTo(connectionLayer);
    }

    // Render unpositioned exits in the sidebar
    if (unpositioned.length) {
        for (const { navMap, hasMap, stageId, destName } of unpositioned) {
            const li = document.createElement('div');
            li.style.cssText = 'padding:2px 0;font-size:0.78rem;';
            if (hasMap) {
                const a = document.createElement('span');
                a.textContent = destName;
                a.style.cssText = 'color:#ff6b35;cursor:pointer;text-decoration:underline dotted;';
                a.addEventListener('click', () => navigateTo(navMap, stageId, arrivalView(mapName, navMap, conn.from_stage ?? null)));
                li.appendChild(a);
            } else {
                li.textContent = destName;
                li.style.color = '#666';
            }
            exitsList.appendChild(li);
        }
        exitsPanel.style.display = '';
    } else {
        exitsPanel.style.display = 'none';
    }
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
        loadBreakTargets(info, currentStageName());
    }
    updateLayersInHash();
}

function swapMapImage(info, imgFile) {
    if (imageOverlay) imageOverlay.remove();
    const bounds = [xy(0, 0), xy(info.img_width, info.img_height)];
    imageOverlay = L.imageOverlay('images/maps/' + imgFile, bounds, { pane: 'mapImagePane' }).addTo(leafletMap);
}

// ── Tile-layer selector (pd maps with multi-layer pieces) ─────────────────────
let _tileLayerSel = null;   // id of currently shown tile-layer image, null = merged default

function buildTileLayerSelector(info) {
    const el = document.getElementById('tile-layer-selector');
    if (!el) return;
    el.innerHTML = '';
    _tileLayerSel = null;
    const tlImages = info.tile_layer_images;
    if (!tlImages || Object.keys(tlImages).length === 0) return;

    const makeBtn = (label, key, imgFile) => {
        const btn = document.createElement('button');
        btn.textContent = label;
        if (key === null) btn.classList.add('active');
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
function resetView() {
    const info = mapParams[_loadedMapName];
    if (!info) return;
    if (info.img_exists) {
        leafletMap.fitBounds([xy(0, 0), xy(info.img_width, info.img_height)]);
    } else {
        leafletMap.setView(xy(info.img_width / 2, info.img_height / 2), 0);
    }
}

function loadMap(mapName) {
    const info = mapParams[mapName];
    if (!info) return;

    _loadedMapName = mapName;
    _loadedStid = currentStageName();
    currentLayer = 0;

    // Update title — always append an ID so the user knows which stage they're on.
    // Prefer the stid from the URL hash (e.g. "st0200"); fall back to the map name.
    const stid = currentStageName();
    const baseName = stid ? stageLabel(info, stid) : (info.name_en ? splitPascalCase(info.name_en) : mapName);
    const sidNum   = stid ? info.stage_ids?.[stid] : null;
    const sidStr   = sidNum != null ? ` - sid${String(sidNum).padStart(4, '0')}` : '';
    const title = baseName + ` (${stid ?? mapName}${sidStr})`;
    document.getElementById('map-title').textContent = title;
    document.title = `${title} — DDON Maps`;

    // Replace image overlay
    if (imageOverlay) imageOverlay.remove();
    const savedView = parseHash().view;
    if (info.img_exists) {
        const bounds = [xy(0, 0), xy(info.img_width, info.img_height)];
        imageOverlay = L.imageOverlay('images/maps/' + info.img_file, bounds, { pane: 'mapImagePane' }).addTo(leafletMap);
        if (savedView) {
            leafletMap.setView(savedView.center, savedView.zoom);
        } else {
            leafletMap.fitBounds(bounds);
        }
    } else {
        imageOverlay = null;
        leafletMap.setView(xy(info.img_width / 2, info.img_height / 2), 0);
    }

    // Restore floor: explicit hash value wins, then arrival-connection detection, else stays 0.
    if (info.floor_obbs) {
        const savedFloor = parseHash().view?.floor;
        if (savedFloor != null) {
            currentLayer = savedFloor;
        } else if (stid) {
            const arrivalStageNo = parseInt(stid.slice(2), 10);
            const conns = connectionData[_loadedMapName] ?? [];
            const arrConn = conns.find(c => c.from_stage === arrivalStageNo && c.x != null && c.z != null);
            if (arrConn) {
                const arrFloor = getEnemyFloor(arrConn.x, arrConn.y ?? 0, arrConn.z, info.floor_obbs);
                if (arrFloor !== null) currentLayer = arrFloor;
            }
        }
        if (currentLayer !== 0) {
            const floorLayer = (info.layers || []).find(l => l.layer === currentLayer && l.img_exists);
            if (floorLayer) swapMapImage(info, floorLayer.img_file);
        }
    }

    // Build floor selector for multi-floor maps
    buildFloorSelector(info);
    buildTileLayerSelector(info);
    buildStageGroupsPanel(info, currentStageName());

    // Reload layers
    loadGrid(info);
    loadPdBoundaries(info);
    loadStageLabels(info);
    loadLandmarks(mapName, info);
    loadConnections(mapName, info);
    loadGatherPoints(info, currentStageName());
    loadNpcShops(info, currentStageName());
    loadSpecialShops(info, currentStageName());
    loadBreakTargets(info, currentStageName());
    // Read openGroups BEFORE loadEnemySpawns — that function calls _updateExpandCollapseBtn
    // → updateLayersInHash which would overwrite the hash (erasing the group list) if read after.
    const { openGroups } = parseHash();

    loadEnemySpawns(info, currentStageName());

    if (openGroups?.length) {
        for (const id of openGroups) if (_groupStore.has(id)) _expandGroupCore(_groupStore.get(id));
        _updateExpandCollapseBtn();
        reapplySpread();
    }

    _spotOpenedGroup = null;
    buildSpotIndex(info);
    if (document.getElementById('spot-panel')?.classList.contains('open')) _runSpotSearch();

    if (_pendingGlobalNavTarget) {
        const t = _pendingGlobalNavTarget;
        _pendingGlobalNavTarget = null;
        setTimeout(() => _navigateToSpot(t), 300);
    }
}

// ── Spot search panel ─────────────────────────────────────────────────────────

let _spotIndex        = [];    // searchable entries for the current map
let _spotHighlights   = [];    // active pulsing ring markers
let _spotOpenedGroup  = null;  // groupId of the enemy group last opened by spot search
const _spotHlLayer  = L.layerGroup().addTo(leafletMap);
let _spotGlobal       = false; // true = global (all stages) search mode
let _globalSpotIndex  = [];    // searchable entries across all maps/stages
let _pendingGlobalNavTarget = null; // deferred navigation target after stage switch

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

function buildSpotIndex(info) {
    _spotIndex = [];
    if (!info.stages?.length) return;

    const stid = currentStageName();
    const stagesToLoad = (stid && info.stages.includes(stid)) ? [stid] : info.stages;

    for (const stageId of stagesToLoad) {
        const stageNo = String(parseInt(stageId.slice(2), 10));
        const serverStageId = stageIds[stageNo];
        const cache = _enemySpawnCache;  // may be null if promise not yet resolved

        // ── Enemies: one entry per emCode per spawn position ────────────────
        const groups = enemyPositions[stageNo];
        if (groups) {
            for (const [groupId, groupData] of Object.entries(groups)) {
                const spawns = groupData.spawns ?? groupData;  // back-compat: array may be direct
                if (!Array.isArray(spawns) || !spawns.length) continue;
                for (let i = 0; i < spawns.length; i++) {
                    const s = spawns[i];
                    const posLatlng = worldToPixel(s.Position.x, s.Position.z, info);
                    const spawnKey  = serverStageId != null ? `${serverStageId},${groupId},${s.posIdx ?? i}` : null;
                    if (cache && spawnKey) {
                        const byEmCode = new Map(); // emCode → Set<level>
                        for (const e of (cache.get(spawnKey) ?? [])) {
                            if (!e.emCode) continue;
                            if (!byEmCode.has(e.emCode)) byEmCode.set(e.emCode, new Set());
                            if (e.lv != null) byEmCode.get(e.emCode).add(e.lv);
                        }
                        for (const [emCode, lvSet] of byEmCode) {
                            const baseName = emNames[emCode]?.name;
                            if (!baseName) continue;
                            const lvs = [...lvSet].sort((a, b) => a - b);
                            const lo = lvs[0], hi = lvs[lvs.length - 1];
                            const lvLabel = lvs.length ? (lo === hi ? `Lv${lo}` : `Lv${lo}-${hi}`) : '';
                            const displayName = lvLabel ? `${baseName} ${lvLabel}` : baseName;
                            _spotIndex.push({
                                type: 'enemy', name: displayName,
                                searchText: `${baseName} ${lvLabel}`.toLowerCase(),
                                latlng: posLatlng, groupId, emCode, spawnKey,
                                worldPos: { x: s.Position.x, y: s.Position.y, z: s.Position.z },
                                previewLines: [
                                    `<b>${displayName}</b>`,
                                    `Code: ${emCode}`,
                                    `Group ${groupId}`,
                                ],
                                stageId,
                            });
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
        if (_shopCache) {
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
function _rebuildSpotIndex() {
    if (_currentMapInfo) {
        buildSpotIndex(_currentMapInfo);
        if (document.getElementById('spot-panel')?.classList.contains('open')) _runSpotSearch();
    }
}
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

function _buildGlobalSpotIndex() {
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
            const groups = enemyPositions[stageNo];
            if (groups) {
                for (const [groupId, groupData] of Object.entries(groups)) {
                    const spawns = groupData.spawns ?? groupData;
                    if (!Array.isArray(spawns) || !spawns.length) continue;
                    for (let i = 0; i < spawns.length; i++) {
                        const s = spawns[i];
                        const spawnKey = serverStageId != null ? `${serverStageId},${groupId},${s.posIdx ?? i}` : null;
                        if (_enemySpawnCache && spawnKey) {
                            const byEmCode = new Map();
                            for (const e of (_enemySpawnCache.get(spawnKey) ?? [])) {
                                if (!e.emCode) continue;
                                if (!byEmCode.has(e.emCode)) byEmCode.set(e.emCode, new Set());
                                if (e.lv != null) byEmCode.get(e.emCode).add(e.lv);
                            }
                            for (const [emCode, lvSet] of byEmCode) {
                                const baseName = emNames[emCode]?.name;
                                if (!baseName) continue;
                                const lvs = [...lvSet].sort((a, b) => a - b);
                                const lo = lvs[0], hi = lvs[lvs.length - 1];
                                const lvLabel = lvs.length ? (lo === hi ? `Lv${lo}` : `Lv${lo}-${hi}`) : '';
                                const displayName = lvLabel ? `${baseName} ${lvLabel}` : baseName;
                                _globalSpotIndex.push({
                                    type: 'enemy', name: displayName,
                                    searchText: `${baseName} ${lvLabel}`.toLowerCase(),
                                    worldPos: { x: s.Position.x, y: s.Position.y, z: s.Position.z },
                                    groupId, emCode, spawnKey,
                                    mapName, stageId, stageNo, locationTag,
                                });
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
            if (_shopCache) {
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
}

function _rebuildGlobalSpotIndex() {
    _buildGlobalSpotIndex();
    if (_spotGlobal && document.getElementById('spot-panel')?.classList.contains('open')) _runSpotSearch();
}

// Parse a search query for exact-match syntax: "quoted phrase" → prefix match on name.
// Returns { term: string, exact: boolean }
function _parseSpotQuery(raw) {
    if (raw.length > 2 && raw.startsWith('"') && raw.endsWith('"'))
        return { term: raw.slice(1, -1), exact: true };
    return { term: raw, exact: false };
}

function _spotEntryMatches(e, term, exact) {
    if (!term) return true;
    if (exact) return e.name.toLowerCase().startsWith(term);
    return e.searchText.includes(term);
}

function _runSpotSearch() {
    const raw       = (document.getElementById('spot-search-input')?.value ?? '').trim().toLowerCase();
    const { term, exact } = _parseSpotQuery(raw);
    const filter    = document.querySelector('.spot-tab.active')?.dataset.filter ?? 'enemy';
    const resultsEl = document.getElementById('spot-results');
    if (!resultsEl) return;

    _clearSpotHighlights();

    if (_spotGlobal) {
        if (!raw) {
            resultsEl.innerHTML = `<div class="spot-empty">Enter a search term to search across all stages.</div>`;
            return;
        }
        const matches = _globalSpotIndex.filter(e => {
            if (filter === 'enemy'  && e.type !== 'enemy')  return false;
            if (filter === 'gather' && e.type !== 'gather') return false;
            if (filter === 'item'   && e.type !== 'item')   return false;
            return _spotEntryMatches(e, term, exact);
        });
        if (!matches.length) {
            resultsEl.innerHTML = `<div class="spot-empty">No matches for <em>${raw}</em> across all stages.</div>`;
            return;
        }
        _renderGlobalResults(matches, resultsEl);
        return;
    }

    const matches = _spotIndex.filter(e => {
        if (filter === 'enemy'  && e.type !== 'enemy')  return false;
        if (filter === 'gather' && e.type !== 'gather') return false;
        if (filter === 'item'   && e.type !== 'item')   return false;
        return _spotEntryMatches(e, term, exact);
    });

    if (!matches.length) {
        resultsEl.innerHTML = `<div class="spot-empty">No matches for <em>${raw}</em>.</div>`;
        return;
    }

    // Group by name+type (items also group by source so enemy/gather/shop stay separate)
    const grouped = new Map();
    for (const m of matches) {
        const key = m.type === 'item' ? `item:${m.source}\0${m.name}` : `${m.type}\0${m.name}`;
        if (!grouped.has(key)) grouped.set(key, []);
        grouped.get(key).push(m);
    }

    const frag = document.createDocumentFragment();
    const summary = document.createElement('div');
    summary.className = 'spot-summary';
    summary.textContent = `${matches.length} result${matches.length !== 1 ? 's' : ''} · ${grouped.size} unique`;
    frag.appendChild(summary);

    for (const items of grouped.values()) {
        const first = items[0];
        const multi = items.length > 1;

        const isBossResult = first.type === 'enemy' && _enemySpawnCache && items.some(it =>
            it.spawnKey && (_enemySpawnCache.get(it.spawnKey) ?? []).some(e => e.isBossGauge || e.isAreaBoss || e.raidBossId > 0));
        const dotHtml = first.type === 'gather'
            ? `<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${GATHER_COLORS[first.gatherType] ?? '#aaa'};flex-shrink:0"></span>`
            : first.type === 'item' && first.source === 'gather'
            ? `<span style="font-size:10px;line-height:1;flex-shrink:0;color:#8c8">🌿</span>`
            : first.type === 'item' && first.source === 'shop'
            ? `<span style="font-size:10px;line-height:1;flex-shrink:0;color:#fc8">🏪</span>`
            : first.type === 'item' && first.source === 'enemy'
            ? `<span style="font-size:10px;line-height:1;flex-shrink:0;color:#c88">⚔</span>`
            : isBossResult
            ? `<span style="font-size:10px;line-height:1;flex-shrink:0;color:#f44">☠</span>`
            : `<span style="font-size:10px;line-height:1;flex-shrink:0;color:#c88">⚔</span>`;

        const previewHtml = multi
            ? [...first.previewLines, `<span style="color:#667">×${items.length} locations</span>`].join('<br>')
            : first.previewLines.join('<br>');

        const row = document.createElement('div');
        row.className = 'spot-result-row';
        row.title = first.name;

        if (multi) {
            // ◀ name  1/N ▶
            let idx = -1;  // -1 = not yet visited; first row click goes to 0, subsequent advance
            const updatePos = (posEl) => { posEl.textContent = `${idx + 1}/${items.length}`; };
            row.innerHTML =
                `${dotHtml}<span class="spot-result-name">${first.name}</span>`
                + `<div class="spot-nav" style="display:flex;align-items:center;gap:2px;flex-shrink:0">`
                + `<button class="spot-nav-btn spot-prev" title="Previous">◀</button>`
                + `<span class="spot-nav-pos" style="font-size:0.68rem;color:#667;min-width:28px;text-align:center">1/${items.length}</span>`
                + `<button class="spot-nav-btn spot-next" title="Next">▶</button>`
                + `</div>`
                + `<div class="spot-preview">${previewHtml}</div>`;

            const posEl  = row.querySelector('.spot-nav-pos');
            const prev   = row.querySelector('.spot-prev');
            const next   = row.querySelector('.spot-next');

            const goTo = (i) => {
                idx = (i + items.length) % items.length;
                updatePos(posEl);
                _navigateToSpot(items[idx]);
            };

            prev.addEventListener('click', e => { e.stopPropagation(); goTo(idx <= 0 ? items.length - 1 : idx - 1); });
            next.addEventListener('click', e => { e.stopPropagation(); goTo(idx + 1); });
            // Row click: first click goes to 0, subsequent clicks advance to next (wraps)
            row.addEventListener('click', e => {
                if (e.target.closest('.spot-nav')) return;  // ignore clicks on the nav buttons
                goTo(idx < 0 ? 0 : idx + 1);
            });
        } else {
            row.innerHTML = `${dotHtml}<span class="spot-result-name">${first.name}</span>`
                + `<div class="spot-preview">${previewHtml}</div>`;
            row.addEventListener('click', () => _navigateToSpot(first));
        }

        row.addEventListener('mouseenter', () => {
            _clearSpotHighlights();
            const drawnGroups = new Set();
            for (const item of items) {
                if (_currentFloorObbs && item.worldPos) {
                    const f = getEnemyFloor(item.worldPos.x, item.worldPos.y, item.worldPos.z, _currentFloorObbs);
                    if (f !== null && f !== currentLayer) continue;
                }
                const grp = item.groupId ? _groupStore.get(item.groupId) : null;
                if (item.groupId && !grp) continue;  // group not on this floor, skip
                if (grp && !grp.isExpanded) {
                    // Collapsed: highlight the chip element directly — no separate marker needed
                    if (!drawnGroups.has(item.groupId)) {
                        drawnGroups.add(item.groupId);
                        _addChipHighlight(grp);
                    }
                } else {
                    _addSpotHighlight(_resolveSpotLatLng(item));
                }
            }
        });
        row.addEventListener('mouseleave', _clearSpotHighlights);

        frag.appendChild(row);
    }

    resultsEl.innerHTML = '';
    resultsEl.appendChild(frag);
}

function _renderGlobalResults(matches, resultsEl) {
    // Group by name + type + source + stage — each row is one name on one stage
    const grouped = new Map();
    for (const m of matches) {
        const key = m.type === 'item'
            ? `item:${m.source}\0${m.name}\0${m.mapName}\0${m.stageId}`
            : `${m.type}\0${m.name}\0${m.mapName}\0${m.stageId}`;
        if (!grouped.has(key)) grouped.set(key, []);
        grouped.get(key).push(m);
    }

    // Sort groups by locationTag (stage name) then by name
    const sortedGroups = [...grouped.values()].sort((a, b) => {
        const locCmp = a[0].locationTag.localeCompare(b[0].locationTag);
        return locCmp !== 0 ? locCmp : a[0].name.localeCompare(b[0].name);
    });

    const uniqueNames = new Set(matches.map(m => m.name)).size;
    const frag = document.createDocumentFragment();
    const summary = document.createElement('div');
    summary.className = 'spot-summary';
    summary.textContent = `${matches.length} result${matches.length !== 1 ? 's' : ''} · ${uniqueNames} unique · ${sortedGroups.length} stage entries`;
    frag.appendChild(summary);

    for (const items of sortedGroups) {
        const first = items[0];
        const multi = items.length > 1;

        const dotHtml = first.type === 'gather'
            ? `<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${GATHER_COLORS[first.gatherType] ?? '#aaa'};flex-shrink:0"></span>`
            : first.type === 'item' && first.source === 'gather'
            ? `<span style="font-size:10px;line-height:1;flex-shrink:0;color:#8c8">🌿</span>`
            : first.type === 'item' && first.source === 'shop'
            ? `<span style="font-size:10px;line-height:1;flex-shrink:0;color:#fc8">🏪</span>`
            : first.type === 'item' && first.source === 'enemy'
            ? `<span style="font-size:10px;line-height:1;flex-shrink:0;color:#c88">⚔</span>`
            : `<span style="font-size:10px;line-height:1;flex-shrink:0;color:#c88">⚔</span>`;

        const row = document.createElement('div');
        row.className = 'spot-result-row';
        row.title = `${first.name} — ${first.locationTag}`;

        if (multi) {
            let idx = -1;
            row.innerHTML =
                `${dotHtml}<div style="flex:1;min-width:0">`
                + `<div class="spot-result-name">${first.name}</div>`
                + `<div class="spot-stage-sub">${first.locationTag}</div>`
                + `</div>`
                + `<div class="spot-nav" style="display:flex;align-items:center;gap:2px;flex-shrink:0">`
                + `<button class="spot-nav-btn spot-prev" title="Previous">◀</button>`
                + `<span class="spot-nav-pos" style="font-size:0.68rem;color:#667;min-width:32px;text-align:center">×${items.length}</span>`
                + `<button class="spot-nav-btn spot-next" title="Next">▶</button>`
                + `</div>`;

            const posEl   = row.querySelector('.spot-nav-pos');
            const stageSub = row.querySelector('.spot-stage-sub');
            const prev    = row.querySelector('.spot-prev');
            const next    = row.querySelector('.spot-next');

            const goTo = (i) => {
                idx = (i + items.length) % items.length;
                posEl.textContent = `${idx + 1}/${items.length}`;
                _navigateToSpotGlobal(items[idx]);
            };

            prev.addEventListener('click', e => { e.stopPropagation(); goTo(idx <= 0 ? items.length - 1 : idx - 1); });
            next.addEventListener('click', e => { e.stopPropagation(); goTo(idx + 1); });
            row.addEventListener('click', e => {
                if (e.target.closest('.spot-nav')) return;
                goTo(idx < 0 ? 0 : idx + 1);
            });
        } else {
            row.innerHTML =
                `${dotHtml}<div style="flex:1;min-width:0">`
                + `<div class="spot-result-name">${first.name}</div>`
                + `<div class="spot-stage-sub">${first.locationTag}</div>`
                + `</div>`;
            row.addEventListener('click', () => _navigateToSpotGlobal(first));
        }
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
        // Close edit panel if open (mutually exclusive)
                panel.classList.add('open');
        toggle.style.display = 'none';
        input.focus();
        _runSpotSearch();
    };
    const closePanel = () => {
        panel.classList.remove('open');
        toggle.style.display = '';
        _clearSpotHighlights();
    };

    toggle.addEventListener('click', openPanel);
    close.addEventListener('click', closePanel);
    input.addEventListener('input', () => {
        if (clearBtn) clearBtn.style.display = input.value ? 'block' : 'none';
        _runSpotSearch();
    });
    if (clearBtn) {
        clearBtn.addEventListener('click', () => {
            input.value = '';
            clearBtn.style.display = 'none';
            input.focus();
            _runSpotSearch();
        });
    }

    document.querySelectorAll('.spot-tab').forEach(btn =>
        btn.addEventListener('click', () => {
            document.querySelectorAll('.spot-tab').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            _runSpotSearch();
        })
    );

    document.querySelectorAll('.spot-scope').forEach(btn =>
        btn.addEventListener('click', () => {
            document.querySelectorAll('.spot-scope').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            _spotGlobal = btn.dataset.scope === 'global';
            if (_spotGlobal && !_globalSpotIndex.length) _buildGlobalSpotIndex();
            _clearSpotHighlights();
            _runSpotSearch();
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
})();

// Pre-build global spot index in the background after startup
setTimeout(_buildGlobalSpotIndex, 0);

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
        };
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
    });
}

_initPanelResize({ handleId: 'sidebar-resize-handle', panelId: 'sidebar',    lsKey: 'ddon-sidebar-width',    minW: 180, maxW: 520, dragDir: 'right' });
_initPanelResize({ handleId: 'edit-resize-handle',    panelId: 'edit-panel', lsKey: 'ddon-edit-panel-width', minW: 240, maxW: 600, dragDir: 'left'  });
_initPanelResize({ handleId: 'spot-resize-handle',    panelId: 'spot-panel', lsKey: 'ddon-spot-panel-width', minW: 200, maxW: 600, dragDir: 'left'  });

// Patch loadMap to keep currentInfo updated
const _origLoadMap = loadMap;
loadMap = function (mapName) {
    _origLoadMap(mapName);
    if (window._setCurrentInfo) window._setCurrentInfo(mapParams[mapName]);
};

// ── Init ──────────────────────────────────────────────────────────────────────
if (!location.hash || location.hash === '#') {
    history.replaceState(null, '', '#field000_m00:st0100');
}
buildSidebar();
loadMap(currentMapName());
