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
const namedParamsById = new Map(namedParamList.map(p => [p.id, p]));
const hmPresetsByEmCode = new Map(hmPresetList.filter(p => p.emCode).map(p => [p.emCode, p]));

function namedParamLabel(p) {
    if (!p) return '0 (None)';
    const trimName = p.name?.trim();
    const typeSuffix = p.type === 'NAMED_TYPE_PREFIX' ? ' [Pfx]'
        : p.type === 'NAMED_TYPE_REPLACE'             ? ' [Rep]'
        : p.type === 'NAMED_TYPE_SUFFIX'              ? ' [Sfx]'
        : '';
    return trimName ? `${p.id}: ${trimName}${typeSuffix}` : `#${p.id}${typeSuffix}`;
}

// ── Leaflet map setup ──────────────────────────────────────────────────────────
const leafletMap = L.map('map', {
    crs: L.CRS.Simple,
    maxZoom: 6,
    minZoom: -3,
    zoomSnap: 0.5,
});
leafletMap.createPane('mapImagePane');
leafletMap.getPane('mapImagePane').style.zIndex = 201;

function xy(x, y) { return L.latLng(y, x); }

L.Control.ResetView = L.Control.extend({
    options: { position: 'topleft' },
    onAdd() {
        const container = L.DomUtil.create('div', 'leaflet-bar leaflet-control');
        const btn = L.DomUtil.create('a', 'leaflet-control-reset-view', container);
        btn.innerHTML = '&#8962;';
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
function worldToPixel(worldX, worldZ, info) {
    let png_y;
    if (info.pd_pieces?.length) {
        const pieces = info.pd_pieces;
        let piece = pieces[0];
        for (const p of pieces) {
            const rangeSize = p.full_size ?? p.size;
            if (worldZ >= p.connect_z + rangeSize && worldZ <= p.connect_z) {
                piece = p;
                break;
            }
        }
        const localZ  = worldZ - piece.connect_z;
        png_y = piece.pixel_y_entrance_v + localZ * info.scale;
        png_y = Math.max(piece.pixel_y_start, Math.min(piece.pixel_y_entrance, png_y));
    } else {
        const scaleZ = info.scale_z ?? info.scale;
        png_y = (info.img_height - info.center_y) - worldZ * scaleZ;
        png_y = info.img_height - png_y;
    }
    const py = info.img_height - png_y;
    const px = worldX * info.scale + info.center_x;
    return xy(px, py);
}

// ── Layer groups ───────────────────────────────────────────────────────────────
let imageOverlay    = null;
let enemyLayer      = L.layerGroup().addTo(leafletMap);
let landmarkLayer   = L.layerGroup().addTo(leafletMap);
let connectionLayer = L.layerGroup().addTo(leafletMap);
let gridLayer        = L.layerGroup();
let territoryLayer   = L.layerGroup();
let stageLabelsLayer = L.layerGroup().addTo(leafletMap);
let gatherLayer       = L.layerGroup();
const _gatherMarkerByKey = new Map();
const _shopMarkerByNpcId = new Map();
let npcShopLayer        = L.layerGroup();
let specialShopLayer    = L.layerGroup();
let breakTargetLayer  = L.layerGroup();
let pdBoundaryLayer = L.layerGroup().addTo(leafletMap);
let spawnRadiiLayer   = L.layerGroup().addTo(leafletMap);
let _spreadOverlay    = L.layerGroup().addTo(leafletMap);

const spawnRenderer = L.canvas({ padding: 0.5 });

// ── Group expand/collapse state ───────────────────────────────────────────────
const _groupStore = new Map();
let _currentMapInfo  = null;

// ── Edit mode is permanently disabled in this viewer build ────────────────────
const _editMode = false;

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
const LAYER_PREFS_KEY = 'ddon-maps-layers';

function getLayersHash() {
    let s = '';
    if (document.getElementById('layer-enemies').checked)       s += 'e';
    if (document.getElementById('layer-landmarks').checked)     s += 'l';
    if (document.getElementById('layer-connections').checked)   s += 'c';
    if (document.getElementById('layer-grid').checked)          s += 'g';
    if (document.getElementById('layer-territory').checked)     s += 't';
    if (document.getElementById('layer-stage-labels').checked)  s += 'a';
    if (document.getElementById('layer-gather').checked)        s += 'r';
    if (document.getElementById('layer-radii').checked)         s += 'i';
    if (document.getElementById('layer-npc-shops').checked)           s += 'n';
    if (document.getElementById('layer-special-shops').checked)       s += 'p';
    if (document.getElementById('layer-break-targets').checked)       s += 'b';
    if (document.getElementById('sidebar').classList.contains('collapsed')) s += 's';
    const openIds = [..._groupStore.values()]
        .filter(g => g.isExpanded)
        .map(g => g.groupId)
        .sort((a, b) => parseInt(a) - parseInt(b));
    if (openIds.length) s += ';' + openIds.join(',');
    return s;
}

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
        territory:    document.getElementById('layer-territory').checked,
        stageLabels:  document.getElementById('layer-stage-labels').checked,
        gather:       document.getElementById('layer-gather').checked,
        radii:        document.getElementById('layer-radii').checked,
        npcShops:      document.getElementById('layer-npc-shops').checked,
        specialShops:  document.getElementById('layer-special-shops').checked,
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

(function applyLayerPrefs() {
    const { layers: urlLayers } = parseHash();
    const stored = loadLayerPrefs();
    const prefs = urlLayers ?? stored ?? {};
    const isOn = (key, defaultOn) => key in prefs ? prefs[key] : defaultOn;

    document.getElementById('layer-enemies').checked       = isOn('enemies',      true);
    document.getElementById('layer-landmarks').checked     = isOn('landmarks',    true);
    document.getElementById('layer-connections').checked   = isOn('connections',  true);
    document.getElementById('layer-grid').checked          = isOn('grid',         false);
    document.getElementById('layer-territory').checked     = isOn('territory',    false);
    document.getElementById('layer-stage-labels').checked  = isOn('stageLabels',  true);
    document.getElementById('layer-gather').checked        = isOn('gather',        false);
    document.getElementById('layer-radii').checked         = isOn('radii',         false);
    document.getElementById('layer-npc-shops').checked      = isOn('npcShops',      true);
    document.getElementById('layer-special-shops').checked  = isOn('specialShops',  false);
    document.getElementById('layer-break-targets').checked  = isOn('breakTargets',  false);

    if (!document.getElementById('layer-landmarks').checked)
        leafletMap.removeLayer(landmarkLayer);
    if (!document.getElementById('layer-connections').checked)
        leafletMap.removeLayer(connectionLayer);
    if (document.getElementById('layer-grid').checked)
        leafletMap.addLayer(gridLayer);
    if (document.getElementById('layer-territory').checked)
        leafletMap.addLayer(territoryLayer);
    if (!document.getElementById('layer-stage-labels').checked)
        leafletMap.removeLayer(stageLabelsLayer);
    if (document.getElementById('layer-gather').checked)
        leafletMap.addLayer(gatherLayer);
    if (document.getElementById('layer-npc-shops').checked)
        leafletMap.addLayer(npcShopLayer);
    if (document.getElementById('layer-special-shops').checked)
        leafletMap.addLayer(specialShopLayer);
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
document.getElementById('layer-npc-shops').addEventListener('change', e => {
    e.target.checked ? leafletMap.addLayer(npcShopLayer) : leafletMap.removeLayer(npcShopLayer);
    saveLayerPrefs();
});
document.getElementById('layer-special-shops').addEventListener('change', e => {
    e.target.checked ? leafletMap.addLayer(specialShopLayer) : leafletMap.removeLayer(specialShopLayer);
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

document.getElementById('layer-territory').addEventListener('change', e => {
    if (e.target.checked) {
        leafletMap.addLayer(territoryLayer);
        for (const g of _groupStore.values())
            if (g.isExpanded && g.territoryRect) territoryLayer.addLayer(g.territoryRect);
    } else {
        territoryLayer.clearLayers();
        leafletMap.removeLayer(territoryLayer);
    }
    saveLayerPrefs();
});

document.getElementById('btn-expand-collapse').addEventListener('click', () => {
    const anyCollapsed = [..._groupStore.values()].some(g => !g.isExpanded);
    if (anyCollapsed) _expandAllGroups(); else _collapseAllGroups();
});

// ── Sidebar map list ───────────────────────────────────────────────────────────
function splitPascalCase(s) {
    let result = s.replace(/([a-z])(to)(the)(?=[A-Z])/g, '$1 $2 $3 ');
    result = result.replace(/([a-z])(to)(?=[A-Z])/g, '$1 $2 ');
    result = result.replace(/([a-z])([A-Z])/g, '$1 $2');
    result = result.replace(/([a-z])(of)(?=the\b|[A-Z\s]|$)/g, '$1 $2');
    result = result.replace(/([a-z])(the)(?=[A-Z\s]|$)/g, '$1 $2');
    result = result.replace(/([ac-z])(by)(?=\s|$)/g, '$1 $2');
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
    const raw = info.stage_names?.[stid] || info.name_en || '';
    return raw ? splitPascalCase(raw) : stid;
}

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
    for (const { key, value } of conditions) {
        if (key === 'stageid') {
            const sid = stid ? info.stage_ids?.[stid] : undefined;
            if (sid === undefined || String(sid) !== value) return false;
        } else if (key === 'stageno') {
            if (!stid) return false;
            if (parseInt(stid.slice(2), 10) !== parseInt(value, 10)) return false;
        } else if (key === 'area') {
            const aname = (info.quest_area_name ?? '').toLowerCase();
            if (!aname.includes(value)) return false;
        }
    }
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

    const pdPieceRe = /^pd\d+_m\d+$/;
    const entries = [];
    for (const [name, info] of Object.entries(mapParams)) {
        if (pdPieceRe.test(name)) continue;
        const stages = info.stages?.length ? info.stages : [null];
        for (const stid of stages) {
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

    const areaMap = new Map();
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
function parseHash() {
    const raw = window.location.hash.slice(1);
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
            npcShops:      flagStr.includes('n'),
            specialShops:  flagStr.includes('p'),
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
const OVERLAP_SPREAD_R = 9;

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
            m._spreadAnchor = anchor;
            m.on('mouseover', () => { anchor.setRadius(8); anchor.setStyle({ weight: 2.5, fillOpacity: 1.0 }); });
            m.on('mouseout',  () => { anchor.setRadius(4); anchor.setStyle({ weight: 1.5, fillOpacity: 0.85 }); });
        });

        group.forEach(m => {
            const spoke = L.polyline([m.getLatLng(), origin], {
                color: m.options.color, weight: 1, opacity: 0.4,
                dashArray: '3 3', interactive: false,
            }).addTo(overlayLayer);
            m._spokeLine = spoke;

            m.on('mouseover', () => spoke.setStyle({ weight: 2.5, opacity: 1.0, dashArray: null }));
            m.on('mouseout',  () => spoke.setStyle({ weight: 1,   opacity: 0.4,  dashArray: '3 3' }));
        });

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

function pointInPolygon(px, py, pts) {
    let inside = false;
    for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) {
        const [xi, yi] = pts[i], [xj, yj] = pts[j];
        if ((yi > py) !== (yj > py) && px < (xj - xi) * (py - yi) / (yj - yi) + xi)
            inside = !inside;
    }
    return inside;
}

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
    const chipColor = `oklch(0.78 0.13 ${(parseInt(groupId, 10) * 137) % 360})`;
    const glows = [];
    if (isBossGroup)      glows.push('0 0 7px 2px rgba(255,60,60,0.85)');
    if (isKeyBearerGroup) glows.push('0 0 7px 2px rgba(255,210,0,0.85)');
    const shadowStyle = glows.length ? `box-shadow:0 0 4px rgba(0,0,0,0.7),${glows.join(',')};` : '';
    const titleAttr = [isBossGroup ? 'Contains boss enemy' : '', isKeyBearerGroup ? 'Key bearer group' : ''].filter(Boolean).join(' · ');
    return L.divIcon({
        className: '',
        html: `<div class="group-chip${expanded ? ' chip-open' : ''}" style="color:${chipColor};${shadowStyle}"${titleAttr ? ` title="${titleAttr}"` : ''}><span class="chip-arrow${expanded ? ' open' : ''}">&#9654;</span>G${groupId} <span class="chip-count">${count}</span></div>`,
        iconSize:   null,
        iconAnchor: expanded ? [0, 22] : [0, yOffset],
    });
}

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

function buildGroupDetails(g) {
    const info  = _currentMapInfo;
    const layer = L.layerGroup();

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
                interactive: false,
            });
            layer.addLayer(poly);
            g.hullPts = hull;
        }
    } else if (g.pts.length === 2) {
        L.polyline(g.pts.map(([px, py]) => xy(px, py)), {
            color: g.color, weight: 1.5, opacity: 0.65, dashArray: '4 3', interactive: false,
        }).addTo(layer);
    }

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

        const serverStageId = stageIds[stageNo];
        const spawnKey = serverStageId != null
            ? `${serverStageId},${g.groupId},${spawn.posIdx ?? idx}` : null;

        let displayIdx = 0;

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
            const hasEnemy  = !spawnCache || entries.some(e => !!e.lv);
            const emCode    = spawnInfo?.emCode ?? (hasEnemy ? spawn.EmName : null);
            const emEntry   = emCode ? emNames[emCode] : null;
            const dispName  = emEntry?.name ?? null;
            const lvText    = spawnInfo?.lv ? ` Lv${spawnInfo.lv}` : '';
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
                ? `<br><span style="color:#333;font-size:12px">${shownName}${lvText}</span>${emCodeLine}` : '';

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

            const buildBossLine = (si) => {
                if (!si) return '';
                const parts = [];
                if (si.isBossGauge) parts.push('Boss Gauge');
                if (si.isBossBGM)   parts.push('Boss BGM');
                if (si.isAreaBoss)  parts.push('Area Boss');
                if (si.raidBossId > 0) parts.push(`Raid Boss ID: ${si.raidBossId}`);
                return parts.length ? `<br><span style="font-size:11px;color:#ff6666">☠ ${parts.join(' · ')}</span>` : '';
            };

            const buildManualSetLine = (si) =>
                si?.isManualSet
                    ? `<br><span style="font-size:11px;color:#b0c4ff">&#128564; Dormant until summoned</span>`
                    : '';

            return `${badge}<br>${groupLabel}, Index: <b>${idx}</b>${subLine}${triggerLine}${cycleHtml}${emLine}${keyLine}${buildBossLine(spawnInfo)}${buildManualSetLine(spawnInfo)}${radiiLine}${orbsLine}${buildDropsHtml(spawnInfo)}`;
        };

        const buildTooltip = (spawnCache) => {
            const entries  = spawnKey && spawnCache ? (spawnCache.get(spawnKey) ?? []) : [];
            const hasEnemy = !spawnCache || entries.some(e => !!e.lv);
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
            .bindPopup(buildEnemyPopup(_enemySpawnCache), { minWidth: 260, maxWidth: 420 })
            .bindTooltip('', { direction: 'top', offset: [0, -8] });

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

                    // Wire cycle prev/next buttons
                    let clickHandler = null;
                    if (clickHandler) el.removeEventListener('click', clickHandler);
                    clickHandler = (e) => {
                        const cycleBtn = e.target.closest('.spawn-prev, .spawn-next');
                        if (!cycleBtn) return;
                        const entries = cache?.get(spawnKey) ?? [];
                        if (entries.length <= 1) return;
                        e.stopPropagation();
                        displayIdx = (displayIdx + (cycleBtn.classList.contains('spawn-prev') ? -1 : 1) + entries.length) % entries.length;
                        const cd = el.querySelector('.leaflet-popup-content');
                        if (cd) {
                            cd.innerHTML = buildEnemyPopup(cache);
                            popup._updateLayout?.();
                            popup._updatePosition?.();
                        }
                    };
                    el.addEventListener('click', clickHandler);
                });
            };
            if (_enemySpawnCache) { bind(_enemySpawnCache); return; }
            _enemySpawnPromise.then(cache => { if (this.isPopupOpen()) bind(cache); });
        });

        marker.on('popupclose', () => {
            if (_activeRadiiMarker === marker) clearSpawnRadii();
        });

        marker._sgKey          = sgKey;
        marker._label          = buildTooltip(_enemySpawnCache);
        marker._origStyle      = { color: isKeyBearer ? '#c8a000' : g.color, weight: isKeyBearer ? 3.5 : 2.5, opacity: 1, fillOpacity: 0.85 };
        marker._spawn          = spawn;
        marker._info           = info;
        marker._spawnKey       = spawnKey;
        marker._naturalLatLng  = latlng;
        marker._naturalTooltip = buildTooltip(_enemySpawnCache);

        if (!g.sgMarkers[sgKey]) g.sgMarkers[sgKey] = [];
        g.sgMarkers[sgKey].push(marker);

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

function _expandGroupCore(g) {
    if (!g.detailsLayer) buildGroupDetails(g);
    const enemiesOn = document.getElementById('layer-enemies').checked;
    if (enemiesOn) g.detailsLayer.addTo(leafletMap);
    if (document.getElementById('layer-territory').checked && g.territoryRect)
        territoryLayer.addLayer(g.territoryRect);
    g.isExpanded = true;
    let topPx = g.pts[0][0], topPy = g.pts[0][1];
    for (const [px, py] of g.pts) { if (py > topPy) { topPy = py; topPx = px; } }
    g.labelMarker.setLatLng(xy(topPx, topPy + 10));
    g.labelMarker.setIcon(makeChipIcon(g.groupId, g.color, g.items.length, true, g.yOffset, g.isKeyBearerGroup, _groupHasBoss(g)));
    for (const [sgKey, markers] of Object.entries(g.sgMarkers)) {
        if (!_sgMarkers[sgKey]) _sgMarkers[sgKey] = [];
        _sgMarkers[sgKey].push(...markers);
    }
    applySubGroupFilter();
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
    _setGroupVisible(g, _activeSubGroupId === null ||
        g.items.some(({ spawn }) => _spawnSubGroupId(spawn) === _activeSubGroupId) ||
        (_activeSubGroupId === 1 && g.areaSpawn));
}

// ── SubGroup filter ────────────────────────────────────────────────────────────
const _spawnSubGroupId = (spawn) =>
    (spawn?.SubGroupNo == null || spawn.SubGroupNo === -1) ? 0 : spawn.SubGroupNo + 1;

let _activeSubGroupId = null;
let _availableSubGroups = [0];

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
    g.labelMarker.setOpacity(visible ? 1 : 0);
    const chipEl = g.labelMarker.getElement();
    if (chipEl) chipEl.style.pointerEvents = visible ? '' : 'none';
    if (g.detailsLayer) {
        for (const layer of g.detailsLayer.getLayers()) {
            if (layer._spawn) continue;
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
let _sgMarkers = {};
let _unhighlightTimer = null;
let _highlightedSet   = new Set();

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
    _clearHighlight();
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
let _activeRadiiMarker = null;

function clearSpawnRadii() {
    spawnRadiiLayer.clearLayers();
    _activeRadiiMarker = null;
}

function showSpawnRadii(marker) {
    if (!document.getElementById('layer-radii').checked) return;
    if (_activeRadiiMarker === marker) {
        clearSpawnRadii();
        return;
    }
    spawnRadiiLayer.clearLayers();
    _activeRadiiMarker = marker;

    const spawn = marker._spawn;
    const info  = marker._info;
    if (!spawn || !info) return;

    const spawnEntries = marker._spawnKey && _enemySpawnCache
        ? (_enemySpawnCache.get(marker._spawnKey) ?? []) : [];
    const hasEnemy = !_enemySpawnCache || spawnEntries.some(e => !!e.lv);
    if (!hasEnemy) { clearSpawnRadii(); return; }

    const emCode = spawnEntries[0]?.emCode ?? spawn.EmName ?? null;
    const radii  = emCode ? (emRadii[emCode] ?? null) : null;
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

let _radiiClickConsumed = false;
leafletMap.on('click', () => {
    if (_radiiClickConsumed) { _radiiClickConsumed = false; return; }
    clearSpawnRadii();
});

function spawnGroupColor(sg) {
    const hue = (sg * 137) % 360;
    const L = sg < 50 ? 0.80 : 0.72;
    const C = sg < 50 ? 0.10 : 0.17;
    return `oklch(${L} ${C} ${hue})`;
}

function groupBorderColor(groupId) {
    const hue = (groupId * 137) % 360;
    return `oklch(0.55 0.13 ${hue})`;
}

// ── Floor OBB test ────────────────────────────────────────────────────────────
function getEnemyFloor(worldX, worldY, worldZ, floorObbs) {
    for (const o of floorObbs) {
        const dx = worldX - o.cx;
        const dy = worldY - o.cy;
        const dz = worldZ - o.cz;
        const lx = dx * o.ax + dz * o.az;
        const lz = dx * o.bx + dz * o.bz;
        if (Math.abs(lx) <= o.ex && Math.abs(dy) <= o.ey && Math.abs(lz) <= o.ez)
            return o.floor_id;
    }
    return null;
}

function loadEnemySpawns(info, stid = null) {
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

    const byGroupId = new Map();
    for (const stageId of stagesToLoad) {
        const stageNo   = String(parseInt(stageId.slice(2), 10));
        const stageData = enemyPositions[stageNo];
        if (!stageData) continue;
        for (const [groupId, groupData] of Object.entries(stageData)) {
            const spawns         = groupData.spawns         ?? groupData;
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

    const centroidBuckets = new Map();
    for (const [groupId, { pts }] of byGroupId) {
        if (!pts.length) continue;
        const cx = pts.reduce((s, p) => s + p[0], 0) / pts.length;
        const cy = pts.reduce((s, p) => s + p[1], 0) / pts.length;
        const key = `${Math.round(cx)}:${Math.round(cy)}`;
        if (!centroidBuckets.has(key)) centroidBuckets.set(key, []);
        centroidBuckets.get(key).push(groupId);
    }

    for (const [groupId, { territory, keyBearerGroup, splitId, areaSpawn, priority, items, pts }] of byGroupId) {
        if (!pts.length) continue;
        const color = groupBorderColor(parseInt(groupId, 10));
        const cx = pts.reduce((s, p) => s + p[0], 0) / pts.length;
        const cy = pts.reduce((s, p) => s + p[1], 0) / pts.length;

        const bucketKey = `${Math.round(cx)}:${Math.round(cy)}`;
        const bucket    = centroidBuckets.get(bucketKey);
        const slotIdx   = bucket.indexOf(groupId);
        const yOffset   = 10 + slotIdx * 20;

        const g = { groupId, color, territory, isKeyBearerGroup: keyBearerGroup, splitId, areaSpawn, priority, items, pts,
                    centroid: { px: cx, py: cy }, yOffset,
                    labelMarker: null, detailsLayer: null, isExpanded: false, sgMarkers: {} };
        _groupStore.set(groupId, g);

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
const HIDDEN_LANDMARK_TYPES = new Set(['TYPE_TEXT', 'TYPE_WATER_LINE', 'TYPE_NONE']);

// ── Live server data (fetched at runtime) ─────────────────────────────────────
const _DEFAULT_GATHERING_URL = 'https://raw.githubusercontent.com/edelarrow/map-spawns/refs/heads/main/Normal%20Channels/GatheringItem.csv';
const _DEFAULT_SPAWNS_URL    = 'https://raw.githubusercontent.com/edelarrow/map-spawns/refs/heads/main/Normal%20Channels/EnemySpawn.json';
const _DEFAULT_SHOP_URL      = 'https://raw.githubusercontent.com/edelarrow/map-spawns/refs/heads/main/Normal%20Channels/Shop.json';
const _DEFAULT_SPECIAL_SHOP_URL = 'https://raw.githubusercontent.com/edelarrow/map-spawns/refs/heads/main/Normal%20Channels/SpecialShops.json';

function showSrcError(label) {
    let box = document.getElementById('src-errors');
    if (!box) {
        box = document.createElement('div');
        box.id = 'src-errors';
        box.style.cssText = 'font-size:0.75rem;padding:2px 8px;';
        const sidebar = document.getElementById('sidebar');
        const anchor  = sidebar.querySelector('#search-box') ?? sidebar.children[1];
        sidebar.insertBefore(box, anchor);
    }
    const item = document.createElement('div');
    item.style.cssText = 'display:flex;align-items:center;gap:5px;margin-bottom:3px;'
        + 'color:#f99;background:#2a0f0f;border-left:3px solid #c0392b;border-radius:2px;padding:3px 6px;';
    item.innerHTML = `⚠ <span style="flex:1"><b>${label}</b> failed to load</span>`
        + `<button data-action="dismiss" style="font-size:0.7rem;padding:1px 4px;cursor:pointer;`
        + `background:none;color:#aaa;border:none">✕</button>`;
    item.querySelector('[data-action="dismiss"]').addEventListener('click', () => item.remove());
    box.appendChild(item);
}

// Gathering items cache
let _gatherItemsCache = null;
const _gatherItemsPromise = fetch(_DEFAULT_GATHERING_URL)
    .then(r => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.text(); })
    .then(text => {
        const lines = text.split('\n');
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
                stageId: cols[iStage], groupId: cols[iGroup], posId: cols[iPos],
                itemId: parseInt(cols[iItem]), itemNum: parseInt(cols[iNum]),
                maxItemNum: parseInt(cols[iMax]), quality: parseInt(cols[iQual]),
                isHidden: cols[iHidden] === 'true' || cols[iHidden] === '1',
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

// Enemy spawn cache
let _enemySpawnCache = null;
let _dropsTablesMap = new Map();
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

        const dropsTables = {};
        for (const dt of (data.dropsTables ?? [])) { dropsTables[dt.id] = dt; _dropsTablesMap.set(dt.id, dt); }
        const result = new Map();
        for (let rawIdx = 0; rawIdx < (data.enemies?.length ?? 0); rawIdx++) {
            const e    = data.enemies[rawIdx];
            const key  = `${e[iStage]},${e[iGroup]},${e[iPosIdx]}`;
            const dtId = e[iDrops];
            const dt   = dtId != null && dtId >= 0 ? dropsTables[dtId] : null;
            const hexStr = e[iEnemyId];
            const emCode = hexStr ? `em${hexStr.slice(2).toLowerCase().padStart(6, '0')}` : null;
            const bloodOrbs = e[iBlood] ?? 0;
            const highOrbs  = e[iHigh]  ?? 0;
            const entry = {
                emCode,
                lv:          e[iLv]         ?? null,
                bloodOrbs,   highOrbs,
                spawnTime:   e[iSpawnTime]  ?? null,
                dropsTableId: dtId ?? -1,
                drops:        dt ? dt.items : [],
                scale:        e[iScale]      ?? 100,
                subGroupId:   e[iSubGroup]   ?? 0,
                namedId:      e[iNamed]      ?? 0,
                raidBossId:   e[iRaidBoss]   ?? 0,
                setType:      e[iSetType]    ?? 0,
                infection:    e[iInfection]  ?? 0,
                isBossGauge:  e[iIsBossG]    ?? false,
                isBossBGM:    e[iIsBossBGM]  ?? false,
                isAreaBoss:   e[iIsAreaBoss] ?? false,
                isManualSet:  e[iIsManualSet] ?? false,
                isBloodOrbEnemy: iIsBloodOrbEnemy >= 0 ? (e[iIsBloodOrbEnemy] ?? false) : bloodOrbs > 0,
                isHighOrbEnemy:  iIsHighOrbEnemy  >= 0 ? (e[iIsHighOrbEnemy]  ?? false) : highOrbs  > 0,
                exp:          e[iExp]         ?? 0,
                repopNum:     e[iRepopNum]    ?? 0,
                repopCount:   e[iRepopCount]  ?? 0,
                targetTypeId: e[iTargetType]  ?? 0,
                hmPreset:     e[iHmPreset]    ?? 0,
                startThink:   e[iStartThink]  ?? 0,
                montage:      e[iMontage]     ?? 0,
                ppDrop:       e[iPPDrop]      ?? 0,
            };
            if (result.has(key)) result.get(key).push(entry);
            else                 result.set(key, [entry]);
        }
        _enemySpawnCache = result;
        return result;
    })
    .catch(() => { showSrcError('Enemy Spawns'); _enemySpawnCache = new Map(); return _enemySpawnCache; });

// Shop cache
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

// Special shop cache
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
    3: 'Shop', 4: 'Item Shop', 5: 'Equipment Shop',
    6: 'Material Shop', 8: 'Weapon Shop', 9: 'Armor Shop',
    19: 'Orb Exchange (Crests)', 20: 'Orb Exchange (Materials)',
    57: 'Play Point Shop', 74: 'Adventure Pass Shop', 97: 'Bitterblack Shop',
};
const NPC_FUNC_COLORS = {
    3: '#ffd700', 4: '#4caf50', 5: '#2196f3', 6: '#ff9800',
    8: '#e91e63', 9: '#9c27b0', 19: '#00bcd4', 20: '#00bcd4',
    57: '#ffeb3b', 74: '#607d8b', 97: '#6a0020',
};
const WALLET_LABELS = {
    1: 'G', 2: 'R', 3: 'BO', 4: 'Tickets', 5: 'Gems',
    6: 'RP', 9: 'HO', 10: 'DP', 11: 'BP', 15: 'Dragon Marks',
};
const SHOP_TYPE_IDS = {
    Unknown: 0, Trinkets: 1,
    EmblemMedalExchangeRathniteFoothills: 2, EmblemMedalExchangeFeryanaWilderness: 3,
    EmblemMedalExchangeMegadosysPlateau: 4, EmblemMedalExchangeUrtecaMountains: 5,
    BitterBlackMaze: 6,
    MedalExchangeHidellPlains: 7, MedalExchangeBriaCoast: 8,
    MedalExchangeVerda: 9, MedalExchangeHarspuds: 10,
    MedalExchangeWideMoor: 11, MedalExchangeLimgom: 12,
    MedalExchangeCrumblingStones: 13, MedalExchangeGarlevMountain: 14,
    MedalExchangeOldSorMountain: 15, MedalExchangeCazhetteForest: 16,
    MedalExchangeTataru: 17, MedalExchangeTamaSpa: 18,
    MedalExchangeGransysOrchard: 19, MedalExchangeGustFront: 20,
    MedalExchangeVolatileVolcano: 21, MedalExchangeCanyon: 22,
    MedalExchangeWestFrontier: 23, ExtremeMission: 27,
};

// ── Gathering node colors/labels ───────────────────────────────────────────────
const GATHER_COLORS = {
    OM_GATHER_GRASS: '#4CAF50', OM_GATHER_FLOWER: '#E91E63', OM_GATHER_MUSHROOM: '#9C27B0',
    OM_GATHER_CLOTH: '#CE93D8', OM_GATHER_CRST_LV1: '#42A5F5', OM_GATHER_CRST_LV2: '#1E88E5',
    OM_GATHER_CRST_LV3: '#1565C0', OM_GATHER_CRST_LV4: '#0D47A1', OM_GATHER_JWL_LV1: '#FFEE58',
    OM_GATHER_JWL_LV2: '#FDD835', OM_GATHER_JWL_LV3: '#F9A825', OM_GATHER_TWINKLE: '#FFF9C4',
    OM_GATHER_TREE_LV1: '#A1887F', OM_GATHER_TREE_LV2: '#795548', OM_GATHER_TREE_LV3: '#5D4037',
    OM_GATHER_TREE_LV4: '#3E2723', OM_GATHER_SAND: '#FF9800', OM_GATHER_SHELL: '#FFCC80',
    OM_GATHER_WATER: '#00BCD4', OM_GATHER_ANTIQUE: '#FF5722', OM_GATHER_BOX: '#8D6E63',
    OM_GATHER_ALCHEMY: '#00BFA5', OM_GATHER_BOOK: '#78909C', OM_GATHER_ONE_OFF: '#B0BEC5',
    OM_GATHER_SHIP: '#29B6F6', OM_GATHER_KEY_LV1: '#EF9A9A', OM_GATHER_KEY_LV2: '#EF5350',
    OM_GATHER_KEY_LV3: '#C62828', OM_GATHER_KEY_LV4: '#B71C1C', OM_GATHER_TREA_IRON: '#E0E0E0',
    OM_GATHER_TREA_OLD: '#BCAAA4', OM_GATHER_TREA_TREE: '#A5D6A7', OM_GATHER_TREA_SILVER: '#CFD8DC',
    OM_GATHER_TREA_GOLD: '#FFD54F', OM_GATHER_CORPSE: '#546E7A', OM_GATHER_DRAGON: '#EF5350',
    CHEST_IRON: '#90A4AE', CHEST_BROWN: '#A1887F', CHEST_TREASURE: '#80CBC4',
    CHEST_BRONZE: '#FFAB40', CHEST_SILVER: '#E0E0E0', CHEST_GOLD: '#FFD700',
    CHEST_PURPLE: '#CE93D8', CHEST_ROUND: '#FFF59D', CHEST_SEALED_ORANGE: '#FF6F00',
    CHEST_SEALED_PURPLE: '#7B1FA2', CHEST_SEALED_PEARL: '#B2EBF2', CHEST_UNKNOWN: '#607D8B',
};

const GATHER_LABELS = {
    OM_GATHER_GRASS: 'Grass / Herb', OM_GATHER_FLOWER: 'Flower', OM_GATHER_MUSHROOM: 'Mushroom',
    OM_GATHER_CLOTH: 'Cloth / Fibre', OM_GATHER_CRST_LV1: 'Crystal (Lv1)', OM_GATHER_CRST_LV2: 'Crystal (Lv2)',
    OM_GATHER_CRST_LV3: 'Crystal (Lv3)', OM_GATHER_CRST_LV4: 'Crystal (Lv4)', OM_GATHER_JWL_LV1: 'Gemstone (Lv1)',
    OM_GATHER_JWL_LV2: 'Gemstone (Lv2)', OM_GATHER_JWL_LV3: 'Gemstone (Lv3)', OM_GATHER_TWINKLE: 'Sparkle Node',
    OM_GATHER_TREE_LV1: 'Lumber (Lv1)', OM_GATHER_TREE_LV2: 'Lumber (Lv2)', OM_GATHER_TREE_LV3: 'Lumber (Lv3)',
    OM_GATHER_TREE_LV4: 'Lumber (Lv4)', OM_GATHER_SAND: 'Sand', OM_GATHER_SHELL: 'Shell',
    OM_GATHER_WATER: 'Water', OM_GATHER_ANTIQUE: 'Antique', OM_GATHER_BOX: 'Box',
    OM_GATHER_ALCHEMY: 'Alchemy Node', OM_GATHER_BOOK: 'Book', OM_GATHER_ONE_OFF: 'One-off Node',
    OM_GATHER_SHIP: 'Maritime Gather', OM_GATHER_KEY_LV1: 'Locked Chest (Lv1)', OM_GATHER_KEY_LV2: 'Locked Chest (Lv2)',
    OM_GATHER_KEY_LV3: 'Locked Chest (Lv3)', OM_GATHER_KEY_LV4: 'Locked Chest (Lv4)',
    OM_GATHER_TREA_IRON: 'Treasure (Iron)', OM_GATHER_TREA_OLD: 'Treasure (Old)',
    OM_GATHER_TREA_TREE: 'Treasure (Wood)', OM_GATHER_TREA_SILVER: 'Treasure (Silver)',
    OM_GATHER_TREA_GOLD: 'Treasure (Gold)', OM_GATHER_CORPSE: 'Examine (Corpse)',
    OM_GATHER_DRAGON: 'Dragon Node', CHEST_IRON: 'Iron Chest', CHEST_BROWN: 'Brown Chest',
    CHEST_TREASURE: 'Treasure Chest', CHEST_BRONZE: 'Bronze Chest', CHEST_SILVER: 'Silver Chest',
    CHEST_GOLD: 'Gold Chest', CHEST_PURPLE: 'Purple Chest', CHEST_ROUND: 'Small Round Chest',
    CHEST_SEALED_ORANGE: 'Sealed Chest (Orange)', CHEST_SEALED_PURPLE: 'Sealed Chest (Purple)',
    CHEST_SEALED_PEARL: 'Pearlescent Chest', CHEST_UNKNOWN: 'Chest',
};

function loadGatherPoints(info, stid = null) {
    gatherLayer.clearLayers();
    _gatherMarkerByKey.clear();
    if (!info.stages?.length) return;

    const floorObbs    = info.floor_obbs ?? null;
    const filterByFloor = floorObbs !== null;
    const stagesToLoad  = (stid && info.stages.includes(stid)) ? [stid] : info.stages;

    for (const stageId of stagesToLoad) {
        const stageNo    = String(parseInt(stageId.slice(2), 10));
        const nodes      = gatherPoints[stageNo];
        if (!nodes) continue;
        const serverStid = stageIds[stageNo];

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
                            '<div style="margin-top:6px;display:flex;flex-direction:column;gap:4px">' +
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
                                return `<div style="display:flex;align-items:flex-start;gap:7px">`
                                    + icon
                                    + `<div style="display:flex;flex-direction:column;gap:1px;min-width:0">${nameLink}${meta}</div>`
                                    + `</div>`;
                            }).join('') +
                            '</div>';
                    }
                }
                return `${badge}${typeLine}${coordLine}${itemsHtml}`;
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

            if (csvKey) {
                marker.on('popupopen', function() {
                    const self = this;
                    const attach = (gatherMap) => {
                        self.getPopup().setContent(buildGatherPopup(gatherMap));
                        self.getPopup().update();
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

    const floorObbs    = info.floor_obbs ?? null;
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

            const questLine  = node.questName ? `<br><span style="color:#c97a00;font-size:10px;font-style:italic">${node.questName.replace(/\n/g, ' ')}</span>` : '';
            const hitsLine   = node.hitNum != null ? `<br><span style="font-size:10px;color:#888">${node.hitNum} hit${node.hitNum !== 1 ? 's' : ''} to destroy</span>` : '';
            const condLine   = (node.questNo || node.layoutFlagNo)
                ? `<br><span style="font-size:10px;color:#888">${node.questNo ? `Quest: ${node.questNo}` : ''}${node.questNo && node.layoutFlagNo ? ' &nbsp;·&nbsp; ' : ''}${node.layoutFlagNo ? `LayoutFlag: ${node.layoutFlagNo}` : ''}</span>` : '';
            const omLine     = node.unitId != null ? `<br><span style="font-size:10px;color:#666">OMID: ${node.unitId}${node.omName ? ` &nbsp;(${node.omName})` : ''}</span>` : '';
            const coordLine  = `<br><span style="font-size:11px;color:#555">X:&nbsp;${node.x.toFixed(0)}&nbsp; Y:&nbsp;${node.y.toFixed(0)}&nbsp; Z:&nbsp;${node.z.toFixed(0)}</span>`;
            const groupLine  = `<br><span style="color:#666;font-size:10px">Group ${node.groupId} · pos ${node.posId}</span>`;
            const popupHtml  = `<span style="font-weight:bold;color:#e65c00">Destroyable Object</span>${questLine}${hitsLine}${condLine}${omLine}${groupLine}${coordLine}`;
            const tooltipText = node.questName ? `Destroyable Object — ${node.questName}` : `Destroyable Object (group ${node.groupId})`;

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

    const floorObbs    = info.floor_obbs ?? null;
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
                    const entry  = itemNames[String(it.ItemId)];
                    const iconNo = entry?.iconNo;
                    const iconFile = iconNo != null ? `ii${String(iconNo).padStart(6, '0')}.png` : null;
                    return iconFile && _iconIdSet.has(iconNo)
                        ? `<img src="images/icons/small/${iconFile}" width="28" height="28" style="vertical-align:middle;margin-right:6px;image-rendering:pixelated">`
                        : `<span style="display:inline-block;width:28px;margin-right:6px"></span>`;
                };

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
                .bindPopup(buildShopPopup(_shopCache), { minWidth: 280 })
                .bindTooltip(`${npcName} — ${funcLabel}`, { direction: 'top', offset: [0, -10] })
                .addTo(npcShopLayer);
            _shopMarkerByNpcId.set(`${stageNo}:${npc.NpcId}`, marker);

            marker.on('popupopen', function() {
                const self = this;
                const rebuild = (cache) => {
                    self.getPopup().setContent(buildShopPopup(cache));
                    self.getPopup().update();
                };
                if (_shopCache) { rebuild(_shopCache); return; }
                _shopPromise.then(cache => { if (self.isPopupOpen()) rebuild(cache); });
            });
        }
    }
}

function loadSpecialShops(info, stid = null) {
    specialShopLayer.clearLayers();
    if (!info.stages?.length) return;

    const floorObbs    = info.floor_obbs ?? null;
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
            const shopType = npc.ShopTypeName ?? npc.ShopType;
            const typeId   = npc.ShopType ?? SHOP_TYPE_IDS[shopType] ?? 0;
            const npcName  = npcNames[String(npc.NpcId)]?.name ?? npc.NpcName ?? `NPC #${npc.NpcId}`;
            const color    = '#c084fc';

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

            const buildSpecialShopPopup = (cache) => {
                const header = `<div style="display:flex;align-items:center;flex-wrap:wrap;gap:2px;padding-bottom:6px;margin-bottom:6px;border-bottom:1px solid #ddd">` +
                    `<span style="display:inline-block;padding:1px 7px;border-radius:3px;background:#c084fc;color:#111;font-weight:bold;font-size:11px;">Appraisals</span>` +
                    `<strong style="font-size:12px;margin-left:6px">${npcName}</strong>` +
                    `<span style="color:#888;font-size:11px;margin-left:6px">${shopType} (ID: ${typeId})</span>` +
                    `</div>`;

                if (!cache) return header + `<div style="padding:10px;color:#888;font-size:11px">Loading…</div>`;
                const shopData = cache.get(shopType);
                if (!shopData?.categories?.length) return header + `<div style="padding:10px;color:#888;font-size:11px">No data for "${shopType}"</div>`;

                const { categories } = shopData;
                let html = header + `<div style="max-height:360px;overflow-y:auto">`;
                for (const cat of categories) {
                    html += `<div style="font-size:11px;font-weight:bold;color:#333;background:#f0f0f0;padding:3px 5px;border-bottom:1px solid #ddd;position:sticky;top:0">${cat.label}</div>`;
                    for (const ap of cat.appraisals) {
                        const costs = ap.base_items.map(bi =>
                            `<span style="display:inline-flex;align-items:center;white-space:nowrap">${itemIcon(bi.item_id)}<span style="font-size:11px">×${bi.amount}</span></span>`
                        ).join(' + ');
                        const isLottery = ap.pool.length > 1;
                        const reward = isLottery
                            ? `<span style="font-size:10px;font-weight:bold;color:#7030b0">${ap.pool.length} possible rewards</span>`
                            : ap.pool.map(pi => `<span style="display:inline-flex;align-items:center">${itemIcon(pi.item_id)}<span style="font-size:11px">${itemRef(pi.item_id, pi.name)} ×${pi.amount}</span></span>`).join('');
                        html += `<div style="display:flex;align-items:center;gap:5px;padding:4px 6px;border-bottom:1px solid rgba(0,0,0,0.06)">` +
                            `<div style="display:inline-flex;align-items:center;gap:2px;flex-shrink:0">${costs}</div>` +
                            `<span style="color:#bbb;flex-shrink:0">→</span>` +
                            `<div style="flex:1;min-width:0">${reward}</div>` +
                            `</div>`;
                    }
                }
                html += `</div>`;
                return html;
            };

            const icon = L.divIcon({
                className: '',
                html: `<div style="width:12px;height:12px;background:${color};border:2px solid #111;transform:rotate(45deg);box-shadow:0 0 4px rgba(192,132,252,0.7);"></div>`,
                iconSize:    [12, 12],
                iconAnchor:  [6, 6],
                popupAnchor: [0, -10],
            });
            const marker = L.marker(latlng, { icon })
                .bindPopup(buildSpecialShopPopup(_specialShopCache), { minWidth: 380 })
                .bindTooltip(`${npcName} — Appraisals (${shopType})`, { direction: 'top', offset: [0, -10] })
                .addTo(specialShopLayer);

            marker.on('popupopen', function() {
                const self = this;
                const rebuild = (cache) => {
                    self.getPopup().setContent(buildSpecialShopPopup(cache));
                    self.getPopup().update();
                };
                if (_specialShopCache) { rebuild(_specialShopCache); return; }
                _specialShopPromise.then(cache => { if (self.isPopupOpen()) rebuild(cache); });
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
            color, fillColor: color, fillOpacity: 0.85, radius: 6, weight: 1.5,
        })
        .bindTooltip(label, { permanent: false, direction: 'top', offset: [0, -6] })
        .addTo(landmarkLayer);
    }
}

// ── Stage connection markers ───────────────────────────────────────────────────
function arrivalView(srcMap, destMap, srcStageNo = null, zoom = 1.5) {
    const destInfo = mapParams[destMap];
    if (!destInfo) return null;
    const destConns = connectionData[destMap] || [];
    const rev = destConns.find(c =>
        c.to_map === srcMap && c.x != null && c.z != null &&
        (srcStageNo == null || c.to_stage === srcStageNo)
    ) ?? destConns.find(c => c.to_map === srcMap && c.x != null && c.z != null);
    if (!rev) return null;
    return { zoom, center: worldToPixel(rev.x, rev.z, destInfo) };
}

function loadConnections(mapName, info) {
    connectionLayer.clearLayers();
    const exitsPanel = document.getElementById('exits-panel');
    const exitsList  = document.getElementById('exits-list');
    exitsList.innerHTML = '';

    const allEntries = connectionData[mapName];
    if (!allEntries) { exitsPanel.style.display = 'none'; return; }

    const stid = currentStageName();
    const activeStageNo = stid ? parseInt(stid.slice(2), 10) : null;
    const stageFiltered = allEntries.filter(c =>
        c.from_stage == null || activeStageNo == null || c.from_stage === activeStageNo
    );

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

        if (conn.x == null || conn.z == null) {
            unpositioned.push({ navMap, hasMap, stageId, destName });
            continue;
        }
        if (info.floor_obbs) {
            const floor = getEnemyFloor(conn.x, conn.y ?? 0, conn.z, info.floor_obbs);
            if (floor !== null && floor !== currentLayer) continue;
        }

        const latlng = worldToPixel(conn.x, conn.z, info);
        const color  = hasMap ? '#ff6b35' : '#666666';

        const icon = L.divIcon({
            className: '',
            html: `<div style="width:14px;height:14px;background:${color};border:2px solid #fff;border-radius:3px;transform:rotate(45deg);box-shadow:0 0 4px rgba(0,0,0,0.7);"></div>`,
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

    if (unpositioned.length) {
        for (const { navMap, hasMap, stageId, destName } of unpositioned) {
            const li = document.createElement('div');
            li.style.cssText = 'padding:2px 0;font-size:0.78rem;';
            if (hasMap) {
                const a = document.createElement('span');
                a.textContent = destName;
                a.style.cssText = 'color:#ff6b35;cursor:pointer;text-decoration:underline dotted;';
                a.addEventListener('click', () => navigateTo(navMap, stageId, arrivalView(mapName, navMap, null)));
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
const GRID_UNIT  = 10;
const GRID_MAJOR = 50;

function pixelToGrid(px, py, imgH) {
    return [Math.floor(px / GRID_UNIT), Math.floor((imgH - py) / GRID_UNIT)];
}

function loadGrid(info) {
    gridLayer.clearLayers();
    const lineStyle   = { color: '#ffffff', weight: 0.5, opacity: 0.2, interactive: false };
    const maxGX = Math.floor(info.img_width  / GRID_UNIT);
    const maxGY = Math.floor(info.img_height / GRID_UNIT);

    for (let gx = 0; gx <= maxGX; gx += GRID_MAJOR) {
        const px = gx * GRID_UNIT;
        L.polyline([xy(px, 0), xy(px, info.img_height)], lineStyle).addTo(gridLayer);
        L.marker(xy(px, info.img_height), {
            icon: L.divIcon({ className: 'grid-label', html: `${gx}`, iconSize: [40, 14], iconAnchor: [-2, 7] }),
            interactive: false,
        }).addTo(gridLayer);
    }
    for (let gy = 0; gy <= maxGY; gy += GRID_MAJOR) {
        const py = info.img_height - gy * GRID_UNIT;
        L.polyline([xy(0, py), xy(info.img_width, py)], lineStyle).addTo(gridLayer);
        L.marker(xy(0, py), {
            icon: L.divIcon({ className: 'grid-label', html: `${gy}`, iconSize: [40, 14], iconAnchor: [42, 7] }),
            interactive: false,
        }).addTo(gridLayer);
    }
}

// ── Floor selector ────────────────────────────────────────────────────────────
let currentLayer = 0;
let _currentFloorObbs = null;

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

function buildTileLayerSelector(info) {
    const el = document.getElementById('tile-layer-selector');
    if (!el) return;
    el.innerHTML = '';
    const tlImages = info.tile_layer_images;
    if (!tlImages || Object.keys(tlImages).length === 0) return;

    const makeBtn = (label, key, imgFile) => {
        const btn = document.createElement('button');
        btn.textContent = label;
        if (key === null) btn.classList.add('active');
        btn.addEventListener('click', () => {
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
    overlay.addEventListener('click', e => { if (e.target === overlay) overlay.style.display = 'none'; });
    document.addEventListener('keydown', e => { if (e.key === 'Escape') overlay.style.display = 'none'; });

    let _activeLayer = null;
    let _compositeGen = 0;

    const layerNav = document.createElement('div');
    layerNav.style.cssText = 'display:none;gap:4px;align-items:center;cursor:default';
    overlay.insertBefore(layerNav, nav);

    function showLayer(piece, pieceIdx, layer) {
        _activeLayer = layer;
        const gen = ++_compositeGen;
        layerNav.querySelectorAll('button').forEach(btn => {
            const active = String(btn.dataset.layer) === String(layer);
            btn.style.background  = active ? '#4a90d9' : '#0f3460';
            btn.style.borderColor = active ? '#4a90d9' : '#1a4a7a';
            btn.style.color       = active ? '#fff'    : '#ccd';
        });

        if (layer !== null) {
            canvas.style.display = 'none';
            img.style.display    = '';
            img.src = `images/maps/${piece.model}_l${layer}.png`;
            title.textContent = `${pieceIdx}: ${piece.model}_l${layer}.png`;
            return;
        }

        if (piece.has_merged) {
            canvas.style.display = 'none';
            img.style.display    = '';
            img.src = `images/maps/${piece.model}_merged.png`;
            title.textContent = `${pieceIdx}: ${piece.model}_merged.png`;
            return;
        }

        const allLayers = piece.layers ?? [piece.layer ?? 0];
        title.textContent = `${pieceIdx}: ${piece.model} (composite)`;
        Promise.all(allLayers.map(lyr => new Promise((res, rej) => {
            const i = new Image();
            i.onload = () => res(i);
            i.onerror = () => rej(new Error(`missing l${lyr}`));
            i.src = `images/maps/${piece.model}_l${lyr}.png`;
        }))).then(imgs => {
            if (_compositeGen !== gen) return;
            canvas.width  = imgs[0].naturalWidth;
            canvas.height = imgs[0].naturalHeight;
            const ctx = canvas.getContext('2d');
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            for (const i of imgs) ctx.drawImage(i, 0, 0);
            img.style.display    = 'none';
            canvas.style.display = '';
        }).catch(() => {
            if (_compositeGen !== gen) return;
            canvas.style.display = 'none';
            img.style.display    = '';
            img.src = `images/maps/${piece.model}_l${piece.layer ?? 0}.png`;
        });
    }

    function showPiece(piece, idx) {
        const allLayers = piece.layers ?? [piece.layer ?? 0];
        if (allLayers.length > 1) {
            layerNav.innerHTML = '';
            layerNav.style.display = 'flex';
            const label = document.createElement('span');
            label.textContent = 'Layers:';
            label.style.cssText = 'color:#888;font-family:monospace;font-size:0.75rem';
            layerNav.appendChild(label);
            const btnC = document.createElement('button');
            btnC.textContent = 'composite';
            btnC.dataset.layer = 'null';
            btnC.style.cssText = 'background:#0f3460;color:#ccd;border:1px solid #1a4a7a;border-radius:4px;padding:2px 7px;font-size:0.72rem;font-family:monospace;cursor:pointer';
            btnC.addEventListener('click', e => { e.stopPropagation(); showLayer(piece, idx, null); });
            layerNav.appendChild(btnC);
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
        nav.querySelectorAll('button').forEach((btn, i) => {
            btn.style.background  = i === idx ? '#4a90d9' : '#0f3460';
            btn.style.borderColor = i === idx ? '#4a90d9' : '#1a4a7a';
            btn.style.color       = i === idx ? '#fff'    : '#ccd';
        });
        showLayer(piece, idx, null);
    }

    return {
        show(pieces, idx) {
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
        const png_y = piece.pixel_y_entrance;
        const py = info.img_height - png_y;
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

    const stid = currentStageName();
    const baseName = stid ? stageLabel(info, stid) : (info.name_en ? splitPascalCase(info.name_en) : mapName);
    const sidNum   = stid ? info.stage_ids?.[stid] : null;
    const sidStr   = sidNum != null ? ` - sid${String(sidNum).padStart(4, '0')}` : '';
    const title = baseName + ` (${stid ?? mapName}${sidStr})`;
    document.getElementById('map-title').textContent = title;
    document.title = `${title} — DDON Maps`;

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

    buildFloorSelector(info);
    buildTileLayerSelector(info);

    loadGrid(info);
    loadPdBoundaries(info);
    loadStageLabels(info);
    loadLandmarks(mapName, info);
    loadConnections(mapName, info);
    loadGatherPoints(info, currentStageName());
    loadNpcShops(info, currentStageName());
    loadSpecialShops(info, currentStageName());
    loadBreakTargets(info, currentStageName());

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

    if (window._setCurrentInfo) window._setCurrentInfo(info);
}

// ── Spot search panel ─────────────────────────────────────────────────────────
let _spotIndex        = [];
let _spotHighlights   = [];
let _spotOpenedGroup  = null;
const _spotHlLayer  = L.layerGroup().addTo(leafletMap);

function _clearSpotHighlights() {
    for (const m of _spotHighlights) _spotHlLayer.removeLayer(m);
    _spotHighlights = [];
    for (const el of _spotHlChipEls) el.classList.remove('spot-hl-chip');
    _spotHlChipEls = [];
}

let _spotHlChipEls = [];

function _addSpotHighlight(latlng) {
    const icon = L.divIcon({ className: 'spot-hl-outer', html: '<div class="spot-hl"></div>', iconSize: [22, 22], iconAnchor: [11, 11] });
    _spotHighlights.push(L.marker(latlng, { icon, interactive: false }).addTo(_spotHlLayer));
}

function _addChipHighlight(g) {
    const el = g.labelMarker.getElement()?.querySelector('.group-chip');
    if (!el) return;
    el.classList.add('spot-hl-chip');
    _spotHlChipEls.push(el);
}

function _resolveSpotLatLng(item) {
    const g = item.groupId ? _groupStore.get(item.groupId) : null;
    if (!g) return item.latlng;
    if (!g.isExpanded) return g.labelMarker.getLatLng();
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
        const focusLatLng = targetMarker?.getLatLng() ?? target.latlng;
        leafletMap.flyTo(focusLatLng, Math.max(leafletMap.getZoom(), 2), { duration: 0.4 });
        _clearSpotHighlights();
        _addSpotHighlight(focusLatLng);
        if (targetMarker) setTimeout(() => targetMarker.openPopup(), 450);
    } else if (target.type === 'item' && target.source === 'enemy' && target.groupId) {
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

function buildSpotIndex(info) {
    _spotIndex = [];
    if (!info.stages?.length) return;

    const stid = currentStageName();
    const stagesToLoad = (stid && info.stages.includes(stid)) ? [stid] : info.stages;

    for (const stageId of stagesToLoad) {
        const stageNo = String(parseInt(stageId.slice(2), 10));
        const serverStageId = stageIds[stageNo];
        const cache = _enemySpawnCache;

        const groups = enemyPositions[stageNo];
        if (groups) {
            for (const [groupId, groupData] of Object.entries(groups)) {
                const spawns = groupData.spawns ?? groupData;
                if (!Array.isArray(spawns) || !spawns.length) continue;
                for (let i = 0; i < spawns.length; i++) {
                    const s = spawns[i];
                    const posLatlng = worldToPixel(s.Position.x, s.Position.z, info);
                    const spawnKey  = serverStageId != null ? `${serverStageId},${groupId},${s.posIdx ?? i}` : null;
                    if (cache && spawnKey) {
                        const byEmCode = new Map();
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
                                previewLines: [`<b>${displayName}</b>`, `Code: ${emCode}`, `Group ${groupId}`],
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

        const nodes = gatherPoints[stageNo];
        if (nodes) {
            for (const node of nodes) {
                const label = GATHER_LABELS[node.type] ?? node.type.replace(/^(OM_GATHER_|CHEST_)/, '').replace(/_/g, ' ');
                _spotIndex.push({
                    type: 'gather', name: label, gatherType: node.type,
                    searchText: label.toLowerCase(),
                    latlng: worldToPixel(node.x, node.z, info),
                    nodeKey: `${stageNo}:${node.groupId}:${node.posId}`,
                    worldPos: { x: node.x, y: node.y, z: node.z },
                    previewLines: [`<b>${label}</b>`, `X: ${node.x.toFixed(0)}  Z: ${node.z.toFixed(0)}`],
                    stageId,
                });
            }
        }

        if (cache && serverStageId != null) {
            for (const [groupId, groupData] of Object.entries(groups ?? {})) {
                const spawns = groupData.spawns ?? groupData;
                if (!Array.isArray(spawns) || !spawns.length) continue;
                for (let i = 0; i < spawns.length; i++) {
                    const s = spawns[i];
                    const spawnKey = `${serverStageId},${groupId},${s.posIdx ?? i}`;
                    const posLatlng = worldToPixel(s.Position.x, s.Position.z, info);
                    const seen = new Set();
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
                                type: 'item', source: 'enemy', name: itemName,
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

        if (_gatherItemsCache && serverStageId != null) {
            for (const node of (gatherPoints[stageNo] ?? [])) {
                const csvKey  = `${serverStageId},${node.groupId},${node.posId}`;
                const nodeItems = _gatherItemsCache.get(csvKey) ?? [];
                const latlng  = worldToPixel(node.x, node.z, info);
                const nodeLabel = GATHER_LABELS[node.type] ?? node.type.replace(/^(OM_GATHER_|CHEST_)/, '').replace(/_/g, ' ');
                for (const it of nodeItems) {
                    if (it.isHidden) continue;
                    const itemName = itemNames[String(it.itemId)]?.name ?? `Item #${it.itemId}`;
                    const qty = it.maxItemNum > it.itemNum ? `×${it.itemNum}–${it.maxItemNum}` : `×${it.itemNum}`;
                    const pct = it.dropChance < 1 ? ` (${Math.round(it.dropChance * 100)}%)` : '';
                    _spotIndex.push({
                        type: 'item', source: 'gather', name: itemName,
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
                        type: 'item', source: 'shop', name: itemName,
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

function _rebuildSpotIndex() {
    if (_currentMapInfo) {
        buildSpotIndex(_currentMapInfo);
        if (document.getElementById('spot-panel')?.classList.contains('open')) _runSpotSearch();
    }
}
_enemySpawnPromise.then(() => {
    _rebuildSpotIndex();
    for (const g of _groupStore.values()) {
        if (_groupHasBoss(g)) {
            g.labelMarker.setIcon(makeChipIcon(g.groupId, g.color, g.items.length, g.isExpanded, g.yOffset, g.isKeyBearerGroup, true));
        }
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
_gatherItemsPromise.then(_rebuildSpotIndex).catch(() => {});
_shopPromise.then(_rebuildSpotIndex).catch(() => {});

function _runSpotSearch() {
    const query     = (document.getElementById('spot-search-input')?.value ?? '').trim().toLowerCase();
    const filter    = document.querySelector('.spot-tab.active')?.dataset.filter ?? 'enemy';
    const resultsEl = document.getElementById('spot-results');
    if (!resultsEl) return;

    _clearSpotHighlights();

    const matches = _spotIndex.filter(e => {
        if (filter === 'enemy'  && e.type !== 'enemy')  return false;
        if (filter === 'gather' && e.type !== 'gather') return false;
        if (filter === 'item'   && e.type !== 'item')   return false;
        return !query || e.searchText.includes(query);
    });

    if (!matches.length) {
        resultsEl.innerHTML = `<div class="spot-empty">No matches for <em>${query}</em>.</div>`;
        return;
    }

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
            let idx = -1;
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
            row.addEventListener('click', e => {
                if (e.target.closest('.spot-nav')) return;
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
                if (item.groupId && !grp) continue;
                if (grp && !grp.isExpanded) {
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

    document.addEventListener('keydown', e => {
        if ((e.ctrlKey || e.metaKey) && e.key === 'f' && !e.shiftKey && _loadedMapName) {
            e.preventDefault();
            panel.classList.contains('open') ? (input.focus(), input.select()) : openPanel();
        }
        if (e.key === 'Escape' && panel.classList.contains('open')) closePanel();
    });
})();

// ── Coordinate readout ────────────────────────────────────────────────────────
(function () {
    const el = document.getElementById('coord-display');
    if (!el) return;

    let currentInfo = null;
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
                if (png_y >= p.pixel_y_start && png_y <= p.pixel_y_entrance) { piece = p; break; }
            }
            wz = piece.connect_z + (png_y - piece.pixel_y_entrance_v) / currentInfo.scale;
        } else {
            const scaleZ = currentInfo.scale_z ?? currentInfo.scale;
            wz = ((currentInfo.img_height - currentInfo.center_y) - py) / scaleZ;
        }
        const [gx, gy] = pixelToGrid(px, py, currentInfo.img_height);
        el.textContent = `(${gx}, ${gy})   world (${wx.toFixed(0)}, ${wz.toFixed(0)})`;
    });

    leafletMap.on('mouseout', () => { el.textContent = ''; });
})();

// ── Init ──────────────────────────────────────────────────────────────────────
if (!location.hash || location.hash === '#') {
    history.replaceState(null, '', '#field000_m00:st0100');
}
buildSidebar();
loadMap(currentMapName());
