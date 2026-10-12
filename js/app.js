/* ─────────────────────────────────────────────────────────
   app.js  –  Core application: routing, state, export
   ─────────────────────────────────────────────────────────*/

const SizingApp = (() => {

    /* ── State ─────────────────────────────────────────── */
    const state = {
        activeModule : 'system-landscape',
        moduleData   : {},           // key → { rows, fileName, headers, analysis }
        completed    : new Set(),
        locked       : new Set([     // modules disabled until System Landscape HW review is reached
            'tech-workload-app', 'tech-workload-db', 'app-workload', 'sizing-projection',
        ]),
    };

    /* ── Module registry ────────────────────────────────── */
    const MODULES = {
        'system-landscape'  : () => SystemLandscapeModule,
        'tech-workload-app' : () => TechWorkloadAppModule,
        'tech-workload-db'  : () => TechWorkloadDbModule,
        'app-workload'      : () => AppWorkloadModule,
        'sizing-projection' : () => SizingProjectionModule,
    };

    const MODULE_ORDER = [
        'system-landscape',
        'tech-workload-app',
        'tech-workload-db',
        'app-workload',
        'sizing-projection',
    ];

    /* ── Public API ─────────────────────────────────────── */
    const api = {

        init() {
            // Nav clicks
            document.querySelectorAll('.nav-item').forEach(el => {
                el.addEventListener('click', () => api.navigateTo(el.dataset.module));
            });

            // Lock downstream modules visually until System Landscape is reviewed
            api._applyLockState();

            // Header buttons
            document.getElementById('btn-new-analysis')
                .addEventListener('click', api.confirmReset);

            document.getElementById('btn-save-session')
                ?.addEventListener('click', api.saveSession);

            const loadBtn   = document.getElementById('btn-load-session');
            const loadInput = document.getElementById('session-file-input');
            loadBtn?.addEventListener('click', () => loadInput?.click());
            loadInput?.addEventListener('change', e => {
                const f = e.target.files && e.target.files[0];
                if (f) api.loadSession(f);
                e.target.value = '';   // allow re-selecting the same file later
            });

            // Render first module
            api.navigateTo('system-landscape');
        },

        navigateTo(moduleId) {
            if (!MODULES[moduleId]) return;
            if (state.locked.has(moduleId)) return;   // blocked while locked

            state.activeModule = moduleId;

            // Clear any module sub-navigation from the previous screen
            api.clearSubNav();

            // Update nav highlights
            document.querySelectorAll('.nav-item').forEach(el => {
                el.classList.toggle('active', el.dataset.module === moduleId);
            });

            // Render into the module mount (keeps the sticky progress bar intact)
            const mod = MODULES[moduleId]();
            const area = document.getElementById('content-area');
            const mount = document.getElementById('module-mount') || area;
            mount.innerHTML = mod.render();
            mod.init(api);

            // Scroll content to top
            area.scrollTop = 0;
        },

        /* ── Module sub-navigation (dynamic jump list under a module) ──
           items = [{ id, label, onClick }]. Rendered as a nested list beneath
           the given module's sidebar item. activeId highlights one entry. */
        setSubNav(moduleId, items, activeId) {
            const navItem = document.querySelector(`.nav-item[data-module="${moduleId}"]`);
            if (!navItem) return;
            let sub = document.getElementById('subnav-' + moduleId);
            if (!sub) {
                sub = document.createElement('ul');
                sub.className = 'nav-sublist';
                sub.id = 'subnav-' + moduleId;
                navItem.insertAdjacentElement('afterend', sub);
            }
            api._subNavHandlers = api._subNavHandlers || {};
            api._subNavHandlers[moduleId] = {};
            sub.innerHTML = (items || []).map((it, i) =>
                `<li class="nav-subitem${it.id === activeId ? ' active' : ''}${it.step ? ' nav-substep' : ''}${it.sub ? ' nav-subchild' : ''}" data-subid="${it.id}">
                    <span class="nav-subdot"></span><span class="nav-subtext">${it.label}</span>
                 </li>`).join('');
            (items || []).forEach(it => { api._subNavHandlers[moduleId][it.id] = it.onClick; });
            sub.querySelectorAll('.nav-subitem').forEach(li => {
                li.addEventListener('click', () => {
                    const h = api._subNavHandlers[moduleId] && api._subNavHandlers[moduleId][li.dataset.subid];
                    if (h) h();
                });
            });
        },

        /* Highlight one sub-nav entry (without rebuilding the list). */
        setSubNavActive(moduleId, activeId) {
            const sub = document.getElementById('subnav-' + moduleId);
            if (!sub) return;
            sub.querySelectorAll('.nav-subitem').forEach(li =>
                li.classList.toggle('active', li.dataset.subid === activeId));
        },

        /* Remove all module sub-navigation lists. */
        clearSubNav() {
            document.querySelectorAll('.nav-sublist').forEach(el => el.remove());
            api._subNavHandlers = {};
        },

        /* Unlock the given module ids (removes disabled nav state) */
        unlockModules(ids) {
            (ids || []).forEach(id => state.locked.delete(id));
            api._applyLockState();
        },

        /* Reflect state.locked onto the sidebar nav items */
        _applyLockState() {
            document.querySelectorAll('.nav-item').forEach(el => {
                el.classList.toggle('disabled', state.locked.has(el.dataset.module));
            });
        },

        /* Mark a module as completed and update progress bar */
        markCompleted(moduleId) {
            state.completed.add(moduleId);
            document.querySelector(`[data-module="${moduleId}"]`)
                ?.classList.add('completed');
            api._updateProgress();
        },

        setModuleData(moduleId, data) {
            state.moduleData[moduleId] = data;
        },

        getModuleData(moduleId) {
            return state.moduleData[moduleId] || null;
        },

        getAllData() {
            return state.moduleData;
        },

        /* ── System Landscape host variables (shared across modules) ──
           Returns the per-host records saved by Module 01:
           { hostname, cpuType, ghz, cores, sapsPerCore, totalSaps,
             memoryMB, component, layer }                              */
        getHostVariables() {
            const d = state.moduleData['system-landscape'];
            return (d && d.hostVariables) || [];
        },

        /* Look up a single landscape host by hostname (case-insensitive,
           tolerant of FQDN vs short-name mismatches). Returns null if none. */
        findHost(hostname) {
            if (!hostname) return null;
            const hv = api.getHostVariables();
            const needle = String(hostname).trim().toLowerCase();
            if (!needle) return null;
            // 1. exact match
            let hit = hv.find(h => (h.hostname || '').toLowerCase() === needle);
            if (hit) return hit;
            // 2. match on the short name (strip domain) on either side
            const short = needle.split('.')[0];
            hit = hv.find(h => (h.hostname || '').toLowerCase().split('.')[0] === short);
            if (hit) return hit;
            // 3. one is a prefix of the other (host vs host_SID_nn remnants)
            hit = hv.find(h => {
                const hn = (h.hostname || '').toLowerCase().split('.')[0];
                return hn && (hn.startsWith(short) || short.startsWith(hn));
            });
            return hit || null;
        },

        getCompletedModules() {
            return state.completed;
        },

        nextModule() {
            const idx = MODULE_ORDER.indexOf(state.activeModule);
            if (idx < MODULE_ORDER.length - 1) {
                api.navigateTo(MODULE_ORDER[idx + 1]);
            }
        },

        /* ── Reset ──────────────────────────────────────── */
        confirmReset() {
            if (!confirm('Start a new analysis? All uploaded data will be cleared.')) return;
            state.moduleData = {};
            state.completed.clear();
            state.locked = new Set([
                'tech-workload-app', 'tech-workload-db', 'app-workload', 'sizing-projection',
            ]);
            // Clear per-module in-memory state (singletons survive navigation,
            // so an explicit reset is needed to wipe parsed/rehydrated views).
            Object.values(MODULES).forEach(getMod => {
                const mod = getMod();
                if (mod && typeof mod.resetState === 'function') mod.resetState();
            });
            ChartFontSizeManager.reset();
            document.querySelectorAll('.nav-item').forEach(el => {
                el.classList.remove('completed');
            });
            api._applyLockState();
            api._updateProgress();
            api.navigateTo('system-landscape');
        },

        /* ── Export full JSON report ─────────────────────── */
        exportReport() {
            if (state.completed.size === 0) {
                alert('No analysis data available yet. Please complete at least one module first.');
                return;
            }

            const report = {
                tool    : 'SAP Sizing Analysis Tool',
                date    : new Date().toISOString(),
                modules : {}
            };

            Object.entries(state.moduleData).forEach(([id, data]) => {
                report.modules[id] = {
                    fileName : data.fileName,
                    rows     : data.rows?.length,
                    headers  : data.headers,
                    analysis : data.analysis || null,
                };
            });

            const blob = new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' });
            const url  = URL.createObjectURL(blob);
            const a    = Object.assign(document.createElement('a'), {
                href    : url,
                download: `sizing-report-${_dateTag()}.json`,
            });
            a.click();
            URL.revokeObjectURL(url);
        },

        /* ── Session save / load ─────────────────────────────
           Serializes the full working state so an analysis can be resumed
           later: the core SizingApp state (moduleData, completed, locked,
           active module) plus each module's own working state via the
           optional getSessionState()/setSessionState() contract.

           Non-serializable artefacts (raw uploaded File bytes, Chart.js
           instances) are intentionally excluded — parsed rows, derived data
           and analysis results are kept, so charts and tables rebuild on
           restore. A step that needs the original file bytes requires a
           re-upload. */
        SESSION_VERSION: 1,

        _sidForFile() {
            const d = state.moduleData['system-landscape'];
            const sid = (d && d.systemInfo && d.systemInfo.sid) || '';
            return String(sid).replace(/[^A-Za-z0-9]/g, '').slice(0, 8).toUpperCase() || 'SAP';
        },

        /* ── Session progress overlay (determinate, stepped) ── */
        _sessionProgress(title, step, pct) {
            const bk = document.getElementById('session-progress-backdrop');
            if (!bk) return;
            bk.style.display = 'flex';
            const set = (id, txt) => { const el = document.getElementById(id); if (el) el.textContent = txt; };
            set('session-progress-title', title);
            set('session-progress-step', step || '');
            set('session-progress-pct', `${Math.round(pct)}%`);
            const fill = document.getElementById('session-progress-fill');
            if (fill) fill.style.width = `${Math.max(0, Math.min(100, pct))}%`;
        },
        _sessionProgressHide() {
            const bk = document.getElementById('session-progress-backdrop');
            if (bk) bk.style.display = 'none';
        },
        /* Yield to the event loop so the progress bar paints between steps. */
        _tick() {
            return new Promise(r => (window.requestAnimationFrame ? requestAnimationFrame(() => r()) : setTimeout(r, 0)));
        },

        async saveSession() {
            try {
                /* Step 1 — Collecting (10%) */
                api._sessionProgress('Saving session', 'Collecting analysis state…', 10);
                await api._tick();
                const moduleState = {};
                Object.entries(MODULES).forEach(([id, getMod]) => {
                    const mod = getMod();
                    if (mod && typeof mod.getSessionState === 'function') {
                        try { moduleState[id] = mod.getSessionState(); }
                        catch (err) { console.warn(`getSessionState failed for ${id}`, err); }
                    }
                });

                const session = {
                    tool          : 'SAP Sizing Analysis Tool',
                    kind          : 'session',
                    version       : api.SESSION_VERSION,
                    savedAt       : new Date().toISOString(),
                    activeModule  : state.activeModule,
                    completed     : Array.from(state.completed),
                    locked        : Array.from(state.locked),
                    moduleData    : state.moduleData,
                    moduleState,
                    chartFontSizes: ChartFontSizeManager.getState(),
                };

                /* Step 2 — Validating / normalising (45%): deep-clone into a
                   JSON-safe structure, converting Date/Map/Set/BigInt, dropping
                   functions/File/Blob/DOM/Chart, and cutting circular refs. */
                api._sessionProgress('Saving session', 'Validating & normalising data…', 45);
                await api._tick();
                const { value: safe, log } = _safeSerialize(session);

                if (log.length) {
                    /* Report what was unsupported — with exact field paths — so the
                       user/developer can see precisely what was transformed/dropped. */
                    const cuts = log.filter(l => l.kind === 'circular' || l.kind === 'function'
                        || l.kind === 'File' || l.kind === 'Blob' || l.kind === 'Node' || l.kind === 'ChartInstance');
                    console.groupCollapsed(`[saveSession] normalised ${log.length} non-JSON value(s)`);
                    log.forEach(l => console.info(`  ${l.kind.padEnd(14)} @ ${l.path || '(root)'}${l.detail ? '  — ' + l.detail : ''}`));
                    console.groupEnd();
                    if (cuts.length) {
                        const first = cuts[0];
                        showToast('Session cleaned', `${cuts.length} non-serializable value(s) were dropped (e.g. ${first.kind} at “${first.path || 'root'}”). Charts/files rebuild on load.`, 'info', 6000);
                    }
                }

                /* Step 2b — Size budgeting (60%): strip data that is duplicated between
                   moduleData and moduleState, then estimate payload size, log large
                   fields, and cap targetRows if still over budget.

                   Root cause of "Invalid string length":
                   state.moduleData['tech-workload-app'].rows  ≡
                   state.moduleData['tech-workload-app'].analysis  ≡
                   moduleState['tech-workload-app'].targetRows
                   — the same large array serialized up to 3×.  We strip `rows` and
                   `analysis` from the moduleData copy (they are reconstructed from
                   moduleState on load) so each dataset is stored exactly once.       */
                api._sessionProgress('Saving session', 'Deduplicating & checking payload size…', 60);
                await api._tick();
                const SIZE_BUDGET_BYTES = 200 * 1024 * 1024;   /* 200 MB soft limit */
                const MAX_TARGET_ROWS   = 50_000;
                const sizeOf = v => {
                    try { return JSON.stringify(v).length; } catch { return 0; }
                };

                /* Strip rows/analysis from moduleData — moduleState carries them. */
                const md = safe.moduleData || {};
                for (const [modId, d] of Object.entries(md)) {
                    if (!d || typeof d !== 'object') continue;
                    if (Array.isArray(d.rows)     && d.rows.length > 0)     delete d.rows;
                    if (Array.isArray(d.analysis) && d.analysis.length > 0) delete d.analysis;
                }

                const ms = safe.moduleState || {};
                const bigFields = [];
                for (const [modId, st] of Object.entries(ms)) {
                    if (!st || typeof st !== 'object') continue;
                    for (const [k, v] of Object.entries(st)) {
                        if (v == null) continue;
                        const bytes = sizeOf(v);
                        if (bytes > 500_000) {
                            bigFields.push({ modId, field: k, bytes });
                        }
                    }
                }
                bigFields.sort((a, b) => b.bytes - a.bytes);
                if (bigFields.length) {
                    console.group('[saveSession] Large session fields (may slow save):');
                    bigFields.forEach(f => console.info(`  ${f.modId}.${f.field}: ${(f.bytes / 1024 / 1024).toFixed(1)} MB`));
                    console.groupEnd();
                }

                /* Apply row cap if targetRows would push us over the budget. */
                let rowsStripped = false;
                let rowsOriginalCount = 0;
                const totalEst = sizeOf(safe);
                if (totalEst > SIZE_BUDGET_BYTES) {
                    for (const [modId, st] of Object.entries(ms)) {
                        if (!st || !Array.isArray(st.targetRows) || st.targetRows.length <= MAX_TARGET_ROWS) continue;
                        rowsOriginalCount = st.targetRows.length;
                        st.targetRows = st.targetRows.slice(0, MAX_TARGET_ROWS);
                        st._targetRowsTruncated = true;
                        st._targetRowsOriginalCount = rowsOriginalCount;
                        rowsStripped = true;
                        console.warn(`[saveSession] targetRows truncated: ${rowsOriginalCount} → ${MAX_TARGET_ROWS} rows (payload was ${(totalEst / 1024 / 1024).toFixed(0)} MB)`);
                    }
                    if (rowsStripped) {
                        showToast('Large dataset', `The monitoring data (${rowsOriginalCount.toLocaleString()} rows) was trimmed to ${MAX_TARGET_ROWS.toLocaleString()} rows to fit within session size limits. Charts will rebuild from the stored subset. Re-upload the original file to restore full chart fidelity.`, 'info', 10000);
                    }
                }

                /* Step 3 — Serializing (75%) */
                api._sessionProgress('Saving session', 'Serializing to JSON…', 75);
                await api._tick();
                let json;
                try {
                    json = JSON.stringify(safe, null, 2);
                } catch (err) {
                    /* If stringify still fails (dataset too large even after row cap),
                       retry once with targetRows stripped entirely. This preserves all
                       mapping, EM config, SAPS overrides and summary stats — only the
                       raw enriched table (and therefore chart rebuilding) is lost. */
                    console.error('saveSession serialize error (first attempt)', err);
                    let retried = false;
                    try {
                        const ms2 = safe.moduleState || {};
                        for (const st of Object.values(ms2)) {
                            if (st && Array.isArray(st.targetRows)) {
                                st._targetRowsOriginalCount = st.targetRows.length;
                                st.targetRows  = [];
                                st.targetCols  = [];
                                st._targetRowsTruncated = true;
                                st._targetRowsDropped   = true;
                            }
                        }
                        json = JSON.stringify(safe, null, 2);
                        retried = true;
                        showToast('Session saved (partial)', 'The monitoring dataset was too large to include — all other analysis state was preserved. Re-upload the original file to restore charts.', 'info', 10000);
                    } catch (err2) {
                        console.error('saveSession serialize error (retry)', err2);
                        showToast('Save failed', `Serialization error: ${err2 && err2.message ? err2.message : err2}`, 'error', 8000);
                        api._sessionProgressHide();
                        return;
                    }
                    if (!retried) return;
                }

                /* Step 4 — Writing (100%) */
                api._sessionProgress('Saving session', 'Writing file…', 100);
                await api._tick();
                const blob = new Blob([json], { type: 'application/json' });
                const url  = URL.createObjectURL(blob);
                const a    = Object.assign(document.createElement('a'), {
                    href    : url,
                    download: `${api._sidForFile()}_session_${_dateTag()}.json`,
                });
                document.body.appendChild(a); a.click(); document.body.removeChild(a);
                setTimeout(() => URL.revokeObjectURL(url), 1000);

                api._sessionProgressHide();
                showToast('Session saved', 'The current analysis was downloaded as a session file.', 'success');
            } catch (err) {
                api._sessionProgressHide();
                console.error('saveSession failed', err);
                showToast('Save failed', err && err.message ? err.message : String(err), 'error', 8000);
            }
        },

        loadSession(file) {
            /* Step 1 — Reading (20%) */
            api._sessionProgress('Loading session', `Reading ${file.name}…`, 20);
            const reader = new FileReader();
            reader.onprogress = e => {
                if (e.lengthComputable) {
                    const pct = 5 + Math.round((e.loaded / e.total) * 15);   /* 5–20% during read */
                    api._sessionProgress('Loading session', `Reading ${file.name}…`, pct);
                }
            };
            reader.onload = async e => {
                try {
                    /* Step 2 — Parsing (40%) */
                    api._sessionProgress('Loading session', 'Parsing JSON…', 40);
                    await api._tick();
                    let session;
                    try { session = JSON.parse(e.target.result); }
                    catch (err) {
                        api._sessionProgressHide();
                        showToast('Load failed', 'The selected file is not valid JSON.', 'error');
                        return;
                    }

                    /* Step 3 — Validating (55%) */
                    api._sessionProgress('Loading session', 'Validating session file…', 55);
                    await api._tick();
                    if (!session || session.kind !== 'session' || !session.moduleData) {
                        api._sessionProgressHide();
                        showToast('Load failed', 'This does not look like a saved session file.', 'error');
                        return;
                    }
                    if (session.version !== api.SESSION_VERSION) {
                        /* Backward/forward-compatible: warn but still attempt. */
                        showToast('Version note', `Session saved by version ${session.version ?? '(unknown)'} — loading with the current version ${api.SESSION_VERSION}.`, 'info', 6000);
                    }

                    /* Revive tagged Date/Map/Set/BigInt wrappers (older files that
                       never used wrappers pass through untouched). */
                    session = _reviveSession(session);

                    /* Step 4 — Restoring (85%) */
                    api._sessionProgress('Loading session', 'Restoring analysis state…', 85);
                    await api._tick();
                    api._applySession(session);

                    /* Step 5 — Done (100%) */
                    api._sessionProgress('Loading session', 'Finalising…', 100);
                    await api._tick();
                    api._sessionProgressHide();
                    showToast('Session loaded', 'The analysis was restored from the session file.', 'success');
                } catch (err) {
                    api._sessionProgressHide();
                    console.error('loadSession failed', err);
                    showToast('Load failed', err && err.message ? err.message : String(err), 'error', 8000);
                }
            };
            reader.onerror = () => { api._sessionProgressHide(); showToast('Load failed', 'The file could not be read.', 'error'); };
            reader.readAsText(file);
        },

        _applySession(session) {
            /* 1. Reset every module's in-memory singleton to a clean baseline. */
            Object.values(MODULES).forEach(getMod => {
                const mod = getMod();
                if (mod && typeof mod.resetState === 'function') mod.resetState();
            });

            /* 2. Restore core SizingApp state. */
            state.moduleData  = session.moduleData || {};
            state.completed   = new Set(session.completed || []);
            state.locked      = new Set(session.locked || []);
            state.activeModule = session.activeModule || 'system-landscape';

            /* 2a. Restore global chart font sizes. */
            if (session.chartFontSizes) ChartFontSizeManager.setState(session.chartFontSizes);

            /* 2b. Re-populate moduleData.rows / analysis from moduleState.targetRows
               for any module that stores them separately (e.g. tech-workload-app).
               Older saves included rows directly in moduleData; newer saves strip them
               to avoid 3× duplication, so we reconstruct from moduleState here. */
            const savedModuleState = session.moduleState || {};
            Object.entries(savedModuleState).forEach(([id, st]) => {
                if (!st) return;
                const mdEntry = state.moduleData[id];
                if (mdEntry && Array.isArray(st.targetRows) && st.targetRows.length > 0) {
                    if (!mdEntry.rows || mdEntry.rows.length === 0) {
                        mdEntry.rows     = st.targetRows;
                        mdEntry.analysis = st.targetRows;
                    }
                }
            });

            /* 3. Restore each module's working state (opt-in). */
            Object.entries(savedModuleState).forEach(([id, st]) => {
                const getMod = MODULES[id];
                const mod = getMod && getMod();
                if (mod && typeof mod.setSessionState === 'function') {
                    try { mod.setSessionState(st); }
                    catch (err) { console.warn(`setSessionState failed for ${id}`, err); }
                }
            });

            /* 4. Reflect progress + lock state onto the sidebar, then render. */
            document.querySelectorAll('.nav-item').forEach(el => {
                el.classList.toggle('completed', state.completed.has(el.dataset.module));
            });
            api._applyLockState();
            api._updateProgress();

            /* Navigate to the saved active module if it is reachable; otherwise
               fall back to the first module. */
            const target = state.locked.has(state.activeModule) ? 'system-landscape' : state.activeModule;
            api.navigateTo(target);
        },

        /* ── Progress bar ───────────────────────────────── */
        _updateProgress() {
            const total = MODULE_ORDER.length;
            const done  = state.completed.size;
            const pct   = Math.round((done / total) * 100);

            document.getElementById('progress-fill').style.width = `${pct}%`;
            document.getElementById('progress-pct').textContent  = `${pct}%`;
            document.getElementById('progress-detail').textContent =
                `${done} of ${total} modules completed`;
        },
    };

    return api;
})();

/* ── Global chart font-size manager ──────────────────── */
/*
   Single source of truth for chart font sizes, shared by all modules.
   Session persistence happens via app.js's own getSessionState/setSessionState
   which are called automatically by saveSession/loadSession.
*/
const ChartFontSizeManager = (() => {
    const DEFAULTS = Object.freeze({ title: 25, axisTitle: 14, axisTicks: 14, legend: 17 });
    const LIMITS   = Object.freeze({ title: { min: 10, max: 40 }, axisTitle: { min: 8, max: 30 }, axisTicks: { min: 8, max: 30 }, legend: { min: 8, max: 30 } });
    let _sizes = { ...DEFAULTS };

    return {
        get()        { return { ..._sizes }; },
        getDefaults(){ return { ...DEFAULTS }; },
        set(key, val){
            const lim = LIMITS[key] || { min: 7, max: 40 };
            _sizes[key] = Math.min(lim.max, Math.max(lim.min, Math.round(Number(val)) || DEFAULTS[key]));
        },
        reset()      { _sizes = { ...DEFAULTS }; },
        getState()   { return { ..._sizes }; },
        setState(st) { if (st) _sizes = { ...DEFAULTS, ...st }; },
    };
})();

/* Returns the font-size control row HTML (placed after the canvas).
   uid must be unique per chart so the row has a stable id. */
function chartFontRowHtml(uid) {
    const sz = ChartFontSizeManager.get();
    return `
        <div class="chart-font-row" id="${uid}-font-row">
            <span class="chart-font-row-label">Font:</span>
            <div class="chart-font-group" title="Font size used for chart title">
                <span class="chart-font-group-label">Title</span>
                <button class="btn btn-sm chart-font-btn" data-fs-action="dec" data-fs-key="title">A−</button>
                <input  type="number" class="chart-font-input" id="${uid}-fs-title"
                        data-fs-key="title" value="${sz.title}" min="10" max="40"
                        title="Font size used for chart title">
                <button class="btn btn-sm chart-font-btn" data-fs-action="inc" data-fs-key="title">A+</button>
            </div>
            <div class="chart-font-sep"></div>
            <div class="chart-font-group" title="Font size used for axis titles and labels">
                <span class="chart-font-group-label">Axis</span>
                <button class="btn btn-sm chart-font-btn" data-fs-action="dec" data-fs-key="axis">A−</button>
                <input  type="number" class="chart-font-input" id="${uid}-fs-axis"
                        data-fs-key="axis" value="${sz.axisTitle}" min="8" max="30"
                        title="Font size used for axis titles and labels">
                <button class="btn btn-sm chart-font-btn" data-fs-action="inc" data-fs-key="axis">A+</button>
            </div>
            <div class="chart-font-sep"></div>
            <div class="chart-font-group" title="Font size used for chart legend">
                <span class="chart-font-group-label">Legend</span>
                <button class="btn btn-sm chart-font-btn" data-fs-action="dec" data-fs-key="legend">A−</button>
                <input  type="number" class="chart-font-input" id="${uid}-fs-legend"
                        data-fs-key="legend" value="${sz.legend}" min="8" max="30"
                        title="Font size used for chart legend">
                <button class="btn btn-sm chart-font-btn" data-fs-action="inc" data-fs-key="legend">A+</button>
            </div>
            <div class="chart-font-sep"></div>
            <button class="btn btn-sm" data-fs-action="reset" title="Reset all font sizes to defaults">Reset</button>
        </div>`;
}

/* Apply current ChartFontSizeManager sizes to a Chart.js instance.
   isPie: true for doughnut/pie (no scales). */
function applyChartFontSizes(chart, isPie) {
    if (!chart) return;
    const sz = ChartFontSizeManager.get();
    if (chart.options.plugins && chart.options.plugins.legend) {
        const leg = chart.options.plugins.legend;
        leg.labels = leg.labels || {};
        leg.labels.font = leg.labels.font || {};
        leg.labels.font.size = sz.legend;
    }
    if (!isPie) {
        const scales = chart.options.scales || {};
        Object.values(scales).forEach(scale => {
            scale.ticks = scale.ticks || {};
            scale.ticks.font = scale.ticks.font || {};
            scale.ticks.font.size = sz.axisTicks;
            if (scale.title && scale.title.display) {
                scale.title.font = scale.title.font || {};
                scale.title.font.size = sz.axisTitle;
            }
        });
    }
    chart.update('none');
}

/* Sync all visible number inputs in a font row to current manager state. */
function _syncFontInputs(row) {
    if (!row) return;
    const sz = ChartFontSizeManager.get();
    const keyMap = { title: sz.title, axis: sz.axisTitle, legend: sz.legend };
    row.querySelectorAll('input[data-fs-key]').forEach(inp => {
        const mapped = keyMap[inp.dataset.fsKey];
        if (mapped !== undefined) inp.value = mapped;
    });
}

/* Re-entry guard: prevents applyAll() from recursing if Chart.js or a plugin
   fires a synchronous update/resize callback while we are already applying
   font sizes. One global flag is fine — font-size changes are user-driven and
   always complete before the next one can start. */
let _applyingFonts = false;

/* Wire a font-size control row for one or more Chart.js instances.
   containerEl  — the card/wrapper that contains the row (search by id suffix).
   getCharts()  — function returning the array of chart instances to update.
   allRows()    — function returning all currently visible font rows (for sync).
   isPie        — whether the chart(s) have no scales (pie/doughnut). */
function wireChartFontRow(containerEl, getCharts, allRowsFn, isPie) {
    const row = containerEl.querySelector('[id$="-font-row"]');
    if (!row) return;

    const STEP = 1;
    const axisKeys = ['axisTitle', 'axisTicks'];

    const applyAll = () => {
        if (_applyingFonts) return;
        _applyingFonts = true;
        try {
            const charts = getCharts();
            charts.forEach(c => {
                try {
                    applyChartFontSizes(c, isPie != null ? isPie : !(c.options && c.options.scales && Object.keys(c.options.scales).length));
                } catch (_) { /* never let one chart block the rest */ }
            });
            /* Sync all rows in the scope (pass allRowsFn if multiple charts share font state) */
            const allRows = allRowsFn ? allRowsFn() : [row];
            allRows.forEach(r => _syncFontInputs(r));
        } finally {
            _applyingFonts = false;
        }
    };

    row.addEventListener('click', e => {
        const btn = e.target.closest('[data-fs-action][data-fs-key]');
        if (!btn) {
            /* Check for Reset */
            const rst = e.target.closest('[data-fs-action="reset"]');
            if (rst) { ChartFontSizeManager.reset(); applyAll(); }
            return;
        }
        const action = btn.dataset.fsAction;
        const fsKey  = btn.dataset.fsKey;
        if (fsKey === 'axis') {
            axisKeys.forEach(k => ChartFontSizeManager.set(k, (action === 'inc' ? 1 : -1) * STEP + ChartFontSizeManager.get()[k]));
        } else {
            const cur = ChartFontSizeManager.get()[fsKey] || 14;
            ChartFontSizeManager.set(fsKey, cur + (action === 'inc' ? STEP : -STEP));
        }
        applyAll();
    });

    row.addEventListener('change', e => {
        const inp = e.target.closest('input[data-fs-key]');
        if (!inp) return;
        const fsKey = inp.dataset.fsKey;
        const val   = parseFloat(inp.value);
        if (isNaN(val)) return;
        if (fsKey === 'axis') {
            axisKeys.forEach(k => ChartFontSizeManager.set(k, val));
        } else {
            ChartFontSizeManager.set(fsKey, val);
        }
        /* Clamp displayed value back to the validated range */
        inp.value = fsKey === 'axis' ? ChartFontSizeManager.get().axisTitle : ChartFontSizeManager.get()[fsKey];
        applyAll();
    });

    /* Apply on next tick to avoid re-entering Chart.js during chart initialisation.
       Deferring with setTimeout(0) breaks any synchronous update→resize→update
       cycles that Chart.js or plugins may fire during new Chart() construction. */
    setTimeout(applyAll, 0);
}

/* ── Shared utilities ─────────────────────────────────── */

function _dateTag() {
    return new Date().toISOString().split('T')[0];
}

/* ── Safe session serialization ──────────────────────────
   JSON.stringify throws on the FIRST circular reference it meets — before a
   replacer function can prune anything — which is the root cause of the
   "could not be serialized to JSON" error. _safeSerialize walks the payload
   itself, producing a JSON-safe deep clone and converting every value JSON
   cannot represent:

     Date       → { __type:'Date',  iso }          (restored to a Date)
     Map        → { __type:'Map',   entries:[…] }  (restored to a Map)
     Set        → { __type:'Set',   values:[…] }   (restored to a Set)
     BigInt     → { __type:'BigInt',value:'…' }    (restored to a BigInt)
     undefined  → dropped (object key omitted; array slot → null)
     function   → dropped
     File/Blob/Node/Chart instance → dropped
     circular reference → dropped, with the exact field path logged

   Every drop/convert is recorded in `log` with its dotted field path, so the
   caller can report precisely what was unsupported. Returns
   { value, log:[{path, kind, detail}] }. */
function _safeSerialize(root) {
    const log = [];
    const seen = new WeakSet();

    const isChart = v => v && typeof v === 'object'
        && v.canvas && typeof v.update === 'function' && typeof v.destroy === 'function';

    const walk = (val, path) => {
        if (val === null) return null;

        const t = typeof val;
        if (t === 'undefined') { log.push({ path, kind: 'undefined' }); return undefined; }
        if (t === 'function')  { log.push({ path, kind: 'function' }); return undefined; }
        if (t === 'bigint')    { log.push({ path, kind: 'BigInt' });   return { __type: 'BigInt', value: val.toString() }; }
        if (t === 'number')    { return isFinite(val) ? val : null; }   /* NaN/Infinity → null (JSON does this anyway) */
        if (t === 'string' || t === 'boolean') return val;

        /* Objects from here down. */
        if (val instanceof Date) {
            log.push({ path, kind: 'Date' });
            return { __type: 'Date', iso: isNaN(val) ? null : val.toISOString() };
        }
        if (typeof File !== 'undefined' && val instanceof File) { log.push({ path, kind: 'File', detail: val.name }); return undefined; }
        if (typeof Blob !== 'undefined' && val instanceof Blob) { log.push({ path, kind: 'Blob' }); return undefined; }
        if (typeof Node !== 'undefined' && val instanceof Node) { log.push({ path, kind: 'Node' }); return undefined; }
        if (isChart(val)) { log.push({ path, kind: 'ChartInstance' }); return undefined; }

        if (seen.has(val)) { log.push({ path, kind: 'circular' }); return undefined; }
        seen.add(val);

        let out;
        if (val instanceof Map) {
            log.push({ path, kind: 'Map' });
            const entries = [];
            let i = 0;
            for (const [k, v] of val.entries()) {
                const cv = walk(v, `${path}.<Map value ${JSON.stringify(k)}>`);
                if (cv !== undefined) entries.push([typeof k === 'object' ? String(k) : k, cv]);
                i++;
            }
            out = { __type: 'Map', entries };
        } else if (val instanceof Set) {
            log.push({ path, kind: 'Set' });
            const values = [];
            let i = 0;
            for (const v of val.values()) {
                const cv = walk(v, `${path}[Set ${i++}]`);
                if (cv !== undefined) values.push(cv);
            }
            out = { __type: 'Set', values };
        } else if (Array.isArray(val)) {
            out = val.map((v, i) => {
                const cv = walk(v, `${path}[${i}]`);
                return cv === undefined ? null : cv;   /* arrays can't have holes in JSON */
            });
        } else {
            /* Plain-ish object (strip non-serializable prototypes implicitly by
               copying own enumerable keys only). */
            out = {};
            for (const k of Object.keys(val)) {
                const cv = walk(val[k], path ? `${path}.${k}` : k);
                if (cv !== undefined) out[k] = cv;   /* omit undefined keys */
            }
        }

        seen.delete(val);   /* allow the same object to appear on sibling branches (shared refs are fine; only true cycles are cut) */
        return out;
    };

    const value = walk(root, '');
    return { value, log };
}

/* Reviver counterpart — turns the tagged wrappers from _safeSerialize back
   into Date / Map / Set / BigInt. Safe on plain objects and on older session
   files that never used the wrappers. */
function _reviveSession(val) {
    if (val === null || typeof val !== 'object') return val;
    if (Array.isArray(val)) return val.map(_reviveSession);
    if (val.__type === 'Date')   return val.iso ? new Date(val.iso) : null;
    if (val.__type === 'BigInt') { try { return BigInt(val.value); } catch { return val.value; } }
    if (val.__type === 'Set')    return new Set((val.values || []).map(_reviveSession));
    if (val.__type === 'Map')    return new Map((val.entries || []).map(([k, v]) => [k, _reviveSession(v)]));
    const out = {};
    for (const k of Object.keys(val)) out[k] = _reviveSession(val[k]);
    return out;
}

/* Show a transient toast notification (top-right).
   type: 'success' | 'error' | 'info' (default). Auto-dismisses after `ms`. */
function showToast(title, message = '', type = 'info', ms = 5000) {
    const stack = document.getElementById('toast-stack');
    if (!stack) return;
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    const icon = type === 'success'
        ? '<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>'
        : type === 'error'
        ? '<circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>'
        : '<circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>';
    toast.innerHTML = `
        <svg class="toast-icon" xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24"
             fill="none" stroke="currentColor" stroke-width="2.2">${icon}</svg>
        <div class="toast-body">
            <div class="toast-title">${title}</div>
            ${message ? `<div class="toast-msg">${message}</div>` : ''}
        </div>
        <button class="toast-close" aria-label="Dismiss">&times;</button>`;
    stack.appendChild(toast);
    requestAnimationFrame(() => toast.classList.add('show'));

    const dismiss = () => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 320);
    };
    toast.querySelector('.toast-close').addEventListener('click', dismiss);
    if (ms > 0) setTimeout(dismiss, ms);
}

/* Parse file (CSV or XLSX) and call back with rows array */
function parseUploadedFile(file, callback) {
    const ext = file.name.split('.').pop().toLowerCase();

    if (ext === 'csv') {
        Papa.parse(file, {
            header       : true,
            skipEmptyLines: true,
            complete     : r => callback(r.data, file.name),
        });
        return;
    }

    if (['xlsx', 'xls'].includes(ext)) {
        const reader = new FileReader();
        reader.onload = e => {
            const wb   = XLSX.read(e.target.result, { type: 'binary' });
            const ws   = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
            callback(rows, file.name);
        };
        reader.readAsBinaryString(file);
        return;
    }

    alert(`Unsupported file format ".${ext}". Please upload a CSV, XLSX, or XLS file.`);
}

/* Wire up a drop-zone + hidden file input pair */
function initUploadZone(zoneId, inputId, onFile) {
    const zone  = document.getElementById(zoneId);
    const input = document.getElementById(inputId);
    if (!zone || !input) return;

    zone.addEventListener('click', () => input.click());

    zone.addEventListener('dragover', e => {
        e.preventDefault();
        zone.classList.add('drag-over');
    });
    zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
    zone.addEventListener('drop', e => {
        e.preventDefault();
        zone.classList.remove('drag-over');
        const f = e.dataTransfer.files[0];
        if (f) onFile(f);
    });

    input.addEventListener('change', e => {
        const f = e.target.files[0];
        if (f) onFile(f);
    });
}

/* Show a file-loaded status row below a zone */
function showFileLoaded(containerId, fileName) {
    const el = document.getElementById(containerId);
    if (!el) return;
    el.style.display = 'flex';
    el.innerHTML = `
        <div class="file-loaded">
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24"
                 fill="none" stroke="currentColor" stroke-width="2.5">
                <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                <polyline points="22 4 12 14.01 9 11.01"/>
            </svg>
            <span>Loaded: <strong>${fileName}</strong></span>
        </div>`;
}

/* Build simple stat-card HTML */
function statCard(label, value, unit = '') {
    return `
        <div class="stat-card">
            <div class="stat-label">${label}</div>
            <div class="stat-value">${value}<span class="stat-unit">${unit}</span></div>
        </div>`;
}

/* Download data as XLSX */
function downloadXlsx(rows, sheetName, fileName) {
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, fileName);
}

/* Detect a column by regex from a list of header names */
function detectCol(headers, regex) {
    return headers.find(h => regex.test(h)) || null;
}

/* Count occurrences of a column's values */
function countBy(rows, col) {
    const out = {};
    rows.forEach(r => {
        const v = (r[col] || 'Unknown').toString().trim();
        out[v] = (out[v] || 0) + 1;
    });
    return out;
}

/* Compute percentile from a numeric array */
function percentile(arr, p) {
    const sorted = [...arr].map(Number).filter(n => !isNaN(n)).sort((a, b) => a - b);
    if (!sorted.length) return 0;
    const idx = Math.ceil((p / 100) * sorted.length) - 1;
    return sorted[Math.max(0, idx)];
}

/* Arithmetic mean of a numeric array (0 for empty) */
function _avg(arr) {
    return arr.length ? arr.reduce((s, v) => s + v, 0) / arr.length : 0;
}

/* SAP-palette chart colours */
const CHART_COLORS = [
    '#0070F2','#10B981','#F59E0B','#EF4444',
    '#8B5CF6','#06B6D4','#F97316','#84CC16',
];

/* ── Boot ─────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => SizingApp.init());
