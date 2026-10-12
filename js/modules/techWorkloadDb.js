/* ─────────────────────────────────────────────────────────
   Module 03 – Technical Workload: Database Layer (SAP HANA)
   ─────────────────────────────────────────────────────────
   HANA-specific workflow:
     • Collects System Landscape info, HANA CPU/Memory history
       (SAP Note 1969700) and the mandatory "7 Memory Component
       Overview" export.
     • Appends SAPS columns (Cores, SAPS/Core, SAPS Deployed,
       SAPS Utilization) derived from the System Landscape
       database-layer hosts.
     • Charts (CPU + Memory tabs) reuse the same framework as the
       Application-layer module (font controls, Copy/PNG/SVG/CSV,
       responsive sizing).
     • Exports <DatabaseSID>_database.xlsx with native, editable,
       data-linked Excel charts (Data + Chart worksheet pairs).

   Shared globals used (defined in app.js): Chart, CHART_COLORS,
   showToast, ChartFontSizeManager, chartFontRowHtml, wireChartFontRow,
   initUploadZone, showFileLoaded, parseUploadedFile, detectCol,
   downloadXlsx, _dateTag, XLSX, JSZip.
   ─────────────────────────────────────────────────────────*/

const TechWorkloadDbModule = {

    _charts: {},

    /* ── Persistent working state ── */
    hanaConfirmed: false,
    landscape:  null,   /* { rows, headers, fileName } */
    cpuUpload:  null,   /* { rows, headers, fileName } — raw CPU/memory history */
    memUpload:  null,   /* { rows, headers, fileName } — 7 Memory Component Overview (mandatory) */
    histUpload: null,   /* { rows, headers, fileName } — optional historic EWA memory */
    targetRows: [],     /* enriched CPU rows (array-of-objects) — named targetRows for session dedup */
    targetCols: [],     /* enriched CPU column order */
    memNormalized: null,/* { labels:[MY_DATE], components:[name], series:{name:[gb]} } */
    histNormalized: null,/* same shape as memNormalized — Historic EWA, kept independent */
    sapsBasis:  null,   /* { cores, sapsPerCore, totalSaps, hosts:[] } from the landscape */

    resetState() {
        this._charts = {};
        this.hanaConfirmed = false;
        this.landscape = this.cpuUpload = this.memUpload = this.histUpload = null;
        this.targetRows = [];
        this.targetCols = [];
        this.memNormalized = null;
        this.histNormalized = null;
        this.sapsBasis = null;
        this._renderedSection = null;
        this._activateToken = 0;
    },

    /* ── Session contract (called by app.js saveSession/loadSession) ── */
    getSessionState() {
        return {
            hanaConfirmed: this.hanaConfirmed,
            landscape:  this.landscape,
            cpuUpload:  this.cpuUpload,
            memUpload:  this.memUpload,
            histUpload: this.histUpload,
            targetRows: this.targetRows,
            targetCols: this.targetCols,
            memNormalized: this.memNormalized,
            histNormalized: this.histNormalized,
            sapsBasis:  this.sapsBasis,
        };
    },
    setSessionState(st) {
        if (!st) return;
        this.hanaConfirmed = !!st.hanaConfirmed;
        this.landscape  = st.landscape  || null;
        this.cpuUpload  = st.cpuUpload  || null;
        this.memUpload  = st.memUpload  || null;
        this.histUpload = st.histUpload || null;
        this.targetRows = st.targetRows || [];
        this.targetCols = st.targetCols || [];
        this.memNormalized = st.memNormalized || null;
        this.histNormalized = st.histNormalized || null;
        this.sapsBasis  = st.sapsBasis  || null;
    },

    /* ════════════════════════════════════════════════════════
       RENDER
       ════════════════════════════════════════════════════════ */
    render() {
        return `
        <div class="module-wrap" id="twd-root">

            <!-- ══ Data-collection view ══ -->
            <div id="twd-collect">
                <div class="module-header">
                    <div class="module-crumb">Analysis › Step 03</div>
                    <h1 class="module-title">Technical Workload – Database Layer</h1>
                    <p class="module-desc">
                        SAP HANA technical workload analysis. Upload the HANA Initial Capacity Check
                        datasets below; CPU SAPS metrics are derived from the System Landscape (Module 01).
                    </p>
                </div>

                <!-- Database info header -->
                <div id="twd-dbinfo"></div>

                <!-- HANA detection notice -->
                <div id="twd-hana-notice"></div>

                <!-- Upload tiles (shown once HANA workflow is active) -->
                <div id="twd-uploads" style="display:none">

                    <div class="section-title" style="margin-top:22px">Infrastructure &amp; Workload Datasets</div>

                    <div class="card" style="margin-bottom:14px">
                        <div class="card-header">
                            <div>
                                <div class="card-title">System Landscape Information</div>
                                <div class="card-sub">RSxxxSTT → HANA → HANA Initial Capacity Check → 1 System / Landscape Information</div>
                            </div>
                            <span class="badge-amber" id="twd-badge-landscape">Mandatory</span>
                        </div>
                        <div class="upload-zone" id="twd-zone-landscape">
                            <input type="file" class="file-input-hidden" id="twd-input-landscape" accept=".txt,.csv,.xlsx,.xls">
                            <div class="upload-zone-title">Drop the System / Landscape Information export here or click to browse</div>
                            <span class="upload-formats">TXT · CSV · XLSX</span>
                        </div>
                        <div id="twd-status-landscape"></div>
                    </div>

                    <div class="card" style="margin-bottom:14px">
                        <div class="card-header">
                            <div>
                                <div class="card-title">CPU Analysis — HANA Resources CPU &amp; Memory History</div>
                                <div class="card-sub">HANA_Resources_CPUAndMemory_History_* · SAP Note 1969700</div>
                            </div>
                            <span class="badge-amber" id="twd-badge-cpu">Mandatory</span>
                        </div>
                        <div class="upload-zone" id="twd-zone-cpu">
                            <input type="file" class="file-input-hidden" id="twd-input-cpu" accept=".txt,.csv,.xlsx,.xls">
                            <div class="upload-zone-title">Drop the CPU &amp; Memory history export here or click to browse</div>
                            <span class="upload-formats">TXT · CSV · XLSX</span>
                        </div>
                        <div class="alert-info" style="margin-top:10px">
                            <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
                            <span>Adjust the script aggregation interval as required. Example: <code>'TS300' TIME_AGGREGATE_BY</code> for 5-minute snapshots.</span>
                        </div>
                        <div id="twd-status-cpu"></div>
                    </div>

                    <div class="card" style="margin-bottom:14px">
                        <div class="card-header">
                            <div>
                                <div class="card-title">Memory Analysis — 7 Memory Component Overview</div>
                                <div class="card-sub">RSxxxSTT → HANA → HANA Initial Capacity Check → 7 Memory Component Overview · <strong>mandatory</strong></div>
                            </div>
                            <span class="badge-amber" id="twd-badge-mem">Mandatory</span>
                        </div>
                        <div class="upload-zone" id="twd-zone-mem">
                            <input type="file" class="file-input-hidden" id="twd-input-mem" accept=".txt,.csv,.xlsx,.xls">
                            <div class="upload-zone-title">Drop the Memory Component Overview export here or click to browse</div>
                            <span class="upload-formats">TXT · CSV · XLSX</span>
                        </div>
                        <div id="twd-status-mem"></div>
                    </div>

                    <div class="card" style="margin-bottom:14px">
                        <div class="card-header">
                            <div>
                                <div class="card-title">Historic EWA Memory Components Overview</div>
                                <div class="card-sub">Optional · daily granularity · ≥ 12 months · exported manually from the EarlyWatch Dashboard</div>
                            </div>
                            <span class="badge-gray">Optional</span>
                        </div>
                        <div class="upload-zone" id="twd-zone-hist">
                            <input type="file" class="file-input-hidden" id="twd-input-hist" accept=".txt,.csv,.xlsx,.xls">
                            <div class="upload-zone-title">Drop the historic EWA memory export here or click to browse</div>
                            <span class="upload-formats">TXT · CSV · XLSX</span>
                        </div>
                        <div id="twd-status-hist"></div>
                    </div>

                    <!-- Validation summary -->
                    <div id="twd-validation" style="margin-top:6px"></div>

                    <div class="action-row">
                        <button class="btn btn-blue" id="twd-btn-charts" disabled>
                            Proceed to Charts
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="9 18 15 12 9 6"/></svg>
                        </button>
                    </div>
                </div>
            </div>

            <!-- ══ Charts view ══ -->
            <div id="twd-charts" style="display:none">
                <div class="module-header" style="margin-top:4px">
                    <div class="module-crumb">Analysis › Step 03 › Charts</div>
                    <h1 class="module-title">Technical Workload – Database (HANA)</h1>
                    <p class="module-desc">Interactive CPU and Memory analysis for the HANA database tier.</p>
                </div>

                <div id="twd-chart-progress" class="proc-progress" style="display:none;margin-bottom:18px">
                    <div class="proc-progress-head">
                        <span id="twd-chart-progress-label">Generating charts…</span>
                        <span id="twd-chart-progress-pct">0%</span>
                    </div>
                    <div class="proc-progress-track"><div class="proc-progress-fill" id="twd-chart-progress-fill"></div></div>
                </div>

                <div class="chart-tabs" id="twd-chart-tabs" role="tablist">
                    <button class="chart-tab is-active" data-tab="cpu" role="tab">CPU Utilization</button>
                    <button class="chart-tab" data-tab="mem" role="tab">Memory Utilization</button>
                </div>

                <div class="action-row" style="margin:4px 0 18px;justify-content:flex-start">
                    <button class="btn btn-blue" id="twd-wb-top" title="Download the HANA database analysis workbook with native, editable Excel charts linked to their data">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                        Download Database Analysis (XLSX)
                    </button>
                    <span style="font-size:11px;color:var(--text-muted)">Includes source + transformed data worksheets and one Data + Chart worksheet pair per dataset (native, editable Excel charts).</span>
                </div>

                <div class="chart-section" data-sec="cpu">
                    <div class="chart-section-body" id="twd-sec-cpu"></div>
                </div>
                <div class="chart-section" data-sec="mem" style="display:none">
                    <div class="chart-section-body" id="twd-sec-mem"></div>
                </div>

                <div class="action-row">
                    <button class="btn" id="twd-charts-back">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 18 9 12 15 6"/></svg>
                        Back to Data Collection
                    </button>
                    <button class="btn btn-blue" id="twd-btn-next">
                        Next: Application Workload
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="9 18 15 12 9 6"/></svg>
                    </button>
                </div>
            </div>
        </div>`;
    },

    /* ════════════════════════════════════════════════════════
       INIT
       ════════════════════════════════════════════════════════ */
    init(app) {
        this.app = app;
        this._charts = this._charts || {};

        this._renderDbInfo(app);
        this._renderHanaNotice(app);
        this._refreshWorkflowVisibility();

        /* Upload zones (TXT/CSV/XLSX via the local _parseFile). */
        initUploadZone('twd-zone-landscape', 'twd-input-landscape', f => this._parseFile(f, (rows, name) => this._onUpload('landscape', rows, name)));
        initUploadZone('twd-zone-cpu',       'twd-input-cpu',       f => this._parseFile(f, (rows, name) => this._onUpload('cpu', rows, name)));
        initUploadZone('twd-zone-mem',       'twd-input-mem',       f => this._parseFile(f, (rows, name) => this._onUpload('mem', rows, name)));
        initUploadZone('twd-zone-hist',      'twd-input-hist',      f => this._parseFile(f, (rows, name) => this._onUpload('hist', rows, name)));

        document.getElementById('twd-btn-charts').addEventListener('click', () => this._showChartsView(app));
        document.getElementById('twd-charts-back').addEventListener('click', () => this._showCollectView());
        document.getElementById('twd-btn-next').addEventListener('click', () => app.nextModule());

        const wbBtn = document.getElementById('twd-wb-top');
        if (wbBtn) wbBtn.addEventListener('click', () => this._buildWorkbook());

        /* Re-hydrate statuses after a session load. */
        ['landscape', 'cpu', 'mem', 'hist'].forEach(kind => {
            const u = this._uploadFor(kind);
            if (u) showFileLoaded(`twd-status-${kind}`, u.fileName);
        });
        this._refreshValidation();
    },

    _uploadFor(kind) {
        return kind === 'landscape' ? this.landscape
             : kind === 'cpu' ? this.cpuUpload
             : kind === 'mem' ? this.memUpload
             : this.histUpload;
    },

    /* ════════════════════════════════════════════════════════
       DB INFO HEADER + HANA DETECTION
       ════════════════════════════════════════════════════════ */
    _systemInfo(app) {
        const d = app.getModuleData && app.getModuleData('system-landscape');
        return (d && d.systemInfo) || {};
    },

    _renderDbInfo(app) {
        const el = document.getElementById('twd-dbinfo');
        if (!el) return;
        const si = this._systemInfo(app);
        const dash = v => (v && String(v).trim()) ? v : '—';
        const anyKnown = si.dbSystem || si.sid || si.dbVersion || si.dbOs;
        if (!anyKnown) { el.innerHTML = ''; return; }
        el.innerHTML = `
            <div class="stats-grid" style="margin-top:14px">
                ${statCard('Database', dash(si.dbSystem))}
                ${statCard('Database SID', dash(si.sid))}
                ${statCard('Database Release', dash(si.dbVersion))}
                ${statCard('DB Server OS &amp; Patch Level', dash(si.dbOs))}
            </div>`;
    },

    /* HANA detected when systemInfo.dbSystem mentions HANA, or any DB-layer host
       component/landscape text does. */
    _hanaDetected(app) {
        const si = this._systemInfo(app);
        const hay = `${si.dbSystem || ''} ${si.product || ''} ${si.dbVersion || ''}`.toLowerCase();
        if (/\bhana\b|hana db/.test(hay)) return true;
        const d = app.getModuleData && app.getModuleData('system-landscape');
        const dbServers = (d && d.dbServers) || [];
        if (dbServers.some(s => /hana/i.test(`${s.component || ''} ${s.model || ''} ${s.os || ''}`))) return true;
        return false;
    },

    _renderHanaNotice(app) {
        const el = document.getElementById('twd-hana-notice');
        if (!el) return;
        if (this._hanaDetected(app) || this.hanaConfirmed) {
            el.innerHTML = `
                <div class="alert-info" style="margin-top:14px">
                    <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
                    <span><strong>SAP HANA database detected.</strong> The HANA analysis workflow is shown below.</span>
                </div>`;
        } else {
            el.innerHTML = `
                <div class="alert-info" style="margin-top:14px;align-items:center">
                    <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
                    <span style="flex:1">A HANA database was not auto-detected from the System Landscape. If this is a HANA system, you can proceed anyway.</span>
                    <button class="btn btn-sm" id="twd-force-hana">Proceed as HANA anyway</button>
                </div>`;
            const btn = document.getElementById('twd-force-hana');
            if (btn) btn.addEventListener('click', () => {
                this.hanaConfirmed = true;
                this._renderHanaNotice(app);
                this._refreshWorkflowVisibility();
            });
        }
    },

    _refreshWorkflowVisibility() {
        const show = this._hanaDetected(this.app) || this.hanaConfirmed;
        const up = document.getElementById('twd-uploads');
        if (up) up.style.display = show ? 'block' : 'none';
    },

    /* ════════════════════════════════════════════════════════
       FILE PARSING (TXT / CSV / XLSX)
       ════════════════════════════════════════════════════════ */
    /* TXT handled locally (delimiter sniff); CSV/XLSX delegate to the global
       parseUploadedFile. Column parsing elsewhere is by NAME, never position. */
    _parseFile(file, cb) {
        const ext = (file.name.split('.').pop() || '').toLowerCase();
        if (ext === 'txt' || ext === 'text' || ext === 'tsv') {
            const reader = new FileReader();
            reader.onload = e => {
                try {
                    const rows = this._parseDelimitedText(String(e.target.result || ''));
                    cb(rows, file.name);
                } catch (err) {
                    showToast('Parse failed', `Could not read ${file.name}: ${err.message}`, 'error');
                }
            };
            reader.onerror = () => showToast('Read failed', `Could not read ${file.name}.`, 'error');
            reader.readAsText(file);
            return;
        }
        /* CSV / XLSX / XLS */
        parseUploadedFile(file, (rows, name) => cb(rows, name));
    },

    /* Parse delimited text into array-of-objects. Sniffs the delimiter from the
       header line (tab, pipe, semicolon, then comma). Skips blank lines. */
    _parseDelimitedText(text) {
        const lines = text.split(/\r\n|\r|\n/).filter(l => l.trim() !== '');
        if (!lines.length) return [];
        const sniff = line => {
            const cands = ['\t', '|', ';', ','];
            let best = ',', bestN = 0;
            cands.forEach(d => { const n = line.split(d).length; if (n > bestN) { bestN = n; best = d; } });
            return best;
        };
        const delim = sniff(lines[0]);
        const split = line => line.split(delim).map(c => c.trim().replace(/^"(.*)"$/, '$1'));
        const headers = split(lines[0]);
        const out = [];
        for (let i = 1; i < lines.length; i++) {
            const cells = split(lines[i]);
            const row = {};
            headers.forEach((h, c) => { row[h] = cells[c] !== undefined ? cells[c] : ''; });
            out.push(row);
        }
        return out;
    },

    _onUpload(kind, rows, fileName) {
        if (!rows || !rows.length) { showToast('Empty file', `${fileName} contained no rows.`, 'error'); return; }
        const headers = Object.keys(rows[0]);
        const payload = { rows, headers, fileName };
        if (kind === 'landscape') this.landscape = payload;
        else if (kind === 'cpu')  this.cpuUpload = payload;
        else if (kind === 'mem')  this.memUpload = payload;
        else                      this.histUpload = payload;
        showFileLoaded(`twd-status-${kind}`, fileName);
        this._refreshValidation();
    },

    /* ════════════════════════════════════════════════════════
       COLUMN RESOLUTION (name-based, release-tolerant)
       ════════════════════════════════════════════════════════ */
    _findKey(headers, patterns) {
        for (const re of patterns) {
            const hit = headers.find(h => re.test(h));
            if (hit) return hit;
        }
        return null;
    },

    _num(v) {
        if (v == null) return NaN;
        const s = String(v).replace(/,/g, '').replace(/[^0-9.\-eE]/g, '');
        return s === '' ? NaN : parseFloat(s);
    },

    /* ════════════════════════════════════════════════════════
       VALIDATION
       ════════════════════════════════════════════════════════ */

    /* Required-column specs. Each entry: { label, pats, numeric?, required? }.
       pats = regex alternatives passed to _findKey (name-based, release-tolerant). */
    _landscapeCols() {
        return [
            { label: 'Database',          pats: [/^database$/i, /database(?!\s*sid)/i, /\bdb\b/i] },
            { label: 'Database SID',      pats: [/database\s*sid/i, /\bdb\s*sid/i, /^sid$/i] },
            { label: 'Database Release',  pats: [/database\s*release/i, /\brelease\b/i, /\bversion\b/i] },
            { label: 'Hostname',          pats: [/host\s*name/i, /^host$/i] },
            { label: 'CPU Model',         pats: [/cpu\s*model/i, /processor\s*model/i, /\bmodel\b/i] },
            { label: 'GHz',               pats: [/\bghz\b/i, /clock/i], numeric: true },
            { label: 'Physical CPUs',     pats: [/physical\s*cpus?/i, /phys.*cpu/i, /sockets?/i], numeric: true },
            { label: 'Logical CPUs',      pats: [/logical\s*cpus?/i, /log.*cpu/i, /threads?/i], numeric: true },
            { label: 'Cores',             pats: [/\bcores?\b/i], numeric: true },
            { label: 'Memory',            pats: [/memory/i, /\bram\b/i, /\bgb\b/i], numeric: true },
            { label: 'OS',                pats: [/operating\s*system/i, /^os$/i, /\bos\b/i] },
            { label: 'OS Level',          pats: [/os\s*level/i, /os\s*patch/i, /patch\s*level/i] },
            { label: 'HANA Revision',     pats: [/hana\s*rev/i, /revision/i] },
        ];
    },

    /* Validate one dataset against its column spec + generic data-quality checks.
       Returns { problems:[], warnings:[], ok:[], cols:{label:resolvedHeader} }.
       `mandatory` governs whether failures are blocking problems or warnings. */
    _validateDataset(name, upload, colSpecs, opts) {
        opts = opts || {};
        const problems = [], warnings = [], ok = [], cols = {};
        const push = opts.mandatory ? problems : warnings;

        if (!upload) {
            if (opts.mandatory) problems.push(`${name}: dataset has not been uploaded.`);
            return { problems, warnings, ok, cols };
        }
        const headers = upload.headers || [];
        const rows = upload.rows || [];

        /* Supported format produced a header + records. */
        if (!headers.length) { push(`${name}: no header row was parsed from ${upload.fileName}.`); return { problems, warnings, ok, cols }; }
        if (!rows.length)    { push(`${name}: no data records found in ${upload.fileName}.`); return { problems, warnings, ok, cols }; }

        /* Required columns present (by name). */
        const missing = [];
        colSpecs.forEach(spec => {
            const hit = this._findKey(headers, spec.pats);
            if (hit) cols[spec.label] = hit; else missing.push(spec.label);
        });
        if (missing.length) push(`${name}: missing required column(s): ${missing.join(', ')}.`);

        /* Numeric-conversion check on resolved numeric columns (sample all rows). */
        colSpecs.filter(s => s.numeric && cols[s.label]).forEach(spec => {
            const col = cols[spec.label];
            const bad = [];
            rows.forEach((r, i) => {
                const v = r[col];
                if (v == null || String(v).trim() === '') return;
                if (isNaN(this._num(v))) bad.push(i + 1);
            });
            if (bad.length) warnings.push(`${name}: column "${col}" has ${bad.length} non-numeric value(s) (row(s) ${bad.slice(0, 8).join(', ')}${bad.length > 8 ? '…' : ''}).`);
        });

        /* Missing mandatory field values per row (blank in a required column). */
        if (!missing.length) {
            const blanks = [];
            rows.forEach((r, i) => {
                const empty = colSpecs.some(s => { const c = cols[s.label]; return c && String(r[c] == null ? '' : r[c]).trim() === ''; });
                if (empty) blanks.push(i + 1);
            });
            if (blanks.length) warnings.push(`${name}: ${blanks.length} row(s) have blank mandatory field(s) (row(s) ${blanks.slice(0, 8).join(', ')}${blanks.length > 8 ? '…' : ''}).`);
        }

        if (!missing.length) ok.push(`${name}: ${rows.length.toLocaleString()} record(s), all required columns resolved.`);
        return { problems, warnings, ok, cols };
    },

    /* Timestamp + duplicate checks for a date(/time)-keyed dataset. */
    _checkTimeAndDupes(name, rows, dateCol, timeCol, warnings) {
        if (!rows || !rows.length || !dateCol) return;
        const seen = new Set();
        let badTs = 0, dupes = 0;
        rows.forEach(r => {
            const dv = String(r[dateCol] == null ? '' : r[dateCol]).trim();
            const tv = timeCol ? String(r[timeCol] == null ? '' : r[timeCol]).trim() : '';
            if (dv && this._timeKey(dv, tv) === 0) badTs++;
            const k = `${dv}|${tv}`;
            if (seen.has(k)) dupes++; else seen.add(k);
        });
        if (badTs)  warnings.push(`${name}: ${badTs} row(s) have an unparseable timestamp.`);
        if (dupes)  warnings.push(`${name}: ${dupes} duplicate timestamp record(s) detected.`);
    },

    _validate() {
        const groups = [];   /* [{ dataset, problems, warnings, ok }] */
        let problems = [], warnings = [], ok = [];

        /* Landscape SAPS basis (from Module 01 DB-layer hosts). */
        const basis = this._computeSapsBasis();
        this.sapsBasis = basis;

        /* 1 — System Landscape (mandatory). */
        const land = this._validateDataset('System Landscape', this.landscape, this._landscapeCols(), { mandatory: true });
        if (!basis || !basis.cores) land.problems.push('System Landscape database-layer host data (Cores / SAPS) is not available — complete Module 01 with a DB host.');
        else land.ok.push(`SAPS basis: ${basis.cores.toLocaleString()} cores · ${Math.round(basis.sapsPerCore).toLocaleString()} SAPS/core (${basis.hosts.length} DB host(s)).`);
        groups.push({ dataset: 'System Landscape Information', ...land });

        /* 2 — CPU Analysis (mandatory): DATE_, TIME, CPU_PCT. */
        const cpuSpec = [
            { label: 'DATE_',   pats: [/^date_?$/i, /date/i] },
            { label: 'TIME',    pats: [/^time$/i, /time/i] },
            { label: 'CPU_PCT', pats: [/cpu_?pct/i, /cpu.*%/i, /cpu/i], numeric: true },
        ];
        const cpu = this._validateDataset('CPU Analysis', this.cpuUpload, cpuSpec, { mandatory: true });
        if (this.cpuUpload) this._checkTimeAndDupes('CPU Analysis', this.cpuUpload.rows, cpu.cols['DATE_'], cpu.cols['TIME'], cpu.warnings);
        groups.push({ dataset: 'CPU Analysis (HANA_Resources_CPUAndMemory_History_*)', ...cpu });

        /* 3 — Memory Component Overview (mandatory): MY_DATE, COMPONENT, MEMORY_SIZE_GB. */
        const memSpec = [
            { label: 'MY_DATE',         pats: [/my_?date/i, /date/i] },
            { label: 'COMPONENT',       pats: [/component/i] },
            { label: 'MEMORY_SIZE_GB',  pats: [/memory_?size_?gb/i, /memory_?size/i, /size.*gb/i, /\bgb\b/i], numeric: true },
        ];
        const mem = this._validateDataset('Memory Component Overview', this.memUpload, memSpec, { mandatory: true });
        if (this.memUpload) this._checkTimeAndDupes('Memory Component Overview', this.memUpload.rows, mem.cols['MY_DATE'], mem.cols['COMPONENT'], mem.warnings);
        groups.push({ dataset: 'Memory Component Overview (7 Memory Component Overview)', ...mem });

        /* 4 — Historic EWA (optional): Date, Component, Memory + consistency check. */
        if (this.histUpload) {
            const histSpec = [
                { label: 'Date',      pats: [/my_?date/i, /date/i] },
                { label: 'Component', pats: [/component/i] },
                { label: 'Memory',    pats: [/memory_?size_?gb/i, /memory_?size/i, /size.*gb/i, /memory/i, /\bgb\b/i], numeric: true },
            ];
            const hist = this._validateDataset('Historic EWA Memory', this.histUpload, histSpec, { mandatory: false });
            this._checkTimeAndDupes('Historic EWA Memory', this.histUpload.rows, hist.cols['Date'], hist.cols['Component'], hist.warnings);
            /* Consistency vs Memory Component Overview component set. */
            if (hist.cols['Component'] && mem.cols['COMPONENT'] && this.memUpload) {
                const norm = s => String(s || '').trim().toLowerCase();
                const histComps = new Set(this.histUpload.rows.map(r => norm(r[hist.cols['Component']])).filter(Boolean));
                const memComps  = new Set(this.memUpload.rows.map(r => norm(r[mem.cols['COMPONENT']])).filter(Boolean));
                const onlyHist = [...histComps].filter(c => !memComps.has(c));
                const onlyMem  = [...memComps].filter(c => !histComps.has(c));
                if (onlyHist.length) hist.warnings.push(`Historic EWA Memory: ${onlyHist.length} component(s) not in the Memory Component Overview.`);
                if (onlyMem.length)  hist.warnings.push(`Historic EWA Memory: ${onlyMem.length} Memory Component Overview component(s) absent from the historic dataset.`);
            }
            groups.push({ dataset: 'Historic EWA Memory Components Overview (optional)', ...hist });
        }

        groups.forEach(g => { problems = problems.concat(g.problems); warnings = warnings.concat(g.warnings); ok = ok.concat(g.ok); });
        const canProceed = problems.length === 0;
        return { canProceed, problems, warnings, ok, groups };
    },

    _refreshValidation() {
        const v = this._validate();
        const el = document.getElementById('twd-validation');
        if (el) {
            const li = arr => arr.map(p => `<li>${p}</li>`).join('');
            const problemsHtml = v.problems.length ? `<div class="alert-warn" style="display:block">
                    <strong>Resolve these before you can proceed to charts:</strong>
                    <ul style="margin:6px 0 0 18px;padding:0">${li(v.problems)}</ul>
                </div>` : '';
            const warningsHtml = v.warnings.length ? `<div class="alert-info" style="display:block;margin-top:8px">
                    <strong>Warnings (non-blocking — review and correct the source data if needed):</strong>
                    <ul style="margin:6px 0 0 18px;padding:0">${li(v.warnings)}</ul>
                </div>` : '';
            const okHtml = v.ok.length ? `<div class="chart-hint" style="margin-top:8px"><ul style="margin:0 0 0 18px;padding:0">${li(v.ok)}</ul></div>` : '';
            el.innerHTML = problemsHtml + warningsHtml + okHtml;
        }
        const btn = document.getElementById('twd-btn-charts');
        if (btn) btn.disabled = !v.canProceed;
    },

    /* ════════════════════════════════════════════════════════
       SAPS BASIS — summed DB-layer System Landscape hosts
       ════════════════════════════════════════════════════════ */
    _computeSapsBasis() {
        const hv = (this.app && this.app.getHostVariables && this.app.getHostVariables()) || [];
        const dbHosts = hv.filter(h => h.layer === 'db');
        if (!dbHosts.length) return null;
        let cores = 0, totalSaps = 0;
        dbHosts.forEach(h => {
            const c = parseInt(h.cores) || 0;
            cores += c;
            totalSaps += Number(h.totalSaps) || (c * (Number(h.sapsPerCore) || 0));
        });
        if (!cores) return null;
        const sapsPerCore = totalSaps / cores;   /* weighted */
        return { cores, sapsPerCore, totalSaps, hosts: dbHosts.map(h => h.hostname) };
    },

    /* ════════════════════════════════════════════════════════
       CPU PROCESSING — append SAPS columns, sort ascending by time
       ════════════════════════════════════════════════════════ */
    _processCpu() {
        if (!this.cpuUpload) return;
        const rows = this.cpuUpload.rows;
        const h = this.cpuUpload.headers;
        const cDate = this._findKey(h, [/^date_?$/i, /date/i]);
        const cTime = this._findKey(h, [/^time$/i, /time/i]);
        const cCpu  = this._findKey(h, [/cpu_?pct/i, /cpu.*%/i, /cpu/i]);
        const basis = this.sapsBasis || this._computeSapsBasis() || { cores: 0, sapsPerCore: 0 };

        const cores = basis.cores;
        const spc   = basis.sapsPerCore;
        const sapsDeployed = cores * spc;

        const r2 = v => Math.round(v * 100) / 100;
        const enriched = rows.map(r => {
            const cpuPct = this._num(r[cCpu]);
            const sapsUtil = isNaN(cpuPct) ? NaN : (cpuPct / 100) * sapsDeployed;
            const ts = this._timeKey(r[cDate], cTime ? r[cTime] : '');
            return {
                __ts: ts,
                __label: `${r[cDate] || ''} ${cTime ? (r[cTime] || '') : ''}`.trim(),
                'DATE_': r[cDate] || '',
                'TIME': cTime ? (r[cTime] || '') : '',
                'CPU_PCT': isNaN(cpuPct) ? '' : cpuPct,
                'Cores': cores,
                'SAPS/Core': r2(spc),
                'SAPS Deployed': r2(sapsDeployed),
                'SAPS Utilization': isNaN(sapsUtil) ? '' : r2(sapsUtil),
            };
        });
        /* Ascending by timestamp (oldest left). */
        enriched.sort((a, b) => a.__ts - b.__ts);

        this.targetCols = ['DATE_', 'TIME', 'CPU_PCT', 'Cores', 'SAPS/Core', 'SAPS Deployed', 'SAPS Utilization'];
        this.targetRows = enriched;
    },

    /* Parse "DATE_ + TIME" into a sortable numeric key. Accepts common HANA
       formats (YYYY-MM-DD / YYYYMMDD / DD.MM.YYYY) + HH:MM[:SS]. Falls back to a
       lexical key so sorting still works. */
    _timeKey(date, time) {
        const d = String(date || '').trim();
        const t = String(time || '').trim();
        let y, mo, da;
        let m;
        if ((m = d.match(/^(\d{4})-(\d{2})-(\d{2})/))) { y = +m[1]; mo = +m[2]; da = +m[3]; }
        else if ((m = d.match(/^(\d{4})(\d{2})(\d{2})$/))) { y = +m[1]; mo = +m[2]; da = +m[3]; }
        else if ((m = d.match(/^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})/))) { da = +m[1]; mo = +m[2]; y = +m[3]; }
        if (y == null) {
            /* Unparseable date — use a stable lexical fallback. */
            return Number(`${d}${t}`.replace(/\D/g, '').slice(0, 14)) || 0;
        }
        const tm = t.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
        const hh = tm ? +tm[1] : 0, mi = tm ? +tm[2] : 0, ss = tm ? (+tm[3] || 0) : 0;
        const ms = Date.UTC(y, mo - 1, da, hh, mi, ss);
        return isNaN(ms) ? 0 : ms;
    },

    /* ════════════════════════════════════════════════════════
       MEMORY PROCESSING — normalize to stacked-area series
       ════════════════════════════════════════════════════════ */
    _processMemory() {
        this.memNormalized = this._normalizeMemory(this.memUpload);
    },

    /* Historic EWA — normalize to the same {labels,components,series} shape, kept
       INDEPENDENT of memNormalized. Only built when the optional file is present. */
    _processHist() {
        this.histNormalized = this.histUpload ? this._normalizeMemory(this.histUpload) : null;
    },

    /* Shared normalizer: Date / Component / Memory → stacked-area series.
       Components are discovered dynamically (never hardcoded); dates sorted asc. */
    _normalizeMemory(upload) {
        if (!upload) return null;
        const rows = upload.rows;
        const h = upload.headers;
        const cDate = this._findKey(h, [/my_?date/i, /date/i]);
        const cComp = this._findKey(h, [/component/i]);
        const cSize = this._findKey(h, [/memory_?size_?gb/i, /memory_?size/i, /size.*gb/i, /memory/i, /\bgb\b/i]);
        if (!cDate || !cComp || !cSize) return null;

        /* Collect distinct dates (with sort key) and components. */
        const dateKey = {};               /* label -> numeric key */
        const byDateComp = {};            /* label -> { component -> gb } */
        const components = [];
        rows.forEach(r => {
            const lbl = String(r[cDate] || '').trim();
            if (!lbl) return;
            const comp = String(r[cComp] || 'Other').trim() || 'Other';
            const gb = this._num(r[cSize]);
            if (!(lbl in dateKey)) { dateKey[lbl] = this._timeKey(lbl, ''); byDateComp[lbl] = {}; }
            if (!components.includes(comp)) components.push(comp);
            byDateComp[lbl][comp] = (byDateComp[lbl][comp] || 0) + (isNaN(gb) ? 0 : gb);
        });

        const labels = Object.keys(dateKey).sort((a, b) => dateKey[a] - dateKey[b]);
        const series = {};
        components.forEach(comp => {
            series[comp] = labels.map(lbl => {
                const v = byDateComp[lbl][comp];
                return v == null ? null : Math.round(v * 100) / 100;
            });
        });

        return { labels, components, series };
    },

    /* ════════════════════════════════════════════════════════
       CHARTS VIEW
       ════════════════════════════════════════════════════════ */
    async _showChartsView(app) {
        /* Build the derived datasets first (also validates). */
        this.sapsBasis = this._computeSapsBasis();
        this._processCpu();
        this._processMemory();
        this._processHist();

        document.getElementById('twd-collect').style.display = 'none';
        document.getElementById('twd-charts').style.display = 'block';
        document.getElementById('content-area')?.scrollTo({ top: 0, behavior: 'smooth' });

        app.markCompleted('tech-workload-db');

        this._wireChartTabs(app);
        this._renderedSection = null;
        await this._activateSection(app, 'cpu', true);
    },

    _showCollectView() {
        document.getElementById('twd-charts').style.display = 'none';
        document.getElementById('twd-collect').style.display = 'block';
        this._teardownCharts();
    },

    _wireChartTabs(app) {
        const strip = document.getElementById('twd-chart-tabs');
        if (!strip || strip._wired) return;
        strip._wired = true;
        strip.querySelectorAll('.chart-tab').forEach(btn => {
            btn.addEventListener('click', () => this._activateSection(app, btn.dataset.tab));
        });
    },

    async _activateSection(app, sectionId, force = false) {
        const SECTIONS = ['cpu', 'mem'];
        if (!SECTIONS.includes(sectionId)) sectionId = 'cpu';

        if (!force && this._renderedSection === sectionId) { this._syncTabUi(sectionId); return; }

        const token = (this._activateToken = (this._activateToken || 0) + 1);
        const current = () => this._activateToken === token;

        this._teardownCharts();
        this._syncTabUi(sectionId);

        const label = { cpu: 'CPU', mem: 'memory' }[sectionId];
        try {
            this._chartProgress(`Preparing ${label} charts…`, 20);
            await this._nextFrame();
            if (!current()) return;
            this._chartProgress(`Building ${label} charts…`, 60);
            await this._nextFrame();
            if (!current()) return;

            if (sectionId === 'cpu') this._renderCpuCharts();
            else                     this._renderMemCharts();

            if (!current()) return;
            this._renderedSection = sectionId;
            this._chartProgress('Charts ready', 100);
            await this._nextFrame();
        } catch (err) {
            showToast('Chart generation issue', (err && err.message) ? err.message : String(err), 'error');
        } finally {
            if (current()) setTimeout(() => { if (current()) this._hideChartProgress(); }, 300);
        }
    },

    _syncTabUi(sectionId) {
        document.querySelectorAll('#twd-chart-tabs .chart-tab').forEach(btn => {
            btn.classList.toggle('is-active', btn.dataset.tab === sectionId);
        });
        document.querySelectorAll('#twd-charts .chart-section').forEach(sec => {
            sec.style.display = sec.dataset.sec === sectionId ? 'block' : 'none';
        });
    },

    _teardownCharts() {
        if (this._charts) Object.values(this._charts).forEach(e => { try { e.chart.destroy(); } catch (_) {} });
        this._charts = {};
    },

    _chartProgress(label, pct) {
        const box = document.getElementById('twd-chart-progress');
        const lab = document.getElementById('twd-chart-progress-label');
        const pe  = document.getElementById('twd-chart-progress-pct');
        const fill = document.getElementById('twd-chart-progress-fill');
        if (box) box.style.display = 'block';
        if (lab) lab.textContent = label;
        const p = Math.max(0, Math.min(100, Math.round(pct)));
        if (pe) pe.textContent = `${p}%`;
        if (fill) fill.style.width = `${p}%`;
    },
    _hideChartProgress() {
        const box = document.getElementById('twd-chart-progress');
        if (box) box.style.display = 'none';
    },
    _nextFrame() {
        return new Promise(resolve => requestAnimationFrame(() => setTimeout(resolve, 0)));
    },

    /* ── CPU charts ── */
    _renderCpuCharts() {
        const sec = document.getElementById('twd-sec-cpu');
        if (!sec) return;
        sec.innerHTML = '';
        if (!this.targetRows.length) {
            sec.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">No CPU data available.</div></div>`;
            return;
        }
        /* Downsample to ~300 points for a readable, responsive on-screen chart
           (clustered columns with thousands of 5-min snapshots are unusable).
           The full dataset is preserved in targetRows, the data table, and the
           workbook Data sheet. */
        const pick = this._downsample(this.targetRows, 300);
        const labels = pick.map(i => this.targetRows[i].__label);
        const cpu = pick.map(i => this.targetRows[i]['CPU_PCT'] === '' ? null : Number(this.targetRows[i]['CPU_PCT']));
        const sapsUtil = pick.map(i => this.targetRows[i]['SAPS Utilization'] === '' ? null : Number(this.targetRows[i]['SAPS Utilization']));
        const sapsDep  = pick.map(i => Number(this.targetRows[i]['SAPS Deployed']) || 0);

        /* Chart 1 — Clustered column: CPU_PCT (%) */
        this._mountClusteredColumn(sec, {
            id: 'hana-cpu-pct', fileSection: 'cpu', fileSeq: 1,
            title: 'HANA CPU Utilization per 5 minute snapshots',
            labels,
            namedSeries: [{ label: 'CPU Utilization (%)', data: cpu, color: '#0070F2' }],
            yLabel: 'CPU Utilization (%)', yMin: 0, yMax: 100, xLabel: 'Date / Time',
            caption: 'Verify that the chart title reflects the TIME_AGGREGATE_BY value used during script execution.',
            snapshotReminder: true,
        });

        /* Chart 2 — Combo: SAPS Utilization (blue columns) + SAPS Deployed (red line) */
        const yMax = this._niceCeil(Math.max(sapsDep[0] || 0, ...sapsUtil.filter(v => v != null)) * 1.1) || 1000;
        this._mountComboChart(sec, {
            id: 'hana-cpu-saps', fileSection: 'cpu', fileSeq: 2,
            title: 'HANA CPU Utilization per 5 minute snapshots',
            labels,
            barLabel: 'SAPS Utilization', barData: sapsUtil, barColor: '#0070F2',
            lineLabel: 'SAPS Deployed', lineData: sapsDep, lineColor: '#EF4444',
            yLabel: 'SAPS', yMax, xLabel: 'Date / Time',
            caption: 'Verify that the chart title reflects the TIME_AGGREGATE_BY value used during script execution.',
            snapshotReminder: true,
        });
    },

    /* ── Memory charts ── */
    _renderMemCharts() {
        const sec = document.getElementById('twd-sec-mem');
        if (!sec) return;
        sec.innerHTML = '';
        const mn = this.memNormalized;
        if (!mn || !mn.labels.length || !mn.components.length) {
            sec.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">No Memory Component Overview data available.</div></div>`;
            return;
        }
        /* Commentary: analysis period + days. */
        const first = mn.labels[0], last = mn.labels[mn.labels.length - 1];
        const k0 = this._timeKey(first, ''), k1 = this._timeKey(last, '');
        let daysTxt = '';
        if (k1 > k0) daysTxt = `${Math.round((k1 - k0) / 86400000) + 1}`;
        else daysTxt = `${mn.labels.length}`;
        const commentary = `
            <div class="chart-stats">
                <div class="chart-stats-title">Analysis Period</div>
                <div class="chart-stats-grid">
                    <div class="chart-stat"><div class="chart-stat-label">Start Date</div><div class="chart-stat-value">${first}</div></div>
                    <div class="chart-stat"><div class="chart-stat-label">End Date</div><div class="chart-stat-value">${last}</div></div>
                    <div class="chart-stat"><div class="chart-stat-label">Days Included</div><div class="chart-stat-value">${daysTxt}</div></div>
                </div>
            </div>`;

        this._mountStackedAreaChart(sec, {
            id: 'hana-mem', fileSection: 'mem', fileSeq: 1,
            title: 'Master Node - Memory Components',
            labels: mn.labels,
            namedSeries: mn.components.map((c, i) => ({ label: c, data: mn.series[c], color: CHART_COLORS[i % CHART_COLORS.length] })),
            yLabel: 'MEMORY_SIZE_GB', xLabel: 'MY_DATE',
            analysisHtml: commentary,
        });
    },

    /* ════════════════════════════════════════════════════════
       CHART MOUNTS (built on the App-layer skeleton)
       ════════════════════════════════════════════════════════ */
    _chartToolbar(uid, cfg, extraControls) {
        const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
        return `
            <div class="chart-toolbar">
                <input class="chart-title-input" id="${uid}-title" value="${esc(cfg.title)}" title="Edit chart title">
                <div class="chart-controls">
                    ${extraControls || ''}
                    <button class="btn btn-sm" id="${uid}-copy" title="Copy a high-resolution image to the clipboard">Copy</button>
                    <button class="btn btn-sm" id="${uid}-png" title="Download a high-resolution PNG">PNG</button>
                    <button class="btn btn-sm" id="${uid}-svg" title="Download a vector-wrapped SVG">SVG</button>
                    <button class="btn btn-sm" id="${uid}-csv" title="Download chart data as CSV">CSV</button>
                    <button class="btn btn-sm chart-collapse-btn" id="${uid}-collapse" title="Hide / show this chart">–</button>
                </div>
            </div>`;
    },

    _wireChartCommon(card, uid, cfg) {
        card.querySelector(`#${uid}-title`).addEventListener('input', e => { this._charts[uid].cfg.title = e.target.value; });
        card.querySelector(`#${uid}-copy`).addEventListener('click', () => {
            this._copyChartPng(uid);
            if (cfg.snapshotReminder) showToast('Verify snapshot interval', 'Confirm the snapshot interval named in the chart title matches the TIME_AGGREGATE_BY value you used.', 'info', 7000);
        });
        card.querySelector(`#${uid}-png`).addEventListener('click', () => this._exportChartPng(uid));
        card.querySelector(`#${uid}-svg`).addEventListener('click', () => this._exportChartSvg(uid));
        card.querySelector(`#${uid}-csv`).addEventListener('click', () => this._exportChartCsv(uid));
        card.querySelector('.chart-canvas-wrap').insertAdjacentHTML('afterend', chartFontRowHtml(uid));
        wireChartFontRow(card, () => [this._charts[uid] && this._charts[uid].chart].filter(Boolean), null, false);
        this._wireChartCollapse(card, uid);
    },

    /* Clustered vertical columns (one or more series side-by-side). */
    _mountClusteredColumn(container, cfg) {
        this._charts = this._charts || {};
        const uid = 'twd-ch-' + cfg.id;
        const specs = cfg.namedSeries.map((s, i) => ({ label: s.label, data: s.data, color: s.color || CHART_COLORS[i % CHART_COLORS.length] }));
        const card = document.createElement('div');
        card.className = 'card chart-card-wrap';
        card.innerHTML =
            this._chartToolbar(uid, cfg,
                `<label>Y max <input type="number" id="${uid}-ymax" value="${cfg.yMax}" step="10"></label>
                 <button class="btn btn-sm" id="${uid}-reset" title="Reset axes">Reset</button>`) +
            `<div class="chart-canvas-wrap chart-tall-wrap"><canvas id="${uid}-canvas"></canvas></div>` +
            (cfg.caption ? `<div class="chart-hint">${cfg.caption}</div>` : '') +
            (cfg.analysisHtml || '');
        container.appendChild(card);

        const ctx = card.querySelector(`#${uid}-canvas`).getContext('2d');
        const chart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: cfg.labels,
                datasets: specs.map(s => ({ label: s.label, data: s.data, backgroundColor: s.color, borderColor: s.color, borderWidth: 0 })),
            },
            options: {
                responsive: true, maintainAspectRatio: false, animation: false,
                plugins: { legend: { position: 'bottom', labels: { font: { size: 10 }, padding: 12, usePointStyle: true } } },
                scales: {
                    x: { stacked: false, title: { display: true, text: cfg.xLabel || '' }, ticks: { maxRotation: 60, autoSkip: true, maxTicksLimit: 16, font: { size: 9 } }, grid: { display: false } },
                    y: { stacked: false, title: { display: true, text: cfg.yLabel || '' }, min: cfg.yMin, max: cfg.yMax, grid: { color: '#EEF2F7' } },
                },
            },
        });
        this._charts[uid] = { chart, cfg: { ...cfg, namedSeries: specs } };

        const applyY = () => { const mx = parseFloat(card.querySelector(`#${uid}-ymax`).value); chart.options.scales.y.max = isNaN(mx) ? undefined : mx; chart.update(); };
        card.querySelector(`#${uid}-ymax`).addEventListener('change', applyY);
        card.querySelector(`#${uid}-reset`).addEventListener('click', () => { card.querySelector(`#${uid}-ymax`).value = cfg.yMax; chart.options.scales.y.max = cfg.yMax; chart.update(); });
        this._wireChartCommon(card, uid, cfg);
    },

    /* Dual-axis combo: clustered columns (left axis) + line (right axis). */
    _mountComboChart(container, cfg) {
        this._charts = this._charts || {};
        const uid = 'twd-ch-' + cfg.id;
        const card = document.createElement('div');
        card.className = 'card chart-card-wrap';
        card.innerHTML =
            this._chartToolbar(uid, cfg,
                `<label>Y max <input type="number" id="${uid}-ymax" value="${cfg.yMax}" step="100"></label>
                 <button class="btn btn-sm" id="${uid}-reset" title="Reset axes">Reset</button>`) +
            `<div class="chart-canvas-wrap chart-tall-wrap"><canvas id="${uid}-canvas"></canvas></div>` +
            (cfg.caption ? `<div class="chart-hint">${cfg.caption}</div>` : '') +
            (cfg.analysisHtml || '');
        container.appendChild(card);

        const ctx = card.querySelector(`#${uid}-canvas`).getContext('2d');
        const chart = new Chart(ctx, {
            data: {
                labels: cfg.labels,
                datasets: [
                    { type: 'bar', label: cfg.barLabel, yAxisID: 'y', data: cfg.barData, backgroundColor: cfg.barColor, borderColor: cfg.barColor, borderWidth: 0 },
                    { type: 'line', label: cfg.lineLabel, yAxisID: 'y', data: cfg.lineData, borderColor: cfg.lineColor, backgroundColor: cfg.lineColor, borderWidth: 2, pointRadius: 0, pointHoverRadius: 3, tension: 0, spanGaps: true },
                ],
            },
            options: {
                responsive: true, maintainAspectRatio: false, animation: false,
                interaction: { mode: 'index', intersect: false },
                plugins: { legend: { position: 'bottom', labels: { font: { size: 10 }, padding: 12, usePointStyle: true } } },
                scales: {
                    x: { title: { display: true, text: cfg.xLabel || '' }, ticks: { maxRotation: 60, autoSkip: true, maxTicksLimit: 16, font: { size: 9 } }, grid: { display: false } },
                    y: { type: 'linear', position: 'left', min: 0, max: cfg.yMax, title: { display: true, text: cfg.yLabel || 'SAPS' }, grid: { color: '#EEF2F7' }, ticks: { font: { size: 10 } } },
                },
            },
        });
        this._charts[uid] = { chart, cfg: { ...cfg, namedSeries: [{ label: cfg.barLabel, data: cfg.barData }, { label: cfg.lineLabel, data: cfg.lineData }] } };

        const applyY = () => { const mx = parseFloat(card.querySelector(`#${uid}-ymax`).value); chart.options.scales.y.max = isNaN(mx) ? undefined : mx; chart.update(); };
        card.querySelector(`#${uid}-ymax`).addEventListener('change', applyY);
        card.querySelector(`#${uid}-reset`).addEventListener('click', () => { card.querySelector(`#${uid}-ymax`).value = cfg.yMax; chart.options.scales.y.max = cfg.yMax; chart.update(); });
        this._wireChartCommon(card, uid, cfg);
    },

    /* Stacked area: one filled series per component, stacked on the Y axis. */
    _mountStackedAreaChart(container, cfg) {
        this._charts = this._charts || {};
        const uid = 'twd-ch-' + cfg.id;
        const specs = cfg.namedSeries.map((s, i) => ({ label: s.label, data: s.data, color: s.color || CHART_COLORS[i % CHART_COLORS.length] }));
        const card = document.createElement('div');
        card.className = 'card chart-card-wrap';
        card.innerHTML =
            this._chartToolbar(uid, cfg,
                `<button class="btn btn-sm" id="${uid}-reset" title="Reset">Reset</button>`) +
            `<div class="chart-canvas-wrap chart-tall-wrap"><canvas id="${uid}-canvas"></canvas></div>` +
            `<div class="chart-hint">Stacked area — one series per memory component. Click a component in the legend to show/hide it.</div>` +
            (cfg.analysisHtml || '');
        container.appendChild(card);

        const hexA = (hex, a) => {
            const m = /^#?([0-9a-f]{6})$/i.exec(hex || '');
            if (!m) return hex;
            const n = parseInt(m[1], 16);
            return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${a})`;
        };

        const ctx = card.querySelector(`#${uid}-canvas`).getContext('2d');
        const chart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: cfg.labels,
                datasets: specs.map(s => ({
                    label: s.label, data: s.data,
                    borderColor: s.color, backgroundColor: hexA(s.color, 0.55),
                    borderWidth: 1, pointRadius: 0, pointHoverRadius: 3, tension: 0.2,
                    fill: true, spanGaps: true,
                })),
            },
            options: {
                responsive: true, maintainAspectRatio: false, animation: false,
                interaction: { mode: 'index', intersect: false },
                plugins: { legend: { position: 'bottom', labels: { font: { size: 10 }, padding: 10, usePointStyle: true } } },
                scales: {
                    x: { title: { display: true, text: cfg.xLabel || 'MY_DATE' }, ticks: { maxRotation: 60, autoSkip: true, maxTicksLimit: 16, font: { size: 9 } }, grid: { display: false } },
                    y: { stacked: true, min: 0, title: { display: true, text: cfg.yLabel || 'MEMORY_SIZE_GB' }, grid: { color: '#EEF2F7' }, ticks: { font: { size: 10 } } },
                },
            },
        });
        this._charts[uid] = { chart, cfg: { ...cfg, namedSeries: specs } };

        card.querySelector(`#${uid}-reset`).addEventListener('click', () => {
            chart.data.datasets.forEach((ds, i) => { ds.hidden = false; const meta = chart.getDatasetMeta(i); if (meta) meta.hidden = false; });
            chart.update();
        });
        this._wireChartCommon(card, uid, cfg);
    },

    /* ════════════════════════════════════════════════════════
       COPY / EXPORT  (ported from the App-layer module)
       ════════════════════════════════════════════════════════ */
    _wireChartCollapse(card, uid) {
        const btn = card.querySelector(`#${uid}-collapse`);
        if (!btn) return;
        btn.addEventListener('click', () => {
            const collapsed = card.classList.toggle('chart-collapsed');
            btn.textContent = collapsed ? '+' : '–';
            btn.title = collapsed ? 'Show this chart' : 'Hide this chart';
        });
    },

    _niceCeil(v) {
        if (!(v > 0)) return 100;
        const exp = Math.floor(Math.log10(v)) - 1;
        const base = Math.pow(10, exp);
        const f = v / base;
        const steps = [10, 12, 15, 20, 25, 30, 40, 50, 60, 75, 80, 100];
        let step = 100;
        for (const s of steps) { if (f <= s) { step = s; break; } }
        return Math.round(step * base);
    },

    _sidTag() {
        const si = this._systemInfo(this.app);
        const sid = si.sid || '';
        return String(sid).replace(/[^A-Za-z0-9]/g, '').slice(0, 3).toLowerCase() || 'sap';
    },

    _chartFileBase(cfg) {
        const now = new Date();
        const dd = String(now.getDate()).padStart(2, '0');
        const mm = String(now.getMonth() + 1).padStart(2, '0');
        const yy = String(now.getFullYear()).slice(-2);
        return `${this._sidTag()}_wa_db_${cfg.fileSection || 'db'}_${cfg.fileSeq || 1}_${dd}-${mm}-${yy}`;
    },

    _chartToCanvas(entry, scale) {
        const src = entry.chart && entry.chart.canvas;
        if (!src) throw new Error('chart canvas unavailable');
        const uid = 'twd-ch-' + entry.cfg.id;
        const titleEl = document.getElementById(`${uid}-title`);
        const title = (titleEl && titleEl.value) || entry.cfg.title || '';
        const titleSize = ChartFontSizeManager.get().title || 18;
        const titleH = Math.round(titleSize * 1.9);
        const canvas = document.createElement('canvas');
        canvas.width = src.width * scale;
        canvas.height = src.height * scale + titleH * scale;
        const ctx = canvas.getContext('2d');
        ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = 'high';
        ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
        if (title) {
            ctx.fillStyle = '#1E293B';
            ctx.font = `700 ${titleSize * scale}px Arial, Helvetica, sans-serif`;
            ctx.textBaseline = 'top'; ctx.textAlign = 'center';
            ctx.fillText(title, canvas.width / 2, Math.round(8 * scale));
            ctx.textAlign = 'left';
        }
        ctx.drawImage(src, 0, titleH * scale, src.width * scale, src.height * scale);
        return canvas;
    },
    _svgToCanvas(entry, scale) {
        try { return Promise.resolve(this._chartToCanvas(entry, scale)); }
        catch (err) { return Promise.reject(err); }
    },
    _downloadBlob(blob, name, toastMsg) {
        const url = URL.createObjectURL(blob);
        const a = Object.assign(document.createElement('a'), { href: url, download: name });
        document.body.appendChild(a); a.click(); document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        if (toastMsg) showToast('Chart saved', toastMsg, 'success');
    },
    _copyChartPng(uid) {
        const entry = this._charts && this._charts[uid];
        if (!entry) return;
        const base = this._chartFileBase(entry.cfg);
        const makeBlob = () => this._svgToCanvas(entry, 4).then(canvas =>
            new Promise((resolve, reject) => { canvas.toBlob(b => b ? resolve(b) : reject(new Error('render failed')), 'image/png'); }));
        const fallbackDownload = (blob, why) => this._downloadBlob(blob, `${base}.png`, why || 'Copy blocked — downloaded a high-res PNG instead.');
        if (navigator.clipboard && window.ClipboardItem) {
            let captured = null;
            const item = new ClipboardItem({ 'image/png': makeBlob().then(b => { captured = b; return b; }) });
            try { window.focus(); } catch (_) {}
            Promise.resolve(navigator.clipboard.write([item]))
                .then(() => showToast('Chart copied', 'Chart image copied — paste it into Word, PowerPoint, or an email.', 'success'))
                .catch(err => {
                    const retry = captured ? Promise.resolve(captured) : makeBlob();
                    retry.then(b => navigator.clipboard.write([new ClipboardItem({ 'image/png': b })])
                        .then(() => showToast('Chart copied', 'Chart image copied — paste it into Word, PowerPoint, or an email.', 'success'))
                        .catch(() => fallbackDownload(b, `Copy blocked (${err && err.name ? err.name : 'clipboard'}) — downloaded a high-res PNG instead.`)))
                        .catch(() => showToast('Copy failed', 'Could not render the chart image.', 'error'));
                });
        } else {
            makeBlob().then(b => fallbackDownload(b, 'Clipboard image not supported — downloaded a high-res PNG instead.'))
                .catch(() => showToast('Copy failed', 'Could not render the chart image.', 'error'));
        }
    },
    _exportChartPng(uid) {
        const entry = this._charts && this._charts[uid];
        if (!entry) return;
        const base = this._chartFileBase(entry.cfg);
        this._svgToCanvas(entry, 4).then(canvas => {
            canvas.toBlob(blob => {
                if (!blob) { showToast('Export failed', 'Could not render the image.', 'error'); return; }
                this._downloadBlob(blob, `${base}.png`, 'High-resolution PNG saved.');
            }, 'image/png');
        }).catch(() => showToast('Export failed', 'Could not render the image.', 'error'));
    },
    _exportChartSvg(uid) {
        const entry = this._charts && this._charts[uid];
        if (!entry) return;
        this._svgToCanvas(entry, 4).then(canvas => {
            const w = canvas.width, h = canvas.height;
            const png = canvas.toDataURL('image/png');
            const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w / 4}" height="${h / 4}" viewBox="0 0 ${w} ${h}"><image width="${w}" height="${h}" href="${png}"/></svg>`;
            this._downloadBlob(new Blob([svg], { type: 'image/svg+xml;charset=utf-8' }), `${this._chartFileBase(entry.cfg)}.svg`, 'Vector-wrapped SVG saved (chart type and colours preserved).');
        }).catch(() => showToast('Export failed', 'Could not render the chart image.', 'error'));
    },
    _exportChartCsv(uid) {
        const entry = this._charts && this._charts[uid];
        if (!entry) return;
        const cfg = entry.cfg;
        const labels = cfg.labels || [];
        const series = (cfg.namedSeries || []).map(s => ({ label: s.label, data: s.data }));
        const cols = ['Date/Time', ...series.map(s => s.label)];
        const rows = [cols.join(',')];
        labels.forEach((lbl, i) => {
            const vals = series.map(s => { const v = s.data ? s.data[i] : ''; return v == null ? '' : v; });
            rows.push([`"${lbl}"`, ...vals].join(','));
        });
        this._downloadBlob(new Blob([rows.join('\r\n')], { type: 'text/csv;charset=utf-8;' }), `${this._chartFileBase(cfg)}.csv`, 'Chart data exported as CSV.');
    },

    /* ════════════════════════════════════════════════════════
       WORKBOOK — <DatabaseSID>_database.xlsx (native charts)
       ════════════════════════════════════════════════════════ */
    _downsample(labels, max) {
        const n = Array.isArray(labels) ? labels.length : 0;
        if (!Number.isFinite(n) || n <= 0) return [];
        if (!Number.isFinite(max) || max < 1) max = 1;
        if (n <= max) { const a = []; for (let i = 0; i < n; i++) a.push(i); return a; }
        const idx = []; const step = (n - 1) / (max - 1);
        for (let k = 0; k < max; k++) { const i = Math.round(k * step); if (i >= 0 && i < n) idx.push(i); }
        return Array.from(new Set(idx));
    },

    async _buildWorkbook() {
        const btn = document.getElementById('twd-wb-top');
        if (btn) btn.disabled = true;
        try {
            this._chartProgress('Collecting data…', 8);
            await this._nextFrame();

            /* Make sure derived datasets are current. */
            if (!this.targetRows.length) { this.sapsBasis = this._computeSapsBasis(); this._processCpu(); }
            if (!this.memNormalized) this._processMemory();

            const wb = XLSX.utils.book_new();
            const descriptors = [];
            /* Every worksheet opens with a 4-row documentation block
               (Data Source / Collection Method / Source Script + blank), then the
               column-header row (row 5), so data begins on row 6. */
            const FIRST_DATA_ROW = 6;
            const colLetter = n => { let s = ''; n += 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
            const MAX_PTS = 300;

            /* Standard 3-line documentation header + a spacer row. */
            const docHeader = (dataSource, collectionMethod, sourceScript) => ([
                ['Data Source', dataSource],
                ['Collection Method', collectionMethod],
                ['Source Script', sourceScript],
                [],
            ]);

            /* ── Source-data worksheets (preserve uploads), documented at the top. ── */
            const addSource = (name, upload, doc) => {
                if (!upload || !upload.rows.length) return;
                try {
                    const headers = upload.headers;
                    const aoa = [
                        ...docHeader(doc.src, doc.method, doc.script),
                        headers,
                        ...upload.rows.slice(0, 100000).map(r => headers.map(h => { const v = r[h]; return v == null ? '' : v; })),
                    ];
                    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name.slice(0, 31));
                } catch (err) { console.warn('[TWDB workbook] source sheet failed', name, err); }
            };
            addSource('Landscape_Source', this.landscape, {
                src: 'RSxxxSTT → HANA → HANA Initial Capacity Check → 1 System / Landscape Information',
                method: `Manual export from the HANA Initial Capacity Check (file: ${this.landscape ? this.landscape.fileName : ''})`,
                script: 'RSxxxSTT (HANA Initial Capacity Check)',
            });
            addSource('CPU_Source', this.cpuUpload, {
                src: 'HANA_Resources_CPUAndMemory_History_* (SAP Note 1969700)',
                method: `SQL script export, TIME_AGGREGATE_BY as configured (file: ${this.cpuUpload ? this.cpuUpload.fileName : ''})`,
                script: 'HANA_Resources_CPUAndMemory_History_* (SAP Note 1969700)',
            });
            addSource('Memory_Source', this.memUpload, {
                src: 'RSxxxSTT → HANA → HANA Initial Capacity Check → 7 Memory Component Overview',
                method: `Manual export from the HANA Initial Capacity Check (file: ${this.memUpload ? this.memUpload.fileName : ''})`,
                script: 'RSxxxSTT (HANA Initial Capacity Check)',
            });
            if (this.histUpload) addSource('Hist_Memory_Source', this.histUpload, {
                src: 'Historic EWA Memory Components Overview (EarlyWatch Alert Dashboard)',
                method: `Manual export from the EWA Dashboard — daily granularity, ≥12 months (file: ${this.histUpload.fileName})`,
                script: 'EarlyWatch Alert Dashboard export',
            });

            this._chartProgress('Writing CPU worksheets…', 35);
            await this._nextFrame();

            /* ── CPU_Data + CPU_Charts (clustered column + combo) ── */
            if (this.targetRows.length) {
                const pick = this._downsample(this.targetRows, MAX_PTS);
                const rowsP = pick.map(i => this.targetRows[i]);
                const header = ['Date / Time', 'CPU_PCT', 'SAPS Utilization', 'SAPS Deployed'];
                const dataAoa = [
                    ...docHeader(
                        'HANA_Resources_CPUAndMemory_History_* (SAP Note 1969700) — transformed; Cores/SAPS from the System Landscape (Module 01)',
                        'SQL script export aggregated by TIME_AGGREGATE_BY; SAPS columns derived from the landscape DB-layer hosts',
                        'HANA_Resources_CPUAndMemory_History_* (SAP Note 1969700)'),
                    header,
                    ...rowsP.map(r => [r.__label, r['CPU_PCT'] === '' ? '' : Number(r['CPU_PCT']), r['SAPS Utilization'] === '' ? '' : Number(r['SAPS Utilization']), Number(r['SAPS Deployed']) || 0]),
                ];
                XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(dataAoa), 'CPU_Data');
                XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['HANA CPU — native Excel charts (linked to CPU_Data)']]), 'CPU_Charts');
                const last = FIRST_DATA_ROW + rowsP.length - 1;
                /* Clustered column for CPU_PCT (col B). */
                descriptors.push({ chartSheet: 'CPU_Charts', dataSheet: 'CPU_Data', firstDataRow: FIRST_DATA_ROW, lastDataRow: last,
                    type: 'clusteredColumn', title: 'HANA CPU Utilization per 5 minute snapshots',
                    seriesCols: [{ col: 'B', name: 'CPU_PCT', color: '#0070F2' }], leftAxisTitle: 'CPU Utilization (%)', catAxisTitle: 'Date / Time', anchorRow: 2 });
                /* Combo: SAPS Utilization bar (C) + SAPS Deployed line (D). */
                descriptors.push({ chartSheet: 'CPU_Charts', dataSheet: 'CPU_Data', firstDataRow: FIRST_DATA_ROW, lastDataRow: last,
                    type: 'comboBarLine', title: 'HANA CPU Utilization per 5 minute snapshots',
                    barCols: [{ col: 'C', name: 'SAPS Utilization', color: '#0070F2' }], lineCols: [{ col: 'D', name: 'SAPS Deployed', color: '#EF4444' }],
                    leftAxisTitle: 'SAPS', rightAxisTitle: 'SAPS', catAxisTitle: 'Date / Time', anchorRow: 36 });
            }

            this._chartProgress('Writing Memory worksheets…', 60);
            await this._nextFrame();

            /* ── Memory_Data + Memory_Charts (stacked area) ── */
            const mn = this.memNormalized;
            if (mn && mn.labels.length && mn.components.length) {
                const pick = this._downsample(mn.labels, MAX_PTS);
                const header = ['MY_DATE', ...mn.components];
                const dataAoa = [
                    ...docHeader(
                        'RSxxxSTT → HANA → HANA Initial Capacity Check → 7 Memory Component Overview',
                        'Manual export from the HANA Initial Capacity Check; normalized to one column per component',
                        'RSxxxSTT (HANA Initial Capacity Check)'),
                    header,
                    ...pick.map(i => [mn.labels[i], ...mn.components.map(c => { const v = mn.series[c][i]; return v == null ? '' : v; })]),
                ];
                XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(dataAoa), 'Memory_Data');
                XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['HANA Memory — native stacked-area chart (linked to Memory_Data)']]), 'Memory_Charts');
                const last = FIRST_DATA_ROW + pick.length - 1;
                const seriesCols = mn.components.map((c, i) => ({ col: colLetter(i + 1), name: c, color: CHART_COLORS[i % CHART_COLORS.length] }));
                descriptors.push({ chartSheet: 'Memory_Charts', dataSheet: 'Memory_Data', firstDataRow: FIRST_DATA_ROW, lastDataRow: last,
                    type: 'stackedArea', title: 'Master Node - Memory Components',
                    seriesCols, leftAxisTitle: 'MEMORY_SIZE_GB', catAxisTitle: 'MY_DATE', anchorRow: 2 });
            }

            /* ── Historic_EWA_Memory (optional, data only — kept independent) ── */
            if (!this.histNormalized && this.histUpload) this._processHist();
            const hn = this.histNormalized;
            if (hn && hn.labels.length && hn.components.length) {
                const header = ['Date', ...hn.components];
                const dataAoa = [
                    ...docHeader(
                        'Historic EWA Memory Components Overview (EarlyWatch Alert Dashboard)',
                        'Manual export from the EWA Dashboard — daily granularity, ≥12 months; normalized to one column per component',
                        'EarlyWatch Alert Dashboard export'),
                    header,
                    ...hn.labels.map((lbl, i) => [lbl, ...hn.components.map(c => { const v = hn.series[c][i]; return v == null ? '' : v; })]),
                ];
                XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(dataAoa), 'Historic_EWA_Memory');
            }

            if (!descriptors.length) {
                this._hideChartProgress(); if (btn) btn.disabled = false;
                showToast('Nothing to export', 'No CPU or Memory data is available for the workbook.', 'info');
                return;
            }

            this._chartProgress('Embedding native Excel charts…', 78);
            await this._nextFrame();
            const dbSid = (this._systemInfo(this.app).sid || this._sidTag()).toString().replace(/[^A-Za-z0-9]/g, '') || 'HANA';
            const fileName = `${dbSid}_database.xlsx`;
            await this._writeXlsxWithCharts(wb, descriptors, fileName);

            this._chartProgress('Workbook ready', 100);
            setTimeout(() => this._hideChartProgress(), 900);
            showToast('Workbook downloaded', `${fileName} generated with ${descriptors.length} native chart(s) linked to their data worksheets.`, 'success');
        } catch (err) {
            this._hideChartProgress();
            console.error('[TWDB workbook] build failed', err);
            showToast('Workbook failed', err && err.message ? err.message : String(err), 'error', 8000);
        } finally {
            if (btn) btn.disabled = false;
        }
    },

    async _writeXlsxWithCharts(wb, descriptors, fileName) {
        const downloadPlain = () => {
            const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = fileName;
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(a.href), 10000);
        };
        if (typeof JSZip === 'undefined') { downloadPlain(); return; }
        try {
            const bytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            const zip = await JSZip.loadAsync(bytes);
            const sheetIndexMap = await this._resolveSheetIndexMap(zip);
            let cId = 0;
            /* Group descriptors by chart sheet so multiple charts can share one sheet's drawing. */
            const bySheet = {};
            descriptors.forEach(d => { (bySheet[d.chartSheet] = bySheet[d.chartSheet] || []).push(d); });

            for (const sheetName of Object.keys(bySheet)) {
                const sheetIdx = sheetIndexMap[sheetName];
                if (!sheetIdx) { console.warn('[TWDB workbook] chart sheet not found:', sheetName); continue; }
                const group = bySheet[sheetName];
                const anchors = [];
                for (const desc of group) {
                    try {
                        cId++;
                        zip.file(`xl/charts/chart${cId}.xml`, this._makeChartXml(desc, cId));
                        zip.file(`xl/charts/_rels/chart${cId}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
                        anchors.push({ cId, anchorRow: desc.anchorRow || 2 });
                    } catch (err) { cId--; console.warn('[TWDB workbook] chart XML failed:', desc.title, err); }
                }
                if (!anchors.length) continue;
                /* One drawing per sheet holding each chart's graphicFrame. */
                const drawId = sheetIdx;
                zip.file(`xl/drawings/drawing${drawId}.xml`, this._makeDrawingXml(anchors));
                zip.file(`xl/drawings/_rels/drawing${drawId}.xml.rels`,
                    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
                    anchors.map((a, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${a.cId}.xml"/>`).join('') +
                    `</Relationships>`);
                const sheetPath = `xl/worksheets/sheet${sheetIdx}.xml`;
                const sheetXml = await zip.file(sheetPath)?.async('string');
                if (sheetXml) zip.file(sheetPath, this._patchSheetXml(sheetXml, 'rId1'));
                const relPath = `xl/worksheets/_rels/sheet${sheetIdx}.xml.rels`;
                const existingRel = await zip.file(relPath)?.async('string');
                zip.file(relPath, this._patchSheetRels(existingRel, `../drawings/drawing${drawId}.xml`));
            }

            if (cId === 0) { downloadPlain(); return; }
            const ctRaw = await zip.file('[Content_Types].xml').async('string');
            zip.file('[Content_Types].xml', this._patchContentTypes(ctRaw, cId, Object.keys(bySheet).map(s => sheetIndexMap[s]).filter(Boolean)));
            const outBuf = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
            const blob = new Blob([outBuf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = fileName;
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(a.href), 10000);
        } catch (err) {
            console.warn('[TWDB workbook] native-chart embedding failed — data-only fallback:', err);
            downloadPlain();
        }
    },

    _xmlEsc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); },

    _txPr(sz, bold) {
        return `<c:txPr><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>` +
            `<a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>` +
            `<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:pPr><a:defRPr sz="${sz}" b="${bold ? 1 : 0}"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
    },

    _makeChartXml(desc, cId) {
        const sz = ChartFontSizeManager.get();
        const q = n => `'${String(n).replace(/'/g, "''")}'`;
        const ds = desc.dataSheet, fdr = desc.firstDataRow, ldr = desc.lastDataRow;
        const ref = (col, r1, r2) => `${q(ds)}!$${col}$${r1}:$${col}$${r2}`;
        const catRef = ref('A', fdr, ldr);
        const esc = s => this._xmlEsc(s);

        const hex = c => String(c || '').replace('#', '').toUpperCase();
        const fillSpPr = c => c ? `<c:spPr><a:solidFill><a:srgbClr val="${hex(c)}"/></a:solidFill></c:spPr>` : '';
        const lineSpPr = c => c ? `<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="${hex(c)}"/></a:solidFill></a:ln></c:spPr>` : '';
        const serFor = (cols, isLine) => (cols || []).map((sc, i) =>
            `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>` +
            `<c:tx><c:strRef><c:f>${q(ds)}!$${sc.col}$${fdr - 1}</c:f></c:strRef></c:tx>` +
            (isLine ? lineSpPr(sc.color) : fillSpPr(sc.color)) +
            (isLine ? `<c:smooth val="0"/>` : '') +
            `<c:cat><c:strRef><c:f>${catRef}</c:f></c:strRef></c:cat>` +
            `<c:val><c:numRef><c:f>${ref(sc.col, fdr, ldr)}</c:f></c:numRef></c:val></c:ser>`).join('');

        const axisTitle = txt => txt
            ? `<c:title><c:tx><c:rich><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>` +
              `<a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>` +
              `<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:pPr><a:defRPr sz="${sz.axisTitle * 100}"/></a:pPr>` +
              `<a:r><a:rPr lang="en-US" sz="${sz.axisTitle * 100}"/><a:t>${esc(txt)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : '';

        const catAx = (id, cross) => `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
            `<c:delete val="0"/><c:axPos val="b"/>${axisTitle(desc.catAxisTitle)}<c:crossAx val="${cross}"/><c:tickLblPos val="low"/>${this._txPr(sz.axisTicks * 100, false)}</c:catAx>`;
        const valAx = (id, cross, pos, title) => `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
            `<c:delete val="0"/><c:axPos val="${pos}"/>${axisTitle(title)}<c:crossAx val="${cross}"/>${this._txPr(sz.axisTicks * 100, false)}</c:valAx>`;

        let plotArea;
        if (desc.type === 'clusteredColumn') {
            plotArea = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
                serFor(desc.seriesCols, false) +
                `<c:overlap val="-27"/><c:axId val="1001"/><c:axId val="1002"/></c:barChart>` +
                catAx(1001, 1002) + valAx(1002, 1001, 'l', desc.leftAxisTitle);
        } else if (desc.type === 'comboBarLine') {
            plotArea = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
                serFor(desc.barCols, false) +
                `<c:overlap val="-27"/><c:axId val="1001"/><c:axId val="1002"/></c:barChart>` +
                `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>` +
                serFor(desc.lineCols, true) +
                `<c:marker val="1"/><c:axId val="1004"/><c:axId val="1003"/></c:lineChart>` +
                catAx(1001, 1002) +
                valAx(1002, 1001, 'l', desc.leftAxisTitle) +
                valAx(1003, 1004, 'r', desc.rightAxisTitle) +
                `<c:catAx><c:axId val="1004"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="1003"/></c:catAx>`;
        } else if (desc.type === 'stackedArea') {
            plotArea = `<c:areaChart><c:grouping val="stacked"/><c:varyColors val="0"/>` +
                serFor(desc.seriesCols, false) +
                `<c:axId val="1001"/><c:axId val="1002"/></c:areaChart>` +
                catAx(1001, 1002) + valAx(1002, 1001, 'l', desc.leftAxisTitle);
        } else { /* line fallback */
            plotArea = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>` +
                serFor(desc.seriesCols, true) +
                `<c:marker val="0"/><c:axId val="1001"/><c:axId val="1002"/></c:lineChart>` +
                catAx(1001, 1002) + valAx(1002, 1001, 'l', desc.leftAxisTitle);
        }

        const titleXml = `<c:title><c:tx><c:rich><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>` +
            `<a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>` +
            `<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:pPr algn="ctr"><a:defRPr sz="${sz.title * 100}" b="1"/></a:pPr>` +
            `<a:r><a:rPr lang="en-US" sz="${sz.title * 100}" b="1"/><a:t>${esc(desc.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <c:autoTitleDeleted val="0"/>
    ${titleXml}
    <c:plotArea><c:layout/>${plotArea}</c:plotArea>
    <c:legend><c:legendPos val="b"/><c:overlay val="0"/>${this._txPr(sz.legend * 100, false)}</c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
</c:chartSpace>`;
    },

    /* One drawing holding each chart's anchored graphicFrame (stacked vertically). */
    _makeDrawingXml(anchors) {
        const frames = anchors.map((a, i) => {
            const top = a.anchorRow;
            const bottom = top + 30;
            return `<xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${top}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>14</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${bottom}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="${a.cId + 1}" name="Chart ${a.cId}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId${i + 1}"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>`;
        }).join('');
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">${frames}</xdr:wsDr>`;
    },

    _patchContentTypes(xml, chartCount, drawingIdxs) {
        let patched = xml;
        for (let i = 1; i <= chartCount; i++) {
            const chartPart = `/xl/charts/chart${i}.xml`;
            if (!patched.includes(chartPart)) patched = patched.replace('</Types>', `  <Override PartName="${chartPart}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>\n</Types>`);
        }
        (drawingIdxs || []).forEach(idx => {
            const drawingPart = `/xl/drawings/drawing${idx}.xml`;
            if (!patched.includes(drawingPart)) patched = patched.replace('</Types>', `  <Override PartName="${drawingPart}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>\n</Types>`);
        });
        return patched;
    },

    async _resolveSheetIndexMap(zip) {
        const wbXml = await zip.file('xl/workbook.xml').async('string');
        const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
        const nameToRid = {}; let m; const re1 = /<sheet\b[^>]*>/g;
        while ((m = re1.exec(wbXml)) !== null) {
            const tag = m[0];
            const nameM = /\bname="([^"]*)"/.exec(tag), ridM = /\br:id="([^"]*)"/.exec(tag);
            if (nameM && ridM) nameToRid[nameM[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")] = ridM[1];
        }
        const ridToTarget = {}; const re2 = /<Relationship\b[^>]*>/g;
        while ((m = re2.exec(relsXml)) !== null) {
            const tag = m[0];
            const idM = /\bId="([^"]*)"/.exec(tag), tgtM = /\bTarget="([^"]*)"/.exec(tag);
            if (idM && tgtM) ridToTarget[idM[1]] = tgtM[1];
        }
        const result = {};
        Object.entries(nameToRid).forEach(([name, rid]) => {
            const target = ridToTarget[rid] || '';
            const idxM = /worksheets\/sheet(\d+)\.xml/i.exec(target);
            if (idxM) result[name] = parseInt(idxM[1]);
        });
        return result;
    },

    _patchSheetXml(xml, drawingRid) {
        let patched = xml;
        if (!patched.includes('xmlns:r=')) patched = patched.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
        patched = patched.replace(/<drawing\s[^>]*\/>/g, '').replace(/<drawing\/>/g, '');
        const tag = `<drawing r:id="${drawingRid}"/>`;
        if (patched.includes('</worksheet>')) patched = patched.replace('</worksheet>', `${tag}</worksheet>`);
        else patched += tag;
        return patched;
    },

    _patchSheetRels(existingXml, drawingTarget) {
        const newRel = `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="${drawingTarget}"/>`;
        if (!existingXml) return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${newRel}</Relationships>`;
        let patched = existingXml.replace(/<Relationship[^>]*Type="[^"]*\/drawing"[^>]*\/>/g, '');
        if (patched.includes('</Relationships>')) patched = patched.replace('</Relationships>', `${newRel}</Relationships>`);
        else patched += newRel;
        return patched;
    },
};
