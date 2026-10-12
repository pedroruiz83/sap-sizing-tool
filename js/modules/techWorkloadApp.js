/* ─────────────────────────────────────────────────────────
   Module 02 – Technical Workload: Application Layer
   ─────────────────────────────────────────────────────────
   Input  : SAP transaction /SDF/SMON_HEADER or /SDF/MON_HEADER
            performance statistics, exported as
              • plain text / enriched text (RTF) — "|" column separator
              • Excel (XLSX / XLS)
              • HTML
   Output : an enriched target workbook (XLSX) that cross-references the
            System Landscape (Module 01) to derive per-host CPU and SAPS
            utilisation plus memory metrics in GB.
   ─────────────────────────────────────────────────────────*/

const TechWorkloadAppModule = {

    /* ── persistent state (survives navigation; cleared by New Analysis) ── */
    sourceTable: '',      // 'SMON' | 'MON'
    targetRows:  [],      // array-of-objects for the output sheet
    targetCols:  [],      // ordered column names for the output sheet
    fileName:    '',

    /* ── optional NMON step (IBM Power Systems only) ── */
    nmonFiles:    [],       // uploaded NMON File objects
    nmonAnalysis: null,     // { workbook summary } after processing, or null

    /* ── optional Step 3 (ST02 Memory Buffer) & Step 4 (ST02 WP Storage) ──
       Uploaded files are stored as { name, text } (serializable raw text) so
       they survive Save/Load Session; File objects would be dropped by the
       serializer. The analyses hold the computed tables; their transient
       SheetJS `workbook` is rebuilt lazily on download after a restore. */
    st02BufFiles:    [],    // [{ name, text }]
    st02BufAnalysis: null,  // { fileName, bufferDetail[], perBuffer[], perServer[], total{}, workbook }
    st02WpFiles:     [],    // [{ name, text }]
    st02WpAnalysis:  null,  // { fileName, servers[], totalUser{}, totalExec{}, workbook }

    /* ── optional SAP System Profile files ── */
    profiles:        [],    // kept profiles: { name, kind, host, instanceNo, sapsysname, params:{} }
    profilesIgnored: [],    // { name, reason } for files whose SAPSYSTEMNAME ≠ in-scope SID

    /* ── virtual↔physical hostname mapping state ── */
    _mappingActive:   false,   // true when the mapping step was used for this build
    _mapping:         {},      // { virtualHost(lowercase) -> physicalHostname }
    _virtualHosts:    [],      // distinct SMON/MON hostnames (original case)
    _removedHosts:    null,    // Set of lowercased physical hostnames the user removed
    _pendingDataRows: null,    // data rows held while awaiting user mapping
    _pendingMeta:     null,    // { source, headers, col } held while awaiting mapping

    resetState() {
        this.sourceTable = '';
        this.targetRows  = [];
        this.targetCols  = [];
        this.fileName    = '';
        this.nmonFiles    = [];
        this.nmonAnalysis = null;
        this.st02BufFiles    = [];
        this.st02BufAnalysis = null;
        this.st02WpFiles     = [];
        this.st02WpAnalysis  = null;
        this.profiles        = [];
        this.profilesIgnored = [];
        this._mappingActive   = false;
        this._mapping         = {};
        this._virtualHosts    = [];
        this._manualVirtualHosts = [];  /* hostnames typed manually (not from SMON/MON) */
        this._removedHosts    = null;
        this._pendingDataRows = null;
        this._pendingMeta     = null;
        this._buildingWithMapping = false;
        this._extMemory       = {};   /* { virtualHost(lowercase) → GB (number|null) } */
        this._colIndex        = null;
        this._summary         = null;
        this._matchStats      = null;
    },

    /* ── Session save/restore contract ──
       Serializable working state only. Raw uploaded NMON File objects and any
       transient pending-mapping buffers are omitted; the processed results
       (targetRows/targetCols, profiles, mapping) are what downstream modules
       and the UI need. A Set is stored as an array and rehydrated on restore. */
    getSessionState() {
        return {
            sourceTable: this.sourceTable,
            targetRows:  this.targetRows,
            targetCols:  this.targetCols,
            fileName:    this.fileName,
            nmonAnalysis: this.nmonAnalysis,
            /* ST02 Step 3/4 — raw file text is serializable; the transient SheetJS
               workbook is stripped and rebuilt lazily on download after restore. */
            st02BufFiles:    this.st02BufFiles || [],
            st02BufAnalysis: this.st02BufAnalysis
                ? (({ workbook, ...rest }) => rest)(this.st02BufAnalysis) : null,
            st02WpFiles:     this.st02WpFiles || [],
            st02WpAnalysis:  this.st02WpAnalysis
                ? (({ workbook, ...rest }) => rest)(this.st02WpAnalysis) : null,
            profiles:        this.profiles,
            profilesIgnored: this.profilesIgnored,
            mappingActive:   this._mappingActive,
            mapping:         this._mapping,
            virtualHosts:    this._virtualHosts,
            manualVirtualHosts: this._manualVirtualHosts || [],
            removedHosts:    this._removedHosts ? Array.from(this._removedHosts) : null,
            extMemory:       this._extMemory,
            colIndex:        this._colIndex || null,
            summary:         this._summary  || null,
            matchStats:      this._matchStats || null,
        };
    },
    setSessionState(st) {
        if (!st) return;
        this.sourceTable = st.sourceTable || '';
        this.targetRows  = st.targetRows  || [];
        this.targetCols  = st.targetCols  || [];
        this.fileName    = st.fileName    || '';
        this.nmonAnalysis = st.nmonAnalysis || null;
        /* ST02 Step 3/4 — restore raw text + computed tables; workbook rebuilt on download. */
        this.st02BufFiles    = st.st02BufFiles    || [];
        this.st02BufAnalysis = st.st02BufAnalysis || null;
        this.st02WpFiles     = st.st02WpFiles     || [];
        this.st02WpAnalysis  = st.st02WpAnalysis  || null;
        this.profiles        = st.profiles        || [];
        this.profilesIgnored = st.profilesIgnored || [];
        this._mappingActive  = !!st.mappingActive;
        this._mapping        = st.mapping || {};
        this._virtualHosts   = st.virtualHosts || [];
        this._manualVirtualHosts = st.manualVirtualHosts || [];
        this._removedHosts   = Array.isArray(st.removedHosts) ? new Set(st.removedHosts) : null;
        this._extMemory      = st.extMemory || {};
        this._colIndex       = st.colIndex  || null;
        this._summary        = st.summary   || null;
        this._matchStats     = st.matchStats || null;
        /* Store truncation flag so _renderResults can show an appropriate banner. */
        this._targetRowsTruncated      = !!st._targetRowsTruncated;
        this._targetRowsDropped        = !!st._targetRowsDropped;
        this._targetRowsOriginalCount  = st._targetRowsOriginalCount || 0;
        /* Raw NMON files are not restored — the processed nmonAnalysis is. */
        this.nmonFiles = [];
    },

    render() {
        return `
        <div class="module-wrap">
            <div class="module-header">
                <div class="module-crumb">Analysis › Step 02</div>
                <h1 class="module-title">Technical Workload – Application Layer</h1>
                <p class="module-desc">
                    Upload the performance statistics from the SAP tables
                    <strong>/SDF/SMON_HEADER</strong> or <strong>/SDF/MON_HEADER</strong>.
                    The tool detects the source table automatically, cross-references the
                    System Landscape, and builds an enriched analysis workbook with per-host
                    CPU / SAPS utilisation and memory metrics.
                </p>
            </div>

            <div id="twa-upload-screen">

            <!-- SAP System Profiles — full card, shown first; collapses to status bar once loaded -->
            <div class="card" style="margin-bottom:12px" id="twa-pfl-inline-card">
                <!-- content injected by _renderPflInlineCard() -->
            </div>

            <div class="alert-info" style="margin-bottom:12px">
                <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                     fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                    <circle cx="12" cy="12" r="10"/>
                    <line x1="12" y1="8" x2="12" y2="12"/>
                    <line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
                <span><strong>Step 2:</strong> export the full result set from <strong>/SDF/SMON_HEADER</strong> or
                <strong>/SDF/MON_HEADER</strong>. Plain-text exports use the <code>|</code>
                column separator — paste or save them as .txt. Excel, HTML and RTF are also accepted.
                Even a reduced set of columns works; the tool applies only the steps that are applicable.</span>
            </div>

            <div class="card" style="margin-top:10px">
                <div class="card-header">
                    <div>
                        <div class="card-title">Performance Statistics Upload</div>
                        <div class="card-sub">/SDF/SMON_HEADER, /SDF/MON_HEADER, or a pre-processed spreadsheet — TXT · RTF · XLSX · XLS · HTML · ZIP · multiple files</div>
                    </div>
                </div>
                <div class="upload-zone" id="twa-zone">
                    <input type="file" class="file-input-hidden" id="twa-input" multiple
                           accept=".txt,.text,.rtf,.csv,.xlsx,.xls,.html,.htm,.zip">
                    <div class="upload-zone-icon">
                        <svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="1.6">
                            <polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>
                        </svg>
                    </div>
                    <div class="upload-zone-text">
                        <div class="upload-zone-title">Upload the "/SDF/SMON or /SDF/MON" statistics (or a pre-processed spreadsheet)</div>
                        <div class="upload-zone-sub">Drag &amp; drop or browse — one or more files, or a .zip bundle (TXT · RTF · XLSX · XLS · HTML)</div>
                    </div>
                    <span class="upload-zone-cta">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="17 8 12 3 7 8"/>
                            <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        Choose file(s)
                    </span>
                </div>
                <div id="twa-status"></div>
            </div>

            </div><!-- /#twa-upload-screen -->

            <!-- Error panel -->
            <div id="twa-error" style="display:none" class="alert-error">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24"
                     fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                    <circle cx="12" cy="12" r="10"/>
                    <line x1="15" y1="9" x2="9" y2="15"/>
                    <line x1="9" y1="9" x2="15" y2="15"/>
                </svg>
                <div id="twa-error-msg"></div>
            </div>

            <!-- ══ SAP System Profiles step (optional, its own screen) ══ -->
            <div id="twa-pfl-view" style="display:none">
                <div class="module-header" style="margin-top:4px">
                    <div class="module-crumb">Analysis › Step 02 › SAP System Profiles (optional)</div>
                    <h1 class="module-title">Upload SAP System Profiles</h1>
                    <p class="module-desc">
                        Optionally upload the SAP system profile files (<code>DEFAULT.PFL</code>,
                        instance profiles <code>&lt;SID&gt;_D##_&lt;host&gt;</code>, ASCS profiles
                        <code>&lt;SID&gt;_ASCS##_&lt;host&gt;</code>). They are classified by type,
                        validated against the system in scope (<code>SAPSYSTEMNAME</code>), and retained
                        for later steps — the <strong>Extended Memory per Server</strong> table auto-fills
                        from <code>em/initial_size_MB</code> when available. Profiles belonging to a
                        different SID are ignored. This step is <strong>optional</strong>.
                    </p>
                </div>

                <div class="card" id="twa-pfl-card">
                    <div class="card-header">
                        <div>
                            <div class="card-title">SAP System Profiles <span style="font-weight:400;color:var(--text-muted)">— optional</span></div>
                            <div class="card-sub">
                                DEFAULT.PFL, instance profiles (<code>&lt;SID&gt;_D##_&lt;host&gt;</code>) and
                                ASCS profiles (<code>&lt;SID&gt;_ASCS##_&lt;host&gt;</code>) — <strong>.PFL · .pfl · no extension</strong> · one or multiple files
                            </div>
                        </div>
                    </div>
                    <div class="upload-zone" id="twa-pfl-zone" style="min-height:120px">
                        <input type="file" class="file-input-hidden" id="twa-pfl-input" multiple
                               accept="*/*">
                        <div class="upload-zone-icon">
                            <svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24"
                                 fill="none" stroke="currentColor" stroke-width="1.6">
                                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                                <polyline points="14 2 14 8 20 8"/>
                            </svg>
                        </div>
                        <div class="upload-zone-text">
                            <div class="upload-zone-title">Upload SAP system profile files</div>
                            <div class="upload-zone-sub">Drag &amp; drop or browse — DEFAULT.PFL, &lt;SID&gt;_D##_&lt;host&gt;, &lt;SID&gt;_ASCS##_&lt;host&gt; · .pfl / no extension · multiple files</div>
                        </div>
                        <span class="upload-zone-cta">
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                                 fill="none" stroke="currentColor" stroke-width="2.2">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="17 8 12 3 7 8"/>
                                <line x1="12" y1="3" x2="12" y2="15"/>
                            </svg>
                            Choose file(s)
                        </span>
                    </div>
                    <div id="twa-pfl-status"></div>
                    <div id="twa-pfl-list" style="display:none;padding:4px 2px 2px"></div>
                </div>

                <div class="action-row" style="margin-top:18px;gap:10px">
                    <button class="btn" id="twa-pfl-back">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="15 18 9 12 15 6"/>
                        </svg>
                        Back
                    </button>
                </div>
            </div>

            <!-- ══ Virtual↔Physical hostname mapping (shown only when needed) ══ -->
            <div id="twa-mapping" style="display:none"></div>

            <!-- ══ SAP Extended Memory configuration step (shown before the build) ══ -->
            <div id="twa-emconfig" style="display:none"></div>

            <!-- Processing progress — sits just above the Enriched Workload Table -->
            <div id="twa-progress" class="proc-progress" style="display:none;margin-top:18px">
                <div class="proc-progress-head">
                    <span id="twa-progress-label">Processing…</span>
                    <span id="twa-progress-pct">0%</span>
                </div>
                <div class="proc-progress-track">
                    <div class="proc-progress-fill" id="twa-progress-fill"></div>
                </div>
            </div>

            <div id="twa-results" style="display:none">

                <div class="section-title" style="margin-top:24px">Analysis Summary</div>
                <div class="stats-grid" id="twa-stats"></div>
                <div class="alert-error" id="twa-truncation-note" style="display:none;margin-top:10px"></div>

                <div class="section-title">
                    Enriched Workload Table
                    <span style="font-size:11px;font-weight:400;color:var(--text-muted);margin-left:6px"
                          id="twa-src-note"></span>
                </div>
                <div class="card" style="padding:0">
                    <div class="card-header" style="padding:14px 18px 12px; border-radius:14px 14px 0 0">
                        <div>
                            <div class="card-title">Target Analysis</div>
                            <div class="card-sub" id="twa-table-sub"></div>
                        </div>
                        <button class="btn btn-blue" id="twa-btn-download">
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                                 fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="7 10 12 15 17 10"/>
                                <line x1="12" y1="15" x2="12" y2="3"/>
                            </svg>
                            Download Analysis (XLSX)
                        </button>
                    </div>
                    <div class="table-wrap" style="border:none;border-radius:0 0 14px 14px;overflow-x:auto;max-height:520px">
                        <table class="data-table" id="twa-table">
                            <thead id="twa-thead"></thead>
                            <tbody id="twa-tbody"></tbody>
                        </table>
                    </div>
                </div>
                <div style="font-size:11px;color:var(--text-muted);margin-top:8px" id="twa-preview-note"></div>

                <!-- Physical/Virtual Server Inventory (only when mapping was used) -->
                <div id="twa-pv-inventory" style="display:none">
                    <div class="section-title" style="margin-top:28px">
                        Server Inventory — Physical ↔ Virtual
                        <span style="font-size:11px;font-weight:400;color:var(--text-muted);margin-left:6px">
                            — landscape hosts with their mapped SMON/MON virtual hostname
                        </span>
                    </div>
                    <div class="card" style="padding:0">
                        <div class="card-header" style="padding:14px 18px 12px; border-radius:14px 14px 0 0">
                            <div>
                                <div class="card-title">Server Inventory</div>
                                <div class="card-sub" id="twa-pv-sub"></div>
                            </div>
                            <div style="display:flex;gap:8px;flex-wrap:wrap">
                                <button class="btn" id="twa-pv-copy">
                                    <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13"
                                         viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                        <rect x="9" y="9" width="13" height="13" rx="2" ry="2"/>
                                        <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>
                                    </svg>
                                    Copy for Excel
                                </button>
                                <button class="btn" id="twa-pv-export">
                                    <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13"
                                         viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                        <polyline points="7 10 12 15 17 10"/>
                                        <line x1="12" y1="15" x2="12" y2="3"/>
                                    </svg>
                                    Export
                                </button>
                            </div>
                        </div>
                        <div class="table-wrap" style="border:none;border-radius:0 0 14px 14px;overflow-x:auto">
                            <table class="data-table" id="twa-pv-table">
                                <thead>
                                    <tr>
                                        <th>Physical Hostname</th>
                                        <th>Virtual Hostname</th>
                                        <th>Hardware Manufacturer</th>
                                        <th>Model</th>
                                        <th>CPU Type</th>
                                        <th>Virtualization</th>
                                        <th>Operating System</th>
                                        <th>GHz</th>
                                        <th>Logical CPUs</th>
                                        <th>Cores</th>
                                        <th>SAPS (TBV)</th>
                                        <th>Memory in MB</th>
                                        <th>Component Type</th>
                                        <th>Layer</th>
                                        <th>SAPS/Core</th>
                                    </tr>
                                </thead>
                                <tbody id="twa-pv-tbody"></tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <div class="action-row">
                    <button class="btn btn-blue" id="twa-btn-charts">
                        Proceed to Charts — Technical Application Workload Analysis
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="9 18 15 12 9 6"/>
                        </svg>
                    </button>
                </div>
            </div>

            <!-- ══ Charts view (Application Workload Analysis) ══ -->
            <div id="twa-charts" style="display:none">
                <div class="module-header" style="margin-top:4px">
                    <div class="module-crumb">Analysis › Step 02 › Charts</div>
                    <h1 class="module-title">Technical Workload – Application Analysis</h1>
                    <p class="module-desc">
                        Interactive charts for the <strong>application-tier</strong> servers only.
                        Database servers are analysed in the next module (Technical Workload – Database).
                        Click a section header to collapse or expand it. Within each chart: drag/scroll to
                        zoom, Shift-drag to pan, edit the title, adjust axis ranges, copy or export (SVG).
                    </p>
                </div>

                <!-- Chart-build progress (shown while charts are being generated) -->
                <div id="twa-chart-progress" class="proc-progress" style="display:none;margin-bottom:18px">
                    <div class="proc-progress-head">
                        <span id="twa-chart-progress-label">Generating charts…</span>
                        <span id="twa-chart-progress-pct">0%</span>
                    </div>
                    <div class="proc-progress-track">
                        <div class="proc-progress-fill" id="twa-chart-progress-fill"></div>
                    </div>
                </div>

                <!-- Analysis tabs: only the active section's charts are rendered/live,
                     which keeps the page responsive (no 40+ live canvases at once). -->
                <div class="chart-tabs" id="twa-chart-tabs" role="tablist">
                    <button class="chart-tab is-active" data-tab="cpu" role="tab">CPU Utilization</button>
                    <button class="chart-tab" data-tab="mem" role="tab">Memory Utilization</button>
                    <button class="chart-tab" data-tab="wp" role="tab">Work Process</button>
                    <button class="chart-tab" data-tab="logins" role="tab">Logins &amp; Sessions</button>
                    <button class="chart-tab" data-tab="network" role="tab">Network</button>
                </div>

                <!-- Analysis workbook download — kept at the top so it is visible
                     without scrolling. Rebuilt on demand so it always reflects the
                     latest chart data. -->
                <div class="action-row" style="margin:4px 0 18px;justify-content:flex-start">
                    <button class="btn btn-blue" id="twa-charts-wb-top" title="Download the Technical Workload Analysis workbook with native, editable Excel charts linked to their data">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="7 10 12 15 17 10"/>
                            <line x1="12" y1="15" x2="12" y2="3"/>
                        </svg>
                        Download Technical Workload Analysis (XLSX)
                    </button>
                    <span style="font-size:11px;color:var(--text-muted)">
                        Includes source data worksheets plus one Data + Chart worksheet pair per chart (native, editable Excel charts).
                    </span>
                </div>

                <!-- CPU Utilization Analysis -->
                <div class="chart-section" data-sec="cpu">
                    <div class="chart-section-head" data-toggle="cpu">
                        <svg class="chart-section-chevron" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                        <span class="chart-section-name">CPU Utilization Analysis</span>
                    </div>
                    <div class="chart-section-body" id="twa-sec-cpu">

                        <!-- Group 1 — Overview: average CPU for all AS -->
                        <div class="chart-group" id="twa-cpugrp-overview">
                            <div class="chart-group-head">
                                <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                                <span>Average CPU Usage — All Application Servers</span>
                                <span class="chart-group-count" id="twa-cpugrp-overview-count">1 chart</span>
                            </div>
                            <div class="chart-group-body" id="twa-cpugrp-overview-body"></div>
                        </div>

                        <!-- Group 2 — Per-host CPU Utilization (%) over Date/Time -->
                        <div class="chart-group" id="twa-cpugrp-perhost">
                            <div class="chart-group-head">
                                <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                                <span>CPU Utilization (%) — Per Server</span>
                                <span class="chart-group-count" id="twa-cpugrp-perhost-count"></span>
                            </div>
                            <div class="chart-group-body" id="twa-cpugrp-perhost-body"></div>
                        </div>

                        <!-- Group 3 — 3D hourly CPU surface per server -->
                        <div class="chart-group" id="twa-cpugrp-3d">
                            <div class="chart-group-head">
                                <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                                <span>3D Hourly CPU Surface — Per Server (Date × Hour × avg CPU%)</span>
                                <span class="chart-group-count" id="twa-cpugrp-3d-count"></span>
                            </div>
                            <div class="chart-group-body" id="twa-cpugrp-3d-body"></div>
                        </div>

                        <!-- Group 4 — SAPS utilization hourly average per server -->
                        <div class="chart-group" id="twa-cpugrp-saps">
                            <div class="chart-group-head">
                                <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                                <span>SAPS Utilization — Hourly Average per Server</span>
                                <span class="chart-group-count" id="twa-cpugrp-saps-count"></span>
                            </div>
                            <div class="chart-group-body" id="twa-cpugrp-saps-body"></div>
                        </div>

                        <!-- Group 5 — Stacked SAPS: all AS combined vs total deployed capacity -->
                        <div class="chart-group" id="twa-cpugrp-stacked">
                            <div class="chart-group-head">
                                <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                                <span>Peak CPU consumption for all AS — Stacked SAPS vs Total Deployed Capacity</span>
                                <span class="chart-group-count" id="twa-cpugrp-stacked-count">1 chart</span>
                            </div>
                            <div class="chart-group-body" id="twa-cpugrp-stacked-body"></div>
                        </div>

                        <!-- Group 6 — Top 20 Relevant CPU Peaks (representative, de-duplicated) -->
                        <div class="chart-group" id="twa-cpugrp-relpeaks">
                            <div class="chart-group-head">
                                <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                                <span>Top 20 Relevant CPU Peaks</span>
                                <span class="chart-group-count" id="twa-cpugrp-relpeaks-count"></span>
                            </div>
                            <div class="chart-group-body" id="twa-cpugrp-relpeaks-body"></div>
                        </div>

                    </div>
                </div>

                <!-- Memory Utilization Analysis (collapsible) -->
                <div class="chart-section" data-sec="mem">
                    <div class="chart-section-head" data-toggle="mem">
                        <svg class="chart-section-chevron" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                        <span class="chart-section-name">Memory Utilization Analysis</span>
                    </div>
                    <div class="chart-section-body" id="twa-sec-mem">
                        <!-- Application-server memory map (Physical Memory + EM Configured) -->
                        <div id="twa-memmap"></div>

                        <!-- Group 1 — per-host memory usage line charts -->
                        <div class="chart-group" id="twa-memgrp-perhost">
                            <div class="chart-group-head">
                                <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                                <span>Memory Usage — Per Server</span>
                                <span class="chart-group-count" id="twa-memgrp-perhost-count"></span>
                            </div>
                            <div class="chart-group-body" id="twa-memgrp-perhost-body"></div>
                        </div>

                        <!-- Group 2 — stacked EM Allocated for all AS -->
                        <div class="chart-group" id="twa-memgrp-em">
                            <div class="chart-group-head">
                                <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                                <span>EM Utilization for all AS (stacked)</span>
                                <span class="chart-group-count" id="twa-memgrp-em-count"></span>
                            </div>
                            <div class="chart-group-body" id="twa-memgrp-em-body"></div>
                        </div>

                        <!-- Group 2b — top-10 EM utilization peak snapshots -->
                        <div class="chart-group" id="twa-memgrp-empeaks">
                            <div class="chart-group-head">
                                <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                                <span>Top 20 Relevant EM Peaks</span>
                                <span class="chart-group-count" id="twa-memgrp-empeaks-count"></span>
                            </div>
                            <div class="chart-group-body" id="twa-memgrp-empeaks-body"></div>
                        </div>

                        <!-- Group 3 — stacked Heap memory for all AS -->
                        <div class="chart-group" id="twa-memgrp-heap">
                            <div class="chart-group-head">
                                <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                                <span>Stacked Heap Utilization for all AS</span>
                                <span class="chart-group-count" id="twa-memgrp-heap-count"></span>
                            </div>
                            <div class="chart-group-body" id="twa-memgrp-heap-body"></div>
                        </div>

                        <!-- Group 3b — top-10 Heap utilization peak snapshots -->
                        <div class="chart-group" id="twa-memgrp-heappeaks">
                            <div class="chart-group-head">
                                <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                                <span>Top 20 Relevant Heap Memory Peaks</span>
                                <span class="chart-group-count" id="twa-memgrp-heappeaks-count"></span>
                            </div>
                            <div class="chart-group-body" id="twa-memgrp-heappeaks-body"></div>
                        </div>

                        <!-- Group 4 — page-out activity per server -->
                        <div class="chart-group" id="twa-memgrp-pageout">
                            <div class="chart-group-head">
                                <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                                <span>Page Out Activity — Per Server</span>
                                <span class="chart-group-count" id="twa-memgrp-pageout-count"></span>
                            </div>
                            <div class="chart-group-body" id="twa-memgrp-pageout-body"></div>
                        </div>
                    </div>
                </div>

                <!-- Work Process Analysis (collapsible) -->
                <div class="chart-section" data-sec="wp">
                    <div class="chart-section-head" data-toggle="wp">
                        <svg class="chart-section-chevron" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                        <span class="chart-section-name">Work Process Analysis</span>
                    </div>
                    <div class="chart-section-body" id="twa-sec-wp"></div>
                </div>

                <!-- Logins and Sessions Analysis (collapsible) -->
                <div class="chart-section" data-sec="logins">
                    <div class="chart-section-head" data-toggle="logins">
                        <svg class="chart-section-chevron" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                        <span class="chart-section-name">Logins and Sessions Analysis</span>
                    </div>
                    <div class="chart-section-body" id="twa-sec-logins"></div>
                </div>

                <!-- Network Analysis (SMON only) -->
                <div class="chart-section" data-sec="network">
                    <div class="chart-section-head" data-toggle="network">
                        <svg class="chart-section-chevron" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                        <span class="chart-section-name">Network Analysis</span>
                    </div>
                    <div class="chart-section-body" id="twa-sec-network"></div>
                </div>

                <!-- Navigation buttons at the END of the charts screen -->
                <div class="action-row" style="margin-top:22px;display:flex;gap:10px">
                    <button class="btn" id="twa-charts-back2">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="15 18 9 12 15 6"/>
                        </svg>
                        Back to Enriched Table
                    </button>
                    <button class="btn btn-blue" id="twa-charts-next2">
                        Next: Technical Workload – Database
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="9 18 15 12 9 6"/>
                        </svg>
                    </button>
                </div>
            </div>

            <!-- ══ NMON view (optional, IBM Power Systems only) ══ -->
            <div id="twa-nmon" style="display:none">
                <div class="module-header" style="margin-top:4px">
                    <div class="module-crumb">Analysis › Step 02 › NMON (optional)</div>
                    <h1 class="module-title">IBM Power Systems — NMON Performance Data</h1>
                    <p class="module-desc">
                        IBM servers were detected in the landscape. Optionally upload NMON
                        performance monitoring files for additional IBM Power Systems analysis.
                        This step is <strong>optional</strong> — the mandatory
                        <strong>/SDF/SMON</strong> / <strong>/SDF/MON</strong> analysis above is unaffected.
                        Multiple files and formats (<strong>.nmon · .txt · .xlsx · .zip</strong>) are supported.
                    </p>
                </div>

                <div class="alert-info" style="margin-bottom:16px">
                    <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                        <circle cx="12" cy="12" r="10"/>
                        <line x1="12" y1="8" x2="12" y2="12"/>
                        <line x1="12" y1="16" x2="12.01" y2="16"/>
                    </svg>
                    <div style="font-size:12px">
                        NMON data from IBM Power Systems enables additional analysis of physical CPU
                        consumption, entitlement vs. virtual CPUs, memory, paging, disk and network.
                        Uploaded files are parsed automatically and consolidated into a downloadable
                        workbook. You may skip this step and proceed to the Database module at any time.
                    </div>
                </div>

                <div class="card" id="twa-nmon-card">
                    <div class="card-header" style="padding:14px 18px 10px">
                        <div>
                            <div class="card-title">NMON Files</div>
                            <div class="card-sub">
                                IBM Power Systems performance data — <strong>.nmon · .txt · .xlsx · .zip</strong>
                                &nbsp;·&nbsp; multiple files supported
                            </div>
                        </div>
                    </div>

                    <div class="upload-zone" id="twa-nmon-zone" style="min-height:120px">
                        <input type="file" class="file-input-hidden" id="twa-nmon-input"
                               accept=".nmon,.txt,.xlsx,.xls,.zip" multiple>
                        <div class="upload-zone-icon">
                            <svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24"
                                 fill="none" stroke="currentColor" stroke-width="1.6">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="17 8 12 3 7 8"/>
                                <line x1="12" y1="3" x2="12" y2="15"/>
                            </svg>
                        </div>
                        <div class="upload-zone-text">
                            <div class="upload-zone-title">Upload NMON files</div>
                            <div class="upload-zone-sub">
                                Drag &amp; drop or browse — .nmon · .txt · .xlsx · .zip · multiple files accepted
                            </div>
                        </div>
                        <span class="upload-zone-cta">
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                                 fill="none" stroke="currentColor" stroke-width="2.2">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="17 8 12 3 7 8"/>
                                <line x1="12" y1="3" x2="12" y2="15"/>
                            </svg>
                            Choose files
                        </span>
                    </div>

                    <!-- Loaded files list -->
                    <div id="twa-nmon-file-list" style="display:none;padding:10px 18px 14px">
                        <div style="font-size:12px;font-weight:600;color:var(--text-muted);
                                    text-transform:uppercase;letter-spacing:.04em;margin-bottom:6px">
                            Files queued for analysis
                        </div>
                        <ul id="twa-nmon-files-ul" style="list-style:none;padding:0;margin:0"></ul>
                    </div>
                </div>

                <!-- NMON processing progress -->
                <div id="twa-nmon-progress" class="proc-progress" style="display:none;margin-top:18px">
                    <div class="proc-progress-head">
                        <span id="twa-nmon-progress-label">Processing NMON files…</span>
                        <span id="twa-nmon-progress-pct">0%</span>
                    </div>
                    <div class="proc-progress-track">
                        <div class="proc-progress-fill" id="twa-nmon-progress-fill"></div>
                    </div>
                </div>

                <!-- NMON analysis summary (after processing) -->
                <div id="twa-nmon-summary" style="display:none"></div>

                <div class="action-row" style="margin-top:20px;display:flex;gap:10px">
                    <button class="btn" id="twa-nmon-back">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="15 18 9 12 15 6"/>
                        </svg>
                        Back to Enriched Table
                    </button>
                    <button class="btn btn-blue" id="twa-nmon-next" style="margin-left:auto">
                        Next: Technical Workload – Database
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="9 18 15 12 9 6"/>
                        </svg>
                    </button>
                </div>
            </div>

            <!-- ══ Step 3 — Memory Buffer Analysis (optional, ST02) ══ -->
            <div id="twa-st02buf" style="display:none">
                <div class="module-header" style="margin-top:4px">
                    <div class="module-crumb">Analysis › Step 02 › Memory Buffer (optional)</div>
                    <h1 class="module-title">Step 3 — Memory Buffer Analysis</h1>
                    <p class="module-desc">
                        Optionally upload the <strong>ST02 → Tune summary</strong> buffer exports
                        (one file per application server). The tool validates each file, consolidates
                        the <strong>Buffer</strong> table across all servers and derives the memory
                        requirement per buffer, per server and for the whole landscape.
                        This step is <strong>optional</strong> — you may skip straight to the next step.
                    </p>
                </div>

                <div class="alert-info" style="margin-bottom:16px">
                    <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                        <circle cx="12" cy="12" r="10"/>
                        <line x1="12" y1="8" x2="12" y2="12"/>
                        <line x1="12" y1="16" x2="12.01" y2="16"/>
                    </svg>
                    <div style="font-size:12px">
                        Each file's first line must read <code>System: &lt;host&gt; Tune summary</code> and the file
                        must contain the <strong>Buffer</strong>, <strong>SAP Memory</strong> and
                        <strong>Call Statistics</strong> sections. Only the Buffer table is used; the
                        <code>Tables</code> classification label row is ignored.
                    </div>
                </div>

                <div class="card" id="twa-st02buf-card">
                    <div class="card-header" style="padding:14px 18px 10px">
                        <div>
                            <div class="card-title">ST02 Buffer Analysis Files</div>
                            <div class="card-sub">
                                SAP ST02 Tune summary exports — <strong>.txt · .csv · .xlsx</strong>
                                &nbsp;·&nbsp; one file per application server · multiple files supported
                            </div>
                        </div>
                    </div>
                    <div class="upload-zone" id="twa-st02buf-zone" style="min-height:120px">
                        <input type="file" class="file-input-hidden" id="twa-st02buf-input"
                               accept=".txt,.text,.csv,.xlsx,.xls" multiple>
                        <div class="upload-zone-icon">
                            <svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24"
                                 fill="none" stroke="currentColor" stroke-width="1.6">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="17 8 12 3 7 8"/>
                                <line x1="12" y1="3" x2="12" y2="15"/>
                            </svg>
                        </div>
                        <div class="upload-zone-text">
                            <div class="upload-zone-title">Upload ST02 Buffer Analysis files</div>
                            <div class="upload-zone-sub">
                                Drag &amp; drop or browse — .txt · .csv · .xlsx · multiple files accepted
                            </div>
                        </div>
                        <span class="upload-zone-cta">
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                                 fill="none" stroke="currentColor" stroke-width="2.2">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="17 8 12 3 7 8"/>
                                <line x1="12" y1="3" x2="12" y2="15"/>
                            </svg>
                            Choose files
                        </span>
                    </div>
                    <div id="twa-st02buf-file-list" style="display:none;padding:10px 18px 14px">
                        <div style="font-size:12px;font-weight:600;color:var(--text-muted);
                                    text-transform:uppercase;letter-spacing:.04em;margin-bottom:6px">
                            Files queued for analysis
                        </div>
                        <ul id="twa-st02buf-files-ul" style="list-style:none;padding:0;margin:0"></ul>
                    </div>
                </div>

                <div id="twa-st02buf-progress" class="proc-progress" style="display:none;margin-top:18px">
                    <div class="proc-progress-head">
                        <span id="twa-st02buf-progress-label">Processing ST02 buffer files…</span>
                        <span id="twa-st02buf-progress-pct">0%</span>
                    </div>
                    <div class="proc-progress-track">
                        <div class="proc-progress-fill" id="twa-st02buf-progress-fill"></div>
                    </div>
                </div>

                <div id="twa-st02buf-summary" style="display:none"></div>

                <div class="action-row" style="margin-top:20px;display:flex;gap:10px">
                    <button class="btn" id="twa-st02buf-back">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="15 18 9 12 15 6"/>
                        </svg>
                        Previous
                    </button>
                    <button class="btn btn-blue" id="twa-st02buf-next" style="margin-left:auto">
                        Next: User Storage for WP Memory (optional)
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="9 18 15 12 9 6"/>
                        </svg>
                    </button>
                </div>
            </div>

            <!-- ══ Step 4 — User Storage for WP Memory Analysis (optional, ST02) ══ -->
            <div id="twa-st02wp" style="display:none">
                <div class="module-header" style="margin-top:4px">
                    <div class="module-crumb">Analysis › Step 02 › WP Memory Storage (optional)</div>
                    <h1 class="module-title">Step 4 — User Storage for WP Memory Analysis</h1>
                    <p class="module-desc">
                        Optionally upload the <strong>ST02 → Storage Details</strong> exports
                        (one file per application server). The tool extracts the user storage for all
                        work processes and the SAP executables size, cross-references the
                        <strong>Physical Memory</strong> from the System Landscape, and summarises the
                        requirement per server and for the whole landscape.
                        This step is <strong>optional</strong>.
                    </p>
                </div>

                <div class="alert-info" style="margin-bottom:16px">
                    <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                        <circle cx="12" cy="12" r="10"/>
                        <line x1="12" y1="8" x2="12" y2="12"/>
                        <line x1="12" y1="16" x2="12.01" y2="16"/>
                    </svg>
                    <div style="font-size:12px">
                        Each file's first line must read <code>System: &lt;host&gt; Date dd.mm.yyyy</code> and the
                        file must contain the <strong>Storage Shared Between Work Processes</strong>,
                        <strong>User Storage for Each of … Work Processes</strong>,
                        <strong>User Storage for all Work Processes</strong> and
                        <strong>Maximum Heap Area for all Work Processes</strong> sections.
                    </div>
                </div>

                <div class="card" id="twa-st02wp-card">
                    <div class="card-header" style="padding:14px 18px 10px">
                        <div>
                            <div class="card-title">ST02 Storage Details Files</div>
                            <div class="card-sub">
                                SAP ST02 Storage Details exports — <strong>.txt · .csv · .xlsx</strong>
                                &nbsp;·&nbsp; one file per application server · multiple files supported
                            </div>
                        </div>
                    </div>
                    <div class="upload-zone" id="twa-st02wp-zone" style="min-height:120px">
                        <input type="file" class="file-input-hidden" id="twa-st02wp-input"
                               accept=".txt,.text,.csv,.xlsx,.xls" multiple>
                        <div class="upload-zone-icon">
                            <svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24"
                                 fill="none" stroke="currentColor" stroke-width="1.6">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="17 8 12 3 7 8"/>
                                <line x1="12" y1="3" x2="12" y2="15"/>
                            </svg>
                        </div>
                        <div class="upload-zone-text">
                            <div class="upload-zone-title">Upload ST02 Storage Details files</div>
                            <div class="upload-zone-sub">
                                Drag &amp; drop or browse — .txt · .csv · .xlsx · multiple files accepted
                            </div>
                        </div>
                        <span class="upload-zone-cta">
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                                 fill="none" stroke="currentColor" stroke-width="2.2">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="17 8 12 3 7 8"/>
                                <line x1="12" y1="3" x2="12" y2="15"/>
                            </svg>
                            Choose files
                        </span>
                    </div>
                    <div id="twa-st02wp-file-list" style="display:none;padding:10px 18px 14px">
                        <div style="font-size:12px;font-weight:600;color:var(--text-muted);
                                    text-transform:uppercase;letter-spacing:.04em;margin-bottom:6px">
                            Files queued for analysis
                        </div>
                        <ul id="twa-st02wp-files-ul" style="list-style:none;padding:0;margin:0"></ul>
                    </div>
                </div>

                <div id="twa-st02wp-progress" class="proc-progress" style="display:none;margin-top:18px">
                    <div class="proc-progress-head">
                        <span id="twa-st02wp-progress-label">Processing ST02 storage files…</span>
                        <span id="twa-st02wp-progress-pct">0%</span>
                    </div>
                    <div class="proc-progress-track">
                        <div class="proc-progress-fill" id="twa-st02wp-progress-fill"></div>
                    </div>
                </div>

                <div id="twa-st02wp-summary" style="display:none"></div>

                <div class="action-row" style="margin-top:20px;display:flex;gap:10px">
                    <button class="btn" id="twa-st02wp-back">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="15 18 9 12 15 6"/>
                        </svg>
                        Previous
                    </button>
                    <button class="btn btn-blue" id="twa-st02wp-next" style="margin-left:auto">
                        Next: Technical Workload – Database
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="9 18 15 12 9 6"/>
                        </svg>
                    </button>
                </div>
            </div>
        </div>`;
    },

    init(app) {
        this.app = app;

        /* Wire the upload zone for MULTIPLE files (and .zip bundles). We bind
           directly here rather than via initUploadZone so the shared single-file
           helper used by other modules is unaffected. */
        const zone  = document.getElementById('twa-zone');
        const input = document.getElementById('twa-input');
        if (zone && input) {
            zone.addEventListener('click', () => input.click());
            zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
            zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
            zone.addEventListener('drop', e => {
                e.preventDefault(); zone.classList.remove('drag-over');
                if (e.dataTransfer.files && e.dataTransfer.files.length) this._handleFiles(e.dataTransfer.files, app);
            });
            input.addEventListener('change', e => {
                if (e.target.files && e.target.files.length) this._handleFiles(e.target.files, app);
            });
        }

        /* Wire the optional SAP System Profile upload zone. */
        this._wirePflZone(app);
        this._wirePflInlineZone(app);
        /* Entry button → open the optional profiles screen; Back → return. */
        document.getElementById('twa-pfl-open')
            ?.addEventListener('click', () => this._showProfilesView());
        document.getElementById('twa-pfl-back')
            ?.addEventListener('click', () => this._hideProfilesView());
        /* Rehydrate the profile list if files were uploaded earlier this session. */
        if ((this.profiles && this.profiles.length) ||
            (this.profilesIgnored && this.profilesIgnored.length)) {
            this._renderProfileList();
        }
        this._updateProfileCount();

        /* Rehydrate on return to this module:
           - a built analysis → show results;
           - an in-progress mapping (awaiting confirmation) → show mapping screen. */
        if (this.targetRows && this.targetRows.length) {
            this._renderResults(app);
        } else if (this._pendingDataRows && this._virtualHosts && this._virtualHosts.length) {
            this._renderMappingScreen(app);
        }
        /* Keep the stepped sidebar sub-nav in sync whenever this module mounts. */
        this._updateSubNav(app, (this.targetRows && this.targetRows.length) ? 'charts' : 'upload');
    },

    /* ══════════════════════════════════════════════════════
       MULTI-FILE / ZIP INTAKE
       Reads every uploaded file (expanding any .zip), parses each to a grid of
       rows, then concatenates all rows that share the same source table so the
       set of files is analysed as if it were one consolidated export.
       ══════════════════════════════════════════════════════ */
    async _handleFiles(fileList, app) {
        this._hideError();
        /* Fresh upload — clear prior state and hide panels. */
        this._mappingActive = false; this._mapping = {}; this._virtualHosts = [];
        this._removedHosts = null; this._pendingDataRows = null; this._pendingMeta = null;
        this.targetRows = []; this.targetCols = [];
        const mapEl = document.getElementById('twa-mapping');
        if (mapEl) { mapEl.style.display = 'none'; mapEl.innerHTML = ''; }
        const resEl = document.getElementById('twa-results');
        if (resEl) resEl.style.display = 'none';

        try {
            /* 1. Expand files: unzip any .zip into its member files. */
            this._showProgress('Reading files…', 3);
            await this._nextFrame();
            const files = [];
            for (const f of Array.from(fileList)) {
                const ext = (f.name.split('.').pop() || '').toLowerCase();
                if (ext === 'zip') {
                    const inner = await this._unzip(f);
                    inner.forEach(x => files.push(x));
                } else {
                    files.push(f);
                }
            }
            if (!files.length) { this._hideProgress(); this._showError('No readable files were found in the upload.'); return; }

            /* 2. Read + parse each file to a grid of rows. */
            const names = [];
            const perFile = [];
            for (let i = 0; i < files.length; i++) {
                this._showProgress(`Reading ${files[i].name} (${i + 1}/${files.length})…`, 5 + Math.round((i / files.length) * 20));
                await this._nextFrame();
                const rows = await this._fileToRows(files[i]);
                if (rows && rows.length) { perFile.push(rows); names.push(files[i].name); }
            }
            if (!perFile.length) { this._hideProgress(); this._showError('None of the uploaded files could be read.'); return; }

            this.fileName = names.length === 1 ? names[0] : `${names.length} files`;
            showFileLoaded('twa-status', names.length === 1 ? names[0] : `${names.length} files: ${names.join(', ')}`);

            /* 3. Consolidate: detect each file's source table + header, then merge
               all data rows (aligned by column NAME, so differing column order or
               reduced column sets across files still combine correctly). */
            const merged = this._consolidateFiles(perFile);
            if (!merged) {
                this._hideProgress();
                /* Help the user diagnose: show the header row we actually saw in
                   the first file, so column-name mismatches are easy to spot. */
                let seen = '';
                const first = perFile[0] || [];
                const escHtml = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
                for (let i = 0; i < Math.min(first.length, 30); i++) {
                    const cells = (first[i] || []).map(c => String(c).trim()).filter(Boolean);
                    if (cells.length >= 3) { seen = cells.map(escHtml).join(' · '); break; }
                }
                this._showError(
                    'Could not recognise any of the uploaded files as /SDF/SMON_HEADER, ' +
                    '/SDF/MON_HEADER, or a pre-processed statistics spreadsheet.' +
                    (seen ? `<br><br><strong>Columns detected in the file:</strong> ${seen}` +
                            `<br><br>A pre-processed file must include a server column (Server Name / ` +
                            `Application Server / Hostname / Host), “CPU Utilization (%)”, and ` +
                            `“SAPS Utilization” or “EM Configured (GB)”.` : '')
                );
                return;
            }
            this.sourceTable = merged.source;
            /* _processGrid expects a grid whose rows[headerIdx] is the header and
               subsequent rows are data; we hand it a synthetic grid starting with
               the header row then all merged data rows. */
            await this._processGrid(merged.grid, app);
        } catch (err) {
            this._hideProgress();
            this._showError(`Failed to process the upload: ${err && err.message ? err.message : err}`);
        }
    },

    /* Decompress a .zip File → array of { name, async-readable } File-like wrappers.
       Only member files with a recognised statistics extension are returned. */
    async _unzip(zipFile) {
        if (typeof JSZip === 'undefined') throw new Error('ZIP support failed to load (JSZip). Reload and try again.');
        const buf = await zipFile.arrayBuffer();
        const zip = await JSZip.loadAsync(buf);
        const out = [];
        const names = Object.keys(zip.files);
        for (const n of names) {
            const entry = zip.files[n];
            if (entry.dir) continue;
            const base = n.split('/').pop();
            if (!base || base.startsWith('.') || base.startsWith('__MACOSX')) continue;
            const ext = (base.split('.').pop() || '').toLowerCase();
            if (!['txt', 'text', 'rtf', 'csv', 'xlsx', 'xls', 'html', 'htm'].includes(ext)) continue;
            /* Build a File-like object exposing the readers _fileToRows needs. */
            const isBinary = ['xlsx', 'xls'].includes(ext);
            const data = isBinary ? await entry.async('binarystring') : await entry.async('string');
            out.push({ name: base, _ext: ext, _binary: isBinary, _data: data, _fromZip: true });
        }
        return out;
    },

    /* Read any supported file (real File or zip-member wrapper) into a grid of
       rows (array-of-arrays). Returns a Promise. */
    _fileToRows(file) {
        const ext = file._ext || (file.name.split('.').pop() || '').toLowerCase();
        const toGridFromBinary = (bin) => {
            const wb = XLSX.read(bin, { type: 'binary' });
            const ws = wb.Sheets[wb.SheetNames[0]];
            return this._gridToRows(XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }));
        };
        const toGridFromText = (text) => {
            if (ext === 'rtf') text = this._stripRtf(text);
            if (['html', 'htm'].includes(ext)) text = this._htmlToText(text);
            return this._textToRows(text);
        };
        /* zip-member wrapper already carries its decoded data */
        if (file._fromZip) {
            return Promise.resolve(file._binary ? toGridFromBinary(file._data) : toGridFromText(file._data));
        }
        /* real File → FileReader */
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            if (['xlsx', 'xls'].includes(ext)) {
                reader.onload = e => { try { resolve(toGridFromBinary(e.target.result)); } catch (err) { reject(err); } };
                reader.onerror = () => reject(new Error('read error'));
                reader.readAsBinaryString(file);
            } else {
                reader.onload = e => { try { resolve(toGridFromText(e.target.result)); } catch (err) { reject(err); } };
                reader.onerror = () => reject(new Error('read error'));
                reader.readAsText(file);
            }
        });
    },

    /* Consolidate multiple parsed grids into one. Each grid is detected for its
       source table + header row; data rows from every file of the SAME source
       table are merged. Columns are aligned by NAME (not position) so files with
       different column orders or reduced columns still combine. The dominant
       source table (most files) wins; files of the other type are skipped.
       Returns { source, grid } where grid[0] is the header and grid[1..] data. */
    _consolidateFiles(perFile) {
        const detected = perFile.map(rows => {
            const d = this._detectSource(rows);
            return d ? { rows, ...d } : null;
        }).filter(Boolean);
        if (!detected.length) return null;

        /* Pick the dominant source table. */
        const counts = {};
        detected.forEach(d => { counts[d.source] = (counts[d.source] || 0) + 1; });
        const source = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0];
        const files = detected.filter(d => d.source === source);

        /* Union of header column names, in the order first seen (keeps the
           canonical SMON/MON order from the first file). Headers are translated
           to canonical field NAMES first, so files that use field LABELS
           ("Application Server Instance", …) merge on the same keys as files that
           use field names — and the consolidated header is itself field names,
           which is what _mapColumns / _buildTarget expect downstream. */
        const colNames = [];
        const seenCol = new Set();
        files.forEach(f => f.headers.forEach(h => {
            const k = this._canonicalFieldName(String(h).trim());
            const key = k.toUpperCase().replace(/\s+/g, '');
            if (key && !seenCol.has(key)) { seenCol.add(key); colNames.push(k); }
        }));

        /* Build the consolidated grid: header row, then every file's data rows
           remapped into the unified column order. */
        const grid = [colNames.slice()];
        files.forEach(f => {
            const colMap = this._mapColumns(f.headers);   /* canonical NAME→index in this file */
            const idxFor = colNames.map(name => {
                const k = name.toUpperCase().replace(/\s+/g, '');
                return colMap[k] != null ? colMap[k] : -1;
            });
            for (let i = f.headerIdx + 1; i < f.rows.length; i++) {
                const src = f.rows[i];
                if (!src || !src.length) continue;
                const line = new Array(colNames.length);
                for (let c = 0; c < idxFor.length; c++) {
                    const si = idxFor[c];
                    line[c] = si >= 0 ? (src[si] != null ? src[si] : '') : '';
                }
                grid.push(line);
            }
        });
        return { source, grid };
    },

    /* ── RTF → plain text (strip control words / groups, keep cell text) ── */
    _stripRtf(rtf) {
        let t = rtf;
        t = t.replace(/\\par[d]?/g, '\n');
        t = t.replace(/\\cell/g, '|');
        t = t.replace(/\\row/g, '\n');
        t = t.replace(/\{\\\*[^{}]*\}/g, '');       /* ignorable groups */
        t = t.replace(/\\'[0-9a-fA-F]{2}/g, ' ');   /* hex escapes */
        t = t.replace(/\\[a-zA-Z]+-?\d* ?/g, '');   /* control words */
        t = t.replace(/[{}]/g, '');
        return t;
    },

    /* ── HTML → text, converting table cells to the "|" separator ── */
    _htmlToText(html) {
        const dom = new DOMParser().parseFromString(html, 'text/html');
        const tables = dom.querySelectorAll('table');
        if (tables.length) {
            let out = '';
            tables.forEach(tbl => {
                tbl.querySelectorAll('tr').forEach(tr => {
                    const cells = [...tr.querySelectorAll('th,td')]
                        .map(c => (c.textContent || '').trim());
                    out += '|' + cells.join('|') + '|\n';
                });
            });
            return out;
        }
        return dom.body ? (dom.body.innerText || dom.body.textContent || '') : html;
    },

    /* ── Plain-text table → array of cell-arrays ──
       Handles the SAP "|"-delimited dump. Lines made only of dashes
       (horizontal rules) and the "Table:" / "Displayed Fields:" banner
       lines are discarded. */
    _textToRows(text) {
        /* Iterate line-by-line over the ORIGINAL string (handling \r\n and \r)
           rather than building a giant replaced copy plus a lines[] array. For a
           month of per-minute snapshots this input can be hundreds of MB, so we
           avoid holding two extra full-size copies in memory at once. */
        const rows = [];
        const n = text.length;
        let i = 0;
        while (i < n) {
            let j = i;
            /* find end of line */
            while (j < n) {
                const ch = text.charCodeAt(j);
                if (ch === 10 || ch === 13) break;   /* \n or \r */
                j++;
            }
            let line = text.slice(i, j).replace(/\s+$/, '');   /* trimEnd */
            /* advance past the line terminator (handle \r\n as one) */
            if (j < n && text.charCodeAt(j) === 13 && j + 1 < n && text.charCodeAt(j + 1) === 10) i = j + 2;
            else i = j + 1;

            if (!line.trim()) continue;
            if (this._isRulerLine(line)) continue;             /* horizontal rule */
            if (line.indexOf('|') >= 0) {
                let cells = line.split('|');
                for (let c = 0; c < cells.length; c++) cells[c] = cells[c].trim();
                if (cells.length && cells[0] === '') cells.shift();
                if (cells.length && cells[cells.length - 1] === '') cells.pop();
                rows.push(cells);
            } else {
                rows.push([line.trim()]);
            }
        }
        return rows;
    },

    /* True if the line is a table ruler: only dashes / spaces / pipes / equals,
       with at least three dash characters. Implemented as a linear character
       scan (NOT a regex) — the SAP dumps contain ruler lines hundreds of chars
       long, and an ambiguous regex like /^[\s|]*[-–—=]{3,}[\s|]*$/ backtracks
       catastrophically on them ("Maximum call stack size exceeded"). */
    _isRulerLine(line) {
        let dashes = 0;
        for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            if (ch === '-' || ch === '–' || ch === '—' || ch === '=') { dashes++; continue; }
            if (ch === ' ' || ch === '\t' || ch === '|') continue;
            return false;   /* some other character ⇒ not a ruler */
        }
        return dashes >= 3;
    },

    /* ── XLSX matrix → cell-arrays, trimming each value ── */
    _gridToRows(grid) {
        return grid.map(r => r.map(c => (c == null ? '' : String(c).trim())));
    },

    /* ══════════════════════════════════════════════════════
       CORE PROCESSING
       ══════════════════════════════════════════════════════ */
    async _processGrid(rows, app) {
        if (!rows || !rows.length) { this._hideProgress(); this._showError('The file appears to be empty.'); return; }

        this._showProgress('Detecting source table…', 25);
        await this._nextFrame();

        /* ── Identify source table & header row ── */
        const detected = this._detectSource(rows);
        if (!detected) {
            this._hideProgress();
            this._showError(
                'Could not recognise this as a /SDF/SMON_HEADER or /SDF/MON_HEADER export, ' +
                'nor as a pre-processed statistics spreadsheet. For a raw export the header row ' +
                'should list SERVER, ACT_WPS, IDLE_TOTAL, …; for a pre-processed file it should ' +
                'list Server Name, CPU Utilization (%), SAPS Utilization, …'
            );
            return;
        }
        const { source, headerIdx, headers } = detected;
        this.sourceTable = source;

        /* ── Pre-processed spreadsheet path ──
           Validate required columns, then build the target directly (no SAP
           derivation, no hostname mapping, no EM-config step — the file already
           carries EM Configured (GB) and all landscape-derived values). */
        if (source === 'PREPARSED') {
            const missing = this._validatePreparsed(headers);
            if (missing.length) {
                this._hideProgress();
                this._showError(
                    'The processed statistics file is missing required column(s): ' +
                    missing.join(', ') +
                    '. Please add the missing column(s) and upload the file again.'
                );
                return;
            }

            /* Collect data rows that have a server value (any alias). */
            const norm = s => String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
            let ppSrvIdx = -1;
            headers.forEach((h, i) => { if (ppSrvIdx < 0 && this._PP_SERVER_ALIASES.includes(norm(h))) ppSrvIdx = i; });
            const ppRows = [];
            for (let i = headerIdx + 1; i < rows.length; i++) {
                const r = rows[i];
                rows[i] = null;
                if (!r || !r.length) continue;
                const sv = ppSrvIdx >= 0 ? String(r[ppSrvIdx] || '').trim() : '';
                if (!sv) continue;
                ppRows.push(r);
            }
            rows.length = 0;
            if (!ppRows.length) {
                this._hideProgress();
                this._showError('No data rows with a server value were found below the header.');
                return;
            }

            this._mappingActive = false;
            this._showProgress(`Building analysis (${ppRows.length} rows)…`, 55);
            await this._nextFrame();
            try {
                const built = this._buildTargetFromPreparsed(headers, ppRows, app);
                ppRows.length = 0;
                this.targetCols = built.cols;
                this.targetRows = built.rows;

                this._showProgress('Finalising…', 90);
                await this._nextFrame();

                app.setModuleData('tech-workload-app', {
                    rows:        built.rows,
                    fileName:    this.fileName,
                    headers:     built.cols,
                    analysis:    built.rows,
                    sourceTable: source,
                });
                app.markCompleted('tech-workload-app');

                this._renderResults(app);
                this._showProgress('Analysis ready', 100);
                setTimeout(() => this._hideProgress(), 700);
                showToast('Target analysis ready',
                    `${built.rows.length.toLocaleString()} rows processed from the pre-processed file. ` +
                    `Click “Download Analysis” to save the file.`,
                'success');
            } catch (err) {
                this._hideProgress();
                this._showError(`Failed to build the analysis: ${err && err.message ? err.message : err}`);
                showToast('Analysis failed', err && err.message ? err.message : String(err), 'error');
            }
            return;
        }

        /* ── Map header names → column index (flexible, case-insensitive) ── */
        const col = this._mapColumns(headers);

        /* ── Data rows: everything after the header that has a SERVER value ──
           Move references into dataRows and null the source slot so the full
           parsed grid can be garbage-collected as we go (large files). */
        const dataRows = [];
        const srvIdx = col.SERVER;
        for (let i = headerIdx + 1; i < rows.length; i++) {
            const r = rows[i];
            rows[i] = null;
            if (!r || !r.length) continue;
            const serverRaw = srvIdx != null ? (r[srvIdx] || '').trim() : '';
            if (!serverRaw) continue;
            dataRows.push(r);
        }
        rows.length = 0;   /* release the (now-empty) source grid */

        if (!dataRows.length) {
            this._hideProgress();
            const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
            const hdrPreview = (headers || []).filter(Boolean).slice(0, 60).map(esc).join(' · ');
            const srvNote = (srvIdx == null)
                ? 'The <strong>SERVER</strong> / “Application Server Instance” column could not be located in the header.'
                : `The SERVER column was located (position ${srvIdx + 1}) but every data row below the header was empty there.`;
            this._showError(
                'No data rows with a SERVER value were found below the header.<br><br>' +
                srvNote +
                (hdrPreview ? `<br><br><strong>Header columns detected:</strong> ${hdrPreview}` : '')
            );
            return;
        }

        /* ── Virtual↔physical hostname check ──
           If a System Landscape exists but the SMON/MON hostnames don't all match
           a landscape host, the customer is likely using virtual hostnames in the
           performance data. Pause and ask the user to map them. */
        this._showProgress('Checking hostnames against the landscape…', 45);
        await this._nextFrame();

        const landscape = app.getHostVariables();
        const distinctHosts = [];
        const seen = new Set();
        for (const r of dataRows) {
            const h = this._hostOf(col.SERVER != null ? (r[col.SERVER] || '') : '');
            const key = h.toLowerCase();
            if (h && !seen.has(key)) { seen.add(key); distinctHosts.push(h); }
        }
        const unmatched = distinctHosts.filter(h => !app.findHost(h));

        if (landscape.length > 0 && unmatched.length > 0) {
            /* Hold the parsed data and show the mapping screen. */
            this._hideProgress();
            this._pendingDataRows = dataRows;
            this._pendingMeta     = { source, headers, col };
            this._virtualHosts    = distinctHosts;

            /* Seed the mapping so the screen opens pre-filled.
               Priority per virtual host:
                 1. a direct name match against a landscape host (findHost), then
                 2. the EWA §2.2 alias map (virtual→physical from Instance Name /
                    SAPDBHOST), captured by Module 01 as `hostAliases`.
               The alias map is keyed physical→virtual, so invert it once. Each
               physical server takes at most one virtual (the screen enforces
               one-to-one), so skip a physical that's already been assigned. */
            this._mapping = {};
            const slData   = app.getModuleData('system-landscape') || {};
            const aliases  = slData.hostAliases || {};           /* physicalLower → virtual */
            const landscapeSet = new Set(landscape.map(h => (h.hostname || '').toLowerCase()));
            const takenPhys    = new Set();                      /* physical already mapped */
            /* virtualLower → physical hostname (original case from the landscape) */
            const virtualToPhys = {};
            Object.keys(aliases).forEach(physLower => {
                const virt = aliases[physLower];
                if (!virt) return;
                const hostRec = landscape.find(h => (h.hostname || '').toLowerCase() === physLower);
                if (hostRec) virtualToPhys[virt.toLowerCase()] = hostRec.hostname;
            });

            distinctHosts.forEach(h => {
                const key = h.toLowerCase();
                /* 1. direct landscape name match */
                const hit = app.findHost(h);
                if (hit) {
                    this._mapping[key] = hit.hostname;
                    takenPhys.add((hit.hostname || '').toLowerCase());
                    return;
                }
                /* 2. EWA alias match (virtual → physical) */
                const phys = virtualToPhys[key];
                if (phys && landscapeSet.has(phys.toLowerCase()) && !takenPhys.has(phys.toLowerCase())) {
                    this._mapping[key] = phys;
                    takenPhys.add(phys.toLowerCase());
                }
            });
            this._renderMappingScreen(app);
            return;
        }

        /* All hosts matched (or no landscape to map against) — collect the SAP
           Extended Memory configuration before building (no mapping needed). */
        this._hideProgress();
        this._mappingActive   = false;
        this._pendingDataRows = dataRows;
        this._pendingMeta     = { source, headers, col };
        this._virtualHosts    = distinctHosts;
        this._renderEmConfigScreen(app);
    },

    /* Confirm handler for the mapping screen's "Confirm mapping & process" button.
       Validates that every virtual host is mapped, then advances to the SAP
       Extended Memory configuration step (the build runs from there). */
    _runBuildWithMapping(app) {
        if (!this._pendingDataRows || !this._pendingMeta) return;

        /* If unmapped virtual hosts remain (e.g. DB servers or infrastructure
           hosts whose hostnames appear in the SMON/MON data but have no
           corresponding physical host in the EWA report), allow proceeding.
           Their rows will be included in the enriched table but will show
           "No match in EWA report" in the SERVER (EWA) column and will have
           blank Cores/SAPS values. */
        const assigned = new Set(Object.keys(this._mapping));
        const unmapped = this._virtualHosts.filter(v => !assigned.has(v.toLowerCase()));
        if (unmapped.length) {
            /* Show an informational note, not a blocking error. */
            this._showMappingError(
                `${unmapped.length} virtual host(s) are not mapped to a physical server: ` +
                unmapped.join(', ') +
                `. These hosts will appear in the enriched table with "No match in EWA report" ` +
                `and blank Cores/SAPS values. You can still proceed.`
            );
            /* Show a "Proceed anyway" button if it doesn't already exist. */
            const existingProceed = document.getElementById('twa-map-proceed-anyway');
            if (!existingProceed) {
                const actRow = document.querySelector('#twa-mapping .action-row');
                if (actRow) {
                    const btn = document.createElement('button');
                    btn.className = 'btn';
                    btn.id = 'twa-map-proceed-anyway';
                    btn.textContent = 'Proceed with unmapped hosts';
                    btn.addEventListener('click', () => {
                        this._clearMappingError();
                        this._mappingActive = true;
                        const mapEl = document.getElementById('twa-mapping');
                        if (mapEl) mapEl.style.display = 'none';
                        this._renderEmConfigScreen(app);
                    });
                    actRow.appendChild(btn);
                }
            }
            return;
        }
        this._clearMappingError();

        this._mappingActive = true;
        /* Hide the mapping screen and advance to the EM configuration step. */
        const mapEl = document.getElementById('twa-mapping');
        if (mapEl) mapEl.style.display = 'none';
        this._renderEmConfigScreen(app);
    },

    /* Build the enriched target table from the held rows. Shared by both the
       mapping path and the direct-match path; branches on this._mappingActive
       only for the hostMapping payload. Invoked by the EM-config step's
       "Build Analysis" button. */
    async _doBuild(app) {
        if (!this._pendingDataRows || !this._pendingMeta) return;

        /* Re-entrancy guard: the build is async (awaits _nextFrame) and the
           button can accumulate listeners across re-renders, so a single click
           may dispatch this handler several times. Without the guard a second
           invocation would run _buildTarget against the SAME dataRows array
           after the first build emptied it. Bail if a build is already in flight. */
        if (this._buildingWithMapping) return;
        this._buildingWithMapping = true;

        const { source, headers, col } = this._pendingMeta;
        /* Take ownership of the held rows synchronously (before any await) so a
           re-entrant call sees _pendingDataRows === null and bails above. */
        const dataRows = this._pendingDataRows;
        this._pendingDataRows = null;

        /* Hide the EM-config screen; show the processing progress below it. */
        const emEl = document.getElementById('twa-emconfig');
        if (emEl) emEl.style.display = 'none';
        this._showProgress(`Building analysis (${dataRows.length} rows)…`, 20);
        await this._nextFrame();

        try {
            const built = this._buildTarget(source, headers, col, dataRows, app);
            dataRows.length = 0;              /* release source rows */
            this.targetCols = built.cols;
            this.targetRows = built.rows;

            this._showProgress('Finalising…', 90);
            await this._nextFrame();

            const payload = {
                rows:        built.rows,
                fileName:    this.fileName,
                headers:     built.cols,
                analysis:    built.rows,
                sourceTable: source,
            };
            if (this._mappingActive) payload.hostMapping = { ...this._mapping };
            app.setModuleData('tech-workload-app', payload);
            app.markCompleted('tech-workload-app');

            this._renderResults(app);
            this._showProgress('Analysis ready', 100);
            setTimeout(() => this._hideProgress(), 700);
            showToast('Target analysis ready',
                `${built.rows.length.toLocaleString()} rows processed` +
                (this._mappingActive ? ' with the hostname mapping' : '') +
                `. Click “Download Analysis” to save the file.`,
            'success');
        } catch (err) {
            this._hideProgress();
            this._showError(`Failed to build the analysis: ${err && err.message ? err.message : err}`);
            showToast('Analysis failed', err && err.message ? err.message : String(err), 'error');
        } finally {
            this._buildingWithMapping = false;
        }
    },

    /* ══════════════════════════════════════════════════════
       SAP SYSTEM PROFILE FILES (optional)
       Uploaded alongside the SMON/MON data. Parsed, classified by
       filename, validated against the in-scope SID (SAPSYSTEMNAME),
       and retained in this.profiles for the Extended Memory step and
       later modules.
       ══════════════════════════════════════════════════════ */

    _wirePflZone(app) {
        const zone  = document.getElementById('twa-pfl-zone');
        const input = document.getElementById('twa-pfl-input');
        if (!zone || !input || zone._wired) return;
        zone._wired = true;
        zone.addEventListener('click', () => input.click());
        zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
        zone.addEventListener('drop', e => {
            e.preventDefault(); zone.classList.remove('drag-over');
            if (e.dataTransfer.files && e.dataTransfer.files.length) this._handleProfileFiles(e.dataTransfer.files, app);
        });
        input.addEventListener('change', e => {
            if (e.target.files && e.target.files.length) this._handleProfileFiles(e.target.files, app);
        });
    },

    _wirePflInlineZone(app) {
        /* The card's inner HTML is replaced dynamically, so wire at the card level
           using event delegation — zone and input are re-created on each render. */
        const card = document.getElementById('twa-pfl-inline-card');
        if (!card || card._inlineWired) return;
        card._inlineWired = true;

        card.addEventListener('click', e => {
            /* Upload zone click → trigger hidden file input */
            if (e.target.closest('#twa-pfl-inline-zone')) {
                document.getElementById('twa-pfl-inline-input')?.click();
                return;
            }
            /* "View Files" toggle */
            if (e.target.closest('#twa-pfl-inline-view')) {
                const detail = document.getElementById('twa-pfl-inline-detail');
                if (detail) detail.style.display = detail.style.display === 'none' ? 'block' : 'none';
                const btn = document.getElementById('twa-pfl-inline-view');
                if (btn) btn.textContent = detail && detail.style.display !== 'none' ? 'Hide Files' : 'View Files';
                return;
            }
            /* "Replace Files" → re-render as upload panel */
            if (e.target.closest('#twa-pfl-inline-replace')) {
                this._pflInlineMode = 'replace';
                this._renderPflInlineCard();
                return;
            }
            /* "Remove All" → clear profiles and re-render upload panel */
            if (e.target.closest('#twa-pfl-inline-remove')) {
                this.profiles = []; this.profilesIgnored = [];
                this._saveProfilesToApp();
                this._renderProfileList();
                this._renderPflInlineCard();
                return;
            }
        });

        card.addEventListener('dragover', e => {
            if (e.target.closest('#twa-pfl-inline-zone')) { e.preventDefault(); e.target.closest('#twa-pfl-inline-zone').classList.add('drag-over'); }
        });
        card.addEventListener('dragleave', e => {
            const zone = e.target.closest('#twa-pfl-inline-zone');
            if (zone) zone.classList.remove('drag-over');
        });
        card.addEventListener('drop', e => {
            const zone = e.target.closest('#twa-pfl-inline-zone');
            if (!zone) return;
            e.preventDefault(); zone.classList.remove('drag-over');
            if (e.dataTransfer.files && e.dataTransfer.files.length) this._handleProfileFiles(e.dataTransfer.files, app);
        });
        card.addEventListener('change', e => {
            if (e.target.id === 'twa-pfl-inline-input' && e.target.files && e.target.files.length) {
                this._handleProfileFiles(e.target.files, app);
            }
        });
    },

    /* In-scope SAP application SID (upper-case, 3 chars) from Module 01. */
    _profileScopeSid() {
        const d = this.app && this.app.getModuleData && this.app.getModuleData('system-landscape');
        const sid = (d && d.systemInfo && d.systemInfo.sid) || '';
        return String(sid).replace(/[^A-Za-z0-9]/g, '').slice(0, 3).toUpperCase();
    },

    async _handleProfileFiles(fileList, app) {
        /* Expand any .zip bundles into individual member files (read as text). */
        const items = [];   /* { name, text } */
        for (const f of Array.from(fileList)) {
            if (/\.zip$/i.test(f.name) && typeof JSZip !== 'undefined') {
                try {
                    const zip = await JSZip.loadAsync(await f.arrayBuffer());
                    for (const n of Object.keys(zip.files)) {
                        const entry = zip.files[n];
                        if (entry.dir) continue;
                        const base = n.split('/').pop();
                        if (!base || base.startsWith('.') || base.startsWith('__MACOSX')) continue;
                        items.push({ name: base, text: await entry.async('string') });
                    }
                } catch (err) {
                    showToast('ZIP read failed', String(err && err.message || err), 'error');
                }
            } else {
                items.push({ name: f.name, text: await f.text() });
            }
        }
        if (!items.length) return;

        const scopeSid = this._profileScopeSid();
        let kept = 0;
        for (const it of items) {
            const prof = this._parseProfile(it.name, it.text);
            /* De-duplicate by filename across repeated uploads. */
            this.profiles = (this.profiles || []).filter(p => p.name !== prof.name);
            this.profilesIgnored = (this.profilesIgnored || []).filter(p => p.name !== prof.name);

            /* SID validation: a profile that declares SAPSYSTEMNAME must match the
               in-scope SID. DEFAULT.PFL is SID-wide and often omits it → keep.
               If we don't know the scope SID yet, keep everything (can't validate). */
            const declared = (prof.sapsysname || '').toUpperCase();
            if (scopeSid && declared && declared !== scopeSid) {
                this.profilesIgnored.push({
                    name: prof.name,
                    reason: `SAPSYSTEMNAME=${declared} ≠ system in scope (${scopeSid})`,
                });
                continue;
            }
            this.profiles.push(prof);
            kept++;
        }

        this._renderProfileList();
        this._saveProfilesToApp();
        this._updateProfileCount();

        const ignored = this.profilesIgnored.length;
        if (ignored) {
            showToast(
                'Some profiles ignored',
                `${ignored} profile file(s) belong to a different SID and were skipped: ` +
                this.profilesIgnored.map(p => p.name).join(', '),
                'info', 7000
            );
        }
        if (kept) {
            showToast('Profiles loaded', `${kept} SAP system profile file(s) retained for analysis.`, 'success');
        }
    },

    /* Parse one profile file: KEY = VALUE lines → params map, plus filename
       classification. Returns { name, kind, host, instanceNo, sapsysname, params }. */
    _parseProfile(name, text) {
        const params = {};
        for (const raw of String(text).split(/\r?\n/)) {
            const line = raw.trim();
            if (!line || line.startsWith('#') || line.startsWith('*')) continue;
            const m = /^([\w/.\-]+)\s*=\s*(.*)$/.exec(line);
            if (m) params[m[1].trim()] = m[2].trim();
        }

        /* Classify by filename. The trailing token is the (virtual/logical) host. */
        const base = name.replace(/\.(pfl)$/i, '');
        let kind = 'OTHER', host = '', instanceNo = '';
        if (/^DEFAULT(\.PFL)?$/i.test(name) || /^DEFAULT$/i.test(base)) {
            kind = 'DEFAULT';
        } else {
            let mm;
            if ((mm = /_ASCS(\d+)_(.+)$/i.exec(base))) {
                kind = 'ASCS'; instanceNo = mm[1]; host = mm[2];
            } else if ((mm = /_D(\d+)_(.+)$/i.exec(base))) {
                kind = 'INSTANCE'; instanceNo = mm[1]; host = mm[2];
            } else {
                /* Fallback: last underscore-separated token as host. */
                const parts = base.split('_');
                if (parts.length >= 2) host = parts[parts.length - 1];
            }
        }

        return {
            name,
            kind,
            host,
            instanceNo,
            sapsysname: params['SAPSYSTEMNAME'] || '',
            params,
        };
    },

    /* Show the optional SAP System Profiles screen. All other views are hidden;
       the parsed data in this.profiles is retained regardless of visibility.
       `origin` records the screen to return to on Back:
         'upload'   — the SMON/MON upload screen (default),
         'emconfig' — the Extended Memory configuration step,
         'mapping'  — the virtual↔physical mapping screen,
         'results'  — the enriched results screen. */
    _showProfilesView(origin) {
        this._pflOrigin = origin || this._inferScreen();
        /* Hide every other screen/section. */
        ['twa-upload-screen', 'twa-results', 'twa-charts', 'twa-nmon',
         'twa-st02buf', 'twa-st02wp', 'twa-mapping', 'twa-emconfig'].forEach(id => {
            const n = document.getElementById(id);
            if (n) n.style.display = 'none';
        });
        const v = document.getElementById('twa-pfl-view');
        if (v) v.style.display = 'block';
        this._renderProfileList();
        if (this.app && this.app.setSubNavActive) this.app.setSubNavActive('tech-workload-app', 'profiles');
        document.getElementById('content-area')?.scrollTo({ top: 0, behavior: 'smooth' });
    },

    /* Which screen is currently visible (used to remember where to return). */
    _inferScreen() {
        const vis = id => {
            const n = document.getElementById(id);
            return n && n.style.display !== 'none';
        };
        if (vis('twa-emconfig')) return 'emconfig';
        if (vis('twa-mapping'))  return 'mapping';
        if (vis('twa-results') || vis('twa-charts')) return 'results';
        return 'upload';
    },

    /* Return from the profiles screen to whichever screen the user came from.
       When returning to the Extended Memory step, re-render it so the table
       picks up any newly-uploaded profile values (em/initial_size_MB etc.). */
    _hideProfilesView() {
        const v = document.getElementById('twa-pfl-view');
        if (v) v.style.display = 'none';
        this._updateProfileCount();

        const origin = this._pflOrigin || 'upload';
        /* The upload screen stays visible above the mapping / EM-config steps
           (that was its behaviour before the profiles screen existed). */
        const up = document.getElementById('twa-upload-screen');
        if (origin === 'emconfig' && this._pendingDataRows) {
            if (up) up.style.display = 'block';
            this._renderEmConfigScreen(this.app);
            return;
        }
        if (origin === 'mapping' && this._pendingDataRows) {
            if (up) up.style.display = 'block';
            this._renderMappingScreen(this.app);
            return;
        }
        if (origin === 'results' && this.targetRows && this.targetRows.length) {
            this._renderResults(this.app);
            return;
        }
        /* Default: back to the SMON/MON upload screen. */
        if (up) up.style.display = 'block';
        if (this.app) this._updateSubNav(this.app, 'upload');
        document.getElementById('content-area')?.scrollTo({ top: 0, behavior: 'smooth' });
    },

    /* Hide only the profiles screen (used when another screen takes over). */
    _hidePflViewOnly() {
        const v = document.getElementById('twa-pfl-view');
        if (v) v.style.display = 'none';
    },

    /* Reflect the retained-profile count on the entry button row. */
    _updateProfileCount() {
        /* Legacy count label (twa-pfl-view screen) */
        const el = document.getElementById('twa-pfl-count');
        if (el) {
            const n = (this.profiles || []).length;
            const ig = (this.profilesIgnored || []).length;
            el.textContent = n
                ? `${n} profile(s) retained${ig ? ` · ${ig} ignored` : ''}`
                : (ig ? `${ig} ignored` : '');
        }
        /* Inline card — full re-render (collapses to status bar when profiles are loaded) */
        this._renderPflInlineCard();
    },

    /* Render the SAP System Profiles inline card.
       - No profiles loaded  → full upload panel (Step 1 banner + upload zone).
       - Profiles loaded      → compact status bar with View / Replace / Remove actions.
       - _pflInlineMode=replace → force the upload panel even when profiles are present. */
    _renderPflInlineCard() {
        const card = document.getElementById('twa-pfl-inline-card');
        if (!card) return;

        const kept    = this.profiles || [];
        const ignored = this.profilesIgnored || [];
        const hasFiles = kept.length > 0 || ignored.length > 0;
        const showUpload = !hasFiles || this._pflInlineMode === 'replace';

        const esc = s => String(s == null ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        const kindLabel = { DEFAULT: 'DEFAULT', ASCS: 'ASCS', INSTANCE: 'Instance', OTHER: 'Other' };

        if (showUpload) {
            card.innerHTML = `
                <div class="card-header">
                    <div>
                        <div class="card-title">SAP System Profiles Upload (Optional)</div>
                        <div class="card-sub">DEFAULT.PFL · instance profiles · START profiles · .pfl · multiple files · ZIP bundles</div>
                    </div>
                </div>
                <div class="alert-info" style="margin-bottom:12px">
                    <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                        <circle cx="12" cy="12" r="10"/>
                        <line x1="12" y1="8" x2="12" y2="12"/>
                        <line x1="12" y1="16" x2="12.01" y2="16"/>
                    </svg>
                    <span><strong>Step 1 (optional):</strong> optionally upload SAP System Profile files to enrich the analysis — <strong>DEFAULT.PFL</strong>, instance profiles, or START profiles.
                    You may continue without uploading files.</span>
                </div>
                <div class="upload-zone" id="twa-pfl-inline-zone">
                    <input type="file" class="file-input-hidden" id="twa-pfl-inline-input" multiple accept=".pfl,.txt,.zip,*">
                    <div class="upload-zone-icon">
                        <svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="1.6">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                            <line x1="12" y1="18" x2="12" y2="12"/>
                            <line x1="9" y1="15" x2="15" y2="15"/>
                        </svg>
                    </div>
                    <div class="upload-zone-text">
                        <div class="upload-zone-title">Upload SAP System Profile files</div>
                        <div class="upload-zone-sub">Drag &amp; drop or browse — one or more files or a ZIP bundle (.pfl · no extension · .txt · .zip)</div>
                    </div>
                    <span class="upload-zone-cta">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="17 8 12 3 7 8"/>
                            <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        Choose file(s)
                    </span>
                </div>`;
        } else {
            /* Compact status bar */
            const n  = kept.length;
            const ig = ignored.length;
            const countLabel = `${n} file${n !== 1 ? 's' : ''}${ig ? ` · ${ig} ignored` : ''}`;

            const fileListHtml = kept.map((p, i) => {
                const bits = [kindLabel[p.kind] || 'Other'];
                if (p.host) bits.push('host: ' + esc(p.host));
                if (p.instanceNo) bits.push('inst: ' + esc(p.instanceNo));
                const emMb = p.params && p.params['em/initial_size_MB'];
                if (emMb) bits.push('em/initial_size_MB: ' + esc(emMb));
                return `<li style="display:flex;align-items:center;gap:8px;padding:3px 0;font-size:13px">
                    <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24"
                         fill="none" stroke="var(--brand)" stroke-width="2.2">
                        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                        <polyline points="22 4 12 14.01 9 11.01"/>
                    </svg>
                    <span><strong>${esc(p.name)}</strong> <span style="color:var(--text-muted)">— ${bits.slice(1).length ? esc(bits[0]) + ' · ' + bits.slice(1).join(' · ') : esc(bits[0])}</span></span>
                    <button onclick="TechWorkloadAppModule._removeProfile(${i})"
                            style="background:none;border:none;cursor:pointer;color:var(--error);font-size:15px;line-height:1;padding:0 4px;margin-left:auto"
                            title="Remove this profile">×</button>
                </li>`;
            }).join('');

            const ignoredHtml = ig ? `
                <div style="margin-top:8px;font-size:12px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em">Ignored — different SID</div>
                <ul style="list-style:none;padding:0;margin:4px 0 0">
                    ${ignored.map(p => `<li style="display:flex;align-items:center;gap:8px;padding:3px 0;font-size:12.5px;color:var(--text-muted)">
                        <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="var(--error)" stroke-width="2.2">
                            <circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>
                        </svg>
                        <span>${esc(p.name)} — ${esc(p.reason)}</span>
                    </li>`).join('')}
                </ul>` : '';

            card.innerHTML = `
                <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24"
                         fill="none" stroke="var(--success,#28a745)" stroke-width="2.2" style="flex-shrink:0">
                        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                        <polyline points="22 4 12 14.01 9 11.01"/>
                    </svg>
                    <span style="font-size:13.5px;font-weight:600;color:var(--text-primary)">SAP System Profiles: Loaded (${countLabel})</span>
                    <div style="display:flex;gap:8px;margin-left:auto;flex-wrap:wrap">
                        <button id="twa-pfl-inline-view" class="btn btn-sm"
                                style="font-size:12px;padding:3px 10px">View Files</button>
                        <button id="twa-pfl-inline-replace" class="btn btn-sm"
                                style="font-size:12px;padding:3px 10px">Replace Files</button>
                        <button id="twa-pfl-inline-remove" class="btn btn-sm"
                                style="font-size:12px;padding:3px 10px;color:var(--error)">Remove Files</button>
                    </div>
                </div>
                <div id="twa-pfl-inline-detail" style="display:none;margin-top:10px;padding-top:10px;border-top:1px solid var(--border)">
                    ${kept.length ? `<div style="font-size:12px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:4px">Retained profiles</div>
                    <ul style="list-style:none;padding:0;margin:0">${fileListHtml}</ul>` : ''}
                    ${ignoredHtml}
                </div>`;
            /* Reset replace mode once the status bar is rendered */
            this._pflInlineMode = null;
        }
    },

    _renderProfileList() {
        const list = document.getElementById('twa-pfl-list');
        if (!list) return;
        const kept = this.profiles || [];
        const ignored = this.profilesIgnored || [];
        if (!kept.length && !ignored.length) { list.style.display = 'none'; list.innerHTML = ''; return; }

        const esc = s => String(s == null ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        const kindLabel = { DEFAULT: 'DEFAULT', ASCS: 'ASCS', INSTANCE: 'Instance', OTHER: 'Other' };

        const keptHtml = kept.map((p, i) => {
            const bits = [kindLabel[p.kind] || 'Other'];
            if (p.host) bits.push('host: ' + esc(p.host));
            if (p.instanceNo) bits.push('inst: ' + esc(p.instanceNo));
            const emMb = p.params['em/initial_size_MB'];
            if (emMb) bits.push('em/initial_size_MB: ' + esc(emMb));
            return `<li style="display:flex;align-items:center;gap:8px;padding:3px 0;font-size:13px">
                <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24"
                     fill="none" stroke="var(--brand)" stroke-width="2.2">
                    <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                    <polyline points="22 4 12 14.01 9 11.01"/>
                </svg>
                <span><strong>${esc(p.name)}</strong> <span style="color:var(--text-muted)">— ${bits.slice(1).length ? esc(bits[0]) + ' · ' + bits.slice(1).join(' · ') : esc(bits[0])}</span></span>
                <button onclick="TechWorkloadAppModule._removeProfile(${i})"
                        style="background:none;border:none;cursor:pointer;color:var(--error);
                               font-size:15px;line-height:1;padding:0 4px;margin-left:auto"
                        title="Remove">×</button>
            </li>`;
        }).join('');

        const ignoredHtml = ignored.length ? `
            <div style="margin-top:8px;font-size:12px;font-weight:600;color:var(--text-muted);
                        text-transform:uppercase;letter-spacing:.04em">Ignored — different SID</div>
            <ul style="list-style:none;padding:0;margin:4px 0 0">
                ${ignored.map(p => `<li style="display:flex;align-items:center;gap:8px;padding:3px 0;font-size:12.5px;color:var(--text-muted)">
                    <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24"
                         fill="none" stroke="var(--error)" stroke-width="2.2">
                        <circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>
                    </svg>
                    <span>${esc(p.name)} — ${esc(p.reason)}</span>
                </li>`).join('')}
            </ul>` : '';

        list.style.display = 'block';
        list.innerHTML = `
            ${kept.length ? `<div style="font-size:12px;font-weight:600;color:var(--text-muted);
                        text-transform:uppercase;letter-spacing:.04em;margin-bottom:4px">Retained profiles</div>
            <ul style="list-style:none;padding:0;margin:0">${keptHtml}</ul>` : ''}
            ${ignoredHtml}`;
    },

    _removeProfile(idx) {
        if (!this.profiles) return;
        this.profiles.splice(idx, 1);
        this._renderProfileList();
        this._saveProfilesToApp();
        this._updateProfileCount();
    },

    _saveProfilesToApp() {
        if (!this.app) return;
        const prev = this.app.getModuleData('tech-workload-app') || {};
        this.app.setModuleData('tech-workload-app', {
            ...prev,
            profiles: (this.profiles || []).map(p => ({
                name: p.name, kind: p.kind, host: p.host,
                instanceNo: p.instanceNo, sapsysname: p.sapsysname, params: p.params,
            })),
            profilesIgnored: (this.profilesIgnored || []).map(p => ({ name: p.name, reason: p.reason })),
        });
    },

    /* Find the best profile for a given (virtual/logical) hostname.
       Prefers the instance (D##) profile — the spec says that is the most
       relevant — then ASCS, then any other host-matched profile. */
    _profileForHost(hostname) {
        const h = String(hostname || '').toLowerCase();
        if (!h || !this.profiles) return null;
        const matches = this.profiles.filter(p => (p.host || '').toLowerCase() === h);
        if (!matches.length) return null;
        return matches.find(p => p.kind === 'INSTANCE')
            || matches.find(p => p.kind === 'ASCS')
            || matches[0];
    },

    /* ══════════════════════════════════════════════════════
       SAP EXTENDED MEMORY CONFIGURATION STEP
       Shown after the hostname mapping is confirmed (and on the
       direct-match path before building). Collects the configured
       SAP Extended Memory (GB) per virtual hostname into
       this._extMemory, which is then written into the "EM Configured
       (GB)" column of the enriched table during _buildTarget.
       ══════════════════════════════════════════════════════ */
    _renderEmConfigScreen(app) {
        const el = document.getElementById('twa-emconfig');
        const res = document.getElementById('twa-results');
        if (res) res.style.display = 'none';
        if (!el) return;
        if (!this._extMemory) this._extMemory = {};
        /* Dismiss the optional profiles screen if it was open; profiles are
           retained in this.profiles regardless. */
        this._hidePflViewOnly();

        const esc = s => String(s == null ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
        const toNum = v => { const n = parseFloat(v); return isNaN(n) ? null : n; };

        /* The hosts to request EM for are the distinct virtual hostnames found in
           the uploaded statistics (same list used by the mapping screen). */
        const hosts = [...this._virtualHosts].sort((a, b) => a.localeCompare(b));

        /* Resolve the profile-derived parameters per host (if any profile matches).
           em/initial_size_MB is preferred; PHYS_MEMSIZE is the fallback. */
        const hostParam = {};   /* key → { emMb|null, physMb|null, profile } */
        let anyEmMb = false, anyPhysOnly = false;
        hosts.forEach(h => {
            const key = h.toLowerCase();
            const prof = this._profileForHost(h);
            const emMb   = prof ? toNum(prof.params['em/initial_size_MB']) : null;
            const physMb = prof ? toNum(prof.params['PHYS_MEMSIZE']) : null;
            hostParam[key] = { emMb, physMb, profile: prof };
            if (emMb != null) anyEmMb = true;
            else if (physMb != null) anyPhysOnly = true;

            /* Seed EM(GB) from em/initial_size_MB when present and the user hasn't
               already entered a value for this host (don't overwrite user edits). */
            if (emMb != null && (this._extMemory[key] == null || this._extMemory[key] === '')) {
                this._extMemory[key] = Math.round(emMb / 1024 * 100) / 100;
            }
        });

        const srcLabel = this.sourceTable === 'MON' ? '/SDF/MON_HEADER' : '/SDF/SMON_HEADER';
        const haveProfiles = (this.profiles || []).length > 0;

        /* Column headers depend on which profile parameters are available. */
        const extraHeads =
            (anyEmMb     ? '<th style="width:170px">em/initial_size_MB</th>' : '') +
            (anyPhysOnly ? '<th style="width:160px">PHYS_MEMSIZE (MB)</th><th style="width:170px">Proposed SAP EM (GB)</th>' : '');

        const rowsHtml = hosts.map(h => {
            const key   = h.toLowerCase();
            const info  = hostParam[key] || {};
            const saved = this._extMemory[key];
            const gbVal = (saved != null && saved !== '') ? saved : '';

            /* em/initial_size_MB cell (editable, drives GB live). */
            let emCell = '';
            if (anyEmMb) {
                emCell = info.emMb != null
                    ? `<td><div style="display:flex;align-items:center;gap:6px">
                            <input type="number" min="0" step="1"
                                   class="editable-input twa-emmb-input"
                                   data-host="${esc(h)}" value="${info.emMb}" style="width:130px">
                            <span style="color:var(--text-muted);font-size:12px">MB</span>
                       </div></td>`
                    : `<td style="color:var(--text-muted)">—</td>`;
            }

            /* PHYS_MEMSIZE + Proposed EM cells (read-only suggestion). */
            let physCells = '';
            if (anyPhysOnly) {
                if (info.emMb == null && info.physMb != null) {
                    const proposed = Math.round(info.physMb * 0.8 / 1024 * 100) / 100;
                    physCells =
                        `<td>${info.physMb}</td>` +
                        `<td><span style="font-weight:600">${proposed}</span>
                             <span style="color:var(--text-muted);font-size:12px"> GB (80%)</span></td>`;
                } else {
                    physCells = `<td style="color:var(--text-muted)">—</td><td style="color:var(--text-muted)">—</td>`;
                }
            }

            return `<tr>
                <td style="font-weight:600;white-space:nowrap">${esc(h)}</td>
                ${emCell}
                ${physCells}
                <td>
                    <div style="display:flex;align-items:center;gap:8px">
                        <input type="number" min="0" step="0.1"
                               class="editable-input twa-extmem-input"
                               data-host="${esc(h)}"
                               value="${gbVal}"
                               placeholder="e.g. 24"
                               style="width:120px">
                        <span style="color:var(--text-muted);font-size:12px">GB</span>
                    </div>
                </td>
            </tr>`;
        }).join('');

        /* Guidance message adapts to what the profiles provided. */
        let profileNote = '';
        if (haveProfiles && anyEmMb) {
            profileNote = `<div class="alert-info" style="margin-bottom:14px">
                <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                     fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                    <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
                <span><strong>Auto-filled from SAP system profiles.</strong> The
                <code>em/initial_size_MB</code> parameter was read from the uploaded profiles and
                converted to GB (÷1024). Both the MB and GB fields stay editable — changing MB
                recalculates GB automatically.</span>
            </div>`;
        }
        if (haveProfiles && anyPhysOnly) {
            profileNote += `<div class="alert-info" style="margin-bottom:14px;background:var(--warn-bg,#fff8e6);border-color:var(--warn,#e0a800)">
                <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                     fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                    <path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
                    <line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
                </svg>
                <span>For one or more hosts the parameter <code>em/initial_size_MB</code> is
                <strong>not documented</strong> in the SAP profiles. A <strong>Proposed SAP EM (GB)</strong>
                has been calculated as <strong>80% of PHYS_MEMSIZE</strong>. Please validate this value in
                the system and enter the confirmed figure in the <strong>SAP Extended Memory (GB)</strong>
                column — it has been left blank for those hosts.</span>
            </div>`;
        }

        el.style.display = 'block';
        el.innerHTML = `
            <div class="section-title" style="margin-top:20px">
                SAP Extended Memory Configuration
                <span style="font-size:11px;font-weight:400;color:var(--text-muted);margin-left:6px">
                    — enter the SAP Extended Memory (EM) configured for each server
                </span>
            </div>
            <div class="alert-info" style="margin-bottom:14px">
                <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                     fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                    <circle cx="12" cy="12" r="10"/>
                    <line x1="12" y1="8" x2="12" y2="12"/>
                    <line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
                <span>Enter the SAP Extended Memory size (in GB) configured for each hostname found in the
                ${srcLabel} statistics (SAP profile parameter <code>em/initial_size_MB</code> converted to GB,
                or the total EM configured in the instance). Each value is written to the
                <strong>EM Configured (GB)</strong> column of the enriched table for that host's rows.</span>
            </div>
            ${profileNote}
            <div class="card" style="padding:0">
                <div class="card-header" style="padding:14px 18px 12px; border-radius:14px 14px 0 0">
                    <div>
                        <div class="card-title">Extended Memory per Server</div>
                        <div class="card-sub">Enter the configured SAP Extended Memory size in GB for each hostname</div>
                    </div>
                    ${anyPhysOnly ? `<button class="btn" id="twa-emconfig-prefill" title="Copies values from 'Proposed SAP EM (GB)' into 'SAP Extended Memory (GB)'. Only fills empty cells. The copied values remain editable.">
                        <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2">
                            <polyline points="17 1 21 5 17 9"/><path d="M3 11V9a4 4 0 0 1 4-4h14"/>
                            <polyline points="7 23 3 19 7 15"/><path d="M21 13v2a4 4 0 0 1-4 4H3"/>
                        </svg>
                        Prefill SAP Extended Memory
                    </button>` : ''}
                </div>
                <div class="table-wrap" style="border:none;border-radius:0 0 14px 14px;overflow-x:auto">
                    <table class="data-table" id="twa-extmem-table">
                        <thead>
                            <tr>
                                <th>Virtual Hostname</th>
                                ${extraHeads}
                                <th style="width:220px">SAP Extended Memory (GB)</th>
                            </tr>
                        </thead>
                        <tbody id="twa-extmem-tbody">${rowsHtml}</tbody>
                    </table>
                </div>
            </div>

            <div class="action-row" style="margin-top:18px;gap:10px">
                <button class="btn btn-blue" id="twa-emconfig-build">
                    Build Analysis
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2.5">
                        <polyline points="9 18 15 12 9 6"/>
                    </svg>
                </button>
                ${this._mappingActive
                    ? `<button class="btn" id="twa-emconfig-back">Back to mapping</button>`
                    : ''}
            </div>`;

        /* Live-save GB on every keystroke so values survive re-renders. */
        el.querySelectorAll('.twa-extmem-input').forEach(input => {
            input.addEventListener('input', () => {
                const key = (input.dataset.host || '').toLowerCase();
                if (!key) return;
                const n = parseFloat(input.value);
                this._extMemory[key] = isNaN(n) ? null : n;
            });
        });

        /* Editing em/initial_size_MB recalculates the matching GB cell live. */
        el.querySelectorAll('.twa-emmb-input').forEach(input => {
            input.addEventListener('input', () => {
                const host = input.dataset.host || '';
                const key  = host.toLowerCase();
                const mb   = parseFloat(input.value);
                const gb   = isNaN(mb) ? '' : Math.round(mb / 1024 * 100) / 100;
                this._extMemory[key] = (gb === '') ? null : gb;
                const gbInput = el.querySelector(`.twa-extmem-input[data-host="${host.replace(/"/g, '\\"')}"]`);
                if (gbInput) gbInput.value = gb === '' ? '' : gb;
            });
        });

        /* "Prefill SAP Extended Memory" button — copies PROPOSED SAP EM (GB) into
           SAP Extended Memory (GB) for rows that are currently empty. */
        const prefillBtn = el.querySelector('#twa-emconfig-prefill');
        if (prefillBtn) {
            prefillBtn.addEventListener('click', () => {
                let filled = 0;
                hosts.forEach(h => {
                    const key  = h.toLowerCase();
                    const info = hostParam[key] || {};
                    /* Proposed value: same formula as rendered in physCells */
                    if (info.emMb == null && info.physMb != null) {
                        const proposed = Math.round(info.physMb * 0.8 / 1024 * 100) / 100;
                        const gbInput = el.querySelector(`.twa-extmem-input[data-host="${h.replace(/"/g, '\\"')}"]`);
                        if (gbInput && (gbInput.value === '' || gbInput.value == null)) {
                            gbInput.value = proposed;
                            this._extMemory[key] = proposed;
                            filled++;
                        }
                    }
                });
                if (filled > 0) {
                    prefillBtn.textContent = `✓ ${filled} row${filled === 1 ? '' : 's'} filled — values copied from Proposed SAP EM (GB) and remain editable`;
                    setTimeout(() => {
                        prefillBtn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="17 1 21 5 17 9"/><path d="M3 11V9a4 4 0 0 1 4-4h14"/><polyline points="7 23 3 19 7 15"/><path d="M21 13v2a4 4 0 0 1-4 4H3"/></svg> Prefill SAP Extended Memory`;
                    }, 3000);
                } else {
                    prefillBtn.textContent = 'No empty cells to fill';
                    setTimeout(() => {
                        prefillBtn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="17 1 21 5 17 9"/><path d="M3 11V9a4 4 0 0 1 4-4h14"/><polyline points="7 23 3 19 7 15"/><path d="M21 13v2a4 4 0 0 1-4 4H3"/></svg> Prefill SAP Extended Memory`;
                    }, 2000);
                }
            });
        }

        el.querySelector('#twa-emconfig-build').addEventListener('click', () => {
            this._saveExtMemory();
            this._doBuild(app);
        });
        const backBtn = el.querySelector('#twa-emconfig-back');
        if (backBtn) backBtn.addEventListener('click', () => {
            this._saveExtMemory();
            el.style.display = 'none';
            this._renderMappingScreen(app);
        });

        document.getElementById('content-area')?.scrollTo({ top: 0, behavior: 'smooth' });
    },

    /* ══════════════════════════════════════════════════════
       VIRTUAL ↔ PHYSICAL HOSTNAME MAPPING SCREEN
       ══════════════════════════════════════════════════════ */
    _renderMappingScreen(app, preserveScroll = false) {
        const el = document.getElementById('twa-mapping');
        const res = document.getElementById('twa-results');
        if (res) res.style.display = 'none';
        if (!el) return;
        /* Dismiss the optional profiles screen if it was open; retained in this.profiles. */
        this._hidePflViewOnly();

        const scroller = document.getElementById('content-area');

        if (!this._removedHosts) this._removedHosts = new Set();
        /* Physical hosts to show = landscape hosts the user has NOT removed. */
        const landscape = app.getHostVariables().filter(h => !this._removedHosts.has((h.hostname || '').toLowerCase()));
        const esc = s => String(s == null ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

        /* Pool = virtual hosts not yet assigned to any physical host */
        const assignedVirtuals = new Set(Object.keys(this._mapping));
        const poolHosts = this._virtualHosts.filter(v => !assignedVirtuals.has(v.toLowerCase()));

        el.style.display = 'block';
        el.innerHTML = `
            <div class="alert-error" style="align-items:flex-start;margin-top:18px">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24"
                     fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                    <path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
                    <line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
                </svg>
                <div>
                    <strong>Hostname mismatch detected.</strong><br>
                    The server names in the ${this.sourceTable === 'MON' ? '/SDF/MON_HEADER' : '/SDF/SMON_HEADER'}
                    statistics do not match the physical server names in the EWA report. This usually means the
                    performance data uses <strong>virtual hostnames</strong>. Map virtual
                    hostnames to their physical server (from the EWA report) below. DB servers,
                    infrastructure hosts, or standalone hosts that do not appear in the SMON/MON data
                    can be left unmapped — they will be included with blank Cores/SAPS values.
                </div>
            </div>

            <div class="section-title" style="margin-top:20px">Map Virtual → Physical Hostnames</div>
            <div class="map-grid">
                <div class="card" style="padding:0">
                    <div class="card-header" style="padding:12px 16px">
                        <div>
                            <div class="card-title">Physical servers (EWA report)</div>
                            <div class="card-sub">Drag a virtual host onto a row, or pick it from the dropdown</div>
                        </div>
                    </div>
                    <div class="table-wrap" style="border:none;overflow-x:auto;max-height:460px">
                        <table class="data-table">
                            <thead>
                                <tr><th style="width:34px"></th><th>Physical Hostname (EWA)</th><th>Virtual Hostname (SMON/MON)</th><th>Or select</th><th>Or type manually</th></tr>
                            </thead>
                            <tbody id="twa-map-tbody">
                                ${landscape.map(h => this._mapRowHtml(h)).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>

                <div class="card map-pool-card">
                    <div class="card-header" style="padding:12px 16px 8px">
                        <div>
                            <div class="card-title">Virtual hostnames (SMON/MON)</div>
                            <div class="card-sub" id="twa-map-pool-sub">${poolHosts.length} unassigned</div>
                        </div>
                    </div>
                    <div class="map-pool" id="twa-map-pool">${this._mapPoolHtml()}</div>
                </div>
            </div>

            <div class="alert-info" id="twa-map-unmapped" style="margin-top:14px"></div>

            <!-- Removed physical servers (with restore) -->
            <div id="twa-map-removed" style="margin-top:10px"></div>

            <!-- Validation error (shown if the user tries to proceed with unmapped hosts) -->
            <div id="twa-map-error" class="alert-error" style="display:none;margin-top:12px">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24"
                     fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                    <circle cx="12" cy="12" r="10"/>
                    <line x1="15" y1="9" x2="9" y2="15"/>
                    <line x1="9" y1="9" x2="15" y2="15"/>
                </svg>
                <div id="twa-map-error-msg"></div>
            </div>

            <div class="action-row" style="margin-top:18px">
                <button class="btn btn-blue" id="twa-map-confirm">
                    Confirm mapping &amp; process
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2.5">
                        <polyline points="20 6 9 17 4 12"/>
                    </svg>
                </button>
            </div>`;

        this._wireMappingEvents(app);
        this._refreshUnmappedNote();
        this._refreshRemovedNote();
        /* Full render only happens on first entry (or restore) — scroll to top so
           the alert is visible. Assignment clicks use _refreshMappingTable, which
           updates in place and never moves the scroll position. */
        if (!preserveScroll) scroller?.scrollTo({ top: 0, behavior: 'smooth' });
    },

    /* ── Mapping-screen piece builders (shared by full render + in-place refresh) ── */
    _mapEsc(s) {
        return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
    },
    /* Dropdown options: hide virtual hosts assigned to a DIFFERENT physical host. */
    _mapVOptions(selectedVirtual) {
        const esc = s => this._mapEsc(s);
        const sv = (selectedVirtual || '').toLowerCase();
        const opts = [`<option value=""${sv ? '' : ' selected'}>— none —</option>`];
        this._virtualHosts.forEach(v => {
            const vk = v.toLowerCase();
            if ((vk in this._mapping) && vk !== sv) return;
            opts.push(`<option value="${esc(v)}"${sv === vk ? ' selected' : ''}>${esc(v)}</option>`);
        });
        return opts.join('');
    },
    _mapRowHtml(h) {
        const esc = s => this._mapEsc(s);
        const phys = h.hostname || '';
        const current = Object.keys(this._mapping)
            .find(vk => (this._mapping[vk] || '').toLowerCase() === phys.toLowerCase()) || '';
        const currentOrig = this._virtualHosts.find(v => v.toLowerCase() === current) || current;
        /* A "resolved" mapping is one that came from the pool (drag/drop/dropdown).
           A "manual" mapping is one the user typed themselves. */
        const isPoolMapped = !!(currentOrig && this._virtualHosts.find(v => v.toLowerCase() === (currentOrig).toLowerCase()));
        const isManualMapped = !!(current && !this._virtualHosts.find(v => v.toLowerCase() === current));
        const manualVal = isManualMapped ? (this._mapping[current] === phys ? current : '') : '';
        /* Lock the manual input when the row already has a pool-based mapping. */
        const manualDisabled = isPoolMapped ? 'disabled style="width:100%;background:var(--bg-page,#f5f7fa);color:var(--text-muted);cursor:not-allowed"'
                                            : 'style="width:100%"';
        /* Show a "manual" badge in the virtual-host chip for manually-added hosts. */
        const isManualHost = currentOrig && (this._manualVirtualHosts || []).includes(currentOrig.toLowerCase());
        const chipExtra = isManualHost ? ' <span style="font-size:9px;font-weight:600;background:var(--brand-light,#e8f0fe);color:var(--brand,#1a73e8);border-radius:3px;padding:1px 4px;vertical-align:middle">manual</span>' : '';
        return `
            <tr data-phys="${esc(phys)}">
                <td style="padding:4px 0;text-align:center;width:34px">
                    <button class="map-row-del" data-phys="${esc(phys)}"
                            title="Remove this physical server (no counterpart in the SMON/MON data)">×</button>
                </td>
                <td style="font-weight:600;white-space:nowrap">${esc(phys)}</td>
                <td>
                    <div class="map-drop" data-phys="${esc(phys)}">
                        ${currentOrig
                            ? `<span class="map-chip assigned" draggable="true" data-virtual="${esc(currentOrig)}">${esc(currentOrig)}${chipExtra}<span class="map-chip-x" data-virtual="${esc(currentOrig)}">×</span></span>`
                            : `<span class="map-drop-hint">Drop a virtual host here</span>`}
                    </div>
                </td>
                <td>
                    <select class="editable-select map-select" data-phys="${esc(phys)}">
                        ${this._mapVOptions(currentOrig)}
                    </select>
                </td>
                <td style="min-width:160px">
                    <input type="text" class="editable-cell map-manual-input" data-phys="${esc(phys)}"
                           value="${esc(manualVal)}"
                           placeholder="${isPoolMapped ? '— assigned via pool —' : 'or type hostname…'}"
                           title="${isPoolMapped ? 'This row already has a mapping from the pool. Clear it (×) to type manually.' : 'Manually enter a virtual hostname (for DB servers, standalone hosts, or hosts not present in /SDF/MON* or /SDF/SMON*)'}"
                           ${manualDisabled}>
                </td>
            </tr>`;
    },
    _mapPoolHtml() {
        const esc = s => this._mapEsc(s);
        const assigned = new Set(Object.keys(this._mapping));
        const pool = this._virtualHosts.filter(v => !assigned.has(v.toLowerCase()));
        return pool.length
            ? pool.map(v => `<span class="map-chip" draggable="true" data-virtual="${esc(v)}">${esc(v)}</span>`).join('')
            : `<span class="map-pool-empty">All virtual hosts have been assigned.</span>`;
    },

    /* Update ONLY the table rows, the pool and the notes — without replacing the
       whole mapping screen, so the scroll position is never disturbed (important
       for landscapes with many servers). Re-wires the events on the new nodes. */
    _refreshMappingTable(app) {
        const tbody = document.getElementById('twa-map-tbody');
        const pool  = document.getElementById('twa-map-pool');
        const poolSub = document.getElementById('twa-map-pool-sub');
        if (!tbody || !pool) return;
        const landscape = app.getHostVariables().filter(h => !this._removedHosts.has((h.hostname || '').toLowerCase()));
        tbody.innerHTML = landscape.map(h => this._mapRowHtml(h)).join('');
        pool.innerHTML = this._mapPoolHtml();
        if (poolSub) {
            const assigned = new Set(Object.keys(this._mapping));
            poolSub.textContent = `${this._virtualHosts.filter(v => !assigned.has(v.toLowerCase())).length} unassigned`;
        }
        this._wireMappingEvents(app);     /* re-bind listeners on the fresh nodes */
        this._refreshUnmappedNote();
        this._refreshRemovedNote();
    },

    /* Attach drag-and-drop + dropdown + confirm handlers to the mapping screen. */
    _wireMappingEvents(app) {
        const el = document.getElementById('twa-mapping');
        if (!el) return;

        /* Dragging a chip (from the pool or an assigned row) */
        el.querySelectorAll('.map-chip').forEach(chip => {
            chip.addEventListener('dragstart', e => {
                e.dataTransfer.setData('text/plain', chip.dataset.virtual);
                e.dataTransfer.effectAllowed = 'move';
                chip.classList.add('dragging');
            });
            chip.addEventListener('dragend', () => chip.classList.remove('dragging'));
        });

        /* Chip "×" returns the virtual host to the pool */
        el.querySelectorAll('.map-chip-x').forEach(x => {
            x.addEventListener('click', e => {
                e.stopPropagation();
                this._assign(x.dataset.virtual, '');   // unassign
                this._refreshMappingTable(app);
            });
        });

        /* Drop targets (physical rows) */
        el.querySelectorAll('.map-drop').forEach(drop => {
            drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('drag-over'); });
            drop.addEventListener('dragleave', () => drop.classList.remove('drag-over'));
            drop.addEventListener('drop', e => {
                e.preventDefault();
                drop.classList.remove('drag-over');
                const virtual = e.dataTransfer.getData('text/plain');
                if (virtual) {
                    this._assign(virtual, drop.dataset.phys);
                    this._refreshMappingTable(app);
                }
            });
        });

        /* Dropdown fallback */
        el.querySelectorAll('.map-select').forEach(sel => {
            sel.addEventListener('change', () => {
                const phys = sel.dataset.phys;
                const virtual = sel.value;
                /* Clear any virtual previously mapped to this physical, then set. */
                Object.keys(this._mapping).forEach(vk => {
                    if ((this._mapping[vk] || '').toLowerCase() === phys.toLowerCase()) delete this._mapping[vk];
                });
                if (virtual) this._assign(virtual, phys);
                this._refreshMappingTable(app);
            });
        });

        /* Allow dropping a chip back onto the pool to unassign it */
        const pool = document.getElementById('twa-map-pool');
        if (pool) {
            pool.addEventListener('dragover', e => { e.preventDefault(); pool.classList.add('drag-over'); });
            pool.addEventListener('dragleave', () => pool.classList.remove('drag-over'));
            pool.addEventListener('drop', e => {
                e.preventDefault();
                pool.classList.remove('drag-over');
                const virtual = e.dataTransfer.getData('text/plain');
                if (virtual) { this._assign(virtual, ''); this._refreshMappingTable(app); }
            });
        }

        /* Delete button — remove a physical server row that has no counterpart
           in the SMON/MON data (e.g. a DB server when the stats cover only app
           servers). Any virtual host mapped to it is returned to the pool. */
        el.querySelectorAll('.map-row-del').forEach(btn => {
            btn.addEventListener('click', e => {
                e.stopPropagation();
                const phys = btn.dataset.phys || '';
                this._removedHosts.add(phys.toLowerCase());
                /* Unassign any virtual host that was mapped to this physical. */
                Object.keys(this._mapping).forEach(vk => {
                    if ((this._mapping[vk] || '').toLowerCase() === phys.toLowerCase()) delete this._mapping[vk];
                });
                this._refreshMappingTable(app);
            });
        });

        /* Restore a previously-removed physical server. */
        el.querySelectorAll('.map-restore').forEach(btn => {
            btn.addEventListener('click', e => {
                e.stopPropagation();
                this._removedHosts.delete((btn.dataset.phys || '').toLowerCase());
                this._refreshMappingTable(app);
            });
        });

        /* Manual text input — lets users type a virtual hostname for physical hosts
           that are not present in /SDF/MON* or /SDF/SMON* (DB servers, standalone
           hosts, infrastructure components, etc.).
           On commit (blur or Enter): add the typed name to _virtualHosts if new,
           record it in _manualVirtualHosts so the UI can mark it as manually-added,
           assign it to the physical host, and refresh the table. */
        el.querySelectorAll('.map-manual-input').forEach(input => {
            if (input.disabled) return;   /* skip locked rows (already pool-mapped) */
            const commit = () => {
                const phys    = input.dataset.phys || '';
                const virtual = (input.value || '').trim();
                if (!virtual) return;
                const vkLower = virtual.toLowerCase();
                if (!this._virtualHosts.find(v => v.toLowerCase() === vkLower)) {
                    this._virtualHosts = [...this._virtualHosts, virtual];
                }
                /* Track this as a manually-added virtual hostname. */
                if (!this._manualVirtualHosts) this._manualVirtualHosts = [];
                if (!this._manualVirtualHosts.includes(vkLower)) {
                    this._manualVirtualHosts = [...this._manualVirtualHosts, vkLower];
                }
                Object.keys(this._mapping).forEach(vk => {
                    if ((this._mapping[vk] || '').toLowerCase() === phys.toLowerCase()) delete this._mapping[vk];
                });
                this._assign(virtual, phys);
                this._refreshMappingTable(app);
            };
            input.addEventListener('blur', commit);
            input.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); commit(); } });
        });

        /* Wire the confirm button ONCE. It lives outside the tbody/pool region
           that _refreshMappingTable replaces, so the element persists across
           re-renders; without this guard every refresh would stack another
           click listener, firing _runBuildWithMapping multiple times per click. */
        const confirmBtn = document.getElementById('twa-map-confirm');
        if (confirmBtn && !confirmBtn._wiredConfirm) {
            confirmBtn._wiredConfirm = true;
            confirmBtn.addEventListener('click', () => this._runBuildWithMapping(app));
        }
    },

    /* Set/clear a mapping entry. virtual→physical ('' physical = unassign).
       A virtual host maps to at most one physical; assigning it elsewhere moves it.
       When unassigning a manually-added virtual host, also remove it from the pool
       (it was added by the user, not discovered from the source file). */
    _assign(virtual, physical) {
        if (!virtual) return;
        const vk = virtual.toLowerCase();
        if (!physical) {
            delete this._mapping[vk];
            /* If this was a manually-added host, purge it entirely from the pool
               so it doesn't linger as an unassigned chip the user can't delete. */
            if (this._manualVirtualHosts && this._manualVirtualHosts.includes(vk)) {
                this._virtualHosts = this._virtualHosts.filter(v => v.toLowerCase() !== vk);
                this._manualVirtualHosts = this._manualVirtualHosts.filter(m => m !== vk);
            }
            return;
        }
        this._mapping[vk] = physical;
    },

    _refreshUnmappedNote() {
        const note = document.getElementById('twa-map-unmapped');
        if (!note) return;
        const assigned = new Set(Object.keys(this._mapping));
        const unmapped = this._virtualHosts.filter(v => !assigned.has(v.toLowerCase()));
        if (!unmapped.length) {
            /* All mapped — clear any prior validation error and show the go-ahead. */
            this._clearMappingError();
            note.innerHTML = `<span>All virtual hosts are mapped. Click <strong>Confirm mapping &amp; process</strong> to continue.</span>`;
            return;
        }
        note.innerHTML =
            `<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                  fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
             </svg>
             <span><strong>${unmapped.length} virtual host(s) are not yet mapped</strong>: ${unmapped.map(u => `<code>${u}</code>`).join(', ')}.
             You can map them now, or click <strong>Confirm mapping &amp; process</strong> to proceed
             — unmapped hosts will appear in the enriched table with blank Cores/SAPS values.</span>`;
    },

    /* Validation-error banner on the mapping screen (unmapped hosts on confirm). */
    _showMappingError(msg) {
        const wrap = document.getElementById('twa-map-error');
        const txt  = document.getElementById('twa-map-error-msg');
        if (txt)  txt.textContent = msg;
        if (wrap) { wrap.style.display = 'flex'; wrap.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); }
    },
    _clearMappingError() {
        const wrap = document.getElementById('twa-map-error');
        if (wrap) wrap.style.display = 'none';
    },

    /* List of removed physical servers, each with a Restore link. */
    _refreshRemovedNote() {
        const note = document.getElementById('twa-map-removed');
        if (!note) return;
        const removed = [...(this._removedHosts || [])];
        if (!removed.length) { note.innerHTML = ''; return; }
        const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
        /* Use the original-case hostname from the landscape for display. */
        const hv = (this.app && this.app.getHostVariables()) || [];
        const disp = key => (hv.find(h => (h.hostname || '').toLowerCase() === key) || {}).hostname || key;
        note.className = 'alert-info';
        note.style.marginTop = '10px';
        note.innerHTML =
            `<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                  fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
             </svg>
             <span><strong>${removed.length} physical server(s) removed</strong> (excluded from the analysis): ` +
             removed.map(k => `<code>${esc(disp(k))}</code> <a href="#" class="map-restore" data-phys="${esc(k)}" style="color:var(--brand);text-decoration:underline">restore</a>`).join(', ') +
             `.</span>`;
    },

    /* ── Field LABEL → canonical field NAME alias table ───────────────────
       SAP exports from /SDF/SMON_HEADER and /SDF/MON_HEADER can carry either the
       technical field NAME (DATUM, SERVER, IDLE_TOTAL, NETRTT, …) or the field
       LABEL / description ("Date of Monitoring Snapshot", "Application Server
       Instance", "CPU Utilization (Idle)", "Network RTT (E. Acc.DB)", …).
       The entire downstream pipeline keys off field names, so when a file uses
       labels we translate each header back to its field name up front. Keys are
       normalised (lower-case, single-spaced); the first field name wins when a
       label is ambiguous (e.g. NETRTT vs NETRTT_FLAG share a label, "Queue"
       repeats) which matches the left-to-right column order of the exports. */
    _FIELD_LABEL_ALIASES: {
        /* ── GUID ── */
        'guid of snapshot analysis': 'GUID',             /* SMON */
        'monitoring guid': 'ANALYSIS_GUID',              /* MON  */
        /* ── date / time ── */
        'date of monitoring snapshot': 'DATUM',          /* SMON */
        'date': 'DATUM',                                 /* MON  */
        'time of monitoring snapshot': 'TIME',           /* SMON */
        'time': 'TIME',                                  /* MON  */
        /* ── server ── */
        'application server instance': 'SERVER',         /* SMON */
        'as instance': 'SERVER',                         /* MON  */
        /* ── work processes ── */
        'act. wps': 'ACT_WPS',                           /* SMON */
        'active wps': 'ACT_WPS',                         /* MON  */
        'dia.wps': 'ACT_DIA',                            /* SMON */
        'active dialog wps': 'ACT_DIA',                  /* MON  */
        'number of free dia wps for rfc with norm': 'TRFC_FREE',  /* SMON */
        'prio normal rfc wps': 'TRFC_FREE',              /* MON  */
        'number of free dia wps for rfc with low': 'RFC_LOW',     /* SMON */
        /* ── CPU ── */
        'cpu utilization (user)': 'USR_TOTAL',           /* SMON */
        'cpu user': 'USR_TOTAL',                         /* MON  */
        'cpu utilization (system)': 'SYS_TOTAL',         /* SMON */
        'cpu system': 'SYS_TOTAL',                       /* MON  */
        'cpu utilization (idle)': 'IDLE_TOTAL',          /* SMON */
        'cpu idle': 'IDLE_TOTAL',                        /* MON  */
        'cpus consumed': 'CPU_CONS',                     /* SMON */
        'available cpus': 'AVAILCPUS',                   /* SMON + MON */
        'ready time in %': 'READY_TIME',                 /* SMON */
        'steal time in seconds': 'STEAL_TIME',           /* SMON */
        /* ── paging / memory ── */
        'paging in activity (in % of ram per hour)': 'PAGE_IN_PERC',  /* SMON */
        'paging in acitvity (in % of ram per hour': 'PAGE_IN_PERC',   /* SMON (older typo'd/truncated label) */
        'paging out': 'PAGE_OUT_PERC',                   /* SMON */
        'free mem.': 'FREE_MEM_PERC',                    /* SMON */
        'paging in (kb/s)': 'IN_KB_SEC',                 /* MON  */
        'paging out (kb/s)': 'OUT_KB_SEC',               /* MON  */
        'free memory in kb': 'FREE_MEM',                 /* MON  */
        'em alloc.': 'EMALLOC',                          /* SMON */
        'ext. mem. allocated': 'EMALLOC',                /* MON  */
        'em attach.': 'EMATTACH',                        /* SMON */
        'ext. mem. attached': 'EMATTACH',                /* MON  */
        'global ext. mem. in kb': 'EMGLOBALKB',          /* MON  */
        'heap memory in kb': 'HEAPSUMKB',                /* SMON + MON */
        'paging mem': 'PAGESHMKB',                       /* SMON */
        'paging shared memory in kb': 'PAGESHMKB',       /* MON  */
        'roll shared memory in kb': 'ROLLSHMKB',         /* MON  */
        'priv modes': 'PRIVWPNO',                        /* SMON + MON */
        'freemem': 'FREE_MEM_MB',                        /* SMON */
        'free(+fs)': 'FREE_MEM_MB_INC_FS',               /* SMON */
        /* ── queues ── */
        'dialog queue length': 'DIAQ',                   /* SMON + MON */
        'average load last 20s': 'DIAAVG20',             /* SMON */
        'average load last 60s': 'DIAAVG60',             /* SMON */
        'update queue length': 'UPDQ',                   /* SMON + MON */
        'enqueue queue length': 'ENQQ',                  /* SMON + MON */
        'queue (inbound)': 'QINLENGTH',                  /* SMON + MON */
        'inbound queue statistics': 'QIN_FLAG',          /* SMON + MON */
        'queue (outbound)': 'QOUTLENGTH',                /* SMON + MON */
        'outbound queue statistics': 'QOUT_FLAG',        /* SMON + MON */
        /* ── status / measurement flags ── */
        'work processes': 'SM50_FLAG',                   /* SMON + MON */
        'queue measured': 'QUEUE_FLAG',                  /* SMON + MON */
        'st06 measured': 'ST06_FLAG',                    /* SMON + MON */
        'top cpu processes': 'TOPCPU_FLAG',              /* SMON + MON */
        'memory consumption measured': 'ST02_FLAG',      /* SMON + MON */
        'memory per modes measured': 'SM04_FLAG',        /* SMON + MON */
        'number of enqueue entries': 'SM12_CNT',         /* SMON + MON */
        'enqueue entries': 'SM12_FLAG',                  /* SMON + MON */
        'enqueue statistics measured': 'SM12_STAT_FLAG', /* SMON + MON */
        /* ── logins / sessions / timing ── */
        'number of logins': 'USERS',                     /* SMON + MON */
        'number of sessions': 'SESSIONS',                /* SMON + MON */
        'time since last measurement in ms': 'DELTATIMEMS', /* SMON */
        /* ── network ── */
        'network rtt (e. acc.db)': 'NETRTT',             /* SMON */
        'network rtt flag (e. acc.db)': 'NETRTT_FLAG',   /* SMON */
    },

    /* Translate a raw header to its canonical SAP field name when it is a known
       field label; otherwise return the header unchanged (already a field name,
       or an unmapped column we carry verbatim). */
    _canonicalFieldName(header) {
        const norm = String(header == null ? '' : header).trim().toLowerCase().replace(/\s+/g, ' ');
        return this._FIELD_LABEL_ALIASES[norm] || header;
    },

    /* ── Detect SMON vs MON and locate the header row ──
       Returns { source:'SMON'|'MON', headerIdx, headers:[] } or null.
       Strategy: find the row that contains the column-name tokens. SMON is
       distinguished by SMON-only columns (CPU_CONS / AVAILCPUS / FREE_MEM_MB),
       MON by MON-only columns (IN_KB_SEC / EMGLOBALKB / a bare FREE_MEM).      */
    _detectSource(rows) {
        /* ── Pre-processed spreadsheet check (runs first) ──
           A file already prepared by the user uses human-friendly headers and
           already carries all derived values. Its signature: a header row with a
           server-column alias + a "CPU Utilization" column + a "SAPS Utilization"
           or "EM Configured" column — tokens the raw SAP exports never have.
           Matching is tolerant (substring on normalised headers) so minor header
           variations ("CPU Utilization (%)", "SAPS Utilization (SAPS)", trailing
           spaces, etc.) are still recognised. */
        const norm = s => String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
        const SERVER_ALIASES = ['server name', 'application server', 'hostname', 'host'];
        for (let i = 0; i < Math.min(rows.length, 60); i++) {
            const cells = (rows[i] || []).map(norm).filter(Boolean);
            if (!cells.length) continue;
            const set = new Set(cells);
            const anyHas = sub => cells.some(c => c.indexOf(sub) >= 0);
            const hasServer = SERVER_ALIASES.some(a => set.has(a)) || anyHas('server name') || anyHas('application server');
            const hasCpu    = anyHas('cpu utilization');
            const hasPp     = anyHas('saps utilization') || anyHas('em configured') || anyHas('saps deployed');
            if (hasServer && hasCpu && hasPp) {
                return { source: 'PREPARSED', headerIdx: i, headers: rows[i].map(c => String(c).trim()) };
            }
        }

        /* Explicit table-name marker — only the first rows carry the banner, so
           scan a bounded slice (avoids building a huge string on large files). */
        const head = rows.slice(0, 60).map(r => r.join(' ')).join('\n').toUpperCase();
        let marker = '';
        if (head.indexOf('/SDF/SMON_HEADER') >= 0) marker = 'SMON';
        else if (head.indexOf('/SDF/MON_HEADER') >= 0) marker = 'MON';

        /* Find a header row: the one containing SERVER plus several known tokens.
           Each cell is first translated through the field-label alias table so a
           file that uses field LABELS ("Application Server Instance", "CPU
           Utilization (Idle)", …) is recognised exactly like one using field
           NAMES. We keep the ORIGINAL header strings in `headers` — _mapColumns
           canonicalises them to field names when building the index. */
        let headerIdx = -1, headers = null;
        for (let i = 0; i < rows.length; i++) {
            const canon = rows[i].map(c => this._canonicalFieldName(String(c).trim()).toUpperCase());
            const has = tok => canon.some(c => c === tok || c.replace(/\s+/g, '') === tok);
            if (has('SERVER') && (has('ACT_WPS') || has('IDLE_TOTAL') || has('USR_TOTAL'))) {
                headerIdx = i;
                headers = rows[i].map(c => c.trim());
                break;
            }
        }
        if (headerIdx < 0) {
            /* No recognisable header row */
            return null;
        }

        /* Decide source from the header columns if no explicit marker. Compare on
           canonical field names so label-format files score correctly. */
        const up = headers.map(h => this._canonicalFieldName(h).toUpperCase());
        const hasCol = name => up.includes(name);
        let source = marker;
        if (!source) {
            const smonSignals = ['CPU_CONS', 'AVAILCPUS', 'FREE_MEM_MB', 'FREE_MEM_MB_INC_FS', 'PAGESHMKB', 'SESSIONS'];
            const monSignals  = ['IN_KB_SEC', 'OUT_KB_SEC', 'EMGLOBALKB'];
            const smonScore = smonSignals.filter(hasCol).length;
            const monScore  = monSignals.filter(hasCol).length;
            if (monScore > smonScore) source = 'MON';
            else if (smonScore > 0)   source = 'SMON';
            else source = hasCol('FREE_MEM') && !hasCol('FREE_MEM_MB') ? 'MON' : 'SMON';
        }
        return { source, headerIdx, headers };
    },

    /* ── Build a name→index lookup, tolerant of spacing/case ──
       Headers that are field LABELS are translated to their canonical field NAME
       first, so a file using labels indexes identically to one using names. When
       a label is ambiguous (two columns share a label, e.g. NETRTT / NETRTT_FLAG)
       the FIRST occurrence wins — matching the left-to-right export order. */
    _mapColumns(headers) {
        const col = {};
        headers.forEach((h, i) => {
            const canon = this._canonicalFieldName(h);
            const key = canon.toUpperCase().replace(/\s+/g, '');
            if (key && col[key] == null) col[key] = i;
        });

        /* Safety net for the SERVER column: it is mandatory (data rows are keyed
           off it) and label spellings vary across SAP releases / export widths.
           If the exact-alias pass missed it, fall back to a fuzzy match on the
           raw headers so an unseen label variant ("Application Server Instance",
           "AS Instance", "Server Name", …) still resolves. */
        if (col.SERVER == null) {
            const norm = s => String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
            let hit = -1;
            headers.forEach((h, i) => {
                if (hit >= 0) return;
                const n = norm(h);
                if (n === 'server' || n === 'as instance' ||
                    n.indexOf('server instance') >= 0 ||
                    n.indexOf('application server') >= 0 ||
                    n === 'server name' || n === 'instance') {
                    hit = i;
                }
            });
            if (hit >= 0) col.SERVER = hit;
        }

        return col;
    },

    /* ══════════════════════════════════════════════════════
       PRE-PROCESSED SPREADSHEET SUPPORT
       A file already prepared by the user, carrying human-friendly columns and
       all derived values. We validate the required columns, then copy them into
       the same enriched target shape the SMON/MON builder produces.
       ══════════════════════════════════════════════════════ */

    /* Server-column aliases accepted in a pre-processed file (normalised). */
    _PP_SERVER_ALIASES: ['server name', 'application server', 'hostname', 'host'],

    /* Required columns (display name → list of accepted normalised header spellings).
       Order defines the message order when reporting what is missing. */
    _preparsedRequired() {
        return [
            { label: 'Date',                 keys: ['date'] },
            { label: 'Time',                 keys: ['time'] },
            { label: 'Hour',                 keys: ['hour', 'hours'] },
            { label: 'Server Name',          keys: this._PP_SERVER_ALIASES },
            { label: 'Active WP',            keys: ['active wp'] },
            { label: 'Active Dialog WP',     keys: ['active dialog wp'] },
            { label: 'WP free for RFC',      keys: ['wp free for rfc'] },
            { label: 'CPU Utilization (%)',  keys: ['cpu utilization (%)'] },
            { label: 'Paging in',            keys: ['paging in'] },
            { label: 'Paging out',           keys: ['paging out'] },
            { label: 'Physical Memory (GB)', keys: ['physical memory (gb)'] },
            { label: 'Free Memory (GB)',     keys: ['free memory (gb)'] },
            { label: 'EM Configured (GB)',   keys: ['em configured (gb)'] },
            { label: 'EM Allocated (GB)',    keys: ['em allocated (gb)'] },
            { label: 'Heap Memory (GB)',     keys: ['heap memory (gb)'] },
            { label: 'Logins',               keys: ['logins'] },
            { label: 'Sessions',             keys: ['sessions'] },
            { label: 'Cores',                keys: ['cores'] },
            { label: 'SAPS/Core',            keys: ['saps/core', 'saps / core', 'saps/core (saps)'] },
            { label: 'SAPS Deployed',        keys: ['saps deployed', 'saps deployed (saps)'] },
            { label: 'SAPS Utilization',     keys: ['saps utilization', 'saps utilization (saps)'] },
        ];
    },

    /* Return the display names of required columns missing from `headers`. */
    _validatePreparsed(headers) {
        const norm = s => String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
        const present = new Set(headers.map(norm));
        return this._preparsedRequired()
            .filter(req => !req.keys.some(k => present.has(k)))
            .map(req => req.label);
    },

    /* Build the enriched target table directly from a pre-processed file.
       Copies values through the column mapping (no SAP derivation) and returns
       { cols, rows } in the same shape as _buildTarget, setting this._colIndex. */
    _buildTargetFromPreparsed(headers, dataRows, app) {
        const norm = s => String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ');
        /* normalised header → source index */
        const srcIdx = {};
        headers.forEach((h, i) => { const k = norm(h); if (k && !(k in srcIdx)) srcIdx[k] = i; });
        const firstIdx = keys => { for (const k of keys) if (k in srcIdx) return srcIdx[k]; return -1; };

        /* Output column → source index (via accepted header spellings). */
        const map = [
            { out: 'Date',                    keys: ['date'] },
            { out: 'TIME',                    keys: ['time'] },
            { out: 'Hours',                   keys: ['hour', 'hours'] },
            { out: 'SERVER',                  keys: this._PP_SERVER_ALIASES },
            { out: 'ACT_WPS',                 keys: ['active wp'] },
            { out: 'ACT_DIA',                 keys: ['active dialog wp'] },
            { out: 'TRFC_FREE',               keys: ['wp free for rfc'] },
            { out: 'CPU Utilization (%)',     keys: ['cpu utilization (%)'] },
            { out: 'Paging in',               keys: ['paging in'] },
            { out: 'Paging out',              keys: ['paging out'] },
            { out: 'Physical Memory (GB)',    keys: ['physical memory (gb)'] },
            { out: 'EM Configured (GB)',      keys: ['em configured (gb)'] },
            { out: 'EM Allocated (GB)',       keys: ['em allocated (gb)'] },
            { out: 'Heap Memory (GB)',        keys: ['heap memory (gb)'] },
            { out: 'Free Memory (GB)',        keys: ['free memory (gb)'] },
            { out: 'USERS',                   keys: ['logins'] },
            { out: 'SESSIONS',                keys: ['sessions'] },
            { out: 'Cores',                   keys: ['cores'] },
            { out: 'SAPS/Core (SAPS)',        keys: ['saps/core', 'saps / core', 'saps/core (saps)'] },
            { out: 'SAPS Deployed (SAPS)',    keys: ['saps deployed', 'saps deployed (saps)'] },
            { out: 'SAPS Utilization (SAPS)', keys: ['saps utilization', 'saps utilization (saps)'] },
        ];

        const outCols = map.map(m => m.out);
        const colPlan = map.map(m => ({ out: m.out, idx: firstIdx(m.keys) }));
        const ci = {};
        outCols.forEach((c, i) => { ci[c] = i; });
        const nCols = outCols.length;

        const oiDate   = ci['Date'],  oiTime = ci['TIME'], oiHours = ci['Hours'], oiSrv = ci['SERVER'];
        const oiCpu    = ci['CPU Utilization (%)'];
        const oiSU     = ci['SAPS Utilization (SAPS)'], oiSD = ci['SAPS Deployed (SAPS)'];
        const siDate   = firstIdx(['date']);
        const siTime   = firstIdx(['time']);

        /* Numeric output columns we coerce via _num (clean thousands separators). */
        const numericOuts = new Set([
            'ACT_WPS','ACT_DIA','TRFC_FREE','CPU Utilization (%)','Paging in','Paging out',
            'Physical Memory (GB)','EM Configured (GB)','EM Allocated (GB)','Heap Memory (GB)',
            'Free Memory (GB)','USERS','SESSIONS','Cores','SAPS/Core (SAPS)',
            'SAPS Deployed (SAPS)','SAPS Utilization (SAPS)',
        ]);
        const numericIdx = colPlan
            .filter(p => numericOuts.has(p.out) && p.idx >= 0)
            .map(p => ({ oi: ci[p.out], si: p.idx }));

        const hostSet = new Set();
        let cpuSum = 0, cpuCnt = 0, cpuMax = -Infinity;
        const cpuSample = [], sapsSample = [];
        const RES_CAP = 50000;
        const sortKeys = new Array(dataRows.length);
        const rowsOut = new Array(dataRows.length);

        /* De-duplicate records repeated across multiple uploaded spreadsheets:
           a row is a duplicate when DATUM + TIME + SERVER all match one already
           written. Keyed on the raw source cells (SERVER case-insensitive); the
           first occurrence wins. Only applied when all three source columns exist. */
        const siSrv = (oiSrv != null && colPlan[oiSrv] && colPlan[oiSrv].idx >= 0) ? colPlan[oiSrv].idx : -1;
        const canDedup = siDate >= 0 && siTime >= 0 && siSrv >= 0;
        const seenKeys = new Set();
        let duplicates = 0;

        for (let ri = 0; ri < dataRows.length; ri++) {
            const r = dataRows[ri];

            if (canDedup) {
                const dupKey = String(r[siDate] ?? '').trim() + '\u0001' +
                               String(r[siTime] ?? '').trim() + '\u0001' +
                               String(r[siSrv]  ?? '').trim().toLowerCase();
                if (seenKeys.has(dupKey)) {
                    duplicates++;
                    dataRows[ri] = null;
                    continue;
                }
                seenKeys.add(dupKey);
            }

            const o = new Array(nCols);

            /* Straight copies first (colPlan is in outCols order, so index = p). */
            for (let p = 0; p < colPlan.length; p++) {
                const pl = colPlan[p];
                o[p] = pl.idx >= 0 ? (r[pl.idx] != null ? r[pl.idx] : '') : '';
            }
            /* Numeric cleanup (overwrite the straight copy for numeric columns).
               Keep decimals — these are already-final values (CPU %, GB, SAPS),
               so e.g. 45.3 or 128.5 must not be truncated to an integer. */
            for (let p = 0; p < numericIdx.length; p++) {
                const { oi, si } = numericIdx[p];
                const n = this._num(r[si], true);
                o[oi] = isNaN(n) ? (r[si] != null ? r[si] : '') : n;
            }

            /* Normalise Date + SERVER, keep Hours/TIME as-is. */
            if (oiDate != null && siDate >= 0) o[oiDate] = this._reformatDate(r[siDate]);
            const hostname = this._hostOf(oiSrv != null && colPlan[oiSrv].idx >= 0 ? (r[colPlan[oiSrv].idx] || '') : '');
            if (oiSrv != null) o[oiSrv] = hostname;
            if (hostname) hostSet.add(hostname);

            sortKeys[ri] = this._dateTimeKey(siDate >= 0 ? r[siDate] : '', siTime >= 0 ? r[siTime] : '');

            /* Summary accumulators. */
            const cpu = parseFloat(o[oiCpu]);
            if (!isNaN(cpu)) {
                cpuSum += cpu; cpuCnt++;
                if (cpu > cpuMax) cpuMax = cpu;
                if (cpuSample.length < RES_CAP) cpuSample.push(cpu);
            }
            const su = parseFloat(o[oiSU]);
            if (!isNaN(su) && sapsSample.length < RES_CAP) sapsSample.push(su);

            dataRows[ri] = null;
            rowsOut[ri] = o;
        }

        /* Chronological sort (same approach as _buildTarget). Compact out the
           holes left by de-duplicated rows first. */
        const order = [];
        for (let i = 0; i < rowsOut.length; i++) {
            if (rowsOut[i] !== undefined) order.push(i);
        }
        order.sort((a, b) => {
            const ka = sortKeys[a], kb = sortKeys[b];
            if (ka === kb) return a - b;
            if (ka == null) return 1;
            if (kb == null) return -1;
            return ka - kb;
        });
        const sorted = new Array(order.length);
        for (let i = 0; i < order.length; i++) sorted[i] = rowsOut[order[i]];
        rowsOut.length = 0;
        for (let i = 0; i < sorted.length; i++) rowsOut[i] = sorted[i];
        sorted.length = 0;

        if (duplicates > 0) {
            try { console.log(`[TechWorkloadApp] skipped ${duplicates} duplicate record(s) (same DATUM+TIME+SERVER).`); } catch (_) {}
        }
        this._duplicatesSkipped = duplicates;

        this._matchStats = { matched: rowsOut.length, unmatched: 0, unmatchedHosts: [] };
        this._summary = {
            dataRows: rowsOut.length,
            hosts: hostSet.size,
            cpuAvg: cpuCnt ? cpuSum / cpuCnt : null,
            cpuMax: cpuMax === -Infinity ? null : cpuMax,
            cpuP95: cpuSample.length ? percentile(cpuSample, 95) : null,
            sapsP95: sapsSample.length ? percentile(sapsSample, 95) : null,
        };
        this._colIndex = ci;
        return { cols: outCols, rows: rowsOut };
    },

    /* ══════════════════════════════════════════════════════
       NUMBER / DATE / SERVER NORMALISERS
       ══════════════════════════════════════════════════════ */

    /* Normalise a SAP-formatted number string using the file's detected
       decimal separator (see _detectDecimalSep).

       Separator semantics once the decimal char is known:
         • decimal = ','  → '.' is thousands, ',' is the decimal point
             "262.026,140" → 262026.140 ; "11,445" → 11.445
         • decimal = '.'  → ',' is thousands, '.' is the decimal point
             "25.392"      → 25.392      ; "11,445" → 11445
       The spec's cleanup columns keep the INTEGER part only (keepDecimals=false),
       so e.g. "262142,044" → 262142 under a comma-decimal locale.

       When the per-file decimal separator is unknown we fall back to the
       heuristic: a lone separator with 1-2 trailing digits is decimal, 3 is
       thousands. Returns a Number or NaN. */
    _num(raw, keepDecimals = true) {
        if (raw == null) return NaN;
        let s = String(raw).trim();
        if (s === '' || s === '-') return NaN;
        s = s.replace(/\s/g, '');

        const dec = this._decimalSep;   /* '.' | ',' | '' (unknown) */
        const hasDot = s.includes('.');
        const hasComma = s.includes(',');

        if (dec === ',') {
            s = s.replace(/\./g, '');          /* drop thousands dots */
            s = s.replace(',', '.');           /* comma → decimal point */
        } else if (dec === '.') {
            s = s.replace(/,/g, '');           /* drop thousands commas */
        } else {
            /* Unknown locale — resolve each value on its own merits. */
            if (hasDot && hasComma) {
                if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
                else                                         s = s.replace(/,/g, '');
            } else if (hasComma) {
                const p = s.split(',');
                if (p.length === 2 && p[1].length > 0 && p[1].length <= 2) s = p[0] + '.' + p[1];
                else s = s.replace(/,/g, '');
            } else if (hasDot) {
                const p = s.split('.');
                if (!(p.length === 2 && p[1].length > 0 && p[1].length <= 2)) s = s.replace(/\./g, '');
            }
        }

        let n = parseFloat(s);
        if (isNaN(n)) return NaN;
        if (!keepDecimals) n = Math.trunc(n);
        return n;
    },

    /* Detect the file-wide decimal separator by scanning numeric-looking cells.
       A value with BOTH '.' and ',' pins the decimal to whichever comes last.
       Returns '.', ',' or '' (undetermined). */
    _detectDecimalSep(dataRows, col) {
        const sampleCols = ['CPU_CONS', 'USR_TOTAL', 'IDLE_TOTAL', 'EMALLOC', 'EMATTACH',
                            'HEAPSUMKB', 'FREE_MEM_MB', 'FREE_MEM', 'FREE_MEM_MB_INC_FS',
                            'AVAILCPUS', 'READY_TIME'];
        let comma = 0, dot = 0;
        for (const r of dataRows) {
            for (const name of sampleCols) {
                if (col[name] == null) continue;
                const v = String(r[col[name]] ?? '').trim();
                if (!v) continue;
                const hasDot = v.includes('.'), hasComma = v.includes(',');
                if (hasDot && hasComma) {
                    if (v.lastIndexOf(',') > v.lastIndexOf('.')) comma += 3; else dot += 3;
                } else if (hasComma) {
                    /* short trailing group ⇒ likely decimal */
                    const p = v.split(',');
                    if (p.length === 2 && p[1].length <= 2) comma += 1;
                } else if (hasDot) {
                    const p = v.split('.');
                    if (p.length === 2 && p[1].length <= 2) dot += 1;
                }
            }
        }
        if (comma > dot) return ',';
        if (dot > comma) return '.';
        return '';
    },

    /* DD.MM.YYYY → DD/MM/YYYY */
    _reformatDate(raw) {
        const s = String(raw || '').trim();
        const m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
        if (m) return `${m[1]}/${m[2]}/${m[3]}`;
        return s;
    },

    /* "HH:MM:SS" → "HH" (24h, two digits) */
    _hourOf(raw) {
        const s = String(raw || '').trim();
        const m = s.match(/^(\d{1,2})[:.]/);
        if (m) return String(m[1]).padStart(2, '0');
        const m2 = s.match(/^(\d{1,2})$/);
        if (m2) return String(m2[1]).padStart(2, '0');
        return '';
    },

    /* "hostname_SID_nn" → "hostname" (keep text before the first underscore) */
    _hostOf(raw) {
        const s = String(raw || '').trim();
        if (!s) return '';
        return s.split('_')[0];
    },

    /* Build a sortable numeric key from the raw DATUM + TIME.
       DATUM is "DD.MM.YYYY" (SAP) or already "MM/DD/YYYY" / "DD/MM/YYYY" in some
       exports; TIME is "HH:MM:SS". Returns YYYYMMDDHHMMSS as a number, or null
       if the date can't be parsed (those rows sort to the end). */
    _dateTimeKey(datum, time) {
        const d = String(datum || '').trim();
        let y, mo, da;
        let m = d.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);        /* DD.MM.YYYY */
        if (m) { da = +m[1]; mo = +m[2]; y = +m[3]; }
        else if ((m = d.match(/^(\d{4})-(\d{2})-(\d{2})$/)))   /* YYYY-MM-DD */
            { y = +m[1]; mo = +m[2]; da = +m[3]; }
        else if ((m = d.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) {
            /* Ambiguous slash format. SAP EWA exports use MM/DD/YYYY; if the
               first field is >12 it must be the day (DD/MM/YYYY). */
            const a = +m[1], b = +m[2];
            if (a > 12) { da = a; mo = b; } else { mo = a; da = b; }
            y = +m[3];
        } else {
            return null;
        }
        const t = String(time || '').trim();
        const tm = t.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
        const hh = tm ? +tm[1] : 0, mi = tm ? +tm[2] : 0, ss = tm && tm[3] ? +tm[3] : 0;
        return ((((y * 100 + mo) * 100 + da) * 100 + hh) * 100 + mi) * 100 + ss;
    },

    /* ══════════════════════════════════════════════════════
       TARGET TABLE BUILDER
       ══════════════════════════════════════════════════════ */
    _buildTarget(source, headers, col, dataRows, app) {
        /* Canonicalise headers to field NAMES first. A file may arrive with field
           LABELS ("Application Server Instance", "CPU Utilization (Idle)", …); the
           rest of this builder keys exclusively off field names, and `col` was
           already built from canonical names by _mapColumns, so translating the
           header array here makes both the label and name formats take the exact
           same code path and produce identically-named output columns. */
        headers = headers.map(h => this._canonicalFieldName(h));

        const has = name => col[name] != null;
        const val = (r, name) => has(name) ? (r[col[name]] ?? '').toString().trim() : '';

        /* Detect the file-wide decimal separator up front so every _num() call
           interprets thousands/decimal punctuation consistently. */
        this._decimalSep = this._detectDecimalSep(dataRows, col);

        /* Columns that get the thousands-separator cleanup (integer part kept).
           Spec step 7 names EMALLOC, EMATTACH, HEAPSUMKB, PAGESHMKB, USERS,
           SESSIONS and FREE_MEM. The SMON free-memory columns (FREE_MEM_MB,
           FREE_MEM_MB_INC_FS) carry the same thousands punctuation, so they are
           cleaned too — their GB derivations depend on it. */
        const CLEAN_INT = ['EMALLOC', 'EMATTACH', 'HEAPSUMKB', 'PAGESHMKB', 'USERS',
                           'SESSIONS', 'FREE_MEM', 'FREE_MEM_MB', 'FREE_MEM_MB_INC_FS',
                           'EMGLOBALKB'];

        /* Build the ordered list of OUTPUT columns.
           We walk the ORIGINAL header order (minus GUID / blank) and inject the
           derived columns at the positions the spec dictates. */
        const guidName = source === 'SMON' ? 'GUID' : 'ANALYSIS_GUID';

        /* Columns to DROP from the target file (queue/flag/status columns that
           are not needed for the sizing analysis). This also keeps the output
           narrow enough for the XLSX writer. SMON drops a few extra columns. */
        const DROP_SMON = ['DIAAVG20','DIAAVG60','UPDQ','ENQQ','SM50_FLAG','QUEUE_FLAG',
            'ST06_FLAG','TOPCPU_FLAG','ST02_FLAG','SM04_FLAG','SM12_CNT','SM12_FLAG',
            'SM12_STAT_FLAG','QINLENGTH','QIN_FLAG','QOUTLENGTH','QOUT_FLAG',
            'READY_TIME','STEAL_TIME'];
        const DROP_MON = ['UPDQ','ENQQ','SM50_FLAG','QUEUE_FLAG','ST06_FLAG','TOPCPU_FLAG',
            'ST02_FLAG','SM04_FLAG','SM12_CNT','SM12_FLAG','SM12_STAT_FLAG','QINLENGTH',
            'QIN_FLAG','QOUTLENGTH','QOUT_FLAG'];
        const DROP = new Set(source === 'MON' ? DROP_MON : DROP_SMON);

        /* Decide which raw columns to carry (skip the blank first col, guid, and
           the dropped queue/flag/status columns). */
        const carried = headers.filter(h => {
            const k = h.toUpperCase().replace(/\s+/g, '');
            if (!k) return false;                 /* blank "| |" column */
            if (k === guidName) return false;      /* drop GUID */
            if (DROP.has(k)) return false;         /* drop queue/flag/status columns */
            return true;
        });

        /* Compose output column order with inserts. */
        const outCols = [];
        const pushOnce = c => { if (!outCols.includes(c)) outCols.push(c); };

        carried.forEach(h => {
            const k = h.toUpperCase().replace(/\s+/g, '');

            if (k === 'DATUM') { pushOnce('Date'); return; }

            /* Physical Memory (GB) + EM Configured (GB) sit to the LEFT of EMALLOC */
            if (k === 'EMALLOC') {
                pushOnce('Physical Memory (GB)');
                pushOnce('EM Configured (GB)');    /* right of Physical Memory, left of EMALLOC */
                pushOnce('EMALLOC');
                pushOnce('EM Allocated (GB)');     /* right of EMALLOC */
                return;
            }

            pushOnce(h);

            if (k === 'SERVER' && this._mappingActive) pushOnce('SERVER (EWA)'); /* right of SERVER */
            if (k === 'TIME')       pushOnce('Hours');                 /* right of TIME */
            if (k === 'IDLE_TOTAL') pushOnce('CPU Utilization (%)');   /* right of IDLE_TOTAL */
            if (k === 'HEAPSUMKB')  pushOnce('Heap Memory (GB)');      /* right of HEAPSUMKB */
            if (k === 'FREE_MEM_MB' && source === 'SMON') pushOnce('Free Memory (GB)');
            if (k === 'FREE_MEM_MB_INC_FS' && source === 'SMON') pushOnce('Free Memory + FS (GB)');
            if (k === 'FREE_MEM' && source === 'MON') pushOnce('Free Memory (GB)');
        });

        /* Landscape-derived trailing columns (step 8) */
        pushOnce('Cores');
        pushOnce('SAPS/Core (SAPS)');
        pushOnce('SAPS Deployed (SAPS)');
        pushOnce('SAPS Utilization (SAPS)');

        /* ── Build each output row as a plain ARRAY (indexed by outCols) ──
           Arrays are used instead of {colName:value} objects: with up to ~1M
           rows (a month of per-minute snapshots × many servers) the per-key
           overhead of objects would exhaust browser memory. We also accumulate
           summary stats here in a single pass so _renderResults never has to
           re-scan the full dataset. */
        const round = (n, d = 2) => (isNaN(n) ? '' : Number(n.toFixed(d)));
        let matched = 0, unmatched = 0;
        const unmatchedHosts = new Set();

        /* Column name → output index (built once). */
        const ci = {};
        outCols.forEach((c, i) => { ci[c] = i; });
        const nCols = outCols.length;

        /* Pre-resolve which carried columns need cleanup and their target index,
           so the hot loop avoids repeated string work per cell. */
        const carriedPlan = carried
            .filter(h => h.toUpperCase().replace(/\s+/g, '') !== 'DATUM')
            .map(h => ({
                k: h.toUpperCase().replace(/\s+/g, ''),
                idx: ci[h],
                clean: CLEAN_INT.includes(h.toUpperCase().replace(/\s+/g, '')),
            }))
            .filter(p => p.idx != null);

        const idxDate = ci['Date'], idxHours = ci['Hours'], idxSrv = ci['SERVER'];
        const idxSrvEwa = ci['SERVER (EWA)'], idxCpuU = ci['CPU Utilization (%)'];
        const idxCores = ci['Cores'], idxSpc = ci['SAPS/Core (SAPS)'];
        const idxDep = ci['SAPS Deployed (SAPS)'], idxUtil = ci['SAPS Utilization (SAPS)'];
        const idxPhysMem = ci['Physical Memory (GB)'], idxEmGb = ci['EM Allocated (GB)'];
        const idxEmCfg = ci['EM Configured (GB)'];
        const idxHeap = ci['Heap Memory (GB)'], idxFree = ci['Free Memory (GB)'];
        const idxFreeFs = ci['Free Memory + FS (GB)'];
        const freeSrcCol = source === 'SMON' ? 'FREE_MEM_MB' : 'FREE_MEM';

        /* Single-pass summary accumulators (avoid a second scan of ~1M rows). */
        const hostSet = new Set();
        let cpuSum = 0, cpuCnt = 0, cpuMax = -Infinity;
        const cpuSample = [], sapsSample = [];   /* reservoir for percentiles */
        const RES_CAP = 50000;

        /* Chronological sort key per row (DATUM + TIME → sortable number), so the
           target file / charts run from the earliest date+time to the latest. */
        const sortKeys = new Array(dataRows.length);

        /* SAP Extended Memory configured (GB) per virtual host (lowercase key),
           entered by the user on the EM-config step. Snapshot once for the loop. */
        const emCfg = this._extMemory || {};

        const rowsOut = new Array(dataRows.length);
        /* De-duplicate records that repeat across multiple uploaded files.
           A record is a duplicate when DATUM + TIME + SERVER all match a row
           already written. Keyed on the RAW cell values (trimmed, SERVER
           case-insensitive) so the comparison is independent of later
           reformatting. The first occurrence wins; later copies are skipped. */
        const seenKeys = new Set();
        let duplicates = 0;
        for (let ri = 0; ri < dataRows.length; ri++) {
            const r = dataRows[ri];

            /* Build the duplicate key from the raw DATUM / TIME / SERVER values.
               Only de-dup when all three source columns are present; otherwise a
               reduced-column export can't be judged and every row is kept. */
            if (has('DATUM') && has('TIME') && has('SERVER')) {
                const dupKey = val(r, 'DATUM') + '\u0001' + val(r, 'TIME') + '\u0001' +
                               val(r, 'SERVER').toLowerCase();
                if (seenKeys.has(dupKey)) {
                    duplicates++;
                    dataRows[ri] = null;      /* free the skipped source row */
                    continue;                 /* rowsOut[ri] stays undefined → compacted out */
                }
                seenKeys.add(dupKey);
            }

            const o = new Array(nCols);

            /* Raw carried values (with per-column cleanup) */
            for (let p = 0; p < carriedPlan.length; p++) {
                const pl = carriedPlan[p];
                let v = val(r, pl.k);
                if (pl.clean) { const n = this._num(v, false); v = isNaN(n) ? v : n; }
                o[pl.idx] = v;
            }

            if (idxDate  != null && has('DATUM')) o[idxDate]  = this._reformatDate(val(r, 'DATUM'));
            if (idxHours != null && has('TIME'))  o[idxHours] = this._hourOf(val(r, 'TIME'));
            sortKeys[ri] = this._dateTimeKey(val(r, 'DATUM'), val(r, 'TIME'));

            const hostname = this._hostOf(val(r, 'SERVER'));
            if (idxSrv != null) o[idxSrv] = hostname;
            if (hostname) hostSet.add(hostname);

            /* CPU Utilization (%) = 100 - IDLE_TOTAL */
            let cpuUtil = NaN;
            if (has('IDLE_TOTAL')) {
                const idle = this._num(val(r, 'IDLE_TOTAL'));
                if (!isNaN(idle)) { cpuUtil = 100 - idle; if (idxCpuU != null) o[idxCpuU] = round(cpuUtil, 2); }
                else if (idxCpuU != null) o[idxCpuU] = '';
            }

            /* Landscape cross-reference (resolve physical host via mapping). */
            let physical = null, host = null;
            if (this._mappingActive) {
                physical = this._mapping[hostname.toLowerCase()] || null;
                host = physical ? app.findHost(physical) : null;
                if (idxSrvEwa != null) o[idxSrvEwa] = physical || 'No match in EWA report';
            } else {
                host = app.findHost(hostname);
            }
            if (host) matched++; else { unmatched++; if (hostname) unmatchedHosts.add(hostname); }

            const cores       = host ? (host.cores || 0) : '';
            const sapsPerCore = host ? (host.sapsPerCore || 0) : '';
            if (idxCores != null) o[idxCores] = cores === '' ? '' : cores;
            if (idxSpc  != null) o[idxSpc]  = sapsPerCore === '' ? '' : sapsPerCore;

            const sapsDeployed = (host && cores && sapsPerCore) ? cores * sapsPerCore : NaN;
            if (idxDep != null) o[idxDep] = isNaN(sapsDeployed) ? '' : sapsDeployed;

            const sapsUtil = (!isNaN(sapsDeployed) && !isNaN(cpuUtil)) ? (cpuUtil / 100) * sapsDeployed : NaN;
            if (idxUtil != null) o[idxUtil] = isNaN(sapsUtil) ? '' : round(sapsUtil, 0);

            /* Memory columns */
            if (idxPhysMem != null) o[idxPhysMem] = host && host.memoryMB ? round(host.memoryMB / 1024, 2) : '';
            if (idxEmCfg != null)   { const g = emCfg[(hostname || '').toLowerCase()]; o[idxEmCfg] = (g == null || g === '') ? '' : g; }
            if (idxEmGb != null)    { const em = this._num(val(r, 'EMALLOC'), false); o[idxEmGb] = isNaN(em) ? '' : round(em / 1024, 2); }
            if (idxHeap != null)    { const hp = this._num(val(r, 'HEAPSUMKB'), false); o[idxHeap] = isNaN(hp) ? '' : round(hp / 1024 / 1024, 3); }
            if (idxFree != null)    { const fm = this._num(val(r, freeSrcCol), false); o[idxFree] = isNaN(fm) ? '' : round(fm / 1024, 2); }
            if (idxFreeFs != null)  { const ff = this._num(val(r, 'FREE_MEM_MB_INC_FS'), false); o[idxFreeFs] = isNaN(ff) ? '' : round(ff / 1024, 2); }

            /* Accumulate stats in-pass */
            if (!isNaN(cpuUtil)) {
                cpuSum += cpuUtil; cpuCnt++;
                if (cpuUtil > cpuMax) cpuMax = cpuUtil;
                if (cpuSample.length < RES_CAP) cpuSample.push(cpuUtil);
            }
            if (!isNaN(sapsUtil) && sapsSample.length < RES_CAP) sapsSample.push(sapsUtil);

            /* Let go of the source row ASAP to free memory as we go. */
            dataRows[ri] = null;
            rowsOut[ri] = o;
        }

        /* ── Chronological sort (earliest Date+Time first, latest last) ──
           Build a compact index of KEPT rows only (duplicates left holes in
           rowsOut), sort it by the per-row key, then materialise the result.
           Rows with an unparseable key sort to the end (stable among themselves). */
        const order = [];
        for (let i = 0; i < rowsOut.length; i++) {
            if (rowsOut[i] !== undefined) order.push(i);
        }
        order.sort((a, b) => {
            const ka = sortKeys[a], kb = sortKeys[b];
            if (ka === kb) return a - b;        /* stable for equal keys */
            if (ka == null) return 1;
            if (kb == null) return -1;
            return ka - kb;
        });
        const sorted = new Array(order.length);
        for (let i = 0; i < order.length; i++) sorted[i] = rowsOut[order[i]];
        /* Replace rowsOut with the compacted, sorted result. */
        rowsOut.length = 0;
        for (let i = 0; i < sorted.length; i++) rowsOut[i] = sorted[i];
        sorted.length = 0;

        if (duplicates > 0) {
            try { console.log(`[TechWorkloadApp] skipped ${duplicates} duplicate record(s) (same DATUM+TIME+SERVER).`); } catch (_) {}
        }
        this._duplicatesSkipped = duplicates;

        this._matchStats = { matched, unmatched, unmatchedHosts: [...unmatchedHosts].slice(0, 50) };
        this._summary = {
            dataRows: rowsOut.length,
            hosts: hostSet.size,
            cpuAvg: cpuCnt ? cpuSum / cpuCnt : null,
            cpuMax: cpuMax === -Infinity ? null : cpuMax,
            cpuP95: cpuSample.length ? percentile(cpuSample, 95) : null,
            sapsP95: sapsSample.length ? percentile(sapsSample, 95) : null,
        };
        /* colIndex lets array-row consumers (preview, CSV, XLSX) resolve a column
           name to its position without rebuilding the map. */
        this._colIndex = ci;
        return { cols: outCols, rows: rowsOut };
    },

    /* ══════════════════════════════════════════════════════
       RESULTS RENDER
       ══════════════════════════════════════════════════════ */
    _renderResults(app) {
        document.getElementById('twa-results').style.display = 'block';
        this._hideError();
        /* Dismiss the optional profiles screen if it was open; retained in this.profiles. */
        this._hidePflViewOnly();

        const rows = this.targetRows, cols = this.targetCols;
        const srcLabel = this.sourceTable === 'PREPARSED' ? 'Pre-processed spreadsheet'
                       : this.sourceTable === 'MON' ? '/SDF/MON_HEADER' : '/SDF/SMON_HEADER';

        const srcNote = document.getElementById('twa-src-note');
        if (srcNote) srcNote.textContent = `— source table: ${srcLabel}`;

        /* Summary stats come from the single-pass accumulator built during
           _buildTarget — we never re-scan the (potentially ~1M-row) dataset. */
        const sm = this._summary || { dataRows: rows.length, hosts: 0, cpuP95: null, cpuMax: null, sapsP95: null };
        const ms = this._matchStats || { matched: 0, unmatched: 0, unmatchedHosts: [] };

        document.getElementById('twa-stats').innerHTML = [
            statCard('Source Table', srcLabel.replace('/SDF/', '')),
            statCard('Data Rows', sm.dataRows),
            statCard('Hosts', sm.hosts),
            sm.cpuP95 != null ? statCard('CPU Util P95', sm.cpuP95.toFixed(1), '%') : '',
            sm.cpuMax != null ? statCard('CPU Util Max', sm.cpuMax.toFixed(1), '%') : '',
            sm.sapsP95 != null ? statCard('SAPS Util P95', Math.round(sm.sapsP95).toLocaleString(), ' SAPS') : '',
        ].filter(Boolean).join('');

        /* Warn when the loaded session had its row data truncated or stripped to
           fit within JSON string-length limits. */
        const truncNote = document.getElementById('twa-truncation-note');
        if (this._targetRowsDropped) {
            if (truncNote) {
                truncNote.style.display = 'flex';
                truncNote.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                    fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                    <path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
                    <line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
                </svg>
                <span><strong>Partial restore:</strong> The monitoring dataset (${this._targetRowsOriginalCount ? this._targetRowsOriginalCount.toLocaleString() + ' rows' : 'large dataset'}) was too large to save in the session file — the enriched table and charts were not restored. Re-upload the original /SDF/SMON_HEADER or /SDF/MON_HEADER file to rebuild charts. All other state (mappings, extended memory, profiles, SAPS overrides) was fully preserved.</span>`;
            }
        } else if (this._targetRowsTruncated && truncNote) {
            truncNote.style.display = 'flex';
            truncNote.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
            </svg>
            <span><strong>Note:</strong> The monitoring dataset (${this._targetRowsOriginalCount ? this._targetRowsOriginalCount.toLocaleString() + ' rows' : 'large dataset'}) was trimmed to ${rows.length.toLocaleString()} rows when saving the session. Charts may show a partial picture. Re-upload the original file for full fidelity.</span>`;
        } else if (truncNote) {
            truncNote.style.display = 'none';
        }

        document.getElementById('twa-table-sub').textContent =
            `${sm.dataRows} row(s) · ${cols.length} columns · ${ms.matched} row(s) matched to landscape hosts` +
            (ms.unmatched ? ` · ${ms.unmatched} unmatched` : '');

        /* Preview table (cap at 200 rows for the DOM; full set still exports).
           Rows are arrays indexed by column position. */
        const PREVIEW = 200;
        document.getElementById('twa-thead').innerHTML =
            `<tr>${cols.map(c => `<th style="white-space:nowrap">${c}</th>`).join('')}</tr>`;
        const previewRows = rows.slice(0, PREVIEW);
        document.getElementById('twa-tbody').innerHTML = previewRows.map(r =>
            `<tr>${r.map(v =>
                `<td style="white-space:nowrap">${v === '' || v == null ? '—' : v}</td>`
            ).join('')}</tr>`
        ).join('');

        const prevNote = document.getElementById('twa-preview-note');
        if (prevNote) {
            let msg = rows.length > PREVIEW
                ? `Showing first ${PREVIEW} of ${rows.length} rows. The downloaded file contains all rows.`
                : `Showing all ${rows.length} rows.`;
            if (this._mappingActive) {
                msg += ` Virtual→physical mapping applied — see the SERVER (EWA) column.`;
            }
            if (this._duplicatesSkipped) {
                msg += ` Removed ${this._duplicatesSkipped} duplicate record(s) (same DATUM + TIME + SERVER across the uploaded files).`;
            }
            if (ms.unmatched) {
                const label = this._mappingActive
                    ? ` Virtual host(s) with no physical match (flagged “No match in EWA report”, Cores/SAPS blank): `
                    : ` Unmatched host(s) (no landscape record — Cores/SAPS left blank): `;
                msg += label + ms.unmatchedHosts.slice(0, 10).join(', ') +
                       (ms.unmatchedHosts.length > 10 ? '…' : '') + '.';
            }
            prevNote.textContent = msg;
        }

        document.getElementById('twa-btn-download').onclick = () => this._download();
        document.getElementById('twa-btn-charts').onclick = () => this._showChartsView(app);

        /* Charts-view navigation (both top and bottom button pairs). */
        const toTable = () => this._showTableView();
        /* After the charts, the flow continues to the optional Step 3 (ST02
           Memory Buffer) → Step 4 (ST02 WP Storage) → optional NMON (IBM only)
           → Database. The charts "Next" always goes to Step 3. */
        const toNext = () => this._showSt02BufView(app);
        ['twa-charts-back', 'twa-charts-back2'].forEach(id => {
            const b = document.getElementById(id); if (b) b.onclick = toTable;
        });
        ['twa-charts-next', 'twa-charts-next2'].forEach(id => {
            const b = document.getElementById(id); if (b) b.onclick = toNext;
        });

        /* Relabel the charts "Next" button — Step 3 (Memory Buffer) always follows. */
        ['twa-charts-next', 'twa-charts-next2'].forEach(id => {
            const b = document.getElementById(id);
            if (b) {
                const svg = b.querySelector('svg');
                b.textContent = 'Next: Memory Buffer Analysis (optional) ';
                if (svg) b.appendChild(svg);
            }
        });

        /* ── Step 3 (ST02 Memory Buffer) view buttons + upload zone ── */
        document.getElementById('twa-st02buf-back')
            ?.addEventListener('click', () => this._showChartsView(app));
        document.getElementById('twa-st02buf-next')
            ?.addEventListener('click', () => this._showSt02WpView(app));
        this._wireSt02BufZone(app);
        if (this.st02BufFiles && this.st02BufFiles.length) this._renderSt02BufFileList();
        if (this.st02BufAnalysis) this._renderSt02BufSummary();

        /* ── Step 4 (ST02 WP Storage) view buttons + upload zone ── */
        document.getElementById('twa-st02wp-back')
            ?.addEventListener('click', () => this._showSt02BufView(app));
        document.getElementById('twa-st02wp-next')
            ?.addEventListener('click', () => {
                if (this._ibmDetected()) this._showNmonView(app);
                else app.nextModule();
            });
        /* The Step 4 "Next" leads to IBM NMON when IBM servers exist. */
        if (this._ibmDetected()) {
            const b = document.getElementById('twa-st02wp-next');
            if (b) {
                const svg = b.querySelector('svg');
                b.textContent = 'Next: IBM NMON (optional) ';
                if (svg) b.appendChild(svg);
            }
        }
        this._wireSt02WpZone(app);
        if (this.st02WpFiles && this.st02WpFiles.length) this._renderSt02WpFileList();
        if (this.st02WpAnalysis) this._renderSt02WpSummary();

        /* NMON view buttons (optional IBM step — Step 5). "Previous" returns to Step 4. */
        document.getElementById('twa-nmon-back')
            ?.addEventListener('click', () => { this._showSt02WpView(app); });
        document.getElementById('twa-nmon-next')
            ?.addEventListener('click', () => app.nextModule());
        this._wireNmonZone(app);
        /* Rehydrate the NMON file list + summary when returning to this module. */
        if (this.nmonFiles && this.nmonFiles.length) this._renderNmonFileList();
        if (this.nmonAnalysis) this._renderNmonSummary(this.nmonAnalysis.fileName);

        /* Physical↔Virtual Server Inventory recap (only when mapping was used) */
        this._renderPhysicalVirtualInventory(app);

        /* Build the left-sidebar jump list for this module. */
        this._updateSubNav(app, 'table');
    },

    /* True when Module 01 flagged IBM Power servers in the landscape. */
    _ibmDetected() {
        const sl = this.app && this.app.getModuleData('system-landscape');
        return !!(sl && sl.hasIBMServers);
    },

    /* ── Toggle between the enriched-table view and the charts view ── */
    async _showChartsView(app) {
        const up = document.getElementById('twa-upload-screen');
        if (up) up.style.display = 'block';
        document.getElementById('twa-results').style.display = 'none';
        const nmon = document.getElementById('twa-nmon');
        if (nmon) nmon.style.display = 'none';
        this._hideStepViews();
        this._hidePflViewOnly();
        document.getElementById('twa-charts').style.display = 'block';
        document.getElementById('content-area')?.scrollTo({ top: 0, behavior: 'smooth' });

        /* Show the Network tab only for SMON source data. */
        const netTab = document.querySelector('#twa-chart-tabs .chart-tab[data-tab="network"]');
        if (netTab) netTab.style.display = this.sourceTable === 'SMON' ? '' : 'none';

        /* Wire the tab strip once. Tabs render lazily so only ONE section's charts
           are ever live in the DOM — this is what keeps the page from freezing
           (previously all four sections built ~40-60 charts at once). */
        this._wireChartTabs(app);

        /* Wire the top "Download Analysis workbook" button (idempotent). The
           workbook is rebuilt on demand at click time, so it always reflects the
           current chart data. */
        const wbBtn = document.getElementById('twa-charts-wb-top');
        if (wbBtn && !wbBtn._wired) {
            wbBtn._wired = true;
            wbBtn.addEventListener('click', () => this._buildAnalysisWorkbook(app));
        }

        /* Reset the "already rendered" tracker so a fresh dataset rebuilds. */
        this._renderedSection = null;

        /* Build only the default (CPU) tab now. */
        await this._activateSection(app, 'cpu', true);
    },

    /* Wire the in-page tab strip (idempotent). */
    _wireChartTabs(app) {
        const strip = document.getElementById('twa-chart-tabs');
        if (!strip || strip._wired) return;
        strip._wired = true;
        strip.querySelectorAll('.chart-tab').forEach(btn => {
            btn.addEventListener('click', () => this._activateSection(app, btn.dataset.tab));
        });
    },

    /* Show one analysis section and build its charts on demand; tear down the
       charts of whatever section was previously shown. Only one section is live
       at a time, so a global Chart.js/Plotly teardown here is correct and simple. */
    async _activateSection(app, sectionId, force = false) {
        const SECTIONS = ['cpu', 'mem', 'wp', 'logins', 'network'];
        if (!SECTIONS.includes(sectionId)) sectionId = 'cpu';

        /* Make sure the charts view is actually visible (sidebar may call this). */
        if (document.getElementById('twa-charts').style.display === 'none') {
            document.getElementById('twa-results').style.display = 'none';
            document.getElementById('twa-charts').style.display = 'block';
            this._wireChartTabs(app);
        }

        /* Re-clicking the active, already-rendered tab is a no-op. */
        if (!force && this._renderedSection === sectionId) {
            this._syncTabUi(sectionId);
            return;
        }

        /* Re-entrancy token: if the user clicks another tab mid-build, only the
           LATEST activation is allowed to drive (and hide) the progress bar, so a
           superseded build can't leave a stale "Building X… 45%" label on screen. */
        const token = (this._activateToken = (this._activateToken || 0) + 1);
        const current = () => this._activateToken === token;

        /* Tear down the previously-built charts (Chart.js + Plotly). */
        this._teardownCharts();

        /* Toggle tab-strip + section visibility immediately. */
        this._syncTabUi(sectionId);
        if (app && app.setSubNavActive) app.setSubNavActive('tech-workload-app', sectionId);

        const label = { cpu: 'CPU', mem: 'memory', wp: 'work-process', logins: 'logins & sessions', network: 'network' }[sectionId];

        try {
            this._chartProgress(`Preparing ${label} charts…`, 15);
            await this._nextFrame();
            if (!current()) return;                     /* superseded by a newer click */

            this._chartProgress(`Building ${label} charts…`, 55);
            await this._nextFrame();                    /* let the bar paint at 55% before the heavy build */
            if (!current()) return;

            if (sectionId === 'cpu')         await this._renderCpuCharts(app);
            else if (sectionId === 'mem')     await this._renderMemCharts(app);
            else if (sectionId === 'wp')      await this._renderWpCharts(app);
            else if (sectionId === 'network') await this._renderNetworkCharts(app);
            else                              await this._renderLoginsCharts(app);

            if (!current()) return;                     /* a newer tab won the race */

            this._chartProgress('Finalising…', 90);
            await this._nextFrame();
            if (!current()) return;

            this._renderedSection = sectionId;
            this._chartProgress('Charts ready', 100);
            await this._nextFrame();
        } catch (err) {
            /* Never leave the bar stuck on an error — surface it and move on. */
            showToast('Chart generation issue', (err && err.message) ? err.message : String(err), 'error');
        } finally {
            /* Only the latest activation hides the bar (prevents a stale build from
               hiding a bar the newest build is still using, or vice-versa). */
            if (current()) {
                setTimeout(() => { if (current()) this._hideChartProgress(); }, 300);
            }
        }

        /* One-time confirmation that charts can be viewed. */
        if (current() && !this._chartsToastShown) {
            this._chartsToastShown = true;
            showToast('Charts ready',
                'Use the tabs (CPU · Memory · Work Process · Logins & Sessions · Network) to switch between analyses.',
                'success');
        }
    },

    /* Destroy all live Chart.js instances and purge all Plotly surfaces. */
    _teardownCharts() {
        if (this._charts) {
            Object.values(this._charts).forEach(e => { try { e.chart.destroy(); } catch (_) {} });
        }
        this._charts = {};
        if (this._surfaces) {
            Object.values(this._surfaces).forEach(e => { try { Plotly.purge(e.plotEl); } catch (_) {} });
        }
        this._surfaces = {};
    },

    /* Reflect the active section in the tab strip + section visibility. */
    _syncTabUi(sectionId) {
        document.querySelectorAll('#twa-chart-tabs .chart-tab').forEach(btn => {
            btn.classList.toggle('is-active', btn.dataset.tab === sectionId);
        });
        document.querySelectorAll('#twa-charts .chart-section').forEach(sec => {
            sec.style.display = sec.dataset.sec === sectionId ? 'block' : 'none';
        });
    },

    /* Progress bar dedicated to the charts view (the shared _showProgress bar
       lives in the results view, which is hidden while charts are shown). */
    _chartProgress(label, pct) {
        const box  = document.getElementById('twa-chart-progress');
        const lab  = document.getElementById('twa-chart-progress-label');
        const pctE = document.getElementById('twa-chart-progress-pct');
        const fill = document.getElementById('twa-chart-progress-fill');
        if (box)  box.style.display = 'block';
        if (lab)  lab.textContent = label;
        const p = Math.max(0, Math.min(100, Math.round(pct)));
        if (pctE) pctE.textContent = `${p}%`;
        if (fill) fill.style.width = `${p}%`;
    },
    _hideChartProgress() {
        const box = document.getElementById('twa-chart-progress');
        if (box) box.style.display = 'none';
    },

    /* Wire the collapse/expand headers for the four chart sections. */
    _wireSectionToggles() {
        document.querySelectorAll('#twa-charts .chart-section-head').forEach(head => {
            if (head._wired) return;
            head._wired = true;
            head.addEventListener('click', () => {
                head.closest('.chart-section').classList.toggle('collapsed');
            });
        });
    },
    _showTableView() {
        const up = document.getElementById('twa-upload-screen');
        if (up) up.style.display = 'block';
        document.getElementById('twa-charts').style.display = 'none';
        const nmon = document.getElementById('twa-nmon');
        if (nmon) nmon.style.display = 'none';
        this._hideStepViews();
        this._hidePflViewOnly();
        document.getElementById('twa-results').style.display = 'block';
        if (this.app) this._updateSubNav(this.app, 'table');
        document.getElementById('content-area')?.scrollTo({ top: 0, behavior: 'smooth' });
    },

    /* ══════════════════════════════════════════════════════
       OPTIONAL NMON STEP (IBM Power Systems)
       Shown only when Module 01 detected IBM servers. The mandatory
       SMON/MON analysis is unaffected; NMON is purely additive.
       ══════════════════════════════════════════════════════ */
    _showNmonView(app) {
        const up = document.getElementById('twa-upload-screen');
        if (up) up.style.display = 'block';
        document.getElementById('twa-results').style.display = 'none';
        document.getElementById('twa-charts').style.display  = 'none';
        this._hideStepViews();
        this._hidePflViewOnly();
        const nmon = document.getElementById('twa-nmon');
        if (nmon) nmon.style.display = 'block';
        if (this.nmonFiles && this.nmonFiles.length) this._renderNmonFileList();
        if (app && app.setSubNavActive) app.setSubNavActive('tech-workload-app', 'nmon');
        document.getElementById('content-area')?.scrollTo({ top: 0, behavior: 'smooth' });
    },

    /* Hide the Step 3 / Step 4 ST02 views (used by every other view switcher so
       the screens stay mutually exclusive). */
    _hideStepViews() {
        ['twa-st02buf', 'twa-st02wp'].forEach(id => {
            const n = document.getElementById(id);
            if (n) n.style.display = 'none';
        });
    },

    /* ══════════════════════════════════════════════════════
       OPTIONAL STEP 3 — ST02 MEMORY BUFFER ANALYSIS
       ══════════════════════════════════════════════════════ */
    _showSt02BufView(app) {
        const up = document.getElementById('twa-upload-screen');
        if (up) up.style.display = 'none';
        document.getElementById('twa-results').style.display = 'none';
        document.getElementById('twa-charts').style.display  = 'none';
        const nmon = document.getElementById('twa-nmon');
        if (nmon) nmon.style.display = 'none';
        this._hidePflViewOnly();
        const wp = document.getElementById('twa-st02wp');
        if (wp) wp.style.display = 'none';
        const buf = document.getElementById('twa-st02buf');
        if (buf) buf.style.display = 'block';
        if (this.st02BufFiles && this.st02BufFiles.length) this._renderSt02BufFileList();
        if (this.st02BufAnalysis) this._renderSt02BufSummary();
        if (app && app.setSubNavActive) app.setSubNavActive('tech-workload-app', 'st02buf');
        document.getElementById('content-area')?.scrollTo({ top: 0, behavior: 'smooth' });
    },

    /* ══════════════════════════════════════════════════════
       OPTIONAL STEP 4 — ST02 USER STORAGE FOR WP MEMORY
       ══════════════════════════════════════════════════════ */
    _showSt02WpView(app) {
        const up = document.getElementById('twa-upload-screen');
        if (up) up.style.display = 'none';
        document.getElementById('twa-results').style.display = 'none';
        document.getElementById('twa-charts').style.display  = 'none';
        const nmon = document.getElementById('twa-nmon');
        if (nmon) nmon.style.display = 'none';
        this._hidePflViewOnly();
        const buf = document.getElementById('twa-st02buf');
        if (buf) buf.style.display = 'none';
        const wp = document.getElementById('twa-st02wp');
        if (wp) wp.style.display = 'block';
        if (this.st02WpFiles && this.st02WpFiles.length) this._renderSt02WpFileList();
        if (this.st02WpAnalysis) this._renderSt02WpSummary();
        if (app && app.setSubNavActive) app.setSubNavActive('tech-workload-app', 'st02wp');
        document.getElementById('content-area')?.scrollTo({ top: 0, behavior: 'smooth' });
    },

    /* ══════════════════════════════════════════════════════════════════
       ST02 SHARED HELPERS (Step 3 + Step 4)
       ══════════════════════════════════════════════════════════════════ */

    /* Strip the "." thousand separator and return an integer KB value. Blank /
       non-numeric cells become 0. SAP uses "," as the decimal separator, but the
       KB columns we read are whole numbers, so a stray "," is treated as a
       separator too (defensive). */
    _st02Kb(s) {
        const t = String(s == null ? '' : s).replace(/[.\s,]/g, '');
        const n = parseInt(t, 10);
        return isNaN(n) ? 0 : n;
    },

    /* KB → GB (binary), rounded to 2 decimals. */
    _st02Gb(kb) {
        return Math.round((Number(kb) / 1024 / 1024) * 100) / 100;
    },

    /* Split a pipe-delimited line into trimmed cells (drops the leading/trailing
       empty cells produced by the bounding "|"). Returns null for non-table lines. */
    _st02PipeCells(line) {
        if (!line || line.indexOf('|') === -1) return null;
        const parts = line.split('|');
        if (parts.length < 3) return null;
        return parts.slice(1, -1).map(c => c.trim());
    },

    /* Normalise an uploaded file into plain text. TXT/CSV are read as-is; XLSX/XLS
       are read via SheetJS and the first sheet is reconstructed into the same
       pipe-delimited layout the text parser expects (each cell joined with "|",
       wrapped so _st02PipeCells sees the bounding separators). */
    async _st02ReadFileText(file) {
        const name = (file.name || '').toLowerCase();
        if (/\.(xlsx|xls)$/.test(name)) {
            const buf = await file.arrayBuffer();
            const wb  = XLSX.read(buf, { type: 'array' });
            const ws  = wb.Sheets[wb.SheetNames[0]];
            if (!ws) return '';
            const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: true, defval: '' });
            return aoa.map(row => {
                const cells = (row || []).map(c => (c == null ? '' : String(c)));
                /* A single-column cell (e.g. the System:/Date header lines) is kept
                   verbatim; multi-column rows are rebuilt as a pipe table row. */
                if (cells.length <= 1) return cells[0] || '';
                return '|' + cells.join('|') + '|';
            }).join('\n');
        }
        return await file.text();
    },

    /* ══════════════════════════════════════════════════════════════════
       STEP 3 — ST02 MEMORY BUFFER: zone, validation, parsing, processing
       ══════════════════════════════════════════════════════════════════ */
    _wireSt02BufZone(app) {
        const zone  = document.getElementById('twa-st02buf-zone');
        const input = document.getElementById('twa-st02buf-input');
        if (!zone || !input || zone._wired) return;
        zone._wired = true;
        zone.addEventListener('click', () => input.click());
        zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
        zone.addEventListener('drop', e => {
            e.preventDefault(); zone.classList.remove('drag-over');
            this._handleSt02BufFiles(Array.from(e.dataTransfer.files), app);
        });
        input.addEventListener('change', e => {
            this._handleSt02BufFiles(Array.from(e.target.files), app);
            e.target.value = '';
        });
    },

    async _handleSt02BufFiles(files, app) {
        if (!files || !files.length) return;
        const ALLOWED = /\.(txt|text|csv|xlsx|xls)$/i;
        const accepted = files.filter(f => ALLOWED.test(f.name));
        if (!accepted.length) {
            showToast('Unsupported file', 'Step 3 accepts .txt, .csv or .xlsx files.', 'error');
            return;
        }
        const existing = new Set((this.st02BufFiles || []).map(f => f.name));
        for (const f of accepted) {
            if (existing.has(f.name)) continue;
            try {
                const text = await this._st02ReadFileText(f);
                (this.st02BufFiles = this.st02BufFiles || []).push({ name: f.name, text });
                existing.add(f.name);
            } catch (err) {
                console.error('ST02 buffer read failed', f.name, err);
                showToast('Read failed', `Could not read ${f.name}.`, 'error');
            }
        }
        this._renderSt02BufFileList();
        this._processSt02Buf(app);
    },

    _renderSt02BufFileList() {
        const list = document.getElementById('twa-st02buf-file-list');
        const ul   = document.getElementById('twa-st02buf-files-ul');
        if (!list || !ul) return;
        const files = this.st02BufFiles || [];
        if (!files.length) { list.style.display = 'none'; return; }
        list.style.display = 'block';
        ul.innerHTML = files.map((f, i) => this._st02FileLi(f.name, 'twa-st02buf', i)).join('');
    },

    _removeSt02BufFile(idx) {
        if (!this.st02BufFiles) return;
        this.st02BufFiles.splice(idx, 1);
        this._renderSt02BufFileList();
        if (!this.st02BufFiles.length) {
            this.st02BufAnalysis = null;
            const sum = document.getElementById('twa-st02buf-summary');
            if (sum) { sum.style.display = 'none'; sum.innerHTML = ''; }
            this._saveSt02ToApp();
        } else if (this.app) {
            this._processSt02Buf(this.app);
        }
    },

    /* Validate one ST02 buffer file. Returns { ok, reason }. */
    _validateSt02BufFile(name, text) {
        const lines = String(text || '').split('\n').map(l => l.replace(/\r$/, ''));
        const l1 = (lines[0] || '').replace(/\s+/g, ' ').trim();
        const l2 = (lines[1] || '').replace(/\s+/g, ' ').trim();
        if (!/^System:\s+\S+\s+Tune summary/i.test(l1)) {
            return { ok: false, reason: 'Line 1 is not "System: <host> Tune summary".' };
        }
        if (!/Snapshot[^:]*:\s*[\d.]+\s+[\d:]+/i.test(l2) || !/Startup:\s*[\d.]+\s+[\d:]+/i.test(l2)) {
            return { ok: false, reason: 'Line 2 is missing the Snapshot and/or Startup date-time.' };
        }
        const whole = text;
        const hasBufferTable = lines.some(l => /\|/.test(l) && /Alloc\.?\s*KB/i.test(l) && /Freesp\.?\s*KB/i.test(l) && /Swaps/i.test(l));
        if (!hasBufferTable) return { ok: false, reason: 'Buffer table (Alloc. KB / Freesp. KB / Swaps) not found.' };
        if (!/SAP Memory/i.test(whole)) return { ok: false, reason: 'SAP Memory section not found.' };
        if (!/Call Stat/i.test(whole)) return { ok: false, reason: 'Call Statistics section not found.' };
        return { ok: true, reason: '' };
    },

    /* Parse one ST02 buffer file into { hostname, snapshot, startup, rows:[{buffer,sizeKB,freeKB,swaps}] }. */
    _parseSt02BufFile(name, text) {
        const lines = String(text || '').split('\n').map(l => l.replace(/\r$/, ''));
        const l1 = (lines[0] || '').replace(/\s+/g, ' ').trim();
        const l2 = (lines[1] || '').replace(/\s+/g, ' ').trim();
        const hostname = (l1.match(/^System:\s+(\S+)\s+Tune summary/i) || [])[1] || name;
        const snapshot = (l2.match(/Snapshot[^:]*:\s*([\d.]+\s+[\d:]+)/i) || [])[1] || '';
        const startup  = (l2.match(/Startup:\s*([\d.]+\s+[\d:]+)/i) || [])[1] || '';

        /* Locate the Buffer-table header row and build a column-name → index map. */
        let hdrIdx = -1, colMap = null;
        for (let i = 0; i < lines.length; i++) {
            const cells = this._st02PipeCells(lines[i]);
            if (!cells) continue;
            if (cells.some(c => /Alloc\.?\s*KB/i.test(c)) && cells.some(c => /Freesp\.?\s*KB/i.test(c)) && cells.some(c => /Swaps/i.test(c))) {
                hdrIdx = i;
                colMap = {
                    buffer: 0,
                    alloc:  cells.findIndex(c => /Alloc\.?\s*KB/i.test(c)),
                    free:   cells.findIndex(c => /Freesp\.?\s*KB/i.test(c)),
                    swaps:  cells.findIndex(c => /Swaps/i.test(c)),
                };
                break;
            }
        }
        const rows = [];
        if (hdrIdx === -1 || !colMap) return { hostname, snapshot, startup, rows };

        let started = false;
        for (let i = hdrIdx + 1; i < lines.length; i++) {
            const raw = lines[i];
            const isSep  = /^\s*-{5,}\s*$/.test(raw);
            const cells  = this._st02PipeCells(raw);
            /* The Buffer table opens with a dashed separator right under the header
               and closes with another one. Skip the leading separator / any non-pipe
               line until data begins; once data has started, a separator or a
               non-pipe line marks the end of the table. */
            if (!cells || isSep) {
                if (started) break;
                continue;
            }
            started = true;
            const name2 = (cells[colMap.buffer] || '').trim();
            if (!name2) continue;                 // blank spacer row inside the table
            if (name2 === 'Tables') continue;      // classification label — not a real buffer
            const sizeKB = this._st02Kb(cells[colMap.alloc]);
            const freeKB = this._st02Kb(cells[colMap.free]);
            const swaps  = this._st02Kb(cells[colMap.swaps]);
            /* Skip rows with no allocation figure at all (pure spacer rows). */
            if (sizeKB === 0 && !/\d/.test(cells[colMap.alloc] || '')) continue;
            rows.push({ buffer: name2, sizeKB, freeKB, swaps });
        }
        return { hostname, snapshot, startup, rows };
    },

    _processSt02Buf(app) {
        const files = this.st02BufFiles || [];
        if (!files.length) return;
        const valid = [], invalid = [];
        files.forEach(f => {
            const v = this._validateSt02BufFile(f.name, f.text);
            if (v.ok) valid.push(f); else invalid.push({ name: f.name, reason: v.reason });
        });

        const build = () => this._buildSt02BufAnalysis(valid);
        if (invalid.length) {
            this._showFileValidationModal(
                'ST02 Buffer file validation',
                valid.map(f => ({ name: f.name })),
                invalid,
                valid.length ? build : null,
            );
        } else {
            build();
        }
    },

    _buildSt02BufAnalysis(validFiles) {
        this._st02BufProgress('Parsing ST02 buffer files…', 20);
        const bufferDetail = [];
        validFiles.forEach(f => {
            const p = this._parseSt02BufFile(f.name, f.text);
            p.rows.forEach(r => bufferDetail.push({
                Hostname: p.hostname,
                'Snapshot DateTime': p.snapshot,
                'Startup DateTime': p.startup,
                Buffer: r.buffer,
                'Buffer Size KB': r.sizeKB,
                'Free Space KB': r.freeKB,
                Swaps: r.swaps,
            }));
        });

        this._st02BufProgress('Aggregating…', 60);
        /* Table 1 — per buffer (across all servers). */
        const byBuffer = new Map();
        bufferDetail.forEach(r => byBuffer.set(r.Buffer, (byBuffer.get(r.Buffer) || 0) + r['Buffer Size KB']));
        const perBuffer = [...byBuffer.entries()].map(([Buffer, kb]) => ({
            Buffer, 'Total Size KB': kb, 'Total Size GB': this._st02Gb(kb),
        })).sort((a, b) => b['Total Size KB'] - a['Total Size KB']);

        /* Table 2 — per server (all buffers). */
        const byServer = new Map();
        bufferDetail.forEach(r => byServer.set(r.Hostname, (byServer.get(r.Hostname) || 0) + r['Buffer Size KB']));
        const perServer = [...byServer.entries()].map(([Hostname, kb]) => ({
            Hostname, 'Total Buffer Size KB': kb, 'Total Buffer Size GB': this._st02Gb(kb),
        })).sort((a, b) => b['Total Buffer Size KB'] - a['Total Buffer Size KB']);

        /* Table 3 — landscape total. */
        const totalKb = bufferDetail.reduce((s, r) => s + r['Buffer Size KB'], 0);
        const total = { 'Total Buffers KB': totalKb, 'Total Buffers GB': this._st02Gb(totalKb) };

        const fileName = `${this._sidTag()}_st02_buffer_memory.xlsx`;
        this.st02BufAnalysis = { fileName, bufferDetail, perBuffer, perServer, total, workbook: null };
        this.st02BufAnalysis.workbook = this._buildSt02BufWorkbook(this.st02BufAnalysis);

        this._st02BufProgress('Done', 100);
        setTimeout(() => this._hideSt02BufProgress(), 500);
        this._renderSt02BufSummary();
        this._saveSt02ToApp();
        showToast('ST02 Buffer analysis ready', `${validFiles.length} file(s) processed, ${bufferDetail.length} buffer rows.`, 'success');
    },

    _buildSt02BufWorkbook(a) {
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(a.bufferDetail, {
            header: ['Hostname', 'Snapshot DateTime', 'Startup DateTime', 'Buffer', 'Buffer Size KB', 'Free Space KB', 'Swaps'],
        }), 'Buffer_Detail');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(a.perBuffer, {
            header: ['Buffer', 'Total Size KB', 'Total Size GB'],
        }), 'Memory per Buffer');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(a.perServer, {
            header: ['Hostname', 'Total Buffer Size KB', 'Total Buffer Size GB'],
        }), 'Memory per Server');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([a.total], {
            header: ['Total Buffers KB', 'Total Buffers GB'],
        }), 'Landscape Total');
        return wb;
    },

    _renderSt02BufSummary() {
        const box = document.getElementById('twa-st02buf-summary');
        const a = this.st02BufAnalysis;
        if (!box || !a) return;
        box.style.display = 'block';
        const nf = n => Number(n).toLocaleString();
        box.innerHTML = `
            <div class="card" style="margin-top:20px">
                <div class="card-header" style="padding:14px 18px 12px">
                    <div>
                        <div class="card-title">ST02 Buffer Memory Workbook</div>
                        <div class="card-sub">${a.fileName} — Buffer_Detail + 3 summary sheets</div>
                    </div>
                    <button class="btn btn-blue" id="twa-st02buf-dl">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
                        </svg>
                        Download Analysis (XLSX)
                    </button>
                </div>
            </div>

            <div class="section-title" style="margin-top:20px">Memory Requirements per Buffer</div>
            ${this._st02HtmlTable(['Buffer', 'Total Size KB', 'Total Size GB'],
                a.perBuffer.map(r => [r.Buffer, nf(r['Total Size KB']), r['Total Size GB']]))}

            <div class="section-title" style="margin-top:20px">Memory Requirements per Server</div>
            ${this._st02HtmlTable(['Hostname', 'Total Buffer Size KB', 'Total Buffer Size GB'],
                a.perServer.map(r => [r.Hostname, nf(r['Total Buffer Size KB']), r['Total Buffer Size GB']]))}

            <div class="section-title" style="margin-top:20px">Landscape Total</div>
            ${this._st02HtmlTable(['Total Buffers KB', 'Total Buffers GB'],
                [[nf(a.total['Total Buffers KB']), a.total['Total Buffers GB']]])}
        `;
        const dl = document.getElementById('twa-st02buf-dl');
        if (dl) dl.onclick = () => {
            if (!this.st02BufAnalysis.workbook) this.st02BufAnalysis.workbook = this._buildSt02BufWorkbook(this.st02BufAnalysis);
            XLSX.writeFile(this.st02BufAnalysis.workbook, this.st02BufAnalysis.fileName);
        };
    },

    _st02BufProgress(label, pct) { this._st02Progress('twa-st02buf', label, pct); },
    _hideSt02BufProgress() { const b = document.getElementById('twa-st02buf-progress'); if (b) b.style.display = 'none'; },

    /* ══════════════════════════════════════════════════════════════════
       STEP 4 — ST02 WP STORAGE: zone, validation, parsing, processing
       ══════════════════════════════════════════════════════════════════ */
    _wireSt02WpZone(app) {
        const zone  = document.getElementById('twa-st02wp-zone');
        const input = document.getElementById('twa-st02wp-input');
        if (!zone || !input || zone._wired) return;
        zone._wired = true;
        zone.addEventListener('click', () => input.click());
        zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
        zone.addEventListener('drop', e => {
            e.preventDefault(); zone.classList.remove('drag-over');
            this._handleSt02WpFiles(Array.from(e.dataTransfer.files), app);
        });
        input.addEventListener('change', e => {
            this._handleSt02WpFiles(Array.from(e.target.files), app);
            e.target.value = '';
        });
    },

    async _handleSt02WpFiles(files, app) {
        if (!files || !files.length) return;
        const ALLOWED = /\.(txt|text|csv|xlsx|xls)$/i;
        const accepted = files.filter(f => ALLOWED.test(f.name));
        if (!accepted.length) {
            showToast('Unsupported file', 'Step 4 accepts .txt, .csv or .xlsx files.', 'error');
            return;
        }
        const existing = new Set((this.st02WpFiles || []).map(f => f.name));
        for (const f of accepted) {
            if (existing.has(f.name)) continue;
            try {
                const text = await this._st02ReadFileText(f);
                (this.st02WpFiles = this.st02WpFiles || []).push({ name: f.name, text });
                existing.add(f.name);
            } catch (err) {
                console.error('ST02 WP read failed', f.name, err);
                showToast('Read failed', `Could not read ${f.name}.`, 'error');
            }
        }
        this._renderSt02WpFileList();
        this._processSt02Wp(app);
    },

    _renderSt02WpFileList() {
        const list = document.getElementById('twa-st02wp-file-list');
        const ul   = document.getElementById('twa-st02wp-files-ul');
        if (!list || !ul) return;
        const files = this.st02WpFiles || [];
        if (!files.length) { list.style.display = 'none'; return; }
        list.style.display = 'block';
        ul.innerHTML = files.map((f, i) => this._st02FileLi(f.name, 'twa-st02wp', i)).join('');
    },

    _removeSt02WpFile(idx) {
        if (!this.st02WpFiles) return;
        this.st02WpFiles.splice(idx, 1);
        this._renderSt02WpFileList();
        if (!this.st02WpFiles.length) {
            this.st02WpAnalysis = null;
            const sum = document.getElementById('twa-st02wp-summary');
            if (sum) { sum.style.display = 'none'; sum.innerHTML = ''; }
            this._saveSt02ToApp();
        } else if (this.app) {
            this._processSt02Wp(this.app);
        }
    },

    _validateSt02WpFile(name, text) {
        const lines = String(text || '').split('\n').map(l => l.replace(/\r$/, ''));
        const l1 = (lines[0] || '').replace(/\s+/g, ' ').trim();
        const l2 = (lines[1] || '').replace(/\s+/g, ' ').trim();
        if (!/^System:\s+\S+\s+Date\s+[\d.]+/i.test(l1)) {
            return { ok: false, reason: 'Line 1 is not "System: <host> Date dd.mm.yyyy".' };
        }
        if (!/Time\s+[\d:]+/i.test(l2)) {
            return { ok: false, reason: 'Line 2 is missing the "Time hh:mm:ss" stamp.' };
        }
        const whole = text;
        const need = [
            [/Storage Shared Between Work Processes/i, '"Storage Shared Between Work Processes" section'],
            [/User Storage for Each of\s+\d+\s+Work Processes/i, '"User Storage for Each of ### Work Processes" section'],
            [/User Storage for all Work Processes/i, '"User Storage for all Work Processes" section'],
            [/Maximum Heap Area for all Work Processes/i, '"Maximum Heap Area for all Work Processes" section'],
        ];
        for (const [re, label] of need) {
            if (!re.test(whole)) return { ok: false, reason: `Missing ${label}.` };
        }
        return { ok: true, reason: '' };
    },

    _parseSt02WpFile(name, text) {
        const lines = String(text || '').split('\n').map(l => l.replace(/\r$/, ''));
        const l1 = (lines[0] || '').replace(/\s+/g, ' ').trim();
        const l2 = (lines[1] || '').replace(/\s+/g, ' ').trim();
        const hostname = (l1.match(/^System:\s+(\S+)\s+Date/i) || [])[1] || name;
        const date = (l1.match(/Date\s+([\d.]+)/i) || [])[1] || '';
        const time = (l2.match(/Time\s+([\d:]+)/i) || [])[1] || '';

        /* The value is the last numeric amount before a standalone "K" on the
           matching line (tolerant of a middle "Shared text" column). */
        const grab = (re) => {
            for (const raw of lines) {
                if (!re.test(raw)) continue;
                const m = raw.match(/([\d.]+)\s*K\b/);
                if (m) return this._st02Kb(m[1]);
            }
            return 0;
        };
        const userStorageKB = grab(/User storage for all work processes/i);
        const execKB = grab(/Size of SAP executables \(Disp\+Work\)/i);
        return { hostname, date, time, userStorageKB, execKB };
    },

    _processSt02Wp(app) {
        const files = this.st02WpFiles || [];
        if (!files.length) return;
        const valid = [], invalid = [];
        files.forEach(f => {
            const v = this._validateSt02WpFile(f.name, f.text);
            if (v.ok) valid.push(f); else invalid.push({ name: f.name, reason: v.reason });
        });

        const build = () => this._buildSt02WpAnalysis(app, valid);
        if (invalid.length) {
            this._showFileValidationModal(
                'ST02 Storage file validation',
                valid.map(f => ({ name: f.name })),
                invalid,
                valid.length ? build : null,
            );
        } else {
            build();
        }
    },

    _buildSt02WpAnalysis(app, validFiles) {
        this._st02WpProgress('Parsing ST02 storage files…', 20);
        const servers = [];
        validFiles.forEach(f => {
            const p = this._parseSt02WpFile(f.name, f.text);
            const host = app.findHost(p.hostname);
            const physGb = host && host.memoryMB ? Math.round((host.memoryMB / 1024) * 100) / 100 : '';
            servers.push({
                Hostname: p.hostname,
                'User Storage for all Work Processes KB': p.userStorageKB,
                'User Storage for all Work Processes GB': this._st02Gb(p.userStorageKB),
                'Size of SAP Executables (Disp+Work) KB': p.execKB,
                'Size of SAP Executables (Disp+Work) GB': this._st02Gb(p.execKB),
                'Physical Memory GB': physGb,
            });
        });

        this._st02WpProgress('Aggregating…', 60);
        const userKb = servers.reduce((s, r) => s + r['User Storage for all Work Processes KB'], 0);
        const execKb = servers.reduce((s, r) => s + r['Size of SAP Executables (Disp+Work) KB'], 0);
        const totalUser = { 'Total KB': userKb, 'Total GB': this._st02Gb(userKb) };
        const totalExec = { 'Total KB': execKb, 'Total GB': this._st02Gb(execKb) };

        const fileName = `${this._sidTag()}_st02_wp_memory.xlsx`;
        this.st02WpAnalysis = { fileName, servers, totalUser, totalExec, workbook: null };
        this.st02WpAnalysis.workbook = this._buildSt02WpWorkbook(this.st02WpAnalysis);

        this._st02WpProgress('Done', 100);
        setTimeout(() => this._hideSt02WpProgress(), 500);
        this._renderSt02WpSummary();
        this._saveSt02ToApp();
        showToast('ST02 WP storage analysis ready', `${validFiles.length} server file(s) processed.`, 'success');
    },

    _buildSt02WpWorkbook(a) {
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(a.servers, {
            header: ['Hostname',
                'User Storage for all Work Processes KB', 'User Storage for all Work Processes GB',
                'Size of SAP Executables (Disp+Work) KB', 'Size of SAP Executables (Disp+Work) GB',
                'Physical Memory GB'],
        }), 'WP_Storage');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([a.totalUser], {
            header: ['Total KB', 'Total GB'],
        }), 'Total User Storage');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([a.totalExec], {
            header: ['Total KB', 'Total GB'],
        }), 'Total SAP Executables');
        return wb;
    },

    _renderSt02WpSummary() {
        const box = document.getElementById('twa-st02wp-summary');
        const a = this.st02WpAnalysis;
        if (!box || !a) return;
        box.style.display = 'block';
        const nf = n => (n === '' || n == null) ? '—' : Number(n).toLocaleString();
        const gf = n => (n === '' || n == null) ? '—' : n;
        box.innerHTML = `
            <div class="card" style="margin-top:20px">
                <div class="card-header" style="padding:14px 18px 12px">
                    <div>
                        <div class="card-title">ST02 WP Memory Workbook</div>
                        <div class="card-sub">${a.fileName} — WP_Storage + 2 total sheets</div>
                    </div>
                    <button class="btn btn-blue" id="twa-st02wp-dl">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
                        </svg>
                        Download Analysis (XLSX)
                    </button>
                </div>
            </div>

            <div class="section-title" style="margin-top:20px">User Storage &amp; SAP Executables per Server</div>
            ${this._st02HtmlTable(
                ['Hostname', 'User Storage KB', 'User Storage GB', 'SAP Exec. KB', 'SAP Exec. GB', 'Physical Memory GB'],
                a.servers.map(r => [
                    r.Hostname,
                    nf(r['User Storage for all Work Processes KB']),
                    gf(r['User Storage for all Work Processes GB']),
                    nf(r['Size of SAP Executables (Disp+Work) KB']),
                    gf(r['Size of SAP Executables (Disp+Work) GB']),
                    gf(r['Physical Memory GB']),
                ]))}

            <div class="section-title" style="margin-top:20px">Landscape Total — User Storage for all Work Processes</div>
            ${this._st02HtmlTable(['Total KB', 'Total GB'], [[nf(a.totalUser['Total KB']), a.totalUser['Total GB']]])}

            <div class="section-title" style="margin-top:20px">Landscape Total — Size of SAP Executables (Disp+Work)</div>
            ${this._st02HtmlTable(['Total KB', 'Total GB'], [[nf(a.totalExec['Total KB']), a.totalExec['Total GB']]])}
        `;
        const dl = document.getElementById('twa-st02wp-dl');
        if (dl) dl.onclick = () => {
            if (!this.st02WpAnalysis.workbook) this.st02WpAnalysis.workbook = this._buildSt02WpWorkbook(this.st02WpAnalysis);
            XLSX.writeFile(this.st02WpAnalysis.workbook, this.st02WpAnalysis.fileName);
        };
    },

    _st02WpProgress(label, pct) { this._st02Progress('twa-st02wp', label, pct); },
    _hideSt02WpProgress() { const b = document.getElementById('twa-st02wp-progress'); if (b) b.style.display = 'none'; },

    /* ══════════════════════════════════════════════════════════════════
       ST02 UI HELPERS (shared)
       ══════════════════════════════════════════════════════════════════ */

    /* A single file-list <li> with a remove button. prefix is 'twa-st02buf' |
       'twa-st02wp'; the matching _removeSt02*File handler is derived from it. */
    _st02FileLi(fileName, prefix, i) {
        const fn = prefix === 'twa-st02buf' ? '_removeSt02BufFile' : '_removeSt02WpFile';
        const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
        return `<li style="display:flex;align-items:center;gap:8px;padding:3px 0;font-size:13px">
            <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24"
                 fill="none" stroke="var(--brand)" stroke-width="2.2">
                <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                <polyline points="22 4 12 14.01 9 11.01"/>
            </svg>
            <span>${esc(fileName)}</span>
            <button onclick="TechWorkloadAppModule.${fn}(${i})"
                    style="background:none;border:none;cursor:pointer;color:var(--error);
                           font-size:15px;line-height:1;padding:0 4px;margin-left:auto"
                    title="Remove">×</button>
        </li>`;
    },

    /* Build a simple read-only HTML table (reuses the module's .data-table style). */
    _st02HtmlTable(cols, rows) {
        const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
        const thead = `<tr>${cols.map(c => `<th>${esc(c)}</th>`).join('')}</tr>`;
        const tbody = rows.length
            ? rows.map(r => `<tr>${r.map(c => `<td>${esc(c)}</td>`).join('')}</tr>`).join('')
            : `<tr><td colspan="${cols.length}" style="text-align:center;color:var(--text-muted)">No data</td></tr>`;
        return `<div class="card" style="padding:0">
            <div class="table-wrap" style="border:none;border-radius:14px;overflow-x:auto;max-height:420px">
                <table class="data-table"><thead>${thead}</thead><tbody>${tbody}</tbody></table>
            </div>
        </div>`;
    },

    _st02Progress(prefix, label, pct) {
        const box  = document.getElementById(prefix + '-progress');
        const lab  = document.getElementById(prefix + '-progress-label');
        const p    = document.getElementById(prefix + '-progress-pct');
        const fill = document.getElementById(prefix + '-progress-fill');
        if (box) box.style.display = 'block';
        if (lab) lab.textContent = label;
        if (p)   p.textContent   = `${Math.round(pct)}%`;
        if (fill) fill.style.width = `${Math.round(pct)}%`;
    },

    /* Persist lightweight ST02 metadata into the module data (merged, not
       overwriting other keys) so it is visible to downstream code / debugging.
       The authoritative session state lives in getSessionState(). */
    _saveSt02ToApp() {
        if (!this.app) return;
        const prev = this.app.getModuleData('tech-workload-app') || {};
        this.app.setModuleData('tech-workload-app', {
            ...prev,
            st02: {
                bufferFileCount: (this.st02BufFiles || []).length,
                bufferTotal: this.st02BufAnalysis ? this.st02BufAnalysis.total : null,
                wpFileCount: (this.st02WpFiles || []).length,
                wpTotalUser: this.st02WpAnalysis ? this.st02WpAnalysis.totalUser : null,
                wpTotalExec: this.st02WpAnalysis ? this.st02WpAnalysis.totalExec : null,
            },
        });
    },

    /* Modal listing valid / invalid files. onProceed (if provided) runs when the
       user clicks "Process valid only"; "Cancel"/backdrop closes without action. */
    _showFileValidationModal(title, valid, invalid, onProceed) {
        const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
        const back = document.createElement('div');
        back.className = 'st02-validate-backdrop';
        const validList = valid.length
            ? `<ul>${valid.map(f => `<li class="st02-validate-ok">✓ ${esc(f.name)}</li>`).join('')}</ul>`
            : `<p style="color:var(--text-muted);font-size:13px;margin:0 0 14px">No valid files.</p>`;
        const invalidList = invalid.length
            ? `<ul>${invalid.map(f => `<li class="st02-validate-bad">✕ ${esc(f.name)} — <span>${esc(f.reason)}</span></li>`).join('')}</ul>`
            : '';
        back.innerHTML = `
            <div class="st02-validate-card" role="dialog" aria-modal="true">
                <h3>${esc(title)}</h3>
                <div style="font-size:12px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:4px">Valid files (${valid.length})</div>
                ${validList}
                <div style="font-size:12px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:4px">Invalid files (${invalid.length})</div>
                ${invalidList}
                <div style="display:flex;gap:10px;justify-content:flex-end;margin-top:8px">
                    <button class="btn" id="twa-st02-modal-cancel">Cancel</button>
                    <button class="btn btn-blue" id="twa-st02-modal-ok" ${valid.length ? '' : 'disabled style="opacity:.5;cursor:not-allowed"'}>Process valid only</button>
                </div>
            </div>`;
        const close = () => { if (back.parentNode) back.parentNode.removeChild(back); };
        back.addEventListener('click', e => { if (e.target === back) close(); });
        document.body.appendChild(back);
        back.querySelector('#twa-st02-modal-cancel').onclick = close;
        const ok = back.querySelector('#twa-st02-modal-ok');
        if (ok) ok.onclick = () => { close(); if (valid.length && onProceed) onProceed(); };
    },

    _wireNmonZone(app) {
        const zone  = document.getElementById('twa-nmon-zone');
        const input = document.getElementById('twa-nmon-input');
        if (!zone || !input || zone._wired) return;
        zone._wired = true;
        zone.addEventListener('click', () => input.click());
        zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
        zone.addEventListener('drop', e => {
            e.preventDefault(); zone.classList.remove('drag-over');
            this._handleNmonFiles(Array.from(e.dataTransfer.files), app);
        });
        input.addEventListener('change', e => {
            this._handleNmonFiles(Array.from(e.target.files), app);
        });
    },

    _handleNmonFiles(files, app) {
        if (!files || !files.length) return;
        const ALLOWED = /\.(nmon|txt|xlsx|xls|zip)$/i;
        const accepted = files.filter(f => ALLOWED.test(f.name));
        if (!accepted.length) {
            showToast('Unsupported file', 'NMON step accepts .nmon, .txt, .xlsx or .zip files.', 'error');
            return;
        }
        const existing = new Set((this.nmonFiles || []).map(f => f.name));
        accepted.forEach(f => {
            if (!existing.has(f.name)) {
                (this.nmonFiles = this.nmonFiles || []).push(f);
                existing.add(f.name);
            }
        });
        this._renderNmonFileList();
        this._saveNmonToApp();
        /* Process automatically on upload. */
        this._processNmon(app);
    },

    _renderNmonFileList() {
        const list = document.getElementById('twa-nmon-file-list');
        const ul   = document.getElementById('twa-nmon-files-ul');
        if (!list || !ul) return;
        const files = this.nmonFiles || [];
        if (!files.length) { list.style.display = 'none'; return; }
        list.style.display = 'block';
        ul.innerHTML = files.map((f, i) => {
            const kb = f.size ? ` (${(f.size / 1024).toFixed(0)} KB)` : '';
            return `<li style="display:flex;align-items:center;gap:8px;padding:3px 0;font-size:13px">
                <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24"
                     fill="none" stroke="var(--brand)" stroke-width="2.2">
                    <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                    <polyline points="22 4 12 14.01 9 11.01"/>
                </svg>
                <span>${f.name}${kb}</span>
                <button onclick="TechWorkloadAppModule._removeNmonFile(${i})"
                        style="background:none;border:none;cursor:pointer;color:var(--error);
                               font-size:15px;line-height:1;padding:0 4px;margin-left:auto"
                        title="Remove">×</button>
            </li>`;
        }).join('');
    },

    _removeNmonFile(idx) {
        if (!this.nmonFiles) return;
        this.nmonFiles.splice(idx, 1);
        this._renderNmonFileList();
        this._saveNmonToApp();
        if (!this.nmonFiles.length) {
            this.nmonAnalysis = null;
            const sum = document.getElementById('twa-nmon-summary');
            if (sum) { sum.style.display = 'none'; sum.innerHTML = ''; }
            this._saveNmonToApp();
        }
    },

    _nmonProgress(label, pct) {
        const box = document.getElementById('twa-nmon-progress');
        const lab = document.getElementById('twa-nmon-progress-label');
        const p   = document.getElementById('twa-nmon-progress-pct');
        const fill= document.getElementById('twa-nmon-progress-fill');
        if (box) box.style.display = 'block';
        if (lab) lab.textContent = label;
        if (p)   p.textContent   = `${Math.round(pct)}%`;
        if (fill)fill.style.width = `${Math.round(pct)}%`;
    },
    _hideNmonProgress() {
        const box = document.getElementById('twa-nmon-progress');
        if (box) box.style.display = 'none';
    },

    /* Parse all queued NMON files and build the consolidated analysis.
       Browser-side implementation mirroring the fields the processor needs for
       later charting: per-timestamp CPU busy, LPAR physical/entitled/virtual
       CPU, memory, paging, disk and network. */
    async _processNmon(app) {
        const files = this.nmonFiles || [];
        if (!files.length) return;
        this._nmonProgress('Reading NMON files…', 5);

        try {
            /* Expand any .zip bundles into individual .nmon/.txt members. */
            const expanded = [];
            for (const f of files) {
                if (/\.zip$/i.test(f.name) && typeof JSZip !== 'undefined') {
                    const zip = await JSZip.loadAsync(f);
                    const names = Object.keys(zip.files).filter(n => /\.(nmon|txt)$/i.test(n) && !zip.files[n].dir);
                    for (const n of names) {
                        const text = await zip.files[n].async('string');
                        expanded.push({ name: n.split('/').pop(), text });
                    }
                } else {
                    const text = await f.text();
                    expanded.push({ name: f.name, text });
                }
            }

            if (!expanded.length) {
                this._hideNmonProgress();
                showToast('No NMON data', 'No .nmon/.txt members were found in the uploaded files.', 'error');
                return;
            }

            /* Parse each file. */
            const parsed = [];
            for (let i = 0; i < expanded.length; i++) {
                this._nmonProgress(`Parsing ${expanded[i].name}…`, 10 + (i / expanded.length) * 70);
                parsed.push(this._parseNmonText(expanded[i].name, expanded[i].text));
                await new Promise(r => setTimeout(r, 0));   /* yield to keep UI responsive */
            }

            this._nmonProgress('Building workbook…', 88);
            const sids = [...new Set(parsed.map(p => p.sid).filter(Boolean))];
            const sidPrefix = sids.length === 1 ? sids[0] : (sids.length > 1 ? 'MULTI' : 'NMON');
            const fileName = `${sidPrefix}_NMON_Analysis.xlsx`;

            const wbInfo = this._buildNmonWorkbook(parsed, fileName);

            this.nmonAnalysis = {
                fileName,
                hosts:    [...new Set(parsed.map(p => p.hostname).filter(Boolean))],
                files:    expanded.length,
                records:  parsed.reduce((s, p) => s + p.cpuRaw.length, 0),
                sids,
                workbook: wbInfo.wb,
            };
            this._saveNmonToApp();

            this._nmonProgress('Analysis ready', 100);
            setTimeout(() => this._hideNmonProgress(), 600);
            this._renderNmonSummary(fileName);
        } catch (err) {
            this._hideNmonProgress();
            showToast('NMON processing failed', String(err && err.message || err), 'error');
        }
    },

    /* Minimal NMON text parser — extracts the fields required for the workbook
       and for later chart generation. Returns a per-host record bundle. */
    _parseNmonText(name, text) {
        const lines = String(text).split(/\r?\n/);
        const rec = {
            file: name, hostname: '', sid: '', lpar: '', os: '', osVer: '',
            machineType: '', model: '', procModel: '',
            physCpus: null, virtCpus: null, entitled: null, memGB: null, virt: '',
            tsStart: '', tsEnd: '',
            timestamps: {},       /* snapId → "DD-MON-YYYY HH:MM:SS" */
            cpuRaw: [], lparRaw: [], memRaw: [], pagingRaw: [], diskRaw: [], netRaw: [],
            warnings: [],
        };

        const meta = {};
        const sections = {};
        for (const raw of lines) {
            if (!raw.trim()) continue;
            const parts = raw.split(',');
            const tag = (parts[0] || '').trim().toUpperCase();
            if (tag === 'AAA' && parts.length >= 3) {
                meta[parts[1].trim().toUpperCase()] = parts.slice(2).join(',').trim();
            } else if (tag === 'BBB' && parts.length >= 3) {
                meta[parts[1].trim().toUpperCase()] = parts.slice(2).join(',').trim();
            } else if (tag === 'ZZZZ' && parts.length >= 3) {
                const snap = parts[1].trim();
                rec.timestamps[snap] = parts.length >= 4
                    ? `${parts[3].trim()} ${parts[2].trim()}`
                    : parts[2].trim();
            } else {
                (sections[tag] = sections[tag] || []).push(parts);
            }
        }

        rec.hostname = (meta['HOST'] || meta['HOSTNAME'] || meta['NODENAME'] || '').toUpperCase();
        rec.os       = meta['OS'] || meta['OSNAME'] || 'AIX';
        rec.osVer    = meta['VER'] || meta['OSVERSION'] || '';
        rec.machineType = meta['MACHINE'] || meta['MACHINETYPE'] || '';
        rec.model    = meta['MODEL'] || '';
        rec.procModel= meta['PROCESSOR'] || meta['PROCESSORTYPE'] || '';
        rec.lpar     = meta['LPAR'] || meta['LPARNAME'] || rec.hostname;
        rec.virt     = meta['VIRTUALIZATION'] || (/(AIX)/i.test(rec.os) ? 'PowerVM LPAR' : '');
        const pInt = v => { const n = parseInt(v, 10); return isNaN(n) ? null : n; };
        const pFlt = v => { const n = parseFloat(v); return isNaN(n) ? null : n; };
        rec.physCpus = pInt(meta['PHYSICALCPU'] || meta['PHYSICALCPUS'] || meta['CPUS']);
        rec.virtCpus = pInt(meta['LOGICALCPU'] || meta['LOGICALCPUS'] || meta['VIRTUALCPUS'] || meta['VCPUS']);
        rec.entitled = pFlt(meta['ENTITLED'] || meta['ENTITLEDCAPACITY'] || meta['ENTITLEDCPU']);
        const memMb  = pFlt(meta['MEMORY'] || meta['TOTALMEM'] || meta['REALMEM']);
        if (memMb != null) rec.memGB = Math.round(memMb / 1024 * 100) / 100;

        /* SID: 3-char token from hostname / lpar / metadata. */
        rec.sid = this._nmonSid(meta['SID'] || meta['SAPSID'] || '') ||
                  this._nmonSid(rec.hostname) || this._nmonSid(rec.lpar) || '';

        const tsVals = Object.values(rec.timestamps);
        if (tsVals.length) { rec.tsStart = tsVals[0]; rec.tsEnd = tsVals[tsVals.length - 1]; }

        /* Header+row helper for a section tag. */
        const readSection = (tag) => {
            const rows = sections[tag] || [];
            let header = null; const out = [];
            for (const r of rows) {
                if ((r[1] || '').trim().toUpperCase() === tag) {
                    header = r.slice(2).map(x => x.trim().toUpperCase());
                } else if (header) {
                    out.push({ snap: (r[1] || '').trim(), vals: r.slice(2) });
                }
            }
            return { header, out };
        };

        /* LPAR capacity (entitled / virtual / physical). */
        const lpar = readSection('LPAR');
        const lparBySnap = {};
        if (lpar.header) {
            for (const row of lpar.out) {
                const d = {}; lpar.header.forEach((h, k) => d[h] = row.vals[k]);
                const ent  = pFlt(d['ENTITLED'] || d['ENTITLEDCAPACITY']);
                const virt = pFlt(d['VIRTUALCPUS'] || d['VCPUS']);
                const phys = pFlt(d['PHYSICALCPU'] || d['PHYSICALCPUS']);
                lparBySnap[row.snap] = { ent, virt, phys };
                if (rec.entitled == null && ent != null) rec.entitled = ent;
                if (rec.virtCpus == null && virt != null) rec.virtCpus = Math.round(virt);
            }
        }

        /* CPU — try CPU_ALL, then CPU_TOTAL, then PROC. */
        let cpu = null;
        for (const tag of ['CPU_ALL', 'CPU_TOTAL', 'PROC']) {
            const s = readSection(tag);
            if (s.header && s.out.length) { cpu = s; break; }
        }
        if (cpu) {
            for (const row of cpu.out) {
                const d = {}; cpu.header.forEach((h, k) => d[h] = row.vals[k]);
                const get = (...keys) => { for (const k of keys) { const v = pFlt(d[k]); if (v != null) return v; } return null; };
                const user = get('USER%', 'USER', 'USR%');
                const sys  = get('SYS%', 'SYS', 'SYSTEM%');
                const wait = get('WAIT%', 'WAIT', 'WAI%');
                let idle   = get('IDLE%', 'IDLE', 'IDL%');
                if (idle == null && user != null) idle = Math.max(0, 100 - (user || 0) - (sys || 0) - (wait || 0));
                const busy = idle != null ? Math.round((100 - idle) * 100) / 100 : null;
                const lp = lparBySnap[row.snap] || {};
                const ent = lp.ent != null ? lp.ent : rec.entitled;
                const virt = lp.virt != null ? lp.virt : rec.virtCpus;
                let phys = lp.phys;
                if (phys == null && busy != null && ent != null) phys = Math.round(busy * ent / 100 * 10000) / 10000;
                const ts = rec.timestamps[row.snap] || '';
                rec.cpuRaw.push({
                    Timestamp: ts, SID: rec.sid, Hostname: rec.hostname,
                    User_Pct: user || 0, System_Pct: sys || 0, Wait_Pct: wait || 0,
                    Idle_Pct: idle || 0, Busy_Pct: busy || 0,
                    LPAR_Physical_CPU: phys != null ? phys : '',
                    LPAR_Entitled_CPU: ent != null ? ent : '',
                    LPAR_Virtual_CPUs: virt != null ? virt : '',
                });
                rec.lparRaw.push({
                    Timestamp: ts, Hostname: rec.hostname,
                    LPAR_Physical_CPU: phys != null ? phys : '',
                    LPAR_Entitled_CPU: ent != null ? ent : '',
                    LPAR_Virtual_CPUs: virt != null ? virt : '',
                });
            }
        } else {
            rec.warnings.push({ Hostname: rec.hostname, Section: 'CPU', Issue: 'No CPU_ALL/CPU_TOTAL/PROC data', Severity: 'Warning' });
        }

        /* Memory. */
        const mem = readSection('MEM');
        if (mem.header) {
            for (const row of mem.out) {
                const d = {}; mem.header.forEach((h, k) => d[h] = row.vals[k]);
                const get = (...keys) => { for (const k of keys) { const v = pFlt(d[k]); if (v != null) return v; } return null; };
                const realMb = get('REALMEM(MB)', 'REALMEM', 'TOTALMEM', 'REAL_MB');
                const freeMb = get('FREEMEM(MB)', 'FREEMEM', 'FREE_MB');
                let usedMb = (realMb != null && freeMb != null) ? Math.max(0, realMb - freeMb) : null;
                const ts = rec.timestamps[row.snap] || '';
                rec.memRaw.push({
                    Timestamp: ts, Hostname: rec.hostname,
                    Total_Memory_GB: realMb != null ? Math.round(realMb / 1024 * 1000) / 1000 : '',
                    Used_Memory_GB:  usedMb != null ? Math.round(usedMb / 1024 * 1000) / 1000 : '',
                    Free_Memory_GB:  freeMb != null ? Math.round(freeMb / 1024 * 1000) / 1000 : '',
                    Used_Memory_Pct: (realMb && usedMb != null) ? Math.round(usedMb / realMb * 10000) / 100 : '',
                });
            }
        } else {
            rec.warnings.push({ Hostname: rec.hostname, Section: 'MEM', Issue: 'No MEM data', Severity: 'Warning' });
        }

        /* Paging. */
        const pg = readSection('PAGING');
        if (pg.header) {
            for (const row of pg.out) {
                const d = {}; pg.header.forEach((h, k) => d[h] = row.vals[k]);
                const get = (...keys) => { for (const k of keys) { const v = pFlt(d[k]); if (v != null) return v; } return null; };
                rec.pagingRaw.push({
                    Timestamp: rec.timestamps[row.snap] || '', Hostname: rec.hostname,
                    PageIn: get('PAGEIN', 'PGIN', 'PAGEINS') || 0,
                    PageOut: get('PAGEOUT', 'PGOUT', 'PAGEOUTS') || 0,
                    PageRate: get('PAGINGRATE', 'PAGING', 'PGRATE') || 0,
                });
            }
        }

        /* Disk busy / read / write / xfer. */
        const diskBusy  = readSection('DISKBUSY');
        const diskRead  = readSection('DISKREAD');
        const diskWrite = readSection('DISKWRITE');
        const diskXfer  = readSection('DISKXFER');
        if (diskBusy.header) {
            const combo = {};
            const fill = (sec, field) => {
                if (!sec.header) return;
                for (const row of sec.out) {
                    const ts = rec.timestamps[row.snap] || '';
                    sec.header.forEach((disk, k) => {
                        const key = row.snap + '|' + disk;
                        (combo[key] = combo[key] || { Timestamp: ts, Hostname: rec.hostname, Disk_Name: disk,
                            Busy_Pct: '', Read_KB_Sec: '', Write_KB_Sec: '', Transfers_Sec: '' });
                        const v = pFlt(row.vals[k]);
                        if (v != null) combo[key][field] = v;
                    });
                }
            };
            fill(diskBusy,  'Busy_Pct');
            fill(diskRead,  'Read_KB_Sec');
            fill(diskWrite, 'Write_KB_Sec');
            fill(diskXfer,  'Transfers_Sec');
            rec.diskRaw = Object.values(combo);
        }

        /* Network. */
        const net = readSection('NET');
        if (net.header) {
            const combo = {};
            for (const row of net.out) {
                const ts = rec.timestamps[row.snap] || '';
                net.header.forEach((col, k) => {
                    const m = /(.+?)-(read|write)/i.exec(col);
                    if (!m) return;
                    const iface = m[1], dir = m[2].toLowerCase();
                    const key = row.snap + '|' + iface;
                    (combo[key] = combo[key] || { Timestamp: ts, Hostname: rec.hostname, Interface: iface, Read_KB_Sec: '', Write_KB_Sec: '' });
                    const v = pFlt(row.vals[k]);
                    if (v != null) combo[key][dir === 'read' ? 'Read_KB_Sec' : 'Write_KB_Sec'] = v;
                });
            }
            rec.netRaw = Object.values(combo);
        }

        return rec;
    },

    _nmonSid(textRaw) {
        const text = String(textRaw || '').toUpperCase();
        const BAD = new Set(['IBM','CPU','MEM','NET','SMP','AIX','SAP','TMP','VAR','OPT','USR','SYS','LOG','RUN','ETH','ENT','HDK','LVM']);
        const pats = [/\bSID[=:\s]+([A-Z][A-Z0-9]{2})\b/, /[_\-/]([A-Z][A-Z0-9]{2})[_\-/]/, /^([A-Z][A-Z0-9]{2})[_\-]/];
        for (const p of pats) {
            const m = p.exec(text);
            if (m && !BAD.has(m[1])) return m[1];
        }
        return '';
    },

    _pctile(arr, p) {
        const a = arr.filter(v => typeof v === 'number' && !isNaN(v)).sort((x, y) => x - y);
        if (!a.length) return '';
        const idx = Math.max(0, Math.ceil(p / 100 * a.length) - 1);
        return Math.round(a[idx] * 100) / 100;
    },

    /* Build a consolidated SheetJS workbook from parsed NMON records. */
    _buildNmonWorkbook(parsed, fileName) {
        const wb = XLSX.utils.book_new();
        const addSheet = (name, rows, cols) => {
            const data = rows && rows.length ? rows : [Object.fromEntries((cols || ['Info']).map(c => [c, '']))];
            const ws = XLSX.utils.json_to_sheet(data, cols ? { header: cols } : undefined);
            XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
        };

        const num = v => (typeof v === 'number' && !isNaN(v)) ? v : null;
        const mean = a => a.length ? Math.round(a.reduce((s, v) => s + v, 0) / a.length * 100) / 100 : '';

        /* Per-host aggregations. */
        const cpuSummary = [], memSummary = [], lparCap = [], kpi = [], exec = [], inv = [], warnings = [];
        const cpuRawAll = [], memRawAll = [], diskRawAll = [], netRawAll = [], lparRawAll = [];

        for (const p of parsed) {
            cpuRawAll.push(...p.cpuRaw);
            memRawAll.push(...p.memRaw);
            diskRawAll.push(...p.diskRaw);
            netRawAll.push(...p.netRaw);
            lparRawAll.push(...p.lparRaw);
            warnings.push(...p.warnings);

            inv.push({
                SID: p.sid, Hostname: p.hostname, LPAR: p.lpar,
                Machine_Type: p.machineType, Processor_Model: p.procModel || p.model,
                Operating_System: p.os, OS_Version: p.osVer,
                Physical_CPUs: p.physCpus != null ? p.physCpus : '',
                Virtual_CPUs: p.virtCpus != null ? p.virtCpus : '',
                Entitled_CPUs: p.entitled != null ? p.entitled : '',
                Installed_Memory_GB: p.memGB != null ? p.memGB : '',
                Virtualization: p.virt,
                Collection_Start: p.tsStart, Collection_End: p.tsEnd,
                Source_File: p.file,
            });

            const busy = p.cpuRaw.map(r => num(r.Busy_Pct)).filter(v => v != null);
            const phys = p.lparRaw.map(r => num(r.LPAR_Physical_CPU)).filter(v => v != null);
            const memP = p.memRaw.map(r => num(r.Used_Memory_Pct)).filter(v => v != null);
            const pageOut = p.pagingRaw.reduce((s, r) => s + (num(r.PageOut) || 0), 0);

            let peakTs = '';
            if (busy.length) { const mx = Math.max(...busy); const row = p.cpuRaw.find(r => num(r.Busy_Pct) === mx); peakTs = row ? row.Timestamp : ''; }

            cpuSummary.push({
                SID: p.sid, Hostname: p.hostname,
                CPU_Avg_Pct: mean(busy), CPU_Min_Pct: busy.length ? Math.min(...busy) : '',
                CPU_Max_Pct: busy.length ? Math.max(...busy) : '',
                CPU_P95_Pct: this._pctile(busy, 95), CPU_P99_Pct: this._pctile(busy, 99),
                Peak_Timestamp: peakTs,
            });
            memSummary.push({
                Hostname: p.hostname, Avg_Memory_Pct: mean(memP),
                Peak_Memory_Pct: memP.length ? Math.max(...memP) : '',
                P95_Memory_Pct: this._pctile(memP, 95),
                Total_Page_In: p.pagingRaw.reduce((s, r) => s + (num(r.PageIn) || 0), 0),
                Total_Page_Out: pageOut,
            });

            const disksByName = {};
            p.diskRaw.forEach(r => { (disksByName[r.Disk_Name] = disksByName[r.Disk_Name] || []).push(r); });
            let diskPeak = '';
            Object.values(disksByName).forEach(rows => {
                const bp = rows.map(r => num(r.Busy_Pct)).filter(v => v != null);
                if (bp.length) diskPeak = Math.max(diskPeak === '' ? 0 : diskPeak, Math.max(...bp));
            });
            const netPeak = (() => {
                const t = p.netRaw.map(r => (num(r.Read_KB_Sec) || 0) + (num(r.Write_KB_Sec) || 0));
                return t.length ? Math.round(Math.max(...t) * 100) / 100 : '';
            })();

            kpi.push({
                Hostname: p.hostname,
                CPU_Avg_Pct: mean(busy), CPU_P95_Pct: this._pctile(busy, 95), CPU_Max_Pct: busy.length ? Math.max(...busy) : '',
                Physical_CPU_Avg: mean(phys), Physical_CPU_P95: this._pctile(phys, 95), Physical_CPU_Max: phys.length ? Math.max(...phys) : '',
                Memory_Avg_Pct: mean(memP), Memory_P95_Pct: this._pctile(memP, 95), Memory_Max_Pct: memP.length ? Math.max(...memP) : '',
                Disk_Max_Busy_Pct: diskPeak, Network_Max_KB_Sec: netPeak, Total_Page_Out: pageOut,
                Collection_Start: p.tsStart, Collection_End: p.tsEnd,
            });

            const comments = [];
            const cpuP95 = this._pctile(busy, 95), memP95 = this._pctile(memP, 95);
            if (typeof cpuP95 === 'number' && cpuP95 > 80) comments.push('Review CPU sizing');
            if (typeof memP95 === 'number' && memP95 > 85) comments.push('Memory pressure detected');
            if (typeof diskPeak === 'number' && diskPeak > 70) comments.push('Storage bottleneck risk');
            if (pageOut > 0) comments.push('Investigate memory pressure (paging detected)');
            if (!comments.length) comments.push('No significant bottlenecks detected');

            exec.push({
                SID: p.sid, Hostname: p.hostname,
                Virtual_CPUs: p.virtCpus != null ? p.virtCpus : '', Entitled_CPUs: p.entitled != null ? p.entitled : '',
                CPU_Avg_Pct: mean(busy), CPU_P95_Pct: cpuP95, CPU_Max_Pct: busy.length ? Math.max(...busy) : '',
                Physical_CPU_Avg: mean(phys), Physical_CPU_P95: this._pctile(phys, 95), Physical_CPU_Max: phys.length ? Math.max(...phys) : '',
                Memory_Avg_Pct: mean(memP), Memory_P95_Pct: memP95, Memory_Max_Pct: memP.length ? Math.max(...memP) : '',
                Highest_Disk_Busy_Pct: diskPeak, Highest_Network_KB_Sec: netPeak, Total_Page_Out: pageOut,
                Comments: comments.join(' | '),
            });

            /* LPAR capacity summary per host. */
            lparCap.push({
                Hostname: p.hostname,
                Physical_CPU_Avg: mean(phys), Physical_CPU_Max: phys.length ? Math.max(...phys) : '',
                Physical_CPU_P95: this._pctile(phys, 95),
                Entitled_CPUs: p.entitled != null ? p.entitled : '', Virtual_CPUs: p.virtCpus != null ? p.virtCpus : '',
            });
        }

        /* Disk & network per-host-per-device summaries. */
        const diskSummary = [];
        const dByKey = {};
        diskRawAll.forEach(r => { (dByKey[r.Hostname + '|' + r.Disk_Name] = dByKey[r.Hostname + '|' + r.Disk_Name] || []).push(r); });
        Object.entries(dByKey).forEach(([k, rows]) => {
            const [host, disk] = k.split('|');
            const bp = rows.map(r => num(r.Busy_Pct)).filter(v => v != null);
            const rd = rows.map(r => num(r.Read_KB_Sec)).filter(v => v != null);
            const wr = rows.map(r => num(r.Write_KB_Sec)).filter(v => v != null);
            diskSummary.push({
                Hostname: host, Disk_Name: disk,
                Avg_Busy_Pct: mean(bp), Peak_Busy_Pct: bp.length ? Math.max(...bp) : '', P95_Busy_Pct: this._pctile(bp, 95),
                Avg_Read_KB_Sec: mean(rd), Avg_Write_KB_Sec: mean(wr),
            });
        });

        const netSummary = [];
        const nByKey = {};
        netRawAll.forEach(r => { (nByKey[r.Hostname + '|' + r.Interface] = nByKey[r.Hostname + '|' + r.Interface] || []).push(r); });
        Object.entries(nByKey).forEach(([k, rows]) => {
            const [host, iface] = k.split('|');
            const tot = rows.map(r => (num(r.Read_KB_Sec) || 0) + (num(r.Write_KB_Sec) || 0));
            netSummary.push({
                Hostname: host, Interface: iface,
                Avg_Total_KB_Sec: mean(tot), Peak_Total_KB_Sec: tot.length ? Math.round(Math.max(...tot) * 100) / 100 : '',
                P95_Total_KB_Sec: this._pctile(tot, 95),
            });
        });

        addSheet('Executive_Summary', exec);
        addSheet('System_Inventory',  inv);
        addSheet('CPU_Raw',           cpuRawAll, ['Timestamp','SID','Hostname','User_Pct','System_Pct','Wait_Pct','Idle_Pct','Busy_Pct','LPAR_Physical_CPU','LPAR_Entitled_CPU','LPAR_Virtual_CPUs']);
        addSheet('CPU_Summary',       cpuSummary);
        addSheet('LPAR_Capacity',     lparRawAll, ['Timestamp','Hostname','LPAR_Physical_CPU','LPAR_Entitled_CPU','LPAR_Virtual_CPUs']);
        addSheet('LPAR_Summary',      lparCap);
        addSheet('Memory_Raw',        memRawAll, ['Timestamp','Hostname','Total_Memory_GB','Used_Memory_GB','Free_Memory_GB','Used_Memory_Pct']);
        addSheet('Memory_Summary',    memSummary);
        addSheet('Disk_Raw',          diskRawAll, ['Timestamp','Hostname','Disk_Name','Busy_Pct','Read_KB_Sec','Write_KB_Sec','Transfers_Sec']);
        addSheet('Disk_Summary',      diskSummary);
        addSheet('Network_Raw',       netRawAll, ['Timestamp','Hostname','Interface','Read_KB_Sec','Write_KB_Sec']);
        addSheet('Network_Summary',   netSummary);
        addSheet('KPI_Summary',       kpi);
        addSheet('Warnings',          warnings, ['Hostname','Section','Issue','Severity']);

        return { wb };
    },

    _renderNmonSummary(fileName) {
        const box = document.getElementById('twa-nmon-summary');
        if (!box || !this.nmonAnalysis) return;
        const a = this.nmonAnalysis;
        box.style.display = 'block';
        box.innerHTML = `
            <div class="section-title" style="margin-top:24px">NMON Analysis Summary</div>
            <div class="stats-grid">
                ${statCard('Files Processed', a.files)}
                ${statCard('Servers Detected', a.hosts.length)}
                ${statCard('CPU Records', a.records.toLocaleString())}
                ${statCard('SIDs', a.sids.length ? a.sids.join(', ') : '—')}
            </div>
            <div class="card" style="margin-top:14px">
                <div class="card-header" style="padding:14px 18px 12px">
                    <div>
                        <div class="card-title">Consolidated Workbook</div>
                        <div class="card-sub">${fileName} — includes LPAR Physical / Entitled / Virtual CPU metrics for chart generation</div>
                    </div>
                    <button class="btn btn-blue" id="twa-nmon-download">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="7 10 12 15 17 10"/>
                            <line x1="12" y1="15" x2="12" y2="3"/>
                        </svg>
                        Download ${fileName}
                    </button>
                </div>
            </div>`;
        document.getElementById('twa-nmon-download')?.addEventListener('click', () => {
            if (this.nmonAnalysis && this.nmonAnalysis.workbook) {
                XLSX.writeFile(this.nmonAnalysis.workbook, this.nmonAnalysis.fileName);
            }
        });
    },

    _saveNmonToApp() {
        if (!this.app) return;
        const prev = this.app.getModuleData('tech-workload-app') || {};
        this.app.setModuleData('tech-workload-app', {
            ...prev,
            nmon: {
                fileCount: (this.nmonFiles || []).length,
                fileNames: (this.nmonFiles || []).map(f => f.name),
                analysis:  this.nmonAnalysis ? {
                    fileName: this.nmonAnalysis.fileName,
                    hosts:    this.nmonAnalysis.hosts,
                    files:    this.nmonAnalysis.files,
                    records:  this.nmonAnalysis.records,
                    sids:     this.nmonAnalysis.sids,
                } : null,
            },
        });
    },

    /* Populate the dynamic left-sidebar navigation for this module as STEPS:
         Step 1 — Upload SMON/MON data (the enriched table lives here)
         Step 2 — Chart visualization (CPU / Memory / Work Process / Logins tabs)
       Chart sub-items activate the matching lazy-rendered tab. */
    _updateSubNav(app, activeId) {
        if (!app || !app.setSubNav) return;
        const hasCharts = !!(this.targetRows && this.targetRows.length);
        /* Switch to the charts view (if needed) and activate the requested tab —
           this lazily builds only that section's charts. */
        const gotoTab = (tab) => {
            if (!hasCharts) return;
            this._activateSection(app, tab);
        };

        const items = [
            { id: 'upload', label: 'Step 1 — Upload SMON / MON data',
              onClick: () => { this._showTableView(); document.getElementById('content-area')?.scrollTo({ top: 0, behavior: 'smooth' }); app.setSubNavActive('tech-workload-app', 'upload'); },
              step: 1 },
            { id: 'profiles', label: 'Upload SAP System Profiles (optional)',
              onClick: () => this._showProfilesView(), sub: true },
        ];
        if (hasCharts) {
            items.push(
                { id: 'charts', label: 'Step 2 — Chart visualization', onClick: () => gotoTab('cpu'), step: 2 },
                { id: 'cpu',    label: 'CPU Utilization', onClick: () => gotoTab('cpu'), sub: true },
                { id: 'mem',    label: 'Memory Utilization', onClick: () => gotoTab('mem'), sub: true },
                { id: 'wp',     label: 'Work Process', onClick: () => gotoTab('wp'), sub: true },
                { id: 'logins', label: 'Logins & Sessions', onClick: () => gotoTab('logins'), sub: true },
            );
            if (this.sourceTable === 'SMON') {
                items.push({ id: 'network', label: 'Network', onClick: () => gotoTab('network'), sub: true });
            }
            /* Optional Step 3 (ST02 Memory Buffer) + Step 4 (ST02 WP Storage).
               Shown once the enriched table exists, consistent with the charts step. */
            items.push(
                { id: 'st02buf', label: 'Step 3 — Memory Buffer Analysis (optional)',
                  onClick: () => this._showSt02BufView(app), step: 3 },
                { id: 'st02wp', label: 'Step 4 — User Storage for WP Memory (optional)',
                  onClick: () => this._showSt02WpView(app), step: 4 },
            );
        }
        /* Optional IBM NMON step (Step 5) — only when the landscape has IBM servers. */
        if (this._ibmDetected()) {
            items.push({
                id: 'nmon',
                label: `Step ${hasCharts ? 5 : 2} — IBM NMON (optional)`,
                onClick: () => { this._showNmonView(app); },
                step: hasCharts ? 5 : 2,
            });
        }
        /* map legacy activeId values onto the stepped ids */
        let active = activeId;
        if (activeId === 'table') active = 'upload';
        app.setSubNav('tech-workload-app', items, active);
    },
    _subActiveFromSel(sel) {
        if (sel.includes('mem')) return 'mem';
        if (sel.includes('wp')) return 'wp';
        if (sel.includes('logins')) return 'logins';
        return 'cpu';
    },

    /* Average of the N highest RAW values of a column across application-tier
       rows over the whole analyzed period. Used by the EM / Heap peak charts,
       whose "Average of the top N peaks" statistic is defined on the N highest
       individual readings in the period (not on the ≤20 displayed bars).
       Returns { avg, count } (count = how many values were actually averaged,
       capped at n) or null when the column/data is unavailable. */
    _avgTopRawByCol(app, colName, n) {
        const ci = this._colIndex || {};
        const iSrv = ci['SERVER'], iEwa = ci['SERVER (EWA)'], iCol = ci[colName];
        if (iSrv == null || iCol == null) return null;
        const rows = this.targetRows || [];
        const isApp = {};
        const vals = [];
        for (let i = 0; i < rows.length; i++) {
            const r = rows[i]; const srv = r[iSrv];
            if (!srv) continue;
            if (isApp[srv] === undefined) isApp[srv] = this._rowLayer(app, srv, iEwa != null ? r[iEwa] : null) === 'app';
            if (!isApp[srv]) continue;
            const v = parseFloat(r[iCol]);
            if (!isNaN(v)) vals.push(v);
        }
        if (!vals.length) return null;
        vals.sort((a, b) => b - a);
        const top = vals.slice(0, Math.min(n, vals.length));
        const avg = top.reduce((s, v) => s + v, 0) / top.length;
        return { avg, count: top.length };
    },

    /* ══════════════════════════════════════════════════════
       CHARTS — CPU Utilization Analysis (application tier)
       ══════════════════════════════════════════════════════ */

    /* Resolve the landscape layer ('app' | 'db' | 'other') for a target row's
       server, honouring the virtual→physical mapping when it was used. */
    _rowLayer(app, serverVal, serverEwaVal) {
        const name = this._mappingActive
            ? (serverEwaVal && serverEwaVal !== 'No match in EWA report' ? serverEwaVal : null)
            : serverVal;
        if (!name) return 'other';
        const h = app.findHost(name);
        return h ? h.layer : 'other';
    },

    /* Build per-server, time-aligned series for the application-tier hosts in a
       single pass. Returns:
         { labels:["DD/MM/YYYY HH:MM",…],
           cpu:        { host:[y|null,…] },   // CPU Utilization (%)
           sapsUtil:   { host:[y|null,…] },   // SAPS Utilization (SAPS)
           sapsDep:    { host:[y|null,…] },   // SAPS Deployed (SAPS)
           hosts: [host,…] }                  // app-tier servers, in first-seen order
       Rows are already chronologically sorted. */
    _appTierSeries(app) {
        const ci = this._colIndex || {};
        const iDate = ci['Date'], iTime = ci['TIME'] != null ? ci['TIME'] : ci['Time'];
        const iSrv = ci['SERVER'], iEwa = ci['SERVER (EWA)'];
        const iCpu = ci['CPU Utilization (%)'];
        const iSU  = ci['SAPS Utilization (SAPS)'];
        const iSD  = ci['SAPS Deployed (SAPS)'];
        const empty = { labels: [], cpu: {}, sapsUtil: {}, sapsDep: {}, hosts: [] };
        if (iDate == null || iSrv == null) return empty;

        const labelIndex = new Map();
        const labels = [];
        const cpu = {}, sapsUtil = {}, sapsDep = {}, hosts = [];
        const isApp = {};
        const rows = this.targetRows;

        /* Pass 1 — ordered labels + app-tier host list. */
        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const srv = r[iSrv];
            if (!srv) continue;
            if (isApp[srv] === undefined) {
                isApp[srv] = this._rowLayer(app, srv, iEwa != null ? r[iEwa] : null) === 'app';
            }
            if (!isApp[srv]) continue;
            const lbl = `${r[iDate] || ''} ${iTime != null ? (r[iTime] || '') : ''}`.trim();
            if (!labelIndex.has(lbl)) { labelIndex.set(lbl, labels.length); labels.push(lbl); }
            if (cpu[srv] === undefined) { cpu[srv] = null; hosts.push(srv); }
        }
        const n = labels.length;
        hosts.forEach(h => {
            cpu[h] = new Array(n).fill(null);
            if (iSU != null) sapsUtil[h] = new Array(n).fill(null);
            if (iSD != null) sapsDep[h] = new Array(n).fill(null);
        });

        /* Pass 2 — fill values aligned to the label index. */
        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const srv = r[iSrv];
            if (!srv || !isApp[srv]) continue;
            const lbl = `${r[iDate] || ''} ${iTime != null ? (r[iTime] || '') : ''}`.trim();
            const pos = labelIndex.get(lbl);
            if (pos == null) continue;
            if (iCpu != null) { const y = parseFloat(r[iCpu]); if (!isNaN(y)) cpu[srv][pos] = y; }
            if (iSU  != null) { const y = parseFloat(r[iSU]);  if (!isNaN(y)) sapsUtil[srv][pos] = y; }
            if (iSD  != null) { const y = parseFloat(r[iSD]);  if (!isNaN(y)) sapsDep[srv][pos] = y; }
        }
        return { labels, cpu, sapsUtil, sapsDep, hosts };
    },

    /* Monthly hourly-average CPU per server: for each app-tier host, the mean of
       "CPU Utilization (%)" across ALL days, bucketed by hour of day (0-23).
       Returns { labels:["00".."23"], series:{host:[24]}, hosts:[...], peak }. */
    _hourlyAvgCpuByHour(app, hosts) {
        const ci = this._colIndex || {};
        const iHours = ci['Hours'], iSrv = ci['SERVER'], iCpu = ci['CPU Utilization (%)'];
        const labels = Array.from({ length: 24 }, (_, h) => String(h).padStart(2, '0'));
        if (iHours == null || iSrv == null || iCpu == null) return { labels, series: {}, hosts: [], peak: 0 };

        const hostSet = new Set(hosts);
        const sums = {}, cnts = {};
        hosts.forEach(h => { sums[h] = new Float64Array(24); cnts[h] = new Float64Array(24); });

        const rows = this.targetRows;
        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const srv = r[iSrv];
            if (!srv || !hostSet.has(srv)) continue;
            const hr = parseInt(r[iHours], 10);
            if (isNaN(hr) || hr < 0 || hr > 23) continue;
            const y = parseFloat(r[iCpu]);
            if (isNaN(y)) continue;
            sums[srv][hr] += y; cnts[srv][hr] += 1;
        }

        const series = {};
        const outHosts = [];
        let peak = 0;
        hosts.forEach(h => {
            const arr = new Array(24).fill(null);
            let any = false;
            for (let hr = 0; hr < 24; hr++) {
                if (cnts[h][hr] > 0) {
                    const v = Math.round((sums[h][hr] / cnts[h][hr]) * 100) / 100;
                    arr[hr] = v; any = true;
                    if (v > peak) peak = v;
                }
            }
            if (any) { series[h] = arr; outHosts.push(h); }
        });
        return { labels, series, hosts: outHosts, peak };
    },

    /* Produce a short load-balancing commentary for the monthly hourly-average
       CPU chart: compares each server's overall mean CPU% and flags imbalance. */
    _loadBalancingAnalysis(hourly) {
        const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        const means = hourly.hosts.map(h => {
            const arr = hourly.series[h].filter(v => v != null);
            const mean = arr.length ? arr.reduce((s, v) => s + v, 0) / arr.length : 0;
            return { host: h, mean };
        });
        if (!means.length) return '';
        means.sort((a, b) => b.mean - a.mean);
        const hi = means[0], lo = means[means.length - 1];
        const avgAll = means.reduce((s, m) => s + m.mean, 0) / means.length;
        /* Spread of per-server means relative to the overall average. */
        const spreadPct = avgAll > 0 ? ((hi.mean - lo.mean) / avgAll) * 100 : 0;

        let verdict, detail;
        if (means.length < 2) {
            verdict = 'Single application server';
            detail = `Only one application server (${esc(hi.host)}) is present, so load balancing cannot be assessed across peers.`;
        } else if (spreadPct <= 25) {
            verdict = 'Load balancing looks healthy';
            detail = `CPU is well distributed: the busiest server (${esc(hi.host)}, ${hi.mean.toFixed(1)}% avg) and the quietest (${esc(lo.host)}, ${lo.mean.toFixed(1)}% avg) differ by only ${spreadPct.toFixed(0)}% of the overall average — the per-server hourly curves track each other closely.`;
        } else if (spreadPct <= 60) {
            verdict = 'Load balancing could be improved';
            detail = `There is a noticeable imbalance: ${esc(hi.host)} averages ${hi.mean.toFixed(1)}% while ${esc(lo.host)} averages ${lo.mean.toFixed(1)}% (a ${spreadPct.toFixed(0)}% spread vs. the overall average of ${avgAll.toFixed(1)}%). Review logon-group / RFC-group distribution so busier servers shed work to quieter ones.`;
        } else {
            verdict = 'Load balancing needs attention';
            detail = `Significant imbalance: ${esc(hi.host)} (${hi.mean.toFixed(1)}% avg) carries far more CPU load than ${esc(lo.host)} (${lo.mean.toFixed(1)}% avg) — a ${spreadPct.toFixed(0)}% spread. Rebalance logon groups / batch server groups so capacity on the quieter servers is actually used.`;
        }

        return `
            <div class="chart-stats">
                <div class="chart-stats-title">Load-balancing analysis</div>
                <div style="font-size:13px;font-weight:600;color:var(--brand);margin-bottom:4px">${esc(verdict)}</div>
                <div style="font-size:12px;color:var(--text-secondary);line-height:1.5">${detail}</div>
            </div>`;
    },

    async _renderCpuCharts(app) {
        /* Chart teardown is handled by _activateSection before this runs; here we
           only (re)initialise the registries and clear this section's bodies so a
           re-render never leaves a stale/empty duplicate card behind. */
        this._charts = this._charts || {};
        this._surfaces = this._surfaces || {};

        const grpOverview = document.getElementById('twa-cpugrp-overview-body');
        const grpPerHost  = document.getElementById('twa-cpugrp-perhost-body');
        const grp3d       = document.getElementById('twa-cpugrp-3d-body');
        const grpSaps     = document.getElementById('twa-cpugrp-saps-body');
        const grpStackedEl = document.getElementById('twa-cpugrp-stacked-body');
        [grpOverview, grpPerHost, grp3d, grpSaps, grpStackedEl].forEach(el => { if (el) el.innerHTML = ''; });

        const data = this._appTierSeries(app);
        const { labels, cpu, hosts } = data;

        const noDataMsg = el => {
            if (el) el.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">No application-tier servers found in the data.</div></div>`;
        };
        if (!hosts.length) {
            [grpOverview, grpPerHost, grp3d, grpSaps].forEach(noDataMsg);
            return;
        }

        const idOf = s => String(s).replace(/[^A-Za-z0-9]/g, '').toLowerCase();
        const setCount = (id, n, label) => {
            const el = document.getElementById(id);
            if (el) el.textContent = `${n} ${label}`;
        };
        let seq = 1;

        /* ── Group 1: Average CPU for all AS (one multi-server overview chart) ── */
        if (grpOverview) {
            this._mountLineChart(grpOverview, {
                id: 'cpu-avg-all-as',
                fileSection: 'cpu',
                fileSeq: seq++,
                title: 'Average CPU usage for all the AS',
                labels,
                series: cpu,
                yLabel: 'CPU utilization (%)',
                yMin: 0, yMax: 100,
                xLabel: 'Date / Time',
                lineWidth: 2.6,
                tall: true,
            });

            /* Monthly hourly-average CPU per server: X = Hour (0-23), Y = mean
               CPU% across all days for that hour. Reveals the load-balancing
               profile — well-balanced systems show the servers' curves tracking
               each other closely. */
            const hourly = this._hourlyAvgCpuByHour(app, hosts);
            if (hourly.hosts.length) {
                this._mountLineChart(grpOverview, {
                    id: 'cpu-monthly-hourly',
                    fileSection: 'cpu',
                    fileSeq: seq++,
                    title: 'Peak CPU consumption for all AS — Monthly hourly averages',
                    labels: hourly.labels,            /* "00".."23" */
                    series: hourly.series,            /* host → [24] */
                    yLabel: 'CPU Utilization (%)',
                    yMin: 0, yMax: this._niceCeil(Math.max(10, hourly.peak * 1.2)),
                    xLabel: 'Hour of day',
                    lineWidth: 2.4,
                    tall: true,
                    analysisHtml: this._loadBalancingAnalysis(hourly),
                });
            }
        }
        setCount('twa-cpugrp-overview-count', grpOverview ? (document.getElementById('twa-cpugrp-overview-body')?.querySelectorAll('.chart-card-wrap').length || 1) : 1, 'charts');

        /* ── Group 2: Per-host CPU Utilization (%) — Date+Time X-axis ──
           Built in a yielding loop so the floating progress bar keeps advancing
           and the UI stays responsive even with many servers. */
        if (grpPerHost) {
            for (let hi = 0; hi < hosts.length; hi++) {
                const h = hosts[hi];
                this._mountLineChart(grpPerHost, {
                    id: `cpu-${idOf(h)}`,
                    fileSection: 'cpu',
                    fileSeq: seq++,
                    title: `CPU Utilization (%) — ${h}`,
                    labels,
                    series: { [h]: cpu[h] },
                    yLabel: 'CPU utilization (%)',
                    yMin: 0, yMax: 100,
                    xLabel: 'Date / Time',
                });
                if (hi % 2 === 1) {
                    this._chartProgress(`Building CPU charts… (${hi + 1}/${hosts.length} servers)`, 55 + Math.round((hi / hosts.length) * 25));
                    await this._nextFrame();
                }
            }
        }
        setCount('twa-cpugrp-perhost-count', hosts.length, hosts.length === 1 ? 'chart' : 'charts');

        /* ── Group 3: 3D CPU surface — one per server (Date × Hour × avg CPU%) ──
           Rendered ON DEMAND: each host shows a lightweight placeholder with a
           "Show 3D surface" button. The Plotly/WebGL surface is built only when
           clicked, so we never create N WebGL contexts up front (the main cause
           of the scroll freeze). Switching tabs purges any that were opened. */
        let cnt3d = 0;
        if (grp3d) {
            if (typeof Plotly !== 'undefined') {
                const esc3d = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
                this._pendingSurfaces = this._pendingSurfaces || {};
                hosts.forEach(h => {
                    const grid = this._hourlyCpuGrid(app, h);
                    if (!grid.dates.length) return;
                    const sid = `cpu3d-${idOf(h)}`;
                    const cfg = { id: sid, fileSection: 'cpu', fileSeq: seq++, title: `Peak CPU usage for ${h}`, hostName: h, grid };
                    this._pendingSurfaces[sid] = cfg;

                    const ph = document.createElement('div');
                    ph.className = 'card chart-card-wrap chart-3d-placeholder';
                    ph.innerHTML = `
                        <div style="padding:24px 16px;text-align:center">
                            <div style="font-weight:600;font-size:13px;margin-bottom:4px">Peak CPU usage for ${esc3d(h)}</div>
                            <div style="font-size:12px;color:var(--text-muted);margin-bottom:14px">3D surface (Date × Hour × avg CPU%). Rendered on demand to keep the page responsive.</div>
                            <button class="btn btn-blue btn-show-3d" data-sid="${sid}">Show 3D surface</button>
                        </div>`;
                    grp3d.appendChild(ph);
                    ph.querySelector('.btn-show-3d').addEventListener('click', () => {
                        const pcfg = this._pendingSurfaces[sid];
                        if (!pcfg) return;
                        const holder = document.createElement('div');
                        ph.replaceWith(holder);
                        this._mountSurfaceChart(holder, pcfg);
                    });
                    cnt3d++;
                });
                if (!cnt3d) {
                    grp3d.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">No hourly data available for 3D charts.</div></div>`;
                }
            } else {
                grp3d.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">3D surface charts require the Plotly library, which failed to load (check your internet connection and reload).</div></div>`;
            }
        }
        setCount('twa-cpugrp-3d-count', cnt3d, cnt3d === 1 ? 'chart' : 'charts');

        /* ── Group 4: SAPS utilization hourly average — one per server ──
           Derived from the same hourly CPU grid as the 3D surface so the
           computation is consistent: SAPSUtil_h = (avgCPU_h / 100) × SAPSDeployed. */
        let cntSaps = 0;
        if (grpSaps) {
            for (let hi = 0; hi < hosts.length; hi++) {
                const h = hosts[hi];
                const hs = this._hourlySapsFromGrid(app, h);
                if (!hs.labels.length) continue;
                let peak = 0;
                for (let i = 0; i < hs.dep.length; i++)  if (hs.dep[i]  != null && hs.dep[i]  > peak) peak = hs.dep[i];
                for (let i = 0; i < hs.util.length; i++) if (hs.util[i] != null && hs.util[i] > peak) peak = hs.util[i];
                const yMax = peak > 0 ? this._niceCeil(peak * 1.1) : 100;
                this._mountLineChart(grpSaps, {
                    id: `saps-${idOf(h)}`,
                    fileSection: 'cpu',
                    fileSeq: seq++,
                    title: `SAPS Utilization — Hourly Average for ${h}`,
                    labels: hs.labels,
                    namedSeries: [
                        { label: `SAPS Utilization (hourly avg) — ${h}`, data: hs.util, color: '#0070F2' },
                        { label: `SAPS Deployed — ${h}`,                 data: hs.dep,  color: '#EF4444' },
                    ],
                    yLabel: 'SAPS',
                    yMin: 0, yMax,
                    xLabel: 'Date / Hour',
                });
                cntSaps++;
                if (hi % 2 === 1) {
                    this._chartProgress(`Building SAPS charts… (${hi + 1}/${hosts.length} servers)`, 80 + Math.round((hi / hosts.length) * 8));
                    await this._nextFrame();
                }
            }
            if (!cntSaps) {
                grpSaps.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">SAPS charts require landscape data (Cores and SAPS/Core) from Module 01.</div></div>`;
            }
        }
        setCount('twa-cpugrp-saps-count', cntSaps, cntSaps === 1 ? 'chart' : 'charts');

        /* ── Group 5: Stacked SAPS — all AS combined vs total deployed capacity ──
           One stacked bar chart: each bar segment = one server's hourly-AVERAGE
           SAPS utilization, computed by averaging the actual "SAPS Utilization
           (SAPS)" column per (Date, Hour) bucket — matching the user's manual
           chart (pivot = AVERAGE of SAPS Utilization by Date + Hour). A flat red
           line shows the total deployed SAPS capacity for the application layer. */
        const grpStacked = document.getElementById('twa-cpugrp-stacked-body');
        if (grpStacked) {
            const sd = this._hourlyStackedSaps(app, hosts);
            if (sd.labels.length && sd.hostSeries.length) {
                this._mountStackedBarChart(grpStacked, {
                    id: 'saps-stacked-all',
                    fileSection: 'cpu',
                    fileSeq: seq++,
                    title: 'Peak CPU consumption for all AS',
                    labels: sd.labels,
                    hostSeries: sd.hostSeries,   /* [{ label, data }] — hourly-average SAPS Utilization */
                    totalDeployed: sd.totalDeployed,
                    yLabel: 'SAPS Utilization (SAPS)',
                    xLabel: 'Date / Hour',
                    statsMetric: 'SAPS Utilization',
                    hint: 'Stacked bars show hourly-average SAPS utilization per server · red line = total deployed SAPS capacity on the application layer',
                });
            } else {
                grpStacked.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">Stacked SAPS chart requires landscape data (Cores and SAPS/Core) from Module 01.</div></div>`;
            }
        }

        /* ── Group 6: Top 20 Relevant CPU Peaks — stacked SAPS, date-ascending ──
           Select the top-20 representative moments by CPU% (same deduplication
           logic as before), then display them as a stacked-column chart showing
           SAPS Utilization per server.  X-axis is sorted chronologically (oldest
           left, newest right) so the chart tells a timeline story. */
        const grpRelPeaks = document.getElementById('twa-cpugrp-relpeaks-body');
        if (grpRelPeaks) {
            grpRelPeaks.innerHTML = '';
            const ci = this._colIndex || {};
            const peaks = (ci['CPU Utilization (%)'] != null)
                ? this._relevantPeaks(app, { valueOf: (r, c) => { const v = parseFloat(r[c['CPU Utilization (%)']]); return isNaN(v) ? null : v; } }, 20)
                : [];
            if (peaks.length && ci['SAPS Utilization (SAPS)'] != null) {
                const sd = this._peaksToStackedSeries(peaks, 'SAPS Utilization (SAPS)', hosts);
                if (sd.labels.length && sd.hostSeries.length) {
                    this._mountStackedBarChart(grpRelPeaks, {
                        id: 'cpu-rel-peaks', fileSection: 'cpu', fileSeq: seq++,
                        title: 'Top 20 Relevant CPU Peaks',
                        labels: sd.labels,
                        hostSeries: sd.hostSeries,
                        yLabel: 'SAPS',
                        xLabel: 'Date / Time',
                        statsMetric: 'SAPS Utilization',
                        hint: 'Top 20 representative CPU peak moments · stacked bars show SAPS Utilization per server · sorted by Date ascending (oldest left, newest right)',
                    });
                } else {
                    this._mountRelevantPeaks(grpRelPeaks, {
                        id: 'cpu-rel-peaks', fileSection: 'cpu', fileSeq: seq++,
                        title: 'Top 20 Relevant CPU Peaks',
                        peaks, valueLabel: 'CPU Utilization (%)', color: '#FFC000',
                    });
                }
            } else {
                this._mountRelevantPeaks(grpRelPeaks, {
                    id: 'cpu-rel-peaks', fileSection: 'cpu', fileSeq: seq++,
                    title: 'Top 20 Relevant CPU Peaks',
                    peaks, valueLabel: 'CPU Utilization (%)', color: '#FFC000',
                });
            }
            setCount('twa-cpugrp-relpeaks-count', peaks.length, peaks.length === 1 ? 'peak' : 'peaks');
        }

        /* Wire sub-group collapse/expand toggles. */
        document.querySelectorAll('#twa-sec-cpu .chart-group-head').forEach(head => {
            if (head._wiredGroup) return;
            head._wiredGroup = true;
            head.addEventListener('click', () => head.closest('.chart-group').classList.toggle('collapsed'));
        });
    },

    /* Build the written commentary shown below the page-out chart. Page-out is an
       undesirable performance behaviour, so we rank servers by their peak page-out
       and report the worst offenders with their peak + how many non-zero readings
       they had. `series` is a host→data map; `unit` labels the values. */
    _pageOutCommentary(series, unit) {
        const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
        const fmt = v => unit && unit.indexOf('%') >= 0
            ? `${Number(v).toLocaleString(undefined, { maximumFractionDigits: 2 })}%`
            : `${Math.round(v).toLocaleString()} ${esc(unit || '')}`.trim();

        const perHost = Object.keys(series).map(h => {
            const data = series[h] || [];
            let peak = 0, hits = 0;
            data.forEach(v => { if (v != null && v > 0) { hits++; if (v > peak) peak = v; } });
            return { host: h, peak, hits };
        }).filter(o => o.peak > 0);

        if (!perHost.length) {
            return `<div class="chart-note">No page-out activity was observed on any application
                server during the measured period — healthy from a memory-pressure standpoint.</div>`;
        }

        perHost.sort((a, b) => b.peak - a.peak);
        const top = perHost.slice(0, Math.min(5, perHost.length));
        const items = top.map(o =>
            `<li><strong>${esc(o.host)}</strong> — peak ${fmt(o.peak)}, ${o.hits.toLocaleString()} reading${o.hits === 1 ? '' : 's'} with page-out activity</li>`
        ).join('');
        const lead = perHost.length === 1
            ? `One application server reported page-out activity.`
            : `${perHost.length} application servers reported page-out activity.`;
        return `<div class="chart-danger">
            <strong>Page-out activity detected.</strong> Paging out is an undesirable behaviour
            that degrades performance — it indicates memory pressure forcing the OS to swap.
            ${lead} Highest offenders by peak:
            <ul style="margin:6px 0 0 18px;padding:0">${items}</ul></div>`;
    },

    /* ══════════════════════════════════════════════════════
       MEMORY UTILIZATION CHARTS
       1. Per-host line chart: Physical Memory / EM Configured /
          EM Allocated / Free Memory(+FS) over Date+Time.
       2. Stacked bar: EM Allocated (GB) across all AS.
       3. Stacked bar: Heap Memory (GB) across all AS.
       A small map table (Physical Memory + EM Configured per host)
       is shown at the top of the section.
       ══════════════════════════════════════════════════════ */
    async _renderMemCharts(app) {
        const grpPerHost = document.getElementById('twa-memgrp-perhost-body');
        const grpEm      = document.getElementById('twa-memgrp-em-body');
        const grpHeap    = document.getElementById('twa-memgrp-heap-body');
        const mapWrap    = document.getElementById('twa-memmap');
        [grpPerHost, grpEm, grpHeap].forEach(el => { if (el) el.innerHTML = ''; });

        const mem = this._appTierMemSeries(app);
        const { labels, hosts, physMem, emCfg, emAlloc, freeMem, freeLabel } = mem;

        const noData = el => { if (el) el.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">No application-tier servers found in the data.</div></div>`; };
        if (!hosts.length) {
            if (mapWrap) mapWrap.innerHTML = '';
            [grpPerHost, grpEm, grpHeap].forEach(noData);
            return;
        }

        const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
        const idOf = s => String(s).replace(/[^A-Za-z0-9]/g, '').toLowerCase();
        const setCount = (id, n, label) => { const el = document.getElementById(id); if (el) el.textContent = `${n} ${label ? label : (n === 1 ? 'chart' : 'charts')}`; };
        const fmtGb = v => (v == null || v === '' || isNaN(v)) ? '—' : `${Number(v).toLocaleString(undefined, { maximumFractionDigits: 2 })} GB`;
        let seq = 1;

        /* ── Map table: Physical Memory (GB) + EM Configured (GB) per server ── */
        if (mapWrap) {
            mapWrap.innerHTML = `
                <div class="card" style="padding:0;margin-bottom:4px">
                    <div class="card-header" style="padding:14px 18px 12px;border-radius:14px 14px 0 0">
                        <div>
                            <div class="card-title">Application Server Memory Map</div>
                            <div class="card-sub">Physical memory (from the landscape) and SAP Extended Memory configured per server</div>
                        </div>
                    </div>
                    <div class="table-wrap" style="border:none;border-radius:0 0 14px 14px;overflow-x:auto">
                        <table class="data-table">
                            <thead>
                                <tr>
                                    <th>Application Server</th>
                                    <th style="width:200px">Physical Memory (GB)</th>
                                    <th style="width:200px">EM Configured (GB)</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${hosts.map(h => `<tr>
                                    <td style="font-weight:600;white-space:nowrap">${esc(h)}</td>
                                    <td>${fmtGb(physMem[h])}</td>
                                    <td>${fmtGb(emCfg[h])}</td>
                                </tr>`).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>`;
        }

        /* ── Group 1: per-host memory usage line charts (yielding loop) ── */
        if (grpPerHost) {
            for (let hi = 0; hi < hosts.length; hi++) {
                const h = hosts[hi];
                const physArr = new Array(labels.length).fill(physMem[h] != null ? physMem[h] : null);
                const emCfgArr = new Array(labels.length).fill(emCfg[h] != null ? emCfg[h] : null);
                const namedSeries = [
                    { label: 'Physical Memory (GB)', data: physArr,     color: '#5B8FF9' },
                    { label: 'EM Configured (GB)',   data: emCfgArr,    color: '#FF0000' },
                    { label: 'EM Allocated (GB)',    data: emAlloc[h],  color: '#F2A815' },
                    { label: freeLabel,              data: freeMem[h],  color: '#30BF78' },
                ];
                /* Y ceiling = max across all four series for this host. */
                let peak = 0;
                namedSeries.forEach(s => s.data.forEach(v => { if (v != null && v > peak) peak = v; }));
                const yMax = peak > 0 ? this._niceCeil(peak * 1.1) : 100;

                /* SAP Extended Memory threshold check: the recommended maximum EM
                   usage is 80% of the configured EM. If this host's EM Allocated
                   peak crosses that, flag it with a red legend below its chart. */
                let emWarnHtml = '';
                const cfgGb = emCfg[h];
                if (cfgGb != null && cfgGb !== '' && !isNaN(cfgGb) && Number(cfgGb) > 0) {
                    const recEm = Number(cfgGb) * 0.8;    /* Maximum recommended EM (GB) */
                    let emPeak = 0;
                    (emAlloc[h] || []).forEach(v => { if (v != null && v > emPeak) emPeak = v; });
                    if (emPeak > recEm) {
                        const g = v => Number(v).toLocaleString(undefined, { maximumFractionDigits: 2 });
                        emWarnHtml = `<div class="chart-danger">
                            <strong>⚠ SAP Extended Memory above the recommended threshold.</strong>
                            EM Allocated peaked at <strong>${g(emPeak)} GB</strong>, above the
                            Maximum recommended EM of <strong>${g(recEm)} GB</strong>
                            (80% of ${g(cfgGb)} GB configured). This server is above our SAP
                            recommended threshold.</div>`;
                    }
                }

                this._mountLineChart(grpPerHost, {
                    id: `mem-${idOf(h)}`,
                    fileSection: 'mem',
                    fileSeq: seq++,
                    title: `Memory usage for ${h}`,
                    labels,
                    namedSeries,
                    yLabel: 'Memory (GB)',
                    yMin: 0, yMax,
                    xLabel: 'Date / Time',
                    analysisHtml: emWarnHtml,
                });
                if (hi % 2 === 1) {
                    this._chartProgress(`Building memory charts… (${hi + 1}/${hosts.length} servers)`, 55 + Math.round((hi / hosts.length) * 25));
                    await this._nextFrame();
                }
            }
        }
        setCount('twa-memgrp-perhost-count', hosts.length);

        /* ── Group 2: stacked EM Allocated (GB) across all AS ── */
        if (grpEm) {
            const hostSeries = hosts
                .map(h => ({ label: h, data: emAlloc[h] }))
                .filter(s => s.data && s.data.some(v => v != null));
            if (hostSeries.length) {
                this._mountStackedBarChart(grpEm, {
                    id: 'mem-em-stacked',
                    fileSection: 'mem',
                    fileSeq: seq++,
                    title: 'EM Utilization for all AS',
                    labels,
                    hostSeries,
                    unit: 'GB',
                    yLabel: 'Memory (GB)',
                    xLabel: 'Date / Time',
                    statsMetric: 'EM Allocated (GB)',
                    hint: 'Stacked bars show EM Allocated (GB) per application server',
                });
                setCount('twa-memgrp-em-count', 1);
            } else {
                grpEm.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">No EM Allocated (GB) data available.</div></div>`;
                setCount('twa-memgrp-em-count', 0);
            }

            /* ── Group 2b: Top 20 Relevant EM Peaks — stacked EM GB, date-ascending ──
               Select the top-20 representative moments by EM% (same deduplication
               logic), then display them as stacked columns showing EM Allocated (GB)
               per server.  X-axis sorted chronologically (oldest left, newest right). */
            const grpEmPeaks = document.getElementById('twa-memgrp-empeaks-body');
            if (grpEmPeaks) {
                grpEmPeaks.innerHTML = '';
                const ci = this._colIndex || {};
                const iAlloc = ci['EM Allocated (GB)'], iCfg = ci['EM Configured (GB)'];
                const peaks = (iAlloc != null && iCfg != null)
                    ? this._relevantPeaks(app, {
                          /* Sort/dedup by EM utilisation ratio (GB / configured GB) so
                             temporal distribution favours the busiest moments. */
                          valueOf: (r) => {
                              const a = parseFloat(r[iAlloc]), c = parseFloat(r[iCfg]);
                              if (isNaN(a) || isNaN(c) || c <= 0) return null;
                              return a / c * 100;
                          },
                          /* rawValueOf: absolute EM GB — used for the "guarantee top 3"
                             seed so the 3 largest actual GB values always appear. */
                          rawValueOf: (r) => {
                              const a = parseFloat(r[iAlloc]);
                              return isNaN(a) ? null : a;
                          },
                          /* Always include the 3 absolute highest EM GB peaks. */
                          guaranteeTop: 3,
                      }, 20)
                    : [];
                if (peaks.length && iAlloc != null) {
                    const sd = this._peaksToStackedSeries(peaks, 'EM Allocated (GB)', hosts);
                    /* Average of the 50 highest EM Allocated (GB) values in the period. */
                    const emTop = this._avgTopRawByCol(app, 'EM Allocated (GB)', 50);
                    if (sd.labels.length && sd.hostSeries.length) {
                        this._mountStackedBarChart(grpEmPeaks, {
                            id: 'mem-em-rel-peaks', fileSection: 'mem', fileSeq: seq++,
                            title: 'Top 20 Relevant EM Peaks',
                            labels: sd.labels,
                            hostSeries: sd.hostSeries,
                            unit: 'GB',
                            yLabel: 'Memory (GB)',
                            xLabel: 'Date / Time',
                            statsMetric: 'EM Allocated (GB)',
                            topNPeaks: 50,
                            avgTopPeaksOverride: emTop ? emTop.avg : null,
                            avgTopPeaksCount: emTop ? emTop.count : 50,
                            hint: 'Top 20 representative EM peak moments · stacked bars show EM Allocated (GB) per server · sorted by Date ascending (oldest left, newest right)',
                        });
                    } else {
                        this._mountRelevantPeaks(grpEmPeaks, {
                            id: 'mem-em-rel-peaks', fileSection: 'mem', fileSeq: seq++,
                            title: 'Top 20 Relevant EM Peaks',
                            peaks, valueLabel: 'EM Utilization (%)', color: '#F2A815',
                        });
                    }
                } else {
                    this._mountRelevantPeaks(grpEmPeaks, {
                        id: 'mem-em-rel-peaks', fileSection: 'mem', fileSeq: seq++,
                        title: 'Top 20 Relevant EM Peaks',
                        peaks, valueLabel: 'EM Utilization (%)', color: '#F2A815',
                    });
                }
                setCount('twa-memgrp-empeaks-count', peaks.length, peaks.length === 1 ? 'peak' : 'peaks');
            }
        }

        /* ── Group 3: stacked Heap Memory (GB) across all AS ── */
        if (grpHeap) {
            const heap = this._stackedMemByHost(app, 'Heap Memory (GB)', hosts, labels);
            const hostSeries = heap.hostSeries.filter(s => s.data && s.data.some(v => v != null));
            if (hostSeries.length) {
                this._mountStackedBarChart(grpHeap, {
                    id: 'mem-heap-stacked',
                    fileSection: 'mem',
                    fileSeq: seq++,
                    title: 'Stacked Heap utilization for all AS',
                    labels,
                    hostSeries,
                    unit: 'GB',
                    yLabel: 'Memory (GB)',
                    xLabel: 'Date / Time',
                    statsMetric: 'Heap Memory (GB)',
                    hint: 'Stacked bars show Heap Memory (GB) per application server',
                });
                setCount('twa-memgrp-heap-count', 1);
            } else {
                grpHeap.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">No Heap Memory (GB) data available.</div></div>`;
                setCount('twa-memgrp-heap-count', 0);
            }

            /* ── Group 3b: Top 20 Relevant Heap Memory Peaks — stacked GB, date-ascending ──
               Select the top-20 representative moments by heap ratio (heap GB / physical GB),
               then display them as stacked columns showing actual Heap Memory (GB) per server.
               X-axis sorted chronologically (oldest left, newest right).
               The 3 absolute largest heap GB values are always guaranteed to appear. */
            const grpHeapPeaks = document.getElementById('twa-memgrp-heappeaks-body');
            if (grpHeapPeaks) {
                grpHeapPeaks.innerHTML = '';
                const ci = this._colIndex || {};
                const iHeap = ci['Heap Memory (GB)'], iPhys = ci['Physical Memory (GB)'];
                const peaks = (iHeap != null)
                    ? this._relevantPeaks(app, {
                          valueOf: (r) => {
                              const h = parseFloat(r[iHeap]);
                              if (isNaN(h) || h <= 0) return null;
                              if (iPhys != null) {
                                  const p = parseFloat(r[iPhys]);
                                  if (!isNaN(p) && p > 0) return h / p * 100;
                              }
                              return h;   /* fallback: use raw GB when no physical column */
                          },
                          rawValueOf: (r) => {
                              const h = parseFloat(r[iHeap]);
                              return isNaN(h) ? null : h;
                          },
                          guaranteeTop: 3,
                      }, 20)
                    : [];
                if (peaks.length && iHeap != null) {
                    const sd = this._peaksToStackedSeries(peaks, 'Heap Memory (GB)', hosts);
                    /* Average of the 30 highest Heap Memory (GB) values in the period. */
                    const heapTop = this._avgTopRawByCol(app, 'Heap Memory (GB)', 30);
                    if (sd.labels.length && sd.hostSeries.length) {
                        this._mountStackedBarChart(grpHeapPeaks, {
                            id: 'mem-heap-rel-peaks', fileSection: 'mem', fileSeq: seq++,
                            title: 'Top 20 Relevant Heap Memory Peaks',
                            labels: sd.labels,
                            hostSeries: sd.hostSeries,
                            unit: 'GB',
                            yLabel: 'Memory (GB)',
                            xLabel: 'Date / Time',
                            statsMetric: 'Heap Memory (GB)',
                            topNPeaks: 30,
                            avgTopPeaksOverride: heapTop ? heapTop.avg : null,
                            avgTopPeaksCount: heapTop ? heapTop.count : 30,
                            hint: 'Top 20 representative Heap Memory peak moments · stacked bars show Heap Memory (GB) per server · sorted by Date ascending (oldest left, newest right)',
                        });
                    } else {
                        this._mountRelevantPeaks(grpHeapPeaks, {
                            id: 'mem-heap-rel-peaks', fileSection: 'mem', fileSeq: seq++,
                            title: 'Top 20 Relevant Heap Memory Peaks',
                            peaks, valueLabel: 'Heap Memory (GB)', color: '#70AD47',
                        });
                    }
                } else {
                    grpHeapPeaks.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">No Heap Memory (GB) data available.</div></div>`;
                }
                setCount('twa-memgrp-heappeaks-count', peaks.length, peaks.length === 1 ? 'peak' : 'peaks');
            }
        }

        /* ── Group 4: Page Out activity — a SINGLE line chart with one line per
           server (different colours), titled "Page out activity for all AS".
           Column + Y-axis label depend on the source table:
             SMON → PAGE_OUT_PERC  → "Page Out (% of RAM)"
             MON  → OUT_KB_SEC     → "Page Out (KB/sec)"
           Zeros are plotted (a flat line at 0), so servers that never page out
           still appear — we only skip a server with NO data points at all. */
        const grpPageout = document.getElementById('twa-memgrp-pageout-body');
        if (grpPageout) {
            const ci = this._colIndex || {};
            const findCol = want => {
                const w = want.toUpperCase().replace(/\s+/g, '');
                for (const k in ci) if (k.toUpperCase().replace(/\s+/g, '') === w) return k;
                return null;
            };
            const colPagePerc = findCol('PAGE_OUT_PERC');
            const colOutKb    = findCol('OUT_KB_SEC');
            let poCol = null, poYLabel = '';
            if (this.sourceTable === 'MON' && colOutKb) {
                poCol = colOutKb; poYLabel = 'Page Out (KB/sec)';
            } else if (colPagePerc) {
                poCol = colPagePerc; poYLabel = 'Page Out (% of RAM)';
            } else if (colOutKb) {
                poCol = colOutKb; poYLabel = 'Page Out (KB/sec)';
            }

            let mounted = false;
            if (poCol) {
                const po = this._stackedMemByHost(app, poCol, hosts, labels);
                /* Keep a server only if it has at least one real data point (incl. 0). */
                const series = {};
                let peak = 0;
                hosts.forEach(h => {
                    const data = (po.hostSeries.find(s => s.label === h) || {}).data;
                    if (data && data.some(v => v != null)) {
                        series[h] = data;
                        data.forEach(v => { if (v != null && v > peak) peak = v; });
                    }
                });
                if (Object.keys(series).length) {
                    const yMax = peak > 0 ? this._niceCeil(peak * 1.1) : (poYLabel.indexOf('%') >= 0 ? 100 : 10);
                    const poUnit = poYLabel.indexOf('%') >= 0 ? '%' : 'KB/sec';
                    this._mountLineChart(grpPageout, {
                        id: 'pageout-all',
                        fileSection: 'mem',
                        fileSeq: seq++,
                        title: 'Page out activity for all AS',
                        labels,
                        series,                 /* host → data, auto-coloured one line per server */
                        yLabel: poYLabel,
                        yMin: 0, yMax,
                        xLabel: 'Date / Time',
                        analysisHtml: this._pageOutCommentary(series, poUnit),
                    });
                    mounted = true;
                }
            }
            if (!mounted) {
                grpPageout.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">No page-out data available (PAGE_OUT_PERC / OUT_KB_SEC not present in this source).</div></div>`;
            }
            setCount('twa-memgrp-pageout-count', mounted ? 1 : 0);
        }

        /* ── Extended Memory Usage Summary ──
           Lists ONLY the servers that exceeded the SAP recommended EM threshold
           (recommended max = 80% of EM Configured). Reports, per offending host:
           Physical Memory, EM Configured, number of EM Allocated readings above
           the recommended threshold, and the highest EM Allocated peak. */
        try {
            const secMem = document.getElementById('twa-sec-mem');
            if (secMem) {
                const offenders = [];
                hosts.forEach(h => {
                    const cfgGb = emCfg[h];
                    if (cfgGb == null || cfgGb === '' || isNaN(cfgGb) || Number(cfgGb) <= 0) return;
                    const recEm = Number(cfgGb) * 0.8;
                    const series = emAlloc[h] || [];
                    let emPeak = 0, exceedCount = 0;
                    series.forEach(v => {
                        if (v == null || isNaN(v)) return;
                        if (v > emPeak) emPeak = v;
                        if (v > recEm) exceedCount++;
                    });
                    if (emPeak > recEm) {
                        offenders.push({ host: h, physMem: physMem[h], emCfg: Number(cfgGb), recEm, exceedCount, emPeak });
                    }
                });
                offenders.sort((a, b) => b.emPeak - a.emPeak);

                if (offenders.length) {
                    const g = v => (v == null || v === '' || isNaN(v)) ? '—'
                        : Number(v).toLocaleString(undefined, { maximumFractionDigits: 2 });
                    const summary = document.createElement('div');
                    summary.className = 'chart-group';
                    summary.innerHTML = `
                        <div class="chart-group-head">
                            <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                            <span>Extended Memory Usage Summary</span>
                        </div>
                        <div class="chart-group-body">
                            <div class="card" style="padding:0">
                                <div class="card-header" style="padding:14px 18px 12px;border-radius:14px 14px 0 0">
                                    <div>
                                        <div class="card-title">Servers above the SAP recommended EM threshold</div>
                                        <div class="card-sub">Recommended maximum EM Allocated = 80% of EM Configured · ${offenders.length} server(s) exceeded it</div>
                                    </div>
                                </div>
                                <div class="table-wrap" style="border:none;border-radius:0 0 14px 14px;overflow-x:auto">
                                    <table class="data-table">
                                        <thead>
                                            <tr>
                                                <th>Hostname</th>
                                                <th style="text-align:right">Physical Memory (GB)</th>
                                                <th style="text-align:right">EM Configured (GB)</th>
                                                <th style="text-align:right">Peaks above recommended threshold</th>
                                                <th style="text-align:right">Highest EM Allocated peak (GB)</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            ${offenders.map(o => `
                                                <tr>
                                                    <td style="font-weight:600">${esc(o.host)}</td>
                                                    <td style="text-align:right">${g(o.physMem)}</td>
                                                    <td style="text-align:right">${g(o.emCfg)}</td>
                                                    <td style="text-align:right">${o.exceedCount.toLocaleString()}</td>
                                                    <td style="text-align:right;color:#C0392B;font-weight:700">${g(o.emPeak)}</td>
                                                </tr>`).join('')}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>`;
                    secMem.appendChild(summary);
                }
            }
        } catch (err) {
            console.warn('[TechWorkloadApp] Extended Memory Usage Summary failed:', err);
        }

        /* Wire sub-group collapse/expand toggles. */
        document.querySelectorAll('#twa-sec-mem .chart-group-head').forEach(head => {
            if (head._wiredGroup) return;
            head._wiredGroup = true;
            head.addEventListener('click', () => head.closest('.chart-group').classList.toggle('collapsed'));
        });
    },

    /* ══════════════════════════════════════════════════════
       WORK PROCESS INVENTORY  (from uploaded SAP instance profiles)
       ──────────────────────────────────────────────────────
       Consolidates the configured work-process parameters across every
       uploaded application-server profile into one row per host, plus
       landscape totals. Built BEFORE any WP charts and used as their source.

       Parameters (missing → 0, logged for troubleshooting):
         rdisp/wp_no_dia  rdisp/wp_no_btc  rdisp/wp_no_vb  rdisp/wp_no_vb2
         rdisp/wp_no_spo  rdisp/wp_no_enq               → Configured_WPs (sum)
         rdisp/wp_no_restricted                          → Dynamic/Restricted (separate)
         rdisp/wp_max_no                                 → Maximum (separate)
    ══════════════════════════════════════════════════════ */
    _WP_PARAMS: [
        { key: 'rdisp/wp_no_dia',  field: 'dia',  label: 'Dialog WPs' },
        { key: 'rdisp/wp_no_btc',  field: 'btc',  label: 'Background WPs' },
        { key: 'rdisp/wp_no_vb',   field: 'vb',   label: 'Update WPs' },
        { key: 'rdisp/wp_no_vb2',  field: 'vb2',  label: 'Update2 WPs' },
        { key: 'rdisp/wp_no_spo',  field: 'spo',  label: 'Spool WPs' },
        { key: 'rdisp/wp_no_enq',  field: 'enq',  label: 'Enqueue WPs' },
    ],

    /* Parse a profile parameter value to an integer. SAP profiles may use
       expressions like "$(rdisp/wp_no_dia)" or embedded arithmetic; we extract
       the leading integer and return null when nothing numeric is present
       (so the caller can treat it as "missing"). */
    _wpNum(raw) {
        if (raw == null) return null;
        const m = String(raw).match(/-?\d+/);
        return m ? parseInt(m[0], 10) : null;
    },

    _buildWpInventory() {
        const profiles = this.profiles || [];
        if (!profiles.length) return { rows: [], totals: null, missingLog: [] };

        /* DEFAULT.PFL values act as a fallback for hosts whose instance profile
           omits a parameter (common in real landscapes). */
        const def = profiles.find(p => p.kind === 'DEFAULT');
        const defParams = (def && def.params) || {};

        /* Instance profiles define a host; group by host (avoid duplicates).
           ASCS/other profiles without WP params are ignored for the inventory
           unless they are the only source for a host. */
        const byHost = new Map();   /* host → merged params object (instance wins over default) */
        const missingLog = [];

        const instanceProfiles = profiles.filter(p => p.kind === 'INSTANCE' && p.host);
        /* Fallback: if no INSTANCE profiles were classified, use any profile that
           carries at least one WP parameter and has a host. */
        const sourceProfiles = instanceProfiles.length ? instanceProfiles
            : profiles.filter(p => p.host && this._WP_PARAMS.some(pp => p.params && p.params[pp.key] != null));

        for (const prof of sourceProfiles) {
            const host = prof.host || (prof.params && (prof.params['SAPLOCALHOST'] || prof.params['SAPLOCALHOSTFULL'])) || '';
            if (!host) continue;
            if (byHost.has(host)) continue;   /* first (instance) profile for a host wins — avoid duplicate rows */
            byHost.set(host, prof);
        }

        const rows = [];
        for (const [host, prof] of byHost.entries()) {
            const p = prof.params || {};
            const read = key => {
                let v = this._wpNum(p[key]);
                if (v == null) v = this._wpNum(defParams[key]);   /* DEFAULT.PFL fallback */
                if (v == null) { missingLog.push({ host, profile: prof.name, param: key }); return 0; }
                return v;
            };
            const vals = {};
            this._WP_PARAMS.forEach(pp => { vals[pp.field] = read(pp.key); });
            const configured = this._WP_PARAMS.reduce((s, pp) => s + vals[pp.field], 0);
            const restricted = read('rdisp/wp_no_restricted');
            const maxWp      = read('rdisp/wp_max_no');
            rows.push({ host, ...vals, configured, restricted, maxWp });
        }

        rows.sort((a, b) => a.host.localeCompare(b.host));

        /* Landscape totals. */
        const sum = f => rows.reduce((s, r) => s + (r[f] || 0), 0);
        const totals = {
            dia: sum('dia'), btc: sum('btc'), vb: sum('vb'), vb2: sum('vb2'),
            spo: sum('spo'), enq: sum('enq'),
            configured: sum('configured'),
            restricted: sum('restricted'),
            maxWp: sum('maxWp'),
        };

        if (missingLog.length) {
            console.groupCollapsed(`[WP inventory] ${missingLog.length} missing parameter(s) treated as 0`);
            missingLog.forEach(m => console.info(`  ${m.param.padEnd(24)} missing on ${m.host} (${m.profile})`));
            console.groupEnd();
        }

        return { rows, totals, missingLog };
    },

    _renderWpInventory(sec) {
        const inv = this._buildWpInventory();
        /* Cache for any subsequent WP visualisation to reuse as its source. */
        this._wpInventory = inv;

        if (!inv.rows.length) {
            sec.insertAdjacentHTML('beforeend', `
                <div class="card" style="margin-bottom:16px">
                    <div class="card-header"><div><div class="card-title">Work Process Inventory</div>
                    <div class="card-sub">Upload SAP instance profiles (DEFAULT.PFL · instance · START) to generate the landscape work-process inventory.</div></div></div>
                    <div style="padding:4px 16px 16px;color:var(--text-muted);font-size:13px">No SAP application-server profiles have been uploaded yet.</div>
                </div>`);
            return;
        }

        const t = inv.totals;
        const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
        const n = v => Number(v || 0).toLocaleString();

        const rowsHtml = inv.rows.map(r => `
            <tr>
                <td>${esc(r.host)}</td>
                <td style="text-align:right">${n(r.dia)}</td>
                <td style="text-align:right">${n(r.btc)}</td>
                <td style="text-align:right">${n(r.vb)}</td>
                <td style="text-align:right">${n(r.vb2)}</td>
                <td style="text-align:right">${n(r.spo)}</td>
                <td style="text-align:right">${n(r.enq)}</td>
                <td style="text-align:right;font-weight:600">${n(r.configured)}</td>
                <td style="text-align:right">${n(r.restricted)}</td>
            </tr>`).join('');

        const card = document.createElement('div');
        card.className = 'card chart-card-wrap';
        card.style.marginBottom = '16px';
        card.innerHTML = `
            <div class="chart-toolbar">
                <div class="card-title" style="font-size:15px">Work Process Inventory — SAP Landscape</div>
                <div class="chart-controls">
                    <button class="btn btn-sm" id="twa-wpinv-csv" title="Download the inventory as CSV">CSV</button>
                </div>
            </div>
            <div class="table-wrap" style="border-radius:12px;overflow-x:auto">
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>Hostname</th>
                            <th style="text-align:right">Dialog WPs</th>
                            <th style="text-align:right">Background WPs</th>
                            <th style="text-align:right">Update WPs</th>
                            <th style="text-align:right">Update2 WPs</th>
                            <th style="text-align:right">Spool WPs</th>
                            <th style="text-align:right">Enqueue WPs</th>
                            <th style="text-align:right">Configured WPs</th>
                            <th style="text-align:right">Dynamic / Restricted WPs</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${rowsHtml}
                        <tr style="font-weight:600;border-top:2px solid var(--border-strong)">
                            <td>Landscape Total</td>
                            <td style="text-align:right">${n(t.dia)}</td>
                            <td style="text-align:right">${n(t.btc)}</td>
                            <td style="text-align:right">${n(t.vb)}</td>
                            <td style="text-align:right">${n(t.vb2)}</td>
                            <td style="text-align:right">${n(t.spo)}</td>
                            <td style="text-align:right">${n(t.enq)}</td>
                            <td style="text-align:right">${n(t.configured)}</td>
                            <td style="text-align:right">${n(t.restricted)}</td>
                        </tr>
                    </tbody>
                </table>
            </div>

            <div class="stats-grid" style="margin-top:16px">
                <div class="stat-card">
                    <div class="stat-label">Total Configured Work Processes</div>
                    <div class="stat-value">${n(t.configured)}</div>
                    <div class="stat-sub" style="font-size:11px;color:var(--text-muted)">currently available work processes</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Total Dynamic / Restricted Work Processes</div>
                    <div class="stat-value">${n(t.restricted)}</div>
                    <div class="stat-sub" style="font-size:11px;color:var(--text-muted)">available for dynamic expansion</div>
                </div>
            </div>
            ${inv.missingLog.length ? `<div class="chart-hint" style="margin-top:10px">${inv.missingLog.length} parameter value(s) were missing in the profiles and treated as 0 — see the browser console for details.</div>` : ''}`;
        sec.appendChild(card);

        card.querySelector('#twa-wpinv-csv')?.addEventListener('click', () => {
            const header = ['Hostname', 'Dialog WPs', 'Background WPs', 'Update WPs', 'Update2 WPs', 'Spool WPs', 'Enqueue WPs', 'Configured WPs', 'Dynamic / Restricted WPs'];
            const lines = [header.join(',')];
            inv.rows.forEach(r => lines.push([`"${r.host}"`, r.dia, r.btc, r.vb, r.vb2, r.spo, r.enq, r.configured, r.restricted].join(',')));
            lines.push(['"Landscape Total"', t.dia, t.btc, t.vb, t.vb2, t.spo, t.enq, t.configured, t.restricted].join(','));
            const blob = new Blob([lines.join('\r\n')], { type: 'text/csv' });
            const url = URL.createObjectURL(blob);
            const a = Object.assign(document.createElement('a'), { href: url, download: `${this._sidTag()}_work-process-inventory.csv` });
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(url), 1000);
            showToast('CSV downloaded', '', 'success', 2500);
        });
    },

    /* ══════════════════════════════════════════════════════
       WORK PROCESS ANALYSIS CHARTS
       1. Active Work Processes Stacked      (stacked bar, ACT_WPS)
       2. Active Dialog Work Processes Stacked (stacked bar, ACT_DIA)
       3. WP free for RFC — normal priority  (line per server, TRFC_FREE)
          + low priority (SMON only, RFC_LOW)
       ══════════════════════════════════════════════════════ */
    _renderWpCharts(app) {
        const sec = document.getElementById('twa-sec-wp');
        if (!sec) return;
        sec.innerHTML = '';

        /* ── Work Process Inventory (from uploaded SAP profiles) ──
           Rendered BEFORE any Work Process charts; serves as the source table
           for subsequent WP visualisations. */
        this._renderWpInventory(sec);

        const data = this._appTierSeries(app);
        const { labels, hosts } = data;
        if (!hosts.length) {
            sec.insertAdjacentHTML('beforeend', `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">No application-tier time-series data found — Work Process charts require SMON/MON data. The inventory above is built from uploaded SAP profiles.</div></div>`);
            return;
        }

        const ci = this._colIndex || {};
        const findCol = want => {
            const w = want.toUpperCase().replace(/\s+/g, '');
            for (const k in ci) if (k.toUpperCase().replace(/\s+/g, '') === w) return k;
            return null;
        };
        const idOf = s => String(s).replace(/[^A-Za-z0-9]/g, '').toLowerCase();
        let seq = 1;

        const mkGroup = (id, title) => {
            const wrap = document.createElement('div');
            wrap.className = 'chart-group';
            wrap.innerHTML = `
                <div class="chart-group-head">
                    <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                    <span>${title}</span>
                </div>
                <div class="chart-group-body" id="${id}"></div>`;
            sec.appendChild(wrap);
            return wrap.querySelector(`#${id}`);
        };

        /* Helper: build { host: [values aligned to labels] } for a raw column. */
        const seriesFor = colName => {
            const resolved = findCol(colName);
            if (!resolved) return null;
            const r = this._stackedMemByHost(app, resolved, hosts, labels);
            const out = {};
            let any = false;
            hosts.forEach(h => {
                const d = (r.hostSeries.find(s => s.label === h) || {}).data;
                if (d && d.some(v => v != null)) { out[h] = d; any = true; }
            });
            return any ? out : null;
        };

        /* ── 1. Active Work Processes Stacked (ACT_WPS) ── */
        {
            const body = mkGroup('twa-wpgrp-actwps', 'Active Work Processes Stacked');
            const s = seriesFor('ACT_WPS');
            if (s) {
                this._mountStackedBarChart(body, {
                    id: 'wp-actwps-stacked', fileSection: 'wp', fileSeq: seq++,
                    title: 'Active Work Processes Stacked',
                    labels,
                    hostSeries: hosts.filter(h => s[h]).map(h => ({ label: h, data: s[h] })),
                    unit: 'WP', yLabel: 'Number of work processes', xLabel: 'Date / Time',
                    statsMetric: 'Active Work Processes',
                    hint: 'Stacked bars show the number of active work processes (ACT_WPS) per application server',
                });
            } else {
                body.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">No ACT_WPS data available.</div></div>`;
            }
        }

        /* ── 2. Active Dialog Work Processes Stacked (ACT_DIA) ── */
        {
            const body = mkGroup('twa-wpgrp-actdia', 'Active Dialog Work Processes Stacked');
            const s = seriesFor('ACT_DIA');
            if (s) {
                this._mountStackedBarChart(body, {
                    id: 'wp-actdia-stacked', fileSection: 'wp', fileSeq: seq++,
                    title: 'Active Dialog Work Processes Stacked',
                    labels,
                    hostSeries: hosts.filter(h => s[h]).map(h => ({ label: h, data: s[h] })),
                    unit: 'WP', yLabel: 'Number of work processes', xLabel: 'Date / Time',
                    statsMetric: 'Active Dialog Work Processes',
                    hint: 'Stacked bars show the number of active dialog work processes (ACT_DIA) per application server',
                });
            } else {
                body.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">No ACT_DIA data available.</div></div>`;
            }
        }

        /* ── 3. WP free for RFC — normal priority (TRFC_FREE) ── */
        {
            const body = mkGroup('twa-wpgrp-rfc', 'Work Processes free for RFC tasks (normal priority)');
            const s = seriesFor('TRFC_FREE');
            if (s) {
                let peak = 0;
                Object.values(s).forEach(arr => arr.forEach(v => { if (v != null && v > peak) peak = v; }));
                this._mountLineChart(body, {
                    id: 'wp-trfc-free', fileSection: 'wp', fileSeq: seq++,
                    title: 'Work Processes free for RFC tasks (normal priority)',
                    labels, series: s,
                    yLabel: 'WP free for RFC', yMin: 0, yMax: this._niceCeil(Math.max(10, peak * 1.1)),
                    xLabel: 'Date / Time',
                });
            } else {
                body.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">No TRFC_FREE data available.</div></div>`;
            }
        }

        /* ── 3b. WP free for RFC — low priority (SMON only, RFC_LOW) ── */
        if (this.sourceTable === 'SMON') {
            const sLow = seriesFor('RFC_LOW');
            if (sLow) {
                const body = mkGroup('twa-wpgrp-rfclow', 'Work Processes free for RFC tasks (low priority)');
                let peak = 0;
                Object.values(sLow).forEach(arr => arr.forEach(v => { if (v != null && v > peak) peak = v; }));
                this._mountLineChart(body, {
                    id: 'wp-rfc-low', fileSection: 'wp', fileSeq: seq++,
                    title: 'Work Processes free for RFC tasks (low priority)',
                    labels, series: sLow,
                    yLabel: 'WP free for RFC (low)', yMin: 0, yMax: this._niceCeil(Math.max(10, peak * 1.1)),
                    xLabel: 'Date / Time',
                });
            }
        }

        /* Wire group toggles. */
        sec.querySelectorAll('.chart-group-head').forEach(head => {
            head.addEventListener('click', () => head.closest('.chart-group').classList.toggle('collapsed'));
        });
    },

    /* ══════════════════════════════════════════════════════
       LOGINS & SESSIONS — dual-axis combo chart
       Bars = Σ USERS (Logins) per timestamp on the left axis,
       Line = Σ SESSIONS per timestamp on the right axis.
       Both axes share the same min/max bounds.
       ══════════════════════════════════════════════════════ */
    _renderLoginsCharts(app) {
        const sec = document.getElementById('twa-sec-logins');
        if (!sec) return;
        sec.innerHTML = '';

        const data = this._appTierSeries(app);
        const { labels, hosts } = data;
        if (!hosts.length) {
            sec.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">No application-tier servers found in the data.</div></div>`;
            return;
        }

        const ci = this._colIndex || {};
        const findCol = want => {
            const w = want.toUpperCase().replace(/\s+/g, '');
            for (const k in ci) if (k.toUpperCase().replace(/\s+/g, '') === w) return k;
            return null;
        };
        const colUsers = findCol('USERS');
        const colSess  = findCol('SESSIONS');
        if (!colUsers || !colSess) {
            sec.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">Logins/Sessions data (USERS / SESSIONS) is not present in this source.</div></div>`;
            return;
        }

        /* Sum USERS and SESSIONS across all app-tier hosts per timestamp label. */
        const uSer = this._stackedMemByHost(app, colUsers, hosts, labels);
        const sSer = this._stackedMemByHost(app, colSess,  hosts, labels);
        const n = labels.length;
        const logins = new Array(n).fill(0), sessions = new Array(n).fill(0);
        const hasData = new Array(n).fill(false);
        uSer.hostSeries.forEach(hs => hs.data.forEach((v, i) => { if (v != null) { logins[i] += v; hasData[i] = true; } }));
        sSer.hostSeries.forEach(hs => hs.data.forEach((v, i) => { if (v != null) { sessions[i] += v; hasData[i] = true; } }));
        /* Null out timestamps that had no data at all (keep the axis clean). */
        for (let i = 0; i < n; i++) if (!hasData[i]) { logins[i] = null; sessions[i] = null; }

        const wrap = document.createElement('div');
        wrap.className = 'chart-group';
        wrap.innerHTML = `
            <div class="chart-group-head">
                <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                <span>Active Logins and Sessions</span>
            </div>
            <div class="chart-group-body" id="twa-loginsgrp-body"></div>`;
        sec.appendChild(wrap);

        this._mountLoginsSessionsChart(wrap.querySelector('#twa-loginsgrp-body'), {
            id: 'logins-sessions', fileSection: 'logins', fileSeq: 1,
            title: 'Active Logins and Sessions',
            labels, logins, sessions, xLabel: 'Date / Time',
        });

        sec.querySelectorAll('.chart-group-head').forEach(head => {
            head.addEventListener('click', () => head.closest('.chart-group').classList.toggle('collapsed'));
        });
    },

    /* Dual-axis combo: Logins (bar, left axis) + Sessions (line, right axis).
       Both axes use the SAME min (0) and max so the two series are visually
       comparable. */
    _mountLoginsSessionsChart(container, cfg) {
        try { this.__mountLoginsSessionsChartInner(container, cfg); } catch (err) {
            console.error('[TechWorkloadApp] _mountLoginsSessionsChart failed for', cfg && cfg.id, err);
        }
    },
    __mountLoginsSessionsChartInner(container, cfg) {
        this._charts = this._charts || {};
        const uid = 'twa-ch-' + cfg.id;
        const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

        let peak = 0;
        cfg.logins.forEach(v => { if (v != null && v > peak) peak = v; });
        cfg.sessions.forEach(v => { if (v != null && v > peak) peak = v; });
        const yMax = peak > 0 ? this._niceCeil(peak * 1.1) : 100;

        const card = document.createElement('div');
        card.className = 'card chart-card-wrap';
        card.innerHTML = `
            <div class="chart-toolbar">
                <input class="chart-title-input" id="${uid}-title" value="${esc(cfg.title)}" title="Edit chart title">
                <div class="chart-controls">
                    <label>Y max <input type="number" id="${uid}-ymax" value="${yMax}" step="50"></label>
                    <button class="btn btn-sm" id="${uid}-reset" title="Reset axes">Reset</button>
                    <button class="btn btn-sm" id="${uid}-copy" title="Copy a high-resolution image to the clipboard">Copy</button>
                    <button class="btn btn-sm" id="${uid}-png" title="Download a high-resolution PNG">PNG</button>
                    <button class="btn btn-sm" id="${uid}-svg" title="Download a vector SVG">SVG</button>
                    <button class="btn btn-sm" id="${uid}-csv" title="Download chart data as CSV">CSV</button>
                    <button class="btn btn-sm chart-collapse-btn" id="${uid}-collapse" title="Hide / show this chart">–</button>
                </div>
            </div>
            <div class="chart-canvas-wrap chart-tall-wrap"><canvas id="${uid}-canvas"></canvas></div>
            <div class="chart-hint">Bars = total Logins (left axis) · line = total Sessions (right axis), summed across all application servers per snapshot · both axes share the same scale.</div>`;
        container.appendChild(card);

        const ctx = card.querySelector(`#${uid}-canvas`).getContext('2d');
        const chart = new Chart(ctx, {
            data: {
                labels: cfg.labels,
                datasets: [
                    {
                        type: 'bar', label: 'Logins', yAxisID: 'yLogins',
                        data: cfg.logins,
                        /* Fixed brand colours — must not be auto-substituted. */
                        backgroundColor: '#0A6CFF', borderColor: '#0A6CFF', borderWidth: 0,
                    },
                    {
                        type: 'line', label: 'Sessions', yAxisID: 'ySessions',
                        data: cfg.sessions,
                        borderColor: '#F4A62A', backgroundColor: '#F4A62A',
                        borderWidth: 1.6, pointRadius: 0, pointHoverRadius: 3, tension: 0.15, spanGaps: true,
                    },
                ],
            },
            options: {
                responsive: true, maintainAspectRatio: false, animation: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    legend: { position: 'bottom', labels: { font: { size: 10 }, padding: 12, usePointStyle: true } },
                    tooltip: { callbacks: { label: c => `${c.dataset.label}: ${Math.round(c.parsed.y).toLocaleString()}` } },
                },
                scales: {
                    x: {
                        title: { display: true, text: cfg.xLabel || 'Date / Time' },
                        ticks: { maxRotation: 60, autoSkip: true, maxTicksLimit: 20, font: { size: 9 } },
                        grid: { display: false },
                    },
                    yLogins: {
                        type: 'linear', position: 'left', min: 0, max: yMax,
                        title: { display: true, text: 'Logins' },
                        grid: { color: '#EEF2F7' }, ticks: { font: { size: 10 }, callback: v => v.toLocaleString() },
                    },
                    ySessions: {
                        type: 'linear', position: 'right', min: 0, max: yMax,
                        title: { display: true, text: 'Sessions' },
                        grid: { drawOnChartArea: false }, ticks: { font: { size: 10 }, callback: v => v.toLocaleString() },
                    },
                },
            },
        });
        this._charts[uid] = { chart, cfg: { ...cfg, _yMax: yMax, namedSeries: [
            { label: 'Logins', data: cfg.logins }, { label: 'Sessions', data: cfg.sessions },
        ] } };

        card.querySelector(`#${uid}-title`).addEventListener('input', e => { this._charts[uid].cfg.title = e.target.value; });
        const applyY = () => {
            const mx = parseFloat(card.querySelector(`#${uid}-ymax`).value);
            chart.options.scales.yLogins.max = isNaN(mx) ? undefined : mx;
            chart.options.scales.ySessions.max = isNaN(mx) ? undefined : mx;
            chart.update();
        };
        card.querySelector(`#${uid}-ymax`).addEventListener('change', applyY);
        card.querySelector(`#${uid}-reset`).addEventListener('click', () => {
            card.querySelector(`#${uid}-ymax`).value = yMax;
            chart.options.scales.yLogins.max = yMax;
            chart.options.scales.ySessions.max = yMax;
            chart.update();
        });
        card.querySelector(`#${uid}-copy`).addEventListener('click', () => this._copyChartPng(uid));
        card.querySelector(`#${uid}-png`).addEventListener('click', () => this._exportChartPng(uid));
        card.querySelector(`#${uid}-svg`).addEventListener('click', () => this._exportChartSvg(uid));
        card.querySelector(`#${uid}-csv`).addEventListener('click', () => this._exportChartCsv(uid));
        card.querySelector('.chart-canvas-wrap').insertAdjacentHTML('afterend', chartFontRowHtml(uid));
        wireChartFontRow(card, () => [this._charts[uid] && this._charts[uid].chart].filter(Boolean), null, false);
        this._wireChartCollapse(card, uid);
    },

    /* ══════════════════════════════════════════════════════
       NETWORK ANALYSIS  (SMON only — requires NETRTT column)
       ══════════════════════════════════════════════════════ */
    _renderNetworkCharts(app) {
        const sec = document.getElementById('twa-sec-network');
        if (!sec) return;
        sec.innerHTML = '';

        if (this.sourceTable !== 'SMON') {
            sec.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">
                Network analysis is only available for <strong>/SDF/SMON_HEADER</strong> data.</div></div>`;
            return;
        }

        const ci = this._colIndex || {};
        const findCol = want => {
            const w = want.toUpperCase().replace(/\s+/g, '');
            for (const k in ci) if (k.toUpperCase().replace(/\s+/g, '') === w) return k;
            return null;
        };
        const colRtt = findCol('NETRTT');
        if (colRtt == null) {
            sec.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">
                NETRTT column not found in this data set.</div></div>`;
            return;
        }

        const data = this._appTierSeries(app);
        const { hosts } = data;
        if (!hosts.length) {
            sec.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">
                No application-tier servers found in the data.</div></div>`;
            return;
        }

        /* Build per-server scatter points keyed by a NUMERIC timestamp (epoch ms),
           not a category string. A month of per-minute SMON data is hundreds of
           thousands of rows; a category x-axis would force Chart.js to index every
           distinct label and freeze the browser (the "stuck at 55%" symptom). A
           linear time axis plots the same points in a fraction of the time. We keep
           the FULL per-host series here (so the user can request a more granular
           view later) and compute per-host stats on the full data; the mount
           downsamples a working copy for display. */
        const iDate = ci['Date'];
        const iTime = ci['TIME'] != null ? ci['TIME'] : ci['Time'];
        const iSrv  = ci['SERVER'];
        const iRtt  = ci[colRtt];
        const rows  = this.targetRows;

        const fullMap = {};   /* host → [{x: epochMs, y, label}] (complete) */
        const stats   = {};   /* host → { sum, cnt, min, max } */
        hosts.forEach(h => { fullMap[h] = []; stats[h] = { sum: 0, cnt: 0, min: Infinity, max: -Infinity }; });
        let rttSum = 0, rttCnt = 0, tMin = Infinity, tMax = -Infinity;

        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const srv = r[iSrv];
            if (!srv || !fullMap[srv]) continue;
            const v = parseFloat(r[iRtt]);
            if (isNaN(v)) continue;
            const date = iDate != null ? r[iDate] : '';
            const time = iTime != null ? r[iTime] : '';
            const t = this._epochMs(date, time, i);
            fullMap[srv].push({ x: t, y: v, label: `${date} ${time}`.trim() });
            const st = stats[srv];
            st.sum += v; st.cnt++;
            if (v < st.min) st.min = v;
            if (v > st.max) st.max = v;
            if (t < tMin) tMin = t;
            if (t > tMax) tMax = t;
            rttSum += v;
            rttCnt++;
        }
        const avgRtt = rttCnt ? Math.round(rttSum / rttCnt) : 0;
        if (!rttCnt) {
            sec.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">
                No NETRTT values are present for the application-tier servers.</div></div>`;
            return;
        }

        /* Per-host summary rows (avg / min / max), only for hosts that have data. */
        const statRows = hosts
            .filter(h => stats[h].cnt > 0)
            .map(h => ({
                host: h,
                avg: Math.round(stats[h].sum / stats[h].cnt),
                min: stats[h].min,
                max: stats[h].max,
                cnt: stats[h].cnt,
            }));

        const wrap = document.createElement('div');
        wrap.className = 'chart-group';
        wrap.innerHTML = `
            <div class="chart-group-head">
                <svg class="chart-group-chevron" xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
                <span>RTT between the AS and the DB</span>
            </div>
            <div class="chart-group-body" id="twa-netgrp-body"></div>`;
        sec.appendChild(wrap);

        this._mountRttScatterChart(wrap.querySelector('#twa-netgrp-body'), {
            id: 'rtt-scatter', fileSection: 'network', fileSeq: 1,
            title: 'RTT between the AS and the DB',
            hosts, fullMap, statRows, avgRtt, tMin, tMax,
        });

        sec.querySelectorAll('.chart-group-head').forEach(head => {
            head.addEventListener('click', () => head.closest('.chart-group').classList.toggle('collapsed'));
        });
    },

    /* Parse an SMON Date ("DD/MM/YYYY" or "YYYY-MM-DD") + Time ("HH:MM:SS") into
       epoch ms for a numeric time axis. Falls back to the row index when the
       stamp can't be parsed, so points still spread out left-to-right. */
    _epochMs(date, time, rowIdx) {
        const d = String(date || '').trim();
        const t = String(time || '').trim();
        let y, mo, da;
        let m = d.match(/^(\d{4})-(\d{2})-(\d{2})$/);          /* YYYY-MM-DD */
        if (m) { y = +m[1]; mo = +m[2]; da = +m[3]; }
        else if ((m = d.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) { /* DD/MM/YYYY */
            da = +m[1]; mo = +m[2]; y = +m[3];
        } else {
            return rowIdx;   /* unparseable — keep insertion order */
        }
        const tm = t.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
        const hh = tm ? +tm[1] : 0, mi = tm ? +tm[2] : 0, ss = tm ? (+tm[3] || 0) : 0;
        const ms = Date.UTC(y, mo - 1, da, hh, mi, ss);
        return isNaN(ms) ? rowIdx : ms;
    },

    /* Uniformly downsample a per-host series map to at most `maxPts` points per
       host (keeps first/last, evenly strides the middle). Returns a NEW map. */
    _downsampleSeries(fullMap, hosts, maxPts) {
        const out = {};
        hosts.forEach(h => {
            const arr = fullMap[h] || [];
            if (arr.length <= maxPts) { out[h] = arr.slice(); return; }
            const step = arr.length / maxPts;
            const ds = [];
            for (let k = 0; k < maxPts; k++) ds.push(arr[Math.floor(k * step)]);
            out[h] = ds;
        });
        return out;
    },

    /* Scatter chart: all AS hosts as individual dot series + overall-average
       reference line. Point size and the number of plotted points per host are
       both adjustable at runtime; changing the point count re-samples from the
       FULL dataset (held in cfg.fullMap) and rebuilds the chart, with the chart
       progress bar shown because a more granular view can be heavy. */
    _mountRttScatterChart(container, cfg) {
        try { this.__mountRttScatterChartInner(container, cfg); } catch (err) {
            console.error('[TechWorkloadApp] _mountRttScatterChart failed for', cfg && cfg.id, err);
        }
    },
    __mountRttScatterChartInner(container, cfg) {
        this._charts = this._charts || {};
        const uid  = 'twa-ch-' + cfg.id;
        const esc  = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
        const { hosts, fullMap, statRows, avgRtt } = cfg;

        /* Peak RTT (over the full data) for the default Y-max. */
        let peak = 0;
        hosts.forEach(h => (fullMap[h] || []).forEach(pt => { if (pt.y > peak) peak = pt.y; }));
        const yMax = this._niceCeil(peak * 1.1) || 2000;

        /* Total available points and the initial per-host cap. */
        const totalPts = hosts.reduce((s, h) => s + (fullMap[h] ? fullMap[h].length : 0), 0);
        const maxHostLen = hosts.reduce((m, h) => Math.max(m, fullMap[h] ? fullMap[h].length : 0), 0);
        const DEFAULT_CAP = 6000;
        let ptCap   = Math.min(DEFAULT_CAP, maxHostLen) || maxHostLen || 1;
        let ptRadius = 2;

        const card = document.createElement('div');
        card.className = 'card chart-card-wrap';
        card.innerHTML = `
            <div class="chart-toolbar">
                <input class="chart-title-input" id="${uid}-title" value="${esc(cfg.title)}" title="Edit chart title">
                <div class="chart-controls">
                    <label>Y max <input type="number" id="${uid}-ymax" value="${yMax}" step="100"></label>
                    <label title="Dot size on the scatter">Dot <input type="number" id="${uid}-dot" value="${ptRadius}" min="0.5" max="10" step="0.5" style="width:52px"></label>
                    <button class="btn btn-sm" id="${uid}-dotminus" title="Reduce the dot size">Smaller dots</button>
                    <label title="Points plotted per server — raise for a more granular analysis">Points/server <input type="number" id="${uid}-pts" value="${ptCap}" min="10" max="${maxHostLen}" step="500" style="width:76px"></label>
                    <button class="btn btn-sm" id="${uid}-ptsapply" title="Apply the point count and rebuild the chart">Apply points</button>
                    <button class="btn btn-sm" id="${uid}-reset" title="Reset axes">Reset</button>
                    <button class="btn btn-sm" id="${uid}-copy" title="Copy a high-resolution image to the clipboard">Copy</button>
                    <button class="btn btn-sm" id="${uid}-png" title="Download a high-resolution PNG">PNG</button>
                    <button class="btn btn-sm" id="${uid}-svg" title="Download a vector SVG">SVG</button>
                    <button class="btn btn-sm" id="${uid}-csv" title="Download chart data as CSV">CSV</button>
                    <button class="btn btn-sm chart-collapse-btn" id="${uid}-collapse" title="Hide / show this chart">–</button>
                </div>
            </div>
            <div class="chart-canvas-wrap chart-tall-wrap"><canvas id="${uid}-canvas"></canvas></div>
            <div class="chart-hint" id="${uid}-hint"></div>
            <div class="net-avg-rtt" id="${uid}-avgrtt" style="margin:6px 0 2px;font-size:13px;font-weight:600">
                Average RTT for all servers: ${avgRtt.toLocaleString()} µs
            </div>
            <div class="net-xservers" id="${uid}-xservers"></div>
            <div class="net-stats-wrap">
                <div class="net-stats-title">Per-server RTT summary (µs)</div>
                <table class="data-table net-stats-table" id="${uid}-stats">
                    <thead><tr><th>Hostname</th><th>Average RTT</th><th>Minimum RTT</th><th>Maximum RTT</th><th>Status</th></tr></thead>
                    <tbody></tbody>
                </table>
                <div class="net-stats-legend" style="margin-top:8px;font-size:11px;color:var(--text-muted);display:flex;gap:16px;flex-wrap:wrap">
                    <span><span style="display:inline-block;width:11px;height:11px;border-radius:2px;background:${this._RTT_COLORS.healthy};vertical-align:middle;margin-right:4px"></span>Healthy (&lt; 350 µs)</span>
                    <span><span style="display:inline-block;width:11px;height:11px;border-radius:2px;background:${this._RTT_COLORS.monitor};vertical-align:middle;margin-right:4px"></span>Monitor (350–750 µs)</span>
                    <span><span style="display:inline-block;width:11px;height:11px;border-radius:2px;background:${this._RTT_COLORS.attention};vertical-align:middle;margin-right:4px"></span>Attention Required (&gt; 750 µs)</span>
                    <span><span style="color:#C0392B;font-weight:700">Red value</span> = well above the overall average</span>
                </div>
            </div>`;
        container.appendChild(card);

        const fmtTick = ms => {
            const d = new Date(ms);
            if (isNaN(d.getTime())) return '';
            const mon = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getUTCMonth()];
            return `${mon} ${d.getUTCDate()}`;
        };

        /* Build the per-server coloured chips shown beneath the Date / Time axis. */
        const renderXServers = () => {
            const el = card.querySelector(`#${uid}-xservers`);
            el.innerHTML = `<span class="net-xservers-lead">Servers:</span>` + hosts.map((h, idx) =>
                `<span class="net-xchip"><span class="net-xdot" style="background:${CHART_COLORS[idx % CHART_COLORS.length]}"></span>${esc(h)}</span>`
            ).join('');
        };

        /* Build the avg/min/max stats table rows with RTT conditional formatting:
           row background by threshold (healthy/monitor/attention) and a red value
           for servers whose average is well above the overall average. */
        const renderStats = () => {
            const tb = card.querySelector(`#${uid}-stats tbody`);
            if (!statRows || !statRows.length) {
                tb.innerHTML = `<tr><td colspan="5" style="color:var(--text-muted)">No data</td></tr>`;
                return;
            }
            tb.innerHTML = statRows.map(s => {
                const fmt = this._rttFormat(s.avg, avgRtt);
                /* Row-level colour: apply the threshold background to the row AND
                   every cell, so the whole row is coloured regardless of any
                   per-cell background in the table CSS. */
                const td = (inner, extra) => `<td style="background:${fmt.bg};${extra || ''}">${inner}</td>`;
                const avgStyle = fmt.outlier ? 'color:#C0392B;font-weight:700' : '';
                return `<tr style="background:${fmt.bg}">` +
                    td(esc(s.host)) +
                    td(s.avg.toLocaleString(), avgStyle) +
                    td(s.min.toLocaleString()) +
                    td(s.max.toLocaleString()) +
                    td(fmt.status) +
                    `</tr>`;
            }).join('');
        };

        const updateHint = (shownPerHost) => {
            const shownTotal = hosts.reduce((s, h) => s + Math.min(shownPerHost, fullMap[h] ? fullMap[h].length : 0), 0);
            card.querySelector(`#${uid}-hint`).innerHTML =
                `Each dot = one snapshot per application server · horizontal line = overall average ` +
                `(${avgRtt.toLocaleString()} µs) · showing <strong>${shownTotal.toLocaleString()}</strong> of ` +
                `${totalPts.toLocaleString()} points (${shownPerHost.toLocaleString()}/server max).`;
        };

        /* (Re)build the Chart.js datasets from a downsampled working map. */
        const buildDatasets = (workMap) => {
            const ds = hosts.map((h, idx) => ({
                type: 'scatter',
                label: h,
                data: (workMap[h] || []).map(pt => ({ x: pt.x, y: pt.y, l: pt.label })),
                backgroundColor: (CHART_COLORS[idx % CHART_COLORS.length] + 'AA'),
                borderColor:      CHART_COLORS[idx % CHART_COLORS.length],
                pointRadius: ptRadius,
                pointHoverRadius: ptRadius + 3,
            }));
            ds.push({
                type: 'line', label: 'Average',
                data: [{ x: cfg.tMin, y: avgRtt }, { x: cfg.tMax, y: avgRtt }],
                borderColor: '#555', backgroundColor: 'transparent',
                borderWidth: 1.4, borderDash: [4, 3], pointRadius: 0, pointHoverRadius: 0, tension: 0,
            });
            return ds;
        };

        let workMap = this._downsampleSeries(fullMap, hosts, ptCap);
        const ctx = card.querySelector(`#${uid}-canvas`).getContext('2d');
        const chart = new Chart(ctx, {
            type: 'scatter',
            data: { datasets: buildDatasets(workMap) },
            options: {
                responsive: true, maintainAspectRatio: false, animation: false,
                parsing: false,
                interaction: { mode: 'nearest', intersect: true },
                plugins: {
                    legend: { position: 'right', labels: { font: { size: 10 }, padding: 10, usePointStyle: true } },
                    tooltip: { callbacks: {
                        title: items => (items[0] && items[0].raw && items[0].raw.l) || '',
                        label: c => `${c.dataset.label}: ${Math.round(c.parsed.y).toLocaleString()} µs`,
                    }},
                },
                scales: {
                    x: {
                        type: 'linear', min: cfg.tMin, max: cfg.tMax,
                        title: { display: true, text: 'Date / Time' },
                        ticks: { maxRotation: 0, autoSkip: true, maxTicksLimit: 14, font: { size: 9 }, callback: v => fmtTick(v) },
                        grid: { display: false },
                    },
                    y: {
                        min: 0, max: yMax,
                        title: { display: true, text: 'Time (microseconds)' },
                        grid: { color: '#EEF2F7' },
                        ticks: { font: { size: 10 }, callback: v => v.toLocaleString() },
                    },
                },
            },
        });

        /* CSV reflects whatever is currently plotted. */
        const rebuildCsv = () => {
            const r = [];
            hosts.forEach(h => (workMap[h] || []).forEach(pt => r.push([pt.label, h, pt.y])));
            this._charts[uid].cfg._netCsvRows = r;
        };

        this._charts[uid] = { chart, cfg: { ...cfg, _yMax: yMax, _netCsvRows: [] } };
        renderXServers();
        renderStats();
        updateHint(ptCap);
        rebuildCsv();

        /* ── Dot-size: live update, no rebuild needed ── */
        const applyDot = (val) => {
            const r = parseFloat(val);
            if (isNaN(r) || r < 0.5) return;
            ptRadius = r;
            card.querySelector(`#${uid}-dot`).value = r;
            chart.data.datasets.forEach(d => { if (d.type === 'scatter') { d.pointRadius = r; d.pointHoverRadius = r + 3; } });
            chart.update();
        };
        card.querySelector(`#${uid}-dot`).addEventListener('change', e => applyDot(e.target.value));
        card.querySelector(`#${uid}-dotminus`).addEventListener('click', () => applyDot(Math.max(0.5, ptRadius - 0.5)));

        /* ── Point count: re-sample from the full data + rebuild (progress bar) ── */
        const applyPoints = async () => {
            const want = Math.max(10, Math.min(maxHostLen, parseInt(card.querySelector(`#${uid}-pts`).value, 10) || ptCap));
            card.querySelector(`#${uid}-pts`).value = want;
            ptCap = want;
            this._chartProgress('Re-sampling network points…', 20);
            await this._nextFrame();
            workMap = this._downsampleSeries(fullMap, hosts, ptCap);
            this._chartProgress('Rebuilding chart…', 65);
            await this._nextFrame();
            chart.data.datasets = buildDatasets(workMap);
            chart.update();
            rebuildCsv();
            updateHint(ptCap);
            this._chartProgress('Chart ready', 100);
            await this._nextFrame();
            setTimeout(() => this._hideChartProgress(), 300);
        };
        card.querySelector(`#${uid}-ptsapply`).addEventListener('click', () => applyPoints());
        card.querySelector(`#${uid}-pts`).addEventListener('keydown', e => { if (e.key === 'Enter') applyPoints(); });

        card.querySelector(`#${uid}-title`).addEventListener('input', e => { this._charts[uid].cfg.title = e.target.value; });
        card.querySelector(`#${uid}-ymax`).addEventListener('change', () => {
            const mx = parseFloat(card.querySelector(`#${uid}-ymax`).value);
            chart.options.scales.y.max = isNaN(mx) ? undefined : mx;
            chart.update();
        });
        card.querySelector(`#${uid}-reset`).addEventListener('click', () => {
            card.querySelector(`#${uid}-ymax`).value = yMax;
            chart.options.scales.y.max = yMax;
            chart.update();
        });
        card.querySelector(`#${uid}-copy`).addEventListener('click', () => this._copyChartPng(uid));
        card.querySelector(`#${uid}-png`).addEventListener('click',  () => this._exportChartPng(uid));
        card.querySelector(`#${uid}-svg`).addEventListener('click',  () => this._exportChartSvg(uid));
        card.querySelector(`#${uid}-csv`).addEventListener('click',  () => this._exportNetworkCsv(uid));
        card.querySelector('.chart-canvas-wrap').insertAdjacentHTML('afterend', chartFontRowHtml(uid));
        wireChartFontRow(card, () => [this._charts[uid] && this._charts[uid].chart].filter(Boolean), null, false);
        this._wireChartCollapse(card, uid);
    },

    /* Export the RTT scatter data: one row per (snapshot, server) with its RTT. */
    _exportNetworkCsv(uid) {
        const entry = this._charts && this._charts[uid];
        if (!entry) return;
        const cfg = entry.cfg;
        const rows = cfg._netCsvRows || [];
        const out = ['Date/Time,Server,NETRTT (microseconds)'];
        rows.forEach(([lbl, srv, y]) => out.push(`"${lbl}","${srv}",${y == null ? '' : y}`));
        const csv = out.join('\r\n');
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        this._downloadBlob(blob, `${this._chartFileBase(cfg)}.csv`, 'Chart data exported as CSV.');
    },

    /* Build per-server, time-aligned memory series for the application-tier hosts.
       Returns:
         { labels:["DD/MM/YYYY HH:MM",…],
           hosts: [host,…],
           physMem: { host: GB|null },   // constant per host (landscape)
           emCfg:   { host: GB|null },   // constant per host (user EM input)
           emAlloc: { host:[y|null,…] }, // EM Allocated (GB) time series
           freeMem: { host:[y|null,…] }, // Free Memory (+FS) time series
           freeLabel: "Free Memory + FS (GB)" | "Free Memory (GB)" } */
    _appTierMemSeries(app) {
        const ci = this._colIndex || {};
        const iDate = ci['Date'], iTime = ci['TIME'] != null ? ci['TIME'] : ci['Time'];
        const iSrv = ci['SERVER'], iEwa = ci['SERVER (EWA)'];
        const iPhys = ci['Physical Memory (GB)'];
        const iEmAlloc = ci['EM Allocated (GB)'];
        const iFreeFs = ci['Free Memory + FS (GB)'];
        const iFree   = ci['Free Memory (GB)'];
        const iFreeUse = iFreeFs != null ? iFreeFs : iFree;
        const freeLabel = iFreeFs != null ? 'Free Memory + FS (GB)' : 'Free Memory (GB)';
        const emCfgMap = this._extMemory || {};

        const empty = { labels: [], hosts: [], physMem: {}, emCfg: {}, emAlloc: {}, freeMem: {}, freeLabel };
        if (iDate == null || iSrv == null) return empty;

        const labelIndex = new Map();
        const labels = [];
        const hosts = [];
        const isApp = {};
        const physMem = {}, emCfg = {}, emAlloc = {}, freeMem = {};
        const rows = this.targetRows;

        /* Pass 1 — ordered labels + app-tier host list + per-host constants. */
        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const srv = r[iSrv];
            if (!srv) continue;
            if (isApp[srv] === undefined) {
                isApp[srv] = this._rowLayer(app, srv, iEwa != null ? r[iEwa] : null) === 'app';
            }
            if (!isApp[srv]) continue;
            const lbl = `${r[iDate] || ''} ${iTime != null ? (r[iTime] || '') : ''}`.trim();
            if (!labelIndex.has(lbl)) { labelIndex.set(lbl, labels.length); labels.push(lbl); }
            if (emAlloc[srv] === undefined) {
                hosts.push(srv);
                emAlloc[srv] = null;
                /* Physical memory: constant per host (prefer the built column, else landscape). */
                let pm = null;
                if (iPhys != null) { const p = parseFloat(r[iPhys]); if (!isNaN(p)) pm = p; }
                physMem[srv] = pm;
                /* EM Configured: the user-entered GB keyed by the virtual host (SERVER). */
                const g = emCfgMap[(srv || '').toLowerCase()];
                emCfg[srv] = (g == null || g === '') ? null : g;
            }
        }
        const n = labels.length;
        hosts.forEach(h => {
            emAlloc[h] = new Array(n).fill(null);
            freeMem[h] = new Array(n).fill(null);
        });

        /* Pass 2 — fill time-series values aligned to the label index. */
        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const srv = r[iSrv];
            if (!srv || !isApp[srv]) continue;
            const lbl = `${r[iDate] || ''} ${iTime != null ? (r[iTime] || '') : ''}`.trim();
            const pos = labelIndex.get(lbl);
            if (pos == null) continue;
            if (iEmAlloc != null) { const y = parseFloat(r[iEmAlloc]); if (!isNaN(y)) emAlloc[srv][pos] = y; }
            if (iFreeUse != null) { const y = parseFloat(r[iFreeUse]); if (!isNaN(y)) freeMem[srv][pos] = y; }
            /* Physical memory can also vary if the landscape value was blank; keep the first non-null. */
            if (physMem[srv] == null && iPhys != null) { const p = parseFloat(r[iPhys]); if (!isNaN(p)) physMem[srv] = p; }
        }
        return { labels, hosts, physMem, emCfg, emAlloc, freeMem, freeLabel };
    },

    /* Build a stacked per-host series for an arbitrary numeric column, aligned to
       the given label axis (Date + Time). Used for the Heap stacked bar chart.
       Returns { hostSeries: [{ label, data:[y|null,…] }] }. */
    _stackedMemByHost(app, colName, hosts, labels) {
        const ci = this._colIndex || {};
        const iDate = ci['Date'], iTime = ci['TIME'] != null ? ci['TIME'] : ci['Time'];
        const iSrv = ci['SERVER'], iCol = ci[colName];
        if (iDate == null || iSrv == null || iCol == null) {
            return { hostSeries: hosts.map(h => ({ label: h, data: new Array(labels.length).fill(null) })) };
        }
        const labelIdx = new Map(labels.map((l, i) => [l, i]));
        const n = labels.length;
        const series = {};
        hosts.forEach(h => { series[h] = new Array(n).fill(null); });

        const rows = this.targetRows;
        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const srv = r[iSrv];
            if (!srv || series[srv] === undefined) continue;
            const lbl = `${r[iDate] || ''} ${iTime != null ? (r[iTime] || '') : ''}`.trim();
            const pos = labelIdx.get(lbl);
            if (pos == null) continue;
            const y = parseFloat(r[iCol]);
            if (!isNaN(y)) series[srv][pos] = y;
        }
        return { hostSeries: hosts.map(h => ({ label: h, data: series[h] })) };
    },

    /* Aggregate a single host's data into an hourly-average CPU grid, exactly as
       an Excel pivot would: rows = Date (ascending), columns = Hour 0-23,
       value = AVG(CPU Utilization %). Minute snapshots are averaged into their
       (date,hour) bucket. Returns { dates:[...], z:[[h0..h23] per date] }. */
    _hourlyCpuGrid(app, hostName) {
        const ci = this._colIndex || {};
        const iDate = ci['Date'], iHours = ci['Hours'], iSrv = ci['SERVER'];
        const iCpu = ci['CPU Utilization (%)'];
        if (iDate == null || iHours == null || iCpu == null || iSrv == null) return { dates: [], z: [] };

        const rows = this.targetRows;
        /* date -> [24] accumulators {sum,count}. Dates are collected in first-seen
           order which is already chronological (rows are pre-sorted). */
        const sums = new Map();   /* dateStr -> Float64Array(24) sum */
        const cnts = new Map();   /* dateStr -> Float64Array(24) count */
        const dates = [];
        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            if (r[iSrv] !== hostName) continue;
            const date = r[iDate];
            const hr = parseInt(r[iHours], 10);
            const y = parseFloat(r[iCpu]);
            if (!date || isNaN(hr) || hr < 0 || hr > 23 || isNaN(y)) continue;
            if (!sums.has(date)) { sums.set(date, new Float64Array(24)); cnts.set(date, new Float64Array(24)); dates.push(date); }
            sums.get(date)[hr] += y;
            cnts.get(date)[hr] += 1;
        }
        const z = dates.map(d => {
            const s = sums.get(d), c = cnts.get(d);
            const row = new Array(24);
            for (let h = 0; h < 24; h++) row[h] = c[h] > 0 ? Math.round((s[h] / c[h]) * 100) / 100 : null;
            return row;
        });
        return { dates, z };
    },

    /* Hourly-average SAPS series derived from the same CPU-grid buckets as
       _hourlyCpuGrid so the computation is identical to the 3D surface chart.
       Formula: SAPSUtil_h = (avgCPU_h / 100) × SAPSDeployed
       where SAPSDeployed comes from the landscape (cores × sapsPerCore).
       Returns { labels: ["DD/MM/YYYY HH:00", …], util: […], dep: […] }.
       Hours with no CPU data are skipped (no null gaps on the X-axis). */
    _hourlySapsFromGrid(app, hostName) {
        /* Resolve the physical host record, honouring the virtual→physical mapping. */
        const physName = this._mappingActive
            ? (this._mapping[hostName.toLowerCase()] || null)
            : hostName;
        const hostRec  = app.findHost(physName || hostName);
        const sapsDep  = (hostRec && hostRec.cores && hostRec.sapsPerCore)
            ? hostRec.cores * hostRec.sapsPerCore
            : null;

        const grid = this._hourlyCpuGrid(app, hostName);
        if (!grid.dates.length) return { labels: [], util: [], dep: [] };

        const labels = [], util = [], dep = [];
        for (let di = 0; di < grid.dates.length; di++) {
            for (let h = 0; h < 24; h++) {
                const cpuAvg = grid.z[di][h];
                if (cpuAvg == null) continue;   /* skip hours with no snapshots */
                labels.push(`${grid.dates[di]} ${String(h).padStart(2, '0')}:00`);
                util.push(sapsDep != null ? Math.round((cpuAvg / 100) * sapsDep) : null);
                dep.push(sapsDep);
            }
        }
        return { labels, util, dep };
    },

    /* Build the data for the stacked SAPS bar chart (Group 5).
       One series per app-tier host: the hourly AVERAGE of the actual
       "SAPS Utilization (SAPS)" column (the per-minute values the file already
       carries), bucketed by (Date, Hour) — identical to the user's Excel pivot
       (AVERAGE of SAPS Utilization grouped by Date + Hour). Averaging each hour
       (not taking the max) is what matches the manually built chart.
       X-axis labels are the chronological union of all (DD/MM/YYYY HH:00) keys.
       Also returns totalDeployed = Σ SAPSDeployed across app-tier landscape hosts.
       Returns { labels:[], hostSeries:[{label, data:[]}], totalDeployed } */
    _hourlyStackedSaps(app, appHosts) {
        const ci = this._colIndex || {};
        const iDate = ci['Date'], iHours = ci['Hours'], iSrv = ci['SERVER'];
        const iSU = ci['SAPS Utilization (SAPS)'];

        const totalDeployed = app.getHostVariables()
            .filter(h => h.layer === 'app')
            .reduce((sum, h) => sum + (h.totalSaps || 0), 0);

        if (iDate == null || iHours == null || iSrv == null || iSU == null) {
            return { labels: [], hostSeries: [], totalDeployed };
        }

        /* Per host: (date HH:00) -> { sum, count } to AVERAGE the SAPS column. */
        const hostSet = new Set(appHosts);
        const sums = new Map();   /* host -> Map(key -> sum) */
        const cnts = new Map();   /* host -> Map(key -> count) */
        const labelSet = new Set();
        const rows = this.targetRows;
        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const srv = r[iSrv];
            if (!srv || !hostSet.has(srv)) continue;
            const date = r[iDate];
            const hr = parseInt(r[iHours], 10);
            if (!date || isNaN(hr) || hr < 0 || hr > 23) continue;
            const y = parseFloat(r[iSU]);
            if (isNaN(y)) continue;
            const key = `${date} ${String(hr).padStart(2, '0')}:00`;
            labelSet.add(key);
            if (!sums.has(srv)) { sums.set(srv, new Map()); cnts.set(srv, new Map()); }
            const sMap = sums.get(srv), cMap = cnts.get(srv);
            sMap.set(key, (sMap.get(key) || 0) + y);
            cMap.set(key, (cMap.get(key) || 0) + 1);
        }

        if (!labelSet.size) return { labels: [], hostSeries: [], totalDeployed };

        /* Chronological label order: key "DD/MM/YYYY HH:00" -> YYYYMMDDHH. */
        const toKey = lbl => {
            const [datePart, timePart] = lbl.split(' ');
            const dp = datePart.split('/');   /* DD, MM, YYYY */
            const hh = timePart ? timePart.split(':')[0] : '00';
            return `${dp[2]}${dp[1]}${dp[0]}${hh}`;
        };
        const labels = [...labelSet].sort((a, b) => toKey(a).localeCompare(toKey(b)));
        const labelIdx = new Map(labels.map((l, i) => [l, i]));
        const n = labels.length;

        const hostSeries = appHosts
            .filter(h => sums.has(h))
            .map(h => {
                const sMap = sums.get(h), cMap = cnts.get(h);
                const dataArr = new Array(n).fill(null);
                sMap.forEach((sum, key) => {
                    const pos = labelIdx.get(key);
                    if (pos != null) dataArr[pos] = Math.round(sum / cMap.get(key));
                });
                return { label: h, data: dataArr };
            })
            .filter(s => s.data.some(v => v != null));

        return { labels, hostSeries, totalDeployed };
    },

    /* ──────────────────────────────────────────────────────
       RELEVANT PEAK ANALYSIS  (CPU / Heap / EM — per-record %)
       ──────────────────────────────────────────────────────
       Extracts per-record utilisation peaks, consolidates near-identical
       measurements of the same event, prioritises representative peaks spread
       across the timeline (different days first, then different times), and
       returns ≤20 results sorted by Date/Time DESCENDING.

       metricCfg = {
         valueOf(row, ci) → number|null   // the % metric for a row (null to skip)
       }
       Returns [{ dt:Date, dateStr, timeStr, label, host, value }] (≤20, desc).
    ────────────────────────────────────────────────────── */
    _CONSOLIDATE_MS: 2 * 60 * 60 * 1000,   /* events within 2h on the same day + similar level → one event */
    _SIMILAR_PCT:    5,                     /* ≤5 percentage-points apart counts as the same utilisation level */

    /* Parse the DD/MM/YYYY date + HH:MM[:SS] time columns into a Date. */
    _parsePeakDateTime(dateStr, timeStr) {
        const d = String(dateStr || '').trim();
        const t = String(timeStr || '').trim();
        const dm = d.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
        if (!dm) return null;
        let [, dd, mm, yy] = dm;
        dd = parseInt(dd, 10); mm = parseInt(mm, 10); yy = parseInt(yy, 10);
        if (yy < 100) yy += 2000;
        const tm = t.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
        const hh = tm ? parseInt(tm[1], 10) : 0;
        const mi = tm ? parseInt(tm[2], 10) : 0;
        const ss = tm && tm[3] ? parseInt(tm[3], 10) : 0;
        const dt = new Date(yy, mm - 1, dd, hh, mi, ss);
        return isNaN(dt) ? null : dt;
    },

    /* Build a stacked-bar dataset from the top-N relevant peaks array.
       peaks     — output of _relevantPeaks (date-descending on entry, sorted here)
       valueCol  — enriched table column name to use for the bar height (e.g.
                   'SAPS Utilization (SAPS)' or 'EM Allocated (GB)')
       appHosts  — ordered list of app-tier server names

       Returns { labels:[], hostSeries:[{label, data:[]}] } with labels in
       chronological ascending order.  For each peak timestamp, looks up every
       app-tier server's value from targetRows (exact date+time+server match;
       null when no row is found for that server at that moment). */
    _peaksToStackedSeries(peaks, valueCol, appHosts) {
        const empty = { labels: [], hostSeries: [] };
        if (!peaks.length || !appHosts.length) return empty;

        const ci = this._colIndex || {};
        const iDate = ci['Date'];
        const iTime = ci['TIME'] != null ? ci['TIME'] : ci['Time'];
        const iSrv  = ci['SERVER'];
        const iVal  = ci[valueCol];
        if (iDate == null || iSrv == null || iVal == null) return empty;

        /* Sort peaks chronologically (ascending). */
        const sorted = peaks.slice().sort((a, b) => a.dt - b.dt);
        const labels = sorted.map(p => p.label);

        /* Build a lookup: "dateStr\x01timeStr\x01host" -> value */
        const lookup = new Map();
        const peakSet = new Set(sorted.map(p => `${p.dateStr}\x01${p.timeStr}`));
        const hostSet = new Set(appHosts);
        const rows = this.targetRows;
        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const srv = r[iSrv];
            if (!srv || !hostSet.has(srv)) continue;
            const dateStr = String(r[iDate] || '').trim();
            const timeStr = iTime != null ? String(r[iTime] || '').trim() : '';
            const dtKey = `${dateStr}\x01${timeStr}`;
            if (!peakSet.has(dtKey)) continue;
            const v = parseFloat(r[iVal]);
            if (!isNaN(v)) {
                lookup.set(`${dtKey}\x01${srv}`, v);
            }
        }

        const hostSeries = appHosts.map(h => ({
            label: h,
            data: sorted.map(p => {
                const key = `${p.dateStr}\x01${p.timeStr}\x01${h}`;
                const v = lookup.get(key);
                return v !== undefined ? v : null;
            }),
        })).filter(s => s.data.some(v => v != null));

        return { labels, hostSeries };
    },

    _relevantPeaks(app, metricCfg, maxN = 20) {
        const ci = this._colIndex || {};
        const iDate = ci['Date'];
        const iTime = ci['TIME'] != null ? ci['TIME'] : ci['Time'];
        const iSrv  = ci['SERVER'], iEwa = ci['SERVER (EWA)'];
        const rows  = this.targetRows || [];
        if (iDate == null || iSrv == null || !rows.length) return [];

        /* 1. Extract all valid per-record peaks (app-tier servers only). */
        const isApp = {};
        const all = [];
        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            const srv = r[iSrv];
            if (!srv) continue;
            if (isApp[srv] === undefined) {
                isApp[srv] = this._rowLayer(app, srv, iEwa != null ? r[iEwa] : null) === 'app';
            }
            if (!isApp[srv]) continue;
            const value = metricCfg.valueOf(r, ci);
            if (value == null || !isFinite(value) || value <= 0) continue;
            const dateStr = r[iDate];
            const timeStr = iTime != null ? r[iTime] : '';
            const dt = this._parsePeakDateTime(dateStr, timeStr);
            if (!dt) continue;
            /* rawValue: absolute value used for the "guarantee top N" seed.
               Falls back to value when not provided. */
            const rawValue = metricCfg.rawValueOf ? metricCfg.rawValueOf(r, ci) : value;
            all.push({ dt, dateStr: String(dateStr).trim(), timeStr: String(timeStr || '').trim(),
                       label: `${String(dateStr).trim()} ${String(timeStr || '').trim()}`.trim(),
                       host: srv, value, rawValue: rawValue ?? value });
        }
        if (!all.length) return [];

        /* 2. Rank candidates by utilisation descending (strongest events first). */
        all.sort((a, b) => b.value - a.value);

        const dayKey = p => `${p.dt.getFullYear()}-${p.dt.getMonth()}-${p.dt.getDate()}`;
        const selected = [];
        const usedDays = new Set();

        /* Consolidation test: a candidate belongs to an already-selected event if
           it is the same host, same day, within the time window AND a similar
           utilisation level. */
        const sameEvent = (c, s) =>
            c.host === s.host &&
            dayKey(c) === dayKey(s) &&
            Math.abs(c.dt - s.dt) <= this._CONSOLIDATE_MS &&
            Math.abs(c.value - s.value) <= this._SIMILAR_PCT;

        const notConsolidated = c => !selected.some(s => sameEvent(c, s));

        /* 2b. Guarantee seed: if metricCfg.guaranteeTop is set, pre-populate
              `selected` with the top-N entries by raw absolute value (not ratio/%).
              This ensures the absolute largest peaks always appear in the chart,
              regardless of the temporal-distribution algorithm. */
        if (metricCfg.guaranteeTop > 0) {
            const byRaw = all.slice().sort((a, b) => b.rawValue - a.rawValue);
            for (const c of byRaw) {
                if (selected.length >= metricCfg.guaranteeTop) break;
                if (notConsolidated(c)) {
                    selected.push(c);
                    usedDays.add(dayKey(c));
                }
            }
        }

        /* 3. Pass 1 — one representative peak per DISTINCT DAY (highest that day),
              prioritising coverage across the timeline. */
        for (const c of all) {
            if (selected.length >= maxN) break;
            const dk = dayKey(c);
            if (usedDays.has(dk)) continue;
            selected.push(c);
            usedDays.add(dk);
        }

        /* 4. Pass 2 — if slots remain, add the next strongest peaks at DIFFERENT
              times, skipping any that consolidate into an already-selected event. */
        if (selected.length < maxN) {
            for (const c of all) {
                if (selected.length >= maxN) break;
                if (selected.includes(c)) continue;
                if (notConsolidated(c)) selected.push(c);
            }
        }

        /* 5. Final ordering: Date/Time ASCENDING (oldest left), capped at maxN. */
        selected.sort((a, b) => a.dt - b.dt);
        return selected.slice(0, maxN);
    },

    /* Render a "Top N Relevant Peaks" column chart + source table.
       cfg = { id, fileSection, fileSeq, title, peaks:[…], valueLabel, color } */
    _mountRelevantPeaks(container, cfg) {
        try { this.__mountRelevantPeaksInner(container, cfg); } catch (err) {
            console.error('[TechWorkloadApp] _mountRelevantPeaks failed for', cfg && cfg.id, err);
        }
    },
    __mountRelevantPeaksInner(container, cfg) {
        this._charts = this._charts || {};
        const uid = 'twa-ch-' + cfg.id;
        const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
        const peaks = cfg.peaks || [];
        const valueLabel = cfg.valueLabel || 'Utilization (%)';
        const color = cfg.color || '#FFC000';

        if (!peaks.length) {
            container.innerHTML = `<div class="card"><div style="padding:16px;color:var(--text-muted);font-size:13px">No ${esc(valueLabel)} peaks to report.</div></div>`;
            return;
        }

        const labels = peaks.map(p => p.label);
        const values = peaks.map(p => Math.round(p.value * 100) / 100);
        const yMax = this._niceCeil(Math.max(...values) * 1.1) || 100;

        const card = document.createElement('div');
        card.className = 'card chart-card-wrap';
        card.innerHTML = `
            <div class="chart-toolbar">
                <input class="chart-title-input" id="${uid}-title" value="${esc(cfg.title)}" title="Edit chart title">
                <div class="chart-controls">
                    <label>Y max <input type="number" id="${uid}-ymax" value="${yMax}" step="5"></label>
                    <label class="chart-xlabels-toggle"><input type="checkbox" id="${uid}-xlabels" checked> X labels</label>
                    <button class="btn btn-sm" id="${uid}-reset" title="Reset axes">Reset</button>
                    <button class="btn btn-sm" id="${uid}-copy" title="Copy a high-resolution image to the clipboard">Copy</button>
                    <button class="btn btn-sm" id="${uid}-png" title="Download a high-resolution PNG">PNG</button>
                    <button class="btn btn-sm" id="${uid}-svg" title="Download a vector SVG">SVG</button>
                    <button class="btn btn-sm" id="${uid}-csv" title="Download chart data as CSV">CSV</button>
                    <button class="btn btn-sm chart-collapse-btn" id="${uid}-collapse" title="Hide / show this chart">–</button>
                </div>
            </div>
            <div class="chart-canvas-wrap chart-tall-wrap"><canvas id="${uid}-canvas"></canvas></div>
            <div class="chart-hint">Representative ${esc(valueLabel)} peaks across the analysis period (duplicates &amp; near-identical timestamps consolidated) · sorted by Date/Time ascending</div>
            <div class="card" style="padding:0;margin-top:12px">
                <div class="table-wrap" style="border:none;border-radius:12px;overflow-x:auto">
                    <table class="data-table">
                        <thead><tr><th>Date/Time</th><th>Hostname</th><th style="text-align:right">${esc(valueLabel)}</th></tr></thead>
                        <tbody>${peaks.map(p => `<tr><td>${esc(p.label)}</td><td>${esc(p.host)}</td><td style="text-align:right">${(Math.round(p.value * 100) / 100).toLocaleString(undefined, { maximumFractionDigits: 2 })}</td></tr>`).join('')}</tbody>
                    </table>
                </div>
            </div>`;
        container.appendChild(card);

        const ctx = card.querySelector(`#${uid}-canvas`).getContext('2d');
        const chart = new Chart(ctx, {
            type: 'bar',
            data: { labels, datasets: [{ label: valueLabel, data: values, backgroundColor: color, borderColor: color, borderWidth: 0 }] },
            options: {
                responsive: true, maintainAspectRatio: false, animation: false,
                plugins: {
                    legend: { display: false },
                    tooltip: { callbacks: {
                        title: items => items.length ? `${peaks[items[0].dataIndex].label} · ${peaks[items[0].dataIndex].host}` : '',
                        label: ctx => `${valueLabel}: ${ctx.parsed.y.toLocaleString(undefined, { maximumFractionDigits: 2 })}`,
                    } },
                },
                scales: {
                    x: { title: { display: true, text: 'Date / Time' },
                         ticks: { maxRotation: 60, autoSkip: false, font: { size: 9 } }, grid: { display: false } },
                    y: { title: { display: true, text: valueLabel }, min: 0, max: yMax,
                         grid: { color: '#EEF2F7' }, ticks: { font: { size: 10 } } },
                },
            },
        });
        this._charts[uid] = { chart, cfg: { ...cfg, _yMax: yMax } };

        card.querySelector(`#${uid}-title`).addEventListener('input', e => { this._charts[uid].cfg.title = e.target.value; });
        const applyY = () => { const mx = parseFloat(card.querySelector(`#${uid}-ymax`).value); chart.options.scales.y.max = isNaN(mx) ? undefined : mx; chart.update(); };
        card.querySelector(`#${uid}-ymax`).addEventListener('change', applyY);
        card.querySelector(`#${uid}-xlabels`).addEventListener('change', e => { chart.options.scales.x.ticks.display = e.target.checked; chart.update(); });
        card.querySelector(`#${uid}-reset`).addEventListener('click', () => {
            card.querySelector(`#${uid}-ymax`).value = yMax;
            card.querySelector(`#${uid}-xlabels`).checked = true;
            chart.options.scales.y.max = yMax;
            chart.options.scales.x.ticks.display = true;
            if (chart.resetZoom) chart.resetZoom();
            chart.update();
        });
        card.querySelector(`#${uid}-copy`).addEventListener('click', () => this._copyChartPng(uid));
        card.querySelector(`#${uid}-png`).addEventListener('click', () => this._exportChartPng(uid));
        card.querySelector(`#${uid}-svg`).addEventListener('click', () => this._exportChartSvg(uid));
        card.querySelector(`#${uid}-csv`).addEventListener('click', () => {
            const lines = [['Date/Time', 'Hostname', valueLabel].join(',')];
            peaks.forEach(p => lines.push([`"${p.label}"`, `"${p.host}"`, Math.round(p.value * 100) / 100].join(',')));
            const blob = new Blob([lines.join('\r\n')], { type: 'text/csv' });
            const url  = URL.createObjectURL(blob);
            const a = Object.assign(document.createElement('a'), { href: url, download: `${this._sidTag()}_${cfg.id}.csv` });
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(url), 1000);
            showToast('CSV downloaded', '', 'success', 2500);
        });
        card.querySelector('.chart-canvas-wrap').insertAdjacentHTML('afterend', chartFontRowHtml(uid));
        wireChartFontRow(card, () => [this._charts[uid] && this._charts[uid].chart].filter(Boolean), null, false);
        this._wireChartCollapse(card, uid);
    },

    /* ──────────────────────────────────────────────────────
       STACKED BAR CHART with an overlaid flat line
       Used for Group 5: combined SAPS utilization across all AS
       + a red reference line for total deployed capacity.
       cfg = { id, title, labels[], hostSeries[{label,data}],
               totalDeployed, yLabel, xLabel, fileSection, fileSeq }
       ────────────────────────────────────────────────────── */
    _mountStackedBarChart(container, cfg) {
        try { this.__mountStackedBarChartInner(container, cfg); } catch (err) {
            console.error('[TechWorkloadApp] _mountStackedBarChart failed for', cfg && cfg.id, err);
        }
    },
    __mountStackedBarChartInner(container, cfg) {
        this._charts = this._charts || {};
        const uid = 'twa-ch-' + cfg.id;
        const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
        const unit = cfg.unit || 'SAPS';
        const hasRef = cfg.totalDeployed != null;   /* draw the red capacity line only when provided */

        /* Compute a sensible Y-axis ceiling. */
        const stackTotals = cfg.labels.map((_, li) =>
            cfg.hostSeries.reduce((s, hs) => s + (hs.data[li] != null ? hs.data[li] : 0), 0)
        );
        const stackPeak = stackTotals.reduce((m, v) => v > m ? v : m, 0);
        const refPeak   = hasRef ? (cfg.totalDeployed || 0) : 0;
        const rawMax    = Math.max(stackPeak, refPeak);
        const yMax      = rawMax > 0 ? this._niceCeil(rawMax * 1.1) : 1000;

        /* ── Peak statistics documented below the chart ──
           • Highest stacked peak (value + the date/time bucket it occurred in).
           • Average of the highest stacked peaks (mean of the top-N bar totals),
             which characterises the sustained busy periods rather than one spike. */
        const fmtNum = v => unit === 'GB'
            ? Number(v).toLocaleString(undefined, { maximumFractionDigits: 2 })
            : Math.round(v).toLocaleString();
        let statsHtml = '';
        {
            let maxIdx = -1, maxVal = -Infinity;
            for (let i = 0; i < stackTotals.length; i++) {
                if (stackTotals[i] > maxVal) { maxVal = stackTotals[i]; maxIdx = i; }
            }
            if (maxIdx >= 0 && maxVal > 0) {
                const TOP_N = cfg.topNPeaks || 10;
                /* "Average of the top N peaks": by default the mean of the top-N
                   stacked bar totals shown here. When cfg.avgTopPeaksOverride is
                   supplied (e.g. the EM/Heap peak charts, which must average the N
                   highest RAW values across the whole analyzed period rather than
                   the ≤20 displayed bars), use that pre-computed value and keep the
                   label N in sync with it. */
                const topN = stackTotals.slice().sort((a, b) => b - a).slice(0, Math.min(TOP_N, stackTotals.length));
                const hasOverride = cfg.avgTopPeaksOverride != null && !isNaN(cfg.avgTopPeaksOverride);
                const avgTop = hasOverride ? Number(cfg.avgTopPeaksOverride) : (topN.reduce((s, v) => s + v, 0) / topN.length);
                const labelN = hasOverride ? (cfg.avgTopPeaksCount || TOP_N) : topN.length;
                const avgSub = hasOverride
                    ? `mean of the ${labelN} highest ${esc(cfg.statsMetric || 'values')} in the period`
                    : `mean of the ${labelN} highest stacked bars`;
                const when = cfg.labels[maxIdx] || '';
                const metric = cfg.statsMetric || 'stacked total';
                statsHtml = `
                    <div class="chart-stats">
                        <div class="chart-stats-title">Peak analysis — ${esc(metric)}</div>
                        <div class="chart-stats-grid">
                            <div class="chart-stat">
                                <div class="chart-stat-label">Highest stacked peak</div>
                                <div class="chart-stat-value">${fmtNum(maxVal)} ${esc(unit)}</div>
                                <div class="chart-stat-sub">on ${esc(when)}</div>
                            </div>
                            <div class="chart-stat">
                                <div class="chart-stat-label">Average of the top ${labelN} peaks</div>
                                <div class="chart-stat-value">${fmtNum(avgTop)} ${esc(unit)}</div>
                                <div class="chart-stat-sub">${avgSub}</div>
                            </div>
                        </div>
                    </div>`;
            }
        }

        const card = document.createElement('div');
        card.className = 'card chart-card-wrap';
        card.innerHTML = `
            <div class="chart-toolbar">
                <input class="chart-title-input" id="${uid}-title" value="${esc(cfg.title)}" title="Edit chart title">
                <div class="chart-controls">
                    <label>Y max <input type="number" id="${uid}-ymax" value="${yMax}" step="100"></label>
                    <button class="btn btn-sm" id="${uid}-reset" title="Reset axes">Reset</button>
                    <button class="btn btn-sm" id="${uid}-copy" title="Copy a high-resolution image to the clipboard">Copy</button>
                    <button class="btn btn-sm" id="${uid}-png" title="Download a high-resolution PNG">PNG</button>
                    <button class="btn btn-sm" id="${uid}-svg" title="Download a vector SVG">SVG</button>
                    <button class="btn btn-sm" id="${uid}-csv" title="Download chart data as CSV">CSV</button>
                    <button class="btn btn-sm chart-collapse-btn" id="${uid}-collapse" title="Hide / show this chart">–</button>
                </div>
            </div>
            <div class="chart-canvas-wrap chart-tall-wrap"><canvas id="${uid}-canvas"></canvas></div>
            <div class="chart-hint">${esc(cfg.hint || ('Stacked bars show hourly-average SAPS utilization per server · red line = total deployed SAPS capacity on the application layer'))} · <strong>click a server in the legend to show/hide it</strong></div>
            ${statsHtml}`;
        container.appendChild(card);

        const ctx = card.querySelector(`#${uid}-canvas`).getContext('2d');
        const n   = cfg.labels.length;

        /* One stacked bar dataset per host, auto-coloured. */
        const barDatasets = cfg.hostSeries.map((hs, i) => ({
            type: 'bar',
            label: hs.label,
            data: hs.data,
            backgroundColor: CHART_COLORS[i % CHART_COLORS.length] + 'CC',  /* slight transparency */
            borderColor:     CHART_COLORS[i % CHART_COLORS.length],
            borderWidth: 0,
            stack: 'stackall',
        }));

        /* Flat reference line for total deployed capacity (optional). */
        const datasets = [...barDatasets];
        if (hasRef) {
            datasets.push({
                type: 'line',
                label: `Total Deployed Capacity — App Layer (${(cfg.totalDeployed || 0).toLocaleString()} ${unit})`,
                data: new Array(n).fill(cfg.totalDeployed || 0),
                borderColor: '#EF4444',
                backgroundColor: 'transparent',
                borderWidth: 2.5,
                borderDash: [6, 4],
                pointRadius: 0,
                pointHoverRadius: 0,
                tension: 0,
                spanGaps: true,
                order: 0,   /* draw on top of bars */
            });
        }

        const chart = new Chart(ctx, {
            type: 'bar',
            data: { labels: cfg.labels, datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                animation: false,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            font: { size: 10 },
                            padding: 12,
                            generateLabels(ch) {
                                return Chart.defaults.plugins.legend.labels.generateLabels(ch).map(l => {
                                    const ds = ch.data.datasets[l.datasetIndex];
                                    /* Draw the deployed-capacity legend entry as a dashed line swatch. */
                                    if (ds && ds.type === 'line') {
                                        l.lineDash = [6, 4];
                                        l.lineWidth = 2.5;
                                        l.pointStyle = 'line';
                                    }
                                    if (ds && ds.hidden) l.fontColor = '#94A3B8';
                                    return l;
                                });
                            },
                        },
                        onHover(e) { e.native && (e.native.target.style.cursor = 'pointer'); },
                        onLeave(e) { e.native && (e.native.target.style.cursor = 'default'); },
                    },
                    tooltip: {
                        mode: 'index',
                        callbacks: {
                            label: ctx => {
                                const v = ctx.parsed.y;
                                if (v == null || isNaN(v)) return null;
                                const val = unit === 'SAPS' ? Math.round(v).toLocaleString() : v.toLocaleString(undefined, { maximumFractionDigits: 2 });
                                return `${ctx.dataset.label}: ${val} ${unit}`;
                            },
                            footer: items => {
                                const stackSum = items
                                    .filter(i => i.dataset.type === 'bar')
                                    .reduce((s, i) => s + (i.parsed.y || 0), 0);
                                if (!stackSum) return [];
                                const val = unit === 'SAPS' ? Math.round(stackSum).toLocaleString() : stackSum.toLocaleString(undefined, { maximumFractionDigits: 2 });
                                return [`Total: ${val} ${unit}`];
                            },
                        },
                    },
                },
                scales: {
                    x: {
                        stacked: true,
                        title: { display: true, text: cfg.xLabel || 'Date / Hour' },
                        ticks: { maxRotation: 60, autoSkip: true, maxTicksLimit: 20, font: { size: 9 } },
                        grid: { display: false },
                    },
                    y: {
                        stacked: true,
                        title: { display: true, text: cfg.yLabel || 'SAPS Utilization (SAPS)' },
                        min: 0,
                        max: yMax,
                        grid: { color: '#EEF2F7' },
                        ticks: { font: { size: 10 }, callback: v => v.toLocaleString() },
                    },
                },
            },
        });
        this._charts[uid] = { chart, cfg: { ...cfg, _yMax: yMax } };

        /* Title editing */
        card.querySelector(`#${uid}-title`).addEventListener('input', e => {
            this._charts[uid].cfg.title = e.target.value;
        });
        /* Y-max input */
        const applyY = () => {
            const mx = parseFloat(card.querySelector(`#${uid}-ymax`).value);
            chart.options.scales.y.max = isNaN(mx) ? undefined : mx;
            chart.update();
        };
        card.querySelector(`#${uid}-ymax`).addEventListener('change', applyY);
        /* Reset */
        card.querySelector(`#${uid}-reset`).addEventListener('click', () => {
            card.querySelector(`#${uid}-ymax`).value = yMax;
            chart.options.scales.y.max = yMax;
            chart.data.datasets.forEach((ds, i) => {
                ds.hidden = false;
                const meta = chart.getDatasetMeta(i);
                if (meta) meta.hidden = false;
            });
            if (chart.resetZoom) chart.resetZoom();
            chart.update();
        });
        /* Copy / PNG / SVG export — reuse the existing line-chart helpers
           (_copyChartPng / _exportChartPng / _exportChartSvg all key off
           this._charts[uid], so the bar chart entry works the same way). */
        card.querySelector(`#${uid}-copy`).addEventListener('click', () => this._copyChartPng(uid));
        card.querySelector(`#${uid}-png`).addEventListener('click', () => this._exportChartPng(uid));
        card.querySelector(`#${uid}-svg`).addEventListener('click', () => this._exportChartSvg(uid));
        card.querySelector(`#${uid}-csv`).addEventListener('click', () => this._exportStackedCsv(uid));
        card.querySelector('.chart-canvas-wrap').insertAdjacentHTML('afterend', chartFontRowHtml(uid));
        wireChartFontRow(card, () => [this._charts[uid] && this._charts[uid].chart].filter(Boolean), null, false);
        this._wireChartCollapse(card, uid);
    },

    /* Hourly-average SAPS series for one host: one point per (Date, Hour) bucket,
       in chronological order. SAPS Utilization is averaged across the minute
       snapshots in each hour; SAPS Deployed is taken as the bucket's value (it
       is constant per host, so its hourly average equals that constant).
       Returns { labels:["DD/MM/YYYY HH:00",…], util:[…], dep:[…] }. */
    _hourlySaps(hostName) {
        const ci = this._colIndex || {};
        const iDate = ci['Date'], iHours = ci['Hours'], iSrv = ci['SERVER'];
        const iSU = ci['SAPS Utilization (SAPS)'], iSD = ci['SAPS Deployed (SAPS)'];
        if (iDate == null || iHours == null || iSrv == null || iSU == null) {
            return { labels: [], util: [], dep: [] };
        }
        const rows = this.targetRows;
        const keyOrder = [];
        const seen = new Set();
        const uSum = new Map(), uCnt = new Map(), dVal = new Map();
        for (let i = 0; i < rows.length; i++) {
            const r = rows[i];
            if (r[iSrv] !== hostName) continue;
            const date = r[iDate];
            const hr = parseInt(r[iHours], 10);
            if (!date || isNaN(hr) || hr < 0 || hr > 23) continue;
            const key = `${date} ${String(hr).padStart(2, '0')}:00`;
            if (!seen.has(key)) { seen.add(key); keyOrder.push(key); uSum.set(key, 0); uCnt.set(key, 0); }
            const u = parseFloat(r[iSU]);
            if (!isNaN(u)) { uSum.set(key, uSum.get(key) + u); uCnt.set(key, uCnt.get(key) + 1); }
            if (iSD != null) { const d = parseFloat(r[iSD]); if (!isNaN(d)) dVal.set(key, d); }
        }
        const labels = keyOrder, util = [], dep = [];
        for (const k of keyOrder) {
            const c = uCnt.get(k);
            util.push(c > 0 ? Math.round((uSum.get(k) / c)) : null);
            dep.push(dVal.has(k) ? dVal.get(k) : null);
        }
        return { labels, util, dep };
    },

    /* Discrete 10% CPU colour bands (matches the Excel pivot legend). */
    _cpuBandColorscale() {
        /* 11 stops → 10 bands (0-10 … 90-100). Blues for low, warm for high. */
        const colors = [
            '#4472C4', /* 0-10  */ '#C0504D', /* 10-20 */ '#808080', /* 20-30 */
            '#F2C80F', /* 30-40 */ '#4472C4', /* 40-50 */ '#70AD47', /* 50-60 */
            '#5B9BD5', /* 60-70 */ '#ED7D31', /* 70-80 */ '#A5A5A5', /* 80-90 */
            '#FFC000', /* 90-100*/
        ];
        const scale = [];
        for (let b = 0; b < 10; b++) {
            scale.push([b / 10, colors[b]]);
            scale.push([(b + 1) / 10, colors[b]]);
        }
        return { scale, colors };
    },

    /* Round a value up to a tidy axis ceiling that stays close to the value
       (uses 2 significant digits with 1/2/2.5/5 steps), so a deployed-SAPS line
       sits comfortably near the top without a huge empty gap above it. */
    _niceCeil(v) {
        if (!(v > 0)) return 100;
        const exp = Math.floor(Math.log10(v)) - 1;   /* 2 significant digits */
        const base = Math.pow(10, exp);
        const f = v / base;                           /* 10 ≤ f < 100 */
        const steps = [10, 12, 15, 20, 25, 30, 40, 50, 60, 75, 80, 100];
        let step = 100;
        for (const s of steps) { if (f <= s) { step = s; break; } }
        return Math.round(step * base);
    },

    /* Build the export file base name per the required convention:
         <appSID>_wa_app_<section>_<seq>_<DD-MM-YY>
       e.g. "EPR_wa_app_cpu_1_02-10-26". SID comes from Module 01; falls back
       to "sap" if unavailable. */
    _chartFileBase(cfg) {
        const sid = this._sidTag();
        const now = new Date();
        const dd = String(now.getDate()).padStart(2, '0');
        const mm = String(now.getMonth() + 1).padStart(2, '0');
        const yy = String(now.getFullYear()).slice(-2);
        const section = cfg.fileSection || 'app';
        const seq = cfg.fileSeq || 1;
        return `${sid}_wa_app_${section}_${seq}_${dd}-${mm}-${yy}`;
    },

    /* The first 3 characters of the application-layer SID (from Module 01's
       System Landscape), lower-cased and alphanumeric. Used as the leading token
       of every exported file name. Falls back to 'sap' when unknown. */
    _sidTag() {
        const d = this.app && this.app.getModuleData && this.app.getModuleData('system-landscape');
        const sid = (d && d.systemInfo && d.systemInfo.sid) || '';
        return String(sid).replace(/[^A-Za-z0-9]/g, '').slice(0, 3).toLowerCase() || 'sap';
    },

    /* ──────────────────────────────────────────────────────
       INTERACTIVE 3D CPU SURFACE (Plotly) — one per host
       X = Hour of Day (0-23), Y = Date (oldest→newest),
       Z = avg CPU Utilization (%). Discrete 10% colour bands.
       ────────────────────────────────────────────────────── */
    _mountSurfaceChart(container, cfg) {
        this._surfaces = this._surfaces || {};
        const uid = 'twa-sf-' + cfg.id;
        const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

        const card = document.createElement('div');
        card.className = 'card chart-card-wrap';
        card.innerHTML = `
            <div class="chart-toolbar">
                <input class="chart-title-input" id="${uid}-title" value="${esc(cfg.title)}" title="Edit chart title">
                <div class="chart-controls">
                    <label>Z min <input type="number" id="${uid}-zmin" value="0" step="5"></label>
                    <label>Z max <input type="number" id="${uid}-zmax" value="100" step="5"></label>
                    <button class="btn btn-sm" id="${uid}-reset" title="Reset the 3D view, title &amp; Z-axis">Reset</button>
                    <button class="btn btn-sm" id="${uid}-copy" title="Copy a high-resolution image to the clipboard">Copy</button>
                    <button class="btn btn-sm" id="${uid}-png" title="Download a high-resolution PNG">PNG</button>
                    <button class="btn btn-sm" id="${uid}-svg" title="Download a vector SVG (for PowerPoint)">SVG</button>
                    <button class="btn btn-sm" id="${uid}-csv" title="Download chart data as CSV">CSV</button>
                    <button class="btn btn-sm chart-collapse-btn" id="${uid}-collapse" title="Hide / show this chart">–</button>
                </div>
            </div>
            <div class="chart-canvas-wrap chart-3d-wrap" id="${uid}-plot"></div>
            <div class="chart-hint">Drag to rotate · scroll to zoom · right-drag to pan · hover a point for Hostname / Date / Hour / avg CPU% · Reset restores the view.</div>`;
        container.appendChild(card);

        const plotEl = card.querySelector(`#${uid}-plot`);
        const bands = this._cpuBandColorscale();

        const x = Array.from({ length: 24 }, (_, h) => h);
        const y = cfg.grid.dates;
        const z = cfg.grid.z;
        const hostName = cfg.hostName;

        const trace = {
            type: 'surface',
            x, y, z,
            cmin: 0, cmax: 100,
            colorscale: bands.scale,
            connectgaps: true,
            contours: { z: { show: false } },
            colorbar: {
                title: { text: 'CPU %', side: 'right' },
                tickmode: 'array',
                tickvals: [5, 15, 25, 35, 45, 55, 65, 75, 85, 95],
                ticktext: ['0-10', '10-20', '20-30', '30-40', '40-50', '50-60', '60-70', '70-80', '80-90', '90-100'],
                tickfont: { size: 10 },
                len: 0.85,
            },
            hovertemplate:
                `<b>${esc(hostName)}</b><br>Date: %{y}<br>Hour: %{x}:00<br>Avg CPU: %{z:.1f}%<extra></extra>`,
        };

        const layout = {
            autosize: true,
            margin: { l: 0, r: 0, t: 10, b: 0 },
            scene: {
                xaxis: { title: { text: 'Hour of Day' }, dtick: 2, tickfont: { size: 10 } },
                yaxis: { title: { text: 'Date' }, tickfont: { size: 9 } },
                zaxis: { title: { text: 'CPU Utilization (%)' }, range: [0, 100], tickfont: { size: 10 } },
                camera: { eye: { x: 1.9, y: -1.5, z: 0.9 } },
                aspectratio: { x: 1.4, y: 1.2, z: 0.7 },
            },
        };
        const config = { displayModeBar: false, responsive: true };

        Plotly.newPlot(plotEl, [trace], layout, config);
        this._surfaces[uid] = { plotEl, cfg, defaultCamera: JSON.parse(JSON.stringify(layout.scene.camera)) };

        card.querySelector(`#${uid}-title`).addEventListener('input', e => {
            this._surfaces[uid].cfg.title = e.target.value;
        });
        const applyZ = () => {
            const mn = parseFloat(card.querySelector(`#${uid}-zmin`).value);
            const mx = parseFloat(card.querySelector(`#${uid}-zmax`).value);
            Plotly.relayout(plotEl, { 'scene.zaxis.range': [isNaN(mn) ? 0 : mn, isNaN(mx) ? 100 : mx] });
        };
        card.querySelector(`#${uid}-zmin`).addEventListener('change', applyZ);
        card.querySelector(`#${uid}-zmax`).addEventListener('change', applyZ);
        card.querySelector(`#${uid}-reset`).addEventListener('click', () => {
            card.querySelector(`#${uid}-zmin`).value = 0;
            card.querySelector(`#${uid}-zmax`).value = 100;
            Plotly.relayout(plotEl, { 'scene.zaxis.range': [0, 100], 'scene.camera': this._surfaces[uid].defaultCamera });
        });
        card.querySelector(`#${uid}-copy`).addEventListener('click', () => this._copySurface(uid));
        card.querySelector(`#${uid}-png`).addEventListener('click', () => this._exportSurface(uid, 'png'));
        card.querySelector(`#${uid}-svg`).addEventListener('click', () => this._exportSurface(uid, 'svg'));
        card.querySelector(`#${uid}-csv`).addEventListener('click', () => this._exportSurfaceCsv(uid));
        this._wireChartCollapse(card, uid);
    },

    /* Per-chart collapse: the chevron/– button hides the canvas + hint, leaving
       just the toolbar, so the user can shrink a long page of charts. */
    _wireChartCollapse(card, uid) {
        const btn = card.querySelector(`#${uid}-collapse`);
        if (!btn) return;
        btn.addEventListener('click', () => {
            const collapsed = card.classList.toggle('chart-collapsed');
            btn.textContent = collapsed ? '+' : '–';
            btn.title = collapsed ? 'Show this chart' : 'Hide this chart';
        });
    },

    _copySurface(uid) {
        const e = this._surfaces && this._surfaces[uid];
        if (!e) return;
        Plotly.toImage(e.plotEl, { format: 'png', scale: 3, width: 1100, height: 680 })
            .then(dataUrl => this._composeTitleAndCopy(dataUrl, e.cfg.title, 1100 * 3, 680 * 3))
            .catch(() => showToast('Copy failed', 'Could not render the 3D chart image.', 'error'));
    },

    _exportSurface(uid, fmt) {
        const e = this._surfaces && this._surfaces[uid];
        if (!e) return;
        const base = this._chartFileBase(e.cfg);
        if (fmt === 'svg') {
            Plotly.toImage(e.plotEl, { format: 'svg', width: 1200, height: 720 })
                .then(dataUrl => {
                    let svg = decodeURIComponent(dataUrl.replace(/^data:image\/svg\+xml,/, ''));
                    svg = this._injectSvgTitle(svg, e.cfg.title);
                    this._downloadBlob(new Blob([svg], { type: 'image/svg+xml;charset=utf-8' }), `${base}.svg`, '');
                    showToast('Chart exported (SVG)', 'Vector SVG saved. In PowerPoint 2016/365: Insert → Pictures, then Convert to Shape to edit.', 'success');
                })
                .catch(() => showToast('Export failed', 'Could not render the SVG.', 'error'));
        } else {
            Plotly.toImage(e.plotEl, { format: 'png', scale: 4, width: 1100, height: 680 })
                .then(dataUrl => this._composeTitleToBlob(dataUrl, e.cfg.title, 1100 * 4, 680 * 4)
                    .then(blob => this._downloadBlob(blob, `${base}.png`, 'High-resolution PNG saved.')))
                .catch(() => showToast('Export failed', 'Could not render the PNG.', 'error'));
        }
    },

    _composeTitleCanvas(dataUrl, title, w, h) {
        return new Promise((resolve, reject) => {
            const img = new Image();
            img.onload = () => {
                const pad = Math.round(h * 0.09);
                const canvas = document.createElement('canvas');
                canvas.width = w; canvas.height = h + pad;
                const ctx = canvas.getContext('2d');
                ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
                ctx.fillStyle = '#1E293B';
                ctx.font = `700 ${Math.round(pad * 0.5)}px Arial, sans-serif`;
                ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
                ctx.fillText(title || '', w / 2, pad / 2);
                ctx.drawImage(img, 0, pad, w, h);
                resolve(canvas);
            };
            img.onerror = reject;
            img.src = dataUrl;
        });
    },
    _composeTitleToBlob(dataUrl, title, w, h) {
        return this._composeTitleCanvas(dataUrl, title, w, h)
            .then(canvas => new Promise(res => canvas.toBlob(res, 'image/png')));
    },
    _composeTitleAndCopy(dataUrl, title, w, h) {
        this._composeTitleToBlob(dataUrl, title, w, h).then(blob => {
            if (!blob) { showToast('Copy failed', 'Could not render the image.', 'error'); return; }
            if (navigator.clipboard && window.ClipboardItem) {
                navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })])
                    .then(() => showToast('Chart copied', 'High-resolution 3D chart (with title) copied to your clipboard.', 'success'))
                    .catch(() => this._downloadBlob(blob, `chart-3d-${_dateTag()}.png`, 'Copy blocked — downloaded a PNG instead.'));
            } else {
                this._downloadBlob(blob, `chart-3d-${_dateTag()}.png`, 'Clipboard not supported — downloaded a PNG instead.');
            }
        });
    },

    _injectSvgTitle(svg, title) {
        const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        const m = svg.match(/width="(\d+(?:\.\d+)?)"/);
        const w = m ? parseFloat(m[1]) : 1200;
        const titleTag = `<text x="${w / 2}" y="34" text-anchor="middle" font-family="Arial, sans-serif" font-size="22" font-weight="700" fill="#1E293B">${esc(title)}</text>`;
        return svg.replace(/(<svg[^>]*>)/, `$1${titleTag}`);
    },

    /* ──────────────────────────────────────────────────────
       Generic interactive line-chart card with:
         • editable title
         • X-range (start/end index) + Y-range (min/max) inputs
         • drag-zoom, wheel-zoom, pan, and Reset
         • Copy image (PNG → clipboard) and Export SVG (vector)
       cfg = { id, title, labels[], series{host:[y]}, yLabel, yMin, yMax, xLabel }
       ────────────────────────────────────────────────────── */
    _mountLineChart(container, cfg) {
        try { this.__mountLineChartInner(container, cfg); } catch (err) {
            console.error('[TechWorkloadApp] _mountLineChart failed for', cfg && cfg.id, err);
        }
    },
    __mountLineChartInner(container, cfg) {
        this._charts = this._charts || {};
        const uid = 'twa-ch-' + cfg.id;
        const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

        const card = document.createElement('div');
        card.className = 'card chart-card-wrap';
        card.innerHTML = `
            <div class="chart-toolbar">
                <input class="chart-title-input" id="${uid}-title" value="${esc(cfg.title)}" title="Edit chart title">
                <div class="chart-controls">
                    <label>Y min <input type="number" id="${uid}-ymin" value="${cfg.yMin}" step="1"></label>
                    <label>Y max <input type="number" id="${uid}-ymax" value="${cfg.yMax}" step="1"></label>
                    <button class="btn btn-sm" id="${uid}-reset" title="Reset zoom &amp; axes">Reset</button>
                    <button class="btn btn-sm" id="${uid}-copy" title="Copy a high-resolution image (with title) to the clipboard">Copy</button>
                    <button class="btn btn-sm" id="${uid}-png" title="Download a high-resolution PNG (works in any PowerPoint)">PNG</button>
                    <button class="btn btn-sm" id="${uid}-svg" title="Download a vector SVG (editable in PowerPoint 2016/365)">SVG</button>
                    <button class="btn btn-sm" id="${uid}-csv" title="Download chart data as CSV">CSV</button>
                    <button class="btn btn-sm chart-collapse-btn" id="${uid}-collapse" title="Hide / show this chart">–</button>
                </div>
            </div>
            <div class="chart-canvas-wrap${cfg.tall ? ' chart-tall-wrap' : ''}"><canvas id="${uid}-canvas"></canvas></div>
            <div class="chart-hint">Drag across the plot to zoom · scroll to zoom · hold Shift and drag to pan · <strong>click a server in the legend to show/hide it</strong> · Reset restores the original view and shows all servers.</div>
            ${cfg.analysisHtml || ''}`;
        container.appendChild(card);

        const ctx = card.querySelector(`#${uid}-canvas`).getContext('2d');
        /* Datasets come either from `namedSeries` (explicit label/data/color per
           line — used for the SAPS util vs deployed chart) or from `series`
           (host→data map, auto-coloured). */
        const specs = cfg.namedSeries
            ? cfg.namedSeries.map((s, i) => ({ label: s.label, data: s.data, color: s.color || CHART_COLORS[i % CHART_COLORS.length] }))
            : Object.keys(cfg.series).map((hostName, i) => ({ label: hostName, data: cfg.series[hostName], color: CHART_COLORS[i % CHART_COLORS.length] }));
        const datasets = specs.map(s => {
            const color = s.color;
            return {
                label: s.label,
                data: s.data,
                borderColor: color,
                backgroundColor: color,
                borderWidth: cfg.lineWidth || 1.4,
                pointRadius: 0,
                pointHoverRadius: 3,
                tension: 0.15,
                spanGaps: true,
                pointStyle: 'line',   /* legend marker drawn as a line, not a dot */
                hidden: false,
            };
        });

        const chart = new Chart(ctx, {
            type: 'line',
            data: { labels: cfg.labels, datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                animation: false,
                parsing: true,
                normalized: true,
                interaction: { mode: 'nearest', intersect: false },
                plugins: {
                    legend: {
                        position: 'bottom',
                        /* line markers + pointer cursor to signal click-to-toggle */
                        labels: {
                            usePointStyle: true,
                            pointStyle: 'line',
                            boxWidth: 36,
                            boxHeight: 3,
                            font: { size: 10 },
                            padding: 12,
                            /* dim the swatch + strike the text for hidden series */
                            generateLabels(ch) {
                                const base = Chart.defaults.plugins.legend.labels.generateLabels(ch);
                                base.forEach(l => {
                                    const ds = ch.data.datasets[l.datasetIndex];
                                    if (ds && ds.hidden) { l.fontColor = '#94A3B8'; }
                                });
                                return base;
                            },
                        },
                        onHover(e) { e.native && (e.native.target.style.cursor = 'pointer'); },
                        onLeave(e) { e.native && (e.native.target.style.cursor = 'default'); },
                    },
                    tooltip: { enabled: true, callbacks: {} },
                    zoom: {
                        zoom: {
                            drag: { enabled: true, backgroundColor: 'rgba(0,112,242,0.12)' },
                            wheel: { enabled: true },
                            mode: 'xy',
                        },
                        pan: { enabled: true, mode: 'xy', modifierKey: 'shift' },
                    },
                },
                scales: {
                    x: {
                        title: { display: true, text: cfg.xLabel || '' },
                        ticks: { maxRotation: 60, autoSkip: true, maxTicksLimit: 16, font: { size: 9 } },
                        grid: { display: false },
                    },
                    y: {
                        title: { display: true, text: cfg.yLabel || '' },
                        min: cfg.yMin, max: cfg.yMax,
                        grid: { color: '#EEF2F7' },
                    },
                },
            },
        });
        this._charts[uid] = { chart, cfg };

        /* Title editing */
        card.querySelector(`#${uid}-title`).addEventListener('input', e => {
            this._charts[uid].cfg.title = e.target.value;
        });
        /* Y-range inputs */
        const applyY = () => {
            const mn = parseFloat(card.querySelector(`#${uid}-ymin`).value);
            const mx = parseFloat(card.querySelector(`#${uid}-ymax`).value);
            chart.options.scales.y.min = isNaN(mn) ? undefined : mn;
            chart.options.scales.y.max = isNaN(mx) ? undefined : mx;
            chart.update();
        };
        card.querySelector(`#${uid}-ymin`).addEventListener('change', applyY);
        card.querySelector(`#${uid}-ymax`).addEventListener('change', applyY);
        /* Reset — axes, zoom, AND show every server again. */
        card.querySelector(`#${uid}-reset`).addEventListener('click', () => {
            card.querySelector(`#${uid}-ymin`).value = cfg.yMin;
            card.querySelector(`#${uid}-ymax`).value = cfg.yMax;
            chart.options.scales.y.min = cfg.yMin;
            chart.options.scales.y.max = cfg.yMax;
            /* Un-hide all datasets so every server is shown again. */
            chart.data.datasets.forEach((ds, i) => {
                ds.hidden = false;
                const meta = chart.getDatasetMeta(i);
                if (meta) meta.hidden = false;
            });
            if (chart.resetZoom) chart.resetZoom();
            chart.update();
        });
        /* Copy image (hi-res PNG, with title) to clipboard */
        card.querySelector(`#${uid}-copy`).addEventListener('click', () => this._copyChartPng(uid));
        /* Download hi-res PNG */
        card.querySelector(`#${uid}-png`).addEventListener('click', () => this._exportChartPng(uid));
        /* Download vector SVG */
        card.querySelector(`#${uid}-svg`).addEventListener('click', () => this._exportChartSvg(uid));
        /* Download chart data as CSV */
        card.querySelector(`#${uid}-csv`).addEventListener('click', () => this._exportChartCsv(uid));
        card.querySelector('.chart-canvas-wrap').insertAdjacentHTML('afterend', chartFontRowHtml(uid));
        wireChartFontRow(card, () => [this._charts[uid] && this._charts[uid].chart].filter(Boolean), null, false);
        this._wireChartCollapse(card, uid);
    },

    /* Copy a HIGH-RESOLUTION image of the chart (including its title) to the
       clipboard. We rasterize the vector SVG at 3× so the clipboard image is
       crisp, and — crucially — includes the title text that lives outside the
       canvas in an <input>. */
    _copyChartPng(uid) {
        const entry = this._charts && this._charts[uid];
        if (!entry) return;
        const scale = 4;   /* 4× of a 1280-wide vector ≈ 5120px — very crisp */
        const base = this._chartFileBase(entry.cfg);

        /* Produce the PNG blob from the chart's vector SVG. */
        const makeBlob = () => this._svgToCanvas(entry, scale).then(canvas =>
            new Promise((resolve, reject) => {
                canvas.toBlob(b => b ? resolve(b) : reject(new Error('render failed')), 'image/png');
            })
        );

        /* Fallback used whenever the clipboard path is unavailable or rejected. */
        const fallbackDownload = (blob, why) =>
            this._downloadBlob(blob, `${base}.png`, why || 'Copy blocked by the browser — downloaded a high-res PNG instead.');

        if (navigator.clipboard && window.ClipboardItem) {
            /* Pass a Promise<Blob> to ClipboardItem and call write() immediately
               (still within the click gesture). This is the pattern browsers
               (Chrome/Edge/Safari) accept — resolving the blob first and writing
               later often loses the user-gesture and silently fails. */
            let captured = null;
            const item = new ClipboardItem({
                'image/png': makeBlob().then(b => { captured = b; return b; }),
            });
            /* Ensure the document is focused — clipboard.write rejects otherwise. */
            try { window.focus(); } catch (_) {}
            Promise.resolve(navigator.clipboard.write([item]))
                .then(() => showToast('Chart copied', 'Chart image copied — paste it into Word, PowerPoint, or an email.', 'success'))
                .catch(err => {
                    /* Retry once by writing an already-resolved blob (some builds
                       dislike the promise form), then fall back to a download. */
                    const retry = captured ? Promise.resolve(captured) : makeBlob();
                    retry.then(b => {
                        navigator.clipboard.write([new ClipboardItem({ 'image/png': b })])
                            .then(() => showToast('Chart copied', 'Chart image copied — paste it into Word, PowerPoint, or an email.', 'success'))
                            .catch(() => fallbackDownload(b, `Copy blocked (${err && err.name ? err.name : 'clipboard'}) — downloaded a high-res PNG instead.`));
                    }).catch(() => showToast('Copy failed', 'Could not render the chart image.', 'error'));
                });
        } else {
            makeBlob()
                .then(b => fallbackDownload(b, 'Clipboard image not supported in this browser — downloaded a high-res PNG instead.'))
                .catch(() => showToast('Copy failed', 'Could not render the chart image.', 'error'));
        }
    },

    /* Rasterize the LIVE Chart.js chart (not a re-synthesized SVG) onto a white
       canvas at the given scale, with the user-edited title drawn top-center
       above it. Rendering the real canvas preserves the exact chart type,
       stacking, colours, legend, and axes — so stacked column charts stay
       stacked columns (the previous SVG path redrew everything as lines).
       Returns a Promise<HTMLCanvasElement> so existing .then() callers work. */
    _svgToCanvas(entry, scale) {
        try {
            return Promise.resolve(this._chartToCanvas(entry, scale));
        } catch (err) {
            return Promise.reject(err);
        }
    },

    /* Draw the live chart canvas + centered title onto a hi-res backing canvas. */
    _chartToCanvas(entry, scale) {
        const src = entry.chart && entry.chart.canvas;
        if (!src) throw new Error('chart canvas unavailable');

        /* Read the live, user-edited title from the toolbar input (fallback cfg). */
        const uid = 'twa-ch-' + entry.cfg.id;
        const titleEl = document.getElementById(`${uid}-title`);
        const title = (titleEl && titleEl.value) || entry.cfg.title || '';

        const titleSize = ChartFontSizeManager.get().title || 18;
        const titleH = Math.round(titleSize * 1.9);   /* band tall enough for the title */

        const canvas = document.createElement('canvas');
        canvas.width  = src.width * scale;
        canvas.height = src.height * scale + titleH * scale;
        const ctx = canvas.getContext('2d');
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'high';

        /* White background (charts are drawn on transparent canvases). */
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);

        /* Title — top-center, horizontally centered over the chart area. */
        if (title) {
            ctx.fillStyle = '#1E293B';
            ctx.font = `700 ${titleSize * scale}px Arial, Helvetica, sans-serif`;
            ctx.textBaseline = 'top';
            ctx.textAlign = 'center';
            ctx.fillText(title, canvas.width / 2, Math.round(8 * scale));
            ctx.textAlign = 'left';
        }

        /* The live chart, drawn below the title band. */
        ctx.drawImage(src, 0, titleH * scale, src.width * scale, src.height * scale);
        return canvas;
    },

    _downloadBlob(blob, name, toastMsg) {
        const url = URL.createObjectURL(blob);
        const a = Object.assign(document.createElement('a'), { href: url, download: name });
        document.body.appendChild(a); a.click(); document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        if (toastMsg) showToast('Chart saved', toastMsg, 'success');
    },

    /* Export the underlying data of a line chart as CSV. */
    _exportChartCsv(uid) {
        const entry = this._charts && this._charts[uid];
        if (!entry) return;
        const cfg = entry.cfg;
        const labels = cfg.labels || [];

        // Collect series — cfg.namedSeries preferred, fall back to cfg.series (object map)
        let series = [];
        if (cfg.namedSeries && Array.isArray(cfg.namedSeries)) {
            series = cfg.namedSeries.map(s => ({ label: s.label, data: s.data }));
        } else if (cfg.series && typeof cfg.series === 'object') {
            series = Object.entries(cfg.series).map(([label, data]) => ({ label, data }));
        }

        const cols = ['Date/Time', ...series.map(s => s.label)];
        const rows = [cols.join(',')];
        labels.forEach((lbl, i) => {
            const vals = series.map(s => {
                const v = s.data ? s.data[i] : '';
                return v == null ? '' : v;
            });
            rows.push([`"${lbl}"`, ...vals].join(','));
        });

        const csv = rows.join('\r\n');
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        this._downloadBlob(blob, `${this._chartFileBase(cfg)}.csv`, 'Chart data exported as CSV.');
    },

    /* Export the underlying data of a 3D surface chart as CSV (Date × 24 hours). */
    _exportSurfaceCsv(uid) {
        const entry = this._surfaces && this._surfaces[uid];
        if (!entry) return;
        const cfg = entry.cfg;
        const grid = cfg.grid;
        if (!grid || !grid.dates || !grid.z) return;

        const hourHeaders = Array.from({ length: 24 }, (_, h) => `Hour ${String(h).padStart(2, '0')}:00`);
        const cols = ['Date', ...hourHeaders];
        const rows = [cols.join(',')];
        grid.dates.forEach((date, di) => {
            const zRow = grid.z[di] || [];
            const vals = Array.from({ length: 24 }, (_, h) => {
                const v = zRow[h];
                return v == null ? '' : v;
            });
            rows.push([`"${date}"`, ...vals].join(','));
        });

        const csv = rows.join('\r\n');
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        this._downloadBlob(blob, `${this._chartFileBase(cfg)}.csv`, 'Chart data exported as CSV.');
    },

    /* Export the underlying data of a stacked bar chart as CSV. */
    _exportStackedCsv(uid) {
        const entry = this._charts && this._charts[uid];
        if (!entry) return;
        const cfg = entry.cfg;
        const labels = cfg.labels || [];
        const hostSeries = cfg.hostSeries || [];
        const hasRef = cfg.totalDeployed != null;
        const unit = cfg.unit || 'SAPS';

        const cols = ['Date/Hour', ...hostSeries.map(s => s.label)];
        if (hasRef) cols.push(`Total Deployed Capacity (${unit})`);
        const rows = [cols.join(',')];
        labels.forEach((lbl, i) => {
            const vals = hostSeries.map(s => {
                const v = s.data ? s.data[i] : '';
                return v == null ? '' : v;
            });
            const row = [`"${lbl}"`, ...vals];
            if (hasRef) row.push(cfg.totalDeployed || 0);
            rows.push(row.join(','));
        });

        const csv = rows.join('\r\n');
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        this._downloadBlob(blob, `${this._chartFileBase(cfg)}.csv`, 'Chart data exported as CSV.');
    },

    /* Export the chart as an SVG that embeds a high-resolution raster of the
       LIVE chart (type/stacking/colours/legend/axes preserved exactly). The
       previous hand-built vector redrew every series as a line, which turned
       stacked column charts into line charts. */
    _exportChartSvg(uid) {
        const entry = this._charts && this._charts[uid];
        if (!entry) return;
        this._svgToCanvas(entry, 4).then(canvas => {
            const w = canvas.width, h = canvas.height;
            const png = canvas.toDataURL('image/png');
            const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w / 4}" height="${h / 4}" viewBox="0 0 ${w} ${h}">` +
                        `<image width="${w}" height="${h}" href="${png}"/></svg>`;
            this._downloadBlob(new Blob([svg], { type: 'image/svg+xml;charset=utf-8' }),
                `${this._chartFileBase(entry.cfg)}.svg`, 'Vector-wrapped SVG saved (chart type and colours preserved).');
        }).catch(() => showToast('Export failed', 'Could not render the chart image.', 'error'));
    },

    /* Export a very high-resolution PNG (universally works in any PowerPoint). */
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

    /* Render the current chart state to a standalone, well-formed vector SVG.
       • Title + axis labels + legend are real <text> (editable in PowerPoint
         after "Convert to Shape"), using the LIVE title from the toolbar input.
       • Lines are <polyline> vectors — infinitely scalable, no quality loss.
       • Honours the live Y-range and the current X zoom window. */
    _buildChartSvg(chart, cfg) {
        /* Read the live, user-edited title from the toolbar input (fallback cfg). */
        const uid = 'twa-ch-' + cfg.id;
        const titleEl = document.getElementById(`${uid}-title`);
        const title = (titleEl && titleEl.value) || cfg.title || '';

        const datasets = chart.data.datasets;
        const labels = cfg.labels;

        /* Legend layout: wrap chips across rows; reserve height for them. */
        const legendPerRow = 4;
        const legendRows = Math.ceil(datasets.length / legendPerRow);
        const legendH = legendRows * 20 + 8;

        const W = 1280;
        const H = 680 + legendH;
        const m = { l: 84, r: 28, t: 64, b: 104 };
        const pw = W - m.l - m.r;
        const ph = H - m.t - m.b - legendH;

        const xs = chart.scales.x, ys = chart.scales.y;
        const xMin = Math.max(0, Math.ceil(xs.min != null ? xs.min : 0));
        const xMax = Math.min(labels.length - 1, Math.floor(xs.max != null ? xs.max : labels.length - 1));
        const yMin = ys.min != null ? ys.min : 0;
        const yMax = ys.max != null ? ys.max : 100;
        const span = Math.max(1, xMax - xMin);
        const px = i => m.l + ((i - xMin) / span) * pw;
        const py = v => m.t + ph - ((v - yMin) / (yMax - yMin || 1)) * ph;
        const esc = s => String(s == null ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
        const f = n => (Math.round(n * 100) / 100);   /* trim coordinate precision */

        this._svgW = W; this._svgH = H;   /* used by the PNG rasterizer */

        const fsSvg = ChartFontSizeManager.get();

        const parts = [];
        parts.push(`<?xml version="1.0" encoding="UTF-8" standalone="no"?>`);
        parts.push(`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="Arial, Helvetica, sans-serif">`);
        parts.push(`<rect x="0" y="0" width="${W}" height="${H}" fill="#ffffff"/>`);
        /* Title (editable text) */
        parts.push(`<text id="chart-title" x="${W / 2}" y="36" text-anchor="middle" font-size="${fsSvg.title}" font-weight="700" fill="#1E293B">${esc(title)}</text>`);

        /* Y grid + labels */
        const ySteps = 5;
        for (let s = 0; s <= ySteps; s++) {
            const v = yMin + (s / ySteps) * (yMax - yMin);
            const y = py(v);
            parts.push(`<line x1="${m.l}" y1="${f(y)}" x2="${m.l + pw}" y2="${f(y)}" stroke="#E7EDF4" stroke-width="1"/>`);
            parts.push(`<text x="${m.l - 10}" y="${f(y + 4)}" text-anchor="end" font-size="${fsSvg.axisTicks}" fill="#64748B">${f(v)}</text>`);
        }

        /* Axis titles */
        parts.push(`<text x="${m.l + pw / 2}" y="${H - legendH - 30}" text-anchor="middle" font-size="${fsSvg.axisTitle}" fill="#334155">${esc(cfg.xLabel || '')}</text>`);
        parts.push(`<text x="22" y="${m.t + ph / 2}" text-anchor="middle" font-size="${fsSvg.axisTitle}" fill="#334155" transform="rotate(-90 22 ${m.t + ph / 2})">${esc(cfg.yLabel || '')}</text>`);

        /* X ticks */
        const nTicks = Math.min(12, span);
        for (let t = 0; t <= nTicks; t++) {
            const i = xMin + Math.round((t / nTicks) * span);
            const x = px(i);
            parts.push(`<line x1="${f(x)}" y1="${m.t + ph}" x2="${f(x)}" y2="${m.t + ph + 5}" stroke="#CBD5E1"/>`);
            parts.push(`<text x="${f(x)}" y="${m.t + ph + 18}" text-anchor="end" font-size="${fsSvg.axisTicks}" fill="#64748B" transform="rotate(-45 ${f(x)} ${m.t + ph + 18})">${esc(labels[i] || '')}</text>`);
        }

        /* Plot border */
        parts.push(`<rect x="${m.l}" y="${m.t}" width="${pw}" height="${ph}" fill="none" stroke="#CBD5E1"/>`);

        /* Series as vector polylines (break on null gaps). Hidden datasets
           (toggled off in the legend) are not plotted. */
        for (let di = 0; di < datasets.length; di++) {
            const ds = datasets[di];
            if (ds.hidden || (chart.getDatasetMeta(di) && chart.getDatasetMeta(di).hidden)) continue;
            const color = ds.borderColor || CHART_COLORS[di % CHART_COLORS.length];
            let pts = [];
            const flush = () => {
                if (pts.length === 1) {
                    const [cx, cy] = pts[0].split(',');
                    parts.push(`<circle cx="${cx}" cy="${cy}" r="1.3" fill="${color}"/>`);
                } else if (pts.length >= 2) {
                    parts.push(`<polyline points="${pts.join(' ')}" fill="none" stroke="${color}" stroke-width="1.4" stroke-linejoin="round" stroke-linecap="round"/>`);
                }
                pts = [];
            };
            for (let i = xMin; i <= xMax; i++) {
                const v = ds.data[i];
                if (v == null || isNaN(v)) { flush(); continue; }
                pts.push(`${f(px(i))},${f(py(v))}`);
            }
            flush();
        }

        /* Legend (wrapping rows, real text). Marker is a short line matching the
           series colour; hidden datasets are shown dimmed. Only datasets that are
           visible are drawn in the plot, but ALL appear in the legend so the
           reader sees the full server list. */
        const colW = pw / legendPerRow;
        const legendTop = H - legendH + 6;
        for (let di = 0; di < datasets.length; di++) {
            const ds = datasets[di];
            const hidden = ds.hidden || (chart.getDatasetMeta(di) && chart.getDatasetMeta(di).hidden);
            const color = ds.borderColor || CHART_COLORS[di % CHART_COLORS.length];
            const row = Math.floor(di / legendPerRow), colN = di % legendPerRow;
            const lx = m.l + colN * colW;
            const ly = legendTop + row * 20;
            const my = ly + 6;
            parts.push(`<line x1="${f(lx)}" y1="${f(my)}" x2="${f(lx + 22)}" y2="${f(my)}" stroke="${color}" stroke-width="2.4" opacity="${hidden ? 0.3 : 1}"/>`);
            parts.push(`<text x="${f(lx + 28)}" y="${f(ly + 11)}" font-size="${fsSvg.legend}" fill="${hidden ? '#94A3B8' : '#334155'}"${hidden ? ' text-decoration="line-through"' : ''}>${esc(ds.label)}</text>`);
        }

        parts.push(`</svg>`);
        return parts.join('');
    },

    /* Snapshot the current Extended Memory input values into this._extMemory
       (called just before the enriched table is built). */
    _saveExtMemory() {
        if (!this._extMemory) this._extMemory = {};
        document.querySelectorAll('.twa-extmem-input').forEach(input => {
            const key = (input.dataset.host || '').toLowerCase();
            if (!key) return;
            const n = parseFloat(input.value);
            this._extMemory[key] = isNaN(n) ? null : n;
        });
    },

    /* Replicated Server Inventory showing PHYSICAL HOSTNAME + VIRTUAL HOSTNAME.
       Rendered only when a virtual↔physical mapping was performed. */
    _renderPhysicalVirtualInventory(app) {
        const wrap = document.getElementById('twa-pv-inventory');
        if (!wrap) return;
        if (!this._mappingActive) { wrap.style.display = 'none'; return; }

        const removed = this._removedHosts || new Set();
        const hv = app.getHostVariables().filter(h => !removed.has((h.hostname || '').toLowerCase()));
        const dash = v => (v !== undefined && v !== null && String(v).trim() !== '' && v !== '-') ? v : '—';

        /* physical(lowercase) → [virtual hostnames mapped to it] */
        const byPhysical = {};
        Object.keys(this._mapping).forEach(vk => {
            const phys = (this._mapping[vk] || '').toLowerCase();
            const orig = this._virtualHosts.find(v => v.toLowerCase() === vk) || vk;
            (byPhysical[phys] || (byPhysical[phys] = [])).push(orig);
        });

        wrap.style.display = 'block';
        const tbody = document.getElementById('twa-pv-tbody');
        if (tbody) {
            tbody.innerHTML = hv.length
                ? hv.map(h => {
                    const virtuals = byPhysical[(h.hostname || '').toLowerCase()] || [];
                    const vLabel = virtuals.length ? virtuals.join(', ') : '— not mapped —';
                    return `
                    <tr>
                        <td style="font-weight:600">${dash(h.hostname)}</td>
                        <td>${vLabel}</td>
                        <td>${dash(h.manufacturer)}</td>
                        <td>${dash(h.model)}</td>
                        <td>${dash(h.cpuType)}</td>
                        <td>${dash(h.virtualization)}</td>
                        <td>${dash(h.os)}</td>
                        <td>${dash(h.ghz)}</td>
                        <td>${dash(h.logicalCpus)}</td>
                        <td>${h.cores || 0}</td>
                        <td><strong>${h.totalSaps ? Number(h.totalSaps).toLocaleString() : '—'}</strong></td>
                        <td>${h.memoryMB ? Number(h.memoryMB).toLocaleString() : '—'}</td>
                        <td>${dash(h.component)}</td>
                        <td>${h.layer === 'db' ? 'Database' : h.layer === 'app' ? 'Application' : 'Other'}</td>
                        <td>${h.sapsPerCore ? Number(h.sapsPerCore).toLocaleString() : '—'}</td>
                    </tr>`;
                }).join('')
                : `<tr><td colspan="15" style="color:var(--text-muted)">No landscape hosts.</td></tr>`;
        }

        /* Unmapped virtual hosts (no physical) listed as extra rows for completeness */
        const assigned = new Set(Object.keys(this._mapping));
        const unmapped = this._virtualHosts.filter(v => !assigned.has(v.toLowerCase()));

        const sub = document.getElementById('twa-pv-sub');
        if (sub) sub.textContent =
            `${hv.length} physical host(s) · ${Object.keys(this._mapping).length} mapped virtual host(s)` +
            (unmapped.length ? ` · ${unmapped.length} unmapped (no match in EWA report)` : '');

        /* Build the export rows once and share with copy + export */
        this._pvExportRows = () => {
            const rows = hv.map(h => ({
                'Physical Hostname'    : h.hostname || '',
                'Virtual Hostname'     : (byPhysical[(h.hostname || '').toLowerCase()] || []).join(', '),
                'Hardware Manufacturer': h.manufacturer   || '',
                'Model'                : h.model          || '',
                'CPU Type'             : h.cpuType        || '',
                'Virtualization'       : h.virtualization || '',
                'Operating System'     : h.os             || '',
                'GHz'                  : h.ghz            || '',
                'Logical CPUs'         : h.logicalCpus    || '',
                'Cores'                : h.cores          || 0,
                'SAPS (TBV)'           : h.totalSaps      || 0,
                'Memory in MB'         : h.memoryMB       || 0,
                'Component Type'       : h.component      || '',
                'Layer'                : h.layer === 'db' ? 'Database' : h.layer === 'app' ? 'Application' : 'Other',
                'SAPS/Core'            : h.sapsPerCore    || 0,
            }));
            /* Unmapped virtual hosts → rows with a "No match" physical marker */
            unmapped.forEach(v => rows.push({
                'Physical Hostname': 'No match in EWA report',
                'Virtual Hostname' : v,
                'Hardware Manufacturer': '', 'Model': '', 'CPU Type': '', 'Virtualization': '',
                'Operating System': '', 'GHz': '', 'Logical CPUs': '', 'Cores': '',
                'SAPS (TBV)': '', 'Memory in MB': '', 'Component Type': '', 'Layer': '', 'SAPS/Core': '',
            }));
            return rows;
        };

        const copyBtn = document.getElementById('twa-pv-copy');
        if (copyBtn) copyBtn.onclick = () => this._copyPvForExcel(copyBtn);
        const expBtn = document.getElementById('twa-pv-export');
        if (expBtn) expBtn.onclick = () => {
            const rows = this._pvExportRows();
            if (rows.length) downloadXlsx(rows, 'Physical-Virtual Hosts', `${this._sidTag()}_physical-virtual-hosts-${_dateTag()}.xlsx`);
        };
    },

    /* Copy the Physical↔Virtual inventory to the clipboard as TSV (Excel-pastable). */
    _copyPvForExcel(btn) {
        const rows = this._pvExportRows ? this._pvExportRows() : [];
        if (!rows.length) return;
        const headers = Object.keys(rows[0]);
        const tsv = [headers.join('\t')]
            .concat(rows.map(r => headers.map(h => String(r[h] ?? '')).join('\t')))
            .join('\r\n');
        const done = () => {
            if (!btn) return;
            const orig = btn.innerHTML;
            btn.innerHTML = '&#10003; Copied';
            setTimeout(() => { btn.innerHTML = orig; }, 1600);
        };
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(tsv).then(done).catch(() => this._copyFallback(tsv, done));
        } else {
            this._copyFallback(tsv, done);
        }
    },
    _copyFallback(tsv, done) {
        const ta = document.createElement('textarea');
        ta.value = tsv; ta.style.position = 'fixed'; ta.style.opacity = '0';
        document.body.appendChild(ta); ta.select();
        try { document.execCommand('copy'); done(); } catch (e) { alert('Copy failed — select the table manually.'); }
        document.body.removeChild(ta);
    },

    /* ══════════════════════════════════════════════════════
       TECHNICAL WORKLOAD ANALYSIS WORKBOOK
       Builds <SID>_TechnicalWorkloadAnalysis.xlsx with:
         • the enriched source-data worksheet(s), plus
         • one <name>_Data worksheet followed immediately by its paired
           <name>_Chart worksheet, for each primary chart.
       Chart worksheets hold NATIVE, editable Excel charts (OOXML) linked to
       the data range on the preceding Data sheet — no images. The OOXML
       injection reuses the proven pattern from appWorkload.js (via JSZip).
    ══════════════════════════════════════════════════════ */

    /* Cap native-chart category count so Excel stays responsive. Evenly samples
       a long series down to at most `max` indices (keeps first/last). Defensive
       against bad lengths so it can never throw "Invalid array length". */
    _downsample(labels, max) {
        const n = Array.isArray(labels) ? labels.length : 0;
        if (!Number.isFinite(n) || n <= 0) return [];
        if (!Number.isFinite(max) || max < 1) max = 1;
        if (n <= max) { const a = []; for (let i = 0; i < n; i++) a.push(i); return a; }
        const idx = [];
        const step = (n - 1) / (max - 1);
        for (let k = 0; k < max; k++) {
            const i = Math.round(k * step);
            if (i >= 0 && i < n) idx.push(i);
        }
        return Array.from(new Set(idx));
    },

    /* Validate a chart-data descriptor before it is turned into worksheets.
       Returns { ok:true } or { ok:false, reason } — never throws. Guards against
       the failure modes that cause "Invalid array length" and corrupt workbooks:
       missing/!array labels or series, mismatched X/Y lengths, negative/huge
       lengths, empty series. */
    _validateChartData(cd) {
        if (!cd || typeof cd !== 'object') return { ok: false, reason: 'descriptor is not an object' };
        if (!cd.key) return { ok: false, reason: 'missing key' };
        if (!Array.isArray(cd.labels)) return { ok: false, reason: 'labels is not an array' };
        const nLabels = cd.labels.length;
        if (!Number.isFinite(nLabels) || nLabels < 1) return { ok: false, reason: `no category labels (${nLabels})` };
        if (nLabels > 1048576) return { ok: false, reason: `too many categories (${nLabels})` };
        if (!Array.isArray(cd.series) || cd.series.length < 1) return { ok: false, reason: 'no series' };
        for (const s of cd.series) {
            if (!s || !Array.isArray(s.data)) return { ok: false, reason: `series "${s && s.name}" has no data array` };
            if (s.data.length !== nLabels) return { ok: false, reason: `series "${s.name}" length ${s.data.length} ≠ categories ${nLabels}` };
        }
        /* At least one real (non-empty) value somewhere. */
        const hasValue = cd.series.some(s => s.data.some(v => v !== '' && v != null && !isNaN(v)));
        if (!hasValue) return { ok: false, reason: 'all series are empty' };
        return { ok: true };
    },

    /* Collect the four primary chart datasets from the live source builders
       (always available from this.targetRows, independent of the active tab).
       Each collector is isolated in try/catch so a single bad chart cannot abort
       the others. Returns an array of validated
       { key, title, type, catLabel, leftAxisTitle, labels, series:[{name,data}] }. */
    _analysisChartData(app) {
        const MAX_PTS = 300;                 /* native-chart category cap */
        const r2 = v => (v == null || v === '' || isNaN(v)) ? '' : Math.round(Number(v) * 100) / 100;
        const out = [];

        /* Run one collector safely; validate its result; log + skip on any problem. */
        const collect = (key, fn) => {
            try {
                const cd = fn();
                if (!cd) { console.info(`[TWA workbook] chart "${key}" skipped — no data.`); return; }
                const v = this._validateChartData(cd);
                if (!v.ok) { console.warn(`[TWA workbook] chart "${key}" skipped — ${v.reason}.`); return; }
                console.info(`[TWA workbook] chart "${key}": ${cd.series.length} series × ${cd.labels.length} categories (type ${cd.type}).`);
                out.push(cd);
            } catch (err) {
                console.warn(`[TWA workbook] chart "${key}" threw and was skipped:`, err);
            }
        };

        /* ── CPU Utilization (line, one series per app host) ── */
        const cpu = this._appTierSeries(app) || { hosts: [], labels: [], cpu: {} };
        collect('CPU_Utilization', () => {
            if (!cpu.hosts.length || !cpu.labels.length) return null;
            const pick = this._downsample(cpu.labels, MAX_PTS);
            if (!pick.length) return null;
            return {
                key: 'CPU_Utilization',
                title: 'Average CPU usage for all the AS',
                type: 'line',
                catLabel: 'Date / Time',
                leftAxisTitle: 'CPU Utilization (%)',
                labels: pick.map(i => cpu.labels[i]),
                series: cpu.hosts.map(h => ({ name: h, data: pick.map(i => r2(cpu.cpu[h] ? cpu.cpu[h][i] : null)) })),
            };
        });

        /* ── Memory Utilization (line, EM Allocated GB per app host) ── */
        collect('Memory_Utilization', () => {
            const mem = this._appTierMemSeries(app) || { hosts: [], labels: [], emAlloc: {} };
            if (!mem.hosts.length || !mem.labels.length) return null;
            const pick = this._downsample(mem.labels, MAX_PTS);
            if (!pick.length) return null;
            return {
                key: 'Memory_Utilization',
                title: 'EM Allocated (GB) for all the AS',
                type: 'line',
                catLabel: 'Date / Time',
                leftAxisTitle: 'EM Allocated (GB)',
                labels: pick.map(i => mem.labels[i]),
                series: mem.hosts.map(h => ({ name: h, data: pick.map(i => r2(mem.emAlloc[h] ? mem.emAlloc[h][i] : null)) })),
            };
        });

        /* ── Work Processes (stacked column, ACT_WPS per app host) ── */
        collect('Work_Processes', () => {
            if (!cpu.hosts.length || !cpu.labels.length) return null;
            /* Resolve the real column key for ACT_WPS (tolerant of spacing/case). */
            const ciMap = this._colIndex || {};
            const want = 'ACTWPS';
            let wpCol = null;
            for (const k in ciMap) { if (k.toUpperCase().replace(/\s+/g, '') === want) { wpCol = k; break; } }
            if (!wpCol) return null;
            const wp = this._stackedMemByHost(app, wpCol, cpu.hosts, cpu.labels);
            const hs = (wp && wp.hostSeries) || [];
            if (!hs.length) return null;
            const pick = this._downsample(cpu.labels, MAX_PTS);
            if (!pick.length) return null;
            return {
                key: 'Work_Processes',
                title: 'Active Work Processes (stacked) for all the AS',
                type: 'stackedColumn',
                catLabel: 'Date / Time',
                leftAxisTitle: 'Active Work Processes',
                labels: pick.map(i => cpu.labels[i]),
                series: hs.map(s => ({ name: s.label, data: pick.map(i => r2(s.data ? s.data[i] : null)) })),
            };
        });

        /* ── Network (line, NETRTT per app host) — SMON only ── */
        collect('Network', () => {
            if (this.sourceTable !== 'SMON') return null;
            const rtt = this._rttByServer(app);
            if (!rtt || !rtt.hosts.length || !rtt.labels.length) return null;
            const pick = this._downsample(rtt.labels, MAX_PTS);
            if (!pick.length) return null;
            return {
                key: 'Network',
                title: 'RTT between the AS and the DB',
                type: 'line',
                catLabel: 'Date / Time',
                leftAxisTitle: 'RTT (µs)',
                labels: pick.map(i => rtt.labels[i]),
                series: rtt.hosts.map(h => ({ name: h, data: pick.map(i => r2(rtt.byHost[h] ? rtt.byHost[h][i] : null)) })),
            };
        });

        return out;
    },

    /* ── RTT conditional-formatting rules (shared by the on-screen summary table
       and the workbook export so the two always agree) ──
         < 350 µs            → Healthy          (light green)
         350–750 µs (incl.)  → Monitor          (light yellow)
         > 750 µs            → Attention Required(light red)
       A server whose average is "significantly above" the overall average
       (≥ 1.5× and at least 150 µs higher) is flagged as an outlier (red value). */
    _RTT_COLORS: { healthy: '#D6F5DD', monitor: '#FFF4CC', attention: '#FCD9D6' },
    /* Excel ARGB equivalents (no leading #, FF alpha) for workbook cell fills. */
    _RTT_FILLS_ARGB: { healthy: 'FFD6F5DD', monitor: 'FFFFF4CC', attention: 'FFFCD9D6' },

    _rttFormat(avg, overallAvg) {
        let status, bg, fill;
        if (avg < 350)        { status = 'Healthy';            bg = this._RTT_COLORS.healthy;   fill = this._RTT_FILLS_ARGB.healthy; }
        else if (avg <= 750)  { status = 'Monitor';            bg = this._RTT_COLORS.monitor;   fill = this._RTT_FILLS_ARGB.monitor; }
        else                  { status = 'Attention Required'; bg = this._RTT_COLORS.attention; fill = this._RTT_FILLS_ARGB.attention; }
        const outlier = overallAvg != null && overallAvg > 0 &&
            avg >= overallAvg * 1.5 && (avg - overallAvg) >= 150;
        return { status, bg, fill, outlier };
    },

    /* Build a time-aligned NETRTT series per application-tier host (SMON).
       Returns { labels, hosts, byHost:{host:[y|null]}, perServerAvg:{host:avg},
       overallAvg } — or null when NETRTT is unavailable. Shared by the Network
       chart data, the on-screen RTT summary, and the workbook export. */
    _rttByServer(app) {
        const ci = this._colIndex || {};
        const iDate = ci['Date'], iTime = ci['TIME'] != null ? ci['TIME'] : ci['Time'];
        const iSrv = ci['SERVER'], iEwa = ci['SERVER (EWA)'];
        /* Resolve NETRTT tolerantly (same logic as the on-screen findCol). */
        let rttKey = null;
        for (const k in ci) { if (k.toUpperCase().replace(/\s+/g, '') === 'NETRTT') { rttKey = k; break; } }
        const iRtt = rttKey != null ? ci[rttKey] : null;
        if (iDate == null || iSrv == null || iRtt == null) return null;

        const labelIndex = new Map(), labels = [], hosts = [], isApp = {}, byHost = {};
        const rows = this.targetRows || [];
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i]; const srv = row[iSrv];
            if (!srv) continue;
            if (isApp[srv] === undefined) isApp[srv] = this._rowLayer(app, srv, iEwa != null ? row[iEwa] : null) === 'app';
            if (!isApp[srv]) continue;
            const lbl = `${row[iDate] || ''} ${iTime != null ? (row[iTime] || '') : ''}`.trim();
            if (!labelIndex.has(lbl)) { labelIndex.set(lbl, labels.length); labels.push(lbl); }
            if (byHost[srv] === undefined) { byHost[srv] = null; hosts.push(srv); }
        }
        const n = labels.length;
        hosts.forEach(h => { byHost[h] = new Array(n).fill(null); });

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i]; const srv = row[iSrv];
            if (!srv || !isApp[srv]) continue;
            const lbl = `${row[iDate] || ''} ${iTime != null ? (row[iTime] || '') : ''}`.trim();
            const pos = labelIndex.get(lbl);
            if (pos == null) continue;
            const y = parseFloat(row[iRtt]); if (!isNaN(y)) byHost[srv][pos] = y;
        }

        /* Per-server average over its non-null samples; overall = mean of all samples. */
        const perServerAvg = {};
        let grandSum = 0, grandCount = 0;
        hosts.forEach(h => {
            let sum = 0, cnt = 0;
            byHost[h].forEach(v => { if (v != null && !isNaN(v)) { sum += v; cnt++; grandSum += v; grandCount++; } });
            perServerAvg[h] = cnt ? sum / cnt : null;
        });
        const overallAvg = grandCount ? grandSum / grandCount : null;

        return { labels, hosts, byHost, perServerAvg, overallAvg };
    },

    /* Build + download the analysis workbook. */
    async _buildAnalysisWorkbook(app) {
        if (!this.targetRows || !this.targetRows.length) {
            showToast('Nothing to export', 'Upload and process a SMON/MON file first.', 'info');
            return;
        }
        const btn = document.getElementById('twa-charts-wb-top');
        if (btn) btn.disabled = true;
        try {
            this._chartProgress('Collecting chart data…', 8);
            await this._nextFrame();

            const charts = this._analysisChartData(app);
            console.info(`[TWA workbook] ${charts.length} valid chart(s) collected for the workbook.`);

            /* ── Build the workbook: source data sheet first, then Data/Chart pairs. ── */
            this._chartProgress('Writing source-data worksheet…', 20);
            await this._nextFrame();
            const wb = XLSX.utils.book_new();

            /* Source data worksheet (the enriched target table), guarded for size
               and for a missing/!array targetCols. */
            const cols = Array.isArray(this.targetCols) ? this.targetCols : [];
            const src  = Array.isArray(this.targetRows) ? this.targetRows : [];
            const SOURCE_ROW_CAP = 100000;    /* keep the workbook openable */
            const nSrc = Math.max(0, Math.min(src.length, SOURCE_ROW_CAP));
            const nCol = Math.max(0, cols.length);
            try {
                const aoa = new Array(nSrc + 1);
                aoa[0] = cols.slice();
                for (let i = 0; i < nSrc; i++) {
                    const row = src[i] || [], line = new Array(nCol);
                    for (let c = 0; c < nCol; c++) { const v = row[c]; line[c] = (v == null) ? '' : v; }
                    aoa[i + 1] = line;
                }
                XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Source_Data');
                console.info(`[TWA workbook] worksheet "Source_Data": ${nSrc} rows × ${nCol} cols.`);
            } catch (err) {
                console.warn('[TWA workbook] Source_Data worksheet failed and was skipped:', err);
            }

            /* ── RTT_Summary worksheet (SMON): per-server RTT with the same
               conditional formatting rules as the on-screen table. The fills and
               red outlier font are applied later via styles.xml patching, so the
               visual indicators survive into Excel. ── */
            let rttFillPlan = null;
            try {
                const rtt = this._rttByServer(app);
                if (rtt && rtt.hosts.length) {
                    const header = ['Hostname', 'Average RTT (µs)', 'Status'];
                    const rowsOut = [];
                    const fills = [];   /* { r (0-based data row), status, outlier } */
                    rtt.hosts.forEach(h => {
                        const avg = rtt.perServerAvg[h];
                        if (avg == null) return;
                        const rounded = Math.round(avg);
                        const fmt = this._rttFormat(rounded, rtt.overallAvg);
                        rowsOut.push([h, rounded, fmt.status]);
                        fills.push({ status: fmt.status, outlier: fmt.outlier });
                    });
                    if (rowsOut.length) {
                        const headerRow = ['RTT between the AS and the DB — per-server summary'];
                        const avgRow = ['Average RTT for all servers (µs)', Math.round(rtt.overallAvg)];
                        const aoaRtt = [headerRow, avgRow, [], header, ...rowsOut];
                        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoaRtt), 'RTT_Summary');
                        /* Data rows start at Excel row 5 (1-based): row1 title, row2 avg,
                           row3 blank, row4 header. Record fills for styles patching. */
                        rttFillPlan = { sheet: 'RTT_Summary', firstDataRow: 5, fills };
                        console.info(`[TWA workbook] worksheet "RTT_Summary": ${rowsOut.length} server rows with conditional formatting.`);
                    }
                }
            } catch (err) {
                console.warn('[TWA workbook] RTT_Summary worksheet failed and was skipped:', err);
            }
            this._rttFillPlan = rttFillPlan;   /* consumed by _writeTwaXlsxWithCharts */

            /* ── Data + Chart worksheet pairs, chart immediately after its data. ── */
            const descriptors = [];
            const FIRST_DATA_ROW = 5;         /* row1 title, row2 type note, row3 blank, row4 header */
            const colLetter = n => {          /* 0-based → A, B, …, Z, AA, … */
                let s = ''; n += 1;
                while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
                return s;
            };

            for (let ci = 0; ci < charts.length; ci++) {
                const cd = charts[ci];
                this._chartProgress(`Writing ${cd.key} worksheets… (${ci + 1}/${charts.length})`, 25 + Math.round((ci / charts.length) * 35));
                await this._nextFrame();

                /* One bad chart must not abort the whole workbook. */
                try {
                    const dataSheet = `${cd.key}_Data`.slice(0, 31);
                    const chartSheet = `${cd.key}_Chart`.slice(0, 31);

                    /* Data sheet: header = [catLabel, ...series names], then one row per label. */
                    const header = [cd.catLabel, ...cd.series.map(s => s.name)];
                    const dataAoa = [
                        [cd.title],
                        [`Chart type: ${cd.type === 'line' ? 'Line' : cd.type === 'stackedColumn' ? 'Stacked Column' : cd.type}`],
                        [],
                        header,
                        ...cd.labels.map((lbl, i) => [lbl, ...cd.series.map(s => s.data[i])]),
                    ];
                    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(dataAoa), dataSheet);

                    /* Chart sheet: a note; the native chart is injected into it later. */
                    XLSX.utils.book_append_sheet(wb,
                        XLSX.utils.aoa_to_sheet([[`${cd.title} — native Excel chart (linked to ${dataSheet})`]]),
                        chartSheet);

                    /* Series columns are B, C, … (column A = category). */
                    const seriesCols = cd.series.map((s, i) => ({ col: colLetter(i + 1), name: s.name }));
                    descriptors.push({
                        chartSheet,
                        dataSheet,
                        firstDataRow: FIRST_DATA_ROW,
                        lastDataRow:  FIRST_DATA_ROW + cd.labels.length - 1,
                        type: cd.type,
                        title: cd.title,
                        seriesCols,
                        leftAxisTitle: cd.leftAxisTitle,
                        catAxisTitle: cd.catLabel,
                    });
                    console.info(`[TWA workbook] worksheets "${dataSheet}" + "${chartSheet}" written (${cd.series.length} series × ${cd.labels.length} rows).`);
                } catch (err) {
                    console.warn(`[TWA workbook] chart "${cd.key}" worksheets failed and were skipped:`, err);
                }
            }

            if (!descriptors.length) {
                this._hideChartProgress();
                if (btn) btn.disabled = false;
                showToast('No charts to export', 'No application-tier chart data was available for the workbook.', 'info');
                return;
            }

            /* ── Inject native charts and download. ── */
            this._chartProgress('Embedding native Excel charts…', 72);
            await this._nextFrame();
            const fileName = `${this._sidTag()}_TechnicalWorkloadAnalysis.xlsx`;
            await this._writeTwaXlsxWithCharts(wb, descriptors, fileName);

            this._chartProgress('Workbook ready', 100);
            setTimeout(() => this._hideChartProgress(), 900);
            showToast('Workbook downloaded',
                `${fileName} generated with ${descriptors.length} native chart${descriptors.length === 1 ? '' : 's'} linked to their data worksheets.`,
                'success');
        } catch (err) {
            this._hideChartProgress();
            console.error('[TechWorkloadApp] workbook build failed', err);
            showToast('Workbook failed', err && err.message ? err.message : String(err), 'error', 8000);
        } finally {
            if (btn) btn.disabled = false;
        }
    },

    /* Inject one native chart per descriptor into the SheetJS workbook via JSZip,
       then trigger the download. Mirrors appWorkload.js's _writeXlsxWithNativeCharts. */
    async _writeTwaXlsxWithCharts(wb, descriptors, fileName) {
        /* Download the plain SheetJS workbook (data only, no native charts). Used
           both when JSZip is unavailable and as the fallback if chart injection
           fails for any reason — so the user always gets a valid workbook. */
        const downloadPlain = () => {
            const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = fileName;
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(a.href), 10000);
        };

        if (typeof JSZip === 'undefined') { downloadPlain(); return; }

        try {
            /* Write directly to a byte array (type:'array') rather than a binary
               string. Copying a multi-MB binary string into a Uint8Array char by
               char is what previously threw "Invalid array length" on large
               workbooks; XLSX's array output avoids that entirely. */
            const bytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            const zip = await JSZip.loadAsync(bytes);
            const sheetIndexMap = await this._twaResolveSheetIndexMap(zip);

            let chartCounter = 0;
            for (const desc of descriptors) {
                try {
                    const sheetIdx = sheetIndexMap[desc.chartSheet];
                    if (!sheetIdx) { console.warn(`[TWA workbook] chart sheet "${desc.chartSheet}" not found in package — chart skipped.`); continue; }
                    chartCounter++;
                    const cId = chartCounter;

                    const chartPath    = `xl/charts/chart${cId}.xml`;
                    const drawingPath  = `xl/drawings/drawing${cId}.xml`;
                    const drawingRel   = `xl/drawings/_rels/drawing${cId}.xml.rels`;
                    const chartRel     = `xl/charts/_rels/chart${cId}.xml.rels`;
                    const sheetPath    = `xl/worksheets/sheet${sheetIdx}.xml`;
                    const sheetRelPath = `xl/worksheets/_rels/sheet${sheetIdx}.xml.rels`;

                    zip.file(chartPath,   this._twaMakeChartXml(desc));
                    zip.file(drawingPath, this._twaChartDrawingXml(cId));
                    zip.file(drawingRel,
                        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
                        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${cId}.xml"/>` +
                        `</Relationships>`);
                    zip.file(chartRel, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);

                    const sheetXml = await zip.file(sheetPath)?.async('string');
                    if (sheetXml) zip.file(sheetPath, this._twaPatchSheetXml(sheetXml, 'rId1'));
                    const existingRel = await zip.file(sheetRelPath)?.async('string');
                    zip.file(sheetRelPath, this._twaPatchSheetRels(existingRel, `../drawings/drawing${cId}.xml`));
                    console.info(`[TWA workbook] native chart ${cId} injected into "${desc.chartSheet}" (type ${desc.type}).`);
                } catch (err) {
                    chartCounter--;   /* undo the counter bump for this failed chart */
                    console.warn(`[TWA workbook] native chart for "${desc.chartSheet}" failed and was skipped:`, err);
                }
            }

            if (chartCounter === 0) {
                console.warn('[TWA workbook] no native charts injected — writing a workbook without charts.');
                /* Still apply RTT conditional formatting and emit via the zip so the
                   styles survive even when no chart was embedded. */
                await this._twaApplyRttStyles(zip, sheetIndexMap);
                const outBuf0 = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
                const blob0 = new Blob([outBuf0], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const a0 = document.createElement('a');
                a0.href = URL.createObjectURL(blob0);
                a0.download = fileName;
                document.body.appendChild(a0); a0.click(); document.body.removeChild(a0);
                setTimeout(() => URL.revokeObjectURL(a0.href), 10000);
                return;
            }

            const ctRaw = await zip.file('[Content_Types].xml').async('string');
            zip.file('[Content_Types].xml', this._twaPatchContentTypes(ctRaw, chartCounter));

            /* Apply RTT conditional formatting (fills + red outlier font) to the
               RTT_Summary worksheet via styles.xml patching. */
            await this._twaApplyRttStyles(zip, sheetIndexMap);

            const outBuf = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
            const blob = new Blob([outBuf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = fileName;
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(a.href), 10000);
        } catch (err) {
            console.warn('[TWA workbook] native-chart embedding failed — falling back to a data-only workbook:', err);
            downloadPlain();
        }
    },

    _twaXmlEsc(s) {
        return String(s == null ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    },

    /* Chart title/axis/legend font sizes (hundredths of a point) from the global
       manager, so the workbook matches the on-screen defaults (Title 25 / Axis 14 /
       Legend 17). Title is centered (Excel default; overlay off). */
    _twaChartTxPr(sizeHundredths, bold) {
        return `<c:txPr>` +
            `<a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>` +
            `<a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>` +
            `<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
            `<a:pPr><a:defRPr sz="${sizeHundredths}" b="${bold ? 1 : 0}"/></a:pPr><a:endParaRPr lang="en-US"/>` +
            `</a:p></c:txPr>`;
    },

    _twaMakeChartXml(desc) {
        const sz = ChartFontSizeManager.get();
        const q = n => `'${String(n).replace(/'/g, "''")}'`;
        const ds = desc.dataSheet, fdr = desc.firstDataRow, ldr = desc.lastDataRow;
        const ref = (col, r1, r2) => `${q(ds)}!$${col}$${r1}:$${col}$${r2}`;
        const catRef = ref('A', fdr, ldr);
        const esc = s => this._twaXmlEsc(s);

        /* One <c:ser> per data column. For line charts emit smooth=0; for stacked
           columns the series live inside a stacked barChart. */
        const makeSer = (sc, i) => {
            const nameRef = `${q(ds)}!$${sc.col}$${fdr - 1}`;
            const valRef  = ref(sc.col, fdr, ldr);
            return `<c:ser>` +
                `<c:idx val="${i}"/><c:order val="${i}"/>` +
                `<c:tx><c:strRef><c:f>${nameRef}</c:f></c:strRef></c:tx>` +
                (desc.type === 'line' ? `<c:smooth val="0"/>` : '') +
                `<c:cat><c:strRef><c:f>${catRef}</c:f></c:strRef></c:cat>` +
                `<c:val><c:numRef><c:f>${valRef}</c:f></c:numRef></c:val>` +
                `</c:ser>`;
        };
        const seriesXml = desc.seriesCols.map(makeSer).join('');

        const axisTitle = (txt) => txt
            ? `<c:title><c:tx><c:rich>` +
              `<a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>` +
              `<a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>` +
              `<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
              `<a:pPr><a:defRPr sz="${sz.axisTitle * 100}"/></a:pPr>` +
              `<a:r><a:rPr lang="en-US" sz="${sz.axisTitle * 100}"/><a:t>${esc(txt)}</a:t></a:r>` +
              `</a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : '';

        const catAx = `<c:catAx>` +
            `<c:axId val="1001"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
            `<c:delete val="0"/><c:axPos val="b"/>` +
            axisTitle(desc.catAxisTitle, 'b') +
            `<c:crossAx val="1002"/><c:tickLblPos val="low"/>` +
            this._twaChartTxPr(sz.axisTicks * 100, false) +
            `</c:catAx>`;
        const valAx = `<c:valAx>` +
            `<c:axId val="1002"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
            `<c:delete val="0"/><c:axPos val="l"/>` +
            axisTitle(desc.leftAxisTitle, 'l') +
            `<c:crossAx val="1001"/>` +
            this._twaChartTxPr(sz.axisTicks * 100, false) +
            `</c:valAx>`;

        let plotArea;
        if (desc.type === 'line') {
            plotArea = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>` +
                seriesXml +
                `<c:marker val="0"/><c:axId val="1001"/><c:axId val="1002"/></c:lineChart>` +
                catAx + valAx;
        } else { /* stackedColumn */
            plotArea = `<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:varyColors val="0"/>` +
                seriesXml +
                `<c:overlap val="100"/><c:axId val="1001"/><c:axId val="1002"/></c:barChart>` +
                catAx + valAx;
        }

        const titleXml = `<c:title><c:tx><c:rich>` +
            `<a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>` +
            `<a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>` +
            `<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
            `<a:pPr algn="ctr"><a:defRPr sz="${sz.title * 100}" b="1"/></a:pPr>` +
            `<a:r><a:rPr lang="en-US" sz="${sz.title * 100}" b="1"/><a:t>${esc(desc.title)}</a:t></a:r>` +
            `</a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <c:autoTitleDeleted val="0"/>
    ${titleXml}
    <c:plotArea><c:layout/>${plotArea}</c:plotArea>
    <c:legend><c:legendPos val="b"/><c:overlay val="0"/>${this._twaChartTxPr(sz.legend * 100, false)}</c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
</c:chartSpace>`;
    },

    _twaChartDrawingXml(cId) {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>14</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>34</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="${cId + 1}" name="Chart ${cId}"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;
    },

    _twaPatchContentTypes(xml, chartCount) {
        let patched = xml;
        for (let i = 1; i <= chartCount; i++) {
            const chartPart = `/xl/charts/chart${i}.xml`;
            const drawingPart = `/xl/drawings/drawing${i}.xml`;
            if (!patched.includes(chartPart)) {
                patched = patched.replace('</Types>',
                    `  <Override PartName="${chartPart}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>\n</Types>`);
            }
            if (!patched.includes(drawingPart)) {
                patched = patched.replace('</Types>',
                    `  <Override PartName="${drawingPart}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>\n</Types>`);
            }
        }
        return patched;
    },

    async _twaResolveSheetIndexMap(zip) {
        const wbXml   = await zip.file('xl/workbook.xml').async('string');
        const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
        const nameToRid = {};
        let m; const re1 = /<sheet\b[^>]*>/g;
        while ((m = re1.exec(wbXml)) !== null) {
            const tag = m[0];
            const nameM = /\bname="([^"]*)"/.exec(tag);
            const ridM  = /\br:id="([^"]*)"/.exec(tag);
            if (nameM && ridM) {
                const name = nameM[1].replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'");
                nameToRid[name] = ridM[1];
            }
        }
        const ridToTarget = {}; const re2 = /<Relationship\b[^>]*>/g;
        while ((m = re2.exec(relsXml)) !== null) {
            const tag = m[0];
            const idM  = /\bId="([^"]*)"/.exec(tag);
            const tgtM = /\bTarget="([^"]*)"/.exec(tag);
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

    _twaPatchSheetXml(xml, drawingRid) {
        let patched = xml;
        if (!patched.includes('xmlns:r=')) {
            patched = patched.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
        }
        patched = patched.replace(/<drawing\s[^>]*\/>/g, '').replace(/<drawing\/>/g, '');
        const drawingTag = `<drawing r:id="${drawingRid}"/>`;
        if (patched.includes('</worksheet>')) patched = patched.replace('</worksheet>', `${drawingTag}</worksheet>`);
        else patched += drawingTag;
        return patched;
    },

    _twaPatchSheetRels(existingXml, drawingTarget) {
        const newRel = `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="${drawingTarget}"/>`;
        if (!existingXml) {
            return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${newRel}</Relationships>`;
        }
        let patched = existingXml.replace(/<Relationship[^>]*Type="[^"]*\/drawing"[^>]*\/>/g, '');
        if (patched.includes('</Relationships>')) patched = patched.replace('</Relationships>', `${newRel}</Relationships>`);
        else patched += newRel;
        return patched;
    },

    /* Apply RTT conditional formatting to the RTT_Summary worksheet by patching
       xl/styles.xml (new fills + a red font + cellXfs) and stamping the matching
       style index onto each data row's cells. No-op if there is no fill plan.
       Any failure is swallowed (styles are cosmetic — the data still exports). */
    async _twaApplyRttStyles(zip, sheetIndexMap) {
        const plan = this._rttFillPlan;
        this._rttFillPlan = null;
        if (!plan || !plan.fills || !plan.fills.length) return;
        try {
            const sheetIdx = sheetIndexMap[plan.sheet];
            if (!sheetIdx) { console.warn('[TWA workbook] RTT_Summary not found for styling.'); return; }

            const stylesPath = 'xl/styles.xml';
            let styles = await zip.file(stylesPath)?.async('string');
            if (!styles) { console.warn('[TWA workbook] styles.xml missing — RTT formatting skipped.'); return; }

            const fills = this._RTT_FILLS_ARGB;
            const countAttr = (xml, tag) => {
                const m = new RegExp(`<${tag}\\b[^>]*\\bcount="(\\d+)"`).exec(xml);
                return m ? parseInt(m[1]) : (new RegExp(`<${tag}\\b`).test(xml) ? (xml.match(new RegExp(`<${tag.replace(/s$/, '')}\\b`, 'g')) || []).length : 0);
            };

            /* ── Fonts: existing count, then append a red bold font. ── */
            const fontCount = countAttr(styles, 'fonts');
            const redFontIdx = fontCount;             /* appended at the end */
            const redFontXml = `<font><b/><color rgb="FFC0392B"/><sz val="11"/><name val="Calibri"/></font>`;
            styles = styles.replace(/(<fonts\b[^>]*\bcount=")(\d+)(")/, (_m, a, c, b) => `${a}${parseInt(c) + 1}${b}`);
            styles = styles.replace('</fonts>', `${redFontXml}</fonts>`);

            /* ── Fills: existing count, then append 3 RTT fills. ── */
            const fillCount = countAttr(styles, 'fills');
            const healthyFill = fillCount, monitorFill = fillCount + 1, attentionFill = fillCount + 2;
            const fillXml = ['healthy', 'monitor', 'attention'].map(k =>
                `<fill><patternFill patternType="solid"><fgColor rgb="${fills[k]}"/><bgColor indexed="64"/></patternFill></fill>`).join('');
            styles = styles.replace(/(<fills\b[^>]*\bcount=")(\d+)(")/, (_m, a, c, b) => `${a}${parseInt(c) + 3}${b}`);
            styles = styles.replace('</fills>', `${fillXml}</fills>`);

            /* ── cellXfs: append xf records. For each status: a fill-only xf and a
               fill+red-font xf (for outlier average cells). ── */
            const xfCount = countAttr(styles, 'cellXfs');
            const xf = (fillId, fontId) =>
                `<xf numFmtId="0" fontId="${fontId || 0}" fillId="${fillId}" borderId="0" xfId="0" applyFill="1"${fontId ? ' applyFont="1"' : ''}/>`;
            const xfs = [
                xf(healthyFill, 0), xf(monitorFill, 0), xf(attentionFill, 0),       /* 0,1,2 — plain */
                xf(healthyFill, redFontIdx), xf(monitorFill, redFontIdx), xf(attentionFill, redFontIdx), /* 3,4,5 — red font */
            ];
            const baseXf = xfCount;
            styles = styles.replace(/(<cellXfs\b[^>]*\bcount=")(\d+)(")/, (_m, a, c, b) => `${a}${parseInt(c) + xfs.length}${b}`);
            styles = styles.replace('</cellXfs>', `${xfs.join('')}</cellXfs>`);
            zip.file(stylesPath, styles);

            /* style index for a (status, outlier) pair → the matching xf we appended. */
            const statusOrder = { 'Healthy': 0, 'Monitor': 1, 'Attention Required': 2 };
            const styleFor = (status, outlier) => baseXf + (outlier ? 3 : 0) + (statusOrder[status] ?? 0);

            /* ── Patch the sheet XML: set s="idx" on A/B/C cells of each data row. ── */
            const sheetPath = `xl/worksheets/sheet${sheetIdx}.xml`;
            let sheetXml = await zip.file(sheetPath)?.async('string');
            if (!sheetXml) { console.warn('[TWA workbook] RTT_Summary sheet XML missing.'); return; }

            plan.fills.forEach((f, i) => {
                const excelRow = plan.firstDataRow + i;      /* 1-based */
                const sIdx = styleFor(f.status, false);      /* row fill on all cells */
                const sIdxAvg = styleFor(f.status, f.outlier); /* avg cell: red font if outlier */
                /* Match the <row r="N"> … </row> block and stamp s= on A,B,C cells. */
                const rowRe = new RegExp(`(<row[^>]*\\br="${excelRow}"[^>]*>)([\\s\\S]*?)(</row>)`);
                sheetXml = sheetXml.replace(rowRe, (_full, open, body, close) => {
                    const cellRe = new RegExp(`<c\\b([^>]*?)\\br="([A-Z]+)${excelRow}"([^>]*?)(/>|>)`, 'g');
                    const patchedBody = body.replace(cellRe,
                        (_cell, pre, col, post, endch) => {
                            const useIdx = (col === 'B') ? sIdxAvg : sIdx;
                            /* remove any existing s="…" then add ours. */
                            const cleanedPre = pre.replace(/\s+s="\d+"/, '');
                            const cleanedPost = post.replace(/\s+s="\d+"/, '');
                            return `<c${cleanedPre} r="${col}${excelRow}"${cleanedPost} s="${useIdx}"${endch}`;
                        });
                    return `${open}${patchedBody}${close}`;
                });
            });
            zip.file(sheetPath, sheetXml);
            console.info(`[TWA workbook] RTT_Summary conditional formatting applied to ${plan.fills.length} rows.`);
        } catch (err) {
            console.warn('[TWA workbook] RTT conditional formatting failed (data still exported):', err);
        }
    },

    async _download() {
        if (!this.targetRows || !this.targetRows.length) {
            this._showError('There is no analysis to download yet. Please upload a SMON/MON file first.');
            return;
        }
        const btn = document.getElementById('twa-btn-download');
        if (btn) btn.disabled = true;
        try {
            this._showProgress('Preparing workbook…', 5);
            await this._nextFrame();

            const cols = this.targetCols;
            const src  = this.targetRows;
            const srcTag = this.sourceTable === 'MON' ? 'mon' : this.sourceTable === 'PREPARSED' ? 'processed' : 'smon';

            /* SheetJS stores every cell as a property on the worksheet object and
               throws "Too many properties to enumerate" once the cell count gets
               very large (hundreds of thousands of rows × dozens of columns).
               So: use XLSX for reasonably sized results, and stream a CSV (which
               has no such limit and opens directly in Excel) for large ones. */
            const cellCount = (src.length + 1) * cols.length;
            const XLSX_CELL_LIMIT = 600000;   /* safely under SheetJS's guard */

            if (cellCount > XLSX_CELL_LIMIT) {
                await this._downloadCsv(cols, src, `${this._sidTag()}_tech-workload-app-${srcTag}-${_dateTag()}.csv`);
                this._showProgress('Download ready', 100);
                setTimeout(() => this._hideProgress(), 900);
                showToast('Download started',
                    `The result is large (${src.length.toLocaleString()} rows), so it was saved as CSV ` +
                    `(opens directly in Excel) to avoid the spreadsheet cell limit.`,
                    'success');
                return;
            }

            /* Build the sheet as an Array-Of-Arrays (header + one array per row)
               and use aoa_to_sheet (no per-object property enumeration).
               targetRows are already arrays aligned to `cols`, so we only need
               to normalise empties — no per-cell object key lookups. */
            const aoa = new Array(src.length + 1);
            aoa[0] = cols.slice();
            const CHUNK = 2000;
            const nc = cols.length;
            for (let start = 0; start < src.length; start += CHUNK) {
                const end = Math.min(start + CHUNK, src.length);
                for (let i = start; i < end; i++) {
                    const r = src[i];
                    const line = new Array(nc);
                    for (let c = 0; c < nc; c++) {
                        const v = r[c];
                        line[c] = (v === undefined || v === null) ? '' : v;
                    }
                    aoa[i + 1] = line;
                }
                this._showProgress('Preparing rows…', 5 + Math.round((end / src.length) * 70));
                await this._nextFrame();
            }

            this._showProgress('Generating XLSX file…', 85);
            await this._nextFrame();

            const ws = XLSX.utils.aoa_to_sheet(aoa);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Technical Workload App');

            this._showProgress('Starting download…', 97);
            await this._nextFrame();

            XLSX.writeFile(wb, `${this._sidTag()}_tech-workload-app-${srcTag}-${_dateTag()}.xlsx`);

            this._showProgress('Download ready', 100);
            setTimeout(() => this._hideProgress(), 900);
            showToast('Download started',
                `The target workbook (${src.length.toLocaleString()} rows) has been generated and saved to your downloads.`,
                'success');
        } catch (err) {
            this._hideProgress();
            this._showError(`Could not generate the workbook: ${err.message}`);
            showToast('Download failed', err.message, 'error');
        } finally {
            if (btn) btn.disabled = false;
        }
    },

    /* Stream a large result set to a CSV file (chunked, so the progress bar
       paints and memory stays bounded). Values are RFC-4180 quoted when they
       contain commas, quotes or newlines. */
    async _downloadCsv(cols, src, fileName) {
        const esc = v => {
            if (v === undefined || v === null) return '';
            const s = String(v);
            return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
        };
        const parts = [];
        parts.push(cols.map(esc).join(',') + '\r\n');

        const CHUNK = 5000;
        const nc = cols.length;
        for (let start = 0; start < src.length; start += CHUNK) {
            const end = Math.min(start + CHUNK, src.length);
            let block = '';
            for (let i = start; i < end; i++) {
                const r = src[i];
                const line = new Array(nc);
                for (let c = 0; c < nc; c++) line[c] = esc(r[c]);
                block += line.join(',') + '\r\n';
            }
            parts.push(block);
            this._showProgress('Writing CSV…', 5 + Math.round((end / src.length) * 90));
            await this._nextFrame();
        }

        /* UTF-8 BOM so Excel detects encoding correctly. */
        const blob = new Blob(['﻿', ...parts], { type: 'text/csv;charset=utf-8;' });
        const url  = URL.createObjectURL(blob);
        const a = Object.assign(document.createElement('a'), { href: url, download: fileName });
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    },

    /* ── Progress helpers ── */
    _showProgress(label, pct) {
        const box  = document.getElementById('twa-progress');
        const lab  = document.getElementById('twa-progress-label');
        const pctE = document.getElementById('twa-progress-pct');
        const fill = document.getElementById('twa-progress-fill');
        if (box)  box.style.display = 'block';
        if (lab)  lab.textContent = label;
        const p = Math.max(0, Math.min(100, Math.round(pct)));
        if (pctE) pctE.textContent = `${p}%`;
        if (fill) fill.style.width = `${p}%`;
    },
    _hideProgress() {
        const box = document.getElementById('twa-progress');
        if (box) box.style.display = 'none';
    },
    /* Yield to the browser so the progress bar repaints between steps. */
    _nextFrame() {
        return new Promise(resolve => requestAnimationFrame(() => setTimeout(resolve, 0)));
    },

    /* ── Error helpers ── */
    _showError(msg) {
        const wrap = document.getElementById('twa-error');
        const txt  = document.getElementById('twa-error-msg');
        if (wrap) wrap.style.display = 'flex';
        if (txt)  txt.innerHTML = msg;
        const res = document.getElementById('twa-results');
        if (res) res.style.display = 'none';
    },
    _hideError() {
        const el = document.getElementById('twa-error');
        if (el) el.style.display = 'none';
    },
};
