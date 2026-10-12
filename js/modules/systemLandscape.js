/* ─────────────────────────────────────────────────────────
   Module 01 – System Landscape  (SAP EWA Report Parser)
   ─────────────────────────────────────────────────────────
   Accepted formats : PDF · DOC/DOCX · HTML / HTM · XML
   Key sections parsed:
     §2.2  Servers in current Landscape  → component-type map
     §2.3  Hardware Configuration         → server inventory table
     Cover page / overview               → SID, Product, DB info
   ─────────────────────────────────────────────────────────*/

const SystemLandscapeModule = {

    /* ── persistent state (used by later modules) ── */
    servers:       [],
    systemInfo:    { sid: '', product: '', dbSystem: '', dbVersion: '', dbOs: '' },
    hostAliases:   {},   /* { physicalLower -> virtualName } from §2.2 Instance Name / SAPDBHOST */
    hasIBMServers: false,
    _slCharts:     {},

    /* Clear working state (used by New Analysis and before a session load). */
    resetState() {
        this.servers       = [];
        this.systemInfo    = { sid: '', product: '', dbSystem: '', dbVersion: '', dbOs: '' };
        this.hostAliases   = {};
        this.hasIBMServers = false;
        this._spcOverride  = {};
        this._slCharts     = {};
    },

    /* ── Session save/restore contract ── (all plain serializable data) */
    getSessionState() {
        return {
            servers:       this.servers,
            systemInfo:    this.systemInfo,
            hostAliases:   this.hostAliases,
            hasIBMServers: this.hasIBMServers,
            spcOverride:   this._spcOverride  || {},
        };
    },
    setSessionState(st) {
        if (!st) return;
        this.servers       = st.servers    || [];
        this.systemInfo    = st.systemInfo || { sid: '', product: '', dbSystem: '', dbVersion: '', dbOs: '' };
        this.hostAliases   = st.hostAliases || {};
        this.hasIBMServers = !!st.hasIBMServers;
        this._spcOverride  = st.spcOverride  || {};
        /* Legacy sessions saved sapsOverride as a separate map — migrate it into
           the per-server _sapsManual flag so old sessions still work correctly. */
        if (st.sapsOverride && Object.keys(st.sapsOverride).length > 0) {
            const legacyByHostname = st.sapsOverride;
            this.servers.forEach(s => {
                if (legacyByHostname[s.hostname]) s._sapsManual = true;
            });
        }
    },

    /* ── SAPS-per-core lookup (SAP benchmark reference values) ─
       Key = CPU model substring, value = SAPS per physical core.
       Users can override per-CPU via the SAPS/core input field.  */
    SAPS_DB: {
        'Xeon Platinum 9480':   4200,
        'Xeon Platinum 8490H':  4000,
        'Xeon Platinum 8480+':  3581,
        'Xeon Platinum 8480':   3500,
        'Xeon Platinum 8462Y+': 3400,
        'Xeon Platinum 8380HL': 2990,
        'Xeon Platinum 8380':   2950,
        'Xeon Platinum 8375C':  3100,
        'Xeon Platinum 8370C':  3050,
        'Xeon Platinum 8362':   2800,
        'Xeon Platinum 8360Y':  2900,
        'Xeon Platinum 8358P':  2850,
        'Xeon Platinum 8358':   2800,
        'Xeon Platinum 8352Y':  2700,
        'Xeon Platinum 8336C':  2700,
        'Xeon Platinum 8280':   2500,
        'Xeon Platinum 8260':   2300,
        'Xeon Gold 6438Y':      3600,
        'Xeon Gold 6438N':      3500,
        'Xeon Gold 6348':       3050,
        'Xeon Gold 6338N':      3100,
        'Xeon Gold 6338':       3000,
        'Xeon Gold 6330':       2900,
        'Xeon Gold 6330N':      2800,
        'Xeon Gold 6230R':      2400,
        'Xeon Gold 6230':       2200,
        'EPYC 9754':            4500,
        'EPYC 9654':            4100,
        'EPYC 9554':            3900,
        'EPYC 9474F':           4000,
        'EPYC 7763':            3200,
        'EPYC 7742':            2900,
        'EPYC 7713':            2800,
        'EPYC 7643':            2700,
        'EPYC 7543':            2600,
        /* Cloud / custom variants */
        'EPYC 7R13':            2900,   /* AWS EPYC Milan-based */
        'EPYC 7R32':            2600,
        'EPYC 7571':            2300,
        'Xeon Platinum 8124M':  2500,
        'Xeon Platinum 8151':   2200,
        'Xeon Platinum 8175M':  2400,   /* AWS Skylake-SP */
        'Xeon Platinum 8259CL': 2600,   /* AWS */
        'Xeon Platinum 8375C':  3100,   /* AWS Ice Lake */
        'Xeon Platinum 8488C':  3600,   /* AWS Sapphire Rapids */
    },

    /* Words that can never be a server hostname */
    NON_HOSTNAME: new Set([
        'earlywatch','alert','report','overview','configuration','section',
        'landscape','hardware','software','network','transport','services',
        'host','hostname','related','system','systems','server','servers',
        'component','components','product','version','database','instance',
        'manufacturer','model','type','memory','logical','physical','cores',
        'processes','operating','description','table','page','current',
        'information','parameters','general','installation','status',
        'priority','critical','warning','active','passive','virtual','machine',
        'class','name','service','enterprise','proliants','proliant','superdome',
        'xeon','intel','amd','suse','linux','redhat','windows','oracle',
        'fujitsu','cisco','lenovo','hewlett','packard','platinum','gold','silver',
        'systems','thinkserver','thinkcentre','poweredge','primergy',
        'alert','attach','none','yes','true','false','note','see',
        'amazon','google','microsoft','azure','vmware','centos',
        'development','develo','production','staging','testing',
    ]),

    /* ══════════════════════════════════════════════════════
       RENDER
       ══════════════════════════════════════════════════════ */
    render() {
        return `
        <div class="module-wrap">

            <div class="module-header" id="sl-module-header">
                <div class="module-crumb">Analysis › Step 01</div>
                <h1 class="module-title">System Landscape</h1>
                <p class="module-desc">
                    Upload your SAP EarlyWatch Alert (EWA) report to automatically extract the
                    hardware configuration and system landscape details.
                    Supported formats: <strong>PDF · DOC · DOCX · HTML · XML</strong>.
                </p>
            </div>

            <!-- Upload card -->
            <div class="card" id="sl-upload-card">
                <div class="card-header">
                    <div>
                        <div class="card-title">EWA Report Upload</div>
                        <div class="card-sub">SAP EarlyWatch Alert report — PDF · DOC · DOCX · HTML · XML</div>
                    </div>
                </div>
                <div class="upload-zone" id="sl-zone">
                    <input type="file" class="file-input-hidden" id="sl-input"
                           accept=".pdf,.doc,.docx,.html,.htm,.xml">
                    <div class="upload-zone-icon">
                        <svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="1.6">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="17 8 12 3 7 8"/>
                            <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                    </div>
                    <div class="upload-zone-text">
                        <div class="upload-zone-title">Upload EWA report</div>
                        <div class="upload-zone-sub">Drag &amp; drop or browse — PDF · DOC · DOCX · HTML · XML</div>
                    </div>
                    <span class="upload-zone-cta">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="17 8 12 3 7 8"/>
                            <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        Choose file
                    </span>
                </div>
                <div id="sl-status"></div>
                <div id="sl-progress" style="display:none" class="parse-progress">
                    <div class="parse-progress-bar"></div>
                    <span id="sl-progress-msg">Analyzing file…</span>
                </div>
            </div>

            <!-- Error panel -->
            <div id="sl-error" style="display:none" class="alert-error">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24"
                     fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                    <circle cx="12" cy="12" r="10"/>
                    <line x1="15" y1="9" x2="9" y2="15"/>
                    <line x1="9" y1="9" x2="15" y2="15"/>
                </svg>
                <div id="sl-error-msg"></div>
            </div>

            <!-- ── Results (hidden until file parsed) ── -->
            <div id="sl-results" style="display:none">

              <!-- ══ VIEW 1: Hardware configuration (main) ══ -->
              <div id="sl-view-main">

                <!-- Summary cards -->
                <div class="section-title" style="margin-top:24px">System Summary</div>
                <div class="stats-grid" id="sl-stats"></div>

                <!-- Step 1: CPU / SAPS summary — shown first so user verifies SAPS/Core before seeing the table -->
                <div class="section-title" style="margin-top:24px">
                    Step 1 — Processor Summary &amp; SAPS Estimation
                </div>
                <div class="alert-info" style="margin-bottom:14px">
                    <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                        <circle cx="12" cy="12" r="10"/>
                        <line x1="12" y1="8" x2="12" y2="12"/>
                        <line x1="12" y1="16" x2="12.01" y2="16"/>
                    </svg>
                    <div style="font-size:12px">
                        Review the detected processor(s) below and verify the <strong>SAPS/Core</strong> value
                        for each CPU type. Values are sourced from the SAP SD Benchmark where available, or
                        estimated from the knowledge base. Adjust any value if needed — the Server Inventory
                        table will update instantly.
                    </div>
                </div>
                <div id="sl-cpu-summary"></div>

                <!-- Step 2: Hardware table -->
                <div class="section-title" style="margin-top:28px">
                    Step 2 — Hardware Configuration
                    <span style="font-size:11px;font-weight:400;color:var(--text-muted);margin-left:6px">
                        — click any cell to edit
                    </span>
                </div>
                <div class="card" style="padding:0">
                    <div class="card-header"
                         style="padding:14px 18px 12px;border-radius:14px 14px 0 0">
                        <div>
                            <div class="card-title">Server Inventory</div>
                            <div class="card-sub" id="sl-table-sub"></div>
                        </div>
                        <div style="display:flex;gap:8px;flex-wrap:wrap">
                            <button class="btn" id="sl-btn-add">+ Add Row</button>
                            <button class="btn" id="sl-btn-copy-excel">
                                <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13"
                                     viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                    <rect x="9" y="9" width="13" height="13" rx="2" ry="2"/>
                                    <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>
                                </svg>
                                Copy for Excel
                            </button>
                            <button class="btn" id="sl-btn-export">
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
                    <div style="overflow-x:auto;border-radius:0 0 14px 14px">
                        <table class="data-table hw-table" id="sl-hw-table">
                            <thead>
                                <tr>
                                    <th></th>
                                    <th class="sl-th-resize" title="Drag the right edge to widen this column">Hostname</th>
                                    <th>Manufacturer</th>
                                    <th>Model</th>
                                    <th>CPU Type</th>
                                    <th>Virtualization</th>
                                    <th>Operating System</th>
                                    <th>GHz</th>
                                    <th>Logical CPUs</th>
                                    <th>Cores</th>
                                    <th>SAPS (TBV)</th>
                                    <th>Memory (MB)</th>
                                    <th>Component Type</th>
                                    <th></th>
                                </tr>
                            </thead>
                            <tbody id="sl-hw-tbody"></tbody>
                        </table>
                    </div>
                </div>

                <!-- Table completion legend -->
                <div class="alert-info" id="sl-table-legend" style="margin-top:18px">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:2px">
                        <circle cx="12" cy="12" r="10"/>
                        <line x1="12" y1="8" x2="12" y2="12"/>
                        <line x1="12" y1="16" x2="12.01" y2="16"/>
                    </svg>
                    <div>
                        <strong>Action required — complete the Hardware Configuration table before proceeding.</strong><br>
                        For each server, ensure the following fields are filled:
                        <strong>Hostname · Cores · SAPS (TBV) · Memory (MB) · Component Type</strong>.
                        Use the SAPS/Core values above for each CPU type — they auto-populate the SAPS column.
                        You may override any SAPS/Core value. Once all rows are complete, click <strong>Next</strong>.
                    </div>
                </div>

                <!-- Validation error panel -->
                <div id="sl-validation-error" style="display:none" class="alert-error">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                        <circle cx="12" cy="12" r="10"/>
                        <line x1="15" y1="9" x2="9" y2="15"/>
                        <line x1="9" y1="9" x2="15" y2="15"/>
                    </svg>
                    <div id="sl-validation-msg"></div>
                </div>

                <div class="action-row" style="margin-top:20px">
                    <button class="btn btn-blue" id="sl-btn-next">
                        Verify HW Configuration
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="9 18 15 12 9 6"/>
                        </svg>
                    </button>
                </div>

              </div><!-- /#sl-view-main -->

              <!-- ══ VIEW 2: HW review & analysis ══ -->
              <div id="sl-view-review" style="display:none">

                <div class="module-header" style="margin-top:4px">
                    <div class="module-crumb">Analysis › Step 01 › Verify</div>
                    <h1 class="module-title">Verify Hardware Configuration</h1>
                    <p class="module-desc">
                        Review the final Server Inventory, system information and deployed
                        capacity below before continuing to the technical workload analysis.
                    </p>
                </div>

                <!-- Top navigation (same actions as the bottom bar, within reach) -->
                <div class="action-row" style="margin-top:4px;margin-bottom:16px;display:flex;gap:10px">
                    <button class="btn" id="sl-btn-back-top">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="15 18 9 12 15 6"/>
                        </svg>
                        Back to Hardware Table
                    </button>
                    <button class="btn btn-blue" id="sl-btn-review-next-top">
                        Next: Technical Workload
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="9 18 15 12 9 6"/>
                        </svg>
                    </button>
                </div>

                <!-- Server Inventory (final values, read-only recap) -->
                <div class="section-title" style="margin-top:4px">Server Inventory</div>
                <div class="card" style="padding:0">
                    <div class="card-header" style="padding:14px 18px 12px;border-radius:14px 14px 0 0">
                        <div>
                            <div class="card-title">Server Inventory</div>
                            <div class="card-sub" id="sl-rv-inv-sub">Final values from the hardware configuration</div>
                        </div>
                        <div style="display:flex;gap:8px;flex-wrap:wrap">
                            <button class="btn" id="sl-rv-btn-copy">
                                <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13"
                                     viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                    <rect x="9" y="9" width="13" height="13" rx="2" ry="2"/>
                                    <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>
                                </svg>
                                Copy for Excel
                            </button>
                            <button class="btn" id="sl-rv-btn-export">
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
                        <table class="data-table" id="sl-rv-inventory-table">
                            <thead>
                                <tr>
                                    <th>Hostname</th>
                                    <th>Manufacturer</th>
                                    <th>Model</th>
                                    <th>CPU Type</th>
                                    <th>Virtualization</th>
                                    <th>Operating System</th>
                                    <th>GHz</th>
                                    <th>Logical CPUs</th>
                                    <th>Cores</th>
                                    <th>SAPS (TBV)</th>
                                    <th>Memory (MB)</th>
                                    <th>Component Type</th>
                                </tr>
                            </thead>
                            <tbody id="sl-rv-inventory-tbody"></tbody>
                        </table>
                    </div>
                </div>

                <div class="section-title" style="margin-top:24px">
                    System Information &amp; Deployed Capacity
                    <span style="font-size:11px;font-weight:400;color:var(--text-muted);margin-left:6px">
                        — review the hardware landscape before continuing
                    </span>
                </div>

                <!-- Application tier -->
                <div class="card" style="margin-bottom:12px">
                    <div class="card-header" style="padding:14px 18px 10px">
                        <div>
                            <div class="card-title">Application Tier</div>
                            <div class="card-sub">SAP system identification — extracted from EWA, editable</div>
                        </div>
                    </div>
                    <div id="sl-rv-sysinfo-app" style="padding:4px 18px 18px"></div>
                </div>

                <!-- Database tier -->
                <div class="card">
                    <div class="card-header" style="padding:14px 18px 10px">
                        <div>
                            <div class="card-title">Database Tier</div>
                            <div class="card-sub">Database identification &amp; configuration — extracted from EWA, editable</div>
                        </div>
                    </div>
                    <div id="sl-rv-sysinfo-db" style="padding:4px 18px 18px"></div>
                </div>

                <!-- SAPS per core -->
                <div class="section-title" style="margin-top:24px">SAPS per core (TBV by HW partner)</div>
                <div class="card" style="padding:0">
                    <div style="overflow-x:auto;border-radius:14px">
                        <table class="data-table" id="sl-rv-saps-table">
                            <thead>
                                <tr><th>CPU Type</th><th>GHz</th><th>SAPS / Core</th><th>Source</th></tr>
                            </thead>
                            <tbody id="sl-rv-saps-tbody"></tbody>
                        </table>
                    </div>
                </div>

                <!-- Deployed capacity -->
                <div class="section-title" style="margin-top:24px">Deployed Capacity Summary</div>
                <div class="card" style="padding:0">
                    <div style="overflow-x:auto;border-radius:14px">
                        <table class="data-table" id="sl-rv-capacity-table">
                            <thead>
                                <tr><th>Scope</th><th>Deployed Capacity (SAPS)</th><th>Deployed Memory (GB)</th><th>Total Cores</th></tr>
                            </thead>
                            <tbody id="sl-rv-capacity-tbody"></tbody>
                        </table>
                    </div>
                </div>

                <div class="action-row" style="margin-top:20px;display:flex;gap:10px">
                    <button class="btn" id="sl-btn-back">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="15 18 9 12 15 6"/>
                        </svg>
                        Back to Hardware Table
                    </button>
                    <button class="btn btn-blue" id="sl-btn-review-next">
                        Next: Technical Workload
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="9 18 15 12 9 6"/>
                        </svg>
                    </button>
                </div>

              </div><!-- /#sl-view-review -->

            </div><!-- /#sl-results -->
        </div>`;
    },

    /* ══════════════════════════════════════════════════════
       INIT
       ══════════════════════════════════════════════════════ */
    init(app) {
        this.app = app;

        const zone  = document.getElementById('sl-zone');
        const input = document.getElementById('sl-input');

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
            if (f) this.handleFile(f);
        });
        input.addEventListener('change', e => {
            const f = e.target.files[0];
            if (f) this.handleFile(f);
        });

        /* Rehydrate: if a report was already parsed earlier this session,
           restore the results view instead of showing an empty upload screen.
           The parsed servers survive navigation because this module object is a
           singleton held in SizingApp's closure. */
        if (this.servers && this.servers.length) {
            this._renderResults(this.servers, this.systemInfo || {});
        }
    },

    /* ══════════════════════════════════════════════════════
       FILE ROUTING
       ══════════════════════════════════════════════════════ */
    handleFile(file) {
        const ext = file.name.split('.').pop().toLowerCase();
        showFileLoaded('sl-status', file.name);
        this._showProgress('Reading file…');
        this._hideError();
        document.getElementById('sl-results').style.display = 'none';

        if (ext === 'pdf') {
            this._parsePDF(file);
        } else if (['html', 'htm'].includes(ext)) {
            const r = new FileReader();
            r.onload = e => {
                try {
                    const dom = new DOMParser().parseFromString(e.target.result, 'text/html');
                    this._processText(dom.body ? (dom.body.innerText || dom.body.textContent) : e.target.result);
                } catch (err) {
                    this._showError(`Failed to parse HTML: ${err.message}`);
                }
            };
            r.readAsText(file);
        } else if (ext === 'xml') {
            const r = new FileReader();
            r.onload = e => {
                const raw = e.target.result;
                /* Word XML documents contain <w:t> elements — detect and extract */
                if (raw.includes('xmlns:w=') || raw.includes('<w:document') || raw.includes('<w:t>') || raw.includes('<w:t ')) {
                    this._processText(this._extractWordXMLText(raw));
                } else {
                    /* Generic XML: try DOMParser plaintext extraction */
                    try {
                        const dom = new DOMParser().parseFromString(raw, 'text/xml');
                        this._processText(dom.documentElement ? (dom.documentElement.textContent || raw) : raw);
                    } catch (_) {
                        this._processText(raw);
                    }
                }
            };
            r.readAsText(file);
        } else if (['doc', 'docx'].includes(ext)) {
            this._parseDOCX(file);
        } else {
            this._hideProgress();
            this._showError(
                `Unsupported file format ".${ext}". ` +
                'Please upload a PDF, DOC, DOCX, HTML, or XML EWA report.'
            );
        }
    },

    async _parsePDF(file) {
        try {
            if (typeof pdfjsLib === 'undefined') {
                throw new Error(
                    'PDF.js library failed to load. ' +
                    'Check your internet connection and reload the page.'
                );
            }
            this._updateProgress('Loading PDF…');
            const ab  = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(ab) }).promise;
            const n   = pdf.numPages;
            let fullText = '';

            for (let pageNum = 1; pageNum <= n; pageNum++) {
                this._updateProgress(`Reading page ${pageNum} of ${n}…`);
                const page    = await pdf.getPage(pageNum);
                const content = await page.getTextContent();
                const vp      = page.getViewport({ scale: 1 });
                const H       = vp.height;

                /* Collect items with coordinates.
                   Flip Y so y=0 is the TOP of the page (PDF origin is bottom-left). */
                const items = content.items
                    .filter(it => it.str && it.str.trim())
                    .map(it => ({
                        x: it.transform[4],
                        y: H - it.transform[5],
                        t: it.str.trim()
                    }));

                /* Group into Y-rows (±4 px tolerance) */
                const yRows = [];
                for (const item of items) {
                    let placed = false;
                    for (const row of yRows) {
                        if (Math.abs(row.y - item.y) <= 4) {
                            row.items.push(item);
                            placed = true;
                            break;
                        }
                    }
                    if (!placed) yRows.push({ y: item.y, items: [item] });
                }
                yRows.sort((a, b) => a.y - b.y);   // top → bottom

                fullText += this._yRowsToText(yRows) + '\n';
            }

            this._processText(fullText);
        } catch (err) {
            this._hideProgress();
            this._showError(
                `Failed to parse PDF: ${err.message}. ` +
                'Please verify this is a valid SAP EarlyWatch Alert report.'
            );
        }
    },

    /* Converts Y-grouped rows from one PDF page into a text string.
       For pages that contain the Host Overview table, uses X-column
       reconstruction so multi-line cells (e.g. "ProLiant / DL380 / Gen11"
       across three visual rows) are reassembled into the correct column
       instead of being interleaved with other columns.               */
    _yRowsToText(yRows) {
        /* Simple left-to-right join per Y-row.
           Field extraction in _parseHardwareLine handles the fragmented
           multi-line cells that result from this approach. */
        return yRows.map(row => {
            row.items.sort((a, b) => a.x - b.x);
            return row.items.map(it => it.t).join(' ').replace(/\s{2,}/g, ' ').trim();
        }).filter(l => l.length > 0).join('\n');
    },

    async _parseDOCX(file) {
        try {
            this._updateProgress('Converting document…');
            const ab    = await file.arrayBuffer();
            const bytes = new Uint8Array(ab);

            /* ── Sniff the real container format (extension is unreliable) ──
               • DOCX (and any Office Open XML) is a ZIP → first bytes "PK".
               • Word 2003 "XML Document" (.doc/.xml) is plain text starting
                 with "<?xml" / "<?mso-application" / "<w:wordDocument".
               • Legacy binary .doc (OLE2) starts with D0 CF 11 E0.
               mammoth.js ONLY understands the ZIP form — handing it a Word
               2003 XML file throws "Can't find end of central directory".   */
            const sig4 = String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]);
            const isZip = sig4.startsWith('PK');
            const isOle = bytes[0] === 0xD0 && bytes[1] === 0xCF &&
                          bytes[2] === 0x11 && bytes[3] === 0xE0;

            /* Decode a text head to detect Word 2003 WordprocessingML XML */
            const head = new TextDecoder('utf-8').decode(bytes.subarray(0, 512));
            const isWordXML =
                /<\?mso-application[^>]*Word\.Document/i.test(head) ||
                /<w:wordDocument/i.test(head) ||
                (/<\?xml/i.test(head) && /xmlns:w=|<w:/i.test(head));

            if (isWordXML) {
                /* Word 2003 XML — decode the WHOLE file and extract via our
                   WordprocessingML text reconstructor (NOT mammoth). */
                const xml = new TextDecoder('utf-8').decode(bytes);
                this._processText(this._extractWordXMLText(xml));
                return;
            }

            if (isZip) {
                if (typeof mammoth === 'undefined') {
                    throw new Error('mammoth.js library failed to load. Check your internet connection.');
                }
                const result = await mammoth.extractRawText({ arrayBuffer: ab });
                if (!result.value || result.value.trim().length === 0) {
                    throw new Error('No text could be extracted from the DOCX.');
                }
                this._processText(result.value);
                return;
            }

            if (isOle) {
                throw new Error(
                    'This is a legacy binary Word (.doc) file, which cannot be read in the ' +
                    'browser. Please open it in Word and "Save As" DOCX, PDF, or Word 2003 XML, ' +
                    'then re-upload.'
                );
            }

            /* Last resort: maybe it is text/XML after all — try the XML extractor */
            const xml = new TextDecoder('utf-8').decode(bytes);
            if (/<w:t[ >\/]/i.test(xml)) {
                this._processText(this._extractWordXMLText(xml));
                return;
            }
            throw new Error('Unrecognised document format.');
        } catch (err) {
            this._hideProgress();
            this._showError(
                `Failed to parse document: ${err.message}. ` +
                'Supported Word formats: DOCX, or Word 2003 XML (.doc/.xml). ' +
                'For a legacy binary .doc, re-save as DOCX or PDF.'
            );
        }
    },

    /* ══════════════════════════════════════════════════════
       WORD XML TEXT EXTRACTOR
       Reconstructs readable table rows from Microsoft Word XML
       (<w:t> elements fragmented across XML tags).
       ══════════════════════════════════════════════════════ */
    _extractWordXMLText(xml) {
        /* Step 1: collect text run content, preserving cell and row boundaries */
        let out = xml;

        /* Normalise self-closing w:t (shouldn't occur but just in case) */
        out = out.replace(/<w:t\s*\/>/g, '');

        /* Extract w:t content — handles attributes like xml:space="preserve" */
        out = out.replace(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g, '$1');

        /* Cell separator → tab, row separator → newline, paragraph → space (keep cell content on one line) */
        out = out.replace(/<\/w:tc>/g, '\t');
        out = out.replace(/<\/w:tr>/g, '\n');
        out = out.replace(/<\/w:p>/g, ' ');

        /* Strip all remaining XML tags */
        out = out.replace(/<[^>]+>/g, '');

        /* Decode common XML entities */
        out = out.replace(/&amp;/g, '&')
                 .replace(/&lt;/g, '<')
                 .replace(/&gt;/g, '>')
                 .replace(/&quot;/g, '"')
                 .replace(/&apos;/g, "'")
                 .replace(/&#160;/g, ' ')
                 .replace(/&#\d+;/g, ' ');

        /* Collapse excessive spaces while keeping newlines AND the tab cell
           delimiters emitted above (</w:tc>→\t). Tabs let the HOST OVERVIEW
           tabular parser read columns by index. Consecutive tabs (empty cells)
           are preserved so column positions don't shift. */
        out = out.split('\n').map(line =>
            line.replace(/ *\t */g, '\t').replace(/ {2,}/g, ' ').replace(/^[ \t]+|[ \t]+$/g, '')
        ).join('\n');

        return out;
    },

    /* ══════════════════════════════════════════════════════
       MAIN TEXT PROCESSING
       ══════════════════════════════════════════════════════ */
    _processText(rawText) {
        const text = rawText
            .replace(/\r\n/g, '\n')
            .replace(/\r/g, '\n')
            /* Collapse runs of spaces, but PRESERVE tabs as cell delimiters
               (the Word-XML extractor emits </w:tc>→\t). Spaces around a tab are
               absorbed into it; consecutive tabs (empty cells) are kept so the
               HOST OVERVIEW tabular parser can read columns by index. */
            .replace(/ *\t */g, '\t')
            .replace(/ {2,}/g, ' ');

        this._updateProgress('Extracting system information…');
        const sysInfo = this._extractSystemInfo(text);

        this._updateProgress('Parsing hardware configuration table…');
        const servers = this._extractHardwareTable(text);

        this._hideProgress();

        if (!servers || servers.length === 0) {
            this._showError(
                'Could not locate the Hardware Configuration data in this file. ' +
                'Please verify that you have uploaded a SAP EarlyWatch Alert report ' +
                'containing section "2.3 Hardware Configuration" or "Host Overview".'
            );
            return;
        }

        this.servers    = this._sortServers(servers);
        this.systemInfo = sysInfo;
        /* Physical→virtual host aliases from §2.2 (Instance Name / SAPDBHOST).
           Used by Module 02 to pre-fill the virtual↔physical mapping screen. */
        this.hostAliases = this._extractHostAliases(text);

        /* Detect IBM Power Systems presence — used to conditionally show the
           optional NMON upload step after "Verify HW Configuration" is clicked. */
        this.hasIBMServers = this._detectIBMServers(this.servers);

        this._renderResults(this.servers, sysInfo);

        this.app.setModuleData('system-landscape', {
            rows:          this.servers,
            fileName:      'EWA Report',
            headers:       Object.keys(this.servers[0] || {}),
            analysis:      this.servers,
            systemInfo:    sysInfo,
            appServers:    this.servers.filter(s => this._isAppLayer(s.component)),
            dbServers:     this.servers.filter(s => this._isDbLayer(s.component)),
            hostVariables: this._hostVariables(),
            hostAliases:   this.hostAliases,   /* { physicalLower -> virtualName } */
            hasIBMServers: this.hasIBMServers,
        });
        this.app.markCompleted('system-landscape');
    },

    /* True when any server in the inventory was manufactured by IBM.
       Case-insensitive. Used to gate the optional NMON upload step. */
    _detectIBMServers(servers) {
        return (servers || []).some(s =>
            /\bIBM\b/i.test(s.manufacturer || '')
        );
    },

    /* Order the inventory: application-layer servers first (A→Z by hostname),
       then the database server(s) last (A→Z). Non-app/non-DB rows sort between
       them. Hostname comparison is case-insensitive and digit-aware so e.g.
       HEIAPPRM2 sorts before HEIAPPRM10. */
    _sortServers(servers) {
        const rank = s => this._isAppLayer(s.component) ? 0 : (this._isDbLayer(s.component) ? 2 : 1);
        const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
        return servers.slice().sort((a, b) => {
            const ra = rank(a), rb = rank(b);
            if (ra !== rb) return ra - rb;
            return collator.compare(a.hostname || '', b.hostname || '');
        });
    },

    /* ══════════════════════════════════════════════════════
       EWA EXTRACTION — System Info
       ══════════════════════════════════════════════════════ */
    _extractSystemInfo(text) {
        const info = { sid: '', product: '', dbSystem: '', dbVersion: '', dbOs: '', dbSid: '' };

        /* ── SID (application tier) ──
           The SAP System ID appears as a suffix on MANY EWA section headings
           — "Performance Indicators for S4P", "Software Configuration for
           S4P", "Workload Distribution S4P", "Performance Overview S4P", "SAP
           System Operating S4P", etc.  This is far more reliable than the
           cover-page label/value table (which the PDF/Word text flattening
           reorders), so we collect SID candidates from these headings and
           take the most frequently occurring one (a majority vote that is
           robust even if a single heading is garbled). */
        info.sid = this._voteSID(text);

        /* Fallback: cover-page heuristics if no heading-based SID was found.
           The SID is a 3-char id (letter + 2 alphanumerics) sitting just
           before the product value, or after an explicit "SID"/"System ID"
           label. */
        if (!info.sid) {
            const sidM = text.match(
                /\b([A-Z][A-Z0-9]{2})\s+SAP\s+(?:S\/?4?HANA|ERP|NetWeaver|Business\s+Suite|Solution\s+Manager|BW)/i
            )
                      || text.match(/\bSID[:\s]+([A-Z][A-Z0-9]{2})\b/i)
                      || text.match(/System\s+ID[:\s]+([A-Z][A-Z0-9]{2})\b/i);
            if (sidM) info.sid = sidM[1].trim().toUpperCase();
        }

        info.product = this._extractProduct(text);
        if (!info.product) {
            /* Legacy fallback: first "Product … SAP <name>" occurrence. */
            const prodM = text.match(/Product\s*[\n:]*\s*(SAP\s+[^\n\r]{3,80})/i);
            if (prodM) info.product = prodM[1].trim().replace(/\s+/g, ' ');
        }

        /* ── Database product + version ──
           The authoritative source is the "DATABASES" table under section
           "2.1 Products and Components in current Landscape":
             Related System  Database System  Database Version  DB ID
             PRM~ABAP         ORACLE           19.27.0.          PRM
           Parse that row first; only fall back to loose whole-text scanning when
           the table is absent. */
        const dbRow = this._extractDatabasesTableRow(text);

        /* Database version — prefer the table's "Database Version" cell. */
        if (dbRow && dbRow.version) {
            info.dbVersion = dbRow.version;
        } else {
            const dbLineM = text.match(/Database\s+Version[^\n]*\n[^\n]*?(\d+\.\d{2}\.\d{3}[\.\d]*)/i)
                         || text.match(/(\d+\.\d{2}\.\d{3}[\.\d]*)/);
            if (dbLineM) info.dbVersion = dbLineM[1];
            if (!info.dbVersion) {
                const verM2 = text.match(/(?:SAP\s+HANA|HANA)\s+(?:Database\s+)?(?:Version\s*)?(\d+\.\d{2}\.\d{3}[\.\d]*)/i);
                if (verM2) info.dbVersion = verM2[1];
            }
        }

        /* Database product — normalize the table's "Database System" cell first;
           fall back to the whole-text heuristic only if the table was not found. */
        info.dbSystem = (dbRow && dbRow.system && this._normalizeDbProduct(dbRow.system))
                        || this._extractDbProduct(text, info.dbVersion);

        info.dbSid = this._extractDbSid(text, info.sid);

        return info;
    },

    /* ──────────────────────────────────────────────────────
       Parse the "DATABASES" table row from the landscape section.
       After PDF/Word text flattening the table reads as:
         Databases Related System Database System Database Version DB ID
         PRM~ABAP ORACLE 19.27.0. PRM
       Returns { system, version, dbId } (raw strings) or null when not found.
       The "Database System" value can be one or more words (e.g. "ORACLE",
       "SAP HANA Database", "Microsoft SQL Server"), so we capture everything
       between the "<SID>~<type>" related-system token and the version number.
       ────────────────────────────────────────────────────── */
    _extractDatabasesTableRow(text) {
        /* Anchor on the Databases table header; the data row follows it. */
        const hdrIdx = text.search(/Database\s+System\s+Database\s+Version\s+DB\s*ID/i);
        let scope;
        if (hdrIdx >= 0) {
            scope = text.slice(hdrIdx, hdrIdx + 600);
        } else {
            /* Looser anchor: a "Databases" heading followed by a related-system row. */
            const dbIdx = text.search(/\bDatabases\b/i);
            if (dbIdx < 0) return null;
            scope = text.slice(dbIdx, dbIdx + 600);
        }

        /* Data row: "<SID>~<type>  <Database System words>  <version>  [<DB ID>]".
           - Related system: letters/digits + "~" + letters/digits
           - Database System: non-greedy run of word tokens (letters, digits,
             spaces, slashes, +) up to the version number
           - Version: starts with a digit, e.g. 19.27.0. or 2.00.079.05 */
        const rowM = scope.match(
            /[A-Z0-9][A-Z0-9_]*~[A-Z0-9]+\s+([A-Za-z][A-Za-z0-9 /+().-]*?)\s+(\d[\d.]*\d|\d)\s*([A-Z][A-Z0-9]{2})?\b/
        );
        if (rowM) {
            return {
                system:  (rowM[1] || '').trim(),
                version: (rowM[2] || '').trim(),
                dbId:    (rowM[3] || '').trim(),
            };
        }
        return null;
    },

    /* Normalize a raw "Database System" string to the canonical product name
       used throughout the tool. Returns '' when unrecognised. */
    _normalizeDbProduct(raw) {
        const s = String(raw || '').trim();
        if (!s) return '';
        if (/HANA|\bHDB\b/i.test(s))                       return 'SAP HANA DB';
        if (/\bOracle\b/i.test(s))                         return 'Oracle';
        if (/IBM\s*DB2|\bDB6\b|\bDB2\b/i.test(s))          return 'IBM DB2 (DB6)';
        if (/Microsoft\s+SQL|\bMSSQL\b|SQL\s*Server/i.test(s)) return 'Microsoft SQL Server';
        if (/\bASE\b|Sybase/i.test(s))                     return 'SAP ASE';
        if (/MaxDB|\bADA\b/i.test(s))                      return 'SAP MaxDB';
        return '';
    },

    /* ──────────────────────────────────────────────────────
       Identify and NORMALIZE the database product.

       Search priority (section headings): Database → Database System →
       Database Information → Technical System Information → System Data →
       Landscape Information. The EWA "Databases" landscape table is the richest
       source:
         Databases Related System Database System Database Version DB ID
         EPR~ABAP SAP HANA Database 2.00.079.05 EHR
       We match explicit product names, normalize them, and finally fall back to
       inferring the product from the database version string.
       ────────────────────────────────────────────────────── */
    _extractDbProduct(text, dbVersion) {
        /* STEP 5 — HANA short-circuit: any strong HANA signal wins. */
        if (/\bHDB\b|HANA\s+DB|SAP\s+HANA|HANA\s+Version|HANA\s+Database/i.test(text)) {
            return 'SAP HANA DB';
        }

        /* STEP 2 — explicit database product names (longest/most specific first). */
        const PATTERNS = [
            [/SAP\s+HANA\s+Database|SAP\s+HANA\b|\bHDB\b/i,        'SAP HANA DB'],
            [/IBM\s+DB2|\bDB6\b|\bDB2\b/i,                          'IBM DB2 (DB6)'],
            [/Microsoft\s+SQL\s+Server|\bSQL\s+Server\b/i,         'Microsoft SQL Server'],
            [/SAP\s+ASE|\bASE\b|\bSybase\b/i,                      'SAP ASE'],
            [/SAP\s+MaxDB|\bMaxDB\b/i,                             'SAP MaxDB'],
            [/Oracle\s+Database|\bOracle\b/i,                      'Oracle'],
        ];
        for (const [re, name] of PATTERNS) {
            if (re.test(text)) return name;
        }

        /* STEP 4 — fall back to inferring from the database version string. */
        const v = (dbVersion || '').toString();
        if (/^\d+\.\d{2}\.\d{3}/.test(v)) return 'SAP HANA DB';   /* HANA dotted release */
        if (/^19\.|^12\.|^11\.2/.test(v)) return 'Oracle';
        if (/^11\.5|^10\.5/.test(v))      return 'IBM DB2 (DB6)';
        if (/^(20\d{2}|15\.|14\.)/.test(v)) return 'Microsoft SQL Server';
        if (/^16\.0/.test(v))             return 'SAP ASE';

        return '';
    },

    /* ──────────────────────────────────────────────────────
       Extract the SAP Product & Release ONLY from the
       "1.1 Products and Components in current Landscape" section.

       The table layout (after Word/PDF text flattening) is:
         Products and Components in current Landscape Product System SAP Product Product Version
         EPR~ABAP SAP S/4HANA 2021
       i.e. the data row is "<System>~<type> <SAP Product name> <Version>".

       We deliberately IGNORE SAP Fiori Add-On rows, Database versions, Main
       Instances, SAP_BASIS and application-server info — the landscape product
       table is the authoritative source. Returns e.g. "SAP S/4HANA 2021".
       ────────────────────────────────────────────────────── */
    _extractProduct(text) {
        const secStart = text.search(/Products\s+and\s+Components\s+in\s+current\s+Landscape/i);
        if (secStart < 0) return '';

        /* Bound the search to this section — stop at the next sub-table so we
           never read a Fiori Add-On / Databases / Main Instances row. */
        const rest = text.slice(secStart, secStart + 1200);
        const endM = rest.search(/Main\s+Instances|SAP\s+Fiori\s+Add-?Ons?|Databases\b|Servers\s+in\s+current\s+Landscape/i);
        const scope = endM > 0 ? rest.slice(0, endM) : rest;

        /* The data row is "<SID>~<type> <SAP Product> <Version>".
           Capture the SAP Product name (starts with "SAP ") up to the trailing
           version token (4-digit year, or NN.N / dotted release). */
        let m = scope.match(
            /[A-Z0-9][A-Z0-9_]*~[A-Z0-9]+\s+(SAP\s+[A-Za-z0-9/][A-Za-z0-9/ .+-]*?)\s+((?:19|20)\d{2}(?:\s+\w+)?|\d{1,2}\.\d{1,2}(?:\.\d+)?|[A-Z]{3,}\s*\d{3,})\b/
        );
        /* Fallback: no "System~type" prefix (some exports drop it) — match the
           product + version pair directly after the column headers. */
        if (!m) {
            m = scope.match(
                /SAP\s+Product\s+Product\s+Version\s+(?:[^\n]*?\s)?(SAP\s+[A-Za-z0-9/][A-Za-z0-9/ .+-]*?)\s+((?:19|20)\d{2}|\d{1,2}\.\d{1,2}(?:\.\d+)?)\b/i
            );
        }
        if (!m) return '';

        const product = m[1].replace(/\s+/g, ' ').trim();
        const version = (m[2] || '').trim();
        /* Guard: don't return a Fiori/HANA line even if it slipped through. */
        if (/fiori|hana\s+database/i.test(product)) return '';
        return version ? `${product} ${version}` : product;
    },

    /* ──────────────────────────────────────────────────────
       the many EWA section headings that carry it as a suffix.

       In an EWA the SID (e.g. "S4P") is repeated throughout the
       report in headings like:
         "Performance Indicators for S4P"
         "Software Configuration for S4P"
         "Workload Distribution S4P"
         "Performance Overview S4P"
         "SAP System Operating S4P"
       Any single heading could be garbled by PDF/Word text
       flattening, so we tally the 3-char candidate captured by
       every pattern across the whole document and return the one
       that occurs most often. This is far more robust than the
       cover-page positional guess.
       ────────────────────────────────────────────────────── */
    _voteSID(text) {
        /* A SID is 3 chars: a letter followed by two alphanumerics.
           Each pattern is a known EWA heading with the SID as a
           suffix. "for" is optional because PDF.js spatial ordering
           sometimes drops/relocates the connective word. */
        const patterns = [
            /Performance\s+Indicators\s+(?:for\s+)?([A-Z][A-Z0-9]{2})\b/gi,
            /Software\s+Configuration\s+(?:for\s+)?([A-Z][A-Z0-9]{2})\b/gi,
            /Workload\s+Distribution\s+(?:for\s+)?([A-Z][A-Z0-9]{2})\b/gi,
            /Performance\s+Overview\s+(?:for\s+)?([A-Z][A-Z0-9]{2})\b/gi,
            /SAP\s+System\s+Operating\s+(?:for\s+)?([A-Z][A-Z0-9]{2})\b/gi,
            /Hardware\s+Capacity\s+(?:for\s+)?([A-Z][A-Z0-9]{2})\b/gi,
            /Service\s+Summary\s+(?:for\s+)?([A-Z][A-Z0-9]{2})\b/gi,
        ];

        /* Common English/EWA words that fit the 3-char shape but are
           never SIDs — exclude so e.g. "Distribution across" cannot
           vote "ACR" and "for the" cannot vote "THE". */
        const BAD = new Set([
            'SAP', 'THE', 'ALL', 'AVG', 'FOR', 'AND', 'ERP', 'RFC', 'SQL',
            'NEW', 'OLD', 'MIN', 'MAX', 'ACR', 'PER', 'TOP', 'KEY', 'ARE',
            'USE', 'DAY', 'ONE', 'TWO', 'HAS', 'WAS', 'ITS', 'OUR',
        ]);

        const votes = {};
        for (const re of patterns) {
            let m;
            while ((m = re.exec(text)) !== null) {
                const cand = m[1].toUpperCase();
                if (BAD.has(cand)) continue;
                votes[cand] = (votes[cand] || 0) + 1;
            }
        }

        const ranked = Object.entries(votes).sort((a, b) => b[1] - a[1]);
        if (ranked.length) {
            console.log('[SystemLandscape] SID votes:', votes, '→ winner:', ranked[0][0]);
            return ranked[0][0];
        }
        return '';
    },

    /* ──────────────────────────────────────────────────────
       Extract the Database SID from the EWA report text.

       Priority chain:
       1. "Databases" table (HANA landscape section) — DB ID column
          e.g. "EPR~ABAP  SAP HANA Database  2.00.079.05  EHR"  → "EHR"
       2. Section heading "SAP HANA Database <SID>" — the SID in the
          heading title, e.g. "SAP HANA Database S4P" → "S4P"
       3. Explicit label fields: HANA System ID, Database Name,
          Instance Name, Database Instance, Connect String
       4. DB Host hostname heuristic (hdb/han/hana/db prefix/suffix)
       5. Fallback to application SID (non-HANA or undetectable)
       ────────────────────────────────────────────────────── */
    _extractDbSid(text, appSid) {
        const BAD_SID = new Set([
            'SAP','THE','ALL','AVG','FOR','AND','ERP','RFC','SQL','NEW','OLD',
            'MIN','MAX','ACR','PER','TOP','KEY','ARE','USE','DAY','ONE','TWO',
            'HAS','WAS','ITS','OUR','NOT','BUT','SET','GET','RUN','SYS','LOG',
            'NET','APP','API','ADM','INT','EXT','DEV','PRO','USR','TBL','IDX',
            'TMP','VAR','DIR','PKG','REG','SVC','HDB','LAN','WAN','NIC','RAM',
        ]);

        /* 1. "Databases" table — look for DB ID column value.
           The table header looks like:
             Related System  Database System  Database Version  DB ID
           followed by data rows with tab or multi-space separation.
           We search a window around the "DB ID" header occurrence.     */
        const dbIdHeaderIdx = text.search(/\bDB\s+ID\b/i);
        if (dbIdHeaderIdx >= 0) {
            /* Take up to 800 chars after the header — covers several data rows */
            const window = text.slice(dbIdHeaderIdx, dbIdHeaderIdx + 800);
            /* Rows after the header: find 3-char SID-shaped token that follows
               a version number (d.dd.ddd format) or appears last on a data line
               that contains "HANA" or a DB product name.                         */
            const rowM = window.match(
                /\d+\.\d{2}\.\d{3}[\.\d]*\s+([A-Z][A-Z0-9]{2})\b/i
            );
            if (rowM) {
                const cand = rowM[1].toUpperCase();
                if (!BAD_SID.has(cand)) return cand;
            }
            /* Broader fallback within the same window: last 3-char token on lines
               that look like database data rows (contain Related System marker ~) */
            const tildeM = window.match(/[A-Z0-9]+~[A-Z0-9]+[^\n]*\s([A-Z][A-Z0-9]{2})\b/i);
            if (tildeM) {
                const cand = tildeM[1].toUpperCase();
                if (!BAD_SID.has(cand)) return cand;
            }
        }

        /* 2. Section heading: "SAP HANA Database <SID>" */
        const hanaHeadingM = text.match(
            /SAP\s+HANA\s+Database\s+([A-Z][A-Z0-9]{2})\b/i
        );
        if (hanaHeadingM) {
            const cand = hanaHeadingM[1].toUpperCase();
            if (!BAD_SID.has(cand)) return cand;
        }

        /* 3. Explicit label fields */
        const labelPatterns = [
            /HANA\s+System\s+ID\s*[:\s]+([A-Z][A-Z0-9]{2})\b/i,
            /Database\s+(?:Name|Instance|Identifier)\s*[:\s]+([A-Z][A-Z0-9]{2})\b/i,
            /Instance\s+Name\s*[:\s]+([A-Z][A-Z0-9]{2})\b/i,
            /Connect\s+String\s*[:\s]+([A-Z][A-Z0-9]{2})\b/i,
            /DB\s+(?:System|Host|Name)\s*[:\s]+([A-Z][A-Z0-9]{2})\b/i,
        ];
        for (const pat of labelPatterns) {
            const m = text.match(pat);
            if (m) {
                const cand = m[1].toUpperCase();
                if (!BAD_SID.has(cand)) return cand;
            }
        }

        /* 4. DB hostname heuristic — extract a SID-shaped suffix/prefix from
           any host that is classified as DB layer */
        for (const s of this.servers || []) {
            if (!this._isDbLayer(s.component)) continue;
            const h = (s.hostname || '').toLowerCase();
            /* Patterns: h<SID>db01, <SID>hdb01, hdb<SID>, etc. */
            const hm = h.match(/^[a-z]{0,3}([a-z][a-z0-9]{2})(?:hdb|han|db)\d*/i)
                    || h.match(/(?:hdb|han|hana|db)([a-z][a-z0-9]{2})\d*/i);
            if (hm) {
                const cand = hm[1].toUpperCase();
                if (!BAD_SID.has(cand) && cand !== appSid) return cand;
            }
        }

        /* 5. Fallback: same as app SID (common for non-HANA or co-located setups) */
        return appSid || '';
    },

    /* ══════════════════════════════════════════════════════
       EWA EXTRACTION — Hardware Table
       ══════════════════════════════════════════════════════ */
    _extractHardwareTable(text) {
        const compMap = this._buildComponentMap(text);

        /* ── Locate the hardware section ── */
        const hwStart = this._findSectionStart(text, [
            /2\.3\s+Hardware\s+Configuration/i,
            /Host\s+Overview/i,
            /Hardware\s+Configuration/i,
        ]);

        const hwEnd = hwStart >= 0
            ? this._findSectionStart(text, [
                /* These work for both PDF (line-anchored) and Word XML (space-joined paragraphs).
                   No ^ anchor so they match even when headings are embedded in flowing text. */
                /\b2\.4[\s\W]/i,
                /Transport\s+Landscape/i,
                /Software\s+Configuration/i,
                /Service\s+Data\s+Quality/i,
                /Service\s+Preparation/i,
                /Hardware\s+Utilization\s+Data/i,
              ], hwStart + 80)
            : -1;

        const sectionText = hwStart >= 0
            ? text.slice(hwStart, hwEnd > hwStart ? hwEnd : hwStart + 16000)
            : text;

        /* ── Split into lines, skip noise ── */
        const lines = sectionText
            .split('\n')
            .map(l => l.trim())
            .filter(l => l.length > 1);

        /*
         * BLOCK-BASED GROUPING
         * Because EWA PDF tables often place each cell on its own line (different Y),
         * we collect ALL lines between two consecutive hostname detections and treat
         * them as a single server entry — no fixed window size.
         */
        const blocks = [];   // [{ hostname, lines[] }]
        let cur = null;
        const seenHosts = new Set();

        for (const line of lines) {
            /* Hard stops — page numbers, confidentiality banners, section headers */
            if (/^\d{1,3}\s*\/\s*\d{1,3}$|^Confidential|^SAP\s+SE|^www\.|©/i.test(line)) {
                continue;
            }
            if (/^Host\s+Hardware|^CPU\s+Type|^Operating\s+System|^Memory\s+in\s+MB/i.test(line)) {
                continue;
            }
            if (/^2\.\d+\s+|^3\.\d+\s+|^4\.\d+\s+/i.test(line)) {
                continue;
            }

            /* Check if this line STARTS a new server block */
            const firstWord = (line.split(/[\s\t]/)[0] || '').replace(/[^\w\-\.]/g, '');
            if (firstWord && this._isServerHostname(firstWord) && !seenHosts.has(firstWord.toLowerCase())) {
                if (cur) blocks.push(cur);
                cur = { hostname: firstWord, lines: [line] };
                seenHosts.add(firstWord.toLowerCase());
            } else if (cur) {
                cur.lines.push(line);
            }
        }
        if (cur) blocks.push(cur);

        /* ── Parse each block into a server entry ── */
        /* First build the structured HOST OVERVIEW map so exact Model and
           Virtualization values override the regex-inferred ones.         */
        const hostOverviewMap = this._extractHostOverviewMap(sectionText);

        const result = [];
        for (const block of blocks) {
            const entryText = block.lines.join(' ').replace(/\s{2,}/g, ' ');
            const server = this._parseHardwareLine(block.hostname, entryText, compMap);
            if (server) {
                /* Apply exact Model / Virtualization from the structured table
                   whenever available — overrides the regex-guessed defaults.  */
                const ov = hostOverviewMap[block.hostname.toLowerCase()];
                if (ov) {
                    if (ov.model)          server.model          = ov.model;
                    if (ov.virtualization) server.virtualization = ov.virtualization;
                }
                result.push(server);
            }
        }

        return result;
    },

    /* ══════════════════════════════════════════════════════
       HOST OVERVIEW STRUCTURED TABLE PARSER

       Extracts exact Model and Virtualization values from the
       "HOST OVERVIEW" table in §2.3, returning a map:
         { hostnameLower → { model, virtualization } }

       The HOST OVERVIEW table has a FIXED column order:
         0 Host | 1 Hardware Manufacturer | 2 Model | 3 CPU Type |
         4 CPU MHz | 5 Virtualization | 6 Operating System |
         7 CPUs | 8 Cores | 9 Memory in MB

       Values are copied VERBATIM from the Model and Virtualization
       columns (no normalization) per the bug-fix spec.

       Three source layouts are supported, discovered automatically:
         • TABULAR  : one table row per text line with TAB-separated cells
                      (Word 2003 XML extractor emits </w:tc>→TAB, </w:tr>→\n).
                      Columns are read directly by index — the most reliable.
         • CELL/LINE: each table cell on its own line (mammoth DOCX export).
                      Rows are reassembled from the known 10-column sequence.
         • INLINE   : all columns space-joined on one line (PDF Y-row join).
                      Columns are isolated structurally (manufacturer / CPU
                      family / OS anchors) rather than by fragile char offsets.
       ══════════════════════════════════════════════════════ */
    _extractHostOverviewMap(sectionText) {
        const map = {};   /* hostLower → { model, virtualization } */

        const rawLines = sectionText.split('\n');

        /* ── Locate the HOST OVERVIEW header line ── */
        let headerIdx = -1;
        for (let i = 0; i < rawLines.length; i++) {
            const l = rawLines[i].toLowerCase();
            if (l.includes('model') && l.includes('host') &&
                (l.includes('virtualization') || l.includes('cpu type'))) {
                headerIdx = i;
                break;
            }
        }
        if (headerIdx < 0) return map;   /* no HOST OVERVIEW table found */

        /* ── Shared vocab for recognising columns by content ──
           Virtualization strings seen in EWA reports include IBM PowerVM
           partition modes ("SPLPAR, UNCAPPED", "DLPAR", "Dedicated"),
           hypervisors (KVM, VMware/vSphere, Hyper-V, Xen) and cloud. */
        const VIRT_RE = /\b(SPLPAR|DLPAR|LPAR|UNCAPPED|CAPPED|PowerVM|Micro[- ]?Partition|Shared\s+Pool|Dedicated|KVM|VMware|vSphere|ESXi|Hyper-?V|Xen|z\/VM|Physical|Bare[- ]?Metal|Amazon\s+VPC|AWS|Azure|HANA\s+Cloud|HCl)\b/i;
        const MFG_RE  = /\b(Amazon\s+EC2|AWS|Google\s+Cloud|Microsoft\s+Azure|Azure|HPE|Hewlett[- ]Packard(?:\s+Enterprise)?|Dell(?:\s+EMC)?|IBM|Cisco\s+UCS|Fujitsu|Lenovo|SuperMicro|Super\s+Micro|Oracle)\b/i;
        const CPU_RE  = /\b(Xeon|EPYC|POWER\d|POWER|Itanium|Opteron|Core\s+i\d|Graviton|Ampere)\b/i;
        const OS_RE   = /\b(SUSE|SLES|Linux|Windows|RedHat|Red\s+Hat|RHEL|AIX|Oracle\s+Linux|IBM\s+i)\b/i;
        const isPureNum = t => /^\d{2,7}$/.test(String(t).trim());

        /* ── Layout discovery: does the header line carry TAB separators? ── */
        const headerHasTabs = /\t/.test(rawLines[headerIdx]);

        /* Column index resolver: map header cell labels → canonical positions.
           Falls back to the fixed EWA order when labels can't be matched.      */
        const resolveCols = (headerCells) => {
            const find = re => headerCells.findIndex(c => re.test(c));
            const idx = {
                host:  find(/^host$|hostname/i),
                mfg:   find(/manufacturer/i),
                model: find(/model/i),
                cpu:   find(/cpu\s*type/i),
                mhz:   find(/mhz|cpu\s*mhz/i),
                virt:  find(/virtual/i),
                os:    find(/operating|^os$/i),
            };
            /* Fallbacks to the documented fixed order */
            if (idx.host  < 0) idx.host  = 0;
            if (idx.model < 0) idx.model = 2;
            if (idx.virt  < 0) idx.virt  = 5;
            return idx;
        };

        const commit = (host, model, virt) => {
            if (!host) return;
            const key = host.toLowerCase();
            const m = (model || '').trim();
            const v = (virt  || '').trim();
            if (!m && !v) return;
            /* First writer wins; never clobber a found value with an empty one. */
            const prev = map[key] || { model: '', virtualization: '' };
            map[key] = {
                model:          prev.model          || m,
                virtualization: prev.virtualization || v,
            };
        };

        /* ────────────────────────────────────────────────────────────
           LAYOUT A — TABULAR (tab-separated, one row per line)
           ──────────────────────────────────────────────────────────── */
        if (headerHasTabs) {
            const headerCells = rawLines[headerIdx].split('\t').map(c => c.trim());
            const col = resolveCols(headerCells);

            for (let i = headerIdx + 1; i < rawLines.length; i++) {
                const line = rawLines[i];
                if (/^\s*2\.\d+[\s\W]|^\s*3\.\d+[\s\W]/.test(line)) break;
                if (!/\t/.test(line)) continue;                 /* not a table row */
                if (/^Confidential|^SAP\s+SE|©/i.test(line.trim())) continue;

                const cells = line.split('\t').map(c => c.trim());
                const host  = (cells[col.host] || '').replace(/[^\w\-\.]/g, '');
                if (!host || !this._isServerHostname(host)) continue;

                commit(host, cells[col.model], cells[col.virt]);
            }

            if (Object.keys(map).length) {
                this._logHostOverview(map);
                return map;
            }
            /* If tabular parse yielded nothing, fall through to the heuristics. */
        }

        /* ────────────────────────────────────────────────────────────
           Build a trimmed, noise-free line list for the fallbacks.
           ──────────────────────────────────────────────────────────── */
        const lines = rawLines
            .map(l => l.replace(/\t/g, ' ').replace(/\s{2,}/g, ' ').trim())
            .filter(l => l.length > 0);

        /* Re-locate the header within the trimmed list. */
        let hIdx = -1;
        for (let i = 0; i < lines.length; i++) {
            const l = lines[i].toLowerCase();
            if (l.includes('model') && l.includes('host') &&
                (l.includes('virtualization') || l.includes('cpu type'))) {
                hIdx = i;
                break;
            }
        }
        if (hIdx < 0) { this._logHostOverview(map); return map; }

        const headerLine = lines[hIdx];

        /* INLINE vs CELL-PER-LINE: if the line AFTER the header is itself a
           header fragment (just a column label), the cells are split per line.
           If the header holds every label on one physical line AND the next
           data line carries a hostname followed by several more tokens, it's
           inline.                                                              */
        const nextDataLine = lines[hIdx + 1] || '';
        const nextTokCount = nextDataLine.split(/\s+/).filter(Boolean).length;
        const headerLabelCount = (headerLine.match(/model|virtual|cpu|operating|memory|cores|manufacturer|host/gi) || []).length;
        const isInline = headerLabelCount >= 3 && nextTokCount >= 4;

        if (isInline) {
            /* ── LAYOUT C — INLINE (PDF) ──
               Group continuation lines into one block per host, then isolate
               the Model column as the tokens between the manufacturer and the
               CPU-family / CPU-MHz anchor. */
            let dataStart = hIdx + 1;
            while (dataStart < lines.length &&
                   /\b(manufacturer|cpu\s*mhz|operating\s+system|memory\s+in\s+mb)\b/i.test(lines[dataStart]) &&
                   !this._isServerHostname((lines[dataStart].split(/\s+/)[0] || '').replace(/[^\w\-\.]/g, ''))) {
                dataStart++;
            }

            const blocks = [];
            let curHost = null, buf = [];
            for (let i = dataStart; i < lines.length; i++) {
                const l = lines[i];
                if (/^2\.\d+[\s\W]|^3\.\d+[\s\W]/.test(l)) break;
                if (/^Confidential|^SAP\s+SE|©/i.test(l)) continue;
                const firstTok = (l.split(/\s+/)[0] || '').replace(/[^\w\-\.]/g, '');
                if (firstTok && this._isServerHostname(firstTok)) {
                    if (curHost) blocks.push({ host: curHost, raw: buf.join(' ') });
                    curHost = firstTok;
                    buf = [l];
                } else if (curHost) {
                    buf.push(l);
                }
            }
            if (curHost) blocks.push({ host: curHost, raw: buf.join(' ') });

            for (const { host, raw } of blocks) {
                /* Virtualization — longest contiguous run of VIRT tokens
                   (captures "SPLPAR, UNCAPPED" and the like). */
                let virt = '';
                const vm = raw.match(VIRT_RE);
                if (vm) {
                    /* Grab the comma/space run around the first virt hit so
                       compound values like "SPLPAR, UNCAPPED" survive intact. */
                    const around = raw.slice(vm.index).match(/^([A-Za-z][\w]*(?:\s*,\s*[A-Za-z][\w]*)*)/);
                    virt = (around ? around[1] : vm[1]).replace(/\s*,\s*/g, ', ').trim();
                }

                /* Model — tokens between manufacturer and the CPU anchor. */
                const toks = raw.split(/\s+/).filter(Boolean);
                const hostPos = toks.findIndex(t =>
                    t.replace(/[^\w\-\.]/g, '').toLowerCase() === host.toLowerCase());
                const after = hostPos >= 0 ? toks.slice(hostPos + 1) : toks;
                let mfgEnd = 0;
                for (let j = 0; j < after.length; j++) {
                    if (MFG_RE.test(after[j]) || /^EC2$/i.test(after[j])) mfgEnd = j + 1;
                    else if (mfgEnd > 0) break;
                }
                let cpuPos = after.findIndex((t, j) => j >= mfgEnd && CPU_RE.test(t));
                if (cpuPos < 0) cpuPos = after.findIndex((t, j) => j >= mfgEnd && isPureNum(t));
                let model = '';
                if (cpuPos > mfgEnd) {
                    model = after.slice(mfgEnd, cpuPos)
                        .filter(t => !isPureNum(t) && !VIRT_RE.test(t) && !OS_RE.test(t))
                        .join(' ')
                        .replace(/\s*-\s*/g, '-')
                        .trim();
                }
                commit(host, model, virt);
            }

        } else {
            /* ── LAYOUT B — CELL-PER-LINE (mammoth DOCX) ──
               Each cell is its own line. Reassemble a row from the known
               10-column sequence starting at each hostname line:
                 0 host | 1 manufacturer | 2 model | 3 cpu | 4 mhz |
                 5 virtualization | 6 os | 7 cpus | 8 cores | 9 memory      */
            let i = hIdx + 1;
            /* Skip trailing header-label lines */
            while (i < lines.length &&
                   /^(model|virtualization|cpu\s*type|cpu\s*mhz|operating\s+system|hardware\s+manufacturer|cpus|cores|memory|host)\b/i.test(lines[i])) {
                i++;
            }

            while (i < lines.length) {
                const l = lines[i];
                if (/^2\.\d+[\s\W]|^3\.\d+[\s\W]/.test(l)) break;
                if (/^Confidential|^SAP\s+SE|©/i.test(l)) { i++; continue; }

                const firstTok = (l.split(/\s+/)[0] || '').replace(/[^\w\-\.]/g, '');
                if (!firstTok || !this._isServerHostname(firstTok)) { i++; continue; }

                /* Collect this host's cells up to (but not including) the next host. */
                const cells = [l.trim()];
                i++;
                while (i < lines.length && cells.length < 11) {
                    const nl = lines[i];
                    if (/^2\.\d+[\s\W]|^3\.\d+[\s\W]/.test(nl)) break;
                    if (/^Confidential|^SAP\s+SE|©/i.test(nl)) { i++; continue; }
                    const ntok = (nl.split(/\s+/)[0] || '').replace(/[^\w\-\.]/g, '');
                    if (ntok && this._isServerHostname(ntok)) break;
                    cells.push(nl.trim());
                    i++;
                }

                /* The hostname cell may also carry the manufacturer inline
                   ("fsascmp03 IBM"). Detect whether cell[1] is the manufacturer
                   or the model to set the column offset. */
                let base = 1;  /* index of manufacturer cell */
                const hostCellRest = cells[0].replace(new RegExp('^' + firstTok, 'i'), '').trim();
                if (hostCellRest && MFG_RE.test(hostCellRest)) {
                    base = 0;  /* manufacturer was inline with host → model is cells[1] */
                }

                /* A single table cell whose text wrapped onto two visual lines
                   (e.g. Model "9043-MRX" → "9043-"/"MRX", or Virtualization
                   "SPLPAR, UNCAPPED" → "SPLPAR,"/"UNCAPPED") can surface as two
                   consecutive cell-lines in the mammoth export. Stitch an obvious
                   continuation fragment back onto the preceding cell so the value
                   is preserved verbatim and column positions don't drift. */
                const stitched = [];
                for (let k = 0; k < cells.length; k++) {
                    const c = cells[k];
                    const prev = stitched[stitched.length - 1];
                    const isContinuation =
                        prev !== undefined && (
                            /[,\-]$/.test(prev) ||                       /* prev ends with , or - */
                            /^(UNCAPPED|CAPPED|Pool|Partition)$/i.test(c) /* known 2nd word */
                        ) && !this._isServerHostname((c.split(/\s+/)[0] || ''));
                    if (isContinuation) {
                        const joiner = /-$/.test(prev) ? '' : ' ';
                        stitched[stitched.length - 1] = (prev + joiner + c).trim();
                    } else {
                        stitched.push(c);
                    }
                }
                const scells = stitched;

                const modelCell = scells[base + 1] || '';   /* 2nd col after host(+mfg) */
                /* Virtualization: prefer the positional cell, else keyword scan. */
                let virt = '';
                const posVirt = scells[base + 4] || '';      /* after model, cpu, mhz */
                if (VIRT_RE.test(posVirt)) {
                    virt = posVirt.replace(/\s*,\s*/g, ', ').trim();
                } else {
                    for (const c of scells) {
                        if (VIRT_RE.test(c)) { virt = c.replace(/\s*,\s*/g, ', ').trim(); break; }
                    }
                }

                /* Model: accept the positional cell only if it isn't a CPU /
                   OS / manufacturer / pure-number cell. */
                let model = '';
                if (modelCell && !CPU_RE.test(modelCell) && !OS_RE.test(modelCell) &&
                    !MFG_RE.test(modelCell) && !VIRT_RE.test(modelCell) && !isPureNum(modelCell)) {
                    model = modelCell.trim();
                }

                commit(firstTok, model, virt);
            }
        }

        this._logHostOverview(map);
        return map;
    },

    _logHostOverview(map) {
        try {
            console.log('[SystemLandscape] host overview map (model/virt):',
                        JSON.parse(JSON.stringify(map)));
        } catch (_) { /* console may be unavailable */ }
    },

    _findSectionStart(text, patterns, afterPos = 0) {
        for (const pat of patterns) {
            const sub = text.slice(afterPos);
            const m   = sub.search(pat);
            if (m >= 0) return afterPos + m;
        }
        return -1;
    },

    /* Returns true if `s` looks like a server hostname, not a section title or keyword */
    _isServerHostname(s) {
        if (!s || s.length < 5 || s.length > 30) return false;
        /* Must contain only word chars, hyphens, dots */
        if (!/^[a-zA-Z][a-zA-Z0-9\-\.]+$/.test(s)) return false;
        /* Exclude known non-hostname words */
        if (this.NON_HOSTNAME.has(s.toLowerCase())) return false;
        /* All-uppercase tokens without any digit are acronyms / section headings, not hostnames.
           Real server hostnames virtually always contain at least one digit (e.g. psapecc022). */
        if (s === s.toUpperCase() && !/\d/.test(s)) return false;
        /* Cloud / hardware model strings that look like hostnames but aren't:
           • AWS EC2 instance types:  m6a.8xlarge, u-6tb1.112xlarge, c5.2xlarge
           • Generic "<letters><digits>.<digits>xlarge" or "<letters><digits>.<digits>large"
           These appear as separate lines in mammoth-extracted DOCX tables and must
           not be mistaken for a new server block. */
        if (/\.\d+x?large$/i.test(s)) return false;
        return true;
    },

    /* Build { hostname.toLowerCase() → componentType } from §2.2 landscape tables.
       A host may carry several roles; roles are collapsed to one label:
         DB SERVERS                      → role 'db'
         COMPONENTS, ABAP SCS row        → role 'scs'
         SAP APPLICATION SERVERS         → role 'app'
       Collapse: db → "DB";  app+scs → "SAP AS / SAP ABAP SCS";
                 scs → "SAP ABAP SCS";  app → "SAP AS".

       Real EWA headings are ALL-CAPS; we match case-insensitively.
       Column order in COMPONENTS table: System | Component | Host | ...
       i.e. the hostname comes AFTER the component label on each row.           */
    _buildComponentMap(text) {
        const roles = {};   /* host -> Set of 'app' | 'scs' | 'db' */
        const addRole = (h, r) => { (roles[h] || (roles[h] = new Set())).add(r); };

        /* Helper: collect every hostname-like token in a block.
           Robust to PDF text-flattening where the "Host" column is not at a
           fixed position — a hostname is a token with letters AND digits,
           length >= 5, not a known keyword. */
        const hostsIn = (block) => {
            const out = [];
            const re = /\b([a-z][a-z0-9\-\.]{4,29})\b/gi;
            let m;
            while ((m = re.exec(block || '')) !== null) {
                const h = m[1].toLowerCase();
                if (this.NON_HOSTNAME.has(h)) continue;
                if (!/\d/.test(h)) continue;          /* real hosts carry a digit */
                if (!/^[a-z]/.test(h)) continue;
                out.push(h);
            }
            return out;
        };

        /* Helper: first hostname-like token in a string (used to find the host
           that FOLLOWS an "ABAP SCS" marker — in PDF text the component label
           and its host land on separate lines, so we cannot require them on
           the same line). */
        const firstHostIn = (block) => hostsIn(block)[0] || null;

        /* Digit-optional host finder, used only for the SCS host that follows an
           "ABAP SCS" marker (some SCS hostnames carry no digit, e.g. "pkgsape").
           Still excludes known non-hostname keywords. */
        const firstHostLoose = (block) => {
            const re = /\b([a-z][a-z0-9\-\.]{4,29})\b/gi;
            let m;
            while ((m = re.exec(block || '')) !== null) {
                const h = m[1].toLowerCase();
                if (this.NON_HOSTNAME.has(h)) continue;
                if (!/^[a-z]/.test(h)) continue;
                return h;
            }
            return null;
        };

        /* ── Locate the three sub-table headings of the "Servers in current
           Landscape" section ("SAP Application Servers", "DB Servers",
           "Components").

           This must work against TWO very different text layouts:

           • PDF (PDF.js): text items are grouped into spatial rows after a
             Y-axis FLIP, so a table's data rows appear BEFORE its heading, and
             the three tables come out in REVERSE page order
             (Components → DB → App).  PDF.js also letter-spaces small-caps
             headings, so the raw text reads "SAP A PPLICATION  S ERVERS",
             "DB  S ERVERS", "C OMPONENTS".
           • DOC/DOCX (Word XML flattening): data rows FOLLOW their heading, in
             natural page order (App → DB → Components), headings are not
             letter-spaced.

           To be robust to both we (a) match headings tolerant of letter
           spacing, (b) pick the tight CLUSTER of the three real headings
           (there are many stray "Components" words elsewhere in the report),
           and (c) detect which side the data rows are on and claim the gap
           between adjacent headings accordingly. ── */

        /* Build a regex that tolerates an optional space between every letter
           of a word (PDF.js small-caps letter spacing). */
        const spaced = (w) => w.split('').map(c => c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('\\s*');
        const reApp = new RegExp('SAP\\s+' + spaced('APPLICATION') + '\\s+' + spaced('SERVERS') + '?', 'i');
        /* DB heading must be the PLURAL "DB Servers" — the singular "DB Server"
           appears in prose ("Max. CPU Utilization on DB Server 3 %") and must
           NOT be mistaken for the landscape sub-table heading. */
        const reDb  = new RegExp('\\b' + spaced('DB') + '\\s+' + spaced('SERVERS'), 'i');
        const reCmp = new RegExp('\\b' + spaced('COMPONENTS') + '?\\b', 'i');
        const reScs = /ABAP\s*SCS|ASCS/gi;

        /* The three landscape sub-table headings are each immediately followed
           by their column labels. Use this to tell a real heading apart from a
           same-worded phrase elsewhere in the report. */
        const isRealHeading = (idx, len) => {
            const after = text.slice(idx + len, idx + len + 80);
            return /Related\s+System|Logical\s+Host|SAPDBHOST|Instance\s+Name|\bHost\b/i.test(after);
        };

        const allIdx = (re) => {
            const out = [];
            const g = new RegExp(re.source, re.flags.includes('g') ? re.flags : re.flags + 'g');
            let m;
            while ((m = g.exec(text)) !== null) { out.push({ index: m.index, len: m[0].length }); if (m.index === g.lastIndex) g.lastIndex++; }
            return out;
        };

        /* App & DB & Components headings all sit in one tight cluster, but the
           report body also contains stray matches ("DB Server 43 %" in a
           performance table, dozens of "Components" words, etc.).  The
           "SAP Application Servers" heading is the most specific and reliably
           unique, so use it as the ANCHOR and pick the DB / Components
           occurrences NEAREST to it (within a window). Restrict DB/Components
           candidates to matches that are FOLLOWED by their table's column
           labels, which eliminates same-worded prose elsewhere in the report. */
        const appHitsAll = allIdx(reApp);
        const dbHits  = allIdx(reDb).filter(h => isRealHeading(h.index, h.len)).map(h => h.index);
        const cmpHits = allIdx(reCmp).filter(h => isRealHeading(h.index, h.len)).map(h => h.index);

        const iApp = appHitsAll.length ? appHitsAll[0].index : -1;

        const nearest = (hits, anchor, window) => {
            if (!hits.length) return -1;
            if (anchor < 0) return hits[0];
            const best = hits.reduce((b, c) =>
                Math.abs(c - anchor) < Math.abs(b - anchor) ? c : b, hits[0]);
            return Math.abs(best - anchor) <= window ? best : -1;
        };

        /* DB Servers heading must be within ~1500 chars of the App Servers
           heading to count as part of the same section cluster.  The
           Components heading sits immediately adjacent to DB Servers, so
           anchor its search on the DB heading (this avoids the nearby stray
           "Products and Components in current Landscape" §2.1 heading, which
           is closer to App than the real letter-spaced "C OMPONENTS"). */
        const iDb  = nearest(dbHits, iApp, 1500);
        const iCmp = nearest(cmpHits, (iDb >= 0 ? iDb : iApp), 1500);

        const marks = [[iApp, 'app'], [iDb, 'db'], [iCmp, 'cmp']]
            .filter(([i]) => i >= 0).sort((a, b) => a[0] - b[0]);

        const MARGIN = 1400;   /* reach for the outermost heading's data block */

        /* Data rows may PRECEDE their heading (PDF: PDF.js groups spatial rows
           after a Y-flip, so rows come out above the heading and the three
           tables are in reverse page order) or FOLLOW it (DOC/DOCX: natural
           order).  Rather than guess the orientation from a fragile heuristic,
           extract roles for BOTH orientations and keep whichever classifies
           more hosts and finds at least one DB host — this is self-correcting
           across report formats. */
        const extract = (rowsFollow) => {
            const r = {};
            const add = (h, role) => { (r[h] || (r[h] = new Set())).add(role); };
            const seg = {};
            marks.forEach(([idx, name], k) => {
                if (rowsFollow) {
                    const end = (k + 1 < marks.length) ? marks[k + 1][0] : idx + MARGIN;
                    seg[name] = text.slice(idx, end);
                } else {
                    const start = (k > 0) ? marks[k - 1][0] : Math.max(0, idx - MARGIN);
                    seg[name] = text.slice(start, idx);
                }
            });
            const appHosts = new Set(hostsIn(seg.app || ''));
            const dbHosts  = new Set(hostsIn(seg.db  || ''));
            for (const h of appHosts) add(h, 'app');
            for (const h of dbHosts)  add(h, 'db');
            let scsCount = 0;
            if (seg.cmp) {
                reScs.lastIndex = 0;
                let m;
                while ((m = reScs.exec(seg.cmp)) !== null) {
                    /* The SCS host sits right AFTER the "ABAP SCS" marker, in a
                       row like "EPR~ABAP ABAP SCS pkgsape pkgsape_EPR_00 …".
                       Some SCS hostnames carry no digit (e.g. "pkgsape"), which
                       the strict digit-requiring hostsIn would drop — so use a
                       digit-optional finder on the text that follows the marker. */
                    const after = seg.cmp.slice(m.index + m[0].length, m.index + m[0].length + 120);
                    const h = firstHostLoose(after) || firstHostIn(seg.cmp.slice(Math.max(0, m.index - 120), m.index + m[0].length + 120));
                    if (h) { add(h, 'scs'); scsCount++; }
                }
            }
            return { roles: r, nApp: appHosts.size, nDb: dbHosts.size, nScs: scsCount };
        };

        /* Score the plausibility of a parse orientation.  A real EWA landscape
           lists MANY application servers but only one or two database hosts
           (and the "Components / ABAP SCS" row names a single host).  The
           wrong orientation inverts this (the app rows get mis-filed under the
           DB heading), so reward app-heavy / DB-light parses. */
        const score = (res) => {
            if (res.nApp === 0 && res.nDb === 0) return -1;
            return res.nApp * 2 - res.nDb * 3 + ((res.nDb >= 1 && res.nDb <= 2) ? 5 : 0);
        };

        const follow  = extract(true);
        const precede = extract(false);
        const chosen  = score(precede) > score(follow) ? precede : follow;
        for (const [h, set] of Object.entries(chosen.roles)) {
            for (const role of set) addRole(h, role);
        }

        /* ── Collapse role sets into Component Type labels ── */
        const map = {};
        for (const [h, set] of Object.entries(roles)) {
            if (set.has('db'))                       map[h] = 'DB';
            else if (set.has('app') && set.has('scs')) map[h] = 'SAP AS / SAP ABAP SCS';
            else if (set.has('scs'))                 map[h] = 'SAP ABAP SCS';
            else                                     map[h] = 'SAP AS';
        }

        /* Diagnostic (visible in browser DevTools console) — helps verify the
           landscape segmentation against any new EWA layout. */
        try {
            console.log('[SystemLandscape] component map:', JSON.parse(JSON.stringify(map)));
            console.log('[SystemLandscape] heading idx — app:', iApp, 'db:', iDb, 'cmp:', iCmp,
                        '| order:', marks.map(m => m[1]).join(',') || '(none)',
                        '| orientation:', chosen === precede ? 'rows-precede (PDF)' : 'rows-follow (DOC)',
                        '| scores follow/precede:', score(follow) + '/' + score(precede));
        } catch (_) { /* console may be unavailable */ }

        return map;
    },

    /* ══════════════════════════════════════════════════════
       EWA EXTRACTION — Physical → Virtual host aliases (§2.2)

       Pulls the virtual (logical) hostname that corresponds to each physical
       host from the "Servers in current Landscape" sub-tables, so Module 02
       can pre-fill its virtual↔physical mapping screen.

       Source columns:
         SAP APPLICATION SERVERS   System | Host | Instance Name | Logical Host | …
         DB SERVERS                Related System | Host | Logical Host (SAPDBHOST)
         COMPONENTS (ABAP SCS row) System | Component | Host | Instance Name | Logical Host

       Each data row is emitted by both the PDF (PDF.js Y-row join) and the
       DOC/DOCX flatteners as a single line beginning with a "SID~stack"
       Related-System token, e.g.:
         S4P~ABAP hec50v038480 vhcxls4pai09_S4P_00 vhcxls4pai09
         S4P~ABAP hec50v018713 vhcxls4pdb
         S4P~ABAP ABAP SCS hec50v018775 vhcxls4pcs_S4P_01 vhcxls4pcs

       The virtual name is the Instance Name with its trailing "_<SID>_<nn>"
       suffix removed (SID = 3 chars, instance number = 2 digits) — which equals
       the Logical Host column — or, for DB servers, the SAPDBHOST token.

       Returns { physicalHost(lowercase) -> virtualName(original case) }.       */
    _extractHostAliases(text) {
        const aliases = {};

        /* A §2.2 data row starts with a "SID~stack" Related-System token
           (e.g. "S4P~ABAP"). Prose and other tables don't carry this token at
           the start of a line, which keeps the scan tightly scoped. */
        const rowRe = /^\s*[A-Za-z0-9]{3}~[A-Za-z0-9]+\b(.*)$/;
        /* Instance-Name suffix: _<SID>_<instanceNo>  (3-char SID, 2-digit nr). */
        const instSuffix = /_[A-Za-z0-9]{3}_\d{2}$/;
        /* Host-like token: letters+digits, len >= 5, not a known keyword. */
        const isHostToken = (t) => {
            const h = t.toLowerCase();
            if (this.NON_HOSTNAME.has(h)) return false;
            if (!/^[a-z][a-z0-9\-.]{4,40}$/i.test(t)) return false;
            if (!/\d/.test(t)) return false;        /* physical hosts carry a digit */
            return true;
        };
        /* Virtual-name token: letters first, len >= 3; may lack a digit
           (e.g. "vhcxls4pdb"). Still excludes known keywords. */
        const isVirtualToken = (t) => {
            const h = t.toLowerCase();
            if (this.NON_HOSTNAME.has(h)) return false;
            return /^[a-z][a-z0-9\-._]{2,60}$/i.test(t);
        };

        for (const rawLine of text.split('\n')) {
            const m = rawLine.match(rowRe);
            if (!m) continue;
            /* Tokens after the Related-System column. A COMPONENTS-table row
               carries a leading component label ("ABAP SCS") before the host;
               findIndex(isHostToken) below skips past it since those words are
               not host-shaped. */
            let toks = m[1].trim().split(/\s+/).filter(Boolean);

            /* Physical host = first host-like token on the row. */
            const physIdx = toks.findIndex(isHostToken);
            if (physIdx < 0) continue;
            const physical = toks[physIdx];

            /* Virtual name = next qualifying token after the physical host.
               Prefer an Instance Name (has the _SID_nn suffix → strip it);
               otherwise take the next host-shaped token (SAPDBHOST / Logical
               Host). Only accept names that differ from the physical host. */
            let virtual = '';
            for (let i = physIdx + 1; i < toks.length; i++) {
                const t = toks[i];
                if (!isVirtualToken(t)) continue;
                const stripped = instSuffix.test(t) ? t.replace(instSuffix, '') : t;
                if (stripped.toLowerCase() === physical.toLowerCase()) continue;
                virtual = stripped;
                break;
            }
            if (!virtual) continue;

            /* First writer wins; don't clobber an app-server alias with a later
               duplicate (the same physical host can appear as both an app CI and
               an ABAP SCS row — keep the first, which is the app-server entry). */
            const key = physical.toLowerCase();
            if (!(key in aliases)) aliases[key] = virtual;
        }

        try {
            console.log('[SystemLandscape] host aliases (physical→virtual):',
                        JSON.parse(JSON.stringify(aliases)));
        } catch (_) { /* console may be unavailable */ }

        return aliases;
    },

    /* ── Parse a single server block (all lines joined as one string) ── */
    _parseHardwareLine(hostname, lineText, compMap) {
        const entry = {
            hostname:      hostname.toUpperCase(),
            manufacturer:  '-',
            model:         '-',
            cpuType:       '-',
            virtualization:'-',
            os:            '-',
            ghz:           '-',
            logicalCpus:   '-',
            cores:         '-',
            saps:          '',
            memoryMB:      '-',
            component:     compMap[hostname.toLowerCase()] || this._guessComponent(hostname),
        };

        /* ── 1. Memory ─ rightmost large integer (5-10 digits) ── */
        const allLargeNums = [...lineText.matchAll(/\b(\d{5,10})\b/g)];
        if (allLargeNums.length > 0) {
            entry.memoryMB = allLargeNums[allLargeNums.length - 1][1];
        }

        /* ── 2. CPU Counts (Logical CPUs + Cores) ── */
        /* Look for the pair just before the memory value */
        const cpuPairM = lineText.match(/\b(\d{2,4})\s+(\d{2,4})\s+\d{5,10}\b/)
                      || lineText.match(/\b(\d{2,4})\s+[-–—]\s+\d{5,10}\b/);

        if (cpuPairM) {
            entry.logicalCpus = cpuPairM[1];
            /* Second group is either cores or a dash */
            if (cpuPairM[2] && /\d/.test(cpuPairM[2])) {
                entry.cores = cpuPairM[2];
            }
            /* If only logical CPUs found, infer cores = logicalCPUs / 2 (HT) */
            if (entry.cores === '-' || entry.cores === '') {
                const n = parseInt(entry.logicalCpus);
                if (!isNaN(n) && n > 0) entry.cores = String(Math.round(n / 2));
            }
        } else {
            /* Fallback: find largest 2-4 digit number that could be a CPU count */
            const cpuM = lineText.match(/\b(224|448|192|128|112|96|64|56|48|32)\b/);
            if (cpuM) {
                entry.logicalCpus = cpuM[1];
                entry.cores = String(Math.round(parseInt(cpuM[1]) / 2));
            }
        }

        /* ── 3. Manufacturer ── */
        const mfgM = lineText.match(
            /\b(Amazon\s+EC2|Amazon\s+Web\s+Services|AWS|Google\s+Cloud|Google\s+Compute|Microsoft\s+Azure|Azure|HPE|Hewlett[- ]Packard(?:\s+Enterprise)?|Dell(?:\s+EMC)?|IBM|Cisco\s+UCS|Fujitsu|Lenovo|SuperMicro|Super\s+Micro|Oracle)\b/i
        );
        if (mfgM) {
            entry.manufacturer = mfgM[1]
                .replace(/Hewlett[- ]Packard(\s+Enterprise)?/i, 'HPE')
                .replace(/Amazon\s+Web\s+Services/i, 'Amazon EC2')
                .replace(/Google\s+Compute/i, 'Google Cloud')
                .replace(/Microsoft\s+Azure/i, 'Azure')
                .replace(/\s+/g, ' ')
                .trim();
        }

        /* ── 4. CPU Type ──
           The EWA "Host Overview" table has SEPARATE "CPU Type" and "CPU MHz"
           columns, e.g. CPU Type = "EPYC 7R13", CPU MHz = "3600".  In the
           PDF.js text the model suffix often lands on a different visual line
           than the "EPYC"/"Xeon" token (e.g. "7R13 Enterprise Server" sits
           above "... EPYC 3600 KVM ..."), so we must stitch the family token
           to its model number while NOT mistaking the 4-digit MHz value for
           the model.

           AWS instance CPUs use models like "7R13" / "7R32" (digit, letter,
           digits).  Classic AMD EPYC models are "7xxx"/"9xxx".  Intel Xeon
           Scalable models are "<4 digits><letter>" (e.g. 8176M, 8488C) or a
           plain >=5000 number. ── */
        {
            const directM = lineText.match(
                /\b(Xeon\s+Platinum\s+\d{3,4}[A-Za-z]?\+?|Xeon\s+Gold\s+\d{3,4}[A-Za-z]?\+?|Xeon\s+Silver\s+\d{3,4}[A-Za-z]?\+?|EPYC\s+\d[A-Za-z]\d{1,3}\w*|EPYC\s+[7-9]\d{3}\w*|POWER\d+[^\s]*|Itanium\s+[\w]+)\b/i
            );
            if (directM) {
                entry.cpuType = directM[1].replace(/\s+/g, ' ').trim();
            } else if (/\bXeon\b/i.test(lineText)) {
                /* "Xeon", its tier, and the model number may each be on a
                   separate visual line.  Collect them from the whole block. */
                const tierM = lineText.match(/\b(Platinum|Gold|Silver)\b/i);
                /* Model: 4 digits + a letter suffix (8176M), or a >=5000 bare
                   number — never a typical MHz value (1800–4999). */
                const modelM = lineText.match(/\b(\d{4}[A-Za-z]\+?|[5-9]\d{3})\b/);
                let cpu = 'Xeon';
                if (tierM)  cpu += ' ' + tierM[1].replace(/^\w/, c => c.toUpperCase());
                if (modelM) cpu += ' ' + modelM[1];
                entry.cpuType = cpu.trim();
            } else if (/\bEPYC\b/i.test(lineText)) {
                /* AWS EPYC instances name their model "7R13" / "7R32" (digit,
                   letter, digits) — match that first.  Only fall back to a
                   classic bare "7xxx"/"9xxx" model when there is NO Intel Xeon
                   context in the block: in the Y-flipped PDF text a neighbour
                   row's Xeon model (e.g. "8176M") can leak into this block and
                   would otherwise be mis-read as an EPYC model. */
                const awsM = lineText.match(/\b(\d[A-Za-z]\d{1,3}[A-Za-z]?)\b/);
                let model = awsM ? awsM[1] : '';
                if (!model && !/\b(Xeon|Platinum|Gold|Silver)\b/i.test(lineText)) {
                    const classicM = lineText.match(/\b([79]\d{3}[A-Za-z]?)\b/);
                    model = classicM ? classicM[1] : '';
                }
                entry.cpuType = ('EPYC' + (model ? ' ' + model : '')).trim();
            }
        }

        /* ── 5. GHz — exclude numbers that are part of CPU model names (e.g. 8480+) ── */
        {
            /* Negative lookahead: skip numbers immediately followed by a letter, +, or digit */
            const mhzPat = /\b([1-9]\d{3})(?![A-Za-z+\d])/g;
            /* Prefer first match found after the CPU type string */
            const cpuTypeStr = entry.cpuType !== '-' ? entry.cpuType : '';
            const cpuPos = cpuTypeStr ? lineText.indexOf(cpuTypeStr) : -1;
            const searchFrom = cpuPos >= 0 ? cpuPos + cpuTypeStr.length : 0;
            const afterCPU = lineText.slice(searchFrom);
            const mhzM = afterCPU.match(/\b([1-9]\d{3})(?![A-Za-z+\d])/);
            if (mhzM) {
                const n = parseInt(mhzM[1]);
                if (n >= 1000 && n <= 9000) entry.ghz = (n / 1000).toFixed(1);
            }
            if (entry.ghz === '-') {
                const allMhz = [...lineText.matchAll(mhzPat)]
                    .map(m => parseInt(m[1]))
                    .filter(n => n >= 1800 && n <= 5200 && n !== 2023 && n !== 2024);
                if (allMhz.length > 0) entry.ghz = (allMhz[0] / 1000).toFixed(1);
            }
        }

        /* ── 6. Operating System — direct match, then fragment-aware SUSE ──
           NOTE: the capture must NOT run on into the trailing numeric columns
           (logical CPUs, cores, memory) that sit on the same PDF text line, e.g.
           "… Server 15 (x86_64) 16 16 128521". Each OS pattern therefore stops at
           an architecture parenthesis or version token rather than greedily
           consuming the next 30 characters. */
        {
            const directOS = lineText.match(
                /(Red\s+Hat\s+Enterprise\s+Linux(?:\s+Server)?(?:\s+release)?(?:\s+[\d.]+)?(?:\s+\([^)]*\))?|SUSE\s+Linux\s+Enterprise\s+Server\s+\d+(?:\s+SP\d+)?(?:\s+\([^)]*\))?|SLES\s+\d+(?:\s+SP\d+)?(?:\s+\([^)]*\))?|Windows\s+Server\s+\d{4}(?:\s+R2)?(?:\s+[A-Za-z]+)?|Oracle\s+Linux(?:\s+Server)?(?:\s+[\d.]+)?|AIX\s+\d[\d.]+)/i
            );
            if (directOS) {
                entry.os = directOS[1].replace(/\s+/g, ' ').trim();
            } else if (/\bSUSE\b/i.test(lineText) || /\bSLES\b/i.test(lineText)) {
                const verM  = lineText.match(/\bServer\s+(\d+)\b/i);
                const archM = lineText.match(/\((x86_64|aarch64|ppc64le)\)/i);
                let os = /\bSLES\b/i.test(lineText) ? 'SLES' : 'SUSE Linux Enterprise Server';
                if (verM)  os += ' ' + verM[1];
                if (archM) os += ' (' + archM[1] + ')';
                entry.os = os;
            }
        }

        /* ── 7. Model — knowledge-based for HPE, generic regex for others ── */
        {
            let model = null;
            if (/\bProLiant\b/i.test(lineText)) {
                const seriesM = lineText.match(/\b(DL|ML|BL|Synergy|XL)\s*(\d{3,4}[a-zA-Z]?)\b/i);
                const genM    = lineText.match(/\bGen\s*(\d{1,2})\b/i);
                model = 'ProLiant';
                if (seriesM) model += ' ' + seriesM[1].toUpperCase() + seriesM[2];
                if (genM)    model += ' Gen' + genM[1];
            } else if (/\bSuperdome\b/i.test(lineText)) {
                const flexM = /\bFlex\b/i.test(lineText);
                const numM  = lineText.match(/\b(280|960|S)\b/);
                model = 'Superdome';
                if (flexM) model += ' Flex';
                if (numM)  model += ' ' + numM[1];
            } else {
                const modelM = lineText.match(
                    /\b([a-z]+\d+[a-z]*\.\d+xlarge|\d+xlarge|PowerEdge\s+\w+(?:\s+\w+)?|ThinkSystem\s+\w+(?:\s+\w+)?|PRIMERGY\s+\w+(?:\s+\w+)?|UCS\s+[CS]\d+[^\s,]{0,10}|Power\s+Systems\s+\w+)\b/i
                );
                if (modelM) model = modelM[1].replace(/\s+/g, ' ').trim();
            }
            if (model) entry.model = model;
        }

        /* ── 7b. Virtualization ── */
        const virtM = lineText.match(
            /\b(KVM|VMware|vSphere|Hyper\-V|HyperV|Xen|LPAR|PowerVM|Physical|Bare\s+Metal|bare-metal|AWS|Amazon\s+VPC)\b/i
        );
        if (virtM) {
            entry.virtualization = virtM[1]
                .replace(/HyperV/i, 'Hyper-V')
                .replace(/Bare\s+Metal|bare-metal/i, 'Physical')
                .replace(/Amazon\s+VPC/i, 'AWS')
                .trim();
        }

        /* ── 8. Auto-calculate SAPS ── */
        const spc   = this._getSAPSPerCore(entry.cpuType);
        const cores = parseInt(entry.cores);
        if (spc && !isNaN(cores) && cores > 0) {
            entry.saps = String(cores * spc);
        }

        /* ── 9. Validation — keep entry only if it has real data ── */
        const inCompMap = !!compMap[hostname.toLowerCase()];
        const hasHW = entry.manufacturer !== '-'
                   || entry.cpuType !== '-'
                   || entry.virtualization !== '-'
                   || (entry.memoryMB !== '-' && parseInt(entry.memoryMB) > 4096);

        if (!inCompMap && !hasHW) return null;

        return entry;
    },

    _guessComponent(hostname) {
        const h = hostname.toLowerCase();
        if (/han|hdb|hana|db\d/.test(h))       return 'DB';
        if (/scs|ascs|cs\d|csc|sapca/.test(h)) return 'SAP ABAP SCS';
        return 'SAP AS';
    },

    /* Layer classification helpers (single source of truth for layer math) */
    _isAppLayer(component) {
        return ['SAP AS', 'SAP AS / SAP ABAP SCS', 'SAP ABAP SCS'].includes(component);
    },
    _isDbLayer(component) {
        return component === 'DB';
    },

    _getSAPSPerCore(cpuType) {
        if (!cpuType || cpuType === '-') return null;
        if (this.SAPS_DB[cpuType]) return this.SAPS_DB[cpuType];
        for (const [key, val] of Object.entries(this.SAPS_DB)) {
            if (cpuType.includes(key) || key.includes(cpuType)) return val;
        }
        const numM = cpuType.match(/(\d{4}[\w\+]*)/);
        if (numM) {
            for (const [key, val] of Object.entries(this.SAPS_DB)) {
                if (key.includes(numM[1])) return val;
            }
        }
        return null;
    },

    /* ── Heuristic SAPS/core estimate when neither benchmark nor SAPS_DB match ──
       Guarantees the AI / knowledge-base (blue flag) case always proposes a value.
       Based on CPU family/generation and clock speed, grounded in observed SD
       benchmark ranges for modern x86 server CPUs (~2000–4500 SAPS/core). */
    _estimateSAPSPerCore(cpuType, ghz) {
        const cpu = (cpuType || '').toLowerCase();
        const g   = parseFloat(ghz) || 0;

        /* Base SAPS/core by family/tier */
        let base = 2600;                                   /* generic modern x86 default */
        if (cpu.includes('epyc')) {
            if (/9\d{3}/.test(cpu))      base = 4000;      /* Genoa / Bergamo */
            else if (/7r13|7\d{3}/.test(cpu)) base = 2900; /* Milan / Rome */
            else                          base = 2600;
        } else if (cpu.includes('xeon')) {
            if (cpu.includes('platinum')) {
                if (/84\d{2}|94\d{2}|88\d{2}/.test(cpu)) base = 3600; /* Sapphire/Emerald Rapids */
                else if (/83\d{2}|82\d{2}|837\d/.test(cpu)) base = 2900; /* Ice/Cascade Lake */
                else                                      base = 2500;
            } else if (cpu.includes('gold')) {
                base = 3000;
            } else {
                base = 2400;
            }
        } else if (cpu.includes('power')) {
            base = 5000;                                   /* IBM POWER9/10 */
        }

        /* Clock adjustment: nudge ±10% around a 2.6 GHz reference */
        if (g > 0) {
            const factor = Math.max(0.85, Math.min(1.15, g / 2.6));
            base = base * factor;
        }
        return Math.round(base / 50) * 50;                 /* round to nearest 50 */
    },

    _renderResults(servers, sysInfo) {
        document.getElementById('sl-results').style.display = 'block';
        /* Hide the upload card once a report has been parsed */
        const upCard = document.getElementById('sl-upload-card');
        if (upCard) upCard.style.display = 'none';

        /* Application layer = SAP AS + SAP AS / SAP ABAP SCS + SAP ABAP SCS */
        const appCnt  = servers.filter(s => this._isAppLayer(s.component)).length;
        const dbCnt   = servers.filter(s => this._isDbLayer(s.component)).length;
        const ascsCnt = servers.filter(s => s.component === 'SAP ABAP SCS' || s.component === 'SAP AS / SAP ABAP SCS').length;
        const totalMB = servers.reduce((s, r) => s + (parseInt(r.memoryMB) || 0), 0);

        document.getElementById('sl-stats').innerHTML = [
            statCard('Total Servers',    servers.length),
            statCard('App Layer',        appCnt,  ' servers'),
            statCard('Database Layer',   dbCnt,   ' servers'),
            ascsCnt ? statCard('ABAP SCS', ascsCnt, ' servers') : '',
            totalMB ? statCard('Total RAM', (totalMB / 1024 / 1024).toFixed(1), ' TB') : '',
        ].join('');

        document.getElementById('sl-table-sub').textContent =
            `${servers.length} server(s) — all cells are editable`;

        this._renderEditableTable(servers);
        this._renderCPUSummaryWithBenchmark(servers);

        document.getElementById('sl-btn-next')
            ?.addEventListener('click', () => this._onNextClick());
        document.getElementById('sl-btn-export')
            ?.addEventListener('click', () => this._downloadAnalysis());
        document.getElementById('sl-btn-add')
            ?.addEventListener('click', () => this._addTableRow());
        document.getElementById('sl-btn-copy-excel')
            ?.addEventListener('click', e => this._copyEditableTableForExcel(e.currentTarget));
        document.getElementById('sl-btn-back')
            ?.addEventListener('click', () => this._showMainScreen());
        document.getElementById('sl-btn-review-next')
            ?.addEventListener('click', () => { this._saveToApp(); this.app.nextModule(); });
        /* Top navigation bar on the review screen — same actions as the bottom */
        document.getElementById('sl-btn-back-top')
            ?.addEventListener('click', () => this._showMainScreen());
        document.getElementById('sl-btn-review-next-top')
            ?.addEventListener('click', () => { this._saveToApp(); this.app.nextModule(); });
    },

    /* ── Editable table ── */
    _renderEditableTable(servers) {
        const tbody = document.getElementById('sl-hw-tbody');
        tbody.innerHTML = servers.map((s, i) => this._rowHTML(s, i)).join('');

        tbody.querySelectorAll('.sl-del-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                this.servers.splice(parseInt(btn.dataset.i), 1);
                this._renderEditableTable(this.servers);
                this._renderCPUSummaryWithBenchmark(this.servers);
                this._saveToApp();
            });
        });
    },

    _rowHTML(s, i) {
        const compOpts = ['SAP AS', 'SAP AS / SAP ABAP SCS', 'SAP ABAP SCS', 'DB']
            .map(o => `<option${s.component === o ? ' selected' : ''}>${o}</option>`)
            .join('');

        const cell = (val, field, w) =>
            `<td style="padding:4px 6px"><input class="editable-cell"
                style="${w === 0 ? '' : `min-width:${w || 70}px`}"
                value="${(val || '').toString().replace(/"/g, '&quot;').replace(/</g, '&lt;')}"
                data-i="${i}" data-f="${field}"
                oninput="SystemLandscapeModule._onInput(${i},'${field}',this.value)"
                onchange="SystemLandscapeModule._onCellCommit(${i},'${field}',this.value)"></td>`;

        return `<tr>
            <td style="padding:4px 0;text-align:center">
                <button class="sl-del-btn" data-i="${i}"
                        style="background:none;border:none;cursor:pointer;color:var(--error);
                               font-size:18px;line-height:1;padding:2px 8px"
                        title="Remove row">×</button>
            </td>
            ${cell(s.hostname,      'hostname',       0)}
            ${cell(s.manufacturer,  'manufacturer',  60)}
            ${cell(s.model,         'model',         110)}
            ${cell(s.cpuType,       'cpuType',       150)}
            ${cell(s.virtualization,'virtualization', 70)}
            ${cell(s.os,            'os',            190)}
            ${cell(s.ghz,           'ghz',            40)}
            ${cell(s.logicalCpus,   'logicalCpus',    50)}
            ${cell(s.cores,         'cores',           50)}
            ${cell(s.saps,          'saps',            80)}
            ${cell(s.memoryMB,      'memoryMB',        90)}
            <td style="padding:4px 6px">
                <select class="editable-select"
                        onchange="SystemLandscapeModule._onInput(${i},'component',this.value)">
                    ${compOpts}
                </select>
            </td>
            <td style="padding:4px 0;text-align:center">
                <button class="sl-del-btn" data-i="${i}"
                        style="background:none;border:none;cursor:pointer;color:var(--error);
                               font-size:18px;line-height:1;padding:2px 8px"
                        title="Remove row">×</button>
            </td>
        </tr>`;
    },

    _onInput(i, field, value) {
        if (!this.servers[i]) return;
        this.servers[i][field] = value;
        /* Keep model state live on every keystroke, but DON'T re-render the
           Processor Summary here — doing it per-keystroke would spawn a line
           for every partial CPU string ("Xe", "Xeo", …). The summary is
           refreshed once editing is committed (blur / Enter) in
           _onCellCommit. */
        this._saveToApp();
    },

    /* Fires on blur / Enter of a Server Inventory cell. When a CPU-affecting
       field (CPU Type or Cores) has been committed, (re)build the Processor
       Summary so a NEWLY TYPED CPU type appears there as its own line, with
       SAPS/Core resolved by the same criteria used originally (SAP SD
       Benchmark → built-in SAPS_DB → heuristic estimate) and the field left
       editable so the user can still override it. */
    _onCellCommit(i, field, value) {
        if (!this.servers[i]) return;
        this.servers[i][field] = value;

        if (field === 'cpuType' || field === 'cores' || field === 'ghz') {
            const s   = this.servers[i];
            const c   = parseInt(s.cores);
            /* Resolve SAPS/Core with the same precedence used elsewhere (manual
               override for this cpu@ghz → benchmark/SAPS_DB lookup). When Cores /
               GHz / CPU type changes, recompute SAPS (TBV) = SAPS/Core × Cores —
               UNLESS the user previously typed SAPS manually for this server, in
               which case that server's SAPS is the source of truth and must be
               preserved. */
            const spc = this._resolveSPC(s);
            const hasManualSaps = !!s._sapsManual;
            if (spc && !isNaN(c) && c > 0 && !hasManualSaps) {
                s.saps = String(c * spc);
                const inp = document.querySelector(`input[data-i="${i}"][data-f="saps"]`);
                if (inp) inp.value = s.saps;
            }
            this._renderCPUSummaryWithBenchmark(this.servers);
            this._saveToApp();
        }

        if (field === 'saps') {
            /* When SAPS (TBV) is manually edited, this server becomes the source
               of truth for SAPS. Set s._sapsManual = true so ALL code paths that
               auto-calculate SAPS skip this server. The flag travels with the
               server object through sorting, deletion, and session save/restore,
               so it never becomes stale the way an index-based map would.
               Back-calculate SAPS/Core and record as a sticky override for this
               processor so the Processor Summary reflects the user's intent. */
            const s      = this.servers[i];
            /* s.saps already holds the exact raw string (set at _onCellCommit top).
               Use parseFloat for the arithmetic so values like "12345.6" work,
               but NEVER write back a rounded/converted value to s.saps. */
            const saps   = parseFloat(s.saps);
            const cores  = parseInt(s.cores);
            if (saps > 0) {
                s._sapsManual = true;
            } else {
                /* Clearing the value releases the manual override. */
                delete s._sapsManual;
            }
            if (!isNaN(saps) && !isNaN(cores) && cores > 0 && saps > 0) {
                const newSpc = Math.round(saps / cores);
                if (!this._spcOverride) this._spcOverride = {};
                this._spcOverride[this._spcKey(s.cpuType, s.ghz)] = newSpc;
                /* Update the SAPS/Core input in the Processor Summary if visible */
                const safeKey = ((s.cpuType || '') + (s.ghz || '')).replace(/[^a-z0-9]/gi, '_');
                const spcInp  = document.getElementById(`spc-${safeKey}`);
                if (spcInp) spcInp.value = newSpc;
            }
            this._renderCPUSummaryWithBenchmark(this.servers);
            /* Refresh the review screen if already open */
            if (document.getElementById('sl-view-review')?.style.display !== 'none'
                && document.getElementById('sl-rv-inventory-tbody')) {
                this._renderReviewScreen();
            }
            this._saveToApp();
        }
    },

    /* Resolve the SAPS/Core for a host: manual override for its exact (cpuType,
       ghz) first, then the built-in SAPS_DB / benchmark cache. Returns 0 when
       unknown (caller then leaves SAPS as-is / to be resolved by the summary). */
    _resolveSPC(s) {
        const ov = this._spcOverride && this._spcOverride[this._spcKey(s.cpuType, s.ghz)];
        if (ov > 0) return ov;
        return this._getSAPSPerCore(s.cpuType || '') || 0;
    },

    _addTableRow() {
        this.servers.push({
            hostname: '', manufacturer: '', model: '', cpuType: '',
            virtualization: '-', os: '', ghz: '', logicalCpus: '', cores: '',
            saps: '', memoryMB: '', component: 'SAP AS',
        });
        this._renderEditableTable(this.servers);
        this._renderCPUSummaryWithBenchmark(this.servers);
        this._saveToApp();
    },

    /* ══════════════════════════════════════════════════════
       BENCHMARK LOOKUP — SAP SD Benchmark CSV
       Columns (semicolon-separated):
         0  Certification Number
         6  CPU Architecture   ← CPU type description
         7  CPU Speed          ← e.g. "2.4 GHz"
         9  Cores
         21 saps               ← total SAPS for that certification run
       ══════════════════════════════════════════════════════ */

    /* Load and parse the benchmark CSV (called once on first render).
       Stores parsed rows in this._benchmarkRows.
       If load fails, falls back silently to SAPS_DB table.             */
    async _loadBenchmarkCSV() {
        if (this._benchmarkRows) return;   /* already loaded */
        this._benchmarkRows = null;        /* mark as attempted */

        /* Try: (1) same-origin relative path (works when served from project root),
                (2) SAP website (may be blocked by CORS/auth) */
        const urls = [
            'additional-directories/bebchmark/export-sd.csv',
            '../additional-directories/bebchmark/export-sd.csv',
            'https://www.sap.com/dmc/exp/2018-benchmark-directory/assets/export-sd.csv',
        ];

        for (const url of urls) {
            try {
                const resp = await fetch(url, { cache: 'default' });
                if (!resp.ok) continue;
                const text = await resp.text();
                this._parseBenchmarkCSV(text);
                if (this._benchmarkRows && this._benchmarkRows.length > 0) return;
            } catch (_) { /* try next */ }
        }
        /* All fetches failed: _benchmarkRows stays null, UI offers manual upload */
    },

    /* Parse the semicolon-delimited CSV and populate this._benchmarkRows */
    _parseBenchmarkCSV(csvText) {
        const lines  = csvText.split('\n').map(l => l.trim()).filter(Boolean);
        if (!lines.length) return;

        /* Detect delimiter: SAP export uses ';' but a user-adjusted file may
           use ',' — pick whichever yields more header columns. */
        const delim = (lines[0].split(';').length >= lines[0].split(',').length) ? ';' : ',';
        const header = lines[0].split(delim).map(h => h.trim().toLowerCase());

        /* Flexible column detection (tolerant of extra / renamed columns) */
        const find = (...res) => {
            for (const re of res) { const i = header.findIndex(h => re.test(h)); if (i >= 0) return i; }
            return -1;
        };
        const idxCpu   = find(/cpu.*arch|processor|cpu.*type|^cpu$/, /cpu/);
        const idxSpeed = find(/cpu.*speed|clock|ghz|mhz|freq/);
        const idxCores = find(/^cores$|number.*cores|core\s*count|total.*cores/, /core/);
        const idxSaps  = find(/^saps$|\bsaps\b/);
        const idxCert  = find(/cert.*(number|no|#)|certification/);
        const idxPdf   = find(/pdf|link|url|certificate/);

        if (idxCpu < 0 || idxCores < 0 || idxSaps < 0) return; /* unexpected format */

        const rows = [];
        for (let i = 1; i < lines.length; i++) {
            const cols = lines[i].split(delim);
            const cpuRaw   = (cols[idxCpu]   || '').trim();
            const speedRaw = (cols[idxSpeed] || '').trim();
            const coresRaw = (cols[idxCores] || '').replace(/,/g, '').trim();
            const sapsRaw  = (cols[idxSaps]  || '').replace(/,/g, '').trim();
            const certNum  = idxCert >= 0 ? (cols[idxCert] || '').trim() : '';
            const pdfLink  = idxPdf  >= 0 ? (cols[idxPdf]  || '').trim() : '';

            const cores = parseInt(coresRaw);
            const saps  = parseInt(sapsRaw);
            if (!cpuRaw || isNaN(cores) || cores <= 0 || isNaN(saps) || saps <= 0) continue;

            /* Parse GHz from speed string — handle "2.4 GHz", "2400 MHz", "2.4" */
            let ghz = 0;
            const speedNum = parseFloat(speedRaw);
            if (!isNaN(speedNum)) {
                ghz = /mhz/i.test(speedRaw) ? speedNum / 1000 : speedNum;
            }

            /* Normalise CPU name: strip trailing "processor", "Processor", punctuation */
            const cpuNorm = cpuRaw.replace(/\s*processor\s*$/i, '').replace(/\s+/g, ' ').trim();

            rows.push({ cpuRaw, cpuNorm, ghz, cores, saps, sapsPerCore: Math.round(saps / cores), certNum, pdfLink });
        }
        this._benchmarkRows = rows;
    },

    /* Given a CPU type string and GHz value from the hardware table,
       find the best matching benchmark entry and return:
       { sapsPerCore, confidence: 'exact'|'close'|'estimated', ref } */
    _lookupBenchmark(cpuType, ghz) {
        const rows = this._benchmarkRows;
        if (!rows || rows.length === 0) return null;

        /* Normalise the query CPU (remove "processor", extra spaces) */
        const queryNorm = (cpuType || '').replace(/\s*processor\s*$/i, '').replace(/\s+/g, ' ').trim().toLowerCase();
        const queryGhz  = parseFloat(ghz) || 0;

        /* ── Pass 1: exact CPU name + GHz within ±0.05 GHz ── */
        const exactMatches = rows.filter(r => {
            const diff = Math.abs(r.ghz - queryGhz);
            return r.cpuNorm.toLowerCase() === queryNorm && diff <= 0.05;
        });
        if (exactMatches.length > 0) {
            /* Average saps/core across all exact matches for robustness */
            const avg = Math.round(exactMatches.reduce((s, r) => s + r.sapsPerCore, 0) / exactMatches.length);
            const best = exactMatches.reduce((a, b) => b.saps > a.saps ? b : a);
            return { sapsPerCore: avg, confidence: 'exact', ref: best.cpuNorm, certNum: best.certNum, pdfLink: best.pdfLink, ghz: best.ghz };
        }

        /* ── Pass 2: exact CPU name, any GHz ── */
        const nameMatches = rows.filter(r => r.cpuNorm.toLowerCase() === queryNorm);
        if (nameMatches.length > 0) {
            /* Pick the entry whose GHz is closest to query */
            nameMatches.sort((a, b) => Math.abs(a.ghz - queryGhz) - Math.abs(b.ghz - queryGhz));
            const best = nameMatches[0];
            return { sapsPerCore: best.sapsPerCore, confidence: 'close', ref: `${best.cpuNorm} @ ${best.ghz} GHz`, certNum: best.certNum, pdfLink: best.pdfLink, ghz: best.ghz };
        }

        /* ── Pass 3: partial CPU name match (model number substring) ── */
        /* Extract model number token from query: e.g. "8490H", "9654", "EPYC 9654" */
        const modelToken = queryNorm.match(/\b([a-z]?\d{4,}[a-z\+]*)\b/i);
        if (modelToken) {
            const token = modelToken[1].toLowerCase();
            const partialMatches = rows.filter(r => r.cpuNorm.toLowerCase().includes(token));
            if (partialMatches.length > 0) {
                partialMatches.sort((a, b) => Math.abs(a.ghz - queryGhz) - Math.abs(b.ghz - queryGhz));
                const best = partialMatches[0];
                return { sapsPerCore: best.sapsPerCore, confidence: 'close', ref: `${best.cpuNorm} @ ${best.ghz} GHz`, certNum: best.certNum, pdfLink: best.pdfLink, ghz: best.ghz };
            }
        }

        return null; /* no match — caller will fall back to SAPS_DB / AI estimate */
    },

    /* ── CPU / SAPS Summary with Benchmark lookup ── */
    async _renderCPUSummaryWithBenchmark(servers) {
        /* Ensure benchmark data is loaded (no-op if already done) */
        if (!this._benchmarkRows) {
            await this._loadBenchmarkCSV();
        }
        this._renderCPUSummary(servers);
    },

    /* ── CPU / SAPS Summary (sync render — called after benchmark data ready) ── */
    _renderCPUSummary(servers) {
        /* Group servers by CPU type */
        const groups = {};
        for (const s of servers) {
            const cpu = (s.cpuType && s.cpuType !== '-') ? s.cpuType : '(not identified)';
            const ghz = s.ghz || '';
            const key = cpu + '||' + ghz;
            if (!groups[key]) groups[key] = { cpu, ghz, count: 0, cores: 0, saps: 0 };
            groups[key].count++;
            groups[key].cores += parseInt(s.cores)  || 0;
            groups[key].saps  += parseInt(s.saps)   || 0;
        }

        const rows = Object.values(groups).map(g => {
            const { cpu, ghz, count, cores, saps } = g;
            const safeKey = (cpu + ghz).replace(/[^a-z0-9]/gi, '_');

            /* ── Determine SAPS/core source and confidence badge ── */
            let spc        = '';
            let confidence = 'ai';      /* default */
            let badgeHtml  = '';
            let refHtml    = '';

            if (cpu !== '(not identified)') {
                /* A manual SAPS/Core override (entered by the user in this
                   panel) always wins — otherwise a re-render would clobber it
                   with the benchmark value and the user's edit would be lost.
                   Overrides are keyed by CPU type AND GHz, because the same
                   CPU model at a different clock speed can deliver a different
                   SAPS/Core, so each (cpu, ghz) row is independent. */
                const ovKey    = this._spcKey(cpu, ghz);
                const override = this._spcOverride && this._spcOverride[ovKey];
                if (override > 0) {
                    spc        = override;
                    confidence = 'manual';
                    badgeHtml  = `<span class="saps-badge saps-badge-manual" title="Manually entered — overrides the benchmark/estimate for this CPU @ this GHz">&#9679; Manual override</span>`;
                } else {
                /* Try SAP SD Benchmark first */
                const bmResult = this._lookupBenchmark(cpu, ghz);
                if (bmResult) {
                    spc        = bmResult.sapsPerCore;
                    confidence = bmResult.confidence;   /* 'exact' or 'close' */
                    const certRef = bmResult.pdfLink
                        ? `<a href="${bmResult.pdfLink}" target="_blank" style="color:inherit">Cert ${bmResult.certNum}</a>`
                        : (bmResult.certNum || '');
                    if (confidence === 'exact') {
                        badgeHtml = `<span class="saps-badge saps-badge-exact" title="Exact match found in SAP SD Benchmark">&#9679; Exact SAP SD Benchmark match</span>`;
                        refHtml   = certRef ? `<div class="saps-ref">Reference: ${certRef}</div>` : '';
                    } else {
                        badgeHtml = `<span class="saps-badge saps-badge-close" title="Closest match in SAP SD Benchmark — ${bmResult.ref}">&#9679; Estimated from SAP SD Benchmark (closest: ${bmResult.ref})</span>`;
                        refHtml   = certRef ? `<div class="saps-ref">Reference: ${certRef}</div>` : '';
                    }
                    /* Cache into SAPS_DB so table cells + _hostVariables pick
                       it up. Refresh the cache when a benchmark file was just
                       (re)loaded so a newly uploaded file's values win. */
                    if (!this.SAPS_DB[cpu] || this._benchmarkJustLoaded) this.SAPS_DB[cpu] = spc;
                } else {
                    /* Fall back to built-in SAPS_DB, then heuristic estimate */
                    spc = this._getSAPSPerCore(cpu);
                    if (!spc) spc = this._estimateSAPSPerCore(cpu, ghz);
                    confidence = 'ai';
                    badgeHtml = `<span class="saps-badge saps-badge-ai" title="No SAP SD Benchmark entry found — value estimated from SAP benchmark knowledge base">&#9679; AI / Knowledge-base estimate</span>`;
                    /* Sync into SAPS_DB so table cells auto-populate */
                    if (spc && !this.SAPS_DB[cpu]) this.SAPS_DB[cpu] = spc;
                }
                } /* end benchmark/estimate branch (override already handled) */

                /* Auto-populate table SAPS cells for servers in THIS group
                   (same cpu + ghz).
                     • On a manual override, overwrite ALL matching cells.
                     • When a SAP SD Benchmark file was just (re)loaded, also
                       overwrite ALL matching cells so the table reflects the
                       newly resolved benchmark SAPS/Core (except hosts whose
                       CPU is manually overridden — handled by that branch).
                     • Otherwise only fill empty cells, so we never clobber a
                       value the user typed directly into the table. */
                if (spc) {
                    const forceAll = override > 0 || this._benchmarkJustLoaded === true;
                    servers.forEach((s, i) => {
                        const sCpu = (s.cpuType && s.cpuType !== '-') ? s.cpuType : '(not identified)';
                        const sGhz = s.ghz || '';
                        /* Never overwrite a server whose SAPS was manually entered,
                           regardless of forceAll — manual SAPS is always the source of truth. */
                        if (s._sapsManual) return;
                        if (sCpu === cpu && sGhz === ghz && (forceAll || !s.saps || s.saps === '')) {
                            const c = parseInt(s.cores);
                            if (!isNaN(c) && c > 0) {
                                s.saps = String(c * spc);
                                const inp = document.querySelector(`input[data-i="${i}"][data-f="saps"]`);
                                if (inp) inp.value = s.saps;
                            }
                        }
                    });
                }
            }

            /* Current total SAPS across THIS group only — must match BOTH
               cpu AND ghz, since the same CPU at a different clock speed is a
               separate row with its own SAPS/Core. */
            const totalSaps = servers
                .filter(s => {
                    const sCpu = (s.cpuType && s.cpuType !== '-') ? s.cpuType : '(not identified)';
                    return sCpu === cpu && (s.ghz || '') === ghz;
                })
                .reduce((acc, s) => acc + (parseInt(s.saps) || 0), 0);

            return `
            <div class="cpu-summary-row">
                <div class="cpu-type-info">
                    <div class="cpu-type-label">${cpu}${ghz && ghz !== '-' ? ` <span style="font-weight:400;font-size:12px;color:var(--text-secondary)">@ ${ghz} GHz</span>` : ''}</div>
                    <div class="cpu-type-sub">${count} server(s) · ${cores.toLocaleString()} physical cores total</div>
                    ${badgeHtml}
                    ${refHtml}
                </div>
                <div class="cpu-saps-field">
                    <label>SAPS / Core</label>
                    <input type="number" class="saps-core-input" id="spc-${safeKey}"
                           value="${spc}" placeholder="e.g. 3581"
                           onchange="SystemLandscapeModule._onSPCChange('${cpu.replace(/'/g,"\\'")}','${String(ghz).replace(/'/g,"\\'")}',this.value)">
                </div>
            </div>`;
        }).join('');

        /* Benchmark upload affordance — ALWAYS keep the upload button visible
           so the user can load a different file at any time; when a file is
           loaded, show a green confirmation above the button. */
        const bmLoadedNote = this._benchmarkRows
            ? `<span style="color:var(--success,#16a34a);font-size:11px;font-weight:600">&#10003; SAP SD Benchmark data loaded (${this._benchmarkRows.length} certifications)</span>`
            : `<span style="color:var(--text-muted);font-size:11px">SAP SD Benchmark could not be fetched automatically.</span>`;
        const bmStatus =
            `<div style="display:flex;flex-direction:column;gap:6px;align-items:flex-end">
                 ${bmLoadedNote}
                 <label class="btn" style="cursor:pointer;white-space:nowrap">
                   <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24"
                        fill="none" stroke="currentColor" stroke-width="2" style="margin-right:4px">
                     <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                     <polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
                   </svg>
                   ${this._benchmarkRows ? 'Upload a different SAP SD Benchmark' : 'Upload SAP SD Benchmark (CSV / XLSX / XLS)'}
                   <input type="file" accept=".csv,.xlsx,.xls" style="display:none"
                          onchange="SystemLandscapeModule._onBenchmarkUpload(this.files[0])">
                 </label>
               </div>`;

        document.getElementById('sl-cpu-summary').innerHTML = `
            <div class="card">
                <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px;gap:16px;flex-wrap:wrap">
                    <div style="font-size:12px;color:var(--text-secondary)">
                        SAPS/Core values are looked up from the SAP SD Benchmark (green = exact match,
                        yellow = closest match, blue = AI knowledge-base estimate).
                        You can override any value — the table updates instantly.
                    </div>
                    <div>${bmStatus}</div>
                </div>
                ${rows || '<p style="color:var(--text-muted);font-size:13px">No data yet.</p>'}
            </div>`;
    },

    /* Manual benchmark upload handler — accepts CSV, XLSX, or XLS and is
       tolerant of extra/re-ordered columns (the user may have adjusted the
       file).  Uses the shared parseUploadedFile (PapaParse / SheetJS) so the
       same column-detection path serves every format. */
    async _onBenchmarkUpload(file) {
        if (!file) return;
        const ext = (file.name.split('.').pop() || '').toLowerCase();

        if (['xlsx', 'xls'].includes(ext)) {
            parseUploadedFile(file, rows => {
                this._benchmarkRowsFromObjects(rows);
                this._applyBenchmarkRefresh();
            });
            return;
        }

        /* CSV / anything else: read as text.  Try the shared object parser
           first (handles comma OR semicolon via PapaParse), then fall back to
           the legacy semicolon splitter for the canonical SAP export. */
        if (ext === 'csv') {
            parseUploadedFile(file, rows => {
                this._benchmarkRowsFromObjects(rows);
                if (!this._benchmarkRows || !this._benchmarkRows.length) {
                    file.text().then(t => { this._parseBenchmarkCSV(t); this._applyBenchmarkRefresh(); });
                } else {
                    this._applyBenchmarkRefresh();
                }
            });
            return;
        }

        const text = await file.text();
        this._parseBenchmarkCSV(text);
        this._applyBenchmarkRefresh();
    },

    /* After a SAP SD Benchmark file is (re)loaded, push the newly resolved
       SAPS/Core values through to the Server Inventory table — recomputing
       SAPS (TBV) = SAPS/Core × Cores for every impacted host — and refresh the
       Processor Summary, the editable table and (if open) the review screen.
       Manually-overridden processors are left untouched. */
    _applyBenchmarkRefresh() {
        this._benchmarkJustLoaded = true;
        /* Re-render the summary first: its auto-populate step rewrites
           this.servers[].saps for every benchmark-resolved host while the
           flag is set. */
        this._renderCPUSummary(this.servers);
        this._benchmarkJustLoaded = false;

        /* Reflect the recomputed SAPS values in the editable table inputs. */
        this._renderEditableTable(this.servers);

        /* Keep the review screen in sync if the user is already there. */
        if (document.getElementById('sl-view-review')?.style.display !== 'none'
            && document.getElementById('sl-rv-inventory-tbody')) {
            this._renderReviewScreen();
        }

        this._saveToApp();
    },

    /* Build this._benchmarkRows from an array of header-keyed row objects
       (as produced by PapaParse / SheetJS).  Columns are matched by flexible
       regex so added columns or renamed headers do not break parsing. */
    _benchmarkRowsFromObjects(objRows) {
        if (!Array.isArray(objRows) || !objRows.length) return;
        const headers = Object.keys(objRows[0]);

        const cpuH   = detectCol(headers, /cpu.*arch|processor|cpu.*type|^cpu$/i) || detectCol(headers, /cpu/i);
        const spdH   = detectCol(headers, /cpu.*speed|clock|ghz|mhz|freq/i);
        const coreH  = detectCol(headers, /^cores$|number.*cores|core\s*count|total.*cores/i) || detectCol(headers, /core/i);
        const sapsH  = detectCol(headers, /^saps$|\bsaps\b/i);
        const certH  = detectCol(headers, /cert.*(number|no|#)|certification/i);
        const pdfH   = detectCol(headers, /pdf|link|url|certificate/i);

        if (!cpuH || !coreH || !sapsH) return;   /* essential columns missing */

        const rows = [];
        for (const r of objRows) {
            const cpuRaw   = (r[cpuH]  ?? '').toString().trim();
            const speedRaw = spdH ? (r[spdH] ?? '').toString().trim() : '';
            const cores    = parseInt((r[coreH] ?? '').toString().replace(/,/g, '').trim());
            const saps     = parseInt((r[sapsH] ?? '').toString().replace(/,/g, '').trim());
            const certNum  = certH ? (r[certH] ?? '').toString().trim() : '';
            const pdfLink  = pdfH  ? (r[pdfH]  ?? '').toString().trim() : '';
            if (!cpuRaw || isNaN(cores) || cores <= 0 || isNaN(saps) || saps <= 0) continue;

            let ghz = 0;
            const speedNum = parseFloat(speedRaw);
            if (!isNaN(speedNum)) ghz = /mhz/i.test(speedRaw) ? speedNum / 1000 : speedNum;

            const cpuNorm = cpuRaw.replace(/\s*processor\s*$/i, '').replace(/\s+/g, ' ').trim();
            rows.push({ cpuRaw, cpuNorm, ghz, cores, saps, sapsPerCore: Math.round(saps / cores), certNum, pdfLink });
        }
        if (rows.length) this._benchmarkRows = rows;
    },

    /* Composite key for a per-processor SAPS/Core override: the same CPU
       model at different clock speeds (GHz) is treated independently. */
    _spcKey(cpuType, ghz) {
        return `${cpuType || ''}||${ghz || ''}`;
    },

    _onSPCChange(cpuType, ghz, newValue) {
        const spc = parseInt(newValue);
        if (isNaN(spc) || spc <= 0) return;

        /* Record as a sticky override keyed by CPU type + GHz, so ONLY the
           edited processor row changes — other rows (including the same CPU
           model at a different GHz) keep their own values and must be updated
           manually. */
        if (!this._spcOverride) this._spcOverride = {};
        this._spcOverride[this._spcKey(cpuType, ghz)] = spc;

        /* Recalculate SAPS (TBV) = SAPS/Core × Cores for every host that uses
           this EXACT processor (same cpuType AND ghz). Comparison is trimmed so
           incidental whitespace differences between the processor-summary group
           key and the stored host value don't block the update. */
        const norm = v => String(v == null ? '' : v).trim();
        const wantCpu = norm(cpuType), wantGhz = norm(ghz);
        let changed = 0;
        this.servers.forEach((s, idx) => {
            if (norm(s.cpuType) === wantCpu && norm(s.ghz) === wantGhz) {
                /* Skip servers whose SAPS was manually typed — their SAPS is the
                   source of truth and must not be overwritten from Processor Summary. */
                if (s._sapsManual) return;
                const c = parseInt(s.cores);
                if (!isNaN(c) && c > 0) { s.saps = String(c * spc); changed++; }
            }
        });

        /* Re-render the Server Inventory table so the SAPS (TBV) column reflects
           the new values immediately (more robust than patching inputs by hand,
           and keeps the DOM in sync after any row sort). */
        if (changed) this._renderEditableTable(this.servers);

        this._renderCPUSummaryWithBenchmark(this.servers);

        /* If the user is already on the review screen, refresh its recap table
           so the SAPS (TBV) column there stays in sync too. */
        if (document.getElementById('sl-view-review')?.style.display !== 'none'
            && document.getElementById('sl-rv-inventory-tbody')) {
            this._renderReviewScreen();
        }

        this._saveToApp();
    },

    /* ── System Info ── */
    _renderSystemInfo(info) {
        const fields = [
            { key: 'sid',       label: 'SAP System ID (SID)',      placeholder: 'e.g. EPC'              },
            { key: 'product',   label: 'SAP Product & Version',    placeholder: 'e.g. SAP S/4HANA 2021' },
            { key: 'dbSystem',  label: 'Database System',          placeholder: 'e.g. SAP HANA Database' },
            { key: 'dbVersion', label: 'Database Version',         placeholder: 'e.g. 2.00.057.00'       },
        ];

        document.getElementById('sl-sysinfo').innerHTML = fields.map(f => `
            <div class="sysinfo-field">
                <label class="sysinfo-label">${f.label}</label>
                <input class="sysinfo-input" type="text"
                       value="${(info[f.key] || '').replace(/"/g, '&quot;')}"
                       placeholder="${f.placeholder}"
                       onchange="SystemLandscapeModule.systemInfo['${f.key}']=this.value;SystemLandscapeModule._saveToApp()">
            </div>`).join('');
    },

    /* ── Charts ── */
    _renderCharts(servers) {
        ['sl-chart-mem','sl-chart-saps','sl-chart-cores','sl-chart-dist'].forEach(id => {
            const c = document.getElementById(id);
            if (c) { const ch = Chart.getChart(c); if (ch) ch.destroy(); }
        });

        const labels = servers.map(s => s.hostname || '?');
        const memGB  = servers.map(s => +((parseInt(s.memoryMB) || 0) / 1024).toFixed(1));
        const cores  = servers.map(s => parseInt(s.cores) || 0);

        this._barChart('sl-chart-mem',   labels, memGB,  'Memory (GB)',      CHART_COLORS[0]);
        this._barChart('sl-chart-cores', labels, cores,  'Physical Cores',   CHART_COLORS[2]);

        const compSAPS = {};
        for (const s of servers) {
            const t = s.component || 'Other';
            compSAPS[t] = (compSAPS[t] || 0) + (parseInt(s.saps) || 0);
        }
        this._doughnutChart('sl-chart-saps', Object.keys(compSAPS), Object.values(compSAPS));

        const typeCounts = {};
        for (const s of servers) {
            typeCounts[s.component || 'Other'] = (typeCounts[s.component || 'Other'] || 0) + 1;
        }
        this._doughnutChart('sl-chart-dist', Object.keys(typeCounts), Object.values(typeCounts));
    },

    _barChart(id, labels, values, label, color) {
        const canvas = document.getElementById(id);
        const ctx = canvas?.getContext('2d');
        if (!ctx) return;
        if (this._slCharts[id]) { try { this._slCharts[id].destroy(); } catch(_) {} }
        this._slCharts[id] = new Chart(ctx, {
            type: 'bar',
            data: {
                labels,
                datasets: [{ label, data: values, backgroundColor: color + 'CC',
                             borderRadius: 4, borderSkipped: false }],
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { beginAtZero: true, grid: { color: '#F1F5F9' } },
                    x: { grid: { display: false },
                         ticks: { maxRotation: 35, font: { size: 10 } } },
                },
            },
        });
        this._addSlFontRow(canvas, id, false);
    },

    _doughnutChart(id, labels, values) {
        const canvas = document.getElementById(id);
        const ctx = canvas?.getContext('2d');
        if (!ctx) return;
        if (this._slCharts[id]) { try { this._slCharts[id].destroy(); } catch(_) {} }
        this._slCharts[id] = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels,
                datasets: [{ data: values, backgroundColor: CHART_COLORS,
                             borderWidth: 2, borderColor: '#fff' }],
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: {
                    legend: { position: 'bottom', labels: { font: { size: 11 }, boxWidth: 12 } },
                },
            },
        });
        this._addSlFontRow(canvas, id, true);
    },

    _addSlFontRow(canvas, uid, _isPie) {
        if (!canvas) return;
        const wrap = canvas.parentElement;
        const existing = wrap.nextElementSibling;
        if (existing && existing.classList.contains('chart-font-row')) existing.remove();
        wrap.insertAdjacentHTML('afterend', chartFontRowHtml(uid));
        const card = wrap.closest('.chart-card') || wrap.parentElement;
        const getAllRows = () => {
            const rows = [];
            const section = card.closest('.charts-grid') || document.getElementById('sl-charts-section');
            if (section) section.querySelectorAll('[id$="-font-row"]').forEach(r => rows.push(r));
            return rows;
        };
        wireChartFontRow(card, () => Object.values(this._slCharts), getAllRows, null);
    },

    /* ── Next button: validate then show review screen ── */
    _onNextClick() {
        const required = ['hostname', 'cores', 'saps', 'memoryMB', 'component'];
        const incomplete = [];

        this.servers.forEach((s, i) => {
            const missing = required.filter(f => {
                const v = (s[f] || '').toString().trim();
                return !v || v === '-';
            });
            if (missing.length > 0) {
                incomplete.push({ row: i + 1, hostname: s.hostname || `Row ${i + 1}`, missing });
            }
        });

        const errEl  = document.getElementById('sl-validation-error');
        const errMsg = document.getElementById('sl-validation-msg');

        if (incomplete.length > 0) {
            const lines = incomplete.map(e =>
                `<strong>${e.hostname}</strong>: missing ${e.missing.map(f => ({
                    hostname: 'Hostname', cores: 'Cores', saps: 'SAPS (TBV)',
                    memoryMB: 'Memory (MB)', component: 'Component Type',
                }[f])).join(', ')}`
            ).join('<br>');

            errMsg.innerHTML =
                `<strong>Cannot proceed — ${incomplete.length} row(s) are incomplete:</strong><br>${lines}<br><br>
                 Please fill in all required fields and try again.`;
            errEl.style.display = 'flex';
            errEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            return;
        }

        /* Validate SAPS/Core for every distinct CPU group */
        const missingSpc = this._cpuGroups()
            .filter(g => !(parseFloat(g.sapsPerCore) > 0))
            .map(g => `${g.cpu}${g.ghz ? ` @ ${g.ghz} GHz` : ''}`);

        if (missingSpc.length > 0) {
            errMsg.innerHTML =
                `<strong>Cannot proceed — SAPS/Core is missing for:</strong><br>` +
                missingSpc.map(c => `<strong>${c}</strong>`).join('<br>') +
                `<br><br>Please enter a SAPS/Core value for each processor in the
                 Processor Summary &amp; SAPS Estimation section.`;
            errEl.style.display = 'flex';
            errEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            return;
        }

        errEl.style.display = 'none';
        this._saveToApp();
        this._showReviewScreen();
    },

    /* ── Distinct CPU groups with resolved SAPS/Core + confidence ──
       Single source for both the review table and SAPS/Core validation. */
    _cpuGroups() {
        const groups = {};
        for (const s of this.servers) {
            const cpu = (s.cpuType && s.cpuType !== '-') ? s.cpuType : '(not identified)';
            const ghz = (s.ghz && s.ghz !== '-') ? s.ghz : '';
            const key = cpu + '||' + ghz;
            if (!groups[key]) groups[key] = { cpu, ghz };
        }
        return Object.values(groups).map(g => {
            let spc = '', source = 'AI estimate';
            if (g.cpu !== '(not identified)') {
                /* Resolve SAPS/Core the SAME way the Processor Summary does and
                   report HOW the shown value was obtained:
                     1. Manual override (user typed it)          → "Entered manually"
                     2. SAP SD Benchmark lookup (uploaded file)  → "SAP SD Benchmark (…)"
                     3. Built-in knowledge-base table            → "AI estimate (knowledge base)"
                     4. Heuristic estimate                       → "AI estimate (heuristic)"
                   Overrides are GHz-aware: the same CPU at a different clock
                   speed is a distinct processor with its own SAPS/Core. */
                const ov = this._spcOverride && this._spcOverride[this._spcKey(g.cpu, g.ghz)];
                if (ov > 0) {
                    spc = ov; source = 'Entered manually';
                } else {
                    const bm = this._lookupBenchmark(g.cpu, g.ghz);
                    if (bm) {
                        spc = bm.sapsPerCore;
                        source = bm.confidence === 'exact'
                            ? 'SAP SD Benchmark (exact match)'
                            : `SAP SD Benchmark (closest: ${bm.ref})`;
                    } else {
                        const kb = this._getSAPSPerCore(g.cpu);
                        if (kb) { spc = kb; source = 'AI estimate (knowledge base)'; }
                        else    { spc = this._estimateSAPSPerCore(g.cpu, g.ghz); source = 'AI estimate (heuristic)'; }
                    }
                }
            }
            return { ...g, sapsPerCore: spc, source };
        });
    },

    /* ── View toggles ── */
    _showMainScreen() {
        const main = document.getElementById('sl-view-main');
        const rev  = document.getElementById('sl-view-review');
        if (main) main.style.display = 'block';
        if (rev)  rev.style.display  = 'none';
        /* Restore the upload card + module header for the main view */
        const hdr = document.getElementById('sl-module-header');
        if (hdr) hdr.style.display = 'block';
        const up = document.getElementById('sl-upload-card');
        if (up) up.style.display = 'block';
        document.getElementById('content-area')?.scrollTo({ top: 0, behavior: 'smooth' });
    },

    _showReviewScreen() {
        this._renderReviewScreen();
        const main = document.getElementById('sl-view-main');
        const rev  = document.getElementById('sl-view-review');
        if (main) main.style.display = 'none';
        if (rev)  rev.style.display  = 'block';
        /* Hide the generic upload header/card — not relevant on the review screen */
        const hdr = document.getElementById('sl-module-header');
        if (hdr) hdr.style.display = 'none';
        const up = document.getElementById('sl-upload-card');
        if (up) up.style.display = 'none';
        /* Unlock downstream modules now that the HW config is reviewable */
        this.app?.unlockModules?.([
            'tech-workload-app', 'tech-workload-db', 'app-workload', 'sizing-projection',
        ]);
        /* Persist final values so downstream modules see them even if the user
           navigates via the sidebar instead of the review Next button. */
        this._saveToApp();
        document.getElementById('content-area')?.scrollTo({ top: 0, behavior: 'smooth' });
    },

    _renderReviewScreen() {
        const info = this.systemInfo || {};

        /* ── Server Inventory recap (final per-host values) ──
           Mirror ALL columns from the first-screen Server Inventory table.
           `hv[i]` (from _hostVariables) is index-aligned with this.servers,
           and carries the computed SAPS (TBV) = SAPS/Core × Cores. */
        const hv = this._hostVariables();
        const gb = mb => (mb / 1024).toFixed(1);
        const invTbody = document.getElementById('sl-rv-inventory-tbody');
        if (invTbody) {
            invTbody.innerHTML = this.servers.length
                ? this.servers.map((s, i) => {
                    const saps = hv[i] ? hv[i].totalSaps : (parseInt(s.saps) || 0);
                    const dash = v => (v !== undefined && v !== null && String(v).trim() !== '' && v !== '-') ? v : '—';
                    return `
                    <tr>
                        <td>${dash(s.hostname)}</td>
                        <td>${dash(s.manufacturer)}</td>
                        <td>${dash(s.model)}</td>
                        <td>${dash(s.cpuType)}</td>
                        <td>${dash(s.virtualization)}</td>
                        <td>${dash(s.os)}</td>
                        <td>${dash(s.ghz)}</td>
                        <td>${dash(s.logicalCpus)}</td>
                        <td>${dash(s.cores)}</td>
                        <td><strong>${saps ? Number(saps).toLocaleString() : '—'}</strong></td>
                        <td>${s.memoryMB ? Number(parseInt(s.memoryMB)).toLocaleString() : '—'}</td>
                        <td>${dash(s.component)}</td>
                    </tr>`;
                }).join('')
                : `<tr><td colspan="12" style="color:var(--text-muted)">No servers parsed.</td></tr>`;
        }
        const invSub = document.getElementById('sl-rv-inv-sub');
        if (invSub) {
            const totSaps = hv.reduce((a, h) => a + (h.totalSaps || 0), 0);
            const totMem  = hv.reduce((a, h) => a + (h.memoryMB  || 0), 0);
            invSub.textContent =
                `${hv.length} server(s) · Total ${totSaps.toLocaleString()} SAPS · ${gb(totMem)} GB RAM`;
        }

        /* DB-host OS (+ patch level): take from the first DB-layer server row */
        const dbHost = this.servers.find(s => this._isDbLayer(s.component));
        const dbOs   = (dbHost && dbHost.os && dbHost.os !== '-') ? dbHost.os : (info.dbOs || '');

        /* ── Application tier information (editable) ── */
        const appFields = [
            { key: 'sid',     label: 'SAP System ID (SID)',   placeholder: 'e.g. S4P'              },
            { key: 'product', label: 'SAP Product & Release', placeholder: 'e.g. SAP S/4HANA 2022' },
        ];
        const mkField = (f) => `
            <div class="sysinfo-field">
                <label class="sysinfo-label">${f.label}</label>
                <input class="editable-cell" style="width:100%;margin-top:4px"
                       value="${(f.value ?? info[f.key] ?? '').toString().replace(/"/g,'&quot;')}"
                       placeholder="${f.placeholder}"
                       onchange="SystemLandscapeModule.systemInfo['${f.key}']=this.value;SystemLandscapeModule._saveToApp()">
            </div>`;
        document.getElementById('sl-rv-sysinfo-app').innerHTML =
            `<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:14px">` +
            appFields.map(mkField).join('') +
            `</div>`;

        /* ── Database tier information (editable) ── */
        const dbFields = [
            { key: 'dbSid',     label: 'Database SID',              placeholder: 'e.g. EHR'               },
            { key: 'dbSystem',  label: 'Database Product',           placeholder: 'e.g. SAP HANA Database' },
            { key: 'dbVersion', label: 'Database Release',           placeholder: 'e.g. 2.00.067.00'       },
            { key: 'dbOs',      label: 'DB Server OS & Patch Level', placeholder: 'e.g. SLES 15 SP4', value: dbOs },
        ];
        document.getElementById('sl-rv-sysinfo-db').innerHTML =
            `<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:14px">` +
            dbFields.map(mkField).join('') +
            `</div>`;

        /* ── SAPS per core table ── */
        const groups = this._cpuGroups().filter(g => g.cpu !== '(not identified)');
        document.getElementById('sl-rv-saps-tbody').innerHTML = groups.length
            ? groups.map(g => `
                <tr>
                    <td>${g.cpu}</td>
                    <td>${g.ghz || '—'}</td>
                    <td><strong>${g.sapsPerCore ? Number(g.sapsPerCore).toLocaleString() : '—'}</strong></td>
                    <td style="font-size:12px;color:var(--text-secondary)">${g.source}</td>
                </tr>`).join('')
            : `<tr><td colspan="4" style="color:var(--text-muted)">No processors identified.</td></tr>`;

        /* ── Deployed capacity ──
           Compute from the SAME per-host variables used downstream so the
           totals always match: SAPS (TBV) per host = SAPS/Core × Cores.
           (`hv` and `gb` are declared at the top of this method.) */
        const sumSaps  = list => list.reduce((a, h) => a + (h.totalSaps || 0), 0);
        const sumMem   = list => list.reduce((a, h) => a + (h.memoryMB  || 0), 0);
        const sumCores = list => list.reduce((a, h) => a + (h.cores     || 0), 0);

        const all = hv;
        const app = hv.filter(h => h.layer === 'app');
        const db  = hv.filter(h => h.layer === 'db');

        const capRows = [
            ['Total deployed capacity on the system',             sumSaps(all), sumMem(all), sumCores(all)],
            ['Total deployed capacity on the application layer',  sumSaps(app), sumMem(app), sumCores(app)],
            ['Total deployed capacity on the DB layer',           sumSaps(db),  sumMem(db),  sumCores(db)],
        ];
        document.getElementById('sl-rv-capacity-tbody').innerHTML = capRows.map(([label, saps, mem, cores]) => `
            <tr>
                <td>${label}</td>
                <td><strong>${saps.toLocaleString()}</strong> SAPS</td>
                <td><strong>${gb(mem)}</strong> GB</td>
                <td><strong>${cores.toLocaleString()}</strong></td>
            </tr>`).join('');

        /* ── Server Inventory copy / export buttons ── */
        const copyBtn = document.getElementById('sl-rv-btn-copy');
        if (copyBtn) copyBtn.onclick = () => this._copyInventoryForExcel(copyBtn);
        const expBtn = document.getElementById('sl-rv-btn-export');
        if (expBtn) expBtn.onclick = () => this._exportInventory();
    },

    /* Column definitions for the review-screen Server Inventory — the single
       source for both the clipboard (TSV) and XLSX exports so they stay in
       sync with the on-screen table. */
    _inventoryExportRows() {
        const hv = this._hostVariables();
        return this.servers.map((s, i) => ({
            'Hostname'        : s.hostname       || '',
            'Manufacturer'    : s.manufacturer   || '',
            'Model'           : s.model          || '',
            'CPU Type'        : s.cpuType        || '',
            'Virtualization'  : s.virtualization || '',
            'Operating System': s.os             || '',
            'GHz'             : s.ghz            || '',
            'Logical CPUs'    : s.logicalCpus    || '',
            'Cores'           : parseInt(s.cores) || 0,
            'SAPS (TBV)'      : hv[i] ? hv[i].totalSaps : (parseInt(s.saps) || 0),
            'Memory (MB)'     : parseInt(s.memoryMB) || 0,
            'Component Type'  : s.component      || '',
        }));
    },

    /* Copy the Server Inventory to the clipboard as TAB-separated values with a
       header row. Pasting into Excel lands each column in its own cell (Excel
       splits on tabs), preserving the table layout — no formatting lost. */
    _copyInventoryForExcel(btn) {
        const rows = this._inventoryExportRows();
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
        const fallback = () => {
            const ta = document.createElement('textarea');
            ta.value = tsv;
            ta.style.position = 'fixed';
            ta.style.opacity = '0';
            document.body.appendChild(ta);
            ta.select();
            try { document.execCommand('copy'); done(); } catch (e) { alert('Copy failed — select the table manually.'); }
            document.body.removeChild(ta);
        };

        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(tsv).then(done).catch(fallback);
        } else {
            fallback();
        }
    },

    /* Download the Server Inventory as XLSX — same column set as the
       first-screen export, using the final review values. */
    _exportInventory() {
        const rows = this._inventoryExportRows();
        if (!rows.length) return;
        downloadXlsx(rows, 'Server Inventory', `server-inventory-${_dateTag()}.xlsx`);
    },

    /* ── Export ── */
    _downloadAnalysis() {
        downloadXlsx(
            this.servers.map(s => ({ ...s })),
            'System Landscape',
            `system-landscape-${_dateTag()}.xlsx`
        );
    },

    /* Copy the editable Step 2 Server Inventory table to the clipboard as
       TAB-separated values, reading current input values directly from the
       DOM so user edits that haven't yet been committed are included. */
    _copyEditableTableForExcel(btn) {
        const headers = [
            'Hostname', 'Manufacturer', 'Model', 'CPU Type', 'Virtualization',
            'Operating System', 'GHz', 'Logical CPUs', 'Cores', 'SAPS (TBV)',
            'Memory (MB)', 'Component Type',
        ];
        const fieldMap = {
            'Hostname':         'hostname',
            'Manufacturer':     'manufacturer',
            'Model':            'model',
            'CPU Type':         'cpuType',
            'Virtualization':   'virtualization',
            'Operating System': 'os',
            'GHz':              'ghz',
            'Logical CPUs':     'logicalCpus',
            'Cores':            'cores',
            'SAPS (TBV)':       'saps',
            'Memory (MB)':      'memoryMB',
            'Component Type':   'component',
        };

        const rows = this.servers.map((s, i) => {
            const row = {};
            for (const h of headers) {
                const f = fieldMap[h];
                /* Read live DOM value first (captures unsaved edits), fall back to model */
                if (f === 'component') {
                    const sel = document.querySelector(
                        `#sl-hw-tbody tr:nth-child(${i + 1}) .editable-select`
                    );
                    row[h] = sel ? sel.value : (s.component || '');
                } else {
                    const inp = document.querySelector(`input[data-i="${i}"][data-f="${f}"]`);
                    row[h] = inp ? inp.value : (s[f] != null ? String(s[f]) : '');
                }
            }
            return row;
        });

        const tsv = [headers.join('\t')]
            .concat(rows.map(r => headers.map(h => String(r[h] ?? '')).join('\t')))
            .join('\r\n');

        const done = () => {
            if (!btn) return;
            const orig = btn.innerHTML;
            btn.innerHTML = '&#10003; Copied!';
            setTimeout(() => { btn.innerHTML = orig; }, 1800);
        };
        const fallback = () => {
            const ta = document.createElement('textarea');
            ta.value = tsv;
            ta.style.cssText = 'position:fixed;opacity:0;top:0;left:0';
            document.body.appendChild(ta);
            ta.select();
            try { document.execCommand('copy'); done(); }
            catch (_) { alert('Copy failed — select the table manually.'); }
            document.body.removeChild(ta);
        };

        if (navigator.clipboard?.writeText) {
            navigator.clipboard.writeText(tsv).then(done).catch(fallback);
        } else {
            fallback();
        }
    },

    /* ── Essential per-host variables for downstream analysis modules ──
       hostname, cpuType, ghz, cores, memoryMB, sapsPerCore, and the
       per-host total SAPS (TBV) = sapsPerCore × cores.
       When s._sapsManual is true the user typed SAPS(TBV) directly —
       preserve the exact value and derive a full-precision sapsPerCore
       from it so downstream capacity figures remain consistent. */
    _hostVariables() {
        const aliases = this.hostAliases || {};
        return this.servers.map(s => {
            const cores = parseInt(s.cores) || 0;
            /* SAPS/Core resolution order (GHz-aware):
               1. manual override for this exact (cpuType, ghz),
               2. built-in SAPS_DB / benchmark lookup by cpuType,
               3. derive from the host's own SAPS ÷ cores (full precision). */
            const ov    = this._spcOverride && this._spcOverride[this._spcKey(s.cpuType, s.ghz)];

            let spc, totalSaps;
            if (s._sapsManual && parseFloat(s.saps) > 0) {
                /* User typed SAPS(TBV) directly — this is the source of truth.
                   Derive SAPS/Core with full float precision; display may round it
                   but calculations must use the exact quotient so that
                   cores × sapsPerCore == totalSaps to full precision. */
                totalSaps = parseFloat(s.saps);
                spc       = cores > 0 ? totalSaps / cores : 0;
            } else if (ov > 0) {
                /* Manual SAPS/Core override in Processor Summary. */
                spc       = ov;
                totalSaps = cores ? spc * cores : (parseFloat(s.saps) || 0);
            } else {
                const dbSpc = this._getSAPSPerCore(s.cpuType) || 0;
                spc = dbSpc
                    || (cores > 0 && parseFloat(s.saps) > 0 ? parseFloat(s.saps) / cores : 0);
                totalSaps = (parseFloat(s.saps) || 0) || (spc && cores ? spc * cores : 0);
            }

            return {
                hostname:    s.hostname || '',
                virtualHost: aliases[(s.hostname || '').toLowerCase()] || '',
                manufacturer: s.manufacturer   || '',
                model:        s.model          || '',
                virtualization: s.virtualization || '',
                os:          s.os              || '',
                logicalCpus: s.logicalCpus     || '',
                cpuType:     s.cpuType  || '',
                ghz:         s.ghz      || '',
                cores,
                memoryMB:    parseInt(s.memoryMB) || 0,
                sapsPerCore: spc || 0,
                totalSaps,
                component:   s.component || '',
                layer:       this._isDbLayer(s.component) ? 'db'
                           : this._isAppLayer(s.component) ? 'app' : 'other',
            };
        });
    },

    _saveToApp() {
        if (!this.app) return;
        const appServers = this.servers.filter(s => this._isAppLayer(s.component));
        const dbServers  = this.servers.filter(s => this._isDbLayer(s.component));

        this.app.setModuleData('system-landscape', {
            rows:          this.servers,
            fileName:      'EWA Report',
            headers:       Object.keys(this.servers[0] || {}),
            analysis:      this.servers,
            systemInfo:    this.systemInfo,
            appServers,    /* application layer (SAP AS + SAP AS / SAP ABAP SCS + SAP ABAP SCS) */
            dbServers,     /* database layer (DB) */
            hostVariables: this._hostVariables(),   /* essential per-host variables */
            hostAliases:   this.hostAliases || {},  /* { physicalLower -> virtualName } */
            hasIBMServers: this.hasIBMServers,
        });
    },

    /* ── Progress / Error helpers ── */
    _showProgress(msg) {
        const el  = document.getElementById('sl-progress');
        const txt = document.getElementById('sl-progress-msg');
        if (el)  el.style.display = 'flex';
        if (txt) txt.textContent  = msg;
    },
    _updateProgress(msg) {
        const txt = document.getElementById('sl-progress-msg');
        if (txt) txt.textContent = msg;
    },
    _hideProgress() {
        const el = document.getElementById('sl-progress');
        if (el) el.style.display = 'none';
    },
    _showError(msg) {
        const wrap = document.getElementById('sl-error');
        const txt  = document.getElementById('sl-error-msg');
        if (wrap) wrap.style.display = 'flex';
        if (txt)  txt.textContent    = msg;
        this._hideProgress();
    },
    _hideError() {
        const el = document.getElementById('sl-error');
        if (el) el.style.display = 'none';
    },
};
