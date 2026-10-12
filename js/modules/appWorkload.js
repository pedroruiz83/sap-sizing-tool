/* ─────────────────────────────────────────────────────────
   Module 04 – Application Workload – Business Metrics
   ─────────────────────────────────────────────────────────
   Two upload zones in Step 1 (timeptype reuses tasktype data):
     • Load History          (dbcpu)   – ST03 Load and History › Load History and Distribution
     • Workload by Task Type (tasktype) – ST03 Workload › Month view
   Step 2 – Transaction Profile (4 optional zones, same task-type format)
   Step 3 – Charts
   ─────────────────────────────────────────────────────────*/

const AppWorkloadModule = {

    /* ── Persistent state ── */
    _step: 1,
    _uploads: {
        dbcpu:      null,   // Load History (monthly rows, date as X axis)
        tasktype:   null,   // Workload by Task Type (task type as X axis)
        dialog:     null,
        http:       null,
        background: null,
        rfc:        null,
    },

    /* Persistent <SID>_WorkloadAnalysis.xlsx workbook (built incrementally; each
       screen writes only its own worksheet) and the cached DB-vs-CPU dataset. */
    _wb: null,
    _dbcpuPrepared: null,

    resetState() {
        this._step = 1;
        this._uploads = { dbcpu: null, tasktype: null, dialog: null, http: null, background: null, rfc: null };
        this._charts = [];
        this._wb = null;
        this._dbcpuPrepared = null;
        this._comboRegistry = {};
    },

    /* ── Session save/restore contract ──
       Returns only serializable working state (parsed uploads + current step).
       The in-memory workbook (_wb) and Chart instances are rebuilt on demand,
       so they are intentionally omitted. */
    getSessionState() {
        return { step: this._step, uploads: this._uploads };
    },
    setSessionState(st) {
        if (!st) return;
        if (typeof st.step === 'number') this._step = st.step;
        if (st.uploads) {
            this._uploads = {
                dbcpu: null, tasktype: null, dialog: null, http: null, background: null, rfc: null,
                ...st.uploads,
            };
            /* Re-derive any profile uploads that lost their derived data during
               serialisation (e.g. very large files that were partially truncated). */
            ['dialog', 'http', 'background', 'rfc'].forEach(key => {
                const u = this._uploads[key];
                if (u && u.rows && u.rows.length && (!u.derived || !u.derived.rows)) {
                    u.derived = this._computeProfileDerived(u.rows, key);
                }
            });
        }
        /* Workbook + charts are regenerated when the Charts screen initialises. */
        this._wb = null;
        this._dbcpuPrepared = null;
        this._comboRegistry = {};
        this._charts = [];
    },

    _charts: [],

    /* ══════════════════════════════════════════════════════
       CANONICAL FIELD MAPS
       TXT short header / XLSX long header → canonical name
    ══════════════════════════════════════════════════════ */

    /* ── Load History (dbcpu) ── */
    _DBCPU_ALIASES: {
        /* TXT */
        'start date':                              'date',
        '# steps':                                 'steps',
        'ø time':                                  'resp_ms',
        'avg. proc. time':                         'proc_ms',
        'ø cpu time':                              'cpu_ms',
        'ø db time':                               'db_ms',
        'ø time/db procedure call':                'db_proc_ms',
        'ø waittime':                              'wait_ms',
        'ø roll in~':                              'rollin_ms',
        'ø roll wait time':                        'rollwait_ms',
        'ø load- + gen. time':                     'loadgen_ms',
        'ø locktime':                              'lock_ms',
        'ø cpic/rfc':                              'rfc_ms',
        'ø time/frontend':                         'frontend_ms',
        'ø gui time':                              'gui_ms',
        '# trips':                                 'roundtrips',
        'db data (kb)':                            'db_kb',
        't vmc cpu':                               'vmc_cpu_s',
        /* XLSX */
        'date of first statistical record':        'date',
        'number of dialog steps':                  'steps',
        'average response time/dialog step (ms)':  'resp_ms',
        'avg. processing time':                    'proc_ms',
        'average cpu time (ms)':                   'cpu_ms',
        'ø db time (ms)':                          'db_ms',
        'ø time/db procedure call (ms)':           'db_proc_ms',
        'average wait time per dialog step (ms)':  'wait_ms',
        'average roll in time (ms)':               'rollin_ms',
        'average roll wait time (ms)':             'rollwait_ms',
        'average load and generation time (ms)':   'loadgen_ms',
        'average lock time per dialog step (ms)':  'lock_ms',
        'average rfc interface time (ms)':         'rfc_ms',
        'average frontend network time (ms)':      'frontend_ms',
        'average gui time per dialog step (ms)':   'gui_ms',
        'number of roundtrips':                    'roundtrips',
        'requested data (kb)':                     'db_kb',
        'total vmc cpu time (s)':                  'vmc_cpu_s',
        /* XLSX — ST03N "Load History (monthly)" actual export headers
           (SAP GUI "Export → Spreadsheet"). These differ from the long
           descriptive labels above, so they are listed explicitly. */
        'date of first record':                    'date',
        'number of steps':                         'steps',
        'ø response time (ms)':                    'resp_ms',
        'ø response time':                         'resp_ms',
        'ø cpu time':                              'cpu_ms',
        'ø db time (ms)':                          'db_ms',
        'ø wait time (ms)':                        'wait_ms',
        'ø wait time':                             'wait_ms',
        'ø roll in time (ms)':                     'rollin_ms',
        'ø roll in time':                          'rollin_ms',
        'ø load + gen. time (ms)':                 'loadgen_ms',
        'ø load + gen. time':                      'loadgen_ms',
        'ø lock time (ms)':                        'lock_ms',
        'ø lock time':                             'lock_ms',
        'avg. cpic/rfc time':                      'rfc_ms',
        'ø fe network time (ms)':                  'frontend_ms',
        'ø fe network time':                       'frontend_ms',
        'ø gui time/dialog step':                  'gui_ms',
        '# roundtrips':                            'roundtrips',
        'db data (kb)':                            'db_kb',
    },

    /* Required canonical fields for Load History validation */
    _DBCPU_REQUIRED: ['date', 'steps', 'resp_ms', 'cpu_ms', 'db_ms'],

    /* ── Workload by Task Type (tasktype / profile uploads) ── */
    _TASKTYPE_ALIASES: {
        /* TXT */
        'task type':                               'task_type',
        '# steps':                                 'steps',
        'ø time':                                  'resp_ms',
        'avg. proc. time':                         'proc_ms',
        'ø cpu time':                              'cpu_ms',
        'ø db time':                               'db_ms',
        'ø time/db procedure call':                'db_proc_ms',
        'ø waittime':                              'wait_ms',
        'ø roll in~':                              'rollin_ms',
        'ø roll wait time':                        'rollwait_ms',
        'ø load- + gen. time':                     'loadgen_ms',
        'ø locktime':                              'lock_ms',
        'ø cpic/rfc':                              'rfc_ms',
        'ø time/frontend':                         'frontend_ms',
        'ø gui time':                              'gui_ms',
        '# trips':                                 'roundtrips',
        'kb':                                      'db_kb',
        '#vmc calls':                              'vmc_calls',
        't vmc cpu':                               'vmc_cpu_s',
        't vmcelaps':                              'vmc_elaps_s',
        'avgvmc cpu':                              'vmc_cpu_avg_ms',
        'avgvmcelap':                              'vmc_elap_avg_ms',
        /* XLSX */
        'task type name':                          'task_type',
        'number of dialog steps':                  'steps',
        'average response time/dialog step (ms)':  'resp_ms',
        'avg. processing time':                    'proc_ms',
        'average cpu time (ms)':                   'cpu_ms',
        'ø db time (ms)':                          'db_ms',
        'ø time/db procedure call (ms)':           'db_proc_ms',
        'average wait time per dialog step (ms)':  'wait_ms',
        'average roll in time (ms)':               'rollin_ms',
        'average roll wait time (ms)':             'rollwait_ms',
        'average load and generation time (ms)':   'loadgen_ms',
        'average lock time per dialog step (ms)':  'lock_ms',
        'average rfc interface time (ms)':         'rfc_ms',
        'average frontend network time (ms)':      'frontend_ms',
        'average gui time per dialog step (ms)':   'gui_ms',
        'number of roundtrips':                    'roundtrips',
        'requested data (kb)':                     'db_kb',
        'number of vmc calls':                     'vmc_calls',
        'total vmc cpu time (s)':                  'vmc_cpu_s',
        'totla vmc elapsed time (s)':              'vmc_elaps_s',   // note: typo in SAP header preserved
        'total vmc elapsed time (s)':              'vmc_elaps_s',
        'average vmc cpu time (ms)':               'vmc_cpu_avg_ms',
        'average vmc elapsed time (ms)':           'vmc_elap_avg_ms',
        /* XLSX — ST03N "Workload per month / Task Type" actual export headers.
           SAP GUI "Export → Spreadsheet" emits these shorter labels. */
        'task type':                               'task_type',
        'number of steps':                         'steps',
        'ø response time (ms)':                    'resp_ms',
        'ø response time':                         'resp_ms',
        'ø cpu time':                              'cpu_ms',
        'ø wait time (ms)':                        'wait_ms',
        'ø wait time':                             'wait_ms',
        'ø roll in time (ms)':                     'rollin_ms',
        'ø roll in time':                          'rollin_ms',
        'ø load + gen. time (ms)':                 'loadgen_ms',
        'ø load + gen. time':                      'loadgen_ms',
        'ø lock time (ms)':                        'lock_ms',
        'ø lock time':                             'lock_ms',
        'avg. cpic/rfc time':                      'rfc_ms',
        'ø fe network time (ms)':                  'frontend_ms',
        'ø fe network time':                       'frontend_ms',
        'ø gui time/dialog step':                  'gui_ms',
        '# roundtrips':                            'roundtrips',
        'db data (kb)':                            'db_kb',
    },
    _TASKTYPE_REQUIRED: ['task_type', 'steps', 'resp_ms', 'cpu_ms', 'db_ms'],

    /* ══════════════════════════════════════════════════════
       CONSOLIDATED VALIDATION MODEL  (_WORKLOAD_MODEL)
       ──────────────────────────────────────────────────────
       Single column-NAME-driven specification for the four
       Transaction Profile upload types. Replaces the former
       _PROFILE_ALIASES / _RFC_ALIASES maps and their flat
       required lists.

       Each field entry:
         canonical  – internal field name used everywhere downstream
         label      – human label (error messages + validation metadata)
         aliases    – normalized header spellings seen in TXT *and* XLSX
                      exports (lower-cased, single-spaced, unit suffix
                      optional — matched via _canonicalize normalization)
         type       – one of:
                        text | identifier | integer | numeric |
                        duration_s | duration_ms | data_kb |
                        data_byte | date | percentage
         required   – true → validation fails if the column is absent
         nullable   – false → empty/invalid cell is flagged (still kept
                      as null); true → empty is acceptable
         txtPos     – OPTIONAL 0-based positional hint, ONLY consulted to
                      disambiguate repeated short labels in TXT exports
                      (e.g. RFC "Ø Time" appears 3× for response /
                      DB-procedure / frontend averages). Column-name
                      resolution always wins when it is unambiguous.

       Future ST03 header changes are absorbed by adding a spelling to
       the relevant `aliases` array — no positional code to touch.
    ══════════════════════════════════════════════════════ */

    /* Shared metric fields common to Dialog / HTTP / Background / RFC.
       Identity columns (trans_rep / job_name / system_id / action /
       user_id) are layered on per type below. */
    _BASE_FIELDS: [
        { canonical: 'steps',            label: '# Steps',                 type: 'integer',     required: true,  nullable: false,
          aliases: ['# steps', 'number of steps', 'number of dialog steps'] },

        { canonical: 'resp_total_s',     label: 'T Response Time (s)',     type: 'duration_s',  required: false, nullable: true,
          aliases: ['t response time', 't response time (s)'] },
        { canonical: 'resp_ms',          label: 'Ø Response Time (ms)',    type: 'duration_ms', required: true,  nullable: false,
          aliases: ['ø time', 'ø response time (ms)', 'ø response time'], txtRepeat: 'ø time', txtPos: 0 },

        { canonical: 'proc_total_s',     label: 'Processing Time (s)',     type: 'duration_s',  required: false, nullable: true,
          aliases: ['process.', 'processing time (s)', 'total processing time (s)'] },
        { canonical: 'proc_ms',          label: 'Avg. Processing Time',    type: 'duration_ms', required: false, nullable: true,
          aliases: ['avg. proc. time', 'avg. processing time'] },

        { canonical: 'cpu_total_s',      label: 'T CPU Time (s)',          type: 'duration_s',  required: true,  nullable: false,
          aliases: ['t cpu~', 't cpu time (s)', 't cpu time'] },
        { canonical: 'cpu_ms',           label: 'Ø CPU Time (ms)',         type: 'duration_ms', required: true,  nullable: false,
          aliases: ['ø cpu~', 'ø cpu time (ms)', 'ø cpu time'] },

        { canonical: 'db_total_s',       label: 'T DB Time (s)',           type: 'duration_s',  required: true,  nullable: false,
          aliases: ['t db time', 't database time (s)', 't database time'] },
        { canonical: 'db_ms',            label: 'Ø DB Time (ms)',          type: 'duration_ms', required: true,  nullable: false,
          aliases: ['ø db time', 'ø db time (ms)'] },

        { canonical: 'db_proc_total_s',  label: 'T DB Procedure Time (s)', type: 'duration_s',  required: false, nullable: true,
          aliases: ['t time', 't db procedure time', 't time for db procedure accesses (s)'] },
        { canonical: 'db_proc_ms',       label: 'Ø Time/DB Procedure (ms)',type: 'duration_ms', required: false, nullable: true,
          aliases: ['ø time/db procedure', 'ø time/db procedure call (ms)'], txtRepeat: 'ø time', txtPos: 1 },

        { canonical: 'wait_ms',          label: 'Ø Wait Time (ms)',        type: 'duration_ms', required: false, nullable: true,
          aliases: ['ø waittime', 'ø wait time (ms)', 'ø wait time'] },

        { canonical: 'roundtrips',       label: '# Roundtrips',            type: 'integer',     required: false, nullable: true,
          aliases: ['# trips', '# roundtrips', 'number of roundtrips'] },

        { canonical: 'frontend_ms',      label: 'Ø FE Network Time (ms)',  type: 'duration_ms', required: false, nullable: true,
          aliases: ['ø time/frontend', 'ø fe network time (ms)', 'average frontend network time (ms)'], txtRepeat: 'ø time', txtPos: 2 },
        { canonical: 'gui_ms',           label: 'Ø GUI Time (ms)',         type: 'duration_ms', required: false, nullable: true,
          aliases: ['ø gui time', 'ø gui time/dialog step', 'average gui time per dialog step (ms)'] },

        { canonical: 'db_kb',            label: 'DB Data (KB)',            type: 'data_kb',     required: false, nullable: true,
          aliases: ['kb', 'db data (kb)', 'requested data (kb)'] },
    ],

    /* VMC + roll-wait + data-volume fields used only by the trans/rep
       profiles (Dialog / HTTP / Background). */
    _PROFILE_EXTRA_FIELDS: [
        { canonical: 'rollwait_total_s', label: 'T Roll Wait Time (s)',    type: 'duration_s',  required: false, nullable: true,
          aliases: ['t roll wait time', 't roll wait time (s)'] },
        { canonical: 'rollwait_ms',      label: 'Ø Roll Wait Time (ms)',   type: 'duration_ms', required: false, nullable: true,
          aliases: ['ø roll wait time'] },
        { canonical: 'input_byte',       label: 'Ø Input (Byte)',          type: 'data_byte',   required: false, nullable: true,
          aliases: ['ø input', 'avg. data volume to app. server (byte)'] },
        { canonical: 'output_byte',      label: 'Ø Output (Byte)',         type: 'data_byte',   required: false, nullable: true,
          aliases: ['ø output', 'average data volume to front end (byte)'] },
        { canonical: 'vmc_calls',        label: '# VMC Calls',             type: 'integer',     required: false, nullable: true,
          aliases: ['#vmc calls', 'number of vmc calls'] },
        { canonical: 'vmc_cpu_s',        label: 'T VMC CPU (s)',           type: 'duration_s',  required: false, nullable: true,
          aliases: ['t vmc cpu', 'total vmc cpu time (s)'] },
        { canonical: 'vmc_elaps_s',      label: 'T VMC Elapsed (s)',       type: 'duration_s',  required: false, nullable: true,
          aliases: ['t vmcelaps', 'totla vmc elapsed time (s)', 'total vmc elapsed time (s)', 'total vmc elapsed time'] },
        { canonical: 'vmc_cpu_avg_ms',   label: 'Avg VMC CPU (ms)',        type: 'duration_ms', required: false, nullable: true,
          aliases: ['avgvmc cpu', 'average vmc cpu time (ms)', 'average vmc cpu time'] },
        { canonical: 'vmc_elap_avg_ms',  label: 'Avg VMC Elapsed (ms)',    type: 'duration_ms', required: false, nullable: true,
          aliases: ['avgvmcelap', 'average vmc elapsed time (ms)', 'average vmc elapsed time'] },
    ],

    /* RFC-only time field: RFC exports carry a total wait column that the
       trans/rep profiles do not. */
    _RFC_EXTRA_FIELDS: [
        { canonical: 'wait_total_s',     label: 'T Wait Time (s)',         type: 'duration_s',  required: false, nullable: true,
          aliases: ['t waittime', 't wait time (s)'] },
    ],

    /* Identity (text / key) columns per upload type. */
    _IDENTITY_FIELDS: {
        transrep: [
            { canonical: 'trans_rep', label: 'Report or Transaction name', type: 'identifier', required: true,  nullable: false,
              aliases: ['trans/rep.', 'report or transaction name', 'report/transaction'] },
        ],
        background: [
            { canonical: 'trans_rep', label: 'Report or Transaction name', type: 'identifier', required: true,  nullable: false,
              aliases: ['trans/rep.', 'report or transaction name', 'report/transaction'] },
            { canonical: 'job_name',  label: 'Name of Background Job',      type: 'identifier', required: false, nullable: true,
              aliases: ['name of background job'] },
        ],
        rfc: [
            { canonical: 'system_id', label: 'System ID',                  type: 'identifier', required: true,  nullable: false,
              aliases: ['system id'] },
            { canonical: 'action',    label: 'Action in Triggering System',type: 'identifier', required: true,  nullable: false,
              aliases: ['action in triggering system', 'action/transaction', 'action'] },
            { canonical: 'user_id',   label: 'User ID',                    type: 'identifier', required: false, nullable: true,
              aliases: ['user id', 'remote user id'] },
        ],
    },

    /* Assembled per-type field lists (identity + metrics, in TXT column
       order so txtPos hints line up for repeated-label disambiguation). */
    get _WORKLOAD_MODEL() {
        if (this.__workloadModelCache) return this.__workloadModelCache;
        const dialog = [
            ...this._IDENTITY_FIELDS.transrep,
            ...this._BASE_FIELDS,
            ...this._PROFILE_EXTRA_FIELDS,
        ];
        const background = [
            ...this._IDENTITY_FIELDS.background,
            ...this._BASE_FIELDS,
            ...this._PROFILE_EXTRA_FIELDS,
        ];
        const rfc = [
            ...this._IDENTITY_FIELDS.rfc,
            ...this._BASE_FIELDS,
            ...this._RFC_EXTRA_FIELDS,
        ];
        this.__workloadModelCache = {
            dialog,
            http: dialog,          /* HTTP/HTTPS shares the trans/rep schema */
            background,
            rfc,
        };
        return this.__workloadModelCache;
    },

    /* ══════════════════════════════════════════════════════
       RENDER
    ══════════════════════════════════════════════════════ */
    render() {
        this._charts = [];
        if (this._step === 1) return this._renderStep1();
        if (this._step === 2) return this._renderDbCpuScreen();
        if (this._step === 3) return this._renderStep2();
        return this._renderStep3();
    },

    /* ── Step 1 ── */
    _renderStep1() {
        return `
        <div class="module-wrap">
            <div class="module-header">
                <div class="module-crumb">Analysis › Section 04 › Step 1 of 4</div>
                <h1 class="module-title">Application Workload – Total View</h1>
                <p class="module-desc">
                    Extract the performance statistics below from SAP transaction
                    <strong>ST03 / ST03N</strong> and upload each file here.
                    Both sections are required to proceed.
                </p>
            </div>

            <div id="aw-validation-error" style="display:none" class="alert-error">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24"
                     fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                    <circle cx="12" cy="12" r="10"/>
                    <line x1="15" y1="9" x2="9" y2="15"/>
                    <line x1="9" y1="9" x2="15" y2="15"/>
                </svg>
                <div id="aw-validation-msg"></div>
            </div>

            ${this._uploadCard('aw-dbcpu', 'dbcpu',
                'Load History',
                `<strong>ST03 / ST03N</strong> → Load and History → Load History and Distribution →
                 TOTAL → Task Type <em>ALL</em> → Month view`,
                'aw-dbcpu-status')}

            ${this._uploadCard('aw-tasktype', 'tasktype',
                'Workload According to Task Type',
                `<strong>ST03 / ST03N</strong> → Workload → Month → choose the last complete month`,
                'aw-tasktype-status')}

            <div class="action-row" style="margin-top:32px">
                <button class="btn btn-blue" id="aw-btn-step2" disabled>
                    DB vs CPU Analysis
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2.5">
                        <polyline points="9 18 15 12 9 6"/>
                    </svg>
                </button>
            </div>
        </div>`;
    },

    /* ══════════════════════════════════════════════════════
       SCREEN 2 — DB vs CPU Time vs Steps (new)
       Prepares the Load History data, writes the
       ST03_DB_vs_CPU_vs_Steps worksheet of <SID>_WorkloadAnalysis.xlsx,
       and renders an editable combination chart.
    ══════════════════════════════════════════════════════ */
    _COMBO_COLORS: { cpu: '#FFC000', db: '#A5A5A5', steps: '#00B0F0' },

    _renderDbCpuScreen() {
        return `
        <div class="module-wrap">
            <div class="module-header">
                <div class="module-crumb">Analysis › Section 04 › Step 2 of 4</div>
                <h1 class="module-title">Application Workload – DB vs CPU Time vs Steps</h1>
                <p class="module-desc">
                    The latest three completed months from the Load History upload are analysed
                    below (the current, incomplete month is excluded). CPU and DB processing time
                    are converted to hours and plotted against the number of dialog steps. The
                    results are written to <strong>ST03_DB_vs_CPU_vs_Steps</strong> in the
                    downloadable workbook.
                </p>
            </div>

            <div id="aw-dbcpu-error" style="display:none" class="alert-error">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24"
                     fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                    <circle cx="12" cy="12" r="10"/>
                    <line x1="15" y1="9" x2="9" y2="15"/>
                    <line x1="9" y1="9" x2="15" y2="15"/>
                </svg>
                <div id="aw-dbcpu-error-msg"></div>
            </div>

            <!-- Summary stat cards -->
            <div class="stats-grid" id="aw-dbcpu-stats" style="margin-bottom:16px"></div>

            <!-- Combination chart mount -->
            <div id="aw-dbcpu-chart"></div>

            <!-- Summary table + # Steps trend observation -->
            <div id="aw-dbcpu-summary"></div>

            <!-- ══ Monthly Analysis (Workload According to Task Type) ══ -->
            <div class="section-title" style="margin-top:34px">Monthly Analysis — Workload According to Task Type</div>
            <div id="aw-ma-pie"></div>
            <div id="aw-ma-timeptype" style="margin-top:28px"></div>

            <div class="action-row" style="margin-top:24px;justify-content:space-between">
                <button class="btn" id="aw-dbcpu-back">
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2.5">
                        <polyline points="15 18 9 12 15 6"/>
                    </svg>
                    Back
                </button>
                <div style="display:flex;gap:10px">
                    <button class="btn" id="aw-dbcpu-download">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="7 10 12 15 17 10"/>
                            <line x1="12" y1="15" x2="12" y2="3"/>
                        </svg>
                        Download Workbook
                    </button>
                    <button class="btn btn-blue" id="aw-dbcpu-next">
                        Next Step
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="9 18 15 12 9 6"/>
                        </svg>
                    </button>
                </div>
            </div>
        </div>`;
    },

    _initDbCpuScreen(app) {
        this.app = app;
        this._charts = [];

        document.getElementById('aw-dbcpu-back')?.addEventListener('click', () => {
            this._step = 1;
            this._updateSubNav(app);
            app.navigateTo('app-workload');
        });
        document.getElementById('aw-dbcpu-next')?.addEventListener('click', () => {
            this._step = 3;
            this._updateSubNav(app);
            this._saveModuleData(app);
            app.navigateTo('app-workload');
        });
        document.getElementById('aw-dbcpu-download')?.addEventListener('click', () => this._downloadWorkbook());

        /* Prepare + validate the dataset, write the worksheet, build the chart. */
        const prep = this._prepareDbCpu();
        if (prep.error) {
            this._showDbCpuError(prep.error);
            const dl = document.getElementById('aw-dbcpu-download');
            if (dl) dl.disabled = true;
            return;
        }
        this._dbcpuPrepared = prep;
        this._writeDbCpuSheet(prep);
        this._renderDbCpuStats(prep);
        this._mountComboChart(document.getElementById('aw-dbcpu-chart'), prep);
        this._renderDbCpuSummary(document.getElementById('aw-dbcpu-summary'), prep);

        /* Monthly Analysis (task-type) — pie + clustered-column charts, and the
           ST03_Monthly_Analysis worksheet. Requires the Task Type upload. */
        const ma = this._prepareMonthlyAnalysis();
        if (ma && !ma.error) {
            this._writeMonthlySheet(ma);
            this._mountTaskTypePie(document.getElementById('aw-ma-pie'), ma);
            this._mountTimePerTaskType(document.getElementById('aw-ma-timeptype'), ma);
        } else if (ma && ma.error) {
            const pie = document.getElementById('aw-ma-pie');
            if (pie) pie.innerHTML = `<div class="chart-note" style="color:var(--text-muted)">${ma.error}</div>`;
        }
    },

    /* ── Parse a DD.MM.YYYY (or DD/MM/YYYY) date to a Date object ── */
    _parseDotDate(raw) {
        if (!raw) return null;
        if (raw instanceof Date) return isNaN(raw) ? null : raw;
        const p = String(raw).trim().replace(/\./g, '/').split('/');
        if (p.length < 3) return null;
        const d = parseInt(p[0], 10), m = parseInt(p[1], 10), y = parseInt(p[2], 10);
        if (isNaN(d) || isNaN(m) || isNaN(y)) return null;
        return new Date(y, m - 1, d);
    },

    /* ── Normalize an XLSX date cell to a DD.MM.YYYY string ──
       Handles a JS Date (from cellDates:true), an Excel serial number
       (days since 1899-12-30), and an already-formatted date string. */
    _xlsxDateToDotStr(cell) {
        const pad = n => String(n).padStart(2, '0');
        const fromDate = dt => `${pad(dt.getDate())}.${pad(dt.getMonth() + 1)}.${dt.getFullYear()}`;

        if (cell instanceof Date && !isNaN(cell)) return fromDate(cell);

        /* Pure number → treat as an Excel serial date (1900 date system). */
        if (typeof cell === 'number' && isFinite(cell)) {
            /* SheetJS exposes the conversion; use it when available, else compute. */
            if (window.XLSX && XLSX.SSF && typeof XLSX.SSF.parse_date_code === 'function') {
                const dc = XLSX.SSF.parse_date_code(cell);
                if (dc && dc.y) return `${pad(dc.d)}.${pad(dc.m)}.${dc.y}`;
            }
            const epoch = new Date(Date.UTC(1899, 11, 30));
            const dt = new Date(epoch.getTime() + cell * 86400000);
            return fromDate(new Date(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate()));
        }

        const s = String(cell ?? '').trim();
        if (!s) return '';
        /* Numeric string that is actually a serial (e.g. "46296"). */
        if (/^\d{4,6}$/.test(s)) return this._xlsxDateToDotStr(parseInt(s, 10));
        return s;   /* already a DD.MM.YYYY / DD/MM/YYYY style string */
    },

    /* ── Data preparation pipeline (spec-driven) ── */
    _prepareDbCpu() {
        const upload = this._uploads.dbcpu;
        if (!upload || !upload.rows || !upload.rows.length) {
            return { error: 'No Load History data is available. Please upload the Load History file in Step 1.' };
        }

        /* 1. Parse dates; keep only rows with a valid date. */
        const dated = [];
        for (const r of upload.rows) {
            const dt = this._parseDotDate(r.date);
            if (dt) dated.push({ ...r, _dt: dt, _ym: dt.getFullYear() * 12 + dt.getMonth() });
        }
        if (!dated.length) {
            return { error: 'Could not parse any Start Date values from the Load History data.' };
        }

        /* 2. Month filtering: drop the current (incomplete) month, sort desc, keep 3. */
        const now = new Date();
        const curYm = now.getFullYear() * 12 + now.getMonth();
        const completed = dated.filter(r => r._ym < curYm);
        completed.sort((a, b) => b._dt - a._dt);   /* descending */
        const currentMonthRemoved = completed.length < dated.length ||
                                     !dated.some(r => r._ym === curYm);
        const top3 = completed.slice(0, 3);
        if (top3.length < 3) {
            return { error: `Only ${top3.length} completed month(s) of Load History are available — at least 3 completed months are required for this analysis.` };
        }

        /* 3. Numeric normalization + 4. calculated columns (all Ø values in ms). */
        const num = v => { const n = this._parseNum(v); return n == null ? 0 : n; };
        const monthRows = top3.map(r => {
            const steps  = num(r.steps);
            const respMs = num(r.resp_ms);
            const cpuMs  = num(r.cpu_ms);
            const dbMs   = num(r.db_ms);
            const respH  = steps * respMs / 1000 / 3600;
            const cpuH   = steps * cpuMs  / 1000 / 3600;
            const dbH    = steps * dbMs   / 1000 / 3600;
            return {
                dt:        r._dt,
                dateLabel: this._fmtMonth(r._dt),
                dateRaw:   String(r.date).replace(/\./g, '/'),
                steps,
                trips:     num(r.roundtrips),
                dbKb:      num(r.db_kb),
                respMs, cpuMs, dbMs,
                respH, cpuH, dbH,
                totalH:    cpuH + dbH,
            };
        });

        /* 5. Averages across the 3 months. */
        const avgCpuH = monthRows.reduce((s, m) => s + m.cpuH, 0) / monthRows.length;
        const avgDbH  = monthRows.reduce((s, m) => s + m.dbH,  0) / monthRows.length;
        const avgRespH = monthRows.reduce((s, m) => s + m.respH, 0) / monthRows.length;

        /* 6. CPU / DB of Response (%) — share of the combined CPU+DB time,
           NOT of the response time. CPU% + DB% ≈ 100%. */
        const totalValue = avgCpuH + avgDbH;
        const cpuPct = totalValue > 0 ? avgCpuH / totalValue * 100 : 0;
        const dbPct  = totalValue > 0 ? avgDbH  / totalValue * 100 : 0;

        /* Ascending order for display (oldest → newest left to right). */
        const monthsAsc = monthRows.slice().sort((a, b) => a.dt - b.dt);

        return {
            monthsAsc,
            avgCpuH, avgDbH, avgRespH, cpuPct, dbPct,
            validation: {
                currentMonthRemoved,
                threeMonths: monthRows.length === 3,
                sortedDesc: true,
                numericsConverted: monthRows.every(m => typeof m.steps === 'number'),
                respCalc: monthRows.every(m => isFinite(m.respH)),
                cpuCalc:  monthRows.every(m => isFinite(m.cpuH)),
                dbCalc:   monthRows.every(m => isFinite(m.dbH)),
                totalCalc: monthRows.every(m => isFinite(m.totalH)),
                avgCalc:  isFinite(avgCpuH) && isFinite(avgDbH),
                pctCalc:  isFinite(cpuPct) && isFinite(dbPct),
            },
        };
    },

    _fmtMonth(dt) {
        const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        return `${months[dt.getMonth()]} ${dt.getFullYear()}`;
    },

    /* ── Summary stat cards ── */
    _renderDbCpuStats(prep, elId = 'aw-dbcpu-stats') {
        const el = document.getElementById(elId);
        if (!el) return;
        const h = n => n.toLocaleString(undefined, { maximumFractionDigits: 1 });
        el.innerHTML = [
            statCard('Months Analysed', prep.monthsAsc.length),
            statCard('Avg CPU Time', h(prep.avgCpuH), ' h/mo'),
            statCard('Avg DB Time', h(prep.avgDbH), ' h/mo'),
            statCard('CPU of Response', prep.cpuPct.toFixed(1), ' %'),
            statCard('DB of Response', prep.dbPct.toFixed(1), ' %'),
        ].join('');
    },

    _showDbCpuError(html) {
        const el = document.getElementById('aw-dbcpu-error');
        const msg = document.getElementById('aw-dbcpu-error-msg');
        if (el && msg) { msg.innerHTML = html; el.style.display = 'flex'; }
    },

    /* ══════════════════════════════════════════════════════
       WORKBOOK PERSISTENCE
    ══════════════════════════════════════════════════════ */

    /* Each data sheet is immediately followed by its chart-data sheet.
       The chart sheet contains the series data formatted for Excel charting
       (ready for "Insert → Chart" without reformatting). */
    _TAB_ORDER: [
        'ST03_DB_vs_CPU_vs_Steps',  'ST03_DB_vs_CPU_Charts',
        'ST03_Monthly_Analysis',    'ST03_Task_Type_Charts',
        'ST03_TP_Dialog',           'ST03_Dialog_Charts',
        'ST03_TP_HTTP_HTTPS',       'ST03_HTTP_HTTPS_Charts',
        'ST03_TP_Background',       'ST03_Background_Charts',
        'ST03_TP_RFC',              'ST03_RFC_Charts',
    ],

    _ensureWorkbook() {
        if (this._wb) return this._wb;
        this._wb = XLSX.utils.book_new();
        /* Pre-create all six tabs (empty placeholders) so the sheet order is fixed. */
        this._TAB_ORDER.forEach(name => {
            const ws = XLSX.utils.aoa_to_sheet([['(pending — populate from the matching screen)']]);
            XLSX.utils.book_append_sheet(this._wb, ws, name);
        });
        return this._wb;
    },

    /* Replace one worksheet's contents (array-of-arrays) without touching others. */
    _writeSheet(name, aoa) {
        const wb = this._ensureWorkbook();
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        if (wb.Sheets[name]) {
            wb.Sheets[name] = ws;                       /* keep position in SheetNames */
            if (!wb.SheetNames.includes(name)) wb.SheetNames.push(name);
        } else {
            XLSX.utils.book_append_sheet(wb, ws, name);
        }
    },

    /* Write the source-data worksheet for the DB-vs-CPU combo chart,
       then immediately write the paired chart-data worksheet. */
    _writeDbCpuSheet(prep) {
        const header = ['Start Date', '# Steps', '# Trips', 'DB Data (KB)',
            'Ø Time (ms)', 'Response Time (Hours)',
            'Ø CPU Time (ms)', 'CPU Time (Hours)',
            'Ø DB Time (ms)', 'DB Time (Hours)',
            'Total CPU+DB Time (Hours)'];
        const r2 = n => Math.round(n * 100) / 100;
        const rows = prep.monthsAsc.map(m => [
            m.dateRaw, m.steps, m.trips, m.dbKb,
            r2(m.respMs), r2(m.respH),
            r2(m.cpuMs), r2(m.cpuH),
            r2(m.dbMs), r2(m.dbH),
            r2(m.totalH),
        ]);
        const aoa = [header, ...rows, [],
            ['Average CPU Time (Hours)', r2(prep.avgCpuH)],
            ['Average DB Time (Hours)',  r2(prep.avgDbH)],
            ['CPU Time (%)',             r2(prep.cpuPct)],
            ['DB Time (%)',              r2(prep.dbPct)],
        ];
        this._writeSheet('ST03_DB_vs_CPU_vs_Steps', aoa);

        /* Chart-data sheet: series ready for "Insert Chart" in Excel.
           Layout: Month | CPU Time (Hours) | DB Time (Hours) | # Steps
           CPU+DB use the left axis (clustered bars); # Steps uses the right (line). */
        const chartHeader = ['Month', 'CPU Time (Hours)', 'DB Time (Hours)', '# Steps'];
        const chartRows   = prep.monthsAsc.map(m => [m.dateLabel, r2(m.cpuH), r2(m.dbH), m.steps]);
        const chartAoa = [
            ['Chart: DB and CPU Time vs No. of Steps'],
            ['Chart type: Clustered Column (CPU + DB hours, left axis) + Line (# Steps, right axis)'],
            [],
            chartHeader,
            ...chartRows,
        ];
        this._writeSheet('ST03_DB_vs_CPU_Charts', chartAoa);
    },

    _sidTag() {
        const d = this.app && this.app.getModuleData && this.app.getModuleData('system-landscape');
        const sid = (d && d.systemInfo && d.systemInfo.sid) || '';
        return String(sid).replace(/[^A-Za-z0-9]/g, '').slice(0, 3).toUpperCase() || 'SAP';
    },

    async _downloadWorkbook() {
        /* Rebuild every data sheet at download time so the workbook is never
           missing content regardless of which screens the user visited. */
        this._rebuildAllSheets();
        if (!this._wb) { showToast('Nothing to download', 'The workbook has not been generated yet.', 'info'); return; }

        if (typeof JSZip !== 'undefined') {
            try {
                const fileName = `${this._sidTag()}_WorkloadAnalysis.xlsx`;
                await this._writeXlsxWithNativeCharts(this._wb, fileName);
                return;
            } catch (err) {
                console.warn('Native chart embedding failed, falling back to plain XLSX:', err);
            }
        }
        XLSX.writeFile(this._wb, `${this._sidTag()}_WorkloadAnalysis.xlsx`);
    },

    /* ══════════════════════════════════════════════════════
       NATIVE EXCEL CHART EMBEDDING
       Each chart sheet gets a proper xl/charts/chartN.xml that references
       the data range on its paired data sheet. Charts are fully editable
       in Excel — no images, no screenshots.
    ══════════════════════════════════════════════════════ */

    /* Descriptor for each chart sheet: data sheet name, row/col layout, chart type. */
    _chartDescriptors() {
        /* These must match exactly what _writeDbCpuSheet / _writeMonthlySheet /
           _writeProfileSheet write.  Row indices are 1-based (Excel rows).
           Layout for all chart sheets:
             Row 1 : chart title text
             Row 2 : chart type note
             Row 3 : (blank or metadata)
             Row 4 : header row  ← firstDataRow = 5
        */
        return [
            {
                chartSheet: 'ST03_DB_vs_CPU_Charts',
                dataSheet:  'ST03_DB_vs_CPU_Charts',
                firstDataRow: 5,           /* row 4 = header, row 5+ = data */
                type: 'comboBarLine',
                title: 'DB and CPU Time vs No. of Steps',
                /* col A = category, col B = CPU bar, col C = DB bar, col D = Steps line */
                barCols:  [{ col: 'B', name: 'CPU Time (Hours)' }, { col: 'C', name: 'DB Time (Hours)' }],
                lineCols: [{ col: 'D', name: '# Steps' }],
                leftAxisTitle: 'Hours per Month',
                rightAxisTitle: '# Steps',
            },
            {
                chartSheet: 'ST03_Task_Type_Charts',
                dataSheet:  'ST03_Task_Type_Charts',
                firstDataRow: 6,           /* rows 1-3 metadata, row 4 blank, row 5 header */
                type: 'clusteredBar',
                title: 'Time per Task Type (Top 5)',
                /* col A = category, col B = CPU bar, col C = DB bar */
                barCols: [{ col: 'B', name: 'CPU Time (Hours)' }, { col: 'C', name: 'DB Time (Hours)' }],
                leftAxisTitle: 'Hours',
            },
            {
                chartSheet: 'ST03_Dialog_Charts',
                dataSheet:  'ST03_Dialog_Charts',
                firstDataRow: 5,
                type: 'clusteredBar',
                title: 'Top 20 Dialog Transactions — CPU vs DB Time',
                barCols: [{ col: 'B', name: 'Total CPU Time (hr)' }, { col: 'C', name: 'Total DB Time (hr)' }],
                leftAxisTitle: 'Time (hr)',
            },
            {
                chartSheet: 'ST03_HTTP_HTTPS_Charts',
                dataSheet:  'ST03_HTTP_HTTPS_Charts',
                firstDataRow: 5,
                type: 'clusteredBar',
                title: 'Top 20 HTTP/HTTPS Transactions — CPU vs DB Time',
                barCols: [{ col: 'B', name: 'Total CPU Time (hr)' }, { col: 'C', name: 'Total DB Time (hr)' }],
                leftAxisTitle: 'Time (hr)',
            },
            {
                chartSheet: 'ST03_Background_Charts',
                dataSheet:  'ST03_Background_Charts',
                firstDataRow: 5,
                type: 'clusteredBar',
                title: 'Top 20 Background Jobs — CPU vs DB Time',
                barCols: [{ col: 'B', name: 'Total CPU Time (hr)' }, { col: 'C', name: 'Total DB Time (hr)' }],
                leftAxisTitle: 'Time (hr)',
            },
            {
                chartSheet: 'ST03_RFC_Charts',
                dataSheet:  'ST03_RFC_Charts',
                firstDataRow: 5,
                type: 'clusteredBar',
                title: 'Top 20 RFC Actions — CPU vs DB Time',
                barCols: [{ col: 'B', name: 'Total CPU Time (hr)' }, { col: 'C', name: 'Total DB Time (hr)' }],
                leftAxisTitle: 'Time (hr)',
            },
        ];
    },

    /* Count actual data rows in a named sheet's array starting at firstDataRow. */
    _countDataRows(wb, sheetName, firstDataRow) {
        const ws = wb.Sheets[sheetName];
        if (!ws || !ws['!ref']) return 0;
        const range = XLSX.utils.decode_range(ws['!ref']);
        /* rows are 0-based in SheetJS range */
        const lastRow1 = range.e.r + 1; /* 1-based */
        return Math.max(0, lastRow1 - firstDataRow + 1);
    },

    async _writeXlsxWithNativeCharts(wb, fileName) {
        /* Step 1: write the workbook to binary */
        const wbBinary = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
        const len = wbBinary.length;
        const buf = new Uint8Array(len);
        for (let i = 0; i < len; i++) buf[i] = wbBinary.charCodeAt(i) & 0xff;

        /* Step 2: open with JSZip */
        const zip = await JSZip.loadAsync(buf);

        /* Step 3: resolve sheet name → file index */
        const sheetIndexMap = await this._resolveSheetIndexMap(zip);

        /* Step 4: inject a native chart for each chart sheet */
        const descriptors = this._chartDescriptors();
        let chartCounter = 0;

        for (const desc of descriptors) {
            const sheetIdx = sheetIndexMap[desc.chartSheet];
            if (!sheetIdx) continue;

            /* Count how many data rows exist — skip if the sheet is empty/pending */
            const nRows = this._countDataRows(wb, desc.dataSheet, desc.firstDataRow);
            if (nRows < 1) continue;

            chartCounter++;
            const cId = chartCounter;

            const chartPath   = `xl/charts/chart${cId}.xml`;
            const drawingPath = `xl/drawings/drawing${cId}.xml`;
            const drawingRel  = `xl/drawings/_rels/drawing${cId}.xml.rels`;
            const chartRel    = `xl/charts/_rels/chart${cId}.xml.rels`;
            const sheetPath    = `xl/worksheets/sheet${sheetIdx}.xml`;
            const sheetRelPath = `xl/worksheets/_rels/sheet${sheetIdx}.xml.rels`;

            /* Build the chart XML */
            const lastDataRow = desc.firstDataRow + nRows - 1;
            zip.file(chartPath, this._makeNativeChartXml(desc, lastDataRow, cId));

            /* Drawing XML: graphicFrame pointing at the chart (not a picture) */
            zip.file(drawingPath, this._makeChartDrawingXml(cId));

            /* drawing → chart relationship */
            zip.file(drawingRel, this._makeChartDrawingRels(cId));

            /* chart relationships (empty — no style/colors file required) */
            zip.file(chartRel, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);

            /* Patch worksheet to reference drawing */
            const sheetXml = await zip.file(sheetPath)?.async('string');
            if (sheetXml) zip.file(sheetPath, this._patchSheetXml(sheetXml, 'rId1'));

            /* Patch / create worksheet relationships */
            const existingRel = await zip.file(sheetRelPath)?.async('string');
            zip.file(sheetRelPath, this._patchSheetRels(existingRel, cId, `../drawings/drawing${cId}.xml`));
        }

        if (chartCounter === 0) {
            XLSX.writeFile(wb, fileName);
            return;
        }

        /* Step 5: patch [Content_Types].xml */
        const ctRaw = await zip.file('[Content_Types].xml').async('string');
        zip.file('[Content_Types].xml', this._patchContentTypesForCharts(ctRaw, chartCounter));

        /* Step 6: generate and download */
        const outBuf = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
        const blob = new Blob([outBuf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = fileName;
        a.click();
        setTimeout(() => URL.revokeObjectURL(a.href), 10000);
    },

    /* Build the xl/charts/chartN.xml OOXML for one chart descriptor. */
    _makeNativeChartXml(desc, lastDataRow, cId) {
        const ds  = desc.dataSheet;          /* sheet name for formula refs */
        const fdr = desc.firstDataRow;       /* first data row (1-based) */
        const ldr = lastDataRow;

        /* Excel formula: 'SheetName'!$A$fdr:$A$ldr */
        const q = n => `'${n.replace(/'/g, "''")}'`;
        const ref = (col, r1, r2) => `${q(ds)}!$${col}$${r1}:$${col}$${r2}`;

        /* Category (X) labels — always column A */
        const catRef = ref('A', fdr, ldr);

        /* ── Clustered bar series ── */
        const barSeries = desc.barCols.map((bc, i) => `
        <c:ser>
            <c:idx val="${i}"/>
            <c:order val="${i}"/>
            <c:tx><c:strRef><c:f>${q(ds)}!$${bc.col}$${fdr - 1}</c:f></c:strRef></c:tx>
            <c:cat><c:strRef><c:f>${catRef}</c:f></c:strRef></c:cat>
            <c:val><c:numRef><c:f>${ref(bc.col, fdr, ldr)}</c:f></c:numRef></c:val>
        </c:ser>`).join('');

        /* ── Line series (only for comboBarLine) ── */
        const lineSeries = (desc.lineCols || []).map((lc, i) => `
        <c:ser>
            <c:idx val="${i}"/>
            <c:order val="${i}"/>
            <c:tx><c:strRef><c:f>${q(ds)}!$${lc.col}$${fdr - 1}</c:f></c:strRef></c:tx>
            <c:smooth val="0"/>
            <c:cat><c:strRef><c:f>${catRef}</c:f></c:strRef></c:cat>
            <c:val><c:numRef><c:f>${ref(lc.col, fdr, ldr)}</c:f></c:numRef></c:val>
        </c:ser>`).join('');

        const titleXml = `
        <c:title>
            <c:tx><c:rich>
                <a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
                <a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
                <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:r><a:t>${this._xmlEsc(desc.title)}</a:t></a:r>
                </a:p>
            </c:rich></c:tx>
            <c:overlay val="0"/>
        </c:title>`;

        const leftAxisTitle = desc.leftAxisTitle
            ? `<c:title><c:tx><c:rich>
                <a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
                <a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
                <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:r><a:t>${this._xmlEsc(desc.leftAxisTitle)}</a:t></a:r>
                </a:p>
            </c:rich></c:tx><c:overlay val="0"/></c:title>` : '';

        const rightAxisTitle = desc.rightAxisTitle
            ? `<c:title><c:tx><c:rich>
                <a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
                <a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
                <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:r><a:t>${this._xmlEsc(desc.rightAxisTitle)}</a:t></a:r>
                </a:p>
            </c:rich></c:tx><c:overlay val="0"/></c:title>` : '';

        let plotAreaXml;
        if (desc.type === 'comboBarLine') {
            plotAreaXml = `
            <c:barChart>
                <c:barDir val="col"/>
                <c:grouping val="clustered"/>
                <c:varyColors val="0"/>
                ${barSeries}
                <c:axId val="1001"/>
                <c:axId val="1002"/>
            </c:barChart>
            <c:lineChart>
                <c:grouping val="standard"/>
                <c:varyColors val="0"/>
                ${lineSeries}
                <c:marker><c:symbol val="circle"/><c:size val="4"/></c:marker>
                <c:smooth val="0"/>
                <c:axId val="1001"/>
                <c:axId val="1003"/>
            </c:lineChart>
            <c:catAx>
                <c:axId val="1001"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="b"/>
                <c:crossAx val="1002"/>
            </c:catAx>
            <c:valAx>
                <c:axId val="1002"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="l"/>
                ${leftAxisTitle}
                <c:crossAx val="1001"/>
            </c:valAx>
            <c:valAx>
                <c:axId val="1003"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="r"/>
                ${rightAxisTitle}
                <c:crossAx val="1001"/>
            </c:valAx>`;
        } else {
            plotAreaXml = `
            <c:barChart>
                <c:barDir val="col"/>
                <c:grouping val="clustered"/>
                <c:varyColors val="0"/>
                ${barSeries}
                <c:axId val="1001"/>
                <c:axId val="1002"/>
            </c:barChart>
            <c:catAx>
                <c:axId val="1001"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="b"/>
                <c:crossAx val="1002"/>
                <c:tickLblPos val="low"/>
            </c:catAx>
            <c:valAx>
                <c:axId val="1002"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="l"/>
                ${leftAxisTitle}
                <c:crossAx val="1001"/>
            </c:valAx>`;
        }

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <c:autoTitleDeleted val="0"/>
    ${titleXml}
    <c:plotArea>
        ${plotAreaXml}
    </c:plotArea>
    <c:legend>
        <c:legendPos val="b"/>
        <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
</c:chartSpace>`;
    },

    _xmlEsc(s) {
        return String(s == null ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    },

    /* Drawing XML: graphicFrame embedding a chart (not a picture). */
    _makeChartDrawingXml(cId) {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>15</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>29</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
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

    /* drawing → chart relationship */
    _makeChartDrawingRels(cId) {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${cId}.xml"/>
</Relationships>`;
    },

    /* Patch [Content_Types].xml to register chart and drawing part types. */
    _patchContentTypesForCharts(xml, chartCount) {
        let patched = xml;

        for (let i = 1; i <= chartCount; i++) {
            const chartPart   = `/xl/charts/chart${i}.xml`;
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

    /* Full mapping: reads workbook.xml and workbook.xml.rels, returns { sheetName: fileIndex }. */
    async _resolveSheetIndexMap(zip) {
        const wbXml   = await zip.file('xl/workbook.xml').async('string');
        const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');

        /* Parse workbook.xml: extract name → rId */
        const nameToRid = {};
        const re1 = /<sheet\b[^>]*>/g;
        let m;
        while ((m = re1.exec(wbXml)) !== null) {
            const tag = m[0];
            const nameM = /\bname="([^"]*)"/.exec(tag);
            const ridM  = /\br:id="([^"]*)"/.exec(tag);
            if (nameM && ridM) {
                const name = nameM[1].replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'");
                nameToRid[name] = ridM[1];
            }
        }

        /* Parse workbook.xml.rels: rId → Target (e.g. worksheets/sheet3.xml) */
        const ridToTarget = {};
        const re2 = /<Relationship\b[^>]*>/g;
        while ((m = re2.exec(relsXml)) !== null) {
            const tag = m[0];
            const idM  = /\bId="([^"]*)"/.exec(tag);
            const tgtM = /\bTarget="([^"]*)"/.exec(tag);
            if (idM && tgtM) ridToTarget[idM[1]] = tgtM[1];
        }

        /* Build final map: sheetName → numeric index extracted from "worksheets/sheetN.xml" */
        const result = {};
        Object.entries(nameToRid).forEach(([name, rid]) => {
            const target = ridToTarget[rid] || '';
            const idxM = /worksheets\/sheet(\d+)\.xml/i.exec(target);
            if (idxM) result[name] = parseInt(idxM[1]);
        });
        return result;
    },

    /* Generate the DrawingML XML that anchors an image starting at cell A1. */
    /* Patch a worksheet XML string to add a <drawing> reference before </worksheet>. */
    _patchSheetXml(xml, drawingRid) {
        const drawingTag = `<drawing r:id="${drawingRid}"/>`;
        let patched = xml;
        if (!patched.includes('xmlns:r=')) {
            patched = patched.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
        }
        patched = patched.replace(/<drawing\s[^>]*\/>/g, '').replace(/<drawing\/>/g, '');
        if (patched.includes('</worksheet>')) {
            patched = patched.replace('</worksheet>', `${drawingTag}</worksheet>`);
        } else {
            patched += drawingTag;
        }
        return patched;
    },

    /* Create or patch a sheet relationships XML to add a drawing relationship. */
    _patchSheetRels(existingXml, _id, drawingTarget) {
        const newRel = `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="${drawingTarget}"/>`;
        if (!existingXml) {
            return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${newRel}
</Relationships>`;
        }
        let patched = existingXml.replace(/<Relationship[^>]*Type="[^"]*\/drawing"[^>]*\/>/g, '');
        if (patched.includes('</Relationships>')) {
            patched = patched.replace('</Relationships>', `  ${newRel}\n</Relationships>`);
        } else {
            patched += newRel;
        }
        return patched;
    },

    /* Idempotently (re)populate all analysis worksheets from current uploads. */
    _rebuildAllSheets() {
        try {
            const prep = this._dbcpuPrepared || (() => { const p = this._prepareDbCpu(); return (p && !p.error) ? p : null; })();
            if (prep) { this._dbcpuPrepared = prep; this._writeDbCpuSheet(prep); }
            const ma = this._prepareMonthlyAnalysis();
            if (ma && !ma.error) this._writeMonthlySheet(ma);
            this._writeAllProfileSheets();
        } catch (err) {
            console.warn('workbook rebuild failed', err);
        }
    },

    /* ══════════════════════════════════════════════════════
       COMBINATION CHART (clustered columns + line, dual axis)
       with editable title, editable per-axis min/max, and
       Copy / PNG / SVG / CSV / Reset controls.
    ══════════════════════════════════════════════════════ */
    _mountComboChart(container, prep) {
        if (!container) return;
        const uid = 'aw-combo';
        const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
        const C = this._COMBO_COLORS;

        const labels = prep.monthsAsc.map(m => m.dateLabel);
        const cpuH   = prep.monthsAsc.map(m => Math.round(m.cpuH * 100) / 100);
        const dbH    = prep.monthsAsc.map(m => Math.round(m.dbH  * 100) / 100);
        const steps  = prep.monthsAsc.map(m => m.steps);

        container.innerHTML = `
            <div class="card chart-card-wrap">
                <div class="chart-toolbar">
                    <input class="chart-title-input" id="${uid}-title" value="DB and CPU Time vs No. of Steps" title="Edit chart title">
                    <div class="chart-controls">
                        <label>Hours min <input type="number" id="${uid}-hmin" value="0" step="1" style="width:68px"></label>
                        <label>Hours max <input type="number" id="${uid}-hmax" placeholder="auto" step="1" style="width:68px"></label>
                        <label>Steps min <input type="number" id="${uid}-smin" value="0" step="1" style="width:72px"></label>
                        <label>Steps max <input type="number" id="${uid}-smax" placeholder="auto" step="1" style="width:72px"></label>
                        <button class="btn btn-sm" id="${uid}-reset" title="Reset axes">Reset</button>
                        <button class="btn btn-sm" id="${uid}-copy" title="Copy a high-resolution image to the clipboard">Copy</button>
                        <button class="btn btn-sm" id="${uid}-png" title="Download a PNG image">PNG</button>
                        <button class="btn btn-sm" id="${uid}-svg" title="Download an SVG image">SVG</button>
                        <button class="btn btn-sm" id="${uid}-csv" title="Download the chart data as CSV">CSV</button>
                    </div>
                </div>
                <div class="chart-canvas-wrap" style="height:380px"><canvas id="${uid}-canvas"></canvas></div>
                ${this._fontSizeToolbarHtml(uid)}
                <div class="chart-hint">Columns = CPU &amp; DB time in hours per month (left axis) · line = dialog steps (right axis). Title and axis ranges are editable.</div>
            </div>`;

        const ctx = container.querySelector(`#${uid}-canvas`).getContext('2d');
        const chart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels,
                datasets: [
                    { type: 'bar',  label: 'CPU Time (Hours)', data: cpuH, backgroundColor: C.cpu, borderColor: C.cpu, yAxisID: 'yHours', order: 2 },
                    { type: 'bar',  label: 'DB Time (Hours)',  data: dbH,  backgroundColor: C.db,  borderColor: C.db,  yAxisID: 'yHours', order: 2 },
                    { type: 'line', label: '# Steps',          data: steps, borderColor: C.steps, backgroundColor: C.steps, yAxisID: 'ySteps', order: 1, tension: 0.3, pointRadius: 4, borderWidth: 2.5 },
                ],
            },
            options: {
                responsive: true, maintainAspectRatio: false, animation: false,
                plugins: {
                    legend: { position: 'bottom', labels: { font: { size: 11 } } },
                    tooltip: { enabled: true },
                },
                scales: {
                    x: { grid: { display: false }, ticks: { font: { size: 11 } } },
                    yHours: { type: 'linear', position: 'left', min: 0, beginAtZero: true,
                              title: { display: true, text: 'Hours per Month' }, grid: { color: '#EEF2F7' } },
                    ySteps: { type: 'linear', position: 'right', min: 0, beginAtZero: true,
                              title: { display: true, text: 'Steps' }, grid: { drawOnChartArea: false } },
                },
            },
        });
        this._charts.push(chart);
        this._comboChart = chart;
        this._comboCsv = { labels, cpuH, dbH, steps };

        /* Title editing (updates the live title text used by exports). */
        this._comboTitle = 'DB and CPU Time vs No. of Steps';
        container.querySelector(`#${uid}-title`).addEventListener('input', e => { this._comboTitle = e.target.value; });

        /* Axis range editing. */
        const applyAxes = () => {
            const n = id => { const v = parseFloat(container.querySelector(`#${uid}-${id}`).value); return isNaN(v) ? undefined : v; };
            chart.options.scales.yHours.min = n('hmin');
            chart.options.scales.yHours.max = n('hmax');
            chart.options.scales.ySteps.min = n('smin');
            chart.options.scales.ySteps.max = n('smax');
            chart.update();
        };
        ['hmin','hmax','smin','smax'].forEach(id =>
            container.querySelector(`#${uid}-${id}`).addEventListener('change', applyAxes));

        container.querySelector(`#${uid}-reset`).addEventListener('click', () => {
            container.querySelector(`#${uid}-hmin`).value = '0';
            container.querySelector(`#${uid}-hmax`).value = '';
            container.querySelector(`#${uid}-smin`).value = '0';
            container.querySelector(`#${uid}-smax`).value = '';
            chart.options.scales.yHours.min = 0; chart.options.scales.yHours.max = undefined;
            chart.options.scales.ySteps.min = 0; chart.options.scales.ySteps.max = undefined;
            chart.update();
        });

        container.querySelector(`#${uid}-copy`).addEventListener('click', () => this._comboCopyPng());
        container.querySelector(`#${uid}-png`).addEventListener('click', () => this._comboExportPng());
        container.querySelector(`#${uid}-svg`).addEventListener('click', () => this._comboExportSvg());
        container.querySelector(`#${uid}-csv`).addEventListener('click', () => this._comboExportCsv());

        this._wireFontSizeControls(container, [chart], false);
    },

    /* ══════════════════════════════════════════════════════
       REUSABLE EDITABLE COMBO-CHART FACTORY
       ──────────────────────────────────────────────────────
       Clustered columns (left axis) + optional line series (right axis),
       with the SAME interaction model as the existing combo chart:
         editable title · per-axis min/max · show/hide X labels ·
         Copy / PNG / SVG / CSV / Reset · responsive · resizable.
       Each instance is fully self-contained (keyed by `uid`), so many
       charts can live on one screen.

       cfg = {
         uid, title, labels:[],
         bars:  [{ label, data:[], color }],     // left (hours) axis
         line:  { label, data:[], color, max } | null,   // right (%) axis
         leftTitle, rightTitle,
         csv:   { headers:[], rows:[[...]] },
         hint,
       }
    ══════════════════════════════════════════════════════ */
    _comboRegistry: {},

    _mountEditableCombo(container, cfg) {
        if (!container) return;
        const uid = cfg.uid;
        const hasLine = !!cfg.line;
        const leftTitle  = cfg.leftTitle  || 'Hours';
        const rightTitle = cfg.rightTitle || 'Percentage (%)';

        container.innerHTML = `
            <div class="card chart-card-wrap">
                <div class="chart-toolbar">
                    <input class="chart-title-input" id="${uid}-title" value="${this._escAttr(cfg.title)}" title="Edit chart title">
                    <div class="chart-controls">
                        <label>Y min <input type="number" id="${uid}-lmin" value="0" step="1" style="width:64px"></label>
                        <label>Y max <input type="number" id="${uid}-lmax" placeholder="auto" step="1" style="width:64px"></label>
                        ${hasLine ? `<label>% min <input type="number" id="${uid}-rmin" value="0" step="1" style="width:60px"></label>
                        <label>% max <input type="number" id="${uid}-rmax" value="${cfg.line.max != null ? cfg.line.max : 100}" step="1" style="width:60px"></label>` : ''}
                        <label class="chart-toggle"><input type="checkbox" id="${uid}-xlabels" checked> X labels</label>
                        <button class="btn btn-sm" id="${uid}-reset" title="Reset axes">Reset</button>
                        <button class="btn btn-sm" id="${uid}-copy" title="Copy a high-resolution image to the clipboard">Copy</button>
                        <button class="btn btn-sm" id="${uid}-png" title="Download a PNG image">PNG</button>
                        <button class="btn btn-sm" id="${uid}-svg" title="Download an SVG image">SVG</button>
                        <button class="btn btn-sm" id="${uid}-csv" title="Download the chart data as CSV">CSV</button>
                    </div>
                </div>
                <div class="chart-canvas-wrap" style="height:380px"><canvas id="${uid}-canvas"></canvas></div>
                ${this._fontSizeToolbarHtml(uid)}
                ${cfg.hint ? `<div class="chart-hint">${cfg.hint}</div>` : ''}
            </div>`;

        const C = this._COMBO_COLORS;
        const datasets = cfg.bars.map(b => ({
            type: 'bar', label: b.label, data: b.data,
            backgroundColor: b.color, borderColor: b.color,
            yAxisID: 'yLeft', order: 2,
            ...(cfg.stacked ? { stack: 'bars' } : {}),
        }));
        if (hasLine) {
            datasets.push({
                type: 'line', label: cfg.line.label, data: cfg.line.data,
                borderColor: cfg.line.color || C.steps, backgroundColor: cfg.line.color || C.steps,
                yAxisID: 'yRight', order: 1, tension: 0.3, pointRadius: 3, borderWidth: 2.5,
            });
        }

        const scales = {
            x: { grid: { display: false }, ticks: { maxRotation: 60, minRotation: 45, font: { size: 10 } },
                 ...(cfg.stacked ? { stacked: true } : {}) },
            yLeft: { type: 'linear', position: 'left', min: 0, beginAtZero: true,
                     title: { display: true, text: leftTitle }, grid: { color: '#EEF2F7' },
                     ...(cfg.stacked ? { stacked: true } : {}) },
        };
        if (hasLine) {
            scales.yRight = { type: 'linear', position: 'right', min: 0,
                              max: cfg.line.max != null ? cfg.line.max : 100, beginAtZero: true,
                              title: { display: true, text: rightTitle }, grid: { drawOnChartArea: false } };
        }

        const ctx = container.querySelector(`#${uid}-canvas`).getContext('2d');
        const chart = new Chart(ctx, {
            type: 'bar',
            data: { labels: cfg.labels, datasets },
            options: {
                responsive: true, maintainAspectRatio: false, animation: false,
                plugins: { legend: { position: 'bottom', labels: { font: { size: 11 } } }, tooltip: { enabled: true } },
                scales,
            },
        });
        this._charts.push(chart);

        const entry = { chart, cfg, title: cfg.title };
        this._comboRegistry[uid] = entry;

        /* Title */
        container.querySelector(`#${uid}-title`).addEventListener('input', e => { entry.title = e.target.value; });

        /* Axis editing */
        const numVal = id => { const el = container.querySelector(`#${uid}-${id}`); if (!el) return undefined; const v = parseFloat(el.value); return isNaN(v) ? undefined : v; };
        const applyAxes = () => {
            chart.options.scales.yLeft.min = numVal('lmin');
            chart.options.scales.yLeft.max = numVal('lmax');
            if (hasLine) { chart.options.scales.yRight.min = numVal('rmin'); chart.options.scales.yRight.max = numVal('rmax'); }
            chart.update();
        };
        ['lmin','lmax','rmin','rmax'].forEach(id => { const el = container.querySelector(`#${uid}-${id}`); if (el) el.addEventListener('change', applyAxes); });

        /* Show/hide X-axis labels */
        container.querySelector(`#${uid}-xlabels`).addEventListener('change', e => {
            chart.options.scales.x.ticks.display = e.target.checked;
            chart.update();
        });

        /* Reset */
        container.querySelector(`#${uid}-reset`).addEventListener('click', () => {
            const set = (id, v) => { const el = container.querySelector(`#${uid}-${id}`); if (el) el.value = v; };
            set('lmin', '0'); set('lmax', '');
            if (hasLine) { set('rmin', '0'); set('rmax', cfg.line.max != null ? cfg.line.max : 100); }
            const xl = container.querySelector(`#${uid}-xlabels`); if (xl) xl.checked = true;
            chart.options.scales.yLeft.min = 0; chart.options.scales.yLeft.max = undefined;
            if (hasLine) { chart.options.scales.yRight.min = 0; chart.options.scales.yRight.max = cfg.line.max != null ? cfg.line.max : 100; }
            chart.options.scales.x.ticks.display = true;
            chart.update();
        });

        /* Export actions reuse the generic renderer. */
        const base = `${this._sidTag()}_${uid}`;
        container.querySelector(`#${uid}-copy`).addEventListener('click', () => this._comboImage(uid, 'copy', base));
        container.querySelector(`#${uid}-png`).addEventListener('click',  () => this._comboImage(uid, 'png',  base));
        container.querySelector(`#${uid}-svg`).addEventListener('click',  () => this._comboImage(uid, 'svg',  base));
        container.querySelector(`#${uid}-csv`).addEventListener('click',  () => this._comboCsvExport(uid, base));

        this._wireFontSizeControls(container, [chart], false);

        /* Section 2 requirement: render the exact plotted Top-N source data
           (same measures + sort order) directly below the chart. */
        if (cfg.sourceTable && cfg.csv) {
            const t = document.createElement('div');
            t.innerHTML = this._sourceTableHtml(cfg.csv, cfg.sourceCaption || 'Source data (plotted rows)');
            container.appendChild(t);
        }

        return entry;
    },

    /* Build a source-data table (used below Section-2 charts). Columns and row
       order mirror the chart's CSV payload exactly. */
    _sourceTableHtml(csv, caption) {
        const fmt = v => typeof v === 'number'
            ? v.toLocaleString('en-US', { maximumFractionDigits: 2 })
            : (v == null ? '' : String(v));
        return `
            <div class="card" style="padding:0;margin-top:12px;margin-bottom:16px">
                <div class="card-header" style="padding:12px 18px"><div class="card-title">${caption}</div></div>
                <div class="table-wrap" style="border:none;border-radius:0 0 14px 14px;overflow-x:auto">
                    <table class="data-table">
                        <thead><tr>${csv.headers.map((h, i) => `<th${i ? ' style="text-align:right"' : ''}>${h}</th>`).join('')}</tr></thead>
                        <tbody>${csv.rows.map(r => `<tr>${r.map((c, i) => `<td${i ? ' style="text-align:right"' : ''}>${fmt(c)}</td>`).join('')}</tr>`).join('')}</tbody>
                    </table>
                </div>
            </div>`;
    },

    _escAttr(s) {
        return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
    },

    /* ══════════════════════════════════════════════════════
       CHART FONT SIZE CONTROLS
       Delegates to the global ChartFontSizeManager + helpers in app.js.
    ══════════════════════════════════════════════════════ */

    /* Returns the font-size control row HTML (placed after the canvas). */
    _fontSizeToolbarHtml(uid) {
        return chartFontRowHtml(uid);
    },

    /* Push current font sizes onto a Chart.js instance and re-render. */
    _applyFontSizesToChart(chart, isPie) {
        applyChartFontSizes(chart, isPie);
    },

    /* Wire font-size buttons for a chart. containerEl is the card/wrapper element. */
    _wireFontSizeControls(containerEl, charts, _isPie) {
        const getAllRows = () => {
            const rows = [];
            document.querySelectorAll('[id$="-font-row"]').forEach(r => rows.push(r));
            return rows;
        };
        wireChartFontRow(containerEl, () => {
            /* Collect all currently mounted charts (array path + combo registry). */
            const all = [];
            if (this._charts && this._charts.length) this._charts.forEach(c => all.push(c));
            return all.length ? all : (charts || []);
        }, getAllRows, null);
    },

    /* Add one font-size control row to an auxiliary section container (Section 3).
       The row is placed at the bottom of the container, after all chart grids.
       uid must be unique per section. Charts are collected from this._charts at wire time. */
    _addSectionFontRow(container, uid) {
        if (!container) return;
        const rowId = `${uid}-font-row`;
        const existing = document.getElementById(rowId);
        if (existing) existing.remove();
        container.insertAdjacentHTML('beforeend', chartFontRowHtml(uid));
        const getAllRows = () => Array.from(document.querySelectorAll('[id$="-font-row"]'));
        wireChartFontRow(container, () => Array.from(this._charts || []).filter(Boolean), getAllRows, null);
    },

    /* Render a registered combo chart (title band + canvas) to a hi-res canvas. */
    _comboRegRenderCanvas(uid, scale = 3) {
        const entry = this._comboRegistry[uid];
        if (!entry) return null;
        const src = entry.chart.canvas;
        const titleH = 42;
        const canvas = document.createElement('canvas');
        canvas.width = src.width * scale;
        canvas.height = src.height * scale + titleH * scale;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.fillStyle = '#1F3864';
        ctx.font = `bold ${(ChartFontSizeManager.get().title || 18) * scale}px Calibri, Arial, sans-serif`;
        /* Title: top-center (horizontally centered over the chart area). */
        ctx.textBaseline = 'top';
        ctx.textAlign = 'center';
        ctx.fillText(entry.title || '', canvas.width / 2, 10 * scale);
        ctx.textAlign = 'left';
        ctx.drawImage(src, 0, titleH * scale, src.width * scale, src.height * scale);
        return canvas;
    },

    _comboImage(uid, mode, base) {
        const canvas = this._comboRegRenderCanvas(uid, 3);
        if (!canvas) return;
        if (mode === 'svg') {
            const w = canvas.width, h = canvas.height;
            const png = canvas.toDataURL('image/png');
            const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w / 3}" height="${h / 3}" viewBox="0 0 ${w} ${h}">` +
                        `<image width="${w}" height="${h}" href="${png}"/></svg>`;
            this._downloadBlob(new Blob([svg], { type: 'image/svg+xml' }), `${base}.svg`, 'SVG downloaded.');
            return;
        }
        canvas.toBlob(blob => {
            if (!blob) { showToast('Export failed', 'Could not render the chart image.', 'error'); return; }
            if (mode === 'copy' && navigator.clipboard && window.ClipboardItem) {
                try { window.focus(); } catch (_) {}
                navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })])
                    .then(() => showToast('Chart copied', 'Chart image copied — paste it into Word or PowerPoint.', 'success'))
                    .catch(() => this._downloadBlob(blob, `${base}.png`, 'Copy blocked — downloaded a PNG instead.'));
            } else {
                this._downloadBlob(blob, `${base}.png`, 'PNG downloaded.');
            }
        }, 'image/png');
    },

    _comboCsvExport(uid, base) {
        const entry = this._comboRegistry[uid];
        if (!entry || !entry.cfg.csv) return;
        const { headers, rows } = entry.cfg.csv;
        const esc = v => typeof v === 'string' ? `"${v.replace(/"/g, '""')}"` : v;
        const lines = [headers.map(esc).join(',')];
        rows.forEach(r => lines.push(r.map(esc).join(',')));
        this._downloadBlob(new Blob([lines.join('\n')], { type: 'text/csv' }), `${base}.csv`, 'CSV downloaded.');
    },

    /* Render the chart (title + canvas) onto a white canvas at hi-res. */
    _comboRenderCanvas(scale = 3) {
        const chart = this._comboChart;
        if (!chart) return null;
        const src = chart.canvas;
        const titleH = 42;
        const canvas = document.createElement('canvas');
        canvas.width = src.width * scale;
        canvas.height = src.height * scale + titleH * scale;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.fillStyle = '#1F3864';
        ctx.font = `bold ${(ChartFontSizeManager.get().title || 18) * scale}px Calibri, Arial, sans-serif`;
        /* Title: top-center (horizontally centered over the chart area). */
        ctx.textBaseline = 'top';
        ctx.textAlign = 'center';
        ctx.fillText(this._comboTitle || '', canvas.width / 2, 10 * scale);
        ctx.textAlign = 'left';
        ctx.drawImage(src, 0, titleH * scale, src.width * scale, src.height * scale);
        return canvas;
    },

    _comboCopyPng() {
        const canvas = this._comboRenderCanvas(3);
        if (!canvas) return;
        const base = `${this._sidTag()}_DB_vs_CPU_vs_Steps`;
        canvas.toBlob(blob => {
            if (!blob) { showToast('Copy failed', 'Could not render the chart image.', 'error'); return; }
            if (navigator.clipboard && window.ClipboardItem) {
                try { window.focus(); } catch (_) {}
                navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })])
                    .then(() => showToast('Chart copied', 'Chart image copied — paste it into Word or PowerPoint.', 'success'))
                    .catch(() => this._downloadBlob(blob, `${base}.png`, 'Copy blocked — downloaded a PNG instead.'));
            } else {
                this._downloadBlob(blob, `${base}.png`, 'Clipboard not supported — downloaded a PNG instead.');
            }
        }, 'image/png');
    },

    _comboExportPng() {
        const canvas = this._comboRenderCanvas(3);
        if (!canvas) return;
        canvas.toBlob(blob => {
            if (blob) this._downloadBlob(blob, `${this._sidTag()}_DB_vs_CPU_vs_Steps.png`, 'PNG downloaded.');
        }, 'image/png');
    },

    /* Minimal standalone SVG (embeds the raster so it drops into PowerPoint). */
    _comboExportSvg() {
        const canvas = this._comboRenderCanvas(3);
        if (!canvas) return;
        const w = canvas.width, h = canvas.height;
        const png = canvas.toDataURL('image/png');
        const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w / 3}" height="${h / 3}" viewBox="0 0 ${w} ${h}">` +
                    `<image width="${w}" height="${h}" href="${png}"/></svg>`;
        this._downloadBlob(new Blob([svg], { type: 'image/svg+xml' }), `${this._sidTag()}_DB_vs_CPU_vs_Steps.svg`, 'SVG downloaded.');
    },

    _comboExportCsv() {
        const d = this._comboCsv;
        if (!d) return;
        const lines = [['Start Date', 'CPU Time (Hours)', 'DB Time (Hours)', '# Steps'].join(',')];
        d.labels.forEach((lbl, i) => lines.push([`"${lbl}"`, d.cpuH[i], d.dbH[i], d.steps[i]].join(',')));
        this._downloadBlob(new Blob([lines.join('\n')], { type: 'text/csv' }),
            `${this._sidTag()}_DB_vs_CPU_vs_Steps.csv`, 'CSV downloaded.');
    },

    _downloadBlob(blob, name, toastMsg) {
        const url = URL.createObjectURL(blob);
        const a = Object.assign(document.createElement('a'), { href: url, download: name });
        document.body.appendChild(a); a.click(); document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        if (toastMsg) showToast('Chart saved', toastMsg, 'success');
    },

    /* ══════════════════════════════════════════════════════
       DB-vs-CPU SUMMARY TABLE + EXTENDED TREND OBSERVATIONS
       Covers the last 3 available months:
         • # Steps trend (increase / decrease / stable)
         • CPU and DB consumption growth vs Steps growth
         • Anomaly detection (month-over-month spikes)
         • Combined narrative observations
    ══════════════════════════════════════════════════════ */
    _renderDbCpuSummary(container, prep) {
        if (!container) return;
        const r1 = n => (Math.round(n * 10) / 10).toLocaleString(undefined, { maximumFractionDigits: 1 });
        const pct = n => (n >= 0 ? '+' : '') + r1(n) + ' %';

        const months = prep.monthsAsc;   /* [oldest, middle, newest] */
        const n = months.length;         /* always 3 */

        /* ── 1. Steps trend ── */
        const steps = months.map(m => m.steps);
        const stepsFirst = steps[0] || 1, stepsLast = steps[n - 1] || 0;
        const stepsPctChange = (stepsLast - stepsFirst) / stepsFirst * 100;
        let stepsTrend, stepsBadge;
        if (stepsPctChange > 10)       { stepsTrend = 'Increasing'; stepsBadge = 'badge-yellow'; }
        else if (stepsPctChange < -10) { stepsTrend = 'Decreasing'; stepsBadge = 'badge-green'; }
        else                           { stepsTrend = 'Stable';     stepsBadge = 'badge-gray';  }

        /* ── 2. CPU trend ── */
        const cpuFirst = months[0].cpuH || 0, cpuLast = months[n - 1].cpuH || 0;
        const cpuPctChange = cpuFirst > 0 ? (cpuLast - cpuFirst) / cpuFirst * 100 : 0;

        /* ── 3. DB trend ── */
        const dbFirst = months[0].dbH || 0, dbLast = months[n - 1].dbH || 0;
        const dbPctChange = dbFirst > 0 ? (dbLast - dbFirst) / dbFirst * 100 : 0;

        /* ── 4. Month-over-month changes for anomaly detection ── */
        const momSteps = [], momCpu = [], momDb = [];
        for (let i = 1; i < n; i++) {
            const prev = months[i - 1], curr = months[i];
            momSteps.push({ from: prev.dateLabel, to: curr.dateLabel,
                            chg: prev.steps > 0 ? (curr.steps - prev.steps) / prev.steps * 100 : 0 });
            momCpu.push({ from: prev.dateLabel, to: curr.dateLabel,
                          chg: prev.cpuH > 0 ? (curr.cpuH - prev.cpuH) / prev.cpuH * 100 : 0 });
            momDb.push({ from: prev.dateLabel, to: curr.dateLabel,
                         chg: prev.dbH > 0 ? (curr.dbH - prev.dbH) / prev.dbH * 100 : 0 });
        }
        const anomalyThreshold = 25;   /* % change considered significant */

        /* ── 5. Build observation bullets ── */
        const obs = [];

        /* Steps trend narrative */
        if (stepsPctChange > 10) {
            obs.push(`<strong># Steps trend: Increasing</strong> — dialog step volume grew by <strong>${r1(stepsPctChange)}%</strong> from ${months[0].dateLabel} to ${months[n - 1].dateLabel}, indicating a growing transactional workload. Capacity planning should account for continued expansion.`);
        } else if (stepsPctChange < -10) {
            obs.push(`<strong># Steps trend: Decreasing</strong> — dialog step volume fell by <strong>${r1(Math.abs(stepsPctChange))}%</strong> from ${months[0].dateLabel} to ${months[n - 1].dateLabel}, suggesting a contracting workload. Current capacity may be over-provisioned.`);
        } else {
            obs.push(`<strong># Steps trend: Stable</strong> — dialog step volume varied by only <strong>${r1(Math.abs(stepsPctChange))}%</strong> across the three months, indicating a steady transactional workload.`);
        }

        /* CPU vs Steps comparison */
        const cpuStepsDiff = Math.abs(cpuPctChange - stepsPctChange);
        if (cpuStepsDiff > 15) {
            if (cpuPctChange > stepsPctChange) {
                obs.push(`<strong>CPU time growing faster than steps</strong> — CPU time increased by <strong>${r1(cpuPctChange)}%</strong> while steps grew by ${r1(stepsPctChange)}%, a divergence of ${r1(cpuStepsDiff)} percentage points. This may indicate increased processing complexity per step, or background jobs being captured in the statistics.`);
            } else {
                obs.push(`<strong>CPU time growing slower than steps</strong> — steps grew by <strong>${r1(stepsPctChange)}%</strong> while CPU time changed by only ${r1(cpuPctChange)}%, suggesting efficiency gains or a shift towards lighter workload types.`);
            }
        } else if (Math.abs(cpuPctChange) > 5) {
            obs.push(`<strong>CPU time</strong> changed by <strong>${pct(cpuPctChange)}</strong> over the period, broadly in line with the ${r1(stepsPctChange)}% change in dialog steps.`);
        }

        /* DB vs Steps comparison */
        const dbStepsDiff = Math.abs(dbPctChange - stepsPctChange);
        if (dbStepsDiff > 15) {
            if (dbPctChange > stepsPctChange) {
                obs.push(`<strong>DB time growing faster than steps</strong> — DB time increased by <strong>${r1(dbPctChange)}%</strong> while steps grew by ${r1(stepsPctChange)}%, a divergence of ${r1(dbStepsDiff)} percentage points. This may indicate table growth, missing indexes, or increased data volumes impacting query performance.`);
            } else {
                obs.push(`<strong>DB time growing slower than steps</strong> — steps grew by <strong>${r1(stepsPctChange)}%</strong> while DB time changed by only ${r1(dbPctChange)}%, suggesting improved database performance or reduced data complexity.`);
            }
        } else if (Math.abs(dbPctChange) > 5) {
            obs.push(`<strong>DB time</strong> changed by <strong>${pct(dbPctChange)}</strong> over the period, broadly consistent with the ${r1(stepsPctChange)}% change in dialog steps.`);
        }

        /* Month-over-month anomaly detection */
        const anomalies = [];
        momSteps.forEach(m => { if (Math.abs(m.chg) >= anomalyThreshold) anomalies.push(`# Steps ${m.from}→${m.to}: ${pct(m.chg)}`); });
        momCpu.forEach(m => { if (Math.abs(m.chg) >= anomalyThreshold) anomalies.push(`CPU Time ${m.from}→${m.to}: ${pct(m.chg)}`); });
        momDb.forEach(m => { if (Math.abs(m.chg) >= anomalyThreshold) anomalies.push(`DB Time ${m.from}→${m.to}: ${pct(m.chg)}`); });
        if (anomalies.length > 0) {
            obs.push(`<strong>Significant month-over-month variation detected</strong> — ${anomalies.join('; ')}. Investigate whether these spikes reflect one-off events (e.g., period-end processing, data migration) or a structural shift in workload.`);
        }

        /* CPU/DB split observation */
        if (prep.dbPct > 60) {
            obs.push(`<strong>DB-bound workload</strong> — DB time accounts for <strong>${r1(prep.dbPct)}%</strong> of combined CPU+DB processing time. Database performance (indexing, statistics, query plans) has a dominant impact on overall response times.`);
        } else if (prep.cpuPct > 60) {
            obs.push(`<strong>CPU-bound workload</strong> — CPU time accounts for <strong>${r1(prep.cpuPct)}%</strong> of combined CPU+DB processing time. Application-level optimisation (ABAP code, buffering) will have the greatest impact on performance.`);
        }

        /* ── 6. Month-by-month detail table ── */
        const detailRows = months.map((m, i) => {
            const prevSteps = i > 0 ? months[i - 1].steps : null;
            const prevCpu   = i > 0 ? months[i - 1].cpuH  : null;
            const prevDb    = i > 0 ? months[i - 1].dbH   : null;
            const fmtChg = (curr, prev) => prev == null ? '—' :
                (prev > 0 ? `<span style="color:${curr > prev ? 'var(--brand)' : curr < prev ? '#888' : 'inherit'}">${pct((curr - prev) / prev * 100)}</span>` : '—');
            return `<tr>
                <td>${m.dateLabel}</td>
                <td style="text-align:right">${m.steps.toLocaleString()}</td>
                <td style="text-align:right">${fmtChg(m.steps, prevSteps)}</td>
                <td style="text-align:right">${r1(m.cpuH)}</td>
                <td style="text-align:right">${fmtChg(m.cpuH, prevCpu)}</td>
                <td style="text-align:right">${r1(m.dbH)}</td>
                <td style="text-align:right">${fmtChg(m.dbH, prevDb)}</td>
            </tr>`;
        }).join('');

        container.innerHTML = `
            <div class="section-title" style="margin-top:24px">Summary</div>
            <div class="card" style="padding:0;margin-bottom:16px">
                <div class="table-wrap" style="border:none;border-radius:14px;overflow-x:auto">
                    <table class="data-table">
                        <thead><tr><th>Metric</th><th style="text-align:right">Value</th></tr></thead>
                        <tbody>
                            <tr><td>Response Time (Hours)</td><td style="text-align:right">${r1(prep.avgRespH)}</td></tr>
                            <tr><td>Average CPU Time (Hours)</td><td style="text-align:right">${r1(prep.avgCpuH)}</td></tr>
                            <tr><td>Average DB Time (Hours)</td><td style="text-align:right">${r1(prep.avgDbH)}</td></tr>
                            <tr><td>CPU Time (%)</td><td style="text-align:right">${r1(prep.cpuPct)} %</td></tr>
                            <tr><td>DB Time (%)</td><td style="text-align:right">${r1(prep.dbPct)} %</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>

            <div class="section-title" style="margin-top:24px">Month-by-Month Comparison</div>
            <div class="card" style="padding:0;margin-bottom:16px">
                <div class="table-wrap" style="border:none;border-radius:14px;overflow-x:auto">
                    <table class="data-table">
                        <thead><tr>
                            <th>Month</th>
                            <th style="text-align:right"># Steps</th>
                            <th style="text-align:right">MoM</th>
                            <th style="text-align:right">CPU Time (h)</th>
                            <th style="text-align:right">MoM</th>
                            <th style="text-align:right">DB Time (h)</th>
                            <th style="text-align:right">MoM</th>
                        </tr></thead>
                        <tbody>${detailRows}</tbody>
                    </table>
                </div>
            </div>

            <div class="section-title" style="margin-top:24px">Observations</div>
            <div style="display:flex;flex-direction:column;gap:8px;margin-bottom:16px">
                ${obs.map(o => `
                <div class="alert-info" style="align-items:flex-start">
                    <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:2px">
                        <polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>
                    </svg>
                    <div>${o}</div>
                </div>`).join('')}
            </div>`;

        this._dbcpuSummaryTrend = { trend: stepsTrend, pctChange: stepsPctChange, cpuPctChange, dbPctChange };
    },

    /* ══════════════════════════════════════════════════════
       MONTHLY ANALYSIS (Workload According to Task Type)
       Produces: Top-5 CPU+DB pie, Time-per-Task-Type clustered
       column chart, supporting tables, observations, and the
       ST03_Monthly_Analysis worksheet.
    ══════════════════════════════════════════════════════ */
    _prepareMonthlyAnalysis() {
        const upload = this._uploads.tasktype;
        if (!upload || !upload.rows || !upload.rows.length) {
            return { error: 'Upload the "Workload According to Task Type" file in Step 1 to generate the Monthly Analysis.' };
        }

        /* Reporting period from the upload meta ("Period: MM/YYYY"). */
        let period = '';
        const pm = String(upload.meta || '').match(/(\d{1,2}\/\d{4})/);
        if (pm) period = pm[1];

        /* Build per-task-type rows with hours (all Ø values are ms). */
        const num = v => { const n = (typeof v === 'number') ? v : this._parseNum(v); return (n == null || !isFinite(n)) ? 0 : n; };
        const rows = upload.rows
            .filter(r => (r.task_type || '').toString().trim() !== '')
            .map(r => {
                const steps = num(r.steps);
                const respH = steps * num(r.resp_ms) / 1000 / 3600;
                const cpuH  = steps * num(r.cpu_ms)  / 1000 / 3600;
                const dbH   = steps * num(r.db_ms)   / 1000 / 3600;
                return {
                    task: String(r.task_type).trim(),
                    steps, respH, cpuH, dbH, totalH: cpuH + dbH,
                };
            });

        /* Sort descending by Total CPU+DB Hours. */
        rows.sort((a, b) => b.totalH - a.totalH);

        const grandTotalH = rows.reduce((s, r) => s + r.totalH, 0);
        const top5 = rows.slice(0, 5);
        const top5TotalH = top5.reduce((s, r) => s + r.totalH, 0);
        const pieRows = top5.map(r => ({
            task: r.task,
            totalH: r.totalH,
            pct: top5TotalH > 0 ? r.totalH / top5TotalH * 100 : 0,
        }));

        /* Time-per-Task-Type totals + Application/Database split (%). */
        const tempCpu = rows.reduce((s, r) => s + r.cpuH, 0);
        const tempDb  = rows.reduce((s, r) => s + r.dbH,  0);
        const tempCpuDb = tempCpu + tempDb;
        const appPct = tempCpuDb > 0 ? Math.round(tempCpu / tempCpuDb * 1000) / 10 : 0;
        const dbPct  = tempCpuDb > 0 ? Math.round(tempDb  / tempCpuDb * 1000) / 10 : 0;

        return {
            period, rows, grandTotalH, top5, top5TotalH, pieRows,
            tempCpu, tempDb, tempCpuDb, appPct, dbPct,
        };
    },

    _mountTaskTypePie(container, ma) {
        if (!container) return;
        const uid = 'aw-ma-pie-c';
        const r1 = n => (Math.round(n * 10) / 10).toLocaleString(undefined, { maximumFractionDigits: 1 });
        const palette = ['#4472C4', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646', '#2C4D75', '#772C2A'];

        const labels = ma.pieRows.map(p => p.task);
        const values = ma.pieRows.map(p => Math.round(p.pct * 10) / 10);

        /* Observations. */
        const largest = ma.pieRows[0];
        const top5OfGrand = ma.grandTotalH > 0 ? ma.top5TotalH / ma.grandTotalH * 100 : 0;
        const concentration = top5OfGrand >= 90 ? 'highly concentrated'
                            : top5OfGrand >= 70 ? 'moderately concentrated' : 'broadly distributed';
        const dominant = largest && largest.pct >= 50
            ? `<strong>${largest.task}</strong> alone accounts for ${r1(largest.pct)}% of the Top-5 CPU+DB workload — a disproportionately high share that should be the primary focus for optimisation and sizing.`
            : 'No single task type dominates the Top-5 workload disproportionately.';

        const periodLbl = ma.period ? ` — ${ma.period}` : '';

        container.innerHTML = `
            <div class="card chart-card-wrap">
                <div class="chart-toolbar">
                    <input class="chart-title-input" id="${uid}-title" value="Workload According to Task Type (CPU+DB Time per Month)${periodLbl}" title="Edit chart title">
                    <div class="chart-controls">
                        <button class="btn btn-sm" id="${uid}-copy" title="Copy a high-resolution image to the clipboard">Copy</button>
                        <button class="btn btn-sm" id="${uid}-png" title="Download a PNG image">PNG</button>
                        <button class="btn btn-sm" id="${uid}-csv" title="Download the chart data as CSV">CSV</button>
                    </div>
                </div>
                <div class="chart-canvas-wrap" style="height:340px"><canvas id="${uid}-canvas"></canvas></div>
                ${this._fontSizeToolbarHtml(uid)}
                <div class="chart-hint">${ma.period ? `Reporting period: <strong>${ma.period}</strong>. ` : ''}Each slice = a Top-5 task type's share of the combined CPU+DB time of the five largest task types.</div>
            </div>

            <div class="card" style="padding:0;margin-top:14px">
                <div class="card-header" style="padding:12px 18px"><div class="card-title">Top-5 Supporting Table</div></div>
                <div class="table-wrap" style="border:none;border-radius:0 0 14px 14px;overflow-x:auto">
                    <table class="data-table">
                        <thead><tr><th>Task Type</th><th style="text-align:right">Total CPU+DB Time (Hours)</th><th style="text-align:right">CPU+DB Time (%)</th></tr></thead>
                        <tbody>
                            ${ma.pieRows.map(p => `<tr><td>${p.task}</td><td style="text-align:right">${r1(p.totalH)}</td><td style="text-align:right">${r1(p.pct)} %</td></tr>`).join('')}
                            <tr style="font-weight:600;border-top:2px solid var(--border-strong)">
                                <td>Top 5 Total</td><td style="text-align:right">${r1(ma.top5TotalH)}</td><td style="text-align:right">100.0 %</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>

            <div class="alert-info" style="margin-top:14px">
                <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                     fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                    <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
                <div>
                    <strong>Observations</strong><br>
                    • Largest workload contributor: <strong>${largest ? largest.task : '—'}</strong> (${largest ? r1(largest.pct) : '0'}% of the Top-5 CPU+DB total).<br>
                    • The Top-5 task types together represent <strong>${r1(top5OfGrand)}%</strong> of the overall CPU+DB workload (grand total ${r1(ma.grandTotalH)} h).<br>
                    • Workload concentration: <strong>${concentration}</strong>.<br>
                    • ${dominant}
                </div>
            </div>`;

        const ctx = container.querySelector(`#${uid}-canvas`).getContext('2d');
        /* Inline plugin: draw each slice's % contribution at its centroid
           (no external datalabels plugin is loaded). */
        const sliceLabels = {
            id: 'awPieLabels',
            afterDatasetsDraw(ch) {
                const g = ch.ctx;
                const meta = ch.getDatasetMeta(0);
                const data = ch.data.datasets[0].data;
                g.save();
                g.font = '600 12px Calibri, Arial, sans-serif';
                g.textAlign = 'center';
                g.textBaseline = 'middle';
                meta.data.forEach((arc, i) => {
                    const v = data[i];
                    if (v == null || v <= 0) return;
                    const p = arc.tooltipPosition();
                    const txt = `${v}%`;
                    /* white halo for contrast over any slice colour */
                    g.lineWidth = 3; g.strokeStyle = 'rgba(255,255,255,0.85)';
                    g.strokeText(txt, p.x, p.y);
                    g.fillStyle = '#1F2937';
                    g.fillText(txt, p.x, p.y);
                });
                g.restore();
            },
        };
        const chart = new Chart(ctx, {
            type: 'pie',
            data: { labels, datasets: [{ data: values, backgroundColor: labels.map((_, i) => palette[i % palette.length]), borderColor: '#fff', borderWidth: 2 }] },
            options: {
                responsive: true, maintainAspectRatio: false,
                layout: { padding: 8 },
                plugins: {
                    legend: { position: 'right', labels: { font: { size: 11 } } },
                    tooltip: { callbacks: { label: c => `${c.label}: ${c.parsed}%` } },
                },
            },
            plugins: [sliceLabels],
        });
        this._charts.push(chart);
        this._maPie = { chart, labels, values, title: `Workload According to Task Type (CPU+DB Time per Month)${periodLbl}` };

        container.querySelector(`#${uid}-title`).addEventListener('input', e => { this._maPie.title = e.target.value; });
        container.querySelector(`#${uid}-copy`).addEventListener('click', () => this._exportSimpleChart(this._maPie, 'copy', `${this._sidTag()}_TaskType_Pie`));
        container.querySelector(`#${uid}-png`).addEventListener('click', () => this._exportSimpleChart(this._maPie, 'png', `${this._sidTag()}_TaskType_Pie`));
        container.querySelector(`#${uid}-csv`).addEventListener('click', () => {
            const lines = [['Task Type', 'CPU+DB Time (%)', 'Total CPU+DB Time (Hours)'].join(',')];
            ma.pieRows.forEach(p => lines.push([`"${p.task}"`, Math.round(p.pct * 10) / 10, Math.round(p.totalH * 10) / 10].join(',')));
            this._downloadBlob(new Blob([lines.join('\n')], { type: 'text/csv' }), `${this._sidTag()}_TaskType_Pie.csv`, 'CSV downloaded.');
        });

        this._wireFontSizeControls(container, [chart], true);
    },

    _mountTimePerTaskType(container, ma) {
        if (!container) return;
        const uid = 'aw-ma-tpt';
        const r1 = n => (Math.round(n * 10) / 10).toLocaleString(undefined, { maximumFractionDigits: 1 });
        const C = this._COMBO_COLORS;
        const periodLbl = ma.period ? ` — ${ma.period}` : '';

        /* Plot only the Top-5 largest task types (rows are sorted desc by Total CPU+DB). */
        const top5 = ma.rows.slice(0, 5);
        const labels = top5.map(r => r.task);
        const cpuH   = top5.map(r => Math.round(r.cpuH * 10) / 10);
        const dbH    = top5.map(r => Math.round(r.dbH  * 10) / 10);

        container.innerHTML = `
            <div class="card chart-card-wrap">
                <div class="chart-toolbar">
                    <input class="chart-title-input" id="${uid}-title" value="Time per Task Type${periodLbl}" title="Edit chart title">
                    <div class="chart-controls">
                        <label>Y min <input type="number" id="${uid}-ymin" value="0" step="1" style="width:68px"></label>
                        <label>Y max <input type="number" id="${uid}-ymax" placeholder="auto" step="1" style="width:68px"></label>
                        <button class="btn btn-sm" id="${uid}-reset" title="Reset axes">Reset</button>
                        <button class="btn btn-sm" id="${uid}-copy" title="Copy a high-resolution image to the clipboard">Copy</button>
                        <button class="btn btn-sm" id="${uid}-png" title="Download a PNG image">PNG</button>
                        <button class="btn btn-sm" id="${uid}-csv" title="Download the chart data as CSV">CSV</button>
                    </div>
                </div>
                <div class="chart-canvas-wrap" style="height:360px"><canvas id="${uid}-canvas"></canvas></div>
                ${this._fontSizeToolbarHtml(uid)}
                <div class="chart-hint">${ma.period ? `Reporting period: <strong>${ma.period}</strong>. ` : ''}Clustered columns — CPU vs DB processing time (hours) for the Top-5 task types.</div>
            </div>

            <div class="alert-info" style="margin-top:14px">
                <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
                     fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                    <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
                <div>
                    <strong>Application Time (%): ${r1(ma.appPct)} %</strong> &nbsp;·&nbsp;
                    <strong>Database Time (%): ${r1(ma.dbPct)} %</strong><br>
                    Of the total CPU+DB processing time (${r1(ma.tempCpuDb)} h), ${r1(ma.appPct)}% is spent in the
                    application (CPU) tier and ${r1(ma.dbPct)}% in the database tier.
                </div>
            </div>

            <div class="card" style="padding:0;margin-top:14px">
                <div class="card-header" style="padding:12px 18px"><div class="card-title">Time per Task Type — Source Table (Top 5)</div></div>
                <div class="table-wrap" style="border:none;border-radius:0 0 14px 14px;overflow-x:auto">
                    <table class="data-table">
                        <thead><tr><th>Task Type</th><th style="text-align:right">CPU Time (Hours)</th><th style="text-align:right">DB Time (Hours)</th><th style="text-align:right">Total CPU+DB Time (Hours)</th></tr></thead>
                        <tbody>
                            ${top5.map(r => `<tr><td>${r.task}</td><td style="text-align:right">${r1(r.cpuH)}</td><td style="text-align:right">${r1(r.dbH)}</td><td style="text-align:right">${r1(r.totalH)}</td></tr>`).join('')}
                            <tr style="font-weight:600;border-top:2px solid var(--border-strong)">
                                <td>Top 5 Total</td><td style="text-align:right">${r1(top5.reduce((s, r) => s + r.cpuH, 0))}</td><td style="text-align:right">${r1(top5.reduce((s, r) => s + r.dbH, 0))}</td><td style="text-align:right">${r1(top5.reduce((s, r) => s + r.totalH, 0))}</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>`;

        const ctx = container.querySelector(`#${uid}-canvas`).getContext('2d');
        const chart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels,
                datasets: [
                    { label: 'CPU Time (Hours)', data: cpuH, backgroundColor: C.cpu, borderColor: C.cpu },
                    { label: 'DB Time (Hours)',  data: dbH,  backgroundColor: C.db,  borderColor: C.db },
                ],
            },
            options: {
                responsive: true, maintainAspectRatio: false, animation: false,
                plugins: { legend: { position: 'bottom', labels: { font: { size: 11 } } } },
                scales: {
                    x: { grid: { display: false }, ticks: { maxRotation: 35, font: { size: 11 } }, title: { display: true, text: 'Task Type' } },
                    y: { beginAtZero: true, min: 0, grid: { color: '#EEF2F7' }, title: { display: true, text: 'Hours' } },
                },
            },
        });
        this._charts.push(chart);
        this._maTpt = { chart, labels, cpuH, dbH, title: `Time per Task Type${periodLbl}` };

        container.querySelector(`#${uid}-title`).addEventListener('input', e => { this._maTpt.title = e.target.value; });
        const applyY = () => {
            const n = id => { const v = parseFloat(container.querySelector(`#${uid}-${id}`).value); return isNaN(v) ? undefined : v; };
            chart.options.scales.y.min = n('ymin'); chart.options.scales.y.max = n('ymax'); chart.update();
        };
        ['ymin', 'ymax'].forEach(id => container.querySelector(`#${uid}-${id}`).addEventListener('change', applyY));
        container.querySelector(`#${uid}-reset`).addEventListener('click', () => {
            container.querySelector(`#${uid}-ymin`).value = '0';
            container.querySelector(`#${uid}-ymax`).value = '';
            chart.options.scales.y.min = 0; chart.options.scales.y.max = undefined; chart.update();
        });
        container.querySelector(`#${uid}-copy`).addEventListener('click', () => this._exportSimpleChart(this._maTpt, 'copy', `${this._sidTag()}_TimePerTaskType`));
        container.querySelector(`#${uid}-png`).addEventListener('click', () => this._exportSimpleChart(this._maTpt, 'png', `${this._sidTag()}_TimePerTaskType`));
        container.querySelector(`#${uid}-csv`).addEventListener('click', () => {
            /* CSV mirrors the plotted Top-5 clustered chart. */
            const lines = [['Task Type', 'CPU Time (Hours)', 'DB Time (Hours)', 'Total CPU+DB Time (Hours)'].join(',')];
            top5.forEach(r => lines.push([`"${r.task}"`, r1(r.cpuH), r1(r.dbH), r1(r.totalH)].join(',')));
            this._downloadBlob(new Blob([lines.join('\n')], { type: 'text/csv' }), `${this._sidTag()}_TimePerTaskType.csv`, 'CSV downloaded.');
        });

        this._wireFontSizeControls(container, [chart], false);
    },

    /* Shared Copy/PNG for the pie + clustered charts (title drawn above canvas). */
    _exportSimpleChart(entry, mode, base) {
        if (!entry || !entry.chart) return;
        const scale = 3, titleH = 42;
        const src = entry.chart.canvas;
        const canvas = document.createElement('canvas');
        canvas.width = src.width * scale; canvas.height = src.height * scale + titleH * scale;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.fillStyle = '#1F3864'; ctx.font = `bold ${(ChartFontSizeManager.get().title || 18) * scale}px Calibri, Arial, sans-serif`;
        ctx.textBaseline = 'top';
        ctx.fillText(entry.title || '', 10 * scale, 10 * scale);
        ctx.drawImage(src, 0, titleH * scale, src.width * scale, src.height * scale);
        canvas.toBlob(blob => {
            if (!blob) return;
            if (mode === 'copy' && navigator.clipboard && window.ClipboardItem) {
                try { window.focus(); } catch (_) {}
                navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })])
                    .then(() => showToast('Chart copied', 'Chart image copied — paste it into Word or PowerPoint.', 'success'))
                    .catch(() => this._downloadBlob(blob, `${base}.png`, 'Copy blocked — downloaded a PNG instead.'));
            } else {
                this._downloadBlob(blob, `${base}.png`, 'PNG downloaded.');
            }
        }, 'image/png');
    },

    /* Write the ST03_Monthly_Analysis worksheet (source tables + KPIs),
       then immediately write the paired Task_Type_Charts worksheet. */
    _writeMonthlySheet(ma) {
        const r2 = n => Math.round(n * 100) / 100;
        const aoa = [];
        aoa.push(['Reporting Period', ma.period || '(not found)']);
        aoa.push([]);
        /* Full per-task-type table (sorted desc by Total CPU+DB Hours). */
        aoa.push(['Task Type', 'CPU Time (Hours)', 'DB Time (Hours)', 'Total CPU+DB Time (Hours)',
                  'Response Time (Hours)', '# Steps']);
        ma.rows.forEach(r => aoa.push([r.task, r2(r.cpuH), r2(r.dbH), r2(r.totalH), r2(r.respH), r.steps]));
        aoa.push(['Grand Total CPU+DB Time (Hours)', '', '', r2(ma.grandTotalH)]);
        aoa.push([]);
        /* Top-5 supporting table (pie source). */
        aoa.push(['Top 5 — Task Type', 'Total CPU+DB Time (Hours)', 'CPU+DB Time (%)']);
        ma.pieRows.forEach(p => aoa.push([p.task, r2(p.totalH), r2(p.pct)]));
        aoa.push(['Top 5 Total CPU+DB Time (Hours)', r2(ma.top5TotalH), 100]);
        aoa.push([]);
        /* Time-per-task-type KPIs. */
        aoa.push(['temp_total_cpu (Hours)', r2(ma.tempCpu)]);
        aoa.push(['temp_total_db (Hours)', r2(ma.tempDb)]);
        aoa.push(['temp_total_cpu_db (Hours)', r2(ma.tempCpuDb)]);
        aoa.push(['Application Time (%)', ma.appPct]);
        aoa.push(['Database Time (%)', ma.dbPct]);
        this._writeSheet('ST03_Monthly_Analysis', aoa);

        /* Chart-data sheet — Top-5 clustered bar (CPU+DB) ready for charting. */
        const top5 = ma.rows.slice(0, 5);
        const chartAoa = [
            ['Chart: Time per Task Type (Top 5)'],
            ['Chart type: Clustered Column — CPU Time (Hours) + DB Time (Hours) per Task Type'],
            ['Reporting Period', ma.period || '(not found)'],
            [],
            ['Task Type', 'CPU Time (Hours)', 'DB Time (Hours)', 'Total CPU+DB Time (Hours)'],
            ...top5.map(r => [r.task, r2(r.cpuH), r2(r.dbH), r2(r.totalH)]),
        ];
        this._writeSheet('ST03_Task_Type_Charts', chartAoa);
    },

    /* ══════════════════════════════════════════════════════
       TRANSACTION-PROFILE WORKBOOK SHEETS
       ST03_TP_Dialog / ST03_TP_HTTP_HTTPS / ST03_TP_Background / ST03_TP_RFC
       Each sheet is written from the derived dataset (sorted DESC by
       Total CPU+DB hr with the accumulation column), with the computed
       columns placed per the ST03_TP_* specification.
    ══════════════════════════════════════════════════════ */
    _r2(n) { return Math.round((n || 0) * 100) / 100; },

    _writeProfileSheet(key) {
        const u = this._uploads[key];
        if (!u || !u.derived) return;
        const r2 = v => this._r2(v);
        const d = u.derived.rows;

        if (key === 'rfc') {
            const header = [
                'System ID', 'Application Server', 'Action in Triggering System', 'User ID', '# Steps',
                'T CPU~ (s)', 'Total CPU Time (hr)',
                'T DB Time (s)', 'Total DB Time (hr)',
                'Total CPU+DB Time (hr)', 'Total CPU+DB Time (%)', 'Acc. Total CPU+DB Time (%)',
            ];
            const rows = d.map(r => [
                r.system_id, r.app_server || '', r.action || '', r.user_id || '', r.steps ?? '',
                r2(r.cpu_total_s), r2(r.cpuH),
                r2(r.db_total_s), r2(r.dbH),
                r2(r.cpuDbH), r2(r.cpuDbPct), r2(r.accPct),
            ]);
            this._writeSheet('ST03_TP_RFC', [header, ...rows, [],
                ['var_total_cpu_db (hr)', r2(u.derived.varTotal)]]);

            /* Chart-data sheet for RFC — Top-20 by CPU+DB. */
            const top20rfc = d.slice(0, 20);
            const rfcChartAoa = [
                ['Chart: Top 20 RFC Actions — CPU vs DB Time'],
                ['Chart type: Clustered Column — Total CPU Time (hr) + Total DB Time (hr) per RFC Action'],
                [],
                ['System ID / Action', 'Total CPU Time (hr)', 'Total DB Time (hr)', 'Total CPU+DB Time (hr)', '# Steps'],
                ...top20rfc.map(r => [`${r.system_id || ''}${r.action ? ' / ' + r.action : ''}`, r2(r.cpuH), r2(r.dbH), r2(r.cpuDbH), r.steps ?? '']),
            ];
            this._writeSheet('ST03_RFC_Charts', rfcChartAoa);
            return;
        }

        /* dialog / http / background */
        const isBg = key === 'background';
        const header = isBg
            ? ['Report or Transaction name', 'Name of Background Job', '# Steps',
               'T CPU~ (s)', 'Total CPU Time (hr)', 'T DB Time (s)', 'Total DB Time (hr)',
               'Total CPU+DB Time (hr)', 'Total CPU+DB Time (%)', 'Acc. Total CPU+DB Time (%)']
            : ['Report or Transaction name', '# Steps',
               'T CPU~ (s)', 'Total CPU Time (hr)', 'Ø CPU~ (ms)', 'Ø CPU (s)',
               'T DB Time (s)', 'Total DB Time (hr)', 'Ø DB Time (ms)', 'Ø DB Time (s)',
               'Total CPU+DB Time (hr)', 'Total CPU+DB Time (%)', 'Acc. Total CPU+DB Time (%)',
               'Ø Roll Wait Time (ms)', 'Ø Roll Wait Time (s)',
               'Ø WaitTime (ms)', 'Ø WaitTime (s)', 'Ø Other Time (s)'];

        const rows = d.map(r => isBg
            ? [r.trans_rep || '', r.job_name || '', r.steps ?? '',
               r2(r.cpu_total_s), r2(r.cpuH), r2(r.db_total_s), r2(r.dbH),
               r2(r.cpuDbH), r2(r.cpuDbPct), r2(r.accPct)]
            : [r.trans_rep || '', r.steps ?? '',
               r2(r.cpu_total_s), r2(r.cpuH), r2(r.cpu_ms), r2(r.cpuS),
               r2(r.db_total_s), r2(r.dbH), r2(r.db_ms), r2(r.dbS),
               r2(r.cpuDbH), r2(r.cpuDbPct), r2(r.accPct),
               r2(r.rollwait_ms), r2(r.rollWaitS),
               r2(r.wait_ms), r2(r.waitS), r2(r.otherS)]);

        const sheet = { dialog: 'ST03_TP_Dialog', http: 'ST03_TP_HTTP_HTTPS', background: 'ST03_TP_Background' }[key];
        this._writeSheet(sheet, [header, ...rows, [],
            ['var_total_cpu_db (hr)', r2(u.derived.varTotal)]]);

        /* Chart-data sheet — Top-20 sorted by Total CPU+DB Time (hr), ready for charting. */
        const chartSheet = { dialog: 'ST03_Dialog_Charts', http: 'ST03_HTTP_HTTPS_Charts', background: 'ST03_Background_Charts' }[key];
        const chartLabel = { dialog: 'Dialog', http: 'HTTP/HTTPS', background: 'Background' }[key];
        const top20 = d.slice(0, 20);
        const chartAoa = [
            [`Chart: Top 20 ${chartLabel} Transactions — CPU vs DB Time`],
            ['Chart type: Clustered Column — Total CPU Time (hr) + Total DB Time (hr) per Transaction/Report'],
            [],
            ['Report / Transaction', 'Total CPU Time (hr)', 'Total DB Time (hr)', 'Total CPU+DB Time (hr)', '# Steps'],
            ...top20.map(r => [r.trans_rep || '', r2(r.cpuH), r2(r.dbH), r2(r.cpuDbH), r.steps ?? '']),
        ];
        this._writeSheet(chartSheet, chartAoa);
    },

    /* Write every available profile sheet (idempotent). */
    _writeAllProfileSheets() {
        ['dialog', 'http', 'background', 'rfc'].forEach(k => {
            if (this._uploads[k] && this._uploads[k].derived) this._writeProfileSheet(k);
        });
    },

    /* ── Step 2 ── */
    _renderStep2() {
        return `
        <div class="module-wrap">
            <div class="module-header">
                <div class="module-crumb">Analysis › Section 04 › Step 3 of 4</div>
                <h1 class="module-title">Application Workload – Transaction Profile</h1>
                <p class="module-desc">
                    Upload transaction profiles for the task types relevant to your landscape.
                    Export each from <strong>ST03 / ST03N → Workload → Month</strong>,
                    filtered by the respective task type.
                    At least one upload is required to proceed.
                </p>
            </div>

            <div id="aw-validation-error" style="display:none" class="alert-error">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24"
                     fill="none" stroke="currentColor" stroke-width="2" style="flex-shrink:0;margin-top:1px">
                    <circle cx="12" cy="12" r="10"/>
                    <line x1="15" y1="9" x2="9" y2="15"/>
                    <line x1="9" y1="9" x2="15" y2="15"/>
                </svg>
                <div id="aw-validation-msg"></div>
            </div>

            ${this._uploadCard('aw-dialog',     'dialog',
                'Dialog Workload',
                `<strong>ST03 / ST03N</strong> → Workload → Total → Month → select month →
                 Analysis view → Transaction Profile → Standard → Task Type: <em>Dialog</em>`,
                'aw-dialog-status')}

            ${this._uploadCard('aw-http',       'http',
                'HTTP / HTTPS Workload',
                `<strong>ST03 / ST03N</strong> → Workload → Total → Month → select month →
                 Analysis view → Transaction Profile → Standard → Task Type: <em>HTTP/HTTPS</em>`,
                'aw-http-status')}

            ${this._uploadCard('aw-background', 'background',
                'Background Workload',
                `<strong>ST03 / ST03N</strong> → Workload → Total → Month → select month →
                 Analysis view → Transaction Profile → Standard → Task Type: <em>Background</em>`,
                'aw-background-status')}

            ${this._uploadCard('aw-rfc',        'rfc',
                'RFC Workload',
                `<strong>ST03 / ST03N</strong> → Workload → Total → Month → select month →
                 Analysis view → Load from External Systems → Aggregation: <em>System – Action – User</em> →
                 Task Type: <em>ALL</em>`,
                'aw-rfc-status')}

            <div class="action-row" style="margin-top:32px;justify-content:space-between">
                <button class="btn" id="aw-btn-back2">
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2.5">
                        <polyline points="15 18 9 12 15 6"/>
                    </svg>
                    Back
                </button>
                <button class="btn btn-blue" id="aw-btn-step3" disabled>
                    Create Charts
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2.5">
                        <polyline points="9 18 15 12 9 6"/>
                    </svg>
                </button>
            </div>
        </div>`;
    },

    /* ── Step 3 ── */
    _renderStep3() {
        const u = this._uploads;
        const section = (title, id, meta) => `
            <div class="chart-group" id="aw-cg-${id}">
                <div class="chart-group-head">
                    <span class="chart-group-title">${title}</span>
                    ${meta ? `<span class="chart-group-count" style="margin-left:12px;color:var(--text-secondary);font-size:12px">${meta}</span>` : ''}
                </div>
                <div class="chart-group-body" id="aw-cb-${id}"></div>
            </div>`;

        /* Collapsible wrapper used by the three top-level sections. */
        const wlSection = (id, title, collapsed, bodyHtml) => `
            <div class="wl-section${collapsed ? ' collapsed' : ''}" id="aw-wl-${id}">
                <div class="wl-section-head">
                    <span class="wl-section-title">${title}</span>
                    <span class="wl-section-actions">
                        <button class="btn btn-sm" data-wl-show="${id}">Show</button>
                        <button class="btn btn-sm" data-wl-hide="${id}">Hide</button>
                    </span>
                </div>
                <div class="wl-section-body">${bodyHtml}</div>
            </div>`;

        const hasTotalWl = u.dbcpu || u.tasktype;
        const hasProfile = u.dialog || u.http || u.background || u.rfc;

        /* SECTION 1 — Total View (expanded): the rich analysis charts with
           their calculations, supporting tables and commentary. */
        const totalBody = hasTotalWl ? `
            ${u.dbcpu ? `
            <div class="section-title">DB and CPU Time vs No. of Steps</div>
            <div class="stats-grid" id="aw-s1-stats" style="margin-bottom:16px"></div>
            <div id="aw-s1-combo"></div>
            <div id="aw-s1-summary"></div>` : ''}
            ${u.tasktype ? `
            <div class="section-title" style="margin-top:28px">Workload According to Task Type (CPU+DB Time per Month)</div>
            <div id="aw-s1-pie"></div>
            <div class="section-title" style="margin-top:28px">Time per Task Type</div>
            <div id="aw-s1-tpt"></div>` : ''}
        ` : `<div class="chart-note" style="color:var(--text-muted)">No Total View uploads were provided in Steps 1–2.</div>`;

        /* SECTION 2 — Transaction Profile (expanded; only provided sources).
           Each chart is followed by its Top-N source table. */
        const profileBody = hasProfile ? `
            ${u.dialog     ? section('Dialog Workload',       'dialog',     u.dialog.meta     || '') : ''}
            ${u.http       ? section('HTTP / HTTPS Workload', 'http',       u.http.meta       || '') : ''}
            ${u.background ? section('Background Workload',   'background', u.background.meta || '') : ''}
            ${u.rfc        ? section('RFC Workload',          'rfc',        u.rfc.meta        || '') : ''}
        ` : `<div class="chart-note" style="color:var(--text-muted)">No Transaction Profile uploads were provided.</div>`;

        /* SECTION 3 — Additional Workload Analysis (collapsed): the supplemental
           per-metric breakdown charts (formerly the Total-View section). */
        const auxBody = (hasTotalWl || hasProfile) ? `
            ${u.dbcpu    ? section('Load History — Breakdown',          'aux-dbcpu',     u.dbcpu.meta    || '') : ''}
            ${u.tasktype ? section('Workload by Task Type — Breakdown', 'aux-tasktype',  u.tasktype.meta || '') : ''}
            ${u.tasktype ? section('Time per Task Type — Breakdown',    'aux-timeptype', u.tasktype.meta || '') : ''}
            ${u.dialog     ? section('Dialog — Breakdown',       'aux-dialog',     '') : ''}
            ${u.http       ? section('HTTP / HTTPS — Breakdown', 'aux-http',       '') : ''}
            ${u.background ? section('Background — Breakdown',   'aux-background', '') : ''}
            ${u.rfc        ? section('RFC — Breakdown',          'aux-rfc',        '') : ''}
        ` : `<div class="chart-note" style="color:var(--text-muted)">No additional charts available.</div>`;

        return `
        <div class="module-wrap">
            <div class="module-header">
                <div class="module-crumb">Analysis › Section 04 › Step 4 of 4</div>
                <h1 class="module-title">Application Workload – Charts</h1>
                <p class="module-desc">
                    Workload analysis charts built from your ST03 uploads.
                </p>
            </div>

            <div class="action-row" style="margin-bottom:20px">
                <button class="btn btn-blue" id="aw-btn-download-top">
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                        <polyline points="7 10 12 15 17 10"/>
                        <line x1="12" y1="15" x2="12" y2="3"/>
                    </svg>
                    Download ${this._sidTag()}_WorkloadAnalysis.xlsx
                </button>
            </div>

            ${wlSection('total',   'Application Workload – Total View', false, totalBody)}
            ${wlSection('profile', 'Transaction Profile',              false, profileBody)}
            ${wlSection('aux',     'Additional Workload Analysis',     true,  auxBody)}

            <div class="action-row" style="margin-top:32px;justify-content:space-between">
                <button class="btn" id="aw-btn-back3">
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2.5">
                        <polyline points="15 18 9 12 15 6"/>
                    </svg>
                    Back
                </button>
                <div style="display:flex;gap:10px">
                    <button class="btn" id="aw-btn-download3">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="7 10 12 15 17 10"/>
                            <line x1="12" y1="15" x2="12" y2="3"/>
                        </svg>
                        Download Workbook
                    </button>
                    <button class="btn btn-blue" id="aw-btn-next">
                        Next: Sizing Projection
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2.5">
                            <polyline points="9 18 15 12 9 6"/>
                        </svg>
                    </button>
                </div>
            </div>
        </div>`;
    },

    /* ══════════════════════════════════════════════════════
       UPLOAD CARD HTML
    ══════════════════════════════════════════════════════ */
    _uploadCard(zoneId, key, title, subtitle, statusId) {
        const loaded = this._uploads[key];
        return `
        <div class="card" style="margin-top:16px">
            <div class="card-header">
                <div>
                    <div class="card-title">${title}</div>
                    <div class="card-sub">${subtitle}</div>
                </div>
                ${loaded ? `<span class="badge-success" style="font-size:12px;padding:3px 10px">Loaded</span>` : ''}
            </div>
            <div class="upload-zone${loaded ? ' upload-zone-done' : ''}" id="${zoneId}-zone">
                <input type="file" class="file-input-hidden" id="${zoneId}-input"
                       accept=".txt,.csv,.xlsx,.xls" data-key="${key}">
                <div class="upload-zone-icon">
                    ${loaded
                        ? `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24"
                                fill="none" stroke="var(--success)" stroke-width="2">
                                <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
                                <polyline points="22 4 12 14.01 9 11.01"/>
                           </svg>`
                        : `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24"
                                fill="none" stroke="currentColor" stroke-width="1.6">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="17 8 12 3 7 8"/>
                                <line x1="12" y1="3" x2="12" y2="15"/>
                           </svg>`
                    }
                </div>
                <div class="upload-zone-text">
                    <div class="upload-zone-title">
                        ${loaded
                            ? `Loaded: <strong>${loaded.fileName}</strong>${loaded.meta ? ` &mdash; <em>${loaded.meta}</em>` : ''}`
                            : `Drop your <strong>${title}</strong> file here or click to browse`}
                    </div>
                    <div class="upload-zone-sub">TXT · CSV · XLSX · XLS</div>
                </div>
                ${loaded ? '' : `<span class="upload-zone-cta">
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                         fill="none" stroke="currentColor" stroke-width="2.2">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                        <polyline points="17 8 12 3 7 8"/>
                        <line x1="12" y1="3" x2="12" y2="15"/>
                    </svg>
                    Choose file
                </span>`}
            </div>
            <div id="${statusId}"></div>
        </div>`;
    },

    /* ══════════════════════════════════════════════════════
       INIT
    ══════════════════════════════════════════════════════ */
    init(app) {
        if (this._step === 1)      this._initStep1(app);
        else if (this._step === 2) this._initDbCpuScreen(app);
        else if (this._step === 3) this._initStep2(app);
        else                       this._initStep3(app);
        this._updateSubNav(app);
    },

    _initStep1(app) {
        const step1Keys = ['dbcpu', 'tasktype'];
        const zoneMap   = { dbcpu: 'aw-dbcpu', tasktype: 'aw-tasktype' };
        const statMap   = { dbcpu: 'aw-dbcpu-status', tasktype: 'aw-tasktype-status' };

        /* Restore status banners for uploads that came from a session restore. */
        step1Keys.forEach(key => {
            const u = this._uploads[key];
            if (u && u.fileName) showFileLoaded(statMap[key], u.fileName + (u.meta ? ` (${u.meta})` : ''));
        });

        step1Keys.forEach(key => {
            initUploadZone(`${zoneMap[key]}-zone`, `${zoneMap[key]}-input`, file => {
                this._parseAndValidate(file, key, statMap[key], () => {
                    this._clearValidationError();
                    this._refreshNextBtn('aw-btn-step2', step1Keys);
                });
            });
        });

        this._refreshNextBtn('aw-btn-step2', step1Keys);

        document.getElementById('aw-btn-step2')?.addEventListener('click', () => {
            this._step = 2;
            app.markCompleted('app-workload');
            this._updateSubNav(app);
            app.navigateTo('app-workload');
        });
    },

    _initStep2(app) {
        this.app = app;
        const profileKeys = ['dialog', 'http', 'background', 'rfc'];
        const zoneMap = { dialog: 'aw-dialog', http: 'aw-http', background: 'aw-background', rfc: 'aw-rfc' };
        const statMap = { dialog: 'aw-dialog-status', http: 'aw-http-status',
                          background: 'aw-background-status', rfc: 'aw-rfc-status' };

        /* Restore status banners for profile uploads that came from a session restore. */
        profileKeys.forEach(key => {
            const u = this._uploads[key];
            if (u && u.fileName) showFileLoaded(statMap[key], u.fileName + (u.meta ? ` (${u.meta})` : ''));
        });

        profileKeys.forEach(key => {
            initUploadZone(`${zoneMap[key]}-zone`, `${zoneMap[key]}-input`, file => {
                this._parseAndValidate(file, key, statMap[key], () => {
                    this._clearValidationError();
                    this._refreshNextBtn('aw-btn-step3', profileKeys, 1);
                });
            });
        });

        this._refreshNextBtn('aw-btn-step3', profileKeys, 1);

        document.getElementById('aw-btn-back2')?.addEventListener('click', () => {
            this._step = 2;
            this._updateSubNav(app);
            app.navigateTo('app-workload');
        });

        document.getElementById('aw-btn-step3')?.addEventListener('click', () => {
            this._step = 4;
            this._updateSubNav(app);
            this._saveModuleData(app);
            app.navigateTo('app-workload');
        });
    },

    _initStep3(app) {
        this.app = app;
        document.getElementById('aw-btn-back3')?.addEventListener('click', () => {
            this._step = 3;
            this._updateSubNav(app);
            app.navigateTo('app-workload');
        });
        document.getElementById('aw-btn-next')?.addEventListener('click', () => app.nextModule());

        /* Ensure the workbook has every analysis tab even if the user skipped
           the DB-vs-CPU screen, then wire the Download buttons. */
        try {
            const prep = this._dbcpuPrepared || this._prepareDbCpu();
            if (prep && !prep.error) { this._dbcpuPrepared = prep; this._writeDbCpuSheet(prep); }
            const ma = this._prepareMonthlyAnalysis();
            if (ma && !ma.error) this._writeMonthlySheet(ma);
            this._writeAllProfileSheets();
        } catch (_) { /* non-fatal — download still offered if a sheet exists */ }
        document.getElementById('aw-btn-download3')?.addEventListener('click', () => this._downloadWorkbook());
        document.getElementById('aw-btn-download-top')?.addEventListener('click', () => this._downloadWorkbook());

        /* Collapsible section Show/Hide controls. */
        document.querySelectorAll('[data-wl-show]').forEach(btn =>
            btn.addEventListener('click', () => document.getElementById('aw-wl-' + btn.dataset.wlShow)?.classList.remove('collapsed')));
        document.querySelectorAll('[data-wl-hide]').forEach(btn =>
            btn.addEventListener('click', () => document.getElementById('aw-wl-' + btn.dataset.wlHide)?.classList.add('collapsed')));

        const u = this._uploads;

        /* ── SECTION 1 — Total View: rich analysis charts + commentary ── */
        const prep = this._dbcpuPrepared || (() => { const p = this._prepareDbCpu(); return (p && !p.error) ? p : null; })();
        if (u.dbcpu && prep) {
            this._dbcpuPrepared = prep;
            this._renderDbCpuStats(prep, 'aw-s1-stats');
            this._mountComboChart(document.getElementById('aw-s1-combo'), prep);
            this._renderDbCpuSummary(document.getElementById('aw-s1-summary'), prep);
        } else if (u.dbcpu) {
            const el = document.getElementById('aw-s1-combo');
            if (el) el.innerHTML = `<div class="chart-note" style="color:var(--text-muted)">At least 3 completed months are required for the DB-vs-CPU analysis.</div>`;
        }
        if (u.tasktype) {
            const ma = this._prepareMonthlyAnalysis();
            if (ma && !ma.error) {
                this._mountTaskTypePie(document.getElementById('aw-s1-pie'), ma);
                this._mountTimePerTaskType(document.getElementById('aw-s1-tpt'), ma);
            } else if (ma && ma.error) {
                const el = document.getElementById('aw-s1-pie');
                if (el) el.innerHTML = `<div class="chart-note" style="color:var(--text-muted)">${ma.error}</div>`;
            }
        }

        /* ── SECTION 2 — Transaction Profile spec charts, each with its
           Top-N source table rendered directly below. ── */
        ['dialog', 'http', 'background', 'rfc'].forEach(key => {
            if (u[key]) this._mountProfileSpecCharts(document.getElementById(`aw-cb-${key}`), u[key], key);
        });

        /* ── SECTION 3 — Additional Workload Analysis: supplemental breakdown
           charts (formerly the Total-View per-metric charts). ── */
        if (u.dbcpu)    this._mountDbcpuCharts(document.getElementById('aw-cb-aux-dbcpu'), u.dbcpu);
        if (u.tasktype) {
            this._mountTasktypeCharts(document.getElementById('aw-cb-aux-tasktype'), u.tasktype);
            this._mountTimePerTaskChart(document.getElementById('aw-cb-aux-timeptype'), u.tasktype);
        }
        ['dialog', 'http', 'background', 'rfc'].forEach(key => {
            if (u[key]) this._mountProfileAuxCharts(document.getElementById(`aw-cb-aux-${key}`), u[key], key);
        });
    },

    /* ══════════════════════════════════════════════════════
       SECTION 2 — TRANSACTION-PROFILE SPEC CHARTS
       Uses the derived dataset (sorted DESC by Total CPU+DB hr) and the
       reusable editable combo factory. Colour convention: CPU = amber,
       DB = gray, Acc.% line = blue (right axis 0–100).
    ══════════════════════════════════════════════════════ */
    _mountProfileSpecCharts(container, upload, key) {
        if (!container) return;
        const derived = (upload.derived && upload.derived.rows) ? upload.derived
                      : this._computeProfileDerived(upload.rows || [], key);
        const rows = derived.rows;
        if (!rows.length) { container.innerHTML = `<div class="chart-note">No data rows found.</div>`; return; }

        const C = this._COMBO_COLORS;
        const r2 = v => this._r2(v);
        const metaHtml = upload.meta ? `<div class="chart-note" style="margin-bottom:10px">${upload.meta}</div>` : '';
        container.innerHTML = metaHtml;

        if (key === 'rfc') { this._mountRfcSpecCharts(container, derived); return; }

        const top = rows.slice(0, 20);
        const labels = top.map(r => key === 'background' ? (r.trans_rep || r.job_name || '—') : (r.trans_rep || '—'));
        const titlePrefix = key === 'background' ? 'Top 20 Background jobs by contribution' : 'Top transactions by contribution';

        /* Chart 1 — Total CPU/DB hr clustered + Acc.% line (0–100). */
        const c1 = document.createElement('div'); container.appendChild(c1);
        this._mountEditableCombo(c1, {
            uid: `tp-${key}-1`,
            title: titlePrefix,
            labels,
            bars: [
                { label: 'Total CPU Time (hr)', data: top.map(r => r2(r.cpuH)), color: C.cpu },
                { label: 'Total DB Time (hr)',  data: top.map(r => r2(r.dbH)),  color: C.db },
            ],
            line: { label: 'Acc. Total CPU+DB Time (%)', data: top.map(r => r2(r.accPct)), color: C.steps, max: 100 },
            leftTitle: 'Time (hr)', rightTitle: 'Percentage (%)',
            csv: {
                headers: ['Report or Transaction name', 'Total CPU Time (hr)', 'Total DB Time (hr)', 'Acc. Total CPU+DB Time (%)'],
                rows: top.map((r, i) => [labels[i], r2(r.cpuH), r2(r.dbH), r2(r.accPct)]),
            },
            sourceTable: true, sourceCaption: `Top ${top.length} records (sorted by Total CPU+DB Time hr)`,
            hint: 'Top 20 by Total CPU+DB Time (hr). Columns = CPU &amp; DB hours (left) · line = accumulated contribution % (right, 0–100).',
        });

        /* Background has only Chart 1 per the spec. */
        if (key === 'background') return;

        /* Chart 2 — Average Response Time: stacked Ø CPU/DB/Other (s) + Acc.% line. */
        const c2 = document.createElement('div'); container.appendChild(c2);
        this._mountEditableCombo(c2, {
            uid: `tp-${key}-2`,
            title: 'Average Response Time',
            labels,
            stacked: true,
            bars: [
                { label: 'Ø CPU (s)',       data: top.map(r => r2(r.cpuS)),   color: '#FFC000' },
                { label: 'Ø DB Time (s)',   data: top.map(r => r2(r.dbS)),    color: '#A5A5A5' },
                { label: 'Ø Other Time (s)',data: top.map(r => r2(r.otherS)), color: '#70AD47' },
            ],
            line: { label: 'Acc. Total CPU+DB Time (%)', data: top.map(r => r2(r.accPct)), color: '#00B0F0', max: 100 },
            leftTitle: 'Response Time (seconds)', rightTitle: 'Percentage (%)',
            csv: {
                headers: ['Report or Transaction name', 'Ø CPU (s)', 'Ø DB Time (s)', 'Ø Other Time (s)', 'Acc. Total CPU+DB Time (%)'],
                rows: top.map(r => [r.trans_rep || '—', r2(r.cpuS), r2(r.dbS), r2(r.otherS), r2(r.accPct)]),
            },
            sourceTable: true, sourceCaption: `Top ${top.length} records (sorted by Total CPU+DB Time hr)`,
            hint: 'Top 20 by Total CPU+DB Time (hr). Stacked columns = average CPU / DB / other time per step in seconds (left) · line = accumulated % (right, 0–100).',
        });
    },

    /* RFC Section-2 spec charts (3 charts: users, systems, actions). */
    _mountRfcSpecCharts(container, derived) {
        const C = this._COMBO_COLORS;
        const r2 = v => this._r2(v);
        const rows = derived.rows;

        /* Chart 1 — Top 20 by User ID. */
        const top20 = rows.slice(0, 20);
        const c1 = document.createElement('div'); container.appendChild(c1);
        this._mountEditableCombo(c1, {
            uid: 'tp-rfc-1',
            title: 'CPU and DB Time per Month for RFC Users',
            labels: top20.map(r => r.user_id || '—'),
            bars: [
                { label: 'Total CPU Time (hr)', data: top20.map(r => r2(r.cpuH)), color: C.cpu },
                { label: 'Total DB Time (hr)',  data: top20.map(r => r2(r.dbH)),  color: C.db },
            ],
            line: null,
            leftTitle: 'Hours',
            csv: {
                headers: ['User ID', 'Total CPU Time (hr)', 'Total DB Time (hr)'],
                rows: top20.map(r => [r.user_id || '—', r2(r.cpuH), r2(r.dbH)]),
            },
            sourceTable: true, sourceCaption: `Top ${top20.length} records (sorted by Total CPU+DB Time hr)`,
            hint: 'Top 20 rows by Total CPU+DB Time (hr), grouped by User ID.',
        });

        /* Chart 2 — Top SAP Systems: System ID exactly 3 alphanumerics, top 5 by CPU+DB hr. */
        const bySystem = {};
        rows.forEach(r => {
            const sid = String(r.system_id || '').trim();
            if (!/^[A-Za-z0-9]{3}$/.test(sid)) return;
            if (!bySystem[sid]) bySystem[sid] = { system_id: sid, cpuH: 0, dbH: 0, cpuDbH: 0 };
            bySystem[sid].cpuH += r.cpuH; bySystem[sid].dbH += r.dbH; bySystem[sid].cpuDbH += r.cpuDbH;
        });
        const sysRows = Object.values(bySystem).sort((a, b) => b.cpuDbH - a.cpuDbH).slice(0, 5);
        const c2 = document.createElement('div'); container.appendChild(c2);
        this._mountEditableCombo(c2, {
            uid: 'tp-rfc-2',
            title: 'Top SAP Systems by RFC Contribution',
            labels: sysRows.map(r => r.system_id),
            bars: [
                { label: 'Total CPU Time (hr)', data: sysRows.map(r => r2(r.cpuH)), color: C.cpu },
                { label: 'Total DB Time (hr)',  data: sysRows.map(r => r2(r.dbH)),  color: C.db },
            ],
            line: null,
            leftTitle: 'Hours',
            csv: {
                headers: ['System ID', 'Total CPU Time (hr)', 'Total DB Time (hr)'],
                rows: sysRows.map(r => [r.system_id, r2(r.cpuH), r2(r.dbH)]),
            },
            sourceTable: true, sourceCaption: `Top ${sysRows.length} systems (sorted by Total CPU+DB Time hr)`,
            hint: 'Top 5 SAP systems (System ID of exactly 3 alphanumeric characters), by Total CPU+DB Time (hr).',
        });

        /* Chart 3 — Top 20 by triggered action. */
        const c3 = document.createElement('div'); container.appendChild(c3);
        this._mountEditableCombo(c3, {
            uid: 'tp-rfc-3',
            title: 'Triggered action by contribution',
            labels: top20.map(r => r.action || '—'),
            bars: [
                { label: 'Total CPU Time (hr)', data: top20.map(r => r2(r.cpuH)), color: C.cpu },
                { label: 'Total DB Time (hr)',  data: top20.map(r => r2(r.dbH)),  color: C.db },
            ],
            line: null,
            leftTitle: 'Hours',
            csv: {
                headers: ['Action in Triggering System', 'Total CPU Time (hr)', 'Total DB Time (hr)'],
                rows: top20.map(r => [r.action || '—', r2(r.cpuH), r2(r.dbH)]),
            },
            sourceTable: true, sourceCaption: `Top ${top20.length} records (sorted by Total CPU+DB Time hr)`,
            hint: 'Top 20 rows by Total CPU+DB Time (hr), grouped by Action in Triggering System.',
        });
    },

    /* ══════════════════════════════════════════════════════
       SECTION 3 — ADDITIONAL (auxiliary) per-metric charts
       The former per-metric breakdowns, kept available but collapsed.
    ══════════════════════════════════════════════════════ */
    _mountProfileAuxCharts(container, upload, key) {
        if (!container) return;
        if (key === 'rfc') { this._mountRfcCharts(container, upload); return; }
        this._mountProfileCharts(container, upload, key);
    },

    /* ══════════════════════════════════════════════════════
       PARSE & VALIDATE
    ══════════════════════════════════════════════════════ */

    /*
     * Dispatches to the right schema + parser:
     *   dbcpu      → legacy _DBCPU_ALIASES   + _DBCPU_REQUIRED    (Step-1 upload)
     *   tasktype   → legacy _TASKTYPE_ALIASES + _TASKTYPE_REQUIRED (Step-1 upload)
     *   dialog / http / background / rfc → _WORKLOAD_MODEL (model-driven,
     *       validated + coerced by _resolveHeaders / _validateHeaders /
     *       _coerceRow, with an inline determinate progress bar for the
     *       large ST03 Transaction Profile exports).
     *
     * Profile uploads parse asynchronously in row batches so the progress bar
     * animates and the UI stays responsive on files with thousands of rows.
     */
    _parseAndValidate(file, key, statusId, onDone) {
        const ext = file.name.split('.').pop().toLowerCase();
        const isProfile = (key === 'dialog' || key === 'http' || key === 'background' || key === 'rfc');

        /* ── Legacy path (Step-1 dbcpu / tasktype) — unchanged, synchronous ── */
        if (!isProfile) {
            let aliasMap, required;
            if (key === 'dbcpu')      { aliasMap = this._DBCPU_ALIASES;    required = this._DBCPU_REQUIRED; }
            else                      { aliasMap = this._TASKTYPE_ALIASES; required = this._TASKTYPE_REQUIRED; }

            if (ext === 'txt' || ext === 'csv') {
                const reader = new FileReader();
                reader.onload = e => {
                    const result = this._parseTxtLegacy(e.target.result, aliasMap, required, key);
                    if (result.error) { this._showValidationError(result.error); return; }
                    this._storeUpload(key, result, file.name, statusId);
                    onDone();
                };
                reader.readAsText(file);
                return;
            }
            if (['xlsx', 'xls'].includes(ext)) {
                const reader = new FileReader();
                reader.onload = e => {
                    const wb   = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
                    const ws   = wb.Sheets[wb.SheetNames[0]];
                    const raw  = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                    const result = this._parseXlsxLegacy(raw, aliasMap, required, key);
                    if (result.error) { this._showValidationError(result.error); return; }
                    this._storeUpload(key, result, file.name, statusId);
                    onDone();
                };
                reader.readAsBinaryString(file);
                return;
            }
            this._showValidationError('Unsupported format. Please upload a TXT, CSV, XLSX, or XLS file.');
            return;
        }

        /* ── Model-driven profile path (dialog / http / background / rfc) ── */
        const model = this._WORKLOAD_MODEL[key];

        const run = async () => {
            try {
                let result;
                if (ext === 'txt' || ext === 'csv') {
                    const text = await this._readFileText(file, key, statusId);
                    result = await this._parseTxtProfile(text, model, key, statusId, file.name);
                } else if (['xlsx', 'xls'].includes(ext)) {
                    const bin = await this._readFileBinary(file, key, statusId);
                    this._showParseProgress(key, { phase: 'Parsing', file: file.name, done: 0, total: 0 });
                    const wb  = XLSX.read(bin, { type: 'binary', cellDates: true });
                    const ws  = wb.Sheets[wb.SheetNames[0]];
                    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                    result = await this._parseXlsxProfile(raw, model, key, statusId, file.name);
                } else {
                    this._showValidationError('Unsupported format. Please upload a TXT, CSV, XLSX, or XLS file.');
                    return;
                }
                if (result.error) { this._hideParseProgress(key); this._showValidationError(result.error); return; }

                /* Aggregation phase — compute the derived columns now so the
                   progress bar covers it and the workbook/charts can reuse it. */
                this._showParseProgress(key, { phase: 'Aggregation', file: file.name, done: result.rows.length, total: result.rows.length });
                await this._yield();
                result.derived = this._computeProfileDerived(result.rows, key);

                this._storeUpload(key, result, file.name, statusId);
                this._hideParseProgress(key);
                onDone();
            } catch (err) {
                this._hideParseProgress(key);
                this._showValidationError('Could not read the file: ' + (err && err.message ? err.message : String(err)));
            }
        };
        run();
    },

    /* ── Async FileReader wrappers that drive the "File loading" phase ── */
    _readFileText(file, key, statusId) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            this._showParseProgress(key, { phase: 'File loading', file: file.name, done: 0, total: file.size || 0, bytes: true });
            reader.onprogress = e => {
                if (e.lengthComputable) this._showParseProgress(key, { phase: 'File loading', file: file.name, done: e.loaded, total: e.total, bytes: true });
            };
            reader.onload  = e => resolve(e.target.result);
            reader.onerror = () => reject(reader.error || new Error('read error'));
            reader.readAsText(file);
        });
    },
    _readFileBinary(file, key, statusId) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            this._showParseProgress(key, { phase: 'File loading', file: file.name, done: 0, total: file.size || 0, bytes: true });
            reader.onprogress = e => {
                if (e.lengthComputable) this._showParseProgress(key, { phase: 'File loading', file: file.name, done: e.loaded, total: e.total, bytes: true });
            };
            reader.onload  = e => resolve(e.target.result);
            reader.onerror = () => reject(reader.error || new Error('read error'));
            reader.readAsBinaryString(file);
        });
    },

    /* Yield to the event loop so the progress bar can paint between batches. */
    _yield() {
        return new Promise(r => (window.requestAnimationFrame ? requestAnimationFrame(() => r()) : setTimeout(r, 0)));
    },

    /* ── TXT parser (legacy: dbcpu / tasktype Step-1 uploads) ──
       Handles the SAP pipe-table format:
         • Skips lines of pure dashes/spaces
         • Reads metadata header (| Instance … | Period … |) for period extraction
         • Finds the actual column header row (row starting with |Start Date or |Task Type)
         • Parses data rows
    */
    _parseTxtLegacy(text, aliasMap, required, key) {
        const lines = text.split(/\r?\n/);

        /* Extract metadata: look for lines like "| Period       02/2024  …" */
        let meta = '';
        for (const line of lines) {
            const mPeriod = line.match(/\|\s*Period\s+([\d/]+)/i);
            if (mPeriod) { meta = `Period: ${mPeriod[1].trim()}`; break; }
            /* Load History uses "Start Date" in data, but Instance/TOTAL in meta */
            const mInstance = line.match(/\|\s*Instance\s+(\S+)/i);
            if (mInstance && key === 'dbcpu') { /* captured but not used as meta label */ }
        }

        /* Find the column header row — the first line that looks like a pipe-delimited header:
           must contain at least 5 pipe chars and contain a known label keyword */
        let headerLineIdx = -1;
        for (let i = 0; i < lines.length; i++) {
            const l = lines[i];
            if ((l.match(/\|/g) || []).length < 4) continue;
            const norm = l.replace(/\|/g, ' ').replace(/\s+/g, ' ').toLowerCase().trim();
            if (norm.includes('start date')  || norm.includes('task type')   ||
                norm.includes('trans/rep')   || norm.includes('trans/rep.')  ||
                norm.includes('report or transaction') ||
                norm.includes('system id')   || norm.includes('action/transaction') ||
                norm.includes('# steps')     || norm.includes('number of dialog')) {
                headerLineIdx = i;
                break;
            }
        }

        if (headerLineIdx < 0) {
            return { error: 'Could not locate the column header row in the TXT file. ' +
                'Make sure this is a plain-text export from ST03 with "|" column separators.' };
        }

        /* Split header row into column names */
        const rawHeaders = lines[headerLineIdx]
            .split('|')
            .map(h => h.trim())
            .filter(h => h.length > 0);

        /* Map to canonical names */
        const canonHeaders = rawHeaders.map(h => this._canonicalize(h, aliasMap));

        /* Validate required columns */
        const missing = required.filter(r => !canonHeaders.includes(r));
        if (missing.length) {
            const missingLabels = missing.map(m => this._canonicalToLabel(m, aliasMap)).join(', ');
            return { error: `The following required columns were not found in the uploaded file: <strong>${missingLabels}</strong>. ` +
                `Please verify this is the correct ST03 export for this section.` };
        }

        /* Parse data rows (lines after header that are pipe-delimited and contain data) */
        const rows = [];
        for (let i = headerLineIdx + 1; i < lines.length; i++) {
            const l = lines[i];
            /* Skip separator lines (all dashes/spaces/pipes) */
            if (/^[\s\-|]+$/.test(l)) continue;
            /* Only rows with pipe delimiters */
            if ((l.match(/\|/g) || []).length < 4) continue;
            const cells = l.split('|').map(c => c.trim()).filter((_, ci) => {
                /* filter: skip the virtual empty columns at start/end of | row | */
                return true;
            });
            /* Realign: the first and last split from " | row | " produce empty strings */
            const dataCells = l.split('|').slice(1, -1).map(c => c.trim());
            if (dataCells.length === 0) continue;
            /* Stop at footer lines (Page: …) */
            if (dataCells[0].toLowerCase().startsWith('page')) break;
            if (dataCells.length < required.length) continue;

            const row = {};
            canonHeaders.forEach((col, ci) => {
                if (!col) return;
                row[col] = this._parseNum(dataCells[ci] ?? '');
            });
            /* Keep string-identity columns as raw text (overwrite the numeric parse) */
            const strCols = ['date', 'task_type', 'trans_rep', 'job_name', 'system_id', 'action', 'user_id'];
            strCols.forEach(sc => {
                const idx = canonHeaders.indexOf(sc);
                if (idx >= 0) row[sc] = (dataCells[idx] ?? '').trim();
            });
            if (Object.keys(row).length > 0) rows.push(row);
        }

        if (!rows.length) {
            return { error: 'The file was parsed successfully but no data rows were found. ' +
                'Check that the export contains at least one measurement row.' };
        }

        return { rows, canonHeaders, meta };
    },

    /* ── XLSX parser (legacy: dbcpu / tasktype Step-1 uploads) ──
       Scans for the header row (first row where a cell matches a known alias),
       then maps to canonical names and parses.
    */
    _parseXlsxLegacy(raw, aliasMap, required, key) {
        if (!raw.length) return { error: 'The spreadsheet appears to be empty.' };

        let headerRowIdx = -1;
        for (let i = 0; i < Math.min(raw.length, 20); i++) {
            const norm = (raw[i][0] ?? '').toString().trim().toLowerCase();
            /* Check if first cell matches a known alias key */
            if (aliasMap[norm] !== undefined) { headerRowIdx = i; break; }
            /* Also check second cell (some exports have an extra leading column) */
            const norm2 = (raw[i][1] ?? '').toString().trim().toLowerCase();
            if (aliasMap[norm2] !== undefined) { headerRowIdx = i; break; }
        }

        if (headerRowIdx < 0) {
            /* Fallback: any row where ≥3 cells match known aliases */
            for (let i = 0; i < Math.min(raw.length, 20); i++) {
                const hits = raw[i].filter(c => aliasMap[String(c).trim().toLowerCase()] !== undefined).length;
                if (hits >= 3) { headerRowIdx = i; break; }
            }
        }

        if (headerRowIdx < 0) {
            return { error: 'Could not locate the column header row in the spreadsheet. ' +
                'Please ensure you are uploading the correct ST03 Excel export.' };
        }

        const rawHeaders = raw[headerRowIdx].map(c => String(c).trim());
        const canonHeaders = rawHeaders.map(h => this._canonicalize(h, aliasMap));

        const missing = required.filter(r => !canonHeaders.includes(r));
        if (missing.length) {
            const missingLabels = missing.map(m => this._canonicalToLabel(m, aliasMap)).join(', ');
            return { error: `The following required columns were not found in the uploaded file: <strong>${missingLabels}</strong>. ` +
                `Please verify this is the correct ST03 export for this section.` };
        }

        /* Extract meta (period) from rows before the header in XLSX — typically not present,
           but we can also extract from a "Period" cell if it appears in the pre-header rows. */
        let meta = '';
        for (let i = 0; i < headerRowIdx; i++) {
            for (const cell of raw[i]) {
                const m = String(cell).match(/period\s+([\d/]+)/i);
                if (m) { meta = `Period: ${m[1].trim()}`; break; }
            }
            if (meta) break;
        }

        const strCols = new Set(['date', 'task_type', 'trans_rep', 'job_name', 'system_id', 'action', 'user_id']);
        const rows = [];
        for (let i = headerRowIdx + 1; i < raw.length; i++) {
            const cells = raw[i];
            if (cells.every(c => String(c).trim() === '')) continue;
            const row = {};
            canonHeaders.forEach((col, ci) => {
                if (!col) return;
                const cell = cells[ci];
                if (col === 'date') {
                    /* The Start Date column may arrive as a JS Date (cellDates:true),
                       an Excel serial number, or an already-formatted string.
                       Normalize all three to DD.MM.YYYY so _parseDotDate works. */
                    row[col] = this._xlsxDateToDotStr(cell);
                } else if (strCols.has(col)) {
                    row[col] = String(cell ?? '').trim();
                } else if (typeof cell === 'number') {
                    /* XLSX stores numbers natively (decimal point already correct);
                       use the value as-is. Running _parseNum on it would wrongly
                       strip the "." as a thousands separator (641.8 → 6418). */
                    row[col] = isFinite(cell) ? cell : null;
                } else {
                    /* String cell (CSV, or text in an XLSX) — apply European-format
                       normalization: "." thousands, "," decimal. */
                    row[col] = this._parseNum(String(cell ?? '').trim());
                }
            });
            rows.push(row);
        }

        if (!rows.length) {
            return { error: 'The spreadsheet was parsed successfully but no data rows were found.' };
        }

        return { rows, canonHeaders, meta };
    },

    /* ══════════════════════════════════════════════════════
       MODEL-DRIVEN PROFILE PARSERS (async, chunked, with progress)
       dialog / http / background / rfc
    ══════════════════════════════════════════════════════ */

    _BATCH: 1000,   /* rows processed per event-loop tick */

    /* TXT (pipe-table) — model-driven, batched. */
    async _parseTxtProfile(text, model, key, statusId, fileName) {
        this._showParseProgress(key, { phase: 'Validation', file: fileName, done: 0, total: 0 });
        await this._yield();

        const lines = text.split(/\r?\n/);

        /* Reporting period from the metadata header. */
        let meta = '';
        for (const line of lines) {
            const mPeriod = line.match(/\|\s*Period\s+([\d/]+)/i);
            if (mPeriod) { meta = `Period: ${mPeriod[1].trim()}`; break; }
        }

        /* Locate the column header row. */
        let headerLineIdx = -1;
        for (let i = 0; i < lines.length; i++) {
            const l = lines[i];
            if ((l.match(/\|/g) || []).length < 4) continue;
            const norm = l.replace(/\|/g, ' ').replace(/\s+/g, ' ').toLowerCase().trim();
            if (norm.includes('report or transaction') || norm.includes('report/transaction') ||
                norm.includes('trans/rep') || norm.includes('system id') ||
                norm.includes('action in triggering') || norm.includes('action/transaction') ||
                norm.includes('# steps') || norm.includes('number of steps')) {
                headerLineIdx = i;
                break;
            }
        }
        if (headerLineIdx < 0) {
            return { error: 'Could not locate the column header row in the TXT file. ' +
                'Make sure this is a plain-text export from ST03 with "|" column separators.' };
        }

        const rawHeaders = lines[headerLineIdx].split('|').map(h => h.trim()).filter(h => h.length > 0);
        const { canon, specs, unresolved } = this._resolveHeaders(rawHeaders, model);
        const validation = this._validateHeaders(canon, model);
        if (!validation.ok) {
            return { error: `The following required columns were not found in the uploaded file: <strong>${validation.missing.join(', ')}</strong>. ` +
                `Please verify this is the correct ST03 export for this section.` };
        }

        /* Collect candidate data lines first (cheap), then coerce in batches. */
        const dataLines = [];
        for (let i = headerLineIdx + 1; i < lines.length; i++) {
            const l = lines[i];
            if (/^[\s\-|]+$/.test(l)) continue;
            if ((l.match(/\|/g) || []).length < 4) continue;
            const dataCells = l.split('|').slice(1, -1).map(c => c.trim());
            if (dataCells.length === 0) continue;
            if (dataCells[0].toLowerCase().startsWith('page')) break;
            dataLines.push(dataCells);
        }

        const total = dataLines.length;
        const rows = [];
        for (let start = 0; start < total; start += this._BATCH) {
            const end = Math.min(start + this._BATCH, total);
            for (let j = start; j < end; j++) {
                const row = this._coerceRow(dataLines[j], specs, false);
                if (Object.keys(row).length) rows.push(row);
            }
            this._showParseProgress(key, { phase: 'Parsing', file: fileName, done: end, total });
            await this._yield();
        }

        if (!rows.length) {
            return { error: 'The file was parsed successfully but no data rows were found. ' +
                'Check that the export contains at least one measurement row.' };
        }

        const validationMeta = this._buildValidationMeta(key, rawHeaders, canon, specs, validation, rows.length);
        return { rows, canonHeaders: canon, meta, validation: validationMeta, unresolved };
    },

    /* XLSX — model-driven, batched. `raw` is an array-of-arrays. */
    async _parseXlsxProfile(raw, model, key, statusId, fileName) {
        if (!raw.length) return { error: 'The spreadsheet appears to be empty.' };
        this._showParseProgress(key, { phase: 'Validation', file: fileName, done: 0, total: 0 });
        await this._yield();

        const idx = this._aliasIndex(model);
        /* Header row = first row (≤20) where a cell resolves to a model field. */
        let headerRowIdx = -1;
        for (let i = 0; i < Math.min(raw.length, 20); i++) {
            const hits = raw[i].filter(c => this._resolveField(c, idx) !== null).length;
            if (hits >= 3) { headerRowIdx = i; break; }
        }
        if (headerRowIdx < 0) {
            return { error: 'Could not locate the column header row in the spreadsheet. ' +
                'Please ensure you are uploading the correct ST03 Excel export.' };
        }

        const rawHeaders = raw[headerRowIdx].map(c => String(c).trim());
        const { canon, specs, unresolved } = this._resolveHeaders(rawHeaders, model);
        const validation = this._validateHeaders(canon, model);
        if (!validation.ok) {
            return { error: `The following required columns were not found in the uploaded file: <strong>${validation.missing.join(', ')}</strong>. ` +
                `Please verify this is the correct ST03 export for this section.` };
        }

        /* Period from any pre-header row. */
        let meta = '';
        for (let i = 0; i < headerRowIdx; i++) {
            for (const cell of raw[i]) {
                const m = String(cell).match(/period\s+([\d/]+)/i);
                if (m) { meta = `Period: ${m[1].trim()}`; break; }
            }
            if (meta) break;
        }

        const total = raw.length - headerRowIdx - 1;
        const rows = [];
        let processed = 0;
        for (let start = headerRowIdx + 1; start < raw.length; start += this._BATCH) {
            const end = Math.min(start + this._BATCH, raw.length);
            for (let i = start; i < end; i++) {
                const cells = raw[i];
                if (cells.every(c => String(c).trim() === '')) { processed++; continue; }
                rows.push(this._coerceRow(cells, specs, true));
                processed++;
            }
            this._showParseProgress(key, { phase: 'Parsing', file: fileName, done: processed, total });
            await this._yield();
        }

        if (!rows.length) {
            return { error: 'The spreadsheet was parsed successfully but no data rows were found.' };
        }

        const validationMeta = this._buildValidationMeta(key, rawHeaders, canon, specs, validation, rows.length);
        return { rows, canonHeaders: canon, meta, validation: validationMeta, unresolved };
    },

    /* ══════════════════════════════════════════════════════
       DERIVED-COLUMN COMPUTATION  (per ST03_TP_* spec)
       ──────────────────────────────────────────────────────
       All source Ø values are already in ms and T values in s (General
       Data Preparation rules). Produces, per row:
         cpuH   = cpu_total_s / 3600          (Total CPU Time hr)
         dbH    = db_total_s  / 3600          (Total DB Time hr)
         cpuDbH = cpuH + dbH                  (Total CPU+DB Time hr)
         cpuDbPct, accPct                     (contribution + accumulation)
         cpuS   = cpu_ms / 1000               (Ø CPU s)
         dbS    = db_ms  / 1000               (Ø DB s)         [trans/rep only]
         rollWaitS = rollwait_ms / 1000       (Ø Roll Wait s)  [trans/rep only]
         waitS     = wait_ms     / 1000       (Ø WaitTime s)   [trans/rep only]
         otherS    = waitS + rollWaitS        (Ø Other s)      [trans/rep only]
       and, for RFC, the System ID / Application Server split.
       Rows are returned sorted DESC by cpuDbH, with accPct accumulated in
       that order. var_total_cpu_db is the sum of cpuDbH across all rows.
    ══════════════════════════════════════════════════════ */
    _computeProfileDerived(rows, key) {
        const n = v => { const x = (typeof v === 'number') ? v : this._parseNum(v); return (x == null || !isFinite(x)) ? 0 : x; };
        const isRfc = key === 'rfc';

        const enriched = rows.map(r => {
            const cpuH = n(r.cpu_total_s) / 3600;
            const dbH  = n(r.db_total_s)  / 3600;
            const cpuDbH = cpuH + dbH;
            const base = {
                ...r,
                cpuH, dbH, cpuDbH,
                cpuS: n(r.cpu_ms) / 1000,
            };
            if (!isRfc) {
                base.dbS       = n(r.db_ms)       / 1000;
                base.rollWaitS = n(r.rollwait_ms) / 1000;
                base.waitS     = n(r.wait_ms)     / 1000;
                base.otherS    = base.waitS + base.rollWaitS;
            } else {
                /* System ID split: "<SID>/<Application Server>" → SID + app server. */
                const raw = String(r.system_id ?? '').trim();
                const slash = raw.indexOf('/');
                if (slash > 0) {
                    base.system_id_full  = raw;
                    base.system_id       = raw.slice(0, slash);
                    base.app_server      = raw.slice(slash + 1);
                } else {
                    base.system_id_full  = raw;
                    base.app_server      = '';
                }
            }
            return base;
        });

        const varTotal = enriched.reduce((s, r) => s + r.cpuDbH, 0);
        enriched.forEach(r => { r.cpuDbPct = varTotal > 0 ? r.cpuDbH / varTotal * 100 : 0; });

        /* Sort DESC by Total CPU+DB hr, then accumulate the % in that order. */
        enriched.sort((a, b) => b.cpuDbH - a.cpuDbH);
        let acc = 0;
        enriched.forEach(r => { acc += r.cpuDbPct; r.accPct = acc; });

        return { rows: enriched, varTotal };
    },

    /* ══════════════════════════════════════════════════════
       DETERMINATE PROGRESS BAR  (reuses .proc-progress component)
       Rendered inline inside the active upload card's status row.
       Shows: phase · current file · records processed / total · percentage.
       Phase → percentage bands:
         File loading 0–15 · Validation 15–20 · Parsing 20–85 ·
         Aggregation 85–95 · Workbook generation 95–100
    ══════════════════════════════════════════════════════ */
    _PHASE_BANDS: {
        'File loading':        [0, 15],
        'Validation':          [15, 20],
        'Parsing':             [20, 85],
        'Aggregation':         [85, 95],
        'Workbook generation': [95, 100],
    },

    _showParseProgress(key, { phase, file, done, total, bytes }) {
        const statusId = {
            dialog: 'aw-dialog-status', http: 'aw-http-status',
            background: 'aw-background-status', rfc: 'aw-rfc-status',
        }[key];
        const el = document.getElementById(statusId);
        if (!el) return;

        const band = this._PHASE_BANDS[phase] || [0, 100];
        const frac = total > 0 ? Math.min(1, done / total) : (phase === 'Validation' ? 1 : 0);
        const pct  = Math.round(band[0] + (band[1] - band[0]) * frac);

        const recordsLine = bytes
            ? `${this._fmtBytes(done)} / ${total ? this._fmtBytes(total) : '—'}`
            : `${done.toLocaleString()} / ${total ? total.toLocaleString() : '—'} records`;

        el.style.display = 'block';
        el.innerHTML = `
            <div class="proc-progress" role="status" aria-live="polite">
                <div class="proc-progress-head">
                    <span class="proc-progress-label">${phase} — <strong>${file}</strong> · ${recordsLine}</span>
                    <span class="proc-progress-pct">${pct}%</span>
                </div>
                <div class="proc-progress-track"><div class="proc-progress-fill" style="width:${pct}%"></div></div>
            </div>`;
    },

    _hideParseProgress(key) {
        /* The success banner (_storeUpload → showFileLoaded) replaces this; on
           error we clear it so a stale bar is not left behind. */
        const statusId = {
            dialog: 'aw-dialog-status', http: 'aw-http-status',
            background: 'aw-background-status', rfc: 'aw-rfc-status',
        }[key];
        const el = document.getElementById(statusId);
        if (el && el.querySelector('.proc-progress')) { el.innerHTML = ''; el.style.display = 'none'; }
    },

    _fmtBytes(b) {
        if (b < 1024) return `${b} B`;
        if (b < 1048576) return `${(b / 1024).toFixed(0)} KB`;
        return `${(b / 1048576).toFixed(1)} MB`;
    },



    /* Normalize a raw header to the form used in the model's `aliases`
       (lower-cased, single-spaced). Unit suffixes are kept — the alias
       lists carry both the suffixed and unsuffixed spellings. */
    _normHeader(raw) {
        return String(raw ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
    },

    /* Build (and cache) a normalized-alias → fieldSpec index for a model. */
    _aliasIndex(model) {
        this.__aliasIndexCache = this.__aliasIndexCache || new Map();
        if (this.__aliasIndexCache.has(model)) return this.__aliasIndexCache.get(model);
        const idx = new Map();
        model.forEach(spec => {
            spec.aliases.forEach(a => {
                const n = this._normHeader(a);
                if (!idx.has(n)) idx.set(n, spec);   /* first spec wins on shared spellings */
            });
        });
        this.__aliasIndexCache.set(model, idx);
        return idx;
    },

    /* Resolve a raw header to a field spec via the alias index, with a
       suffix-tolerant fallback (strip trailing "(ms)"/"(kb)"/"(s)"). */
    _resolveField(rawHeader, idx) {
        const n = this._normHeader(rawHeader);
        if (idx.has(n)) return idx.get(n);
        const stripped = n.replace(/\s*\((?:ms|kb|s|byte)\)\s*$/, '');
        if (idx.has(stripped)) return idx.get(stripped);
        return null;
    },

    /* Column-NAME-driven header resolution for a whole header row.
       Returns { canon:[], specs:[], unresolved:[] } where canon[i] is the
       canonical name for rawHeaders[i] ('' if unresolved).

       Repeated short labels (e.g. RFC TXT "Ø Time" ×3) resolve to the same
       spec; when a spec would be assigned more than once we fall back to the
       spec's `txtPos` ordering so the 1st/2nd/3rd occurrence map to the
       response / db-procedure / frontend averages respectively. This replaces
       the former hard-coded RFC_POS positional block. */
    _resolveHeaders(rawHeaders, model) {
        const idx = this._aliasIndex(model);
        const canon = new Array(rawHeaders.length).fill('');
        const specs = new Array(rawHeaders.length).fill(null);
        const unresolved = [];

        /* Repeated-label groups: in TXT exports a short label (e.g. "Ø Time")
           appears several times for distinct metrics. Specs that declare a
           `txtRepeat` trigger belong to one group, ordered by `txtPos`. The
           Nth occurrence of that bare label in the header row maps to the Nth
           spec in the group. XLSX headers are unique so this never fires. */
        const repeatGroups = {};   /* trigger spelling → [specs sorted by txtPos] */
        model.forEach(s => {
            if (!s.txtRepeat) return;
            const trig = this._normHeader(s.txtRepeat);
            (repeatGroups[trig] = repeatGroups[trig] || []).push(s);
        });
        Object.values(repeatGroups).forEach(g => g.sort((a, b) => (a.txtPos || 0) - (b.txtPos || 0)));
        const repeatCursor = {};   /* trigger spelling → next group index to assign */

        rawHeaders.forEach((h, i) => {
            const norm = this._normHeader(h);
            /* A bare repeated-label header is resolved positionally within its group. */
            if (repeatGroups[norm]) {
                const grp = repeatGroups[norm];
                const cur = repeatCursor[norm] || 0;
                const spec = grp[Math.min(cur, grp.length - 1)];
                repeatCursor[norm] = cur + 1;
                canon[i] = spec.canonical;
                specs[i] = spec;
                return;
            }
            /* Otherwise resolve by unique alias (exact → suffix-tolerant). */
            const spec = this._resolveField(h, idx);
            if (!spec) { if (String(h).trim()) unresolved.push(h); return; }
            canon[i] = spec.canonical;
            specs[i] = spec;
        });

        return { canon, specs, unresolved };
    },

    /* Validate required columns against a resolved canon list.
       Returns { ok, missing:[labels], warnings:[] }. */
    _validateHeaders(canon, model) {
        const present = new Set(canon.filter(Boolean));
        const missing = model
            .filter(s => s.required && !present.has(s.canonical))
            .map(s => s.label);
        return { ok: missing.length === 0, missing, warnings: [] };
    },

    /* Coerce one raw cell to the type declared by its field spec.
       `native` is true for XLSX numeric cells (already a JS number). */
    _coerceValue(spec, cell, native) {
        const t = spec.type;
        if (t === 'text' || t === 'identifier') {
            return String(cell ?? '').trim();
        }
        if (t === 'date') {
            return this._xlsxDateToDotStr(cell);
        }
        /* numeric families: integer | numeric | duration_s | duration_ms |
           data_kb | data_byte | percentage */
        if (native && typeof cell === 'number') {
            return isFinite(cell) ? cell : null;   /* XLSX native number — already normalized */
        }
        return this._parseNum(String(cell ?? '').trim());   /* European-format string */
    },

    /* Coerce a whole data row (array of raw cells) using the resolved specs. */
    _coerceRow(dataCells, specs, native) {
        const row = {};
        specs.forEach((spec, ci) => {
            if (!spec) return;
            row[spec.canonical] = this._coerceValue(spec, dataCells[ci], native);
        });
        return row;
    },

    /* Build the reusable validation-metadata block for an upload. */
    _buildValidationMeta(type, rawHeaders, canon, specs, validation, recordCount) {
        const resolved = [];
        rawHeaders.forEach((h, i) => {
            if (!canon[i]) return;
            resolved.push({ raw: String(h).trim(), canonical: canon[i], type: specs[i] ? specs[i].type : null, label: specs[i] ? specs[i].label : null });
        });
        return {
            type,
            resolved,
            missing: validation.missing || [],
            warnings: validation.warnings || [],
            recordCount,
        };
    },

    /* ── Normalize a raw header string to its canonical name ── */
    _canonicalize(raw, aliasMap) {
        const key = raw.trim().toLowerCase()
            .replace(/\s+/g, ' ')
            /* strip trailing unit suffixes to help fuzzy match */
            .replace(/\s*\(ms\)\s*$/, '')
            .replace(/\s*\(kb\)\s*$/, '')
            .replace(/\s*\(s\)\s*$/, '');
        /* Exact match first */
        if (aliasMap[key]) return aliasMap[key];
        /* Try the key with unit suffix present */
        const keyWithUnit = raw.trim().toLowerCase().replace(/\s+/g, ' ');
        if (aliasMap[keyWithUnit]) return aliasMap[keyWithUnit];
        /* Partial/prefix match — last resort */
        for (const [alias, canon] of Object.entries(aliasMap)) {
            if (key.startsWith(alias) || alias.startsWith(key)) return canon;
        }
        return '';   // unmapped column — kept but ignored
    },

    /* Reverse lookup: canonical → first matching alias label (for error messages) */
    _canonicalToLabel(canon, aliasMap) {
        for (const [alias, c] of Object.entries(aliasMap)) {
            if (c === canon) return alias;
        }
        return canon;
    },

    /* Parse European-format numbers: 1.234.567,89 → 1234567.89 */
    _parseNum(s) {
        if (s === '' || s == null) return null;
        const clean = String(s).trim()
            .replace(/\./g, '')   // remove thousands separator (dot)
            .replace(',', '.');   // convert decimal comma to dot
        const n = parseFloat(clean);
        return isNaN(n) ? null : n;
    },

    /* ── Store a successfully parsed upload ── */
    _storeUpload(key, result, fileName, statusId) {
        this._uploads[key] = { ...result, fileName };
        showFileLoaded(statusId, fileName + (result.meta ? ` (${result.meta})` : ''));
        showToast('Loaded', `${fileName} — ${result.rows.length} rows${result.meta ? ', ' + result.meta : ''}`, 'success', 3000);
        /* Publish the reusable validation metadata so later import-processing
           phases can reuse the resolved schema without re-sniffing headers. */
        this._publishValidationMeta();
    },

    /* Aggregate every profile upload's validation metadata into module data. */
    _publishValidationMeta() {
        if (!this.app || !this.app.setModuleData) return;
        const existing = this.app.getModuleData('app-workload') || {};
        const vmeta = {};
        ['dialog', 'http', 'background', 'rfc'].forEach(k => {
            const u = this._uploads[k];
            if (u && u.validation) vmeta[k] = u.validation;
        });
        this.app.setModuleData('app-workload', { ...existing, validationMeta: vmeta });
    },


    /* ── Validation error UI ── */
    _showValidationError(html) {
        const el = document.getElementById('aw-validation-error');
        const msg = document.getElementById('aw-validation-msg');
        if (el && msg) { msg.innerHTML = html; el.style.display = 'flex'; el.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); }
        showToast('Validation Error', 'One or more required columns are missing — see the error message above.', 'error', 8000);
    },

    _clearValidationError() {
        const el = document.getElementById('aw-validation-error');
        if (el) el.style.display = 'none';
    },

    /* ── Next button enable/disable ── */
    _refreshNextBtn(btnId, keys, minRequired) {
        const loaded   = keys.filter(k => this._uploads[k] != null).length;
        const required = minRequired ?? keys.length;
        const btn = document.getElementById(btnId);
        if (btn) btn.disabled = loaded < required;
    },

    /* ── Sub-nav ── */
    _updateSubNav(app) {
        const step = this._step;
        const u    = this._uploads;
        const items = [
            {
                id: 'aw-step1', label: 'Step 1 – Application Workload', step: true,
                onClick: () => { this._step = 1; this._updateSubNav(app); app.navigateTo('app-workload'); },
            },
            {
                id: 'aw-step2', label: 'Step 2 – DB vs CPU Analysis', step: true,
                onClick: () => {
                    if (u.dbcpu && u.tasktype) { this._step = 2; this._updateSubNav(app); app.navigateTo('app-workload'); }
                },
            },
            {
                id: 'aw-step3', label: 'Step 3 – Transaction Profile', step: true,
                onClick: () => {
                    if (u.dbcpu && u.tasktype) { this._step = 3; this._updateSubNav(app); app.navigateTo('app-workload'); }
                },
            },
            {
                id: 'aw-step4', label: 'Step 4 – Charts', step: true,
                onClick: () => {
                    const hasProfile = ['dialog', 'http', 'background', 'rfc'].some(k => u[k] != null);
                    if (hasProfile) { this._step = 4; this._updateSubNav(app); app.navigateTo('app-workload'); }
                },
            },
        ];
        const activeId = step === 1 ? 'aw-step1' : step === 2 ? 'aw-step2'
                       : step === 3 ? 'aw-step3' : 'aw-step4';
        app.setSubNav('app-workload', items, activeId);
    },

    /* ── Save module data ── */
    _saveModuleData(app) {
        const u = this._uploads;
        const allRows = [];
        ['dbcpu', 'tasktype', 'dialog', 'http', 'background', 'rfc'].forEach(k => {
            if (u[k]) u[k].rows.forEach(r => allRows.push({ _section: k, ...r }));
        });
        const vmeta = {};
        ['dialog', 'http', 'background', 'rfc'].forEach(k => {
            if (u[k] && u[k].validation) vmeta[k] = u[k].validation;
        });
        app.setModuleData('app-workload', {
            rows: allRows, fileName: ['dbcpu','tasktype','dialog','http','background','rfc']
                .filter(k => u[k]).map(k => u[k].fileName).join(', '),
            headers: Object.keys(allRows[0] || {}), analysis: null, uploads: u,
            validationMeta: vmeta,
        });
    },

    /* ══════════════════════════════════════════════════════
       CHART RENDERING — Step 3
    ══════════════════════════════════════════════════════ */

    /* ── Load History: line chart — X = date, Y = resp/cpu/db time, secondary Y = steps ── */
    _mountDbcpuCharts(container, upload) {
        if (!container || !upload) return;
        const { rows, meta } = upload;

        /* Sort chronologically */
        const sorted = rows.slice().sort((a, b) => {
            const da = this._parseDate(a.date), db = this._parseDate(b.date);
            return da - db;
        });

        const labels    = sorted.map(r => r.date);
        const steps     = sorted.map(r => r.steps ?? 0);
        const respTime  = sorted.map(r => r.resp_ms ?? 0);
        const cpuTime   = sorted.map(r => r.cpu_ms ?? 0);
        const dbTime    = sorted.map(r => r.db_ms ?? 0);
        const procTime  = sorted.map(r => r.proc_ms ?? 0);

        const metaHtml  = meta ? `<div class="chart-note" style="margin-bottom:10px">${meta}</div>` : '';

        container.innerHTML = metaHtml + `
            <div class="charts-grid" style="grid-template-columns:1fr;margin-bottom:16px">
                <div class="chart-card">
                    <div class="chart-title">Response, CPU &amp; DB Time per Month (ms) vs Dialog Steps</div>
                    <div class="chart-wrap" style="height:300px"><canvas id="aw-chart-dbcpu-main"></canvas></div>
                </div>
            </div>
            <div class="charts-grid" style="grid-template-columns:repeat(auto-fill,minmax(320px,1fr));margin-bottom:16px">
                <div class="chart-card">
                    <div class="chart-title">Dialog Steps per Month</div>
                    <div class="chart-wrap" style="height:220px"><canvas id="aw-chart-dbcpu-steps"></canvas></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">Avg Response Time (ms)</div>
                    <div class="chart-wrap" style="height:220px"><canvas id="aw-chart-dbcpu-resp"></canvas></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">Avg CPU Time (ms)</div>
                    <div class="chart-wrap" style="height:220px"><canvas id="aw-chart-dbcpu-cpu"></canvas></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">Avg DB Time (ms)</div>
                    <div class="chart-wrap" style="height:220px"><canvas id="aw-chart-dbcpu-db"></canvas></div>
                </div>
            </div>
            ${this._dataTable(sorted, ['date','steps','resp_ms','cpu_ms','db_ms','proc_ms','wait_ms','rollin_ms'], 'Load History Data')}`;

        /* Main combo chart — dual axis */
        this._createChart('aw-chart-dbcpu-main', {
            type: 'bar',
            data: {
                labels,
                datasets: [
                    { type: 'bar',  label: 'Dialog Steps',      data: steps,    backgroundColor: CHART_COLORS[0] + '99', yAxisID: 'ySteps', order: 2 },
                    { type: 'line', label: 'Avg Response (ms)',  data: respTime, borderColor: CHART_COLORS[3], backgroundColor: 'transparent', yAxisID: 'yTime', tension: 0.3, pointRadius: 3, order: 1 },
                    { type: 'line', label: 'Avg CPU (ms)',       data: cpuTime,  borderColor: CHART_COLORS[2], backgroundColor: 'transparent', yAxisID: 'yTime', tension: 0.3, pointRadius: 3, order: 1 },
                    { type: 'line', label: 'Avg DB (ms)',        data: dbTime,   borderColor: CHART_COLORS[1], backgroundColor: 'transparent', yAxisID: 'yTime', tension: 0.3, pointRadius: 3, order: 1 },
                ],
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { position: 'bottom', labels: { font: { size: 11 } } } },
                scales: {
                    x: { grid: { display: false }, ticks: { maxRotation: 35, font: { size: 10 } } },
                    ySteps: { type: 'linear', position: 'left',  beginAtZero: true, grid: { color: '#F1F5F9' }, title: { display: true, text: 'Dialog Steps', font: { size: 10 } } },
                    yTime:  { type: 'linear', position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: 'Time (ms)', font: { size: 10 } } },
                },
            },
        });

        this._createBarChart('aw-chart-dbcpu-steps', labels, steps,    'Dialog Steps',  CHART_COLORS[0]);
        this._createBarChart('aw-chart-dbcpu-resp',  labels, respTime, 'Avg Resp (ms)', CHART_COLORS[3]);
        this._createBarChart('aw-chart-dbcpu-cpu',   labels, cpuTime,  'Avg CPU (ms)',  CHART_COLORS[2]);
        this._createBarChart('aw-chart-dbcpu-db',    labels, dbTime,   'Avg DB (ms)',   CHART_COLORS[1]);
        this._addSectionFontRow(container, 'aw-aux-dbcpu');
    },

    /* ── Workload by Task Type: horizontal bar charts ── */
    _mountTasktypeCharts(container, upload) {
        if (!container || !upload) return;
        const { rows, meta } = upload;

        const labels   = rows.map(r => r.task_type || '—');
        const steps    = rows.map(r => r.steps    ?? 0);
        const resp     = rows.map(r => r.resp_ms  ?? 0);
        const cpu      = rows.map(r => r.cpu_ms   ?? 0);
        const db       = rows.map(r => r.db_ms    ?? 0);

        const metaHtml = meta ? `<div class="chart-note" style="margin-bottom:10px">${meta}</div>` : '';

        container.innerHTML = metaHtml + `
            <div class="charts-grid" style="grid-template-columns:repeat(auto-fill,minmax(340px,1fr));margin-bottom:16px">
                <div class="chart-card">
                    <div class="chart-title">Dialog Steps by Task Type</div>
                    <div class="chart-wrap" style="height:280px"><canvas id="aw-chart-tt-steps"></canvas></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">Avg Response Time by Task Type (ms)</div>
                    <div class="chart-wrap" style="height:280px"><canvas id="aw-chart-tt-resp"></canvas></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">Avg CPU Time by Task Type (ms)</div>
                    <div class="chart-wrap" style="height:280px"><canvas id="aw-chart-tt-cpu"></canvas></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">Avg DB Time by Task Type (ms)</div>
                    <div class="chart-wrap" style="height:280px"><canvas id="aw-chart-tt-db"></canvas></div>
                </div>
            </div>
            ${this._dataTable(rows, ['task_type','steps','resp_ms','cpu_ms','db_ms','proc_ms','wait_ms','rfc_ms'], 'Workload by Task Type')}`;

        this._createHBarChart('aw-chart-tt-steps', labels, steps, 'Steps',    CHART_COLORS[0]);
        this._createHBarChart('aw-chart-tt-resp',  labels, resp,  'Resp (ms)',CHART_COLORS[3]);
        this._createHBarChart('aw-chart-tt-cpu',   labels, cpu,   'CPU (ms)', CHART_COLORS[2]);
        this._createHBarChart('aw-chart-tt-db',    labels, db,    'DB (ms)',  CHART_COLORS[1]);
        this._addSectionFontRow(container, 'aw-aux-tasktype');
    },

    /* ── Time per Task Type: stacked bar — each metric as a series, X = task type ── */
    _mountTimePerTaskChart(container, upload) {
        if (!container || !upload) return;
        const { rows, meta } = upload;

        const labels  = rows.map(r => r.task_type || '—');

        /* Build time-component stacks (exclude steps/non-time fields) */
        const timeFields = [
            { field: 'proc_ms',    label: 'Processing',  color: CHART_COLORS[0] },
            { field: 'db_ms',      label: 'DB',          color: CHART_COLORS[1] },
            { field: 'cpu_ms',     label: 'CPU',         color: CHART_COLORS[2] },
            { field: 'wait_ms',    label: 'Wait',        color: CHART_COLORS[3] },
            { field: 'rfc_ms',     label: 'RFC',         color: CHART_COLORS[4] },
            { field: 'rollin_ms',  label: 'Roll In',     color: CHART_COLORS[5] },
            { field: 'loadgen_ms', label: 'Load+Gen',    color: CHART_COLORS[6] },
            { field: 'gui_ms',     label: 'GUI',         color: CHART_COLORS[7] },
        ].filter(f => rows.some(r => (r[f.field] ?? 0) > 0));

        /* Also compute response vs processing gap */
        const respVsProc = rows.map(r => Math.max(0, (r.resp_ms ?? 0) - (r.proc_ms ?? 0)));

        const metaHtml = meta ? `<div class="chart-note" style="margin-bottom:10px">${meta}</div>` : '';

        container.innerHTML = metaHtml + `
            <div class="charts-grid" style="grid-template-columns:1fr;margin-bottom:16px">
                <div class="chart-card">
                    <div class="chart-title">Time Components per Task Type (ms) — stacked</div>
                    <div class="chart-wrap" style="height:320px"><canvas id="aw-chart-tptt-stack"></canvas></div>
                </div>
            </div>
            <div class="charts-grid" style="grid-template-columns:repeat(auto-fill,minmax(340px,1fr));margin-bottom:16px">
                <div class="chart-card">
                    <div class="chart-title">Avg Response Time (ms)</div>
                    <div class="chart-wrap" style="height:240px"><canvas id="aw-chart-tptt-resp"></canvas></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">Response vs Processing Time Gap (ms)</div>
                    <div class="chart-wrap" style="height:240px"><canvas id="aw-chart-tptt-gap"></canvas></div>
                </div>
            </div>`;

        /* Stacked chart */
        this._createChart('aw-chart-tptt-stack', {
            type: 'bar',
            data: {
                labels,
                datasets: timeFields.map(f => ({
                    label: f.label,
                    data: rows.map(r => r[f.field] ?? 0),
                    backgroundColor: f.color + 'CC',
                    borderRadius: 0,
                    borderSkipped: false,
                })),
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { position: 'bottom', labels: { font: { size: 10 } } } },
                scales: {
                    x: { stacked: true, grid: { display: false }, ticks: { maxRotation: 35, font: { size: 10 } } },
                    y: { stacked: true, beginAtZero: true, grid: { color: '#F1F5F9' }, title: { display: true, text: 'Time (ms)', font: { size: 10 } } },
                },
            },
        });

        this._createHBarChart('aw-chart-tptt-resp', labels, rows.map(r => r.resp_ms ?? 0), 'Resp (ms)', CHART_COLORS[3]);
        this._createHBarChart('aw-chart-tptt-gap',  labels, respVsProc,                     'Gap (ms)',  CHART_COLORS[4]);
        this._addSectionFontRow(container, 'aw-aux-timeptype');
    },

    /* ── Transaction Profile charts (Dialog / HTTP / Background / RFC) ──
       X-axis = transaction/report name; 4 top-N horizontal bar charts
       + totals combo chart + data table.
       Background uploads also carry a job_name column shown in the table.
    */
    _mountProfileCharts(container, upload, key) {
        if (!container || !upload) return;
        const { rows, meta } = upload;

        if (!rows.length) {
            container.innerHTML = `<div class="chart-note">No data rows found.</div>`;
            return;
        }

        const isRfc = key === 'rfc';

        /* For RFC the label field may not be trans_rep yet — fall back gracefully */
        const labelField = rows[0].trans_rep !== undefined ? 'trans_rep'
                         : rows[0].task_type !== undefined ? 'task_type' : null;
        const hasJobName = rows.some(r => r.job_name != null && r.job_name !== '');

        /* Sort by total response time desc (or steps if resp unavailable) */
        const sorted = rows.slice().sort((a, b) =>
            ((b.resp_total_s ?? b.resp_ms ?? b.steps ?? 0)) -
            ((a.resp_total_s ?? a.resp_ms ?? a.steps ?? 0))
        );

        /* Top-N for charts (cap at 20 for readability) */
        const TOP = Math.min(sorted.length, 20);
        const top = sorted.slice(0, TOP);
        const labels = top.map(r => labelField ? (r[labelField] || '—') : '—');

        const p = key.toUpperCase();
        const metaHtml = meta ? `<div class="chart-note" style="margin-bottom:10px">${meta}</div>` : '';

        /* RFC Workload — dedicated renderer using system_id / action / user_id schema */
        if (isRfc) {
            this._mountRfcCharts(container, upload);
            return;
        }

        container.innerHTML = metaHtml + `
            <!-- Combo: total times bar + avg response line -->
            <div class="charts-grid" style="grid-template-columns:1fr;margin-bottom:16px">
                <div class="chart-card">
                    <div class="chart-title">Top ${TOP} by Total Response Time (s) — with Avg Response Time (ms)</div>
                    <div class="chart-wrap" style="height:300px"><canvas id="aw-chart-${p}-combo"></canvas></div>
                </div>
            </div>
            <div class="charts-grid" style="grid-template-columns:repeat(auto-fill,minmax(320px,1fr));margin-bottom:16px">
                <div class="chart-card">
                    <div class="chart-title">Top ${TOP} by Dialog Steps</div>
                    <div class="chart-wrap" style="height:280px"><canvas id="aw-chart-${p}-steps"></canvas></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">Avg Response Time / Step (ms)</div>
                    <div class="chart-wrap" style="height:280px"><canvas id="aw-chart-${p}-resp"></canvas></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">Avg CPU Time / Step (ms)</div>
                    <div class="chart-wrap" style="height:280px"><canvas id="aw-chart-${p}-cpu"></canvas></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">Avg DB Time / Step (ms)</div>
                    <div class="chart-wrap" style="height:280px"><canvas id="aw-chart-${p}-db"></canvas></div>
                </div>
            </div>
            ${this._dataTable(sorted,
                hasJobName
                    ? ['trans_rep','job_name','steps','resp_total_s','resp_ms','cpu_total_s','cpu_ms','db_total_s','db_ms','proc_total_s','proc_ms','wait_ms','rollwait_ms','db_kb']
                    : ['trans_rep','steps','resp_total_s','resp_ms','cpu_total_s','cpu_ms','db_total_s','db_ms','proc_total_s','proc_ms','wait_ms','rollwait_ms','db_kb'],
                'Transaction Profile Data')}`;

        /* Combo chart — dual axis */
        this._createChart(`aw-chart-${p}-combo`, {
            type: 'bar',
            data: {
                labels,
                datasets: [
                    { type: 'bar',  label: 'Total Resp (s)',  data: top.map(r => r.resp_total_s ?? 0), backgroundColor: CHART_COLORS[0] + '99', yAxisID: 'yTotal',  order: 2 },
                    { type: 'line', label: 'Avg Resp (ms)',   data: top.map(r => r.resp_ms      ?? 0), borderColor: CHART_COLORS[3], backgroundColor: 'transparent', yAxisID: 'yAvg', tension: 0.3, pointRadius: 4, order: 1 },
                ],
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { position: 'bottom', labels: { font: { size: 11 } } } },
                scales: {
                    x: { grid: { display: false }, ticks: { maxRotation: 45, font: { size: 9 } } },
                    yTotal: { type: 'linear', position: 'left',  beginAtZero: true, grid: { color: '#F1F5F9' }, title: { display: true, text: 'Total (s)',  font: { size: 10 } } },
                    yAvg:   { type: 'linear', position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, title: { display: true, text: 'Avg (ms)', font: { size: 10 } } },
                },
            },
        });

        this._createHBarChart(`aw-chart-${p}-steps`, labels, top.map(r => r.steps    ?? 0), 'Steps',     CHART_COLORS[0]);
        this._createHBarChart(`aw-chart-${p}-resp`,  labels, top.map(r => r.resp_ms  ?? 0), 'Resp (ms)', CHART_COLORS[3]);
        this._createHBarChart(`aw-chart-${p}-cpu`,   labels, top.map(r => r.cpu_ms   ?? 0), 'CPU (ms)',  CHART_COLORS[2]);
        this._createHBarChart(`aw-chart-${p}-db`,    labels, top.map(r => r.db_ms    ?? 0), 'DB (ms)',   CHART_COLORS[1]);
        this._addSectionFontRow(container, `aw-aux-${key}`);
    },
    /* ── RFC Workload charts ──
       Aggregation: System – Action – User
       Key dims: system_id (calling system), action (RFC function/transaction), user_id
       Charts:
         1. Top N actions — combo: total response time bar + avg response time line
         2. Top N actions by steps (hbar)
         3. Top N actions by avg response time (hbar)
         4. Top N actions by avg CPU time (hbar)
         5. Top N actions by avg DB time (hbar)
         6. RFC steps by calling system (hbar)
         7. Data table (all columns)
    */
    _mountRfcCharts(container, upload) {
        if (!container || !upload) return;
        const { rows, meta } = upload;

        if (!rows.length) {
            container.innerHTML = `<div class="chart-note">No data rows found.</div>`;
            return;
        }

        const metaHtml = meta ? `<div class="chart-note" style="margin-bottom:10px">${meta}</div>` : '';

        /* Aggregate by action (sum steps/totals, derive weighted averages) */
        const byAction = {};
        for (const r of rows) {
            const key = (r.action || '—').trim();
            if (!byAction[key]) {
                byAction[key] = {
                    action: key, steps: 0,
                    resp_total_s: 0, cpu_total_s: 0, db_total_s: 0,
                    wait_total_s: 0, db_proc_total_s: 0, db_kb: 0,
                };
            }
            const a = byAction[key];
            a.steps           += r.steps           ?? 0;
            a.resp_total_s    += r.resp_total_s     ?? 0;
            a.cpu_total_s     += r.cpu_total_s      ?? 0;
            a.db_total_s      += r.db_total_s       ?? 0;
            a.wait_total_s    += r.wait_total_s     ?? 0;
            a.db_proc_total_s += r.db_proc_total_s  ?? 0;
            a.db_kb           += r.db_kb            ?? 0;
        }
        for (const a of Object.values(byAction)) {
            a.resp_ms = a.steps > 0 ? (a.resp_total_s  * 1000 / a.steps) : 0;
            a.cpu_ms  = a.steps > 0 ? (a.cpu_total_s   * 1000 / a.steps) : 0;
            a.db_ms   = a.steps > 0 ? (a.db_total_s    * 1000 / a.steps) : 0;
        }

        /* Aggregate by calling system */
        const bySystem = {};
        for (const r of rows) {
            const key = (r.system_id || '—').trim();
            if (!bySystem[key]) bySystem[key] = { system_id: key, steps: 0, resp_total_s: 0 };
            bySystem[key].steps        += r.steps        ?? 0;
            bySystem[key].resp_total_s += r.resp_total_s ?? 0;
        }

        const actionRows = Object.values(byAction).sort((a, b) => b.resp_total_s - a.resp_total_s);
        const systemRows = Object.values(bySystem).sort((a, b) => b.steps - a.steps);

        const TOP    = Math.min(actionRows.length, 20);
        const SYS_TOP = Math.min(systemRows.length, 15);
        const top    = actionRows.slice(0, TOP);
        const topSys = systemRows.slice(0, SYS_TOP);
        const topLabels = top.map(r => r.action);

        container.innerHTML = metaHtml + `
            <div class="charts-grid" style="grid-template-columns:1fr;margin-bottom:16px">
                <div class="chart-card">
                    <div class="chart-title">Top ${TOP} Actions — Total Response Time (s) &amp; Avg Response Time (ms)</div>
                    <div class="chart-wrap" style="height:300px"><canvas id="aw-chart-RFC-combo"></canvas></div>
                </div>
            </div>
            <div class="charts-grid" style="grid-template-columns:repeat(auto-fill,minmax(320px,1fr));margin-bottom:16px">
                <div class="chart-card">
                    <div class="chart-title">Top ${TOP} Actions by # Steps</div>
                    <div class="chart-wrap" style="height:280px"><canvas id="aw-chart-RFC-steps"></canvas></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">Top ${TOP} Actions — Avg Response Time (ms)</div>
                    <div class="chart-wrap" style="height:280px"><canvas id="aw-chart-RFC-resp"></canvas></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">Top ${TOP} Actions — Avg CPU Time (ms)</div>
                    <div class="chart-wrap" style="height:280px"><canvas id="aw-chart-RFC-cpu"></canvas></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">Top ${TOP} Actions — Avg DB Time (ms)</div>
                    <div class="chart-wrap" style="height:280px"><canvas id="aw-chart-RFC-db"></canvas></div>
                </div>
            </div>
            <div class="charts-grid" style="grid-template-columns:1fr;margin-bottom:16px">
                <div class="chart-card">
                    <div class="chart-title">RFC Steps by Calling System (top ${SYS_TOP})</div>
                    <div class="chart-wrap" style="height:${Math.max(200, SYS_TOP * 28)}px"><canvas id="aw-chart-RFC-sys"></canvas></div>
                </div>
            </div>
            ${this._dataTable(rows,
                ['system_id','action','user_id','steps','resp_total_s','resp_ms',
                 'cpu_total_s','cpu_ms','db_total_s','db_ms','db_proc_total_s','db_proc_ms',
                 'wait_total_s','wait_ms','roundtrips','frontend_ms','gui_ms','db_kb'],
                'RFC Workload Data')}`;

        /* Combo chart */
        this._createChart('aw-chart-RFC-combo', {
            type: 'bar',
            data: {
                labels: topLabels,
                datasets: [
                    { type: 'bar',  label: 'Total Resp (s)', data: top.map(r => r.resp_total_s ?? 0),
                      backgroundColor: CHART_COLORS[0] + '99', yAxisID: 'yTotal', order: 2 },
                    { type: 'line', label: 'Avg Resp (ms)',  data: top.map(r => r.resp_ms      ?? 0),
                      borderColor: CHART_COLORS[3], backgroundColor: 'transparent',
                      yAxisID: 'yAvg', tension: 0.3, pointRadius: 4, order: 1 },
                ],
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { position: 'bottom', labels: { font: { size: 11 } } } },
                scales: {
                    x: { grid: { display: false }, ticks: { maxRotation: 45, font: { size: 9 } } },
                    yTotal: { type: 'linear', position: 'left',  beginAtZero: true,
                              grid: { color: '#F1F5F9' },
                              title: { display: true, text: 'Total (s)', font: { size: 10 } } },
                    yAvg:   { type: 'linear', position: 'right', beginAtZero: true,
                              grid: { drawOnChartArea: false },
                              title: { display: true, text: 'Avg (ms)', font: { size: 10 } } },
                },
            },
        });

        this._createHBarChart('aw-chart-RFC-steps', topLabels, top.map(r => r.steps   ?? 0), 'Steps',     CHART_COLORS[0]);
        this._createHBarChart('aw-chart-RFC-resp',  topLabels, top.map(r => r.resp_ms ?? 0), 'Resp (ms)', CHART_COLORS[3]);
        this._createHBarChart('aw-chart-RFC-cpu',   topLabels, top.map(r => r.cpu_ms  ?? 0), 'CPU (ms)',  CHART_COLORS[2]);
        this._createHBarChart('aw-chart-RFC-db',    topLabels, top.map(r => r.db_ms   ?? 0), 'DB (ms)',   CHART_COLORS[1]);
        this._createHBarChart('aw-chart-RFC-sys',
            topSys.map(r => r.system_id),
            topSys.map(r => r.steps),
            'Steps', CHART_COLORS[4]);
        this._addSectionFontRow(container, 'aw-aux-rfc');
    },

    _createChart(id, config) {
        const ctx = document.getElementById(id)?.getContext('2d');
        if (!ctx) return null;
        const chart = new Chart(ctx, config);
        this._charts.push(chart);
        return chart;
    },

    _createBarChart(id, labels, values, label, color) {
        this._createChart(id, {
            type: 'bar',
            data: { labels, datasets: [{ label, data: values, backgroundColor: color + 'CC', borderRadius: 4, borderSkipped: false }] },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { beginAtZero: true, grid: { color: '#F1F5F9' } },
                    x: { grid: { display: false }, ticks: { maxRotation: 35, font: { size: 10 } } },
                },
            },
        });
    },

    _createHBarChart(id, labels, values, label, color) {
        /* Sort descending for readability */
        const pairs = labels.map((l, i) => ({ l, v: values[i] })).sort((a, b) => b.v - a.v);
        this._createChart(id, {
            type: 'bar',
            data: {
                labels: pairs.map(p => p.l),
                datasets: [{ label, data: pairs.map(p => p.v), backgroundColor: color + 'CC', borderRadius: 4, borderSkipped: false }],
            },
            options: {
                indexAxis: 'y',
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    x: { beginAtZero: true, grid: { color: '#F1F5F9' } },
                    y: { grid: { display: false }, ticks: { font: { size: 10 } } },
                },
            },
        });
    },

    /* Canonical field → display label */
    _DISPLAY_NAMES: {
        date: 'Date', task_type: 'Task Type',
        trans_rep: 'Trans / Report', job_name: 'Background Job',
        system_id: 'System ID', action: 'Action / Transaction', user_id: 'User ID',
        steps: 'Steps',
        resp_total_s: 'Total Resp (s)', resp_ms: 'Avg Resp (ms)',
        proc_total_s: 'Total Proc (s)', proc_ms: 'Avg Proc (ms)',
        cpu_total_s:  'Total CPU (s)',  cpu_ms:  'Avg CPU (ms)',
        db_total_s:   'Total DB (s)',   db_ms:   'Avg DB (ms)',
        db_proc_total_s: 'DB Proc Total (s)', db_proc_ms: 'DB Proc Avg (ms)',
        rollwait_total_s: 'Roll Wait Total (s)', rollwait_ms: 'Roll Wait Avg (ms)',
        wait_total_s: 'Total Wait (s)', wait_ms: 'Avg Wait (ms)',
        roundtrips: 'Roundtrips',
        frontend_ms: 'Frontend (ms)', gui_ms: 'GUI (ms)',
        db_kb: 'DB Data (KB)', input_byte: 'Avg Input (B)', output_byte: 'Avg Output (B)',
        rollin_ms: 'Roll In (ms)', loadgen_ms: 'Load+Gen (ms)', lock_ms: 'Lock (ms)', rfc_ms: 'RFC (ms)',
        vmc_calls: 'VMC Calls', vmc_cpu_s: 'VMC CPU (s)', vmc_elaps_s: 'VMC Elap (s)',
        vmc_cpu_avg_ms: 'VMC CPU Avg (ms)', vmc_elap_avg_ms: 'VMC Elap Avg (ms)',
    },

    /* Data table — uses _DISPLAY_NAMES for column headers */
    _dataTable(rows, fields, title) {
        const visible = fields.filter(f => rows.some(r => r[f] != null && r[f] !== ''));
        if (!visible.length) return '';
        const fmt = v => v == null ? '' : typeof v === 'number'
            ? v.toLocaleString('en-US', { maximumFractionDigits: 1 }) : v;
        const dn = this._DISPLAY_NAMES;
        return `
            <div class="table-wrap" style="border-radius:10px;margin-top:8px">
                <table class="data-table">
                    <thead><tr>${visible.map(f => `<th>${dn[f] || f}</th>`).join('')}</tr></thead>
                    <tbody>${rows.slice(0, 60).map(r =>
                        `<tr>${visible.map(f => `<td>${fmt(r[f])}</td>`).join('')}</tr>`
                    ).join('')}</tbody>
                </table>
            </div>
            ${rows.length > 60 ? `<div class="chart-note" style="margin-top:6px">Showing 60 of ${rows.length} rows.</div>` : ''}`;
    },

    /* Parse a date string like "02/01/2024" or "01/02/2024" to a sortable number */
    _parseDate(s) {
        if (!s) return 0;
        const p = String(s).trim().split(/[\/\-]/);
        if (p.length < 3) return 0;
        /* Assume DD/MM/YYYY or MM/DD/YYYY — both ok for sorting since year dominates */
        return parseInt(p[2], 10) * 10000 + parseInt(p[1], 10) * 100 + parseInt(p[0], 10);
    },
};

/* ── Backward-compat stubs ── */
function _userCategory(n) { return n < 50 ? 'Small (<50)' : n < 200 ? 'Medium (50–200)' : n < 1000 ? 'Large (200–1000)' : 'Enterprise (>1000)'; }
function _txnCategory(n)  { return n < 1000 ? 'Low' : n < 10000 ? 'Medium' : n < 50000 ? 'High' : 'Very High'; }
