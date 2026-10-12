/* ─────────────────────────────────────────────────────────
   Module 05 – Sizing Projection
   ─────────────────────────────────────────────────────────
   Aggregates data from all prior modules and generates
   hardware sizing recommendations plus growth projections.
   Also supports uploading an optional additional file
   (e.g. growth plan, future workload forecast).
   ─────────────────────────────────────────────────────────*/

const SizingProjectionModule = {

    _charts: {},

    render() {
        return `
        <div class="module-wrap">
            <div class="module-header">
                <div class="module-crumb">Analysis › Step 05</div>
                <h1 class="module-title">Sizing Projection</h1>
                <p class="module-desc">
                    Consolidated sizing recommendations based on all collected data.
                    Review the hardware projections, apply growth factors, and download
                    the final sizing report.
                </p>
            </div>

            <!-- Data completeness panel -->
            <div class="card" style="margin-bottom:18px">
                <div class="card-header">
                    <div>
                        <div class="card-title">Input Data Completeness</div>
                        <div class="card-sub">Modules with data loaded feed into the projection</div>
                    </div>
                    <button class="btn btn-blue" id="sp-btn-calc">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2">
                            <polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>
                        </svg>
                        Generate Projection
                    </button>
                </div>
                <div id="sp-completeness"></div>
            </div>

            <!-- Optional growth file upload -->
            <div class="card" style="margin-bottom:18px">
                <div class="card-header">
                    <div>
                        <div class="card-title">Growth Plan Upload <span class="badge badge-gray" style="font-size:10px;margin-left:6px">Optional</span></div>
                        <div class="card-sub">Upload a growth forecast or future workload estimate</div>
                    </div>
                </div>
                <div class="upload-zone" id="sp-zone" style="padding:24px">
                    <input type="file" class="file-input-hidden" id="sp-input" accept=".csv,.xlsx,.xls">
                    <div class="upload-zone-title" style="font-size:14px">Upload Growth Forecast (optional)</div>
                    <div class="upload-zone-sub">CSV or XLSX with Year and growth factor columns</div>
                    <span class="upload-formats">CSV · XLSX · XLS</span>
                </div>
                <div id="sp-status"></div>
            </div>

            <!-- Projection results -->
            <div id="sp-results" style="display:none">

                <div class="section-title">Sizing Recommendations</div>
                <div class="rec-grid" id="sp-recs"></div>

                <div class="section-title">Projection Charts</div>
                <div class="charts-grid">
                    <div class="chart-card">
                        <div class="chart-title">Current vs. Recommended Capacity</div>
                        <div class="chart-wrap"><canvas id="sp-chart-capacity"></canvas></div>
                    </div>
                    <div class="chart-card">
                        <div class="chart-title">3-Year Growth Projection</div>
                        <div class="chart-wrap"><canvas id="sp-chart-growth"></canvas></div>
                    </div>
                    <div class="chart-card">
                        <div class="chart-title">Resource Utilization Overview</div>
                        <div class="chart-wrap"><canvas id="sp-chart-util"></canvas></div>
                    </div>
                    <div class="chart-card">
                        <div class="chart-title">Module Data Coverage</div>
                        <div class="chart-wrap"><canvas id="sp-chart-coverage"></canvas></div>
                    </div>
                </div>

                <div class="section-title">Detailed Projection Table</div>
                <div class="card" style="padding:0">
                    <div class="card-header" style="padding:14px 18px 12px; border-radius:14px 14px 0 0">
                        <div>
                            <div class="card-title">Hardware Sizing Projection</div>
                            <div class="card-sub">All values include 20% headroom buffer</div>
                        </div>
                        <button class="btn" id="sp-btn-download-proj">
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                                 fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="7 10 12 15 17 10"/>
                                <line x1="12" y1="15" x2="12" y2="3"/>
                            </svg>
                            Download Full Report
                        </button>
                    </div>
                    <div class="table-wrap" style="border:none;border-radius:0 0 14px 14px">
                        <table class="data-table" id="sp-proj-table">
                            <thead>
                                <tr>
                                    <th>Component</th>
                                    <th>Current (Est.)</th>
                                    <th>Recommended</th>
                                    <th>With 20% Buffer</th>
                                    <th>3-Year Target</th>
                                    <th>Status</th>
                                </tr>
                            </thead>
                            <tbody id="sp-proj-tbody"></tbody>
                        </table>
                    </div>
                </div>

                <div class="action-row">
                    <button class="btn btn-blue" id="sp-btn-export-full">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24"
                             fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="7 10 12 15 17 10"/>
                            <line x1="12" y1="15" x2="12" y2="3"/>
                        </svg>
                        Export Complete Sizing Report
                    </button>
                    <button class="btn" id="sp-btn-restart">
                        Start New Analysis
                    </button>
                </div>
            </div>
        </div>`;
    },

    init(app) {
        this._charts = {};
        // Show completeness status immediately
        this._renderCompleteness(app);

        // Optional growth file
        initUploadZone('sp-zone', 'sp-input', file => {
            parseUploadedFile(file, (rows, name) => {
                showFileLoaded('sp-status', name);
                app.setModuleData('sizing-projection-growth', { rows, fileName: name });
            });
        });

        document.getElementById('sp-btn-calc').addEventListener('click', () => {
            this._generateProjection(app);
        });
    },

    _renderCompleteness(app) {
        const modules = [
            { id: 'system-landscape',  label: 'System Landscape' },
            { id: 'tech-workload-app', label: 'Technical Workload (App)' },
            { id: 'tech-workload-db',  label: 'Technical Workload (DB)' },
            { id: 'app-workload',      label: 'Application Workload' },
        ];

        const completed = app.getCompletedModules();

        const rows = modules.map(m => {
            const done = completed.has(m.id);
            const data = app.getModuleData(m.id);
            return `
                <div style="display:flex;align-items:center;gap:12px;padding:9px 0;
                            border-bottom:1px solid var(--border);">
                    <div style="width:22px;height:22px;border-radius:50%;
                                background:${done ? 'var(--success)' : 'var(--border)'};
                                display:flex;align-items:center;justify-content:center;flex-shrink:0">
                        ${done ? `<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24"
                                      fill="none" stroke="white" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>`
                                : `<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24"
                                       fill="none" stroke="var(--text-muted)" stroke-width="2">
                                       <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                                   </svg>`}
                    </div>
                    <span style="font-size:13px;font-weight:500;color:${done ? 'var(--text-primary)' : 'var(--text-muted)'}">
                        ${m.label}
                    </span>
                    <span style="margin-left:auto;font-size:12px;color:var(--text-muted)">
                        ${done && data ? `${data.rows?.length ?? 0} records` : done ? 'Loaded' : 'No data'}
                    </span>
                    <span class="badge ${done ? 'badge-green' : 'badge-gray'}">${done ? 'Ready' : 'Missing'}</span>
                </div>`;
        });

        document.getElementById('sp-completeness').innerHTML =
            `<div style="padding:4px 0">${rows.join('')}</div>
             <p style="font-size:12px;color:var(--text-muted);margin-top:12px">
                ${completed.size} of 4 input modules loaded.
                ${completed.size < 4 ? 'You can still generate a projection from partial data.' : 'All modules ready.'}
             </p>`;
    },

    _generateProjection(app) {
        const slData  = app.getModuleData('system-landscape');
        const appData = app.getModuleData('tech-workload-app');
        const dbData  = app.getModuleData('tech-workload-db');
        const wlData  = app.getModuleData('app-workload');

        // Derive sizing signals
        const systemCount = slData?.rows?.length ?? 0;

        // App CPU P95
        const appAnalysis = appData?.analysis ?? [];
        const cpuP95Row   = appAnalysis.find(r => /cpu/i.test(r.Metric)) || {};
        const cpuP95      = parseFloat(cpuP95Row.P95) || 70;
        const memP95Row   = appAnalysis.find(r => /mem/i.test(r.Metric)) || {};
        const memP95      = parseFloat(memP95Row.P95) || 75;

        // DB
        const dbRows   = dbData?.rows ?? [];
        const dbHdrs   = dbData?.headers ?? [];
        const sizeCol  = detectCol(dbHdrs, /size(_?gb)?|data_?size/i);
        const totalDbGB = sizeCol
            ? dbRows.reduce((s, r) => s + (parseFloat(r[sizeCol]) || 0), 0)
            : 0;

        // Users
        const totals   = wlData?.totals ?? {};
        const totalUsers = totals.users ?? 0;

        // Sizing model (simplified, demonstrative)
        const HEADROOM   = 1.20;  // 20 % buffer
        const GROWTH_3Y  = 1.40;  // 40 % 3-year growth assumption

        const estAppCPU  = Math.ceil(cpuP95 * HEADROOM);
        const estAppMem  = Math.ceil(memP95 * HEADROOM);
        const estDbSize  = Math.ceil(totalDbGB * HEADROOM) || 0;
        const appCores   = Math.max(4, Math.ceil((cpuP95 / 80) * systemCount * 2));
        const appMemGB   = Math.max(16, Math.ceil(memP95 / 75 * 32));
        const dbMemGB    = Math.max(32, Math.ceil(totalDbGB * 0.5));

        const BUFFER     = 1.20;
        const rows3y     = [
            { Component: 'Application Servers – CPU',   Current: `${appCores} vCPUs`,  Recommended: `${Math.ceil(appCores * BUFFER)} vCPUs`, WithBuffer: `${Math.ceil(appCores * BUFFER)} vCPUs`, Target3Y: `${Math.ceil(appCores * BUFFER * GROWTH_3Y)} vCPUs`, _status: cpuP95 > 80 ? 'critical' : cpuP95 > 60 ? 'warn' : 'ok' },
            { Component: 'Application Servers – Memory', Current: `${appMemGB} GB`,    Recommended: `${Math.ceil(appMemGB * BUFFER)} GB`,  WithBuffer: `${Math.ceil(appMemGB * BUFFER)} GB`,  Target3Y: `${Math.ceil(appMemGB * BUFFER * GROWTH_3Y)} GB`,  _status: memP95 > 85 ? 'critical' : memP95 > 65 ? 'warn' : 'ok' },
            { Component: 'Database – Memory',            Current: `${dbMemGB} GB`,     Recommended: `${Math.ceil(dbMemGB * BUFFER)} GB`,   WithBuffer: `${Math.ceil(dbMemGB * BUFFER)} GB`,   Target3Y: `${Math.ceil(dbMemGB * BUFFER * GROWTH_3Y)} GB`,  _status: 'ok' },
            { Component: 'Database – Storage',           Current: `${Math.max(0, totalDbGB).toFixed(0)} GB`, Recommended: `${Math.ceil(totalDbGB * BUFFER || 100)} GB`, WithBuffer: `${Math.ceil(totalDbGB * BUFFER || 100)} GB`, Target3Y: `${Math.ceil((totalDbGB * BUFFER || 100) * GROWTH_3Y)} GB`, _status: 'ok' },
            { Component: 'User Licences (concurrent ~10%)', Current: `${totalUsers}`,  Recommended: `${Math.ceil(totalUsers * 0.10)} concurrent`, WithBuffer: `${Math.ceil(totalUsers * 0.12)} concurrent`, Target3Y: `${Math.ceil(totalUsers * 0.12 * GROWTH_3Y)} concurrent`, _status: 'ok' },
        ];

        // Recommendations
        const recs = [
            { label: 'Recommended App vCPUs', value: `${Math.ceil(appCores * BUFFER)}`,  desc: `Based on P95 CPU ${cpuP95.toFixed(0)}%` },
            { label: 'Recommended App Memory', value: `${Math.ceil(appMemGB * BUFFER)} GB`, desc: `Based on P95 Mem ${memP95.toFixed(0)}%` },
            { label: 'Recommended DB Storage', value: `${Math.ceil(totalDbGB * BUFFER || 100)} GB`, desc: `Total DB + 20% buffer` },
            { label: '3-Year DB Storage Target', value: `${Math.ceil((totalDbGB * BUFFER || 100) * GROWTH_3Y)} GB`, desc: '40% growth assumption' },
        ];

        document.getElementById('sp-recs').innerHTML = recs.map(r => `
            <div class="rec-card">
                <div class="rec-card-label">${r.label}</div>
                <div class="rec-card-value">${r.value}</div>
                <div class="rec-card-desc">${r.desc}</div>
            </div>`).join('');

        // Table
        const statusBadge = s =>
            s === 'critical' ? '<span class="badge badge-red">Critical</span>' :
            s === 'warn'     ? '<span class="badge badge-yellow">Review</span>' :
                               '<span class="badge badge-green">OK</span>';

        document.getElementById('sp-proj-tbody').innerHTML = rows3y.map(r => `
            <tr>
                <td><strong>${r.Component}</strong></td>
                <td>${r.Current}</td>
                <td>${r.Recommended}</td>
                <td>${r.WithBuffer}</td>
                <td>${r.Target3Y}</td>
                <td>${statusBadge(r._status)}</td>
            </tr>`).join('');

        // Charts
        this._capacityChart(appCores, appMemGB, dbMemGB);
        this._growthChart(totalDbGB, appMemGB);
        this._utilChart(cpuP95, memP95);
        this._coverageChart(app);

        // Show results
        document.getElementById('sp-results').style.display = 'block';
        document.getElementById('sp-results').scrollIntoView({ behavior: 'smooth', block: 'start' });

        // Save & mark complete
        const analysisRows = rows3y.map(({ _status, ...rest }) => rest);
        app.setModuleData('sizing-projection', { rows: analysisRows, fileName: 'generated', headers: Object.keys(analysisRows[0]), analysis: analysisRows });
        app.markCompleted('sizing-projection');

        // Download
        document.getElementById('sp-btn-download-proj').addEventListener('click', () => {
            downloadXlsx(analysisRows, 'Sizing Projection', `sizing-projection-${_dateTag()}.xlsx`);
        });

        document.getElementById('sp-btn-export-full').addEventListener('click', () => {
            app.exportReport();
        });

        document.getElementById('sp-btn-restart').addEventListener('click', () => {
            app.confirmReset();
        });
    },

    _capacityChart(cpuCores, memGB, dbMemGB) {
        const canvas = document.getElementById('sp-chart-capacity');
        const ctx = canvas?.getContext('2d');
        if (!ctx) return;
        if (this._charts.capacity) { try { this._charts.capacity.destroy(); } catch(_) {} }
        this._charts.capacity = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['App vCPUs', 'App Memory (GB)', 'DB Memory (GB)'],
                datasets: [
                    { label: 'Current Estimate', data: [cpuCores, memGB, dbMemGB],
                      backgroundColor: CHART_COLORS[0] + '88', borderRadius: 3 },
                    { label: 'Recommended (w/ buffer)', data: [Math.ceil(cpuCores*1.2), Math.ceil(memGB*1.2), Math.ceil(dbMemGB*1.2)],
                      backgroundColor: CHART_COLORS[0], borderRadius: 3 },
                ],
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { position: 'top', labels: { font: { size: 11 } } } },
                scales: {
                    y: { beginAtZero: true, grid: { color: '#F1F5F9' } },
                    x: { grid: { display: false } },
                },
            },
        });
        this._addFontRow(canvas, 'sp-cap', false);
    },

    _growthChart(dbGB, memGB) {
        const canvas = document.getElementById('sp-chart-growth');
        const ctx = canvas?.getContext('2d');
        if (!ctx) return;
        const base = Math.max(dbGB, 100);
        const years = ['Current', 'Year 1', 'Year 2', 'Year 3'];
        const rates = [1, 1.15, 1.28, 1.40];
        if (this._charts.growth) { try { this._charts.growth.destroy(); } catch(_) {} }
        this._charts.growth = new Chart(ctx, {
            type: 'line',
            data: {
                labels: years,
                datasets: [
                    {
                        label: 'DB Storage (GB)',
                        data: rates.map(r => Math.ceil(base * r)),
                        borderColor: CHART_COLORS[0], backgroundColor: CHART_COLORS[0] + '18',
                        tension: 0.3, fill: true, pointRadius: 5,
                    },
                    {
                        label: 'App Memory (GB)',
                        data: rates.map(r => Math.ceil(memGB * r)),
                        borderColor: CHART_COLORS[1], backgroundColor: CHART_COLORS[1] + '18',
                        tension: 0.3, fill: false, pointRadius: 5,
                    },
                ],
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { position: 'top', labels: { font: { size: 11 } } } },
                scales: {
                    y: { beginAtZero: true, grid: { color: '#F1F5F9' } },
                    x: { grid: { display: false } },
                },
            },
        });
        this._addFontRow(canvas, 'sp-grow', false);
    },

    _utilChart(cpuP95, memP95) {
        const canvas = document.getElementById('sp-chart-util');
        const ctx = canvas?.getContext('2d');
        if (!ctx) return;
        if (this._charts.util) { try { this._charts.util.destroy(); } catch(_) {} }
        this._charts.util = new Chart(ctx, {
            type: 'radar',
            data: {
                labels: ['CPU Utilization', 'Memory Utilization', 'Storage Load', 'I/O Pressure', 'User Concurrency'],
                datasets: [{
                    label: 'Current Load (%)',
                    data: [cpuP95, memP95, 55, 48, 62],
                    borderColor: CHART_COLORS[0],
                    backgroundColor: CHART_COLORS[0] + '28',
                    pointBackgroundColor: CHART_COLORS[0],
                }],
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    r: { beginAtZero: true, max: 100, ticks: { stepSize: 25, font: { size: 10 } } },
                },
            },
        });
        this._addFontRow(canvas, 'sp-util', true);
    },

    _coverageChart(app) {
        const canvas = document.getElementById('sp-chart-coverage');
        const ctx = canvas?.getContext('2d');
        if (!ctx) return;

        const modules = [
            { label: 'System Landscape',   id: 'system-landscape' },
            { label: 'App Workload',        id: 'tech-workload-app' },
            { label: 'DB Workload',         id: 'tech-workload-db' },
            { label: 'Biz Workload',        id: 'app-workload' },
        ];

        const completed = app.getCompletedModules();
        const values    = modules.map(m => completed.has(m.id) ? 100 : 0);

        if (this._charts.coverage) { try { this._charts.coverage.destroy(); } catch(_) {} }
        this._charts.coverage = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: modules.map(m => m.label),
                datasets: [{
                    label: 'Data Coverage (%)',
                    data: values,
                    backgroundColor: values.map(v => v === 100 ? CHART_COLORS[1] : '#E2E8F0'),
                    borderRadius: 4, borderSkipped: false,
                }],
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { beginAtZero: true, max: 100, grid: { color: '#F1F5F9' },
                         ticks: { callback: v => v + '%' } },
                    x: { grid: { display: false } },
                },
            },
        });
        this._addFontRow(canvas, 'sp-cov', false);
    },

    _addFontRow(canvas, uid, _isPie) {
        if (!canvas) return;
        const wrap = canvas.parentElement;
        const existing = wrap.nextElementSibling;
        if (existing && existing.classList.contains('chart-font-row')) existing.remove();
        wrap.insertAdjacentHTML('afterend', chartFontRowHtml(uid));
        const card = wrap.closest('.chart-card') || wrap.parentElement;
        const getAllRows = () => {
            const rows = [];
            const section = card.closest('.charts-grid') || document.getElementById('sp-results');
            if (section) section.querySelectorAll('[id$="-font-row"]').forEach(r => rows.push(r));
            return rows;
        };
        wireChartFontRow(card, () => Object.values(this._charts), getAllRows, null);
    },
};
