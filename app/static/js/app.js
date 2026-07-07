// ============================================================
// 全局状态
// ============================================================
let currentView = 'execute';    // execute | search | summary
let currentIndex = 0;
let totalCount = 0;
let dirty = false;
let selectedResult = '';
let allStatuses = [];

let searchPage = 0;
let searchPageSize = 30;
let searchLastParams = null;

// ============================================================
// 视图切换
// ============================================================
function switchView(view) {
    if (dirty && view !== 'execute') {
        showUnsavedDialog().then(confirmed => {
            if (confirmed) _doSwitch(view);
        });
        return;
    }
    _doSwitch(view);
}

function _doSwitch(view) {
    currentView = view;

    document.getElementById('viewExecute').style.display = view === 'execute' ? 'block' : 'none';
    document.getElementById('viewSearch').style.display = view === 'search' ? 'block' : 'none';
    document.getElementById('viewSummary').style.display = view === 'summary' ? 'block' : 'none';
    var issuesDiv = document.getElementById('viewIssues');
    if (issuesDiv) issuesDiv.style.display = view === 'issues' ? 'block' : 'none';

    document.getElementById('tabExecute').classList.toggle('active', view === 'execute');
    document.getElementById('tabSearch').classList.toggle('active', view === 'search');
    document.getElementById('tabSummary').classList.toggle('active', view === 'summary');
    var tabIssues = document.getElementById('tabIssues');
    if (tabIssues) tabIssues.classList.toggle('active', view === 'issues');

    if (view === 'search') {
        loadFilterOptions();
        searchPage = 0;
        doSearch();
    } else if (view === 'summary') {
        loadSummary();
    } else if (view === 'issues') {
        loadIssues();
    }

    window.scrollTo({top: 0, behavior: 'smooth'});
}

// ============================================================
// 初始化
// ============================================================
async function init() {
    try {
        const resp = await fetch('/api/init');
        const data = await resp.json();

        if (!data.loaded) {
            // 没有活跃文件 → 显示上传页
            if (data.has_active) {
                // 有文件但读不到数据（可能是空 Excel）
                showEmptyState();
            } else {
                showUploadPage();
            }
            return;
        }

        totalCount = data.total;
        document.getElementById('topbarFilename').textContent =
            '📂 ' + (data.filename || '已加载');
        document.getElementById('progressWrap').style.display = 'block';
        document.getElementById('navBar').style.display = 'flex';
        document.getElementById('navHint').style.display = 'block';

        await refreshStatuses();
        await loadCaseTitles();
        await loadTestCase(0);
    } catch (err) {
        showEmptyState();
        console.error('初始化失败:', err);
    }
}

function showEmptyState() {
    const container = document.getElementById('execContainer');
    container.innerHTML = `
        <div class="state-message">
            <div class="icon">&#128203;</div>
            <h2>未读取到用例数据</h2>
            <p>当前活跃文件中未检测到有效用例数据。<br>请确保第一行为表头，从第二行开始为用例数据。<br>你也可以重新上传一个新的测试用例文件。</p>
            <button class="btn btn-save" style="margin-top:16px;" onclick="showReplaceDialog()">
                &#128194; 上传新文件
            </button>
        </div>
    `;
}

// ============================================================
// 上传页
// ============================================================
function showUploadPage() {
    const container = document.getElementById('execContainer');
    document.getElementById('progressWrap').style.display = 'none';
    document.getElementById('navBar').style.display = 'none';
    document.getElementById('navHint').style.display = 'none';

    container.innerHTML = `
        <div class="upload-page" id="uploadPage">
            <div class="upload-card" id="uploadCard">
                <div class="upload-icon">🚀</div>
                <h2>欢迎使用测试用例记录表</h2>
                <p class="upload-desc">请上传你的测试用例 Excel 文件开始使用<br>支持 .xlsx 和 .xls 格式</p>
                <div class="upload-zone" id="uploadZone">
                    <div class="upload-zone-icon">📂</div>
                    <div class="upload-zone-text">拖拽 Excel 文件到此处</div>
                    <div class="upload-zone-hint">或点击下方按钮选择文件</div>
                    <input type="file" id="fileInput" accept=".xlsx,.xls" style="display:none"
                           onchange="handleFileSelect(this)">
                    <button class="btn btn-save" onclick="document.getElementById('fileInput').click()"
                            style="margin-top:16px;">
                        📁 选择文件
                    </button>
                </div>
                <div class="upload-status" id="uploadStatus" style="display:none;"></div>
            </div>
        </div>
        <style>
            .upload-page {
                display: flex; align-items: center; justify-content: center;
                min-height: 70vh; padding: 40px 20px;
            }
            .upload-card {
                background: var(--surface); border-radius: var(--radius);
                box-shadow: var(--shadow-lg); padding: 44px 40px;
                max-width: 520px; width: 100%; text-align: center;
                border: 1px solid var(--border-light);
            }
            .upload-icon { font-size: 56px; margin-bottom: 16px; }
            .upload-card h2 { font-size: 24px; margin-bottom: 10px; color: var(--text); font-weight: 800; }
            .upload-desc { font-size: 14px; color: var(--text-secondary); margin-bottom: 32px; line-height: 1.7; }
            .upload-zone {
                border: 2px dashed var(--border); border-radius: var(--radius);
                padding: 36px 24px; transition: all .25s;
                cursor: pointer; background: var(--bg);
            }
            .upload-zone:hover, .upload-zone.drag-over {
                border-color: var(--primary); background: var(--primary-light);
                transform: translateY(-2px); box-shadow: var(--shadow-md);
            }
            .upload-zone-icon { font-size: 44px; margin-bottom: 10px; }
            .upload-zone-text { font-size: 17px; font-weight: 700; color: var(--text); margin-bottom: 6px; }
            .upload-zone-hint { font-size: 13px; color: var(--text-secondary); }
            .upload-status { margin-top: 18px; padding: 12px; border-radius: var(--radius-sm); font-size: 14px; font-weight: 600; }
            .upload-status.loading { background: #eff6ff; color: #1d4ed8; }
            .upload-status.error { background: #fef2f2; color: #991b1b; }
        </style>
    `;

    // 拖拽事件
    const zone = document.getElementById('uploadZone');
    if (zone) {
        zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('drag-over'); });
        zone.addEventListener('dragleave', () => { zone.classList.remove('drag-over'); });
        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('drag-over');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                uploadFile(files[0]);
            }
        });
    }
}

async function handleFileSelect(input) {
    if (input.files.length > 0) {
        await uploadFile(input.files[0]);
    }
}

async function uploadFile(file) {
    // 检查文件类型
    const name = file.name.toLowerCase();
    if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
        showUploadStatus('不支持的文件格式，请上传 .xlsx 或 .xls 文件', 'error');
        return;
    }

    // 检查是否已有测试结果 → 二次确认
    try {
        const checkResp = await fetch('/api/check-results');
        const checkData = await checkResp.json();
        if (checkData.has_results) {
            const confirmed = await showReplaceConfirmDialog(checkData.filename);
            if (!confirmed) return;
        }
    } catch (err) {
        // 检查失败不影响上传
    }

    showUploadStatus('⏳ 正在上传并解析文件...', 'loading');

    const formData = new FormData();
    formData.append('file', file);

    try {
        const resp = await fetch('/api/upload', { method: 'POST', body: formData });
        const data = await resp.json();

        if (data.success) {
            showUploadStatus('✅ 上传成功！共 ' + data.total + ' 条用例', '');
            totalCount = data.total;
            document.getElementById('topbarFilename').textContent =
                '📂 ' + (data.filename || '已加载');
            document.getElementById('progressWrap').style.display = 'block';
            document.getElementById('navBar').style.display = 'flex';
            document.getElementById('navHint').style.display = 'block';

            // 短暂延迟后切换到执行页
            setTimeout(async () => {
                await refreshStatuses();
                await loadCaseTitles();
                await loadTestCase(0);
            }, 600);
        } else {
            showUploadStatus('❌ ' + (data.error || '上传失败'), 'error');
        }
    } catch (err) {
        showUploadStatus('❌ 网络错误: ' + err.message, 'error');
    }
}

function showUploadStatus(msg, type) {
    const el = document.getElementById('uploadStatus');
    if (!el) return;
    el.style.display = 'block';
    el.textContent = msg;
    el.className = 'upload-status ' + type;
}

// ============================================================
// 更换文件对话框
// ============================================================
function showReplaceDialog() {
    // 构造一个隐藏的 file input
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls';
    input.onchange = async () => {
        if (input.files.length > 0) {
            // 检查是否有结果
            try {
                const checkResp = await fetch('/api/check-results');
                const checkData = await checkResp.json();
                if (checkData.has_results) {
                    const confirmed = await showReplaceConfirmDialog(checkData.filename);
                    if (!confirmed) return;
                }
            } catch (err) {}
            await uploadFile(input.files[0]);
        }
    };
    input.click();
}

function showReplaceConfirmDialog(filename) {
    return new Promise((resolve) => {
        const overlay = document.createElement('div');
        overlay.className = 'modal-overlay';
        overlay.innerHTML = `
            <div class="modal-box">
                <h3>⚠️ 确认替换文件</h3>
                <p>当前活跃文件 <strong>${escapeHtml(filename)}</strong> 中已有测试执行结果。<br><br>
                   上传新文件后将覆盖当前文件，<br><strong>已有执行结果将会丢失</strong>。<br>
                   如需保留，请先备份原文件。</p>
                <div class="modal-actions">
                    <button class="btn-modal-ghost" id="modalCancel">取消</button>
                    <button class="btn-modal-primary" id="modalConfirm"
                            style="background:#dc2626;">确认替换</button>
                </div>
            </div>`;
        document.body.appendChild(overlay);

        overlay.querySelector('#modalCancel').onclick = () => {
            document.body.removeChild(overlay);
            resolve(false);
        };
        overlay.querySelector('#modalConfirm').onclick = () => {
            document.body.removeChild(overlay);
            resolve(true);
        };
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) { document.body.removeChild(overlay); resolve(false); }
        });
    });
}

// 全局都可以触发更换文件（从 topbar 按钮）
async function triggerReplaceFile() {
    if (dirty) {
        const confirmed = await showUnsavedDialog();
        if (!confirmed) return;
    }
    showReplaceDialog();
}

// ============================================================
// 执行测试 - 加载用例
// ============================================================
async function loadTestCase(index) {
    if (dirty) {
        const confirmed = await showUnsavedDialog();
        if (!confirmed) return;
    }

    try {
        const resp = await fetch('/api/testcase/' + index);
        if (!resp.ok) throw new Error('加载失败');

        const tc = await resp.json();
        currentIndex = index;
        selectedResult = tc._saved_result || '';
        renderTestCase(tc);
        updateNavButtons();
        syncCaseJumpSelect(index);
        dirty = false;
        window.scrollTo({top: 0, behavior: 'smooth'});
    } catch (err) {
        showToast('加载用例失败: ' + err.message, 'error');
    }
}

function renderTestCase(tc) {
    const container = document.getElementById('execContainer');
    const fields = tc._display_fields || [];
    const total = tc._total || totalCount;

    let fieldsHtml = '';
    for (const f of fields) {
        let valueHtml = escapeHtml(f.value);
        if (f.css_class.includes('field-steps')) {
            valueHtml = formatSteps(f.value);
        }
        fieldsHtml += `
            <div class="field-row ${f.css_class}">
                <div class="field-label">${f.label}</div>
                <div class="field-value">${valueHtml}</div>
            </div>`;
    }

    const savedBadge = tc._saved_result
        ? `<span class="result-badge result-badge-${resultCss(tc._saved_result)}">已执行: ${escapeHtml(tc._saved_result)}</span>`
        : '';

    container.innerHTML = `
        <div class="card">
            <div class="card-header">
                <div class="card-navigator">
                    第 <strong>${tc._index + 1}</strong> 条 / 共 ${total} 条
                </div>
                ${savedBadge}
            </div>
            <div class="card-body">${fieldsHtml}</div>
        </div>

        <div class="action-card">
            <h3>&#9989; 测试结果</h3>
            <div class="result-group">
                <button class="result-btn ${selectedResult === '通过' ? 'selected-pass' : ''}"
                        onclick="selectResult('通过', this)">
                    <span class="result-emoji">&#9989;</span>通过
                </button>
                <button class="result-btn ${selectedResult === '失败' ? 'selected-fail' : ''}"
                        onclick="selectResult('失败', this)">
                    <span class="result-emoji">&#10060;</span>失败
                </button>
                <button class="result-btn ${selectedResult === '阻塞' ? 'selected-block' : ''}"
                        onclick="selectResult('阻塞', this)">
                    <span class="result-emoji">&#128683;</span>阻塞
                </button>
                <button class="result-btn ${selectedResult === '跳过' ? 'selected-skip' : ''}"
                        onclick="selectResult('跳过', this)">
                    <span class="result-emoji">&#9193;</span>跳过
                </button>
            </div>

            <h3 style="margin-bottom:8px">&#128221; 实际结果</h3>
            <textarea class="remark-area" id="remarkInput"
                      placeholder="请描述测试执行过程中的实际现象、发现的问题等..."
                      oninput="markDirty()">${escapeHtml(tc._saved_actual_result || '')}</textarea>

            <!-- 新增字段 -->
            <div class="extra-fields" style="margin-top:14px;display:grid;grid-template-columns:1fr 1fr;gap:10px;">
                <div>
                    <label class="extra-label">&#128100; 测试人员</label>
                    <input type="text" class="extra-input" id="inputTester" placeholder="填写测试人员姓名"
                           value="${escapeHtml(tc._saved_tester || '')}" oninput="markDirty()">
                </div>
                <div>
                    <label class="extra-label">&#128030; BugID</label>
                    <input type="text" class="extra-input" id="inputBugId" placeholder="如 PROJ-001"
                           value="${escapeHtml(tc._saved_bug_id || '')}" oninput="markDirty()">
                </div>
                <div>
                    <label class="extra-label">&#128257; Bug频率</label>
                    <select class="extra-input" id="inputBugFreq" onchange="markDirty()">
                        <option value="">-- 请选择 --</option>
                        <option value="必现" ${tc._saved_bug_frequency === '必现' ? 'selected' : ''}>必现</option>
                        <option value="高频" ${tc._saved_bug_frequency === '高频' ? 'selected' : ''}>高频</option>
                        <option value="偶现" ${tc._saved_bug_frequency === '偶现' ? 'selected' : ''}>偶现</option>
                        <option value="1次" ${tc._saved_bug_frequency === '1次' ? 'selected' : ''}>1次</option>
                    </select>
                </div>
                <div>
                    <label class="extra-label">&#128339; 问题时间</label>
                    <input type="text" class="extra-input" id="inputIssueTime" placeholder="选择失败时自动记录"
                           value="${escapeHtml(tc._saved_issue_time || '')}" oninput="markDirty()">
                </div>
            </div>

            <div style="margin-top:14px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
                <button class="btn btn-save" id="btnSave" onclick="saveResult()">
                    &#128190; 确认并保存
                </button>
                <span id="saveStatus" style="font-size:13px;color:var(--text-secondary);"></span>
            </div>
        </div>`;

    document.getElementById('remarkInput').addEventListener('keydown', function(e) {
        if ((e.ctrlKey || e.metaKey) && e.key === 's') {
            e.preventDefault();
            saveResult();
        }
    });
}

function formatSteps(text) {
    const lines = text.split('\n').filter(l => l.trim() !== '');
    if (lines.length <= 1) return escapeHtml(text).replace(/\n/g, '<br>');
    let html = '<ol style="padding-left:20px;margin:0;">';
    for (const line of lines) {
        html += `<li style="margin-bottom:4px;">${escapeHtml(line.trim())}</li>`;
    }
    html += '</ol>';
    return html;
}

function selectResult(value, btnEl) {
    selectedResult = value;
    markDirty();
    document.querySelectorAll('.result-btn').forEach(b => b.className = 'result-btn');
    const classMap = {'通过':'selected-pass','失败':'selected-fail','阻塞':'selected-block','跳过':'selected-skip'};
    btnEl.className = 'result-btn ' + (classMap[value] || '');

    // 失败时自动记录问题时间
    if (value === '失败') {
        const now = new Date();
        const timeStr = now.getFullYear() + '-' +
            String(now.getMonth()+1).padStart(2,'0') + '-' +
            String(now.getDate()).padStart(2,'0') + ' ' +
            String(now.getHours()).padStart(2,'0') + ':' +
            String(now.getMinutes()).padStart(2,'0') + ':' +
            String(now.getSeconds()).padStart(2,'0');
        const issueInput = document.getElementById('inputIssueTime');
        if (issueInput) issueInput.value = timeStr;
    } else {
        // 非失败时清空问题时间
        const issueInput = document.getElementById('inputIssueTime');
        if (issueInput) issueInput.value = '';
    }
}

function resultCss(r) {
    const m = {'通过':'pass','失败':'fail','阻塞':'block','跳过':'skip'};
    return m[r] || '';
}

function markDirty() { dirty = true; }

// ============================================================
// 保存结果
// ============================================================
async function saveResult() {
    if (!selectedResult) {
        showToast('请先选择测试结果', 'warn');
        return;
    }
    const actualResult = document.getElementById('remarkInput').value.trim();
    const tester = document.getElementById('inputTester') ? document.getElementById('inputTester').value.trim() : '';
    const bugId = document.getElementById('inputBugId') ? document.getElementById('inputBugId').value.trim() : '';
    const bugFreq = document.getElementById('inputBugFreq') ? document.getElementById('inputBugFreq').value : '';
    const issueTime = document.getElementById('inputIssueTime') ? document.getElementById('inputIssueTime').value.trim() : '';
    const btnSave = document.getElementById('btnSave');
    const saveStatus = document.getElementById('saveStatus');
    btnSave.disabled = true;
    btnSave.textContent = '⏳ 保存中...';

    try {
        const resp = await fetch('/api/save', {
            method:'POST',
            headers:{'Content-Type':'application/json'},
            body:JSON.stringify({
                index:currentIndex, result:selectedResult,
                actual_result: actualResult,
                tester: tester, bug_id: bugId,
                bug_frequency: bugFreq, issue_time: issueTime,
            }),
        });
        const data = await resp.json();
        if (data.success) {
            dirty = false;
            showToast('✅ 已保存！测试结果已写入 Excel', 'success');
            saveStatus.textContent = '✅ 已保存 ' + new Date().toLocaleTimeString('zh-CN');
            await refreshStatuses();
            allStatuses[currentIndex] = selectedResult;
            // 保存后自动跳转到下一条
            const nextIndex = currentIndex + 1;
            if (nextIndex < totalCount) {
                await loadTestCase(nextIndex);
            } else {
                showToast('🎉 已到达最后一条用例', 'success');
            }
        } else {
            showToast('❌ 保存失败: ' + data.error, 'error');
        }
    } catch (err) {
        showToast('❌ 网络错误: ' + err.message, 'error');
    } finally {
        btnSave.disabled = false;
        btnSave.textContent = '💾 确认并保存';
    }
}

// ============================================================
// 导航
// ============================================================
async function goTo(delta) {
    if (dirty) {
        const confirmed = await showUnsavedDialog();
        if (!confirmed) return;
    }
    const newIndex = currentIndex + delta;
    if (newIndex < 0 || newIndex >= totalCount) return;
    loadTestCase(newIndex);
}

function updateNavButtons() {
    document.getElementById('btnPrev').disabled = currentIndex <= 0;
    document.getElementById('btnNext').disabled = currentIndex >= totalCount - 1;
}

async function loadCaseTitles() {
    try {
        const resp = await fetch('/api/titles');
        if (!resp.ok) return;
        const titles = await resp.json();
        const sel = document.getElementById('caseJumpSelect');
        if (!sel) return;
        const placeholder = sel.options[0] ? sel.options[0].outerHTML : '<option value="">&#128269; 选择用例跳转...</option>';
        let html = placeholder;
        for (const item of titles) {
            const label = `${item.id} - ${item.title}`;
            html += `<option value="${item.index}">${escapeHtml(label)}</option>`;
        }
        sel.innerHTML = html;
    } catch (err) {
        console.error('加载用例列表失败:', err);
    }
}

function syncCaseJumpSelect(index) {
    const sel = document.getElementById('caseJumpSelect');
    if (!sel) return;
    sel.value = String(index);
}

// ============================================================
// 进度条
// ============================================================
async function refreshStatuses() {
    try {
        const resp = await fetch('/api/all-status');
        allStatuses = await resp.json();
        renderProgress();
    } catch (err) {
        console.error('获取状态失败:', err);
    }
}

function renderProgress() {
    const total = allStatuses.length;
    if (total === 0) return;
    const counts = {'通过':0,'失败':0,'阻塞':0,'跳过':0,'':0};
    for (const s of allStatuses) {
        if (counts.hasOwnProperty(s)) counts[s]++;
        else counts['']++;
    }
    const executed = total - counts[''];
    const pct = Math.round((executed / total) * 100);

    document.getElementById('progressText').textContent = `已执行 ${executed} / ${total} 条`;
    document.getElementById('progressPercent').textContent = pct + '%';
    document.getElementById('progressFill').style.width = pct + '%';
    document.getElementById('progressFill').className = 'progress-fill' + (pct === 100 ? ' done' : '');
    document.getElementById('countPass').textContent = counts['通过'];
    document.getElementById('countFail').textContent = counts['失败'];
    document.getElementById('countBlock').textContent = counts['阻塞'];
    document.getElementById('countSkip').textContent = counts['跳过'];
    document.getElementById('countNone').textContent = counts[''];
}

// ============================================================
// 搜索筛选
// ============================================================
let searchDebounceTimer = null;

function doSearch() {
    clearTimeout(searchDebounceTimer);
    searchDebounceTimer = setTimeout(_doSearch, 200);
}

async function _doSearch() {
    const keyword = document.getElementById('searchKeyword').value.trim();
    const resultFilter = document.getElementById('filterResult').value;
    const priorityFilter = document.getElementById('filterPriority').value;
    const moduleFilter = document.getElementById('filterModule').value;

    searchLastParams = {keyword, resultFilter, priorityFilter, moduleFilter, page:searchPage};

    try {
        const resp = await fetch('/api/search', {
            method:'POST',
            headers:{'Content-Type':'application/json'},
            body:JSON.stringify({
                keyword, result_filter:resultFilter,
                priority_filter:priorityFilter, module_filter:moduleFilter,
                page:searchPage, page_size:searchPageSize,
            }),
        });
        const data = await resp.json();
        renderSearchResults(data);
    } catch (err) {
        console.error('搜索失败:', err);
    }
}

function renderSearchResults(data) {
    const container = document.getElementById('searchContainer');
    const {results, total_matched, page, total_pages} = data;

    document.getElementById('searchInfo').textContent =
        total_matched > 0 ? `找到 ${total_matched} 条匹配用例` : '';

    if (results.length === 0) {
        container.innerHTML = `
            <div class="state-message">
                <div class="icon">&#128269;</div>
                <h2>无匹配结果</h2>
                <p>尝试更换关键词或放宽筛选条件</p>
            </div>`;
        document.getElementById('searchPagination').style.display = 'none';
        return;
    }

    // 高亮关键词
    const kw = (searchLastParams && searchLastParams.keyword) ? searchLastParams.keyword.toLowerCase() : '';

    let html = '<div class="search-result-list">';
    for (const r of results) {
        const badge = r.result
            ? `<span class="result-badge result-badge-${resultCss(r.result)}">${escapeHtml(r.result)}</span>`
            : '';
        const snippet = r.snippet ? highlightText(r.snippet, kw) : '';
        const titleDisplay = kw ? highlightText(r.name || r.title, kw) : escapeHtml(r.name || r.title);
        const idDisplay = r.id || '#' + (r.index + 1);

        html += `
            <div class="search-result-item" onclick="jumpToExecute(${r.index})" title="点击跳转到该用例">
                <div class="sr-left">
                    <span class="sr-id">${escapeHtml(idDisplay)}</span>
                    <div class="sr-title">${titleDisplay}</div>
                    <div class="sr-meta">
                        ${r.module ? escapeHtml(r.module) + ' &middot; ' : ''}
                        ${r.priority ? escapeHtml(r.priority) : ''}
                    </div>
                    ${snippet ? `<div class="sr-snippet">${snippet}</div>` : ''}
                </div>
                <div class="sr-right">
                    ${badge}
                    <span style="font-size:12px;color:var(--text-secondary);">#${r.index + 1}</span>
                </div>
            </div>`;
    }
    html += '</div>';
    container.innerHTML = html;

    // 分页
    const pagDiv = document.getElementById('searchPagination');
    if (total_pages > 1) {
        pagDiv.style.display = 'flex';
        pagDiv.innerHTML = `
            <button class="btn btn-outline btn-sm" onclick="searchGoPage(0)" ${page === 0 ? 'disabled' : ''}>首页</button>
            <button class="btn btn-outline btn-sm" onclick="searchGoPage(${page - 1})" ${page === 0 ? 'disabled' : ''}>上一页</button>
            <span class="page-info">第 ${page + 1} / ${total_pages} 页（共 ${total_matched} 条）</span>
            <button class="btn btn-outline btn-sm" onclick="searchGoPage(${page + 1})" ${page >= total_pages - 1 ? 'disabled' : ''}>下一页</button>
            <button class="btn btn-outline btn-sm" onclick="searchGoPage(${total_pages - 1})" ${page >= total_pages - 1 ? 'disabled' : ''}>末页</button>
        `;
    } else {
        pagDiv.style.display = 'none';
    }
}

function searchGoPage(p) {
    searchPage = p;
    if (searchLastParams) searchLastParams.page = p;
    _doSearch();
    window.scrollTo({top:0, behavior:'smooth'});
}

function clearSearch() {
    document.getElementById('searchKeyword').value = '';
    document.getElementById('filterResult').value = '';
    document.getElementById('filterPriority').value = '';
    document.getElementById('filterModule').value = '';
    searchPage = 0;
    doSearch();
}

function highlightText(text, keyword) {
    if (!keyword) return escapeHtml(text);
    const escaped = escapeHtml(text);
    const escapedKw = escapeHtml(keyword);
    const regex = new RegExp('(' + escapedKw.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'gi');
    return escaped.replace(regex, '<em class="highlight">$1</em>');
}

function jumpToExecute(index) {
    switchView('execute').then ? switchView('execute').then(() => loadTestCase(index)) : (() => {
        _doSwitch('execute');
        loadTestCase(index);
    })();
}

async function loadFilterOptions() {
    try {
        const resp = await fetch('/api/filter-options');
        const data = await resp.json();
        // 填充优先级下拉
        const priSel = document.getElementById('filterPriority');
        priSel.innerHTML = '<option value="">全部优先级</option>';
        for (const p of data.priorities) {
            priSel.innerHTML += `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`;
        }
        // 填充模块下拉
        const modSel = document.getElementById('filterModule');
        modSel.innerHTML = '<option value="">全部模块</option>';
        for (const m of data.modules) {
            modSel.innerHTML += `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`;
        }
    } catch (err) {
        console.error('加载筛选选项失败:', err);
    }
}

// ============================================================
// 汇总统计
// ============================================================
async function loadSummary() {
    const container = document.getElementById('summaryContainer');
    container.innerHTML = '<div class="state-message"><div class="icon">⏳</div><h2>加载中...</h2></div>';

    try {
        const resp = await fetch('/api/summary');
        const data = await resp.json();
        renderSummary(data);
    } catch (err) {
        container.innerHTML = '<div class="state-message"><div class="icon">❌</div><h2>加载失败</h2></div>';
    }
}

function renderSummary(data) {
    const container = document.getElementById('summaryContainer');
    const {total, counts, by_priority, by_module, execution_rate, pass_rate} = data;

    // 环形进度条 SVG
    const ringSVG = (pct, color) => {
        const r = 26, circ = 2 * Math.PI * r;
        const offset = circ * (1 - pct / 100);
        return `<svg width="64" height="64" viewBox="0 0 64 64">
            <circle class="bg" cx="32" cy="32" r="${r}"/>
            <circle class="fg" cx="32" cy="32" r="${r}"
                stroke="${color}" stroke-dasharray="${circ}" stroke-dashoffset="${offset}"/>
        </svg>`;
    };

    let html = '';

    // 统计卡片
    html += `<div class="summary-grid">
        <div class="summary-stat stat-total"><div class="stat-number">${total}</div><div class="stat-label">用例总数</div></div>
        <div class="summary-stat stat-pass"><div class="stat-number">${counts['通过']}</div><div class="stat-label">通过</div></div>
        <div class="summary-stat stat-fail"><div class="stat-number">${counts['失败']}</div><div class="stat-label">失败</div></div>
        <div class="summary-stat stat-block"><div class="stat-number">${counts['阻塞']}</div><div class="stat-label">阻塞</div></div>
        <div class="summary-stat stat-skip"><div class="stat-number">${counts['跳过']}</div><div class="stat-label">跳过</div></div>
    </div>`;

    // 指标行
    html += `<div class="summary-metrics">
        <div class="metric-card">
            <div class="metric-ring">
                ${ringSVG(execution_rate, '#4f46e5')}
                <div class="pct" style="color:#4f46e5;">${execution_rate}%</div>
            </div>
            <div class="metric-text">
                <h4>执行进度</h4>
                <p>已执行 ${total - counts['未执行']} / ${total} 条</p>
            </div>
        </div>
        <div class="metric-card">
            <div class="metric-ring">
                ${ringSVG(pass_rate, '#16a34a')}
                <div class="pct" style="color:#16a34a;">${pass_rate}%</div>
            </div>
            <div class="metric-text">
                <h4>通过率</h4>
                <p>已执行用例中的通过比例</p>
            </div>
        </div>
    </div>`;

    // 按优先级汇总
    const priKeys = Object.keys(by_priority).sort();
    if (priKeys.length > 0) {
        html += `<div class="summary-table-wrap"><h3>⚡ 按优先级分布</h3>
        <table class="summary-table">
            <tr><th>优先级</th><th>用例数</th><th style="color:var(--pass)">通过</th><th style="color:var(--fail)">失败</th><th style="color:var(--block)">阻塞</th><th>跳过</th><th>未执行</th><th>进度</th></tr>`;
        for (const p of priKeys) {
            const d = by_priority[p];
            const totalP = d['通过'] + d['失败'] + d['阻塞'] + d['跳过'] + d['未执行'];
            const execP = totalP - d['未执行'];
            const pctP = totalP > 0 ? Math.round((execP / totalP) * 100) : 0;
            html += `<tr>
                <td><strong>${escapeHtml(p)}</strong></td>
                <td>${totalP}</td>
                <td class="cell-pass">${d['通过']}</td>
                <td class="cell-fail">${d['失败']}</td>
                <td class="cell-block">${d['阻塞']}</td>
                <td>${d['跳过']}</td>
                <td>${d['未执行']}</td>
                <td class="bar-cell">
                    <div class="mini-bar"><div class="mini-bar-fill pass" style="width:${pctP}%"></div></div>
                    <span style="font-size:11px;">${pctP}%</span>
                </td>
            </tr>`;
        }
        html += `</table></div>`;
    }

    // 按模块汇总
    const modKeys = Object.keys(by_module).sort();
    if (modKeys.length > 0) {
        html += `<div class="summary-table-wrap"><h3>&#128193; 按模块分布</h3>
        <table class="summary-table">
            <tr><th>模块</th><th>用例数</th><th style="color:var(--pass)">通过</th><th style="color:var(--fail)">失败</th><th style="color:var(--block)">阻塞</th><th>跳过</th><th>未执行</th><th>进度</th></tr>`;
        for (const m of modKeys) {
            const d = by_module[m];
            const totalM = d['通过'] + d['失败'] + d['阻塞'] + d['跳过'] + d['未执行'];
            const execM = totalM - d['未执行'];
            const pctM = totalM > 0 ? Math.round((execM / totalM) * 100) : 0;
            html += `<tr>
                <td><strong>${escapeHtml(m)}</strong></td>
                <td>${totalM}</td>
                <td class="cell-pass">${d['通过']}</td>
                <td class="cell-fail">${d['失败']}</td>
                <td class="cell-block">${d['阻塞']}</td>
                <td>${d['跳过']}</td>
                <td>${d['未执行']}</td>
                <td class="bar-cell">
                    <div class="mini-bar"><div class="mini-bar-fill pass" style="width:${pctM}%"></div></div>
                    <span style="font-size:11px;">${pctM}%</span>
                </td>
            </tr>`;
        }
        html += `</table></div>`;
    }

    container.innerHTML = html;
}

// ============================================================
// 未保存提示
// ============================================================
function showUnsavedDialog() {
    return new Promise((resolve) => {
        const overlay = document.createElement('div');
        overlay.className = 'modal-overlay';
        overlay.innerHTML = `
            <div class="modal-box">
                <h3>⚠️ 有未保存的更改</h3>
                <p>当前用例的结果或备注尚未保存，<br>切换后将丢失这些更改。是否先保存？</p>
                <div class="modal-actions">
                    <button class="btn-modal-ghost" id="modalDiscard">不保存，直接切换</button>
                    <button class="btn-modal-primary" id="modalSave">先保存</button>
                </div>
            </div>`;
        document.body.appendChild(overlay);

        overlay.querySelector('#modalSave').onclick = async () => {
            document.body.removeChild(overlay);
            await saveResult();
            resolve(true);
        };
        overlay.querySelector('#modalDiscard').onclick = () => {
            document.body.removeChild(overlay);
            dirty = false;
            resolve(true);
        };
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) { document.body.removeChild(overlay); resolve(false); }
        });
    });
}

// ============================================================
// Toast
// ============================================================
function showToast(msg, type) {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = 'toast toast-' + type;
    toast.textContent = msg;
    container.appendChild(toast);
    setTimeout(() => { if (toast.parentNode) toast.parentNode.removeChild(toast); }, 2300);
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.appendChild(document.createTextNode(str));
    return div.innerHTML;
}

// ============================================================
// 刷新数据（重新加载 Excel）
// ============================================================
async function reloadData() {
    const btn = document.getElementById('btnReload');
    btn.disabled = true;
    btn.textContent = '⏳ 刷新中...';

    try {
        const resp = await fetch('/api/reload');
        const data = await resp.json();

        if (data.success) {
            totalCount = data.total;
            document.getElementById('topbarFilename').textContent =
                '📂 ' + (data.filename || '已加载');
            document.getElementById('progressWrap').style.display = 'block';
            document.getElementById('navBar').style.display = 'flex';
            document.getElementById('navHint').style.display = 'block';

            // 重置到第一条
            currentIndex = 0;
            selectedResult = '';
            dirty = false;

            await refreshStatuses();
            await loadTestCase(0);

            // 如果当前在搜索或汇总页，也刷新
            if (currentView === 'search') {
                await loadFilterOptions();
                searchPage = 0;
                await _doSearch();
            } else if (currentView === 'summary') {
                await loadSummary();
            }

            showToast('✅ 已刷新！共 ' + data.total + ' 条用例', 'success');
        } else {
            showToast('❌ 刷新失败: ' + (data.error || '未知错误'), 'error');
        }
    } catch (err) {
        showToast('❌ 刷新失败: ' + err.message, 'error');
    } finally {
        btn.disabled = false;
        btn.textContent = '↻ 刷新';
    }
}

// ============================================================
// 键盘快捷键
// ============================================================
document.addEventListener('keydown', function(e) {
    if (currentView !== 'execute') return;

    const tag = document.activeElement.tagName;
    const isInput = tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT';

    if (!isInput) {
        if (e.key === 'ArrowLeft') { e.preventDefault(); goTo(-1); return; }
        if (e.key === 'ArrowRight') { e.preventDefault(); goTo(1); return; }
    }

    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        if (currentView === 'execute') saveResult();
    }
});

// ============================================================
// 问题列表
// ============================================================
async function loadIssues() {
    const container = document.getElementById('issuesContainer');
    container.innerHTML = '<div class="state-message"><div class="icon">⏳</div><h2>加载中...</h2></div>';

    try {
        const resp = await fetch('/api/all-status');
        const statuses = await resp.json();
        const total = statuses.length;

        // 统计
        let passCount = 0, failCount = 0, blockCount = 0, skipCount = 0, noneCount = 0;
        const issueIndices = [];
        for (let i = 0; i < statuses.length; i++) {
            const s = statuses[i];
            if (s === '通过') passCount++;
            else if (s === '失败') { failCount++; issueIndices.push({index: i, result: s}); }
            else if (s === '阻塞') { blockCount++; issueIndices.push({index: i, result: s}); }
            else if (s === '跳过') skipCount++;
            else noneCount++;
        }

        // 判断逻辑:
        // 1. 全部通过 (passCount == total) → 庆祝动画
        // 2. 全部执行完毕且没有失败/阻塞 (noneCount == 0 && failCount == 0 && blockCount == 0) → 庆祝动画
        // 3. 有失败或阻塞 → 展示问题列表
        // 4. 还没执行完毕，且没有失败阻塞 → 暂无问题
        const hasIssues = failCount > 0 || blockCount > 0;
        const executed = total - noneCount;
        const allPass = (passCount + skipCount === total && !hasIssues); // 通过+跳过=全部，没有失败阻塞

        if (allPass && noneCount === 0) {
            // 全部执行完毕且无失败阻塞 → 庆祝
            showCelebrate(container);
            return;
        }

        if (hasIssues) {
            // 有问题 → 展示列表
            let html = '<div style="margin-bottom:16px;">';
            html += '<h3 style="margin-bottom:8px;">&#9888; 共 ' + issueIndices.length + ' 条需要关注</h3>';
            html += '<p style="font-size:13px;color:var(--text-secondary);">以下用例未通过，修复后可在此页面进行回归测试</p>';
            html += '</div>';
            html += '<div class="issue-list">';

            for (const item of issueIndices) {
                try {
                    const tcResp = await fetch('/api/testcase/' + item.index);
                    const tc = await tcResp.json();
                    html += '<div class="issue-item" onclick="jumpToExecute(' + item.index + ')">';
                    html += '<div class="issue-header">';
                    html += '<span class="issue-id">' + escapeHtml(tc.id || tc.col_0 || '#' + (item.index+1)) + '</span>';
                    html += '<span class="result-badge result-badge-' + resultCss(item.result) + '">' + escapeHtml(item.result) + '</span>';
                    html += '</div>';
                    html += '<div class="issue-title">' + escapeHtml(tc.name || tc.title || tc.col_2 || '(无标题)') + '</div>';
                    html += '<div class="issue-meta">';
                    if (tc.module) html += '<span>&#128193; ' + escapeHtml(tc.module) + '</span>';
                    if (tc.priority) html += '<span>&#9889; ' + escapeHtml(tc.priority) + '</span>';
                    if (tc._saved_bug_id) html += '<span>&#128030; ' + escapeHtml(tc._saved_bug_id) + '</span>';
                    if (tc._saved_issue_time) html += '<span>&#128339; ' + escapeHtml(tc._saved_issue_time) + '</span>';
                    html += '<span>#' + (item.index + 1) + '</span>';
                    html += '</div>';
                    html += '</div>';
                } catch (e) {
                    // skip individual errors
                }
            }
            html += '</div>';
            container.innerHTML = html;
        } else {
            // 还未执行完毕，且暂无失败阻塞 → 暂无问题
            container.innerHTML = '<div class="state-message">' +
                '<div class="icon">&#9989;</div>' +
                '<h2>暂无问题</h2>' +
                '<p>已执行 ' + executed + ' / ' + total + ' 条，目前没有失败或阻塞的用例。</p>' +
                '<p style="font-size:13px;color:var(--text-secondary);margin-top:4px;">完成全部测试后将自动判断结果</p>' +
                '</div>';
        }
    } catch (err) {
        container.innerHTML = '<div class="state-message"><div class="icon">&#10060;</div><h2>加载失败</h2></div>';
    }
}

function showCelebrate(container) {
    const colors = ['#22c55e','#4f46e5','#f59e0b','#ef4444','#3b82f6','#8b5cf6','#ec4899','#14b8a6'];
    let confettiHtml = '';
    for (let i = 0; i < 60; i++) {
        const color = colors[Math.floor(Math.random() * colors.length)];
        const dx = (Math.random() - 0.5) * 400;
        const dy = -(Math.random() * 300 + 50);
        const rot = (Math.random() - 0.5) * 720;
        const delay = Math.random() * 1.5;
        const size = 6 + Math.random() * 10;
        confettiHtml += '<div class="confetti" style="background:' + color +
            ';width:' + size + 'px;height:' + size + 'px' +
            ';left:50%;top:50%' +
            ';--dx:' + dx + 'px' +
            ';--dy:' + dy + 'px' +
            ';--rot:' + rot + 'deg' +
            ';animation-delay:' + delay + 's' +
            ';"></div>';
    }

    container.innerHTML = '<div class="celebrate-wrap">' +
        confettiHtml +
        '<div style="font-size:64px;z-index:1;">&#127881;</div>' +
        '<div class="celebrate-text">恭喜！所有测试全部通过！</div>' +
        '<div class="celebrate-sub">&#127942; 所有用例执行完毕且无失败、无阻塞，测试任务圆满完成</div>' +
        '</div>';
}

// ============================================================
// 启动
// ============================================================
init();
