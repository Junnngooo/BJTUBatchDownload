// ── 全局状态 ──
let sessionId = '';
let courses = [];
let courseFiles = {};
let selectedFiles = new Set();
let downloadedSet = new Set(); // 持久化已下载 rpId 集合
let isDownloading = false;
let completedDl = 0, totalDl = 0;
let isScanningAll = false;
let scanStopRequested = false;
let connState = { mis: null, course: null }; // null=checking, true=ok, false=err

// ── 启动 ──
document.addEventListener('DOMContentLoaded', async () => {
  bindEvents();
  await loadStoredSettings();
  await init();
});

async function init() {
  showMainContent();
  window._courseLoading = true;
  renderCourses([]);

  checkAndUpdateConn(); // independent — always resolves and updates UI

  getPageInfo()
    .then(info => {
      window._courseLoading = false;
      if (info?.sessionId) sessionId = info.sessionId;
      renderCourses(info?.courses || []);
    })
    .catch(() => { window._courseLoading = false; renderCourses([]); });
}

// Direct connectivity check from the page itself (extension host_permissions covers both domains).
// Avoids service-worker message-passing entirely — no wakeup race, no silent swallow.
async function checkAndUpdateConn() {
  setConnStatus('mis', 'checking');
  setConnStatus('course', 'checking');

  async function probe(url) {
    try {
      const r = await fetch(url, {
        credentials: 'include', redirect: 'follow',
        cache: 'no-store', signal: AbortSignal.timeout(8000)
      });
      return r.status < 500;
    } catch { return false; }
  }

  const [mis, course] = await Promise.all([
    probe('https://mis.bjtu.edu.cn/'),
    probe('http://123.121.147.7:88/ve/')
  ]);
  updateConnUI({ mis, course });
}

function updateConnUI(conn) {
  connState.mis    = !!conn.mis;
  connState.course = !!conn.course;

  setConnStatus('mis', conn.mis ? 'ok' : 'err');
  document.getElementById('misLabel').textContent = conn.mis ? 'MIS ✓' : 'MIS 不可达';

  setConnStatus('course', conn.course ? 'ok' : 'err');
  document.getElementById('courseLabel').textContent = conn.course ? '课程平台 ✓' : '课程平台 不可达';
}

function setConnStatus(which, state) {
  const dot = document.getElementById(which + 'Dot');
  dot.className = 'conn-dot ' + state;
}

// ── 从课程平台标签页提取信息 ──
async function getPageInfo() {
  return await chrome.runtime.sendMessage({ action: 'getPageData' });
}

// ── 登录提示 / 主体切换 ──
function showLoginWarning(title, detail) {
  document.getElementById('warningTitle').textContent = title;
  document.getElementById('warningDetail').textContent = detail;
  document.getElementById('loginWarning').classList.remove('hidden');
  document.getElementById('mainContent').classList.add('hidden');
}

function showMainContent() {
  document.getElementById('loginWarning').classList.add('hidden');
  document.getElementById('mainContent').classList.remove('hidden');
}

// ── 课程列表渲染 ──
function renderCourses(list) {
  courses = list || [];
  document.getElementById('courseCount').textContent = courses.length;
  const el = document.getElementById('courseList');

  if (!courses.length) {
    el.innerHTML = window._courseLoading
      ? '<div class="empty-tip">正在自动加载课程列表，请稍候...</div>'
      : `<div class="empty-tip">
           未检测到课程，请确认已登录 MIS 系统且处于校园网 / VPN 环境，然后点击 ↻ 重新检测。<br><br>
           <a href="http://123.121.147.7:88/ve/back/coursePlatform/coursePlatform.shtml?method=toCoursePlatformIndex"
              target="_blank" style="color:#4a7fd4;font-size:12px">手动打开课程中心</a>
         </div>`;
    return;
  }

  el.innerHTML = courses.map((c, i) => `
    <div class="course-item" data-num="${esc(c.courseNum)}">
      <input type="checkbox" class="course-chk" data-num="${esc(c.courseNum)}">
      <div class="course-item-body">
        <div class="course-name" title="${esc(c.name)}">${esc(c.name)}</div>
        <div class="course-meta">
          <code style="font-size:10px">${esc(c.courseNum)}</code>
          <span id="cs-${esc(c.courseNum)}"></span>
        </div>
      </div>
      <button class="btn btn-xs course-scan-btn" data-num="${esc(c.courseNum)}">扫描</button>
    </div>
  `).join('');

  el.querySelectorAll('.course-item').forEach(item => {
    item.querySelector('.course-name').addEventListener('click', () => {
      highlightCourse(item.dataset.num);
      scrollToGroup(item.dataset.num);
    });
  });

  el.querySelectorAll('.course-scan-btn').forEach(btn => {
    btn.addEventListener('click', () => scanOneCourse(btn.dataset.num));
  });
}

function highlightCourse(num) {
  document.querySelectorAll('.course-item').forEach(el => el.classList.remove('active'));
  document.querySelector(`.course-item[data-num="${num}"]`)?.classList.add('active');
}

function scrollToGroup(num) {
  document.getElementById(`group-${num}`)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ── 扫描文件 ──
async function scanOneCourse(courseNum) {
  const course = courses.find(c => c.courseNum === courseNum);
  if (!course) return;
  setCourseStatus(courseNum, 'scanning', '扫描中...');
  switchTab('files');

  const res = await chrome.runtime.sendMessage({ action: 'scanCourseFiles', sessionId, course });

  if (res?.success) {
    courseFiles[courseNum] = res.files;
    const canDl = res.files.filter(f => f.canDownload).length;
    setCourseStatus(courseNum, 'done', `${canDl}/${res.files.length} 可下载`);
    renderFileGroup(course, res.files);
    checkDownloadedByHistory(courseNum).catch(() => {});
    updateFileStats();
  } else {
    setCourseStatus(courseNum, 'error', '失败');
    addLog(`扫描失败: ${course.name} — ${res?.error || '未知'}`, 'error');
  }
}

async function scanAllCourses() {
  const btn = document.getElementById('btnScanAll');

  if (isScanningAll) {
    scanStopRequested = true;
    btn.textContent = '停止中...';
    btn.disabled = true;
    return;
  }

  isScanningAll = true;
  scanStopRequested = false;
  btn.textContent = '停止扫描';
  btn.classList.remove('btn-primary');
  btn.classList.add('btn-danger');

  switchTab('files');
  document.getElementById('filePanel').innerHTML = '';

  for (const c of courses) {
    if (scanStopRequested) break;
    await scanOneCourse(c.courseNum);
  }

  isScanningAll = false;
  scanStopRequested = false;
  btn.disabled = false;
  btn.textContent = '扫描全部课件';
  btn.classList.remove('btn-danger');
  btn.classList.add('btn-primary');
  updateFileStats();
}

function setCourseStatus(num, type, text) {
  const el = document.getElementById(`cs-${num}`);
  if (!el) return;
  el.className = `course-status status-${type}`;
  el.textContent = text;
}

// ── 文件渲染 ──
function renderFileGroup(course, files) {
  const panel = document.getElementById('filePanel');
  document.getElementById(`group-${course.courseNum}`)?.remove();
  panel.querySelector('.empty-tip')?.remove();

  const canDl = files.filter(f => f.canDownload).length;
  const locked = files.length - canDl;
  const num = course.courseNum;

  const g = document.createElement('div');
  g.className = 'course-group';
  g.id = `group-${num}`;

  const header = document.createElement('div');
  header.className = 'course-group-header';
  header.innerHTML = `
    <div>
      <div class="course-group-title">${esc(course.name)}</div>
      <div class="course-group-meta">${canDl} 可下载 · ${locked} 受限 · 共 ${files.length} 个</div>
    </div>
    <div class="course-group-right">
      <span class="course-group-toggle" id="tog-${esc(num)}">▾</span>
    </div>`;
  header.addEventListener('click', () => toggleGroup(num));

  const body = document.createElement('div');
  body.id = `gbody-${esc(num)}`;

  if (files.length === 0) {
    body.innerHTML = '<div class="empty-tip" style="padding:20px">该课程暂无课件</div>';
  } else {
    const table = document.createElement('table');
    table.className = 'file-table';
    table.innerHTML = `
      <thead><tr>
        <th style="width:30px"><input type="checkbox" id="chk-all-${esc(num)}"></th>
        <th>文件名</th><th style="width:82px">状态</th>
      </tr></thead>
      <tbody>
        ${files.map(f => `
          <tr>
            <td>${f.canDownload
              ? `<input type="checkbox" class="file-chk"
                  data-rpid="${esc(f.rpId)}" data-course="${esc(num)}"
                  data-name="${esc(f.name)}" data-folder="${esc(f.folderPath||'')}"
                  data-filetype="${esc(f.fileType||'')}">`
              : '<input type="checkbox" disabled>'
            }</td>
            <td class="file-name">
              <span class="file-name-text" title="${esc(f.name)}">${esc(f.name)}</span>
              ${f.folderPath ? `<span class="file-path">${esc(f.folderPath)}</span>` : ''}
            </td>
            <td id="st-${esc(f.rpId)}">${
              !f.canDownload ? '<span class="tag-locked">受限</span>'
              : downloadedSet.has(f.rpId) ? '<span class="tag-downloaded">已下载</span>'
              : '<span class="tag-ok">可下载</span>'
            }</td>
          </tr>`).join('')}
      </tbody>`;

    const chkAll = table.querySelector(`#chk-all-${esc(num)}`);
    chkAll?.addEventListener('change', () => {
      table.querySelectorAll('.file-chk').forEach(c => { c.checked = chkAll.checked; onFileCheck(c); });
    });
    table.querySelectorAll('.file-chk').forEach(c => c.addEventListener('change', () => onFileCheck(c)));

    body.appendChild(table);
  }

  g.appendChild(header);
  g.appendChild(body);
  panel.appendChild(g);
  g.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function toggleGroup(num) {
  const body = document.getElementById(`gbody-${num}`);
  const tog = document.getElementById(`tog-${num}`);
  if (!body) return;
  const hidden = body.style.display === 'none';
  body.style.display = hidden ? '' : 'none';
  if (tog) tog.textContent = hidden ? '▾' : '▸';
}

function onFileCheck(el) {
  el.checked ? selectedFiles.add(el.dataset.rpid) : selectedFiles.delete(el.dataset.rpid);
}

function updateFileStats() {
  const total = Object.values(courseFiles).reduce((s, f) => s + f.length, 0);
  const dl = Object.values(courseFiles).reduce((s, f) => s + f.filter(x => x.canDownload).length, 0);
  document.getElementById('fileStats').textContent = `共 ${total} 个文件，${dl} 个可下载`;
}

// ── 下载 ──
function buildFilesFromCourses(nums) {
  return nums.flatMap(num => {
    const c = courses.find(x => x.courseNum === num);
    return (courseFiles[num] || [])
      .filter(f => f.canDownload)
      .map(f => ({ ...f, fileName: f.name, courseName: c?.name || num }));
  });
}

function buildSelectedFilesList() {
  const files = [];
  document.querySelectorAll('.file-chk:checked').forEach(el => {
    const c = courses.find(x => x.courseNum === el.dataset.course);
    files.push({
      rpId: el.dataset.rpid,
      name: el.dataset.name,
      fileName: el.dataset.name,
      folderPath: el.dataset.folder,
      fileType: el.dataset.filetype || '',
      courseName: c?.name || el.dataset.course
    });
  });
  return files;
}

function showConnWarning() {
  document.getElementById('connAlert').classList.remove('hidden');
}
function hideConnWarning() {
  document.getElementById('connAlert').classList.add('hidden');
}

function startDownload(files) {
  if (!files.length) { alert('没有可下载的文件'); return; }
  if (isDownloading) { alert('下载进行中，请等待或先停止'); return; }
  const duplicates = files.filter(f => downloadedSet.has(f.rpId));
  if (duplicates.length > 0) {
    showDuplicateModal(files, duplicates, (finalFiles) => {
      if (finalFiles && finalFiles.length > 0) {
        uncheckAllFiles();
        doStartDownload(finalFiles);
      }
    });
  } else {
    uncheckAllFiles();
    doStartDownload(files);
  }
}

function doStartDownload(files) {
  isDownloading = true;
  totalDl = files.length; completedDl = 0;
  updateProgressBar(0, totalDl);
  switchTab('progress');
  document.getElementById('btnStopDownload').classList.remove('hidden');
  const rootFolder = document.getElementById('rootFolder').value.trim() || '课程资料';
  addLog(`准备下载 ${files.length} 个文件`, 'info');
  chrome.runtime.sendMessage({ action: 'startDownload', files, sessionId, rootFolder });
}

function uncheckAllFiles() {
  document.querySelectorAll('.file-chk:checked').forEach(c => { c.checked = false; onFileCheck(c); });
  document.querySelectorAll('[id^="chk-all-"]').forEach(c => c.checked = false);
  document.querySelectorAll('.course-chk:checked').forEach(c => c.checked = false);
  selectedFiles.clear();
}

function showDuplicateModal(allFiles, duplicates, callback) {
  const modal = document.getElementById('dupModal');
  const list  = document.getElementById('dupModalList');
  document.getElementById('dupModalSubtitle').textContent =
    `以下 ${duplicates.length} 个文件已下载过，请勾选要重新下载的文件：`;

  const dupSet = new Set(duplicates.map(f => f.rpId));
  list.innerHTML = duplicates.map(f => `
    <label class="dup-item">
      <input type="checkbox" class="dup-chk" data-rpid="${esc(f.rpId)}" checked>
      <span class="dup-name" title="${esc(f.name)}">${esc(f.name)}</span>
      <span class="dup-course">${esc(f.courseName || '')}</span>
    </label>`).join('');

  modal.classList.remove('hidden');

  const cleanup = (finalFiles) => { modal.classList.add('hidden'); callback(finalFiles); };

  document.getElementById('dupSelectAll').onclick   = () => list.querySelectorAll('.dup-chk').forEach(c => c.checked = true);
  document.getElementById('dupDeselectAll').onclick = () => list.querySelectorAll('.dup-chk').forEach(c => c.checked = false);
  document.getElementById('dupCancel').onclick  = () => cleanup(null);
  document.getElementById('dupConfirm').onclick = () => {
    const kept = new Set([...list.querySelectorAll('.dup-chk:checked')].map(c => c.dataset.rpid));
    cleanup(allFiles.filter(f => !dupSet.has(f.rpId) || kept.has(f.rpId)));
  };
}

function markFileDownloaded(rpId) {
  const td = document.getElementById(`st-${rpId}`);
  if (td) td.innerHTML = '<span class="tag-downloaded">已下载</span>';
}

// 扫描完成后，与 chrome.downloads 历史交叉验证已下载状态
// 即使 storage 被清除，只要文件还在下载历史中就能正确标记
async function checkDownloadedByHistory(courseNum) {
  const files = courseFiles[courseNum];
  if (!files?.length) return;

  const rootFolder = document.getElementById('rootFolder').value.trim() || '课程资料';
  const course = courses.find(c => c.courseNum === courseNum);
  const sanitize = s => s.replace(/[\\/:*?"<>|]/g, '_').trim();

  let completed;
  try {
    completed = await chrome.downloads.search({ state: 'complete', limit: 5000 });
  } catch { return; }

  // 统一为正斜杠 + 小写，便于 includes 比对
  const dlPaths = completed.map(d => d.filename.replace(/\\/g, '/').toLowerCase());

  let anyNew = false;
  for (const file of files) {
    if (!file.canDownload || downloadedSet.has(file.rpId)) continue;
    const parts = [rootFolder, course?.name || courseNum, file.folderPath, file.name]
      .filter(Boolean).map(sanitize);
    const basePath = parts.join('/').toLowerCase();
    if (dlPaths.some(p => p.includes(basePath))) {
      downloadedSet.add(file.rpId);
      markFileDownloaded(file.rpId);
      anyNew = true;
    }
  }
  if (anyNew) chrome.storage.local.set({ downloadedRpIds: [...downloadedSet] });
}

// ── Keep-alive ──
async function loadStoredSettings() {
  const { keepAliveEnabled, lastKeepAlive, downloadedRpIds } = await chrome.storage.local.get(['keepAliveEnabled', 'lastKeepAlive', 'downloadedRpIds']);
  document.getElementById('keepAliveToggle').checked = !!keepAliveEnabled;
  if (lastKeepAlive) document.getElementById('keepAliveInfo').textContent = `上次: ${lastKeepAlive}`;
  if (Array.isArray(downloadedRpIds)) downloadedRpIds.forEach(id => downloadedSet.add(id));
}

function updateKeepAliveInfo(status, time) {
  const info = document.getElementById('keepAliveInfo');
  info.textContent = status === 'ok' ? `上次: ${time}` : '刷新失败';
  info.style.color = status === 'ok' ? '#9ca3af' : '#dc2626';
}

// ── 进度监听 ──
chrome.runtime.onMessage.addListener((msg) => {
  if (msg.action === 'downloadProgress') {
    if (msg.type === 'start') { totalDl = msg.total; completedDl = 0; }
    else if (msg.type === 'progress') {
      completedDl = msg.completed; totalDl = msg.total;
      updateProgressBar(msg.completed, msg.total);
      const icon = { success: '✓', error: '✗', downloading: '↓' }[msg.status] || '·';
      const detail = msg.msg ? ` (${msg.msg})` : '';
      addLog(`${icon} ${msg.file}${detail}`, msg.status === 'error' ? 'error' : msg.status === 'success' ? 'success' : 'info');
      if (msg.status === 'success' && msg.rpId) {
        downloadedSet.add(msg.rpId);
        markFileDownloaded(msg.rpId);
      }
    } else if (msg.type === 'done') {
      isDownloading = false;
      document.getElementById('btnStopDownload').classList.add('hidden');
      addLog(`下载完成：${msg.completed}/${msg.total}`, 'info');
    } else if (msg.type === 'hint') {
      addLog(msg.msg, 'warn');
    }
  }
  if (msg.action === 'keepAliveStatus') {
    updateKeepAliveInfo(msg.status, msg.time);
  }
});

function updateProgressBar(done, total) {
  const pct = total > 0 ? Math.round(done / total * 100) : 0;
  document.getElementById('progressBar').style.width = pct + '%';
  document.getElementById('progressText').textContent = `${done} / ${total} (${pct}%)`;
  const icon = document.getElementById('progressIcon');
  if (icon) icon.style.left = pct + '%';
}

function addLog(msg, type = 'info') {
  const log = document.getElementById('progressLog');
  const d = document.createElement('div');
  d.className = `log-line log-${type}`;
  d.innerHTML = `<span class="log-time">${now()}</span><span class="log-msg">${esc(msg)}</span>`;
  log.appendChild(d);
  log.scrollTop = log.scrollHeight;
}

// ── Tab 切换 ──
function switchTab(name) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === name));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.toggle('active', c.id === `tab${name[0].toUpperCase() + name.slice(1)}`));
}

// ── 事件绑定 ──
function bindEvents() {
  document.getElementById('btnRecheck').addEventListener('click', () => { hideConnWarning(); checkAndUpdateConn(); });
  document.getElementById('btnRetryCheck').addEventListener('click', () => { hideConnWarning(); init(); });
  document.getElementById('connAlertClose').addEventListener('click', hideConnWarning);
  document.getElementById('btnScanAll').addEventListener('click', scanAllCourses);

  document.getElementById('btnSelectAllCourses').addEventListener('click', () => {
    document.querySelectorAll('.course-chk').forEach(c => c.checked = true);
  });
  document.getElementById('btnDeselectCourses').addEventListener('click', () => {
    document.querySelectorAll('.course-chk').forEach(c => c.checked = false);
  });
  document.getElementById('btnDownloadSelected').addEventListener('click', () => {
    const nums = [...document.querySelectorAll('.course-chk:checked')].map(c => c.dataset.num);
    if (!nums.length) { alert('请勾选课程'); return; }
    const unscanned = nums.filter(n => !courseFiles[n]);
    if (unscanned.length) {
      const names = unscanned.map(n => courses.find(c => c.courseNum === n)?.name || n).join('\n');
      alert(`请先扫描以下课程:\n${names}`);
      return;
    }
    startDownload(buildFilesFromCourses(nums));
  });

  document.getElementById('btnSelectAllFiles').addEventListener('click', () => {
    document.querySelectorAll('.file-chk').forEach(c => { c.checked = true; onFileCheck(c); });
    document.querySelectorAll('[id^="chk-all-"]').forEach(c => c.checked = true);
  });
  document.getElementById('btnDeselectFiles').addEventListener('click', () => {
    document.querySelectorAll('.file-chk').forEach(c => { c.checked = false; onFileCheck(c); });
    document.querySelectorAll('[id^="chk-all-"]').forEach(c => c.checked = false);
    selectedFiles.clear();
  });
  document.getElementById('btnDownloadFiles').addEventListener('click', () => {
    startDownload(buildSelectedFilesList());
  });

  document.getElementById('btnStopDownload').addEventListener('click', () => {
    chrome.runtime.sendMessage({ action: 'stopDownload' });
    isDownloading = false;
    document.getElementById('btnStopDownload').classList.add('hidden');
    addLog('已请求停止', 'warn');
  });
  document.getElementById('btnClearLog').addEventListener('click', () => {
    document.getElementById('progressLog').innerHTML = '';
  });

  document.getElementById('keepAliveToggle').addEventListener('change', function() {
    chrome.runtime.sendMessage({ action: 'setKeepAlive', enabled: this.checked });
    addLog(this.checked ? '防过期已开启（每5分钟自动刷新 MIS 会话）' : '防过期已关闭', 'info');
  });

  document.getElementById('btnPingNow').addEventListener('click', async () => {
    const btn = document.getElementById('btnPingNow');
    btn.disabled = true; btn.textContent = '...';
    const res = await chrome.runtime.sendMessage({ action: 'pingNow' });
    btn.disabled = false; btn.textContent = '刷新';
    if (res?.ok) {
      updateKeepAliveInfo('ok', res.time);
      addLog('手动刷新 MIS 会话成功', 'success');
    } else {
      addLog(`手动刷新失败: ${res?.error || '未知'}`, 'error');
    }
  });

  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => switchTab(btn.dataset.tab));
  });
}

// ── 工具函数 ──
function esc(s) {
  return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function fileIcon() { return ''; }

function fileTypeName(f) {
  const t = (f.fileType || '').trim().toLowerCase();
  if (t) return t.toUpperCase();
  const nameParts = (f.name || '').split('.');
  if (nameParts.length > 1) return nameParts.pop().toUpperCase().substring(0, 6);
  return '—';
}
function now() {
  return new Date().toLocaleTimeString('zh-CN',{hour12:false});
}
