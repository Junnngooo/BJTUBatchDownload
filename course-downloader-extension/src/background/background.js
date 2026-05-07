const COURSE_BASE = 'http://123.121.147.7:88/ve';
const MIS_HOME    = 'https://mis.bjtu.edu.cn/home/';
const MIS_PING    = 'https://mis.bjtu.edu.cn/';
const COURSEWARE_PAGE = '10450';

let pageTabId = null;
let apiSessionId = ''; // 每次扫描后缓存，用于下载

// ── 点击图标：打开/聚焦完整页面 ──
chrome.action.onClicked.addListener(async () => {
  if (pageTabId !== null) {
    try {
      const tab = await chrome.tabs.get(pageTabId);
      await chrome.tabs.update(pageTabId, { active: true });
      await chrome.windows.update(tab.windowId, { focused: true });
      return;
    } catch { pageTabId = null; }
  }
  const tab = await chrome.tabs.create({ url: chrome.runtime.getURL('src/page/page.html') });
  pageTabId = tab.id;
});

chrome.tabs.onRemoved.addListener((tabId) => {
  if (tabId === pageTabId) pageTabId = null;
});

// ── Keep-alive alarm ──
chrome.alarms.get('misKeepAlive', (alarm) => {
  if (!alarm) chrome.alarms.create('misKeepAlive', { periodInMinutes: 5 });
});

chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name !== 'misKeepAlive') return;
  const { keepAliveEnabled } = await chrome.storage.local.get('keepAliveEnabled');
  if (!keepAliveEnabled) return;
  await pingMisNow();
});

// ── 消息路由 ──
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  // service worker 重启后 pageTabId 丢失；每次收到页面消息时顺带恢复
  if (sender.tab && sender.url?.startsWith(chrome.runtime.getURL('src/page/'))) {
    pageTabId = sender.tab.id;
  }
  switch (msg.action) {
    case 'checkConnectivity':
      checkConnectivity().then(sendResponse);
      return true;
    case 'getPageData':
      getPageDataFromTab().then(sendResponse);
      return true;
    case 'scanCourseFiles':
      scanCourseFiles(msg.course).then(sendResponse);
      return true;
    case 'startDownload':
      startDownload(msg.files, msg.sessionId, msg.rootFolder);
      sendResponse({ ok: true });
      break;
    case 'stopDownload':
      stopFlag = true;
      sendResponse({ ok: true });
      break;
    case 'setKeepAlive':
      chrome.storage.local.set({ keepAliveEnabled: msg.enabled });
      if (msg.enabled) pingMisNow();
      sendResponse({ ok: true });
      break;
    case 'pingNow':
      pingMisNow().then(sendResponse);
      return true;
  }
});

// ── 联通性检测 ──
async function checkConnectivity() {
  const result = { mis: false, course: false, misReason: '', courseReason: '' };
  const to = () => AbortSignal.timeout ? AbortSignal.timeout(6000) : undefined;

  try {
    const r = await fetch(MIS_HOME, { credentials: 'include', redirect: 'follow', cache: 'no-store', signal: to() });
    result.mis = r.status < 500;
  } catch { result.misReason = '无法连接 MIS'; }

  try {
    const r = await fetch(`${COURSE_BASE}/`, { credentials: 'include', redirect: 'follow', cache: 'no-store', signal: to() });
    result.course = r.status < 500;
    if (!result.course) result.courseReason = `HTTP ${r.status}`;
  } catch { result.courseReason = '无法连接，请确认在校园网/VPN环境下'; }

  return result;
}

// ── Keep-alive ping ──
async function pingMisNow() {
  try {
    await fetch(MIS_PING, { credentials: 'include', cache: 'no-store' });
    const ts = new Date().toLocaleTimeString('zh-CN');
    await chrome.storage.local.set({ lastKeepAlive: ts, keepAliveStatus: 'ok' });
    broadcastToPage({ action: 'keepAliveStatus', status: 'ok', time: ts });
    return { ok: true, time: ts };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ── 获取课程列表 ──
async function getPageDataFromTab() {
  const INDEX_URL = `${COURSE_BASE}/back/coursePlatform/coursePlatform.shtml?method=toCoursePlatformIndex`;
  let bgTab;

  try {
    bgTab = await chrome.tabs.create({ url: INDEX_URL, active: false });
    await waitTabComplete(bgTab.id, 15000);

    const [{ result }] = await chrome.scripting.executeScript({
      target: { tabId: bgTab.id },
      func: async () => {
        const BASE = 'http://123.121.147.7:88/ve';

        // 从页面服务端渲染的 JS 中提取 API sessionId（32位大写十六进制）
        let sid = '';
        for (const s of document.scripts) {
          const m = (s.textContent || '').match(/setRequestHeader\s*\(\s*["']sessionId["']\s*,\s*["']([A-F0-9]{32})["']\s*\)/i);
          if (m) { sid = m[1]; break; }
        }
        if (!sid) sid = document.getElementById('sessionId')?.value || '';

        // 等待页面自身的 getCourseList AJAX 完成（最多 10 秒）
        const items = await new Promise(resolve => {
          const deadline = Date.now() + 10000;
          const tick = () => {
            const els = document.querySelectorAll('.courseItem');
            if (els.length || Date.now() > deadline) resolve(els);
            else setTimeout(tick, 300);
          };
          tick();
        });

        if (items.length) {
          const courses = [];
          items.forEach(el => {
            const oc = el.getAttribute('onclick') || '';
            const m = oc.match(/goPage\(\s*'([^']*)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'\s*\)/);
            if (!m) return;
            const nameEl = el.querySelector('.course-text[title]') || el.querySelector('.course-text');
            const name = (nameEl?.getAttribute('title') || nameEl?.textContent || '').trim();
            if (name) courses.push({ cId: m[1], courseNum: m[2], xkhId: m[3], xqCode: m[4], name });
          });
          if (courses.length) return { sessionId: sid, courses, tabFound: true };
        }

        // DOM 无数据时回退：用提取到的 sid 直接调 API
        try {
          const xqRes = await fetch(`${BASE}/back/rp/common/teachCalendar.shtml?method=queryCurrentXq`);
          const xqData = await xqRes.json().catch(() => ({}));
          let xqCode = '';
          if (xqData.STATUS === '0' && Array.isArray(xqData.result)) {
            const cur = xqData.result.find(r => r.currentFlag == 2) || xqData.result[0];
            xqCode = cur?.xqCode || '';
          }
          if (xqCode) {
            const listRes = await fetch(
              `${BASE}/back/coursePlatform/course.shtml?method=getCourseList&pagesize=100&page=1&xqCode=${xqCode}`,
              { headers: sid ? { sessionId: sid } : {} }
            );
            const data = await listRes.json().catch(() => ({}));
            if (data.STATUS === '0' && Array.isArray(data.courseList) && data.courseList.length) {
              return {
                sessionId: sid,
                courses: data.courseList.map(item => ({
                  cId: String(item.id || ''),
                  courseNum: item.course_num || '',
                  xkhId: item.fz_id || '',
                  xqCode: item.xq_code || xqCode,
                  name: item.name || ''
                })).filter(c => c.name),
                tabFound: true
              };
            }
          }
        } catch {}

        return { sessionId: sid, courses: [], tabFound: true };
      }
    });

    if (result?.sessionId) apiSessionId = result.sessionId;
    return result || { sessionId: '', courses: [], tabFound: true };
  } catch (e) {
    return { sessionId: '', courses: [], tabFound: true, error: e.message };
  } finally {
    if (bgTab) chrome.tabs.remove(bgTab.id).catch(() => {});
  }
}

// ── 扫描课程电子课件 ──
async function scanCourseFiles(course) {
  const url =
    `${COURSE_BASE}/back/coursePlatform/coursePlatform.shtml` +
    `?method=toCoursePlatform&courseToPage=${COURSEWARE_PAGE}` +
    `&courseId=${encodeURIComponent(course.courseNum)}` +
    `&dataSource=1` +
    `&cId=${encodeURIComponent(course.cId)}` +
    `&xkhId=${encodeURIComponent(course.xkhId)}` +
    `&xqCode=${encodeURIComponent(course.xqCode)}`;

  let tab;
  try {
    tab = await chrome.tabs.create({ url, active: false });
    await waitTabComplete(tab.id, 20000);

    const [{ result }] = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      args: [{
        courseNum: course.courseNum,
        cId: course.cId,
        xkhId: course.xkhId,
        xqCode: course.xqCode
      }],
      func: async (c) => {
        const BASE = 'http://123.121.147.7:88/ve';
        const sid = document.getElementById('sessionId')?.value || '';

        async function fetchList(upId) {
          const qs = new URLSearchParams({
            courseId: c.courseNum, cId: c.cId,
            xkhId: c.xkhId, xqCode: c.xqCode,
            docType: '1', up_id: String(upId), searchName: ''
          });
          const res = await fetch(
            `${BASE}/back/coursePlatform/courseResource.shtml?method=stuQueryUploadResourceForCourseList&${qs}`,
            { headers: sid ? { sessionId: sid } : {} }
          );
          return res.json().catch(() => ({}));
        }

        const files = [];
        const root = await fetchList(0);

        function toFile(item, folderPath) {
          const t = item.RP_PRIX;
          // Try to extract extension from rpName itself (e.g. "第一章.pptx")
          const nameExt = (item.rpName || '').match(/(\.[a-zA-Z0-9]{2,6})$/)?.[1]?.toLowerCase() || '';
          // RP_PRIX may also come as "PPT", "PDF", etc. or be absent/"undefined"
          const rpType = (t && t !== 'undefined' && t !== 'null') ? t.toLowerCase() : '';
          return {
            rpId: item.rpId,
            name: item.rpName,
            fileType: rpType || nameExt.replace(/^\./, ''),
            canDownload: item.stu_download == '2',
            folderPath
          };
        }

        (root.resList || []).forEach(item => files.push(toFile(item, '')));

        for (const bag of (root.bagList || [])) {
          const sub = await fetchList(bag.id);
          (sub.resList || []).forEach(item => files.push(toFile(item, bag.bag_name)));
        }

        return { files, sid };
      }
    });

    if (result?.sid) apiSessionId = result.sid;
    return { success: true, files: result?.files || [], course };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    if (tab) chrome.tabs.remove(tab.id).catch(() => {});
  }
}

// 等待标签页加载完成
async function waitTabComplete(tabId, timeout = 20000) {
  try {
    const tab = await chrome.tabs.get(tabId);
    if (tab.status === 'complete') return;
  } catch { return; }

  return new Promise((resolve) => {
    const timer = setTimeout(() => {
      chrome.tabs.onUpdated.removeListener(listener);
      resolve();
    }, timeout);

    const listener = (id, info) => {
      if (id === tabId && info.status === 'complete') {
        clearTimeout(timer);
        chrome.tabs.onUpdated.removeListener(listener);
        setTimeout(resolve, 800);
      }
    };
    chrome.tabs.onUpdated.addListener(listener);
  });
}

// ── 下载批次复用的标签页 ──
let dlTabId = null; // 整批复用，避免每个文件各开一个后台 tab

// 确保有可用的 123.121.147.7 tab 供 executeScript 注入
// 返回新开的 tab（调用方负责关闭），若复用已有 tab 则返回 null
async function ensureDlTab() {
  const tabs = await chrome.tabs.query({});
  const exist = tabs.find(t =>
    t.url?.includes('123.121.147.7') &&
    t.id !== pageTabId
  );
  if (exist) { dlTabId = exist.id; return null; }

  const bg = await chrome.tabs.create({
    url: `${COURSE_BASE}/back/coursePlatform/coursePlatform.shtml?method=toCoursePlatformIndex`,
    active: false
  });
  await waitTabComplete(bg.id, 15000);
  dlTabId = bg.id;
  return bg;
}

// ── 获取文件真实下载 URL（从注入页上下文发请求，保证同源 + Cookie）──
async function resolveRpUrl(rpId) {
  if (!dlTabId) return { url: '', ext: '' };
  try { await chrome.tabs.get(dlTabId); } catch { return { url: '', ext: '' }; }

  try {
    const [{ result }] = await chrome.scripting.executeScript({
      target: { tabId: dlTabId },
      args: [rpId, apiSessionId],
      func: async (rpId, sid) => {
        const BASE = 'http://123.121.147.7:88/ve';
        const headers = sid ? { sessionId: sid } : {};

        // 1. 获取真实下载 URL
        let url = '';
        try {
          const res = await fetch(
            `${BASE}/back/resourceSpace.shtml?method=rpinfoDownloadUrl&rpId=${rpId}`,
            { method: 'POST', headers, signal: AbortSignal.timeout(10000) }
          );
          const d = await res.json().catch(() => ({}));
          url = d.rpUrl || '';
          if (!url) return { url: '', ext: '' };
          if (!url.startsWith('http')) url = BASE + (url.startsWith('/') ? '' : '/') + url;
        } catch { return { url: '', ext: '' }; }

        // 2. HEAD 请求：读 Content-Disposition / Content-Type / 重定向目标路径
        const SCRIPT_EXTS = new Set(['.shtml','.html','.htm','.php','.asp','.aspx','.jsp','.do','.action','.cgi','.pl']);
        const CT_MAP = {
          'application/pdf': '.pdf',
          'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx',
          'application/vnd.ms-powerpoint': '.ppt',
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx',
          'application/msword': '.doc',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
          'application/vnd.ms-excel': '.xls',
          'application/zip': '.zip',
          'application/x-zip-compressed': '.zip',
          'video/mp4': '.mp4',
          'video/x-msvideo': '.avi',
          'image/jpeg': '.jpg',
          'image/png': '.png',
          'text/plain': '.txt',
        };

        let ext = '';
        let headContentType = '';
        try {
          const head = await fetch(url, { method: 'HEAD', signal: AbortSignal.timeout(6000) });
          headContentType = (head.headers.get('content-type') || '').split(';')[0].trim();

          // Priority 1: Content-Disposition filename
          const cd = head.headers.get('content-disposition') || '';
          const mcd = cd.match(/filename\*=UTF-8''([^;\r\n]+)/i)
            || cd.match(/filename\*=GBK''([^;\r\n]+)/i)
            || cd.match(/filename="([^"]+)"/i)
            || cd.match(/filename=([^;\r\n]+)/i);
          if (mcd) {
            try {
              const fn = decodeURIComponent(mcd[1].trim().replace(/^["']|["']$/g, ''));
              const me = fn.match(/(\.[a-zA-Z0-9]{2,6})$/);
              if (me) ext = me[1].toLowerCase();
            } catch {
              const fn = mcd[1].trim().replace(/^["']|["']$/g, '');
              const me = fn.match(/(\.[a-zA-Z0-9]{2,6})$/);
              if (me) ext = me[1].toLowerCase();
            }
          }

          // Priority 2: Content-Type mapping
          if (!ext && CT_MAP[headContentType]) ext = CT_MAP[headContentType];

          // Priority 3: redirect destination URL path
          if (!ext) {
            try {
              const finalPath = new URL(head.url).pathname;
              const m = finalPath.match(/(\.[a-zA-Z0-9]{2,6})$/);
              if (m && !SCRIPT_EXTS.has(m[1].toLowerCase())) ext = m[1].toLowerCase();
            } catch {}
          }
        } catch {}

        // Priority 4: magic byte detection — only when all else failed
        // Read 2048 bytes: enough to cover ZIP local file headers (which store directory
        // names as plain text), letting us distinguish docx/pptx/xlsx without decompression.
        if (!ext) {
          try {
            const peek = await fetch(url, {
              method: 'GET',
              headers: { Range: 'bytes=0-2047' },
              signal: AbortSignal.timeout(8000)
            });
            const buf = await peek.arrayBuffer();
            const b = new Uint8Array(buf);
            const hex = Array.from(b.slice(0, 8)).map(x => x.toString(16).padStart(2,'0')).join('');

            if      (hex.startsWith('25504446')) ext = '.pdf';       // %PDF
            else if (hex.startsWith('ffd8ff'))   ext = '.jpg';       // JPEG
            else if (hex.startsWith('89504e47')) ext = '.png';       // PNG
            else if (hex.startsWith('52617221')) ext = '.rar';       // Rar!
            else if (hex.startsWith('377abcaf')) ext = '.7z';        // 7z
            else if (hex.startsWith('1f8b'))     ext = '.gz';        // gzip
            else if (hex.startsWith('d0cf11e0')) {
              // OLE2: read root directory CLSID at sector 0 offset 80 = file offset 592
              // CLSID first-4 bytes (little-endian DWORD): Word=06090200, PPT=108d8164, Excel=20080200
              let oleExt = '';
              if (b.length >= 596) {
                const c4 = Array.from(b.slice(592, 596)).map(x => x.toString(16).padStart(2,'0')).join('');
                if      (c4 === '06090200') oleExt = '.doc';
                else if (c4 === '108d8164') oleExt = '.ppt';
                else if (c4 === '20080200') oleExt = '.xls';
              }
              if (!oleExt) {
                // Fallback: scan directory sector (offset 512+) for UTF-16LE stream names
                const u16 = new TextDecoder('utf-16le', { fatal: false }).decode(b.slice(512));
                if      (u16.includes('WordDocument')) oleExt = '.doc';
                else if (u16.includes('PowerPoint'))   oleExt = '.ppt';
                else if (u16.includes('Workbook') || u16.includes('Book')) oleExt = '.xls';
                else                                   oleExt = '.ppt';
              }
              ext = oleExt;
            }
            else if (hex.startsWith('504b0304')) {
              // ZIP-based Office: local file headers store directory names as plain text.
              // Scan for the first characteristic subdirectory to tell formats apart.
              const text = new TextDecoder('utf-8', { fatal: false }).decode(b);
              if      (text.includes('word/')) ext = '.docx';
              else if (text.includes('ppt/'))  ext = '.pptx';
              else if (text.includes('xl/'))   ext = '.xlsx';
              else                             ext = '.zip';
            }
          } catch {}
        }

        return { url, ext };
      }
    });
    return result || { url: '', ext: '' };
  } catch { return { url: '', ext: '' }; }
}

// ── 下载队列 ──
let isDownloading = false;
let stopFlag = false;
let completedCount = 0;
let totalCount = 0;

async function startDownload(files, sessionId, rootFolder) {
  if (isDownloading) return;
  isDownloading = true;
  stopFlag = false;
  completedCount = 0;
  totalCount = files.length;

  broadcastToPage({ action: 'downloadProgress', type: 'start', total: totalCount });

  // 整批共用一个 tab，避免每文件反复开/关
  let openedBgTab = null;
  try {
    openedBgTab = await ensureDlTab();
  } catch (e) {
    broadcastToPage({ action: 'downloadProgress', type: 'progress', file: '初始化', status: 'error', completed: 0, total: totalCount, msg: '无法打开课程平台标签页: ' + e.message });
  }

  for (const file of files) {
    if (stopFlag) break;
    await downloadSingleFile(file, rootFolder);
  }

  if (openedBgTab) chrome.tabs.remove(openedBgTab.id).catch(() => {});
  dlTabId = null;
  isDownloading = false;
  broadcastToPage({ action: 'downloadProgress', type: 'done', completed: completedCount, total: totalCount });
}

async function downloadSingleFile(file, rootFolder) {
  broadcastToPage({ action: 'downloadProgress', type: 'progress', file: file.name, status: 'downloading', completed: completedCount, total: totalCount });

  const fail = (msg) => {
    broadcastToPage({ action: 'downloadProgress', type: 'progress', file: file.name, status: 'error', completed: ++completedCount, total: totalCount, msg });
  };

  try {
    const { url, ext: serverExt } = await resolveRpUrl(file.rpId);
    if (!url) { fail('未获取到下载链接'); return; }

    // 扩展名优先级：HEAD(CD/CT) > RP_PRIX > URL路径(排除服务端脚本后缀)
    // URL路径排在最后，因为下载代理 URL 形如 download.shtml?...，路径不是真实文件名
    const SERVER_SCRIPT_EXTS = new Set(['.shtml', '.html', '.htm', '.php', '.asp', '.aspx', '.jsp', '.do', '.action']);
    let ext = serverExt || '';
    if (!ext && file.fileType) {
      ext = '.' + file.fileType.toLowerCase().replace(/^\./, '');
    }
    if (!ext) {
      try {
        const candidate = new URL(url).pathname.match(/(\.[a-zA-Z0-9]{2,6})$/)?.[1]?.toLowerCase();
        if (candidate && !SERVER_SCRIPT_EXTS.has(candidate)) ext = candidate;
      } catch {}
    }

    let dlName = (file.fileName || file.name || 'download').trim();
    if (ext && !/\.[a-zA-Z0-9]{2,6}$/.test(dlName)) dlName += ext;

    const sanitize = s => s.replace(/[\\/:*?"<>|]/g, '_').trim();
    const parts = [rootFolder, file.courseName, file.folderPath, dlName].filter(Boolean);
    const filename = parts.map(sanitize).join('/');

    await new Promise((resolve) => {
      let slowTimer, verySlowTimer;

      const done = (status, msg) => {
        clearTimeout(slowTimer);
        clearTimeout(verySlowTimer);
        if (status === 'success') {
          chrome.storage.local.get({ downloadedRpIds: [] }, ({ downloadedRpIds }) => {
            const s = new Set(downloadedRpIds);
            s.add(file.rpId);
            chrome.storage.local.set({ downloadedRpIds: [...s] });
          });
        }
        broadcastToPage({ action: 'downloadProgress', type: 'progress', file: file.name, rpId: file.rpId, status, completed: ++completedCount, total: totalCount, msg });
        resolve();
      };

      chrome.downloads.download({ url, filename, conflictAction: 'uniquify', saveAs: false }, (dlId) => {
        if (chrome.runtime.lastError || !dlId) {
          done('error', chrome.runtime.lastError?.message || '下载启动失败');
          return;
        }

        slowTimer = setTimeout(() => {
          broadcastToPage({ action: 'downloadProgress', type: 'hint', msg: `${file.name}：下载较慢，可能是学校网络较慢或文件较大` });
        }, 45 * 1000);

        verySlowTimer = setTimeout(() => {
          broadcastToPage({ action: 'downloadProgress', type: 'hint', msg: `${file.name}：下载时间过长，请确认能正常登录课程平台` });
        }, 3 * 60 * 1000);

        // 安全兜底：若 onChanged 始终未触发（SW 生命周期问题），30分钟后强制推进
        const safeTimer = setTimeout(() => {
          chrome.downloads.onChanged.removeListener(onChange);
          done('success');
        }, 30 * 60 * 1000);

        const onChange = (delta) => {
          if (delta.id !== dlId) return;
          const s = delta.state?.current;
          if (s === 'complete' || s === 'interrupted') {
            clearTimeout(safeTimer);
            chrome.downloads.onChanged.removeListener(onChange);
            done(s === 'complete' ? 'success' : 'error', s === 'interrupted' ? '下载中断' : '');
          }
        };
        chrome.downloads.onChanged.addListener(onChange);
      });
    });
  } catch (e) {
    fail(e.message || '未知错误');
  }
}

// ── 广播到插件页面 ──
// 不依赖 pageTabId——SW 重启后变量丢失；先试缓存值，失败则动态查找
async function broadcastToPage(data) {
  const tryTab = async (id) => {
    try { await chrome.tabs.sendMessage(id, data); return true; } catch { return false; }
  };
  if (pageTabId !== null && await tryTab(pageTabId)) return;
  const tabs = await chrome.tabs.query({ url: chrome.runtime.getURL('src/page/page.html') }).catch(() => []);
  if (tabs.length) { pageTabId = tabs[0].id; await tryTab(pageTabId); }
}
