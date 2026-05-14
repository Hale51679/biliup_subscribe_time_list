// Cloudflare Worker — B站 API 代理 + 前端页面
// 部署: https://dash.cloudflare.com → Workers & Pages → 创建/更新 Worker
// 访问 Worker 域名即打开前端页面，?url= 参数走 API 代理

addEventListener("fetch", event => {
    event.respondWith(handleRequest(event.request))
})

// ── 内嵌前端 HTML ──────────────────────────────────────────
const HTML_PAGE = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>B站关注列表导出工具</title>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/qrcodejs@1.0.0/qrcode.min.js"></script>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, "Microsoft YaHei", sans-serif; background: #f4f5f7; color: #1f2a3a; min-height: 100vh; }
.container { max-width: 640px; margin: 0 auto; padding: 16px; }
.header { background: #00A1D6; text-align: center; padding: 20px 16px; }
.header h1 { color: #fff; font-size: 20px; }
.header p { color: #cce8f7; font-size: 13px; margin-top: 4px; }
.tabs { display: flex; background: #fff; border-bottom: 1px solid #e8ecf1; margin-top: 12px; }
.tab-btn { flex: 1; padding: 12px; text-align: center; font-size: 14px; font-weight: 600;
  cursor: pointer; color: #8e9aaf; border: none; background: none; transition: all .2s; }
.tab-btn.active { color: #00A1D6; border-bottom: 2px solid #00A1D6; }
.tab-btn:hover { color: #00A1D6; }
.tab-content { display: none; }
.tab-content.active { display: block; }
.card { background: #fff; border: 1px solid #e8ecf1; border-radius: 8px; padding: 16px; margin-top: 12px; }
.card-title { font-size: 13px; font-weight: 700; margin-bottom: 6px; }
.card-sub { font-size: 12px; color: #8e9aaf; margin-bottom: 4px; }
label { display: block; font-size: 13px; font-weight: 600; margin-bottom: 4px; margin-top: 12px; }
label:first-child { margin-top: 0; }
input, textarea { width: 100%; padding: 10px 12px; font-size: 14px; border: 1px solid #e8ecf1;
  border-radius: 6px; outline: none; transition: border .2s; font-family: "Consolas", monospace; }
input:focus, textarea:focus { border-color: #00A1D6; }
textarea { resize: vertical; min-height: 60px; }
.btn { display: inline-flex; align-items: center; justify-content: center; gap: 6px;
  padding: 10px 24px; font-size: 14px; font-weight: 600; border: none; border-radius: 6px;
  cursor: pointer; transition: all .2s; width: 100%; }
.btn-primary { background: #00A1D6; color: #fff; }
.btn-primary:hover { background: #0085b3; }
.btn-primary:disabled { background: #b0d4e6; cursor: not-allowed; }
.btn-secondary { background: #fff; color: #00A1D6; border: 1px solid #00A1D6; }
.btn-secondary:hover { background: #f0f9ff; }
.status { padding: 10px 14px; border-radius: 6px; font-size: 13px; margin-top: 8px; display: none; }
.status.info { display: block; background: #f0f9ff; color: #004085; border: 1px solid #b8daff; }
.status.success { display: block; background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
.status.error { display: block; background: #fff3f3; color: #721c24; border: 1px solid #f5c6cb; }
.status.warning { display: block; background: #fffbe6; color: #856404; border: 1px solid #ffeeba; }
.hidden { display: none !important; }
.mt-8 { margin-top: 8px; }
.mt-12 { margin-top: 12px; }
.flex { display: flex; gap: 8px; }
.log-box { background: #1a1d23; color: #ccc; border-radius: 6px; padding: 10px; font-size: 12px;
  font-family: "Consolas", monospace; height: 140px; overflow-y: auto; white-space: pre-wrap;
  word-break: break-all; margin-top: 8px; }
.progress-bar { height: 6px; background: #e8ecf1; border-radius: 3px; margin-top: 8px; overflow: hidden; }
.progress-bar-inner { height: 100%; background: #00A1D6; border-radius: 3px; width: 0%; transition: width .3s; }
.hint { font-size: 12px; color: #8e9aaf; margin-top: 4px; }
.qr-wrap { display: flex; justify-content: center; padding: 16px 0 8px; }
.qr-wrap #qrcode { padding: 12px; background: #fff; border: 1px solid #e8ecf1; border-radius: 8px; display: inline-block; }
.qr-wrap #qrcode img, .qr-wrap #qrcode canvas { display: block; }
.result-card { padding: 14px; border-radius: 8px; margin-top: 8px; }
.result-card.following { background: #e8f7fd; border: 1px solid #b8daff; }
.result-card.not-following { background: #fff3f3; border: 1px solid #f5c6cb; }
.result-card h3 { margin: 0 0 4px 0; font-size: 15px; }
.result-card p { margin: 2px 0; font-size: 13px; }
.sessdata-box { background: #f0f8ff; border: 1px solid #b8daff; border-radius: 6px;
  padding: 8px 10px; font-family: Consolas, monospace; font-size: 12px;
  word-break: break-all; margin-top: 6px; }
</style>
</head>
<body>
<div class="header">
  <h1>B站关注列表导出</h1>
  <p>批量导出你关注的 UP 主名称、UID 与关注时间</p>
</div>

<div class="container">
  <div class="tabs">
    <button class="tab-btn active" onclick="switchTab('qr', this)">扫码登录</button>
    <button class="tab-btn" onclick="switchTab('batch', this)">批量导出</button>
    <button class="tab-btn" onclick="switchTab('single', this)">单个查询</button>
  </div>

  <div id="tabQr" class="tab-content active">
    <div class="card">
      <div class="card-title">扫码登录</div>
      <div class="card-sub">用 B站 App 扫码，自动获取 SESSDATA 和 UID</div>
      <div class="qr-wrap"><div id="qrcode"></div></div>
      <div id="qrStatus" class="status"></div>
      <div id="qrActions">
        <button class="btn btn-primary" onclick="startQRLogin()" id="qrBtn">生成二维码</button>
      </div>
      <div id="qrResult" class="hidden">
        <div style="margin-top:8px;">
          <div style="font-size:12px;color:#8e9aaf;">SESSDATA</div>
          <div id="qrSessdata" class="sessdata-box"></div>
        </div>
        <div style="margin-top:6px;font-size:12px;" id="qrUid"></div>
        <div class="hint mt-8">已自动填入「批量导出」和「单个查询」</div>
      </div>
      <pre id="qrLog" class="log-box hidden"></pre>
    </div>
  </div>

  <div id="tabBatch" class="tab-content">
    <div class="card">
      <div class="card-title">批量导出关注列表</div>
      <div class="card-sub">导出结果为 Excel 文件，包含 UP 主名称、UID、关注时间</div>
      <label>SESSDATA</label>
      <input type="text" id="batchSessdata" placeholder="从扫码登录获取，或手动填写" autocomplete="off">
      <div class="hint">登录B站 - F12 - Application - Cookies - bilibili.com - 复制 SESSDATA</div>
      <label>你的 B站 UID</label>
      <input type="text" id="batchUid" placeholder="扫码自动填入，或手动填写">
      <div class="flex mt-12">
        <button class="btn btn-primary" onclick="startBatchExport()" id="batchBtn">开始导出</button>
      </div>
      <div id="batchStatus" class="status"></div>
      <div id="batchProgress" class="progress-bar hidden"><div class="progress-bar-inner" id="batchProgressInner"></div></div>
      <pre id="batchLog" class="log-box hidden"></pre>
      <div id="batchResult" class="hidden">
        <div class="status success" id="batchDoneMsg"></div>
        <button class="btn btn-secondary mt-8" onclick="downloadExcel()">下载 Excel 文件</button>
      </div>
    </div>
  </div>

  <div id="tabSingle" class="tab-content">
    <div class="card">
      <div class="card-title">查询是否关注某个 UP 主</div>
      <div class="card-sub">输入 UP 主 UID，查看关注状态和关注时间</div>
      <label>SESSDATA</label>
      <input type="text" id="singleSessdata" placeholder="从扫码登录获取，或手动填写" autocomplete="off">
      <label>UP主 UID</label>
      <input type="text" id="singleUid" placeholder="要查询的 UP 主 UID">
      <div class="flex mt-12">
        <button class="btn btn-primary" onclick="startSingleQuery()" id="singleBtn">查询</button>
      </div>
      <div id="singleStatus" class="status"></div>
      <div id="singleResult"></div>
    </div>
  </div>

  <div style="text-align:center; padding: 20px 0; color: #8e9aaf; font-size: 12px;">
    数据来源 Bilibili API
  </div>
</div>

<script>
async function api(method, biliUrl, cookie) {
  var params = new URLSearchParams()
  params.set('url', biliUrl)
  if (cookie) params.set('cookie', cookie)
  var resp = await fetch('/?' + params.toString(), { method: method })
  if (!resp.ok) {
    var t = await resp.text()
    throw new Error('请求失败 (' + resp.status + '): ' + t.slice(0, 100))
  }
  var text = await resp.text()
  try { return JSON.parse(text) } catch(e) { return text }
}
function setStatus(id, type, msg) { var el = document.getElementById(id); el.className = 'status ' + type; el.textContent = msg }
function logEl(id, msg) { var el = document.getElementById(id); el.classList.remove('hidden'); el.textContent += msg + '\\n'; el.scrollTop = el.scrollHeight }
function clearEl(id) { var el = document.getElementById(id); el.textContent = ''; el.classList.add('hidden') }
function switchTab(name, btn) {
  document.querySelectorAll('.tab-btn').forEach(function(b) { b.classList.remove('active') })
  document.querySelectorAll('.tab-content').forEach(function(c) { c.classList.remove('active') })
  document.getElementById('tab' + name.charAt(0).toUpperCase() + name.slice(1)).classList.add('active')
  btn.classList.add('active')
}
var qrPollTimer = null
async function startQRLogin() {
  var btn = document.getElementById('qrBtn'); btn.disabled = true; btn.textContent = '生成中...'
  setStatus('qrStatus', 'info', '正在请求二维码...')
  document.getElementById('qrResult').classList.add('hidden'); clearEl('qrLog'); document.getElementById('qrcode').innerHTML = ''
  try {
    var resp = await api('GET', 'https://passport.bilibili.com/x/passport-login/web/qrcode/generate?source=main-fe-header')
    if (resp.code !== 0) throw new Error(resp.message || '生成二维码失败')
    new QRCode(document.getElementById('qrcode'), { text: resp.data.url, width: 180, height: 180, correctLevel: QRCode.CorrectLevel.H })
    setStatus('qrStatus', 'info', '请打开B站App扫描二维码'); btn.textContent = '重新生成'; btn.disabled = false
    logEl('qrLog', '二维码已生成'); logEl('qrLog', '请打开B站App扫码'); startQrPoll(resp.data.qrcode_key)
  } catch(e) { setStatus('qrStatus', 'error', '失败: ' + e.message); btn.textContent = '生成二维码'; btn.disabled = false }
}
async function startQrPoll(key) {
  if (qrPollTimer) { clearInterval(qrPollTimer); qrPollTimer = null }
  var elapsed = 0
  qrPollTimer = setInterval(async function() {
    elapsed += 2
    if (elapsed > 180) { clearInterval(qrPollTimer); qrPollTimer = null; setStatus('qrStatus', 'error', '扫码超时，请重新生成'); return }
    try {
      var resp = await api('GET', 'https://passport.bilibili.com/x/passport-login/web/qrcode/poll?qrcode_key=' + key)
      var c = resp.data && resp.data.code
      if (c === 0) {
        clearInterval(qrPollTimer); qrPollTimer = null
        var s = resp._sessdata || ''
        if (!s) { setStatus('qrStatus', 'error', '获取 SESSDATA 失败，请手动填写'); return }
        logEl('qrLog', '登录成功！')
        document.getElementById('qrSessdata').textContent = s; document.getElementById('qrResult').classList.remove('hidden')
        document.getElementById('batchSessdata').value = s; document.getElementById('singleSessdata').value = s
        try {
          var n = await api('GET', 'https://api.bilibili.com/x/web-interface/nav', 'SESSDATA=' + s)
          if (n.code === 0 && n.data) { document.getElementById('batchUid').value = n.data.mid; document.getElementById('qrUid').textContent = 'UID: ' + n.data.mid }
        } catch(_) {}
        setStatus('qrStatus', 'success', '登录成功！SESSDATA 和 UID 已自动填入'); logEl('qrLog', 'SESSDATA 和 UID 已自动填入各Tab')
      } else if (c === 86090) { setStatus('qrStatus', 'info', '已扫码，请在手机上确认...') }
      else if (c === 86038) { clearInterval(qrPollTimer); qrPollTimer = null; setStatus('qrStatus', 'error', '二维码已过期，请重新生成'); logEl('qrLog', '二维码已过期') }
    } catch(e) {}
  }, 2000)
}
var exportResults = []
async function startBatchExport() {
  var s = document.getElementById('batchSessdata').value.trim(), u = document.getElementById('batchUid').value.trim(), btn = document.getElementById('batchBtn')
  if (!s || !u) { setStatus('batchStatus', 'error', '请填写 SESSDATA 和 UID'); return }
  btn.disabled = true; btn.textContent = '导出中...'
  setStatus('batchStatus', 'info', '正在获取关注列表...'); document.getElementById('batchProgress').classList.remove('hidden')
  clearEl('batchLog'); document.getElementById('batchResult').classList.add('hidden'); exportResults = []
  try {
    var f = await api('GET', 'https://api.bilibili.com/x/relation/followings?vmid=' + u + '&pn=1&ps=50&order=desc', 'SESSDATA=' + s)
    if (f.code !== 0) throw new Error(f.message || '获取关注列表失败')
    var list = f.data.list || [], total = f.data.total || list.length
    logEl('batchLog', '共获取 ' + total + ' 个UP主，开始查询关注时间...')
    var allItems = [].concat(list), page = 2
    while (allItems.length < total) {
      var r = await api('GET', 'https://api.bilibili.com/x/relation/followings?vmid=' + u + '&pn=' + page + '&ps=50&order=desc', 'SESSDATA=' + s)
      if (r.code === 0 && r.data.list) allItems.push.apply(allItems, r.data.list)
      page++; if (page > 50) break
    }
    var results = []
    for (var i = 0; i < allItems.length; i++) {
      var up = allItems[i]
      document.getElementById('batchProgressInner').style.width = ((i + 1) / allItems.length * 100) + '%'
      try {
        var t = await api('GET', 'https://api.bilibili.com/x/space/acc/relation?mid=' + up.mid, 'SESSDATA=' + s)
        var ft = '未知'
        if (t.code === 0 && t.data && t.data.relation) { var ts = t.data.relation.mtime; if (ts) ft = new Date(ts * 1000).toLocaleString('zh-CN', { hour12: false }) }
        results.push({ name: up.uname, mid: up.mid, follow_time: ft })
      } catch(e) { results.push({ name: up.uname, mid: up.mid, follow_time: '查询失败' }) }
      await new Promise(function(r) { setTimeout(r, 500) })
    }
    exportResults = results; document.getElementById('batchProgress').classList.add('hidden')
    setStatus('batchStatus', 'success', '导出完成！共 ' + results.length + ' 条记录')
    document.getElementById('batchDoneMsg').textContent = '导出完成！共 ' + results.length + ' 条记录'
    document.getElementById('batchResult').classList.remove('hidden')
  } catch(e) { document.getElementById('batchProgress').classList.add('hidden'); setStatus('batchStatus', 'error', '' + e.message) }
  finally { btn.disabled = false; btn.textContent = '开始导出' }
}
function downloadExcel() {
  if (!exportResults.length) return
  var d = exportResults.map(function(r, i) { return { '序号': i + 1, 'UP主名称': r.name, 'UID': r.mid, '关注时间': r.follow_time } })
  var ws = XLSX.utils.json_to_sheet(d); ws['!cols'] = [{wch:6}, {wch:24}, {wch:14}, {wch:20}]
  var wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '关注列表'); XLSX.writeFile(wb, 'b站关注列表.xlsx')
}
async function startSingleQuery() {
  var s = document.getElementById('singleSessdata').value.trim(), tu = document.getElementById('singleUid').value.trim(), btn = document.getElementById('singleBtn')
  if (!s || !tu) { setStatus('singleStatus', 'error', '请填写 SESSDATA 和 UID'); return }
  btn.disabled = true; btn.textContent = '查询中...'; setStatus('singleStatus', 'info', '正在查询...'); document.getElementById('singleResult').innerHTML = ''
  try {
    var info = await api('GET', 'https://api.bilibili.com/x/space/acc/info?mid=' + tu, 'SESSDATA=' + s)
    var name = (info.code === 0 && info.data) ? info.data.name : '未知用户'
    var rel = await api('GET', 'https://api.bilibili.com/x/space/acc/relation?mid=' + tu, 'SESSDATA=' + s)
    var following = false, ft = ''
    if (rel.code === 0 && rel.data && rel.data.relation) {
      following = rel.data.relation.attribute !== 0
      if (following && rel.data.relation.mtime) ft = new Date(rel.data.relation.mtime * 1000).toLocaleString('zh-CN', { hour12: false })
    }
    if (following) {
      document.getElementById('singleResult').innerHTML = '<div class="result-card following"><h3>' + name + '</h3><p><b>UID:</b> ' + tu + '</p><p style="color:#00A1D6;">已关注 - ' + ft + '</p></div>'
      setStatus('singleStatus', 'success', '已关注 ' + name)
    } else {
      document.getElementById('singleResult').innerHTML = '<div class="result-card not-following"><h3>' + name + '</h3><p><b>UID:</b> ' + tu + '</p><p style="color:#dc3545;">未关注</p></div>'
      setStatus('singleStatus', 'warning', '你尚未关注 ' + name)
    }
  } catch(e) { setStatus('singleStatus', 'error', '' + e.message) }
  finally { btn.disabled = false; btn.textContent = '查询' }
}
</script>
</body>
</html>`

// ── 主处理逻辑 ──────────────────────────────────────────────

async function handleRequest(request) {
    const url = new URL(request.url)

    // 根路径 → 返回前端页面
    if (request.method === "GET" && url.pathname === "/" && !url.searchParams.has("url")) {
        return new Response(HTML_PAGE, {
            headers: { "Content-Type": "text/html; charset=utf-8" },
        })
    }

    // CORS 预检
    if (request.method === "OPTIONS") {
        return new Response(null, {
            headers: {
                "Access-Control-Allow-Origin": "*",
                "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
                "Access-Control-Allow-Headers": "Content-Type, Cookie",
            },
        })
    }

    const target = url.searchParams.get("url")
    if (!target) {
        return new Response("Missing ?url= parameter", { status: 400 })
    }

    // 构建转发请求
    const headers = new Headers({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": "https://www.bilibili.com/",
        "Origin": "https://www.bilibili.com",
    })

    const clientCookie = url.searchParams.get("cookie")
    if (clientCookie) {
        headers.set("Cookie", clientCookie)
    }

    try {
        const resp = await fetch(decodeURIComponent(target), { headers })
        const body = await resp.text()

        const contentType = resp.headers.get("content-type") || ""
        const isJSON = contentType.includes("json") || body.trim().startsWith("{")

        const corsHeaders = {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type, Cookie",
        }

        if (isJSON) {
            const responseHeaders = {
                ...corsHeaders,
                "Content-Type": "application/json; charset=utf-8",
            }

            // 扫码轮询接口：提取 Set-Cookie 中的 SESSDATA 塞入响应体
            if (target.includes("qrcode/poll")) {
                const setCookie = resp.headers.get("Set-Cookie")
                let sessdata = ""
                if (setCookie) {
                    for (const part of setCookie.split(";")) {
                        const trimmed = part.trim()
                        if (trimmed.startsWith("SESSDATA=")) {
                            sessdata = trimmed.substring(9)
                            break
                        }
                    }
                }
                if (sessdata) {
                    try {
                        const json = JSON.parse(body)
                        json._sessdata = sessdata
                        return new Response(JSON.stringify(json), { headers: responseHeaders })
                    } catch (_) {}
                }
            }

            return new Response(body, { headers: responseHeaders })
        } else {
            return new Response(body, { headers: { ...corsHeaders, "Content-Type": contentType || "text/plain" } })
        }
    } catch (e) {
        return new Response(JSON.stringify({ error: e.message }), {
            status: 500,
            headers: { "Access-Control-Allow-Origin": "*", "Content-Type": "application/json" },
        })
    }
}
