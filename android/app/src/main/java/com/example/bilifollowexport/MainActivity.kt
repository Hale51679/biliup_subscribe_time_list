package com.example.bilifollowexport

import android.content.ClipData
import android.content.ClipboardManager
import android.content.Context
import android.os.Bundle
import android.widget.Toast
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.ViewCompat
import androidx.core.view.WindowInsetsCompat
import androidx.lifecycle.lifecycleScope
import com.google.android.material.button.MaterialButton
import com.google.android.material.textfield.TextInputEditText
import com.google.gson.Gson
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.delay
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import okhttp3.OkHttpClient
import okhttp3.Request
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale
import java.util.concurrent.TimeUnit

// ── B站 API 返回数据模型 ──────────────────────────────────────────────

data class FollowingItem(
    val mid: Long,
    val uname: String
)

data class RelationInfo(
    val mtime: Long?,    // 关注时间戳
    val attribute: Int?  // 0=未关注 非0=已关注
)

// ── 主页 ──────────────────────────────────────────────────────────────

class MainActivity : AppCompatActivity() {

    // 视图
    private lateinit var etSessdata: TextInputEditText
    private lateinit var etUid: TextInputEditText
    private lateinit var etTargetUid: TextInputEditText
    private lateinit var btnStart: MaterialButton
    private lateinit var btnQuery: MaterialButton
    private lateinit var btnQrLogin: MaterialButton
    private lateinit var btnPause: MaterialButton
    private lateinit var btnCopy: MaterialButton
    private lateinit var progressLayout: android.widget.LinearLayout
    private lateinit var progressBar: android.widget.ProgressBar
    private lateinit var tvProgressText: android.widget.TextView
    private lateinit var tvLog: android.widget.TextView

    // 暂停控制
    @Volatile
    private var isExportPaused = false

    // 纯结果（不含控制台日志）
    private val resultBuffer = StringBuilder()

    // 网络
    private val client = OkHttpClient.Builder()
        .connectTimeout(10, TimeUnit.SECONDS)
        .readTimeout(15, TimeUnit.SECONDS)
        .build()

    private val gson = Gson()

    // B站 API 常量
    private val API_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_main)
        ViewCompat.setOnApplyWindowInsetsListener(findViewById(R.id.main)) { v, insets ->
            val systemBars = insets.getInsets(WindowInsetsCompat.Type.systemBars())
            v.setPadding(systemBars.left, systemBars.top, systemBars.right, systemBars.bottom)
            insets
        }

        initViews()
        setupListeners()
    }

    private fun initViews() {
        etSessdata = findViewById(R.id.etSessdata)
        etUid = findViewById(R.id.etUid)
        etTargetUid = findViewById(R.id.etTargetUid)
        btnStart = findViewById(R.id.btnStart)
        btnQuery = findViewById(R.id.btnQuery)
        btnQrLogin = findViewById(R.id.btnQrLogin)
        btnPause = findViewById(R.id.btnPause)
        btnCopy = findViewById(R.id.btnCopy)
        progressLayout = findViewById(R.id.progressLayout)
        progressBar = findViewById(R.id.progressBar)
        tvProgressText = findViewById(R.id.tvProgressText)
        tvLog = findViewById(R.id.tvLog)
    }

    private fun setupListeners() {
        btnStart.setOnClickListener {
            val sessdata = etSessdata.text?.toString()?.trim() ?: ""
            val uid = etUid.text?.toString()?.trim() ?: ""
            if (sessdata.isEmpty() || uid.isEmpty()) {
                toast("请填写 SESSDATA 和 UID")
                return@setOnClickListener
            }
            startExport(sessdata, uid)
        }

        btnQuery.setOnClickListener {
            val sessdata = etSessdata.text?.toString()?.trim() ?: ""
            val targetUid = etTargetUid.text?.toString()?.trim() ?: ""
            if (sessdata.isEmpty() || targetUid.isEmpty()) {
                toast("请填写 SESSDATA 和目标 UID")
                return@setOnClickListener
            }
            querySingle(sessdata, targetUid)
        }

        btnQrLogin.setOnClickListener {
            val dialog = QrLoginDialog()
            dialog.onLoginSuccess = { sessdata ->
                if (sessdata.isNotEmpty()) {
                    etSessdata.setText(sessdata)
                    toast("✅ 扫码登录成功！")
                    log("✅ 扫码登录成功，正在获取 UID...")

                    // 自动获取当前用户的 UID
                    lifecycleScope.launch {
                        val uid = withContext(Dispatchers.IO) {
                            getMyUid(sessdata)
                        }
                        if (uid != null) {
                            etUid.setText(uid.toString())
                            log("✅ 已自动填入 UID: $uid")
                        } else {
                            log("⚠️ 获取 UID 失败，请手动填写")
                        }
                    }
                } else {
                    toast("⚠️ 获取 SESSDATA 失败，请手动输入")
                }
            }
            dialog.show(supportFragmentManager, "QrLogin")
        }

        btnPause.setOnClickListener {
            isExportPaused = !isExportPaused
            btnPause.text = if (isExportPaused) "▶️ 继续" else "⏸ 暂停"
            if (isExportPaused) {
                log("⏸ 已暂停，点击「继续」恢复导出")
            } else {
                log("▶️ 继续导出...")
            }
        }

        btnCopy.setOnClickListener {
            val text = resultBuffer.toString()
            if (text.isBlank()) {
                toast("还没有可复制的内容")
                return@setOnClickListener
            }
            val clipboard = getSystemService(Context.CLIPBOARD_SERVICE) as ClipboardManager
            val clip = ClipData.newPlainText("B站关注列表", text)
            clipboard.setPrimaryClip(clip)
            toast("✅ 已复制 ${text.lines().size - 1} 条结果到剪贴板")
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  批量导出
    // ═══════════════════════════════════════════════════════════════

    private fun startExport(sessdata: String, uid: String) {
        btnStart.isEnabled = false
        progressLayout.visibility = android.view.View.VISIBLE
        tvLog.text = ""
        isExportPaused = false
        btnPause.text = "⏸ 暂停"
        resultBuffer.clear()

        lifecycleScope.launch {
            try {
                val headers = makeHeaders(sessdata)
                log("正在获取关注列表...")

                // 1. 获取关注列表
                val followings = withContext(Dispatchers.IO) {
                    getFollowingList(uid, headers)
                }
                if (followings.isEmpty()) {
                    log("❌ 获取关注列表失败，请检查 SESSDATA 和 UID")
                    toast("获取失败，请检查登录信息")
                    return@launch
                }

                log("✅ 共 ${followings.size} 个UP主，开始查询关注时间...\n")

                // 2. 逐个查询关注时间
                for ((i, up) in followings.withIndex()) {
                    // 暂停检查 - 如果暂停则循环等待直到恢复
                    while (isExportPaused) {
                        delay(500)
                    }

                    val (mtime, isFollowing) = withContext(Dispatchers.IO) {
                        getFollowTime(up.mid.toString(), headers)
                    }
                    val timeStr = if (mtime != null) timestampToStr(mtime) else "未知"
                    val status = if (isFollowing) timeStr else "❌ 未关注"
                    val line = "[${i + 1}/${followings.size}] ${up.uname} → $status"
                    val cleanLine = "${up.uname}\t$status"

                    // 保存纯结果（含表头）
                    if (i == 0) resultBuffer.appendLine("UP主名称\t关注时间")
                    resultBuffer.appendLine(cleanLine)

                    // 更新UI - 逐行追加
                    withContext(Dispatchers.Main) {
                        tvLog.append("\n$line")
                        progressBar.progress = ((i + 1) * 100 / followings.size)
                        tvProgressText.text = "${i + 1}/${followings.size}"
                        // 自动滚到日志底部
                        findViewById<androidx.core.widget.NestedScrollView>(R.id.logScrollView).post {
                            findViewById<androidx.core.widget.NestedScrollView>(R.id.logScrollView).fullScroll(
                                android.view.View.FOCUS_DOWN
                            )
                        }
                    }

                    // 随机延时避免风控（1.2~2.2秒）
                    delay(1200L + kotlin.random.Random.nextLong(1000))
                }

                log("\n🎉 导出完成！共 ${followings.size} 条记录")
                toast("导出完成！")

            } catch (e: Exception) {
                log("❌ 错误: ${e.message}")
                toast("出错: ${e.message}")
            } finally {
                btnStart.isEnabled = true
                progressLayout.visibility = android.view.View.GONE
            }
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  单个查询
    // ═══════════════════════════════════════════════════════════════

    private fun querySingle(sessdata: String, targetUid: String) {
        btnQuery.isEnabled = false
        lifecycleScope.launch {
            try {
                val headers = makeHeaders(sessdata)

                // 获取UP主名称
                val name = withContext(Dispatchers.IO) {
                    getUserInfo(targetUid, headers)
                }

                // 查询关注状态
                val (mtime, isFollowing) = withContext(Dispatchers.IO) {
                    getFollowTime(targetUid, headers)
                }

                val result = buildString {
                    appendLine("═══════════════════════════")
                    appendLine("  UP主: $name")
                    appendLine("  UID: $targetUid")
                    appendLine("  ─────────────────────────")
                    if (isFollowing) {
                        appendLine("  ✅ 已关注")
                        if (mtime != null) {
                            appendLine("  关注时间: ${timestampToStr(mtime)}")
                        }
                    } else {
                        appendLine("  ❌ 未关注")
                    }
                    appendLine("═══════════════════════════")
                }

                log(result)
            } catch (e: Exception) {
                log("❌ 查询失败: ${e.message}")
            } finally {
                btnQuery.isEnabled = true
            }
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  B站 API 核心方法
    // ═══════════════════════════════════════════════════════════════

    private fun makeHeaders(sessdata: String): Map<String, String> {
        val sd = sessdata.trim()
            .removePrefix("SESSDATA=")
            .removePrefix("sessdata=")
        return mapOf(
            "User-Agent" to API_UA,
            "Referer" to "https://www.bilibili.com",
            "Cookie" to "SESSDATA=$sd"
        )
    }

    private fun getFollowingList(uid: String, headers: Map<String, String>): List<FollowingItem> {
        val result = mutableListOf<FollowingItem>()
        var page = 1
        val pageSize = 50

        while (true) {
            val url = "https://api.bilibili.com/x/relation/followings?vmid=$uid&pn=$page&ps=$pageSize&order=desc"
            val resp = requestGet(url, headers)
            val data = gson.fromJson(resp, Map::class.java)
            val code = ((data["code"] as? Double)?.toInt()) ?: -1

            if (code != 0) {
                log("❌ 获取关注列表失败: code=$code")
                return emptyList()
            }

            val dataObj = data["data"] as? Map<*, *>
            val listRaw = dataObj?.get("list") as? List<*>
            val items = mutableListOf<FollowingItem>()

            if (listRaw != null) {
                for (item in listRaw) {
                    val map = item as? Map<*, *> ?: continue
                    val mid = (map["mid"] as? Double)?.toLong() ?: continue
                    val uname = map["uname"] as? String ?: ""
                    items.add(FollowingItem(mid, uname))
                }
            }

            if (items.isEmpty()) break
            result.addAll(items)

            val total = (dataObj?.get("total") as? Double)?.toInt() ?: 0
            log("  已获取 ${result.size}/$total 个UP主")
            if (result.size >= total) break
            page++
            Thread.sleep(kotlin.random.Random.nextLong(800, 1500))
        }

        return result
    }

    private fun getFollowTime(targetUid: String, headers: Map<String, String>): Pair<Long?, Boolean> {
        val url = "https://api.bilibili.com/x/space/acc/relation?mid=$targetUid"
        var mtime: Long? = null
        var isFollowing = false

        for (attempt in 1..5) {
            try {
                val resp = requestGet(url, headers)
                val data = gson.fromJson(resp, Map::class.java)
                val code = ((data["code"] as? Double)?.toInt()) ?: -1

                if (code == -799) {
                    // 风控限流 - 指数退避 + 随机抖动
                    val wait = (attempt * 3L) + kotlin.random.Random.nextLong(3)
                    log("⚠️ 风控限流，等待 ${wait}s（第${attempt}次重试）...")
                    Thread.sleep(wait * 1000)
                    continue
                }

                if (code == 0) {
                    val dataObj = data["data"] as? Map<*, *>
                    val relation = dataObj?.get("relation") as? Map<*, *>
                    if (relation != null) {
                        mtime = (relation["mtime"] as? Double)?.toLong()
                        // mtime=0 表示没有关注时间
                        if (mtime != null && mtime == 0L) mtime = null
                        val attr = (relation["attribute"] as? Double)?.toInt() ?: 0
                        isFollowing = attr != 0
                    }
                }
                break
            } catch (e: Exception) {
                log("⚠️ 请求失败: ${e.message}")
                Thread.sleep(2000)
            }
        }

        return Pair(mtime, isFollowing)
    }

    private fun getMyUid(sessdata: String): Long? {
        return try {
            val headers = makeHeaders(sessdata)
            val url = "https://api.bilibili.com/x/web-interface/nav"
            val request = Request.Builder()
                .url(url)
                .apply {
                    for ((k, v) in headers) addHeader(k, v)
                }
                .build()
            val resp = client.newCall(request).execute().use { it.body?.string() ?: "{}" }
            val data = gson.fromJson(resp, Map::class.java)
            if ((data["code"] as? Double)?.toInt() == 0) {
                val dataObj = data["data"] as? Map<*, *>
                val mid = dataObj?.get("mid")
                // mid 可能返回 Double 或 String
                (mid as? Double)?.toLong() ?: (mid as? String)?.toLongOrNull()
            } else null
        } catch (e: Exception) {
            null
        }
    }

    private fun getUserInfo(uid: String, headers: Map<String, String>): String {
        val url = "https://api.bilibili.com/x/space/acc/info?mid=$uid"
        return try {
            val resp = requestGet(url, headers)
            val data = gson.fromJson(resp, Map::class.java)
            if ((data["code"] as? Double)?.toInt() == 0) {
                val dataObj = data["data"] as? Map<*, *>
                dataObj?.get("name") as? String ?: "未知用户"
            } else "未知用户"
        } catch (e: Exception) {
            "未知用户"
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  工具方法
    // ═══════════════════════════════════════════════════════════════

    private fun requestGet(url: String, headers: Map<String, String>): String {
        val request = Request.Builder()
            .url(url)
            .apply {
                for ((k, v) in headers) {
                    addHeader(k, v)
                }
            }
            .build()

        return client.newCall(request).execute().use { response ->
            response.body?.string() ?: "{}"
        }
    }

    private fun timestampToStr(timestamp: Long): String {
        val sdf = SimpleDateFormat("yyyy-MM-dd HH:mm:ss", Locale.CHINA)
        return sdf.format(Date(timestamp * 1000))
    }

    private fun log(msg: String) {
        runOnUiThread {
            val current = tvLog.text.toString()
            tvLog.text = if (current == "等待操作...") msg else "$current\n$msg"
        }
    }

    private fun toast(msg: String) {
        Toast.makeText(this, msg, Toast.LENGTH_SHORT).show()
    }
}
