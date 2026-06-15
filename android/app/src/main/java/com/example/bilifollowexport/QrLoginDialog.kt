package com.example.bilifollowexport

import android.graphics.Bitmap
import android.graphics.Color
import android.os.Bundle
import android.view.LayoutInflater
import android.view.ViewGroup
import android.widget.ImageButton
import android.widget.ImageView
import android.widget.TextView
import androidx.fragment.app.DialogFragment
import androidx.lifecycle.lifecycleScope
import com.google.android.material.button.MaterialButton
import com.google.zxing.BarcodeFormat
import com.google.zxing.qrcode.QRCodeWriter
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.Job
import kotlinx.coroutines.delay
import kotlinx.coroutines.isActive
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import okhttp3.OkHttpClient
import okhttp3.Request
import java.util.concurrent.TimeUnit

class QrLoginDialog : DialogFragment() {

    private lateinit var ivQrCode: ImageView
    private lateinit var tvStatus: TextView
    private lateinit var tvTimer: TextView
    private lateinit var btnRefresh: MaterialButton
    private lateinit var btnManualInput: MaterialButton
    private lateinit var btnClose: ImageButton

    private var qrcodeKey: String? = null
    private var qrUrl: String? = null
    private var pollingJob: Job? = null
    private var timerJob: Job? = null
    private var qrExpireTime = 0L

    private val client = OkHttpClient.Builder()
        .connectTimeout(15, TimeUnit.SECONDS)
        .readTimeout(30, TimeUnit.SECONDS)
        .build()

    private val API_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

    var onLoginSuccess: ((sessdata: String) -> Unit)? = null

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setStyle(STYLE_NORMAL, android.R.style.Theme_Material_Light_Dialog)
    }

    override fun onCreateView(
        inflater: LayoutInflater,
        container: ViewGroup?,
        savedInstanceState: Bundle?
    ): android.view.View? {
        val view = inflater.inflate(R.layout.dialog_qr_login, container, false)

        ivQrCode = view.findViewById(R.id.ivQrCode)
        tvStatus = view.findViewById(R.id.tvQrStatus)
        tvTimer = view.findViewById(R.id.tvQrTimer)
        btnRefresh = view.findViewById(R.id.btnRefreshQr)
        btnManualInput = view.findViewById(R.id.btnManualInput)
        btnClose = view.findViewById(R.id.btnClose)

        btnClose.setOnClickListener { dismiss() }
        btnRefresh.setOnClickListener { generateQrCode() }
        btnManualInput.setOnClickListener { dismiss() }

        dialog?.setCanceledOnTouchOutside(false)

        // 自动开始生成二维码
        view.post { generateQrCode() }

        return view
    }

    override fun onStart() {
        super.onStart()
        dialog?.window?.setLayout(
            ViewGroup.LayoutParams.MATCH_PARENT,
            ViewGroup.LayoutParams.WRAP_CONTENT
        )
    }

    override fun onDestroyView() {
        pollingJob?.cancel()
        timerJob?.cancel()
        super.onDestroyView()
    }

    private fun generateQrCode() {
        pollingJob?.cancel()
        timerJob?.cancel()
        qrcodeKey = null
        qrUrl = null

        tvStatus.text = "正在生成二维码..."
        tvTimer.text = ""
        ivQrCode.setImageDrawable(null)
        btnRefresh.isEnabled = false

        lifecycleScope.launch {
            try {
                val (key, url) = withContext(Dispatchers.IO) {
                    requestQrGenerate()
                }

                qrcodeKey = key
                qrUrl = url
                qrExpireTime = System.currentTimeMillis() + 180_000L

                // 生成二维码图片
                val qrBitmap = withContext(Dispatchers.Default) {
                    generateQrBitmap(url, 480)
                }

                if (!isAdded) return@launch

                ivQrCode.setImageBitmap(qrBitmap)
                tvStatus.text = "请用 B站 App 扫码"
                btnRefresh.isEnabled = true

                // 同时启动倒计时和轮询
                startTimer()
                startPolling()
            } catch (e: Exception) {
                if (isAdded) {
                    tvStatus.text = "❌ 生成失败: ${e.message}"
                    btnRefresh.isEnabled = true
                }
            }
        }
    }

    private fun startTimer() {
        timerJob = lifecycleScope.launch {
            while (isActive) {
                val remaining = (qrExpireTime - System.currentTimeMillis()) / 1000
                if (remaining <= 0) {
                    tvTimer.text = "二维码已过期，请点击刷新"
                    break
                }
                val min = remaining / 60
                val sec = remaining % 60
                tvTimer.text = "二维码有效 ${min}:%02d".format(sec)
                delay(1000)
            }
        }
    }

    private fun startPolling() {
        pollingJob = lifecycleScope.launch {
            while (isActive) {
                val key = qrcodeKey ?: break

                val result = withContext(Dispatchers.IO) {
                    pollQrLogin(key)
                }

                when (result.status) {
                    "ok" -> {
                        if (isAdded) {
                            tvStatus.text = "✅ 登录成功！"
                            tvTimer.text = ""
                            onLoginSuccess?.invoke(result.sessdata)
                            delay(800)
                            dismiss()
                        }
                        break
                    }
                    "scanned" -> {
                        tvStatus.text = "📱 已扫码，请在手机上确认..."
                    }
                    "expired" -> {
                        tvStatus.text = "⏰ 二维码已过期，点击刷新"
                        tvTimer.text = ""
                        break
                    }
                    else -> {
                        if (tvStatus.text != "📱 已扫码，请在手机上确认...") {
                            tvStatus.text = "等待扫码..."
                        }
                    }
                }

                delay(2000)
            }
        }
    }

    // ── 网络请求 ──

    private fun requestQrGenerate(): Pair<String, String> {
        val request = Request.Builder()
            .url("https://passport.bilibili.com/x/passport-login/web/qrcode/generate?source=main-fe-header")
            .header("User-Agent", API_UA)
            .build()

        return client.newCall(request).execute().use { response ->
            val body = response.body?.string() ?: "{}"
            val gson = com.google.gson.Gson()
            val data = gson.fromJson(body, Map::class.java)
            val dataObj = data["data"] as? Map<*, *>
            val key = dataObj?.get("qrcode_key") as? String
                ?: throw Exception("获取二维码Key失败")
            val url = dataObj?.get("url") as? String
                ?: throw Exception("获取二维码URL失败")
            Pair(key, url)
        }
    }

    private data class PollResult(val status: String, val sessdata: String)

    private fun pollQrLogin(qrcodeKey: String): PollResult {
        val request = Request.Builder()
            .url("https://passport.bilibili.com/x/passport-login/web/qrcode/poll?qrcode_key=$qrcodeKey")
            .header("User-Agent", API_UA)
            .build()

        return client.newCall(request).execute().use { response ->
            val body = response.body?.string() ?: "{}"
            val gson = com.google.gson.Gson()
            val data = gson.fromJson(body, Map::class.java)
            val dataObj = data["data"] as? Map<*, *>
            val code = (dataObj?.get("code") as? Double)?.toInt() ?: -1

            when (code) {
                0 -> {
                    var sessdata = ""
                    for (cookie in response.headers("Set-Cookie")) {
                        if (cookie.trim().startsWith("SESSDATA=")) {
                            sessdata = cookie.split("SESSDATA=")[1].split(";")[0]
                            sessdata = java.net.URLDecoder.decode(sessdata, "UTF-8")
                            break
                        }
                    }
                    PollResult("ok", sessdata)
                }
                86038 -> PollResult("expired", "")
                86090 -> PollResult("scanned", "")
                else -> PollResult("waiting", "")
            }
        }
    }

    // ── 生成二维码 Bitmap ──

    private fun generateQrBitmap(content: String, size: Int): Bitmap? {
        return try {
            val writer = QRCodeWriter()
            val bitMatrix = writer.encode(content, BarcodeFormat.QR_CODE, size, size)
            val bitmap = Bitmap.createBitmap(size, size, Bitmap.Config.RGB_565)
            for (x in 0 until size) {
                for (y in 0 until size) {
                    bitmap.setPixel(x, y, if (bitMatrix[x, y]) Color.BLACK else Color.WHITE)
                }
            }
            bitmap
        } catch (e: Exception) {
            null
        }
    }
}
