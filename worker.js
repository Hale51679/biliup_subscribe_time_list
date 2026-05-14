// Cloudflare Worker — B站 API 代理
// 部署方法: https://dash.cloudflare.com → Workers & Pages → 创建 Worker → 粘贴此代码 → 部署
addEventListener("fetch", event => {
    event.respondWith(handleRequest(event.request))
})

async function handleRequest(request) {
    const url = new URL(request.url)
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

    // 如果前端传了 Cookie，透传给 B站 API
    const clientCookie = url.searchParams.get("cookie")
    if (clientCookie) {
        headers.set("Cookie", clientCookie)
    }

    try {
        const resp = await fetch(decodeURIComponent(target), { headers })
        const body = await resp.text()

        // 判断响应类型
        const contentType = resp.headers.get("content-type") || ""
        const isJSON = contentType.includes("json") || body.trim().startsWith("{")

        // 构造响应，添加 CORS 头
        const corsHeaders = {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
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
                    // Set-Cookie 可能有多段，找到 SESSDATA=xxx 部分
                    for (const part of setCookie.split(";")) {
                        const trimmed = part.trim()
                        if (trimmed.startsWith("SESSDATA=")) {
                            sessdata = trimmed.substring(9)
                            break
                        }
                    }
                }
                if (sessdata) {
                    // 把 SESSDATA 注入到 JSON 响应体中
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
