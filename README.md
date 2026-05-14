# 🎬 B站关注列表导出工具

批量导出 B站关注列表，包含 **UP主名称、UID、关注时间**，支持导出为 Excel 文件。

## 📦 下载与使用

### 方式一：Windows 客户端（GUI）

直接从 Release 下载即开即用：

[⬇️ 下载 v2.0 (B.exe)](https://github.com/Hale51679/biliup_subscribe_time_list/releases/download/v2.0/B.exe)

> 无需安装 Python，下载后双击运行即可。如遇杀软误报请添加信任。

### 方式二：在线网页版（Streamlit）

扫码登录即可使用，无需下载任何文件：

[🌐 在线使用](https://biliup-subscribe-timelist.streamlit.app/) — 绑定本仓库部署


### 方式三：本地运行源码

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py    # Web 版
# 或
python bilibili_gui.py            # GUI 版
```

## ✨ 功能特点

- **扫码登录** — 工具内直接生成二维码，B站 App 扫码即登，无需手动填 Cookie
- **批量导出** — 一键导出所有关注 UP 主的名称、UID 和关注时间，实时显示进度
- **单个查询** — 输入 UID 快速查询是否关注了某个 UP 主及具体关注时间
- **Excel 导出** — 带格式的表格文件，按关注时间排序
- **自动获取 UID** — 扫码后自动填入你的 UID，防止填错
- **风控应对** — 自动重试 + 指数退避，应对 B 站 API 限流（-799）

## 📄 Excel 输出格式

| 序号 | UP主名称 | UID | 关注时间 |
|---|---|---|---|
| 1 | UP主A | 123456 | 2021-03-15 10:30:00 |
| 2 | UP主B | 789012 | 2022-08-20 14:22:00 |

## 🔧 技术栈

- **前端** — Streamlit / Tkinter
- **后端** — Python + Requests
- **登录** — B站 OAuth 扫码登录
- **导出** — openpyxl

## 📝 免责声明

本工具仅用于个人学习与数据备份，请合理使用 B站 API，避免高频请求。请遵守 B站用户协议。
