"""
B站关注列表导出工具 - GUI版本 v2
打包命令：
  pyinstaller --onefile --windowed --name "B站关注导出" --icon "bilibili_icon.ico" bilibili_gui.py
"""

# ── DPI感知（必须在任何tkinter导入之前）──
import ctypes, sys
try:
    if sys.platform == "win32":
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import requests
import time
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BLUE  = "#00A1D6"
BLUE2 = "#0095C5"
BG    = "#f5f5f5"
WHITE = "#ffffff"

# ───────────────────────────── 爬虫核心逻辑 ─────────────────────────────

def make_headers(sessdata):
    sessdata = sessdata.strip()
    if sessdata.lower().startswith("sessdata="):
        sessdata = sessdata[9:]
    return {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://www.bilibili.com",
        "Cookie": f"SESSDATA={sessdata}",
    }

def get_following_list(uid, headers, log):
    followings = []
    page, page_size = 1, 50
    while True:
        url = f"https://api.bilibili.com/x/relation/followings?vmid={uid}&pn={page}&ps={page_size}&order=desc"
        resp = requests.get(url, headers=headers, timeout=10)
        data = resp.json()
        if data.get("code") != 0:
            log(f"❌ 获取关注列表失败: {data.get('message')}")
            return []
        items = data["data"].get("list", [])
        if not items:
            break
        for item in items:
            followings.append({"mid": item["mid"], "name": item["uname"]})
        total = data["data"].get("total", 0)
        log(f"📋 已获取 {len(followings)}/{total} 个UP主")
        if len(followings) >= total:
            break
        page += 1
        time.sleep(0.5)
    return followings

def get_follow_time(target_uid, headers, log, max_retries=5):
    url = f"https://api.bilibili.com/x/space/acc/relation?mid={target_uid}"
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.get(url, headers=headers, timeout=10)
            data = resp.json()
            code = data.get("code")
            if code == -799:
                wait = attempt * 2
                log(f"  ⚠️ 风控限流，等待 {wait}s 后重试（{attempt}/{max_retries}）...")
                time.sleep(wait)
                continue
            if code == 0:
                relation = data.get("data", {}).get("relation")
                if not relation:
                    return None, False
                mtime = relation.get("mtime", 0)
                attribute = relation.get("attribute", 0)
                is_following = attribute != 0
                return (mtime if mtime and mtime > 0 else None), is_following
            log(f"  ⚠️ uid={target_uid} 异常 code={code}: {data.get('message')}")
            return None, False
        except Exception as e:
            log(f"  ⚠️ 查询 {target_uid} 出错: {e}")
            time.sleep(2)
    log(f"  ❌ uid={target_uid} 重试 {max_retries} 次后仍失败")
    return None, False

def get_user_info(uid, headers):
    url = f"https://api.bilibili.com/x/space/acc/info?mid={uid}"
    try:
        resp = requests.get(url, headers=headers, timeout=10)
        data = resp.json()
        if data.get("code") == 0:
            return data["data"].get("name", "未知用户")
    except Exception:
        pass
    return "未知用户"

def timestamp_to_str(ts):
    if ts:
        return datetime.datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")
    return "未知"

def export_to_excel(data, filename):
    wb = Workbook()
    ws = wb.active
    ws.title = "关注列表"
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", start_color="00A1D6")
    header_align = Alignment(horizontal="center", vertical="center")
    cell_align = Alignment(horizontal="left", vertical="center")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    headers = ["序号", "UP主名称", "UID", "关注时间"]
    col_widths = [8, 30, 18, 22]
    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border
        ws.column_dimensions[cell.column_letter].width = w
    ws.row_dimensions[1].height = 28
    alt_fill = PatternFill("solid", start_color="F0F9FF")
    for i, row in enumerate(data, 1):
        excel_row = i + 1
        values = [i, row["name"], row["mid"], row["follow_time"]]
        fill = alt_fill if i % 2 == 0 else None
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=excel_row, column=col, value=val)
            cell.font = Font(name="Arial", size=10)
            cell.alignment = cell_align if col != 1 else Alignment(horizontal="center", vertical="center")
            cell.border = border
            if fill:
                cell.fill = fill
        ws.row_dimensions[excel_row].height = 20
    ws.freeze_panes = "A2"
    wb.save(filename)

# ───────────────────────────── GUI ─────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("B站关注列表导出工具")
        self.resizable(True, True)
        self.minsize(520, 640)
        # 设置窗口图标（打包后从exe同目录读取，开发时从脚本同目录读取）
        try:
            import os
            base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            icon_path = os.path.join(base, "bilibili_icon.ico")
            self.iconbitmap(icon_path)
        except Exception:
            pass
        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _entry(self, parent, textvariable=None, show=None):
        kw = dict(font=("Consolas", 10), relief="flat", bg=WHITE,
                  highlightthickness=1, highlightbackground="#cccccc", highlightcolor=BLUE)
        if textvariable: kw["textvariable"] = textvariable
        if show: kw["show"] = show
        return tk.Entry(parent, **kw)

    def _text_box(self, parent, height=3):
        return tk.Text(parent, height=height, font=("Consolas", 10), relief="flat",
                       bg=WHITE, highlightthickness=1, highlightbackground="#cccccc",
                       highlightcolor=BLUE, wrap="word")

    def _btn(self, parent, text, command, small=False):
        b = tk.Button(parent, text=text,
                      font=("微软雅黑", 9 if small else 11, "bold"),
                      bg=BLUE, fg=WHITE, relief="flat", cursor="hand2",
                      activebackground=BLUE2, activeforeground=WHITE, command=command)
        b.bind("<Enter>", lambda e: b.configure(bg=BLUE2))
        b.bind("<Leave>", lambda e: b.configure(bg=BLUE))
        return b

    def _log_widget(self, parent):
        frame = tk.Frame(parent, bg=BG)
        box = tk.Text(frame, font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
                      relief="flat", state="disabled", wrap="word")
        scroll = tk.Scrollbar(frame, command=box.yview)
        box.configure(yscrollcommand=scroll.set)
        box.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        return frame, box

    def _build_ui(self):
        self.configure(bg=BG)

        # 标题
        title_frame = tk.Frame(self, bg=BLUE, pady=14)
        title_frame.pack(fill="x")
        tk.Label(title_frame, text="B站关注列表导出工具",
                 font=("微软雅黑", 16, "bold"), bg=BLUE, fg=WHITE).pack()
        tk.Label(title_frame, text="导出你关注的UP主名称、UID与关注时间",
                 font=("微软雅黑", 9), bg=BLUE, fg="#d0f0ff").pack()

        # Tab样式
        style = ttk.Style()
        style.theme_use("default")
        style.configure("TNotebook", background=BG, borderwidth=0)
        style.configure("TNotebook.Tab", font=("微软雅黑", 10),
                        padding=[16, 6], background="#e0e0e0", foreground="#555")
        style.map("TNotebook.Tab",
                  background=[("selected", BLUE)],
                  foreground=[("selected", WHITE)])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        tab_batch  = tk.Frame(nb, bg=BG)
        tab_single = tk.Frame(nb, bg=BG)
        nb.add(tab_batch,  text="  批量导出  ")
        nb.add(tab_single, text="  单个查询  ")

        self._build_batch_tab(tab_batch)
        self._build_single_tab(tab_single)

        tk.Label(self, text="💡 获取SESSDATA：浏览器登录B站 → F12 → Application → Cookies → bilibili.com → 复制SESSDATA值",
                 font=("微软雅黑", 8), bg=BG, fg="#999", wraplength=500).pack(pady=(4, 8))

    # ── 批量导出 ──────────────────────────────────────────────────────

    def _build_batch_tab(self, parent):
        p = dict(padx=24)
        tk.Frame(parent, bg=BG, height=10).pack()

        tk.Label(parent, text="你的B站 UID", font=("微软雅黑", 10), bg=BG, anchor="w").pack(fill="x", **p)
        self.uid_var = tk.StringVar()
        self._entry(parent, textvariable=self.uid_var).pack(fill="x", ipady=5, pady=(2,10), **p)

        tk.Label(parent, text="SESSDATA", font=("微软雅黑", 10), bg=BG, anchor="w").pack(fill="x", **p)
        self.sessdata_text = self._text_box(parent, height=3)
        self.sessdata_text.pack(fill="x", pady=(2,10), **p)

        tk.Label(parent, text="导出路径", font=("微软雅黑", 10), bg=BG, anchor="w").pack(fill="x", **p)
        pf = tk.Frame(parent, bg=BG)
        pf.pack(fill="x", pady=(2,10), **p)
        self.path_var = tk.StringVar(value="b站关注列表.xlsx")
        self._entry(pf, textvariable=self.path_var).pack(side="left", fill="x", expand=True, ipady=5)
        self._btn(pf, " 浏览 ", self._browse, small=True).pack(side="left", padx=(6,0), ipady=4)

        self.b_progress = ttk.Progressbar(parent, mode="determinate")
        self.b_progress.pack(fill="x", pady=(4,2), **p)
        self.b_status_var = tk.StringVar(value="就绪")
        tk.Label(parent, textvariable=self.b_status_var,
                 font=("微软雅黑", 9), bg=BG, fg="#666").pack(anchor="w", **p)

        log_frame, self.b_log_box = self._log_widget(parent)
        log_frame.pack(fill="both", expand=True, pady=(6,0), **p)

        self.run_btn = self._btn(parent, "  开始导出  ", self._start_batch)
        self.run_btn.pack(pady=14)

    # ── 单个查询 ──────────────────────────────────────────────────────

    def _build_single_tab(self, parent):
        p = dict(padx=24)
        tk.Frame(parent, bg=BG, height=10).pack()

        tk.Label(parent, text="SESSDATA", font=("微软雅黑", 10), bg=BG, anchor="w").pack(fill="x", **p)
        self.s_sessdata_text = self._text_box(parent, height=3)
        self.s_sessdata_text.pack(fill="x", pady=(2,10), **p)

        tk.Label(parent, text="UP主 UID（输入后按回车或点击查询）",
                 font=("微软雅黑", 10), bg=BG, anchor="w").pack(fill="x", **p)
        qf = tk.Frame(parent, bg=BG)
        qf.pack(fill="x", pady=(2,10), **p)
        self.query_var = tk.StringVar()
        qe = self._entry(qf, textvariable=self.query_var)
        qe.pack(side="left", fill="x", expand=True, ipady=5)
        qe.bind("<Return>", lambda e: self._start_single())
        self._btn(qf, " 查询 ", self._start_single, small=True).pack(side="left", padx=(6,0), ipady=4)

        # 结果卡片
        card = tk.Frame(parent, bg=WHITE, highlightthickness=1, highlightbackground="#e0e0e0")
        card.pack(fill="x", **p)
        inner = tk.Frame(card, bg=WHITE, padx=16, pady=14)
        inner.pack(fill="x")

        def card_col(row_frame, label, var, col, is_status=False):
            f = tk.Frame(row_frame, bg=WHITE)
            f.grid(row=0, column=col, sticky="w", padx=(0,40))
            tk.Label(f, text=label, font=("微软雅黑", 9), bg=WHITE, fg="#999").pack(anchor="w")
            lbl = tk.Label(f, textvariable=var, font=("微软雅黑", 12, "bold"), bg=WHITE, fg="#222")
            lbl.pack(anchor="w")
            row_frame.columnconfigure(col, weight=1)
            return lbl

        row1 = tk.Frame(inner, bg=WHITE); row1.pack(fill="x")
        self.s_name_var = tk.StringVar(value="—")
        self.s_uid_var  = tk.StringVar(value="—")
        card_col(row1, "UP主名称", self.s_name_var, 0)
        card_col(row1, "UID",     self.s_uid_var,  1)

        tk.Frame(inner, bg="#eeeeee", height=1).pack(fill="x", pady=10)

        row2 = tk.Frame(inner, bg=WHITE); row2.pack(fill="x")
        self.s_status_var = tk.StringVar(value="—")
        self.s_time_var   = tk.StringVar(value="—")
        self.s_status_lbl = card_col(row2, "关注状态", self.s_status_var, 0)
        card_col(row2, "关注时间", self.s_time_var, 1)

        log_frame, self.s_log_box = self._log_widget(parent)
        log_frame.pack(fill="both", expand=True, pady=(12,0), **p)
        tk.Frame(parent, bg=BG, height=14).pack()

    # ── 通用工具 ──────────────────────────────────────────────────────

    def _browse(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="b站关注列表.xlsx", title="选择导出位置")
        if path:
            self.path_var.set(path)

    def _log(self, box, msg):
        box.configure(state="normal")
        box.insert("end", msg + "\n")
        box.see("end")
        box.configure(state="disabled")

    def _clear_log(self, box):
        box.configure(state="normal")
        box.delete("1.0", "end")
        box.configure(state="disabled")

    # ── 批量导出线程 ──────────────────────────────────────────────────

    def _start_batch(self):
        uid = self.uid_var.get().strip()
        sessdata = self.sessdata_text.get("1.0", "end").strip()
        path = self.path_var.get().strip()
        if not uid or not sessdata:
            messagebox.showerror("缺少信息", "请填写 UID 和 SESSDATA")
            return
        if not path.endswith(".xlsx"):
            path += ".xlsx"
            self.path_var.set(path)
        self.run_btn.configure(state="disabled", text="  导出中...  ")
        self.b_progress["value"] = 0
        self._clear_log(self.b_log_box)
        threading.Thread(target=self._run_batch, args=(uid, sessdata, path), daemon=True).start()

    def _run_batch(self, uid, sessdata, path):
        headers = make_headers(sessdata)
        log = lambda m: self._log(self.b_log_box, m)
        try:
            self.b_status_var.set("正在获取关注列表...")
            followings = get_following_list(uid, headers, log)
            if not followings:
                messagebox.showerror("失败", "未能获取关注列表，请检查 UID 和 SESSDATA 是否正确")
                return
            total = len(followings)
            log(f"\n✅ 共获取 {total} 个UP主，开始查询关注时间...\n")
            results = []
            for i, up in enumerate(followings, 1):
                mtime, _ = get_follow_time(up["mid"], headers, log)
                follow_time = timestamp_to_str(mtime)
                results.append({"name": up["name"], "mid": up["mid"], "follow_time": follow_time})
                log(f"  [{i:3d}/{total}] {up['name'][:16]:<16} → {follow_time}")
                self.b_status_var.set(f"查询关注时间 {i}/{total}")
                self.b_progress["value"] = i / total * 100
                time.sleep(0.8)
            self.b_status_var.set("正在写入 Excel...")
            export_to_excel(results, path)
            log(f"\n🎉 导出完成！共 {len(results)} 条记录")
            log(f"📁 文件已保存到：{path}")
            self.b_status_var.set(f"完成！已导出 {len(results)} 条记录")
            messagebox.showinfo("完成", f"导出成功！共 {len(results)} 条记录\n\n文件位置：\n{path}")
        except Exception as e:
            log(f"\n❌ 发生错误: {e}")
            messagebox.showerror("错误", str(e))
        finally:
            self.run_btn.configure(state="normal", text="  开始导出  ")

    # ── 单个查询线程 ──────────────────────────────────────────────────

    def _start_single(self):
        target_uid = self.query_var.get().strip()
        sessdata = self.s_sessdata_text.get("1.0", "end").strip()
        if not target_uid or not sessdata:
            messagebox.showerror("缺少信息", "请填写 SESSDATA 和要查询的UP主UID")
            return
        self._clear_log(self.s_log_box)
        self.s_name_var.set("查询中...")
        self.s_uid_var.set(target_uid)
        self.s_status_var.set("—")
        self.s_time_var.set("—")
        threading.Thread(target=self._run_single, args=(target_uid, sessdata), daemon=True).start()

    def _run_single(self, target_uid, sessdata):
        headers = make_headers(sessdata)
        log = lambda m: self._log(self.s_log_box, m)
        try:
            log(f"🔍 正在查询 uid={target_uid}...")
            name = get_user_info(target_uid, headers)
            mtime, is_following = get_follow_time(target_uid, headers, log)
            self.s_name_var.set(name)
            self.s_uid_var.set(str(target_uid))
            if not is_following:
                self.s_status_var.set("未关注")
                self.s_status_lbl.configure(fg="#f25d8e")
                self.s_time_var.set("—")
                log(f"✅ {name}（{target_uid}）：你尚未关注该UP主")
            else:
                self.s_status_var.set("已关注 ✓")
                self.s_status_lbl.configure(fg="#00a1d6")
                time_str = timestamp_to_str(mtime)
                self.s_time_var.set(time_str)
                log(f"✅ {name}（{target_uid}）：关注于 {time_str}")
        except Exception as e:
            log(f"❌ 查询失败: {e}")
            self.s_name_var.set("查询失败")


if __name__ == "__main__":
    app = App()
    app.mainloop()
