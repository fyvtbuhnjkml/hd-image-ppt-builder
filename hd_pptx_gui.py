# -*- coding: utf-8 -*-
"""HD Image PPT Builder v3 - FINAL VERSION
Key fixes over v2:
1. DispatchEx (new PowerPoint instance) + retry loop for COM call rejection
2. Wait 5s after killing existing PowerPoint
3. Filter PNGs: keep only the largest per slide number (deduplicates thumbnails)
4. Show full error in scrollable text widget
5. Run COM in main thread via root.after() (no threading issues)
"""
import os, sys, glob, shutil, time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import pythoncom
import win32com.client


def kill_powerpoint():
    """Kill all PowerPoint instances and wait."""
    try:
        import psutil
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] and 'POWERPNT' in proc.info['name'].upper():
                proc.kill()
        time.sleep(5)
    except ImportError:
        os.system('taskkill /f /im POWERPNT.EXE >nul 2>&1')
        time.sleep(5)


def render_slides(src_pptx, render_dir, width_px, height_px, status_cb=None):
    """Render slides using PowerPoint COM (DispatchEx + retry)."""
    pythoncom.CoInitialize()
    pres = None
    app = None
    try:
        kill_powerpoint()

        if status_cb:
            status_cb("启动 PowerPoint...")

        # Use DispatchEx for a NEW instance
        max_retries = 3
        for attempt in range(max_retries):
            try:
                app = win32com.client.DispatchEx("PowerPoint.Application")
                break
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(3)
                else:
                    raise RuntimeError("无法启动 PowerPoint：%s" % str(e))

        if status_cb:
            status_cb("打开演示文稿...")

        # Open with retry — CRITICAL: normalize path to backslashes (COM requirement)
        src_pptx = os.path.normpath(src_pptx)
        for attempt in range(max_retries):
            try:
                pres = app.Presentations.Open(src_pptx, True, False, False)
                break
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(3)
                else:
                    raise RuntimeError("无法打开文件：%s" % str(e))

        slide_count = pres.Slides.Count

        # Clean render dir
        if os.path.exists(render_dir):
            shutil.rmtree(render_dir, ignore_errors=True)
        os.makedirs(render_dir, exist_ok=True)

        if status_cb:
            status_cb("正在渲染 %d 页（%d x %d px）..." % (slide_count, width_px, height_px))

        # Export
        pres.Export(render_dir, "PNG", width_px, height_px)

        # Collect PNGs - deduplicate by slide number, keep largest
        all_pngs = glob.glob(os.path.join(render_dir, "*.png")) + glob.glob(os.path.join(render_dir, "*.PNG"))
        png_by_num = {}  # slide_num -> (file_size, filepath)
        for p in all_pngs:
            base = os.path.basename(p)
            digits = "".join(c for c in base if c.isdigit())
            if not digits:
                continue
            num = int(digits)
            size = os.path.getsize(p)
            if num not in png_by_num or size > png_by_num[num][0]:
                png_by_num[num] = (size, p)

        # Extract clean mapping
        png_files = {num: path for num, (size, path) in png_by_num.items()}

        if not png_files:
            raise RuntimeError("渲染完成但未找到图片文件")

        pres.Close()
        app.Quit()

        return slide_count, png_files

    except Exception as e:
        if pres:
            try:
                pres.Close()
            except Exception:
                pass
        if app:
            try:
                app.Quit()
            except Exception:
                pass
        raise
    finally:
        pythoncom.CoUninitialize()


def build_hd_pptx(src_pptx, png_files, out_pptx, status_cb=None):
    """Build new PPTX with full-bleed images + verbatim notes."""
    from pptx import Presentation
    if status_cb:
        status_cb("生成防错位 PPTX...")
    src = Presentation(src_pptx)
    W, H = src.slide_width, src.slide_height
    new = Presentation()
    new.slide_width = W
    new.slide_height = H
    blank = new.slide_layouts[6]
    for i, slide in enumerate(src.slides, 1):
        png = png_files.get(i)
        if png is None:
            continue
        ns = new.slides.add_slide(blank)
        ns.shapes.add_picture(png, 0, 0, width=W, height=H)
        if slide.has_notes_slide:
            nt = slide.notes_slide.notes_text_frame.text
            if nt.strip():
                ns.notes_slide.notes_text_frame.text = nt
    new.save(out_pptx)
    return len(list(src.slides))


# ── GUI ──
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("HD Image PPT Builder v3")
        self.root.geometry("620x560")
        self.root.resizable(False, False)
        self.root.configure(bg="#f5f5f5")

        # title
        tk.Label(self.root, text="高清防错位 PPT 转换工具", font=("微软雅黑", 16, "bold"),
                 bg="#f5f5f5", fg="#2D2D8A").pack(pady=(15, 3))
        tk.Label(self.root, text="每页渲染为高清图片 → 防错位纯图版（保留演讲备注）",
                 font=("微软雅黑", 10), bg="#f5f5f5", fg="#666").pack(pady=(0, 10))

        # file
        f1 = tk.Frame(self.root, bg="#f5f5f5")
        f1.pack(pady=3, padx=25, fill="x")
        tk.Label(f1, text="选择 PPT：", font=("微软雅黑", 10), bg="#f5f5f5").pack(anchor="w")
        self.file_var = tk.StringVar()
        tk.Entry(f1, textvariable=self.file_var, width=58, font=("Consolas", 9)).pack(side="left", fill="x", expand=True)
        tk.Button(f1, text="浏览", command=self.browse_file, font=("微软雅黑", 9)).pack(side="right", padx=(5, 0))

        # output
        f2 = tk.Frame(self.root, bg="#f5f5f5")
        f2.pack(pady=3, padx=25, fill="x")
        tk.Label(f2, text="输出（留空自动命名）：", font=("微软雅黑", 10), bg="#f5f5f5").pack(anchor="w")
        self.out_var = tk.StringVar()
        tk.Entry(f2, textvariable=self.out_var, width=58, font=("Consolas", 9)).pack(side="left", fill="x", expand=True)

        # DPI
        f3 = tk.Frame(self.root, bg="#f5f5f5")
        f3.pack(pady=8, padx=25, anchor="w")
        tk.Label(f3, text="清晰度：", font=("微软雅黑", 10), bg="#f5f5f5").pack(side="left")
        self.dpi_var = tk.IntVar(value=5120)
        for val, lab in [(2560, "2K"), (3840, "4K（推荐）"), (5120, "5K（最清晰）")]:
            tk.Radiobutton(f3, text=lab, variable=self.dpi_var, value=val,
                           font=("微软雅黑", 10), bg="#f5f5f5").pack(side="left", padx=8)

        # run
        self.run_btn = tk.Button(self.root, text="▶  开始转换", font=("微软雅黑", 13, "bold"),
                                 bg="#2D2D8A", fg="white", width=22, height=2,
                                 command=self.run, cursor="hand2")
        self.run_btn.pack(pady=10)

        # status
        self.status_var = tk.StringVar(value="请选择文件...")
        tk.Label(self.root, textvariable=self.status_var, font=("微软雅黑", 10),
                 bg="#f5f5f5", fg="#2D2D8A").pack(pady=3)

        # progress
        self.progress = ttk.Progressbar(self.root, mode="indeterminate", length=500)
        self.progress.pack(pady=3)

        # log area (scrollable)
        log_frame = tk.Frame(self.root, bg="#f5f5f5")
        log_frame.pack(pady=5, padx=25, fill="both", expand=True)
        self.log_text = tk.Text(log_frame, height=8, font=("Consolas", 9), bg="#1a1a2e", fg="#00ff00",
                                state="disabled", wrap="word")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self.root.update_idletasks()

    def browse_file(self):
        f = filedialog.askopenfilename(title="选择 PPT", filetypes=[("PPT 文件", "*.pptx")])
        if f:
            self.file_var.set(f)
            base = os.path.splitext(f)[0]
            if not self.out_var.get():
                self.out_var.set(base + "_HD_image.pptx")

    def run(self):
        src = os.path.normpath(self.file_var.get().strip())
        if not src or not os.path.exists(src):
            messagebox.showerror("错误", "请先选择有效的 PPT 文件！")
            return
        out = os.path.normpath(self.out_var.get().strip() or os.path.splitext(src)[0] + "_HD_image.pptx")
        dpi = self.dpi_var.get()

        self.run_btn.config(state="disabled")
        self.progress.start(30)
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

        # Run in main thread via after()
        def do_work():
            try:
                out_path = os.path.normpath(out)
                self.log("=== HD Image PPT Builder v3 ===")
                self.log("源: %s" % src)
                self.log("输出: %s" % out_path)
                self.log("清晰度: %d px" % dpi)
                self.log("")

                from pptx import Presentation
                tmp = Presentation(src)
                h_ratio = tmp.slide_height / tmp.slide_width if tmp.slide_width else 0.5625
                h_px = int(dpi * h_ratio)
                del tmp

                src_dir = os.path.dirname(src)
                src_name = os.path.splitext(os.path.basename(src))[0]
                render_dir = os.path.normpath(os.path.join(src_dir, src_name + "_hd_png"))

                def status(msg):
                    self.set_status(msg)
                    self.log(msg)

                count, png_files = render_slides(src, render_dir, dpi, h_px, status)
                self.log("渲染完成: %d 页, %d 张图片" % (count, len(png_files)))

                n = build_hd_pptx(src, png_files, out_path, status)

                size_mb = round(os.path.getsize(out_path) / 1048576, 1)
                self.log("")
                self.log("=== 完成! ===")
                self.log("输出: %s (%.1f MB)" % (out_path, size_mb))
                self.log("页数: %d | 备注: 已保留" % n)
                self.set_status("✔ 转换成功！")

            except Exception as e:
                self.log("")
                self.log("=== 出错 ===")
                self.log(str(e))
                self.set_status("✖ 出错，详见下方日志")

            finally:
                self.progress.stop()
                self.run_btn.config(state="normal")

        self.root.after(200, do_work)

    def set_status(self, msg):
        self.status_var.set(msg)
        self.root.update_idletasks()

    def go(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = App()
    app.go()
