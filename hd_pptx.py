# -*- coding: utf-8 -*-
"""
HD Image PPT Builder
=====================
将任意 PPTX 的每一页渲染为高清图片，然后生成一个"纯图片+备注"的防错位 PPTX。
适用于会议放映，确保排版零偏移。

用法（命令行）：
    python hd_pptx.py <input.pptx> [output.pptx] [--dpi 300]

示例：
    python hd_pptx.py "我的报告.pptx"
    python hd_pptx.py "我的报告.pptx" "输出_高清.pptx" --dpi 3840

参数：
    input.pptx   源文件路径（必需）
    output.pptx  输出路径（可选，默认 = 源文件名_HD_image.pptx）
    --dpi N      渲染宽度像素（默认 5120，即 5K；可选 3840=4K, 2560=2K 等）

输出：
    <output.pptx>          每页一张高清整页图 + 逐字备注
    <input>_hd_png/        渲染出的高清 PNG 文件夹（备用）

限制：
    需要本机安装 Microsoft PowerPoint（COM 调用）
    不修改源文件
"""
import os, sys, glob, re, shutil, subprocess

# ── 配置 ──
POWERPOINT_NAMES = ["POWERPNT.EXE"]
TEMP_RENDER_DIR = "_hd_render_temp"

# ── 参数解析 ──
def parse_args():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        sys.exit(1)
    src = os.path.abspath(args[0])
    if not src.lower().endswith(".pptx"):
        print("错误：输入文件必须是 .pptx 格式")
        sys.exit(1)
    if not os.path.exists(src):
        print("错误：文件不存在:", src)
        sys.exit(1)
    out = None
    dpi = 5120
    i = 1
    while i < len(args):
        if args[i] == "--dpi" and i + 1 < len(args):
            dpi = int(args[i + 1])
            i += 2
        elif not out:
            out = os.path.abspath(args[i])
            i += 1
        else:
            i += 1
    if not out:
        base = os.path.splitext(src)[0]
        out = base + "_HD_image.pptx"
    return src, out, dpi


# ── Step 1: Render slides to high-res PNG via PowerPoint COM ──
def render_slides(src_pptx, render_dir, width_px, height_px):
    """Uses PowerPoint COM to export each slide as PNG at the target resolution."""
    ps_script = f'''
$ErrorActionPreference = "Stop"
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 3
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open("{src_pptx}", $true, $false, $false)
if (Test-Path "{render_dir}") {{ Remove-Item -Recurse -Force "{render_dir}" }}
New-Item -ItemType Directory -Force -Path "{render_dir}" | Out-Null
$pres.Export("{render_dir}", "PNG", {width_px}, {height_px})
Write-Host "SLIDES:" $pres.Slides.Count
$pres.Close()
$app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
'''
    script_path = os.path.join(os.path.dirname(render_dir), "_render_temp.ps1")
    with open(script_path, "w", encoding="utf-8-sig") as f:
        f.write(ps_script)

    print("正在渲染幻灯片（%d x %d px）..." % (width_px, height_px))
    result = os.system('powershell -ExecutionPolicy Bypass -File "%s"' % script_path)
    if result != 0:
        print("错误：PowerPoint COM 渲染失败（退出码 %d）" % result)
        print("请确认：1) 已安装 Microsoft PowerPoint；2) 没有其他 PowerPoint 实例占用文件")
        sys.exit(1)

    # collect PNGs
    png_files = {}
    for p in glob.glob(os.path.join(render_dir, "*.png")) + glob.glob(os.path.join(render_dir, "*.PNG")):
        base = os.path.basename(p)
        digits = "".join(c for c in base if c.isdigit())
        if digits:
            png_files[int(digits)] = p

    if not png_files:
        print("错误：渲染完成但没有找到 PNG 文件")
        sys.exit(1)

    print("渲染完成：%d 页" % len(png_files))
    return png_files


# ── Step 2: Build new PPTX with full-bleed images + verbatim notes ──
def build_hd_pptx(src_pptx, png_files, out_pptx):
    from pptx import Presentation
    from pptx.util import Emu

    src = Presentation(src_pptx)
    W = src.slide_width
    H = src.slide_height

    new = Presentation()
    new.slide_width = W
    new.slide_height = H
    blank = new.slide_layouts[6]

    for i, slide in enumerate(src.slides, 1):
        png = png_files.get(i)
        if png is None:
            print("警告：第 %d 页渲染文件缺失，跳过" % i)
            continue
        ns = new.slides.add_slide(blank)
        ns.shapes.add_picture(png, 0, 0, width=W, height=H)
        # copy notes verbatim
        if slide.has_notes_slide:
            nt = slide.notes_slide.notes_text_frame.text
            if nt.strip():
                ns.notes_slide.notes_text_frame.text = nt

    new.save(out_pptx)
    return len(new.slides.__iter__.__self__._sldIdLst) if hasattr(new.slides, '_sldIdLst') else len(list(new.slides))


# ── Main ──
def main():
    src, out, dpi = parse_args()

    # calculate height from slide aspect ratio
    from pptx import Presentation
    tmp = Presentation(src)
    w_emu = tmp.slide_width
    h_emu = tmp.slide_height
    aspect = h_emu / w_emu if w_emu else 0.5625
    height_px = int(dpi * aspect)
    del tmp

    # render dir: next to source file
    src_dir = os.path.dirname(src)
    src_name = os.path.splitext(os.path.basename(src))[0]
    render_dir = os.path.join(src_dir, src_name + "_hd_png")

    print("=" * 50)
    print("HD Image PPT Builder")
    print("=" * 50)
    print("源文件:  %s" % src)
    print("输出:    %s" % out)
    print("分辨率:  %d x %d px (%.1fK)" % (dpi, height_px, dpi / 1000.0))
    print("")

    # Step 1: render
    png_files = render_slides(src, render_dir, dpi, height_px)

    # Step 2: build
    count = build_hd_pptx(src, png_files, out)

    # done
    size_mb = round(os.path.getsize(out) / 1048576, 1)
    print("")
    print("=" * 50)
    print("完成！")
    print("输出文件: %s (%.1f MB)" % (out, size_mb))
    print("页数:     %d" % count)
    print("备注:     已逐字保留")
    print("=" * 50)


if __name__ == "__main__":
    main()
