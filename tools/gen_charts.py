# -*- coding: utf-8 -*-
"""
生成课程 PPT 用的数据可视化图表（08-design-assets/charts/）

全部由 case_data.py 推导，与 demo 输入/输出数字严格一致。
使用 Pillow 绘制，微软雅黑字体，科技蓝主题。

License: MIT
"""
import os
import sys

from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import case_data as C

ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
OUT = os.path.join(ROOT, "08-design-assets", "charts")
os.makedirs(OUT, exist_ok=True)

FONT_PATH = "C:/Windows/Fonts/msyh.ttc"
FONT_BOLD = "C:/Windows/Fonts/msyhbd.ttc"

# 主题色
NAVY = (27, 58, 107)
BLUE = (46, 107, 230)
CYAN = (23, 192, 199)
LIGHT = (234, 241, 253)
GREY = (85, 85, 85)
RED = (192, 57, 43)
GREEN = (39, 174, 96)
WHITE = (255, 255, 255)
INK = (34, 42, 53)

W, H = 1280, 720


def font(size, bold=False):
    try:
        return ImageFont.truetype(FONT_BOLD if bold else FONT_PATH, size)
    except Exception:
        return ImageFont.load_default()


def new_canvas(title, subtitle=""):
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    # 顶部标题条
    d.rectangle([0, 0, W, 92], fill=NAVY)
    d.text((40, 22), title, font=font(30, True), fill=WHITE)
    if subtitle:
        d.text((40, 60), subtitle, font=font(15), fill=LIGHT)
    # 底部口径条
    d.rectangle([0, H - 40, W, H], fill=LIGHT)
    return img, d


def footer(d, text):
    d.text((40, H - 30), text, font=font(14), fill=GREY)


# ============================================================
# 图 1：各渠道曝光量（横向条形）
# ============================================================
def chart_exposure():
    img, d = new_canvas("各渠道曝光量对比",
                        "单位：次　｜　数据来自 05_渠道曝光与线索数据.xlsx（公式计算）")
    t = C.channel_totals()
    data = [(ch, t[ch][0]) for ch in C.CHANNELS]
    data.sort(key=lambda x: x[1], reverse=True)
    maxv = max(v for _, v in data)
    x0, y0 = 300, 150
    bar_h, gap = 56, 36
    for i, (ch, v) in enumerate(data):
        y = y0 + i * (bar_h + gap)
        bw = int((v / maxv) * (W - x0 - 160))
        # 条
        d.rectangle([x0, y, x0 + bw, y + bar_h], fill=BLUE)
        # 渠道名
        d.text((40, y + 12), ch, font=font(22, True), fill=INK)
        # 数值
        d.text((x0 + bw + 12, y + 12), f"{v:,}", font=font(22, True), fill=BLUE)
    # 强调：抖音曝光最高
    top_ch, top_v = data[0]
    d.text((40, H - 110), f"曝光最高：{top_ch}（{top_v:,} 次），但转化率最低（见下页）",
           font=font(16), fill=RED)
    footer(d, "统计周期：%s　｜　口径：各渠道后台导出，人工初筛" % C.PERIOD_RANGE)
    p = os.path.join(OUT, "chart_exposure.png")
    img.save(p)
    return p


# ============================================================
# 图 2：各渠道转化率 + 整体对比（纵条 + 参考线）
# ============================================================
def chart_conversion():
    img, d = new_canvas("各渠道转化率对比",
                        "单位：%　｜　危险点：整体转化率不能用算术平均替代")
    t = C.channel_totals()
    chs = C.CHANNELS
    vals = [t[ch][3] for ch in chs]
    grand = C.grand_total()[3]            # 37.0 加权
    avg = sum(vals) / len(vals)           # ~41.8 算术平均（错误）

    x0, y0, plot_w, plot_h = 120, 150, W - 240, 460
    ymax = 70.0
    # 网格
    for g in range(0, int(ymax) + 1, 10):
        gy = y0 + plot_h - int(g / ymax * plot_h)
        d.line([x0, gy, x0 + plot_w, gy], fill=(220, 220, 220), width=1)
        d.text((x0 - 38, gy - 8), f"{g}%", font=font(13), fill=GREY)
    # 纵条
    n = len(chs)
    slot = plot_w / n
    bw = slot * 0.55
    for i, (ch, v) in enumerate(zip(chs, vals)):
        cx = x0 + slot * i + slot / 2
        bh = int(v / ymax * plot_h)
        bx = cx - bw / 2
        by = y0 + plot_h - bh
        col = BLUE if v >= grand else (120, 160, 220)
        d.rectangle([bx, by, bx + bw, y0 + plot_h], fill=col)
        d.text((cx, by - 26), f"{v:.1f}%", font=font(18, True), fill=INK)
        d.text((cx, y0 + plot_h + 8), ch, font=font(15, True), fill=INK)
    # 参考线：加权平均（正确）
    gy_g = y0 + plot_h - int(grand / ymax * plot_h)
    d.line([x0, gy_g, x0 + plot_w, gy_g], fill=GREEN, width=3)
    d.text((x0 + plot_w - 250, gy_g - 26), f"整体（加权）{grand:.1f}%", font=font(16, True), fill=GREEN)
    # 参考线：算术平均（错误）
    gy_a = y0 + plot_h - int(avg / ymax * plot_h)
    d.line([x0, gy_a, x0 + plot_w, gy_a], fill=RED, width=3)
    d.text((x0 + plot_w - 270, gy_a + 6), f"算术平均（误）{avg:.1f}%", font=font(16, True), fill=RED)
    footer(d, "口径：转化率 = 有效线索 ÷ 线索 ×100%　整体=合计有效线索÷合计线索（加权）")
    p = os.path.join(OUT, "chart_conversion.png")
    img.save(p)
    return p


# ============================================================
# 图 3：市场活动完成情况（环形图）
# ============================================================
def chart_campaigns():
    img, d = new_canvas("本月市场活动完成情况",
                        "共 %d 项　｜　完成率 = 已完成 ÷ 总数" % C.CAMPAIGN_TOTAL)
    done = C.CAMPAIGN_DONE
    delayed = C.CAMPAIGN_DELAYED
    total = C.CAMPAIGN_TOTAL
    cx, cy, r = 360, 380, 180
    # 环
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=LIGHT)
    # 完成扇区（从 -90° 顺时针）
    import math
    start = -90
    end_done = start + 360 * done / total
    d.pieslice([cx - r, cy - r, cx + r, cy + r], start, end_done, fill=BLUE)
    d.ellipse([cx - r + 60, cy - r + 60, cx + r - 60, cy + r - 60], fill=WHITE)
    d.text((cx - 40, cy - 22), f"{done/total:.0%}", font=font(40, True), fill=NAVY)
    d.text((cx - 60, cy + 24), "完成率", font=font(18), fill=GREY)
    # 图例
    lx = 640
    ly = 300
    d.rectangle([lx, ly, lx + 26, ly + 26], fill=BLUE)
    d.text((lx + 38, ly + 2), f"已完成：{done} 项", font=font(22, True), fill=INK)
    d.rectangle([lx, ly + 60, lx + 26, ly + 86], fill=RED)
    d.text((lx + 38, ly + 62), f"延期：{delayed} 项（校园联名，待 Logo 规范确认）",
           font=font(22, True), fill=INK)
    d.text((lx + 38, ly + 110), f"完成率 = {done} ÷ {total} = {done/total:.0%}（仅计已完成）",
           font=font(18), fill=GREY)
    footer(d, "延期不计入完成率，是会议纪要中须单独标记的待确认事项")
    p = os.path.join(OUT, "chart_campaigns.png")
    img.save(p)
    return p


if __name__ == "__main__":
    C.verify()
    ps = [chart_exposure(), chart_conversion(), chart_campaigns()]
    for p in ps:
        print("[OK] %s" % os.path.relpath(p, ROOT))
