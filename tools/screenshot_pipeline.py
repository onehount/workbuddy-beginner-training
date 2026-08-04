# -*- coding: utf-8 -*-
"""
截图批处理工具链（MIT）

把学员运行 WorkBuddy 后采集的原始截图加工为：
  - ANNOTATED：原图 + 红色定位框
  - COMPOSITE：原图 + 局部放大 + 科技蓝边框 + Step 编号 + 说明

用法：
  python screenshot_pipeline.py annotate  [--spec 07-screenshots/screenshot-spec.json]
  python screenshot_pipeline.py composite [--spec 07-screenshots/screenshot-spec.json]

依赖：Pillow（已随项目 venv 安装）

License: MIT
"""
import os
import sys
import json

from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
RAW_DIR = os.path.join(ROOT, "07-screenshots", "raw")
ANN_DIR = os.path.join(ROOT, "07-screenshots", "annotated")
CMP_DIR = os.path.join(ROOT, "07-screenshots", "composite")
FONT_PATH = "C:/Windows/Fonts/msyh.ttc"
FONT_BOLD = "C:/Windows/Fonts/msyhbd.ttc"

# 颜色
RED = (220, 53, 69)
NAVY = (27, 58, 107)
BLUE = (46, 107, 230)
WHITE = (255, 255, 255)


def font(size, bold=False):
    try:
        return ImageFont.truetype(FONT_BOLD if bold else FONT_PATH, size)
    except Exception:
        return ImageFont.load_default()


def load_spec(path=None):
    if path is None:
        path = os.path.join(ROOT, "07-screenshots", "screenshot-spec.json")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def annotate(spec):
    """给每张原始截图加红色定位框，输出到 annotated/"""
    os.makedirs(ANN_DIR, exist_ok=True)
    for item in spec:
        raw_path = os.path.join(RAW_DIR, item["raw"])
        if not os.path.exists(raw_path):
            print(f"[SKIP] {item['raw']} not found")
            continue
        img = Image.open(raw_path).convert("RGB")
        d = ImageDraw.Draw(img)
        x, y, w, h = item["box"]
        d.rectangle([x, y, x + w, y + h], outline=RED, width=3)
        # 标注编号和标题
        d.text((10, 10), f"{item['id']} · {item['title']}", font=font(20, True), fill=RED)
        out_name = item["raw"].replace(".png", "_annotated.png")
        img.save(os.path.join(ANN_DIR, out_name))
        print(f"[OK] annotated/{out_name}")
    print(f"\nDone: {len(spec)} images annotated → {os.path.relpath(ANN_DIR, ROOT)}")


def composite(spec):
    """生成合成图：原图 + 放大区域 + 边框 + 说明，用于独立展示"""
    os.makedirs(CMP_DIR, exist_ok=True)
    for item in spec:
        raw_path = os.path.join(RAW_DIR, item["raw"])
        if not os.path.exists(raw_path):
            print(f"[SKIP] {item['raw']} not found")
            continue
        img = Image.open(raw_path).convert("RGB")
        iw, ih = img.size
        # 目标尺寸：1920x1080 的展示画布
        cw, ch = 1920, 1080
        canvas = Image.new("RGB", (cw, ch), WHITE)
        d = ImageDraw.Draw(canvas)

        # 缩放原图至左侧区域
        scale = min((cw - 500) / iw, (ch - 120) / ih)
        dw, dh = int(iw * scale), int(ih * scale)
        ox, oy = 30, 60
        canvas.paste(img.resize((dw, dh)), (ox, oy))

        # 蓝色外框
        d.rectangle([ox - 4, oy - 4, ox + dw + 4, oy + dh + 4], outline=BLUE, width=3)

        # 右侧信息区
        rx = ox + dw + 40
        # Step 编号
        d.text((rx, 70), f"Step {item['id']}", font=font(36, True), fill=NAVY)
        # 标题
        d.text((rx, 130), item["title"], font=font(28, True), fill=(34, 42, 53))
        # 说明要点
        by = 200
        for note in item.get("notes", []):
            d.text((rx, by), f"• {note}", font=font(22), fill=(85, 85, 85))
            by += 40

        # 底部口径条
        d.rectangle([0, ch - 50, cw, ch], fill=(234, 241, 253))
        d.text((30, ch - 38),
               f"截图编号 {item['id']} | 原始文件: {item['raw']} | "
               f"本图由 screenshot_pipeline.py 自动生成",
               font=font(16), fill=(85, 85, 85))

        out_name = item["raw"].replace(".png", "_composite.png")
        canvas.save(os.path.join(CMP_DIR, out_name))
        print(f"[OK] composite/{out_name}")
    print(f"\nDone: {len(spec)} composites → {os.path.relpath(CMP_DIR, ROOT)}")


if __name__ == "__main__":
    C.verify()
    spec = load_spec()
    mode = sys.argv[1] if len(sys.argv) > 1 else "annotate"
    if mode == "annotate":
        annotate(spec)
    elif mode == "composite":
        composite(spec)
    else:
        print("Usage: python screenshot_pipeline.py [annotate|composite]")
