# -*- coding: utf-8 -*-
"""
生成讲师材料（03-instructor-guide/）

- instructor-manual.md    讲师使用手册
- full-speaking-script.md 逐页讲稿（含每页时长，合计 60 分钟）
- timing-guide.md         时间分配表
- classroom-faq.md        课堂常见问题

License: MIT
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gen_course_ppt as G

ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
OUT = os.path.join(ROOT, "03-instructor-guide")

# 每页时长（分钟），8 章合计 60 分钟
TIMING = {
    1: 0.5, 2: 0.3, 3: 1.2, 4: 1.5, 5: 1.5,            # 第一章 5
    6: 0.5, 7: 1.6, 8: 1.6, 9: 1.6, 10: 1.7,           # 第二章 7
    11: 0.5, 12: 1.7, 13: 1.7, 14: 1.7, 15: 1.7, 16: 1.7,  # 第三章 9
    17: 0.5, 18: 1.5, 19: 1.5, 20: 1.5,                # 第四章 5
    21: 0.5, 22: 1.7, 23: 1.7, 24: 1.7, 25: 1.7,
    26: 1.7, 27: 1.7, 28: 1.7, 29: 1.7,                # 第五章 14
    30: 0.5, 31: 1.4, 32: 1.4, 33: 1.4, 34: 1.4, 35: 1.4,  # 第六章 9
    36: 0.5, 37: 1.4, 38: 1.4, 39: 1.4, 40: 1.3,       # 第七章 6
    41: 0.4, 42: 1.5, 43: 1.5, 44: 1.1, 45: 0.5,       # 第八章 5
}

CHAPTER_TIME = {
    "第一章 为什么学习 WorkBuddy": 5, "第二章 认识 WorkBuddy": 7,
    "第三章 核心方法": 9, "第四章 统一案例介绍": 5,
    "第五章 会议纪要重点 Demo": 14, "第六章 其他办公成果": 9,
    "第七章 综合 PPT 案例": 6, "第八章 检查、复现与寄语": 5,
}


def gen_script():
    lines = ["# 逐页讲稿（full-speaking-script.md）", "",
             "> 总时长 60 分钟。每页标注建议时长；讲解时以「核心信息」为主，"
             "不必逐字照读。截图页请切换到对应标注截图（见 screenshot-index.md）。", ""]
    for i, s in enumerate(G.SLIDES, start=1):
        t = s["type"]
        mins = TIMING.get(i, 1.0)
        lines.append(f"## 第 {i} 页 · {s['title']}　（{mins} 分钟）")
        lines.append("")
        # 过渡开场
        if t == "divider":
            lines.append(f"【章节页】进入{s['title']}。先点明本章要解决的问题，再翻页。")
        elif t == "cover":
            lines.append("【开场】欢迎大家。今天用 60 分钟，带大家用 WorkBuddy 完整走一遍"
                         "真实办公任务：从一句模糊需求，到一份正式汇报。")
        elif t == "closing":
            lines.append("【结语】放慢语速。告诉新人：工具会迭代，但「看清目标、检查结果」的能力长期有用。")
        else:
            lines.append(f"【讲解】{s.get('note','')}")
        # 内容要点
        if s.get("bullets"):
            lines.append("")
            lines.append("要点：")
            for b in s["bullets"]:
                lines.append(f"- {b}")
        if s.get("shot"):
            lines.append("")
            lines.append(f"【演示】展示截图 {s['shot']}（RAW → ANNOTATED → COMPOSITE）。"
                         + (f" 说明：{s['side']}" if s.get('side') else ""))
        lines.append("")
    with open(os.path.join(OUT, "full-speaking-script.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return os.path.join(OUT, "full-speaking-script.md")


def gen_timing():
    lines = ["# 时间分配表（timing-guide.md）", "",
             "总时长 **60 分钟**，按 8 章拆分。每章留有缓冲，若超时优先压缩第六章演示。", ""]
    lines.append("| 章 | 主题 | 页数 | 建议时长 |")
    lines.append("|----|------|------|----------|")
    ch_pages = {}
    for i, s in enumerate(G.SLIDES, start=1):
        ch = s.get("chapter") or ("封面/寄语" if s["type"] in ("cover", "closing") else s["title"])
        ch_pages.setdefault(ch, []).append(i)
    for ch, pg in ch_pages.items():
        dur = sum(TIMING.get(p, 0) for p in pg)
        lines.append(f"| — | {ch} | {len(pg)} | {dur:.1f} 分钟 |")
    lines.append("")
    lines.append("| 合计 | 8 章 | 45 | 60.0 分钟 |")
    lines.append("")
    lines.append("### 逐页时长")
    lines.append("")
    lines.append("| 页 | 标题 | 时长 |")
    lines.append("|----|------|------|")
    for i, s in enumerate(G.SLIDES, start=1):
        lines.append(f"| {i} | {s['title']} | {TIMING.get(i,0):.1f} 分钟 |")
    with open(os.path.join(OUT, "timing-guide.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return os.path.join(OUT, "timing-guide.md")


def gen_manual():
    txt = """# 讲师使用手册（instructor-manual.md）

## 一、课前准备
- 本课程的 PPT、PDF、截图占位、Demo 输入/输出、提示词均已随仓库提供，**无需现场安装任何软件**。
- 将 `02-slides/workbuddy-beginner-training.pdf` 用于投影或发放；`pptx` 用于二次编辑。
- 演示机提前打开 `07-screenshots/composite/` 中的成品图（如已采集），按 `screenshot-index.md` 编号对照。
- 若截图尚未采集，PPT 中的占位框已标明编号，可口头说明「此处展示 WorkBuddy 操作截图」。

## 二、讲授原则
- **不讲代码、不现场操作**：按方案设计，所有 Demo 提前完成并截图；课堂只讲流程与方法。
- **每页一个知识点**：遇到学员追问，记到「待课后解答」，避免拖堂。
- **反复强化三句话**：本质、效率、未来；以及三个核心动作（先提问、传材料、再检查）。
- **不展示失败过程**：只展示验证过的成功流程（方案已确认「不展示失败过程」）。

## 三、课堂节奏
- 严格按 `timing-guide.md` 控制；第五章（会议纪要）最重，占 14 分钟。
- 总缓冲极小，若某章超时，优先压缩第六章「其他办公成果」的演示，快速带过。

## 四、配套文件
- `full-speaking-script.md`：逐页讲稿与每页时长。
- `slide-notes.md`：每页一句讲解要点（速查）。
- `slide-outline.md`：页面索引。
- `classroom-faq.md`：学员常见问题与标准回答。

## 五、常见风险
- 学员问「能不能现场试一下」→ 说明课后可在 GitHub 下载材料自行复现。
- 学员担心隐私 → 强调所有案例数据为虚构，并指向 `SECURITY.md` 与 `11-quality-assurance/privacy-review.md`。
- 时间不够 → 跳过第六章部分演示页，保留第五章与第七章。
"""
    p = os.path.join(OUT, "instructor-manual.md")
    with open(p, "w", encoding="utf-8") as f:
        f.write(txt)
    return p


def gen_faq():
    txt = """# 课堂常见问题（classroom-faq.md）

**Q1：WorkBuddy 会把我们的数据泄露吗？**
A：本教学案例全部为虚构数据，不涉及真实信息。实际使用中，应遵循公司数据安全规范，不上传账号密码、客户隐私等敏感内容；详见 `SECURITY.md`。

**Q2：提示词需要背下来吗？**
A：不需要。仓库 `06-prompts/` 提供 8 份可直接复制的提示词。理解「先提问、再出方案、最后生成」的结构比背文字更重要。

**Q3：WorkBuddy 生成的内容出错怎么办？**
A：必须人工检查。重点核对数字、人名、日期、决定，并把「待确认」与「已决定」分开。本课第八章有检查清单。

**Q4：完全没有基础能学会吗？**
A：可以。本课程面向零基础新员工，不要求编程。只要能描述任务目标、上传文件，就能使用。

**Q5：它和 Excel / Word 有什么不同？**
A：办公软件是工具，要你一步步操作；WorkBuddy 能理解任务、跨文件整合、产出初稿，你负责判断与定稿。

**Q6：模型能力弱时生成质量差怎么办？**
A：把复杂任务拆小（本课采用分步骤策略）：一次只做一个文件、一类输出。先出方案，再逐份生成。

**Q7：生成的 PPT / 报告能直接发给领导吗？**
A：先人工复核数据与表述，再对外。WorkBuddy 产出的是「初稿」，最终责任在人。

**Q8：课后怎么练习？**
A：打开 GitHub 仓库 `workbuddy-beginner-training`，按 README 的复现路径，用 `04-demo-inputs` 与 `06-prompts` 自行走一遍。
"""
    p = os.path.join(OUT, "classroom-faq.md")
    with open(p, "w", encoding="utf-8") as f:
        f.write(txt)
    return p


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    for fn in (gen_script, gen_timing, gen_manual, gen_faq):
        p = fn()
        print(f"[OK] {os.path.relpath(p, ROOT)}")
    print("\n讲师材料生成完成。")
