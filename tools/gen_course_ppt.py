# -*- coding: utf-8 -*-
"""
生成课程 PPT 与 PDF（02-slides/）

- workbuddy-beginner-training.pptx  （科技风，16:9，约 38–40 页）
- workbuddy-beginner-training.pdf   （与 PPT 内容一致，可打印）
- slide-outline.md                 （页面索引）
- slide-notes.md                   （讲师备注）

本版本不含截图占位。Demo 章节以「处理流程 + 可视化结果 + 错误纠正」呈现。
图表来自 tools/gen_charts.py（Pillow 绘制，与 case_data.py 数字一致）。

License: MIT
"""
import os
import sys

from pptx import Presentation
from pptx.util import Inches, Pt as PptPt
from pptx.dml.color import RGBColor as PptRGB
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas as RLCanvas

ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
SLIDE_DIR = os.path.join(ROOT, "02-slides")
CHART_DIR = os.path.join(ROOT, "08-design-assets", "charts")

# ── 主题色 ──
NAVY = PptRGB(0x1B, 0x3A, 0x6B)
BLUE = PptRGB(0x2E, 0x6B, 0xE6)
CYAN = PptRGB(0x17, 0xC0, 0xC7)
LIGHT = PptRGB(0xEA, 0xF1, 0xFD)
WHITE = PptRGB(0xFF, 0xFF, 0xFF)
GREY = PptRGB(0x55, 0x55, 0x55)
RED = PptRGB(0xC0, 0x39, 0x2B)
GREEN = PptRGB(0x27, 0xAE, 0x60)
INK = PptRGB(0x22, 0x2A, 0x35)

NAVY_H = (0x1B, 0x3A, 0x6B)
BLUE_H = (0x2E, 0x6B, 0xE6)
LIGHT_H = (0xEA, 0xF1, 0xFD)
GREY_H = (0x55, 0x55, 0x55)
RED_H = (0xC0, 0x39, 0x2B)
GREEN_H = (0x27, 0xAE, 0x60)


# ============================================================
# 幻灯片内容（单一来源，无截图占位）
# ============================================================
SLIDES = [
    # ═══════════ 封面 ═══════════
    {"type": "cover", "kicker": "WorkBuddy 零基础办公任务实战教程",
     "title": "从需求澄清到综合汇报",
     "bullets": ["面向第一次接触 WorkBuddy 的新员工",
                 "60 分钟 · 不讲代码，只讲如何把工作交给 AI 完成",
                 "统一案例：青屿茶研市场部 7 月月度复盘"],
     "note": "欢迎开场，说明本课程目标与形式：用统一案例讲清完整任务闭环。"},

    # ═══════════ 第一章 为什么学习 WorkBuddy ═══════════
    {"type": "divider", "kicker": "第一章", "title": "为什么学习 WorkBuddy",
     "note": "先讲动机，再讲方法。"},
    {"type": "content", "chapter": "第一章 为什么学习 WorkBuddy",
     "kicker": "学习价值", "title": "本次课程能带来什么",
     "bullets": ["学会把模糊的工作目标，变成可执行的任务",
                 "掌握一套可复用的提示词，覆盖纪要、总结、通知、分析、汇报",
                 "拿到完整案例与公开仓库，课后能自行复现"],
     "note": "强调「可复现」是本课程的核心交付。"},
    {"type": "content", "chapter": "第一章 为什么学习 WorkBuddy",
     "kicker": "学习价值", "title": "新员工常见办公困境",
     "bullets": ["任务说不清：知道要结果，却写不出需求",
                 "材料散乱：会议记录、数据表、个人记录各在一处",
                 "反复排版：同一内容要改成纪要、总结、PPT",
                 "不敢交付：担心 AI 编数字、编人名、编决定"],
     "note": "这些困境正是 WorkBuddy 要解决的。"},
    {"type": "content", "chapter": "第一章 为什么学习 WorkBuddy",
     "kicker": "核心信息", "title": "本质、效率与未来",
     "bullets": ["本质：WorkBuddy 不是替你思考，而是把目标、材料、约束变成可执行任务",
                 "效率：减少澄清、整理、排版、转换中的机械劳动",
                 "未来：能定义目标、组织材料、检查结果，并与 AI 协同"],
     "note": "三句话分别对应开场、中段、寄语页，须反复强化。"},

    # ═══════════ 第二章 认识 WorkBuddy ═══════════
    {"type": "divider", "kicker": "第二章", "title": "认识 WorkBuddy",
     "note": "建立对工具角色的基本认知。"},
    {"type": "content", "chapter": "第二章 认识 WorkBuddy",
     "kicker": "角色", "title": "WorkBuddy 在办公任务中的角色",
     "bullets": ["一个能读文件、写文件、按步骤执行任务的协作助手",
                 "你给目标与材料，它产出初稿与成果",
                 "你保留判断、责任与最终决定权"],
     "note": "强调人是主导，WB 是协作者。"},
    {"type": "content", "chapter": "第二章 认识 WorkBuddy",
     "kicker": "区别", "title": "与普通问答工具的区别",
     "bullets": ["问答工具：一问一答，不记上下文、不碰你的文件",
                 "WorkBuddy：能上传附件、跨步骤推进、产出可下载文件",
                 "关键差异：任务连续性 + 文件处理能力"],
     "note": "用「能不能处理我的文件」来区分。"},
    {"type": "content", "chapter": "第二章 认识 WorkBuddy",
     "kicker": "分工", "title": "人、WorkBuddy 和 Agent 的分工",
     "bullets": ["人：定目标、供材料、做判断、负最终责任",
                 "WorkBuddy：拆任务、整材料、出初稿、做格式",
                 "Agent：在授权范围内自动跑子任务、批量处理"],
     "note": "职责边界要讲清，避免「全交给 AI」的误解。"},
    {"type": "content", "chapter": "第二章 认识 WorkBuddy",
     "kicker": "边界", "title": "该交 / 不该完全交",
     "bullets": ["适合交：会议纪要、工作总结、通知、数据分析、PPT 初稿",
                 "别全交：涉及真实隐私的信息",
                 "别全交：需要你拍板的决策",
                 "别全交：对外正式承诺"],
     "note": "落到「哪些信息不能上传」。"},

    # ═══════════ 第三章 核心方法 ═══════════
    {"type": "divider", "kicker": "第三章", "title": "核心方法",
     "note": "本课程最关键的一章。"},
    {"type": "content", "chapter": "第三章 核心方法",
     "kicker": "闭环", "title": "完整任务闭环",
     "bullets": ["提出任务 → 让 WB 提问 → 回答关键问题",
                 "获得方案 → 上传附件 → 分步骤执行",
                 "查看结果 → 提出修改 → 人工检查 → 导出文件"],
     "note": "这张图是后面所有 Demo 的骨架，建议板书。"},
    {"type": "content", "chapter": "第三章 核心方法",
     "kicker": "为什么", "title": "为什么不能一开始就要求生成最终文件",
     "bullets": ["目标不清，生成物必然跑偏",
                 "材料未给，AI 容易用「合理的虚构」填补空白",
                 "先出方案，能把你的隐性要求显性化"],
     "note": "用例：直接要「写份纪要」vs 先提问。"},
    {"type": "content", "chapter": "第三章 核心方法",
     "kicker": "先提问", "title": "先让 WorkBuddy 提问",
     "bullets": ["写不清需求时，先说：请先提问，帮我明确任务",
                 "WB 会问背景、目标、对象、材料、格式、时限、限制、口径、确认项、验收",
                 "信息足够后，它先给方案，不急着出文件"],
     "note": "对应提示词 06-prompts/01-interview-first.md。"},
    {"type": "content", "chapter": "第三章 核心方法",
     "kicker": "提示词", "title": "提问式提示词全貌",
     "bullets": ["一句话：我需要完成一次复盘，请先以提问方式帮我明确 10 个问题",
                 "每次提适量问题，不一次轰炸",
                 "方案含：任务顺序、输入、每步输出、确认节点、交付物、风险"],
     "note": "完整提示词在仓库 06-prompts 中可直接复制。"},
    {"type": "content", "chapter": "第三章 核心方法",
     "kicker": "方案", "title": "WB 返回的问题 → 形成可执行方案",
     "bullets": ["WB 返回的问题覆盖背景、对象、材料、文件、时限、口径等",
                 "你回答后，它给出分步方案",
                 "确认方案之前，不要让它开始生成终稿"],
     "note": "方案确认是人工控制点。"},

    # ═══════════ 第四章 统一案例介绍 ═══════════
    {"type": "divider", "kicker": "第四章", "title": "统一案例介绍",
     "note": "进入具体案例，后面 Demo 都围绕它。"},
    {"type": "content", "chapter": "第四章 统一案例介绍",
     "kicker": "背景", "title": "案例背景：青屿茶研",
     "bullets": ["虚构新式茶饮品牌，市场部做 7 月月度复盘",
                 "需要把会议记录、渠道数据、个人记录整理成正式成果",
                 "全部人物、数据均为虚构，仅用于教学"],
     "note": "务必提醒学员这是虚构案例。"},
    {"type": "content", "chapter": "第四章 统一案例介绍",
     "kicker": "活动", "title": "本月市场活动（5 项）",
     "bullets": ["夏日新品试饮 · 小红书种草 · 微信社群优惠 · 商场快闪：已完成",
                 "校园联名：因合作方 Logo 规范未确认，延期",
                 "完成率 = 4 ÷ 5 = 80%（仅计已完成）"],
     "note": "延期不计入完成率，这是后面纪要的要点。"},
    {"type": "content", "chapter": "第四章 统一案例介绍",
     "kicker": "总览", "title": "输入与产出总览",
     "bullets": ["输入：案例背景、会议记录、参会表、活动表、渠道数据、个人记录、4 份模板",
                 "产出：会议纪要、工作总结、活动通知、分析报告、月度复盘 PPT",
                 "所有数字由表格公式计算，非文字写入"],
     "note": "点出「数字必须可溯源」。"},

    # ═══════════ 第五章 会议纪要重点 Demo ═══════════
    {"type": "divider", "kicker": "第五章", "title": "会议纪要重点 Demo",
     "note": "本章展示处理流程、可复用结构、以及最易犯的错误。"},
    {"type": "content", "chapter": "第五章 会议纪要重点 Demo",
     "kicker": "流程", "title": "会议纪要处理流程（5 步）",
     "bullets": ["① 创建任务并命名（如「7 月市场复盘会议纪要」），给出清晰边界",
                 "② 上传原始记录 + 模板（模板决定结构，原始记录决定内容）",
                 "③ 输入约束：只使用原始材料，不得虚构发言 / 数字 / 决定",
                 "④ 先输出纪要结构 + 待确认事项列表，经确认后再生成正式文档",
                 "⑤ 人工核对数字、人名、日期；与数据表交叉验证关键指标"],
     "note": "强调第③④步是控制点，不是一次性生成。"},
    {"type": "content", "chapter": "第五章 会议纪要重点 Demo",
     "kicker": "结构", "title": "会议纪要的可复用结构（6 部分）",
     "bullets": ["一、会议基本信息（时间 / 地点 / 主持人 / 参会人员）",
                 "二、讨论内容（逐项列出，标注事实 vs 讨论）",
                 "三、已确认决定（逐条编号，有出处）",
                 "四、行动事项（序号 / 事项 / 负责人 / 截止 / 备注）",
                 "五、待确认事项（单独标记，不写入决定）",
                 "六、下次安排"],
     "note": "这个结构适用于大多数内部会议纪要。"},
    {"type": "compare", "chapter": "第五章 会议纪要重点 Demo",
     "kicker": "纠错", "title": "常见错误①：「待确认」被写成「决定」",
     "left_label": "❌ 错误做法", "left_items": [
         "原始记录中写的是「快闪预算追加金额待定」",
         "但生成的纪要里写成了「决定追加预算 XX 元」",
         "原因：AI 把模糊表述补成了确定结论",
         "后果：未经授权的承诺可能引发责任纠纷",
     ], "right_label": "✅ 正确做法", "right_items": [
         "将该项归入「五、待确认事项」",
         "引用原话：「快闪预算追加金额待定」",
         "备注：缺具体金额，需后续专项确认",
         "原则：源材料未明确的，一律标记为待确认",
     ],
     "note": "这是会议纪要中最常见的 AI 幻觉类型。"},

    # ═══════════ 第六章 其他办公成果 ═══════════
    {"type": "divider", "kicker": "第六章", "title": "其他办公成果",
     "note": "快速带过，重点放在数据分析的可视化结果和纠错上。"},
    {"type": "chart", "chapter": "第六章 其他办公成果",
     "kicker": "曝光", "title": "数据分析可视化：各渠道曝光量",
     "image": "chart_exposure.png",
     "caption": "抖音曝光量最高（262,000 次），但转化率最低——高流量 ≠ 高质量。",
     "note": "引导学员看图说话，不要念数字。"},
    {"type": "chart", "chapter": "第六章 其他办公成果",
     "kicker": "转化率", "title": "数据分析可视化：转化率对比",
     "image": "chart_conversion.png",
     "caption": "绿线 = 整体加权（正确 37.0%）；红线 = 算术平均（错误 ~41.8%）。",
     "bullets": ["微信社群转化率最高 60.0%，客群精准",
                 "抖音转化率最低 25.0%，需优化内容而非加预算",
                 "整体转化率受抖音大基数拉低，不可用算术平均替代"],
     "note": "这张图是本课程最重要的「纠错」可视化。"},
    {"type": "compare", "chapter": "第六章 其他办公成果",
     "kicker": "纠错", "title": "常见错误②：整体转化率算成算术平均",
     "left_label": "❌ 错误做法", "left_items": [
         "五个渠道转化率求平均：(34.1+25.0+45.0+60.0+45.0) ÷ 5 = 41.8%",
         "报告结论：「整体转化率为 41.8%」",
         "问题：忽略了各渠道线索量差异巨大的事实",
         "后果：决策依据失真，可能误导资源分配",
     ], "right_label": "✅ 正确做法", "right_items": [
         "整体转化率 = 总有效线索 ÷ 总线索 = 1,744 ÷ 4,710 = 37.0%",
         "报告结论：「整体转化率为 37.0%（加权）」",
         "说明：因抖音低转化高曝光拉低加权结果",
         "原则：任何「合计」类指标都要检查是否需要加权",
     ],
     "note": "这是数据分析中最隐蔽也最常见的错误。"},
    {"type": "content", "chapter": "第六章 其他办公成果",
     "kicker": "其他", "title": "其他成果要点一览",
     "bullets": ["工作总结：区分「完成 / 参与 / 协助」，校园联名标为进行中（延期）",
                 "活动通知：仅含已确认事项，不含待确认项",
                 "分析报告：区分「事实 / 判断 / 建议」，不虚构预算或销售额",
                 "以上三份均可在仓库 05-demo-outputs 中查看原始版与审核版对比"],
     "note": "提示学员课后打开 repo 对比 original 与 reviewed-final。"},

    # ═══════════ 第七章 综合 PPT 案例 ═══════════
    {"type": "divider", "kicker": "第七章", "title": "综合 PPT 案例",
     "note": "把多份材料汇成一份汇报。"},
    {"type": "content", "chapter": "第七章 综合 PPT 案例",
     "kicker": "流程", "title": "从多份材料到统一汇报",
     "bullets": ["输入：纪要 + 总结 + 通知 + 分析报告（4 份已完成的文档）",
                 "目标：汇成一份 10–12 页的月度复盘 PPT",
                 "步骤：先让 WB 给逐页大纲 → 你确认/修改 → 再生成",
                 "原则：复用的不是某一篇文档，而是这套方法"],
     "note": "升华到方法论层面。"},
    {"type": "content", "chapter": "第七章 综合 PPT 案例",
     "kicker": "结构", "title": "综合 PPT 的标准结构（12 页）",
     "bullets": ["封面 → 目录 → 案例背景 → 本月活动（表格）→ 渠道概览",
                 "→ 各渠道转化率（图表）→ 已确认决定 → 行动事项（表格）",
                 "→ 待确认事项 → 关键结论 → 下月计划 → 寄语",
                 "每页一个核心结论；数据页标统计周期与口径"],
     "note": "成品见 05-demo-outputs/reviewed-final/07_月度复盘PPT_人工审核版.pptx。"},

    # ═══════════ 第八章 检查、复现与寄语 ═══════════
    {"type": "divider", "kicker": "第八章", "title": "检查、复现与寄语",
     "note": "收尾，落到行动。"},
    {"type": "content", "chapter": "第八章 检查、复现与寄语",
     "kicker": "四查", "title": "人工检查清单（四查）",
     "bullets": ["① 数字：每个值能在源表中找到，合计 = 各项相加",
                 "② 日期：年份/月份对，截止日是未来时间，逻辑不自相矛盾",
                 "③ 人名：与参会人员表一致，负责人没搞反，动词用对",
                 "④ 结论：每条有原文出处，「建议」≠「决定」，「待确认」≠ 确定"],
     "note": "详见 06-prompts/08-review-checklist.md。"},
    {"type": "flow", "chapter": "第八章 检查、复现与寄语",
     "kicker": "流程", "title": "错误纠正工作流",
     "steps": ["WorkBuddy\n原始输出", "四查自检\n(数字/日期/\n人名/结论)", "发现错误\n(定位行号)", "纠正话术\n(精确指出+\n统一修改)", "人工终稿\n(通读+导出)"],
     "note": "纠正话术模板见 08-review-checklist 第三节。"},
    {"type": "content", "chapter": "第八章 检查、复现与寄语",
     "kicker": "复现", "title": "GitHub 资料使用方法",
     "bullets": ["打开仓库 workbuddy-beginner-training",
                 "README 给出三条复现路径（按角色选择）",
                 "下载 inputs，按 prompts 自行复现",
                 "对比 05-demo-outputs 中 original 与 reviewed-final 的差异"],
     "note": "演示仓库结构即可。"},
    {"type": "content", "chapter": "第八章 检查、复现与寄语",
     "kicker": "动作", "title": "三个核心动作",
     "bullets": ["先提问，再生成方案",
                 "上传材料，明确限制",
                 "人工检查，再交付"],
     "note": "让学员记住这三句话。"},

    # ═══════════ 结语 ═══════════
    {"type": "closing", "kicker": "寄语", "title": "写给刚起步的你",
     "bullets": ["初入职场，不必一次写出完美指令",
                 "理解工作本质，知道目标与判断标准",
                 "保持好奇，持续学习，与 AI 协同成长"],
     "note": "放慢语速，给学员留出思考空间。"},
]


# ============================================================
# PPTX 渲染
# ============================================================
def render_pptx(path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    total = len(SLIDES)

    def footer(slide, idx, chapter):
        box = slide.shapes.add_textbox(Inches(0.4), Inches(7.05), Inches(9), Inches(0.35))
        tf = box.text_frame; p = tf.paragraphs[0]; p.text = chapter or ""
        p.font.size = PptPt(11); p.font.color.rgb = GREY; p.font.name = "微软雅黑"
        box2 = slide.shapes.add_textbox(Inches(12.3), Inches(7.05), Inches(0.8), Inches(0.35))
        tf2 = box2.text_frame; p2 = tf2.paragraphs[0]; p2.text = f"{idx}/{total}"
        p2.font.size = PptPt(11); p2.font.color.rgb = GREY; p2.font.name = "微软雅黑"
        p2.alignment = PP_ALIGN.RIGHT

    def bg(slide, color):
        box = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        box.fill.solid(); box.fill.fore_color.rgb = color
        box.line.fill.background(); box.shadow.inherit = False
        return box

    def top_bar(slide, title, kicker):
        bar = slide.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(1.15))
        bar.fill.solid(); bar.fill.fore_color.rgb = NAVY; bar.line.fill.background()
        tf = bar.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.5)
        p = tf.paragraphs[0]; p.text = title
        p.font.size = PptPt(30); p.font.bold = True; p.font.color.rgb = WHITE
        p.font.name = "微软雅黑"
        sp = tf.add_paragraph(); sp.text = kicker
        sp.font.size = PptPt(13); sp.font.color.rgb = LIGHT; sp.font.name = "微软雅黑"

    for idx, s in enumerate(SLIDES, start=1):
        slide = prs.slides.add_slide(blank)
        t = s["type"]

        if t == "cover":
            bg(slide, NAVY)
            add_text(slide, 1.0, 2.2, 11.3, 0.6, s["kicker"], 22, LIGHT)
            add_text(slide, 1.0, 2.8, 11.3, 1.2, s["title"], 44, WHITE, bold=True)
            add_text(slide, 1.0, 4.2, 11.3, 2.0,
                     "\n".join("· " + b for b in s["bullets"]), 20, LIGHT)
            footer(slide, idx, "WorkBuddy 零基础办公任务实战教程")

        elif t == "divider":
            bg(slide, NAVY)
            add_text(slide, 1.0, 2.6, 11.3, 0.8, s["kicker"], 26, CYAN, bold=True)
            add_text(slide, 1.0, 3.4, 11.3, 1.4, s["title"], 46, WHITE, bold=True)
            footer(slide, idx, s["title"])

        elif t == "content":
            bg(slide, WHITE)
            top_bar(slide, s["title"], s["kicker"])
            add_bullets(slide, 0.8, 1.6, 11.7, 5.2, s["bullets"], size=22)
            footer(slide, idx, s.get("chapter", ""))

        elif t == "chart":
            bg(slide, WHITE)
            top_bar(slide, s["title"], s["kicker"])
            img_path = os.path.join(CHART_DIR, s["image"])
            if os.path.exists(img_path):
                pic = slide.shapes.add_picture(
                    img_path, Inches(0.6), Inches(1.45),
                    width=Inches(9.2))
                # 图片下方说明
                cap_y = 6.0
                add_text(slide, 0.8, cap_y, 9.0, 0.5, s.get("caption", ""), 14, BLUE)
            else:
                add_text(slide, 0.8, 2.0, 10, 3, f"[图表缺失: {s['image']}]", 18, RED)
            # 右侧要点（如果有）
            if s.get("bullets"):
                add_bullets(slide, 10.0, 1.6, 3.0, 4.8, s["bullets"], size=16)
            footer(slide, idx, s.get("chapter", ""))

        elif t == "compare":
            bg(slide, WHITE)
            top_bar(slide, s["title"], s["kicker"])
            # 左列（红色标题 + 错误做法）
            lh = min(len(s["left_items"]) * 0.65 + 0.8, 4.8)
            left_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                              Inches(0.6), Inches(1.5), Inches(5.9), Inches(lh))
            left_box.fill.solid(); left_box.fill.fore_color.rgb = PptRGB(0xFD, 0xED, 0xEC)
            left_box.line.color.rgb = RED; left_box.line.width = PptPt(1.5)
            add_text(slide, 0.85, 1.58, 5.5, 0.45, s["left_label"], 18, RED, bold=True)
            add_bullets(slide, 0.85, 2.1, 5.5, lh - 0.5, s["left_items"], size=16, color=RED)
            # 右列（绿色标题 + 正确做法）
            rh = min(len(s["right_items"]) * 0.65 + 0.8, 4.8)
            right_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                               Inches(6.8), Inches(1.5), Inches(5.9), Inches(rh))
            right_box.fill.solid(); right_box.fill.fore_color.rgb = PptRGB(0xE8, 0xF8, 0xEE)
            right_box.line.color.rgb = GREEN; right_box.line.width = PptPt(1.5)
            add_text(slide, 7.05, 1.58, 5.5, 0.45, s["right_label"], 18, GREEN, bold=True)
            add_bullets(slide, 7.05, 2.1, 5.5, rh - 0.5, s["right_items"], size=16, color=GREEN)
            footer(slide, idx, s.get("chapter", ""))

        elif t == "flow":
            bg(slide, WHITE)
            top_bar(slide, s["title"], s["kicker"])
            steps = s["steps"]
            n = len(steps)
            sw = 1.95          # step box width (inches)
            gap = 0.32         # gap between boxes
            total_w = n * sw + (n - 1) * gap
            sx = (13.333 - total_w) / 2  # center horizontally
            sy = 2.4           # top y
            sh = 2.8           # step box height
            arrow_color = CYAN
            for i, label in enumerate(steps):
                x = sx + i * (sw + gap)
                # 圆角矩形
                box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                             Inches(x), Inches(sy), Inches(sw), Inches(sh))
                if i == 0:
                    box.fill.solid(); box.fill.fore_color.rgb = NAVY
                    txt_color = WHITE
                elif i == n - 1:
                    box.fill.solid(); box.fill.fore_color.rgb = GREEN
                    txt_color = WHITE
                else:
                    box.fill.solid(); box.fill.fore_color.rgb = LIGHT
                    txt_color = INK
                box.line.color.rgb = BLUE; box.line.width = PptPt(1.5)
                # 多行文本
                tb = slide.shapes.add_textbox(Inches(x + 0.08), Inches(sy + 0.15),
                                              Inches(sw - 0.16), Inches(sh - 0.3))
                tf = tb.text_frame; tf.word_wrap = True
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                p = tf.paragraphs[0]; p.text = label
                p.font.size = PptPt(14); p.font.color.rgb = txt_color
                p.font.bold = True; p.font.name = "微软雅黑"
                p.alignment = PP_ALIGN.CENTER
                # 箭头（除最后一个）
                if i < n - 1:
                    ax = x + sw + 0.04
                    arr = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                                  Inches(ax), Inches(sy + sh / 2 - 0.18),
                                                  Inches(gap - 0.08), Inches(0.36))
                    arr.fill.solid(); arr.fill.fore_color.rgb = arrow_color
                    arr.line.fill.background()
            # 底部说明
            add_text(slide, 0.8, 5.8, 11.7, 0.8,
                     "纠正话术示例：「文档第 X 处的数字 ___ 与源表不符，请检查全部不一致项后统一修改。」",
                     14, BLUE)
            footer(slide, idx, s.get("chapter", ""))

        elif t == "closing":
            bg(slide, NAVY)
            add_text(slide, 1.0, 2.0, 11.3, 0.6, s["kicker"], 22, CYAN, bold=True)
            add_text(slide, 1.0, 2.7, 11.3, 1.0, s["title"], 40, WHITE, bold=True)
            add_text(slide, 1.0, 3.9, 11.3, 2.4,
                     "\n".join("· " + b for b in s["bullets"]), 22, LIGHT)
            footer(slide, idx, "寄语")

    prs.save(path)
    return path


def add_text(slide, l, t, w, h, text, size, color, bold=False, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = text; p.alignment = align
    p.font.size = PptPt(size); p.font.color.rgb = color
    p.font.bold = bold; p.font.name = "微软雅黑"
    return box


def add_bullets(slide, l, t, w, h, items, size=20, color=GREY):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame; tf.word_wrap = True
    first = True
    for it in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.text = "• " + it
        p.font.size = PptPt(size); p.font.color.rgb = color; p.font.name = "微软雅黑"
        p.space_after = PptPt(8)
    return box


# ============================================================
# PDF 渲染（reportlab，与 PPT 内容一致）
# ============================================================
def render_pdf(path):
    pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
    FONT = 'STSong-Light'
    pw, ph = A4[0], A4[1]
    c = RLCanvas.Canvas(path, pagesize=(pw, ph))
    total = len(SLIDES)

    def rect(x, y, w, h, fill, stroke=None, dash=False):
        if fill:
            c.setFillColorRGB(*[v / 255 for v in fill])
            c.rect(x, y, w, h, fill=1, stroke=0)
        if stroke:
            c.setStrokeColorRGB(*[v / 255 for v in stroke])
            c.setLineWidth(2)
            if dash:
                c.setDash(4, 3)
            c.rect(x, y, w, h, fill=0, stroke=1)
            c.setDash()

    def text(x, y, s, size, color, bold=False):
        c.setFillColorRGB(*[v / 255 for v in color])
        c.setFont(FONT, size)
        c.drawString(x, y, s)

    def wrap_lines(items, max_chars=36):
        """简单换行"""
        out = []
        for it in items:
            while len(it) > max_chars:
                out.append(it[:max_chars] + "…")
                it = "…" + it[max_chars:]
            out.append(it)
        return out

    def footer_p(idx, chapter):
        c.setFillColorRGB(*[v / 255 for v in GREY_H])
        c.setFont(FONT, 10)
        c.drawString(30, 22, chapter)
        c.drawRightString(pw - 30, 22, f"{idx}/{total}")

    for idx, s in enumerate(SLIDES, start=1):
        t = s["type"]
        if t == "cover":
            rect(0, 0, pw, ph, NAVY_H)
            text(60, ph - 150, s["kicker"], 22, (0xEA, 0xF1, 0xFD))
            text(60, ph - 220, s["title"], 40, (0xFF, 0xFF, 0xFF), bold=True)
            yy = ph - 300
            for b in s["bullets"]:
                text(70, yy, "· " + b, 18, (0xEA, 0xF1, 0xFD)); yy -= 30
            footer_p(idx, "WorkBuddy 零基础办公任务实战教程")

        elif t == "divider":
            rect(0, 0, pw, ph, NAVY_H)
            text(60, ph - 220, s["kicker"], 26, (0x17, 0xC0, 0xC7), bold=True)
            text(60, ph - 300, s["title"], 44, (0xFF, 0xFF, 0xFF), bold=True)
            footer_p(idx, s["title"])

        elif t == "content":
            rect(0, 0, pw, ph, (0xFF, 0xFF, 0xFF))
            rect(0, ph - 90, pw, 90, NAVY_H)
            text(50, ph - 62, s["title"], 26, (0xFF, 0xFF, 0xFF), bold=True)
            text(50, ph - 82, s["kicker"], 12, (0xEA, 0xF1, 0xFD))
            yy = ph - 130
            for b in s["bullets"]:
                text(60, yy, "• " + b, 18, GREY_H); yy -= 34
            footer_p(idx, s.get("chapter", ""))

        elif t == "chart":
            rect(0, 0, pw, ph, (0xFF, 0xFF, 0xFF))
            rect(0, ph - 90, pw, 90, NAVY_H)
            text(50, ph - 62, s["title"], 24, (0xFF, 0xFF, 0xFF), bold=True)
            text(50, ph - 82, s["kicker"], 12, (0xEA, 0xF1, 0xFD))
            # 嵌入图片
            img_path = os.path.join(CHART_DIR, s["image"])
            if os.path.exists(img_path):
                from reportlab.lib.utils import ImageReader
                ir = ImageReader(img_path)
                iw, ih = ir.getSize()
                scale = min((pw - 100) / iw, (ph - 250) / ih) * 0.88
                c.drawImage(ir, 50, ph - 430, iw * scale, ih * scale)
            else:
                text(50, ph - 280, "[图表缺失]", 18, RED_H)
            if s.get("caption"):
                text(50, ph - 450, s["caption"], 13, BLUE_H)
            yy = ph - 130
            for b in s.get("bullets", []):
                text(540, yy, "• " + b, 14, GREY_H); yy -= 26
            footer_p(idx, s.get("chapter", ""))

        elif t == "compare":
            rect(0, 0, pw, ph, (0xFF, 0xFF, 0xFF))
            rect(0, ph - 90, pw, 90, NAVY_H)
            text(50, ph - 62, s["title"], 24, (0xFF, 0xFF, 0xFF), bold=True)
            text(50, ph - 82, s["kicker"], 12, (0xEA, 0xF1, 0xFD))
            # 左列
            rect(40, ph - 420, pw / 2 - 60, 310, (0xFD, 0xED, 0xEC), stroke=RED_H)
            text(52, ph - 400, s["left_label"], 16, RED_H, bold=True)
            ly = ph - 370
            for item in wrap_lines(s["left_items"]):
                text(56, ly, "· " + item, 14, RED_H); ly -= 24
            # 右列
            rect(pw / 2 + 20, ph - 420, pw / 2 - 60, 310, (0xE8, 0xF8, 0xEE), stroke=GREEN_H)
            text(pw / 2 + 32, ph - 400, s["right_label"], 16, GREEN_H, bold=True)
            ry = ph - 370
            for item in wrap_lines(s["right_items"]):
                text(pw / 2 + 36, ry, "· " + item, 14, GREEN_H); ry -= 24
            footer_p(idx, s.get("chapter", ""))

        elif t == "flow":
            rect(0, 0, pw, ph, (0xFF, 0xFF, 0xFF))
            rect(0, ph - 90, pw, 90, NAVY_H)
            text(50, ph - 62, s["title"], 24, (0xFF, 0xFF, 0xFF), bold=True)
            text(50, ph - 82, s["kicker"], 12, (0xEA, 0xF1, 0xFD))
            steps = s["steps"]
            n = len(steps)
            bw = (pw - 120) / n - 20
            bh = 160
            bx0 = 60
            by = ph - 320
            for i, lbl in enumerate(steps):
                bx = bx0 + i * (bw + 20)
                fc = NAVY_H if i == 0 else (GREEN_H if i == n - 1 else LIGHT_H)
                tc = (0xFF, 0xFF, 0xFF) if i in (0, n - 1) else INK
                rect(bx, by, bw, bh, fc, stroke=BLUE_H)
                lines = lbl.split("\n")
                cy = by + bh - 25
                for ln in lines:
                    c.drawCentredString(bx + bw / 2, cy, ln)
                    cy -= 18
                if i < n - 1:
                    ax = bx + bw + 3
                    c.setFillColorRGB(*[v / 255 for v in CYAN])
                    c.drawString(ax, by + bh / 2 - 5, "→")
            text(50, ph - 450,
                 "纠正话术示例：「文档第 X 处的数字 ___ 与源表不符，请检查全部不一致项后统一修改。」",
                 13, BLUE_H)
            footer_p(idx, s.get("chapter", ""))

        elif t == "closing":
            rect(0, 0, pw, ph, NAVY_H)
            text(60, ph - 160, s["kicker"], 22, (0x17, 0xC0, 0xC7), bold=True)
            text(60, ph - 220, s["title"], 36, (0xFF, 0xFF, 0xFF), bold=True)
            yy = ph - 290
            for b in s["bullets"]:
                text(70, yy, "· " + b, 20, (0xEA, 0xF1, 0xFD)); yy -= 32
            footer_p(idx, "寄语")

        c.showPage()
    c.save()
    return path


# ============================================================
# 大纲 / 备注
# ============================================================
def write_outline(path):
    lines = ["# 课程页面索引（slide-outline.md）", "",
             "> 与 `workbuddy-beginner-training.pptx` / `.pdf` 严格对应。", "",
             "> 本版本不含截图占位。Demo 章节以处理流程 + 可视化图表 + 错误纠正呈现。", ""]
    type_cn = {"cover": "封面", "divider": "章节", "content": "内容",
               "chart": "图表", "compare": "对比", "flow": "流程", "closing": "结语"}
    for i, s in enumerate(SLIDES, start=1):
        extra = ""
        if t := s.get("image"):
            extra += f" | 图:{t}"
        if s.get("steps"):
            extra += f" | 流程({len(s['steps'])}步)"
        lines.append(f"{i:02d} | {type_cn[s['type']]} | {s['title']}{extra}")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return path


def write_notes(path):
    lines = ["# 讲师备注（slide-notes.md）", "",
             "> 每页一句到两句的讲解要点，配合 `full-speaking-script.md` 使用。", ""]
    for i, s in enumerate(SLIDES, start=1):
        lines.append(f"## 第 {i} 页 · {s['title']}")
        lines.append("")
        lines.append(s.get("note", ""))
        lines.append("")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return path


if __name__ == "__main__":
    os.makedirs(SLIDE_DIR, exist_ok=True)
    p1 = render_pptx(os.path.join(SLIDE_DIR, "workbuddy-beginner-training.pptx"))
    p2 = render_pdf(os.path.join(SLIDE_DIR, "workbuddy-beginner-training.pdf"))
    p3 = write_outline(os.path.join(SLIDE_DIR, "slide-outline.md"))
    p4 = write_notes(os.path.join(SLIDE_DIR, "slide-notes.md"))
    print(f"[OK] {os.path.relpath(p1, ROOT)}  ({len(SLIDES)} 页)")
    print(f"[OK] {os.path.relpath(p2, ROOT)}")
    print(f"[OK] {os.path.relpath(p3, ROOT)}")
    print(f"[OK] {os.path.relpath(p4, ROOT)}")
