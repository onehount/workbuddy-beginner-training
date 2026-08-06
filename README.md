# WorkBuddy 零基础办公任务实战教程

> 以新式茶饮品牌市场部门月度复盘为例，学习如何让 WorkBuddy 理解任务、制定方案、处理材料并生成办公成果。

[![Content License: CC BY 4.0](https://img.shields.io/badge/Content-CC%20BY%204.0-blue.svg)](./LICENSE-CONTENT)
[![Code License: MIT](https://img.shields.io/badge/Code-MIT-green.svg)](./LICENSE-CODE)
[![Version](https://img.shields.io/badge/version-v1.0-informational.svg)](./CHANGELOG.md)

---

## ⚠️ 虚构声明

> **本项目中的品牌、人员、部门、活动、数据和文件均为虚构，仅用于教学演示，与任何真实企业无关。**

案例品牌「青屿茶研 QINGYU TEA LAB」为完全虚构的新式茶饮品牌。所有人名、岗位、活动、渠道数据、会议记录均为教学构造，不对应任何现实组织或个人。

---

## 这个项目是什么

这是一套 **60 分钟、面向零基础新员工** 的 WorkBuddy 实体教学材料，完整开源。

它不教你"按钮怎么点"，而是回答三个问题：

| | |
|---|---|
| **本质** | WorkBuddy 的本质不是替人思考，而是把人的目标、材料和约束，转化为一系列可执行任务。 |
| **效率** | 效率不是一次生成更多文字，而是减少任务澄清、材料整理、重复排版和多文件转换中的机械劳动。 |
| **未来** | 未来的工作能力不仅是会用某个软件，更是能清楚定义目标、组织材料、检查结果，并与 AI 协同完成任务。 |

## 你能拿到什么

- ✅ 一套可直接开讲的 **教学 PPT**（PPTX + PDF）与 **逐页讲师稿**
- ✅ 一套完整的 **Demo 输入材料**（会议记录、数据表、5 类文档模板）
- ✅ 一套对应的 **Demo 输出成果**（WorkBuddy 原始版 + 人工审核版对照）
- ✅ **8 份可直接复制的提示词**，每份只做一件事
- ✅ 全流程 **操作截图规范 + 批量标注/放大/合成脚本**
- ✅ **学员自行复现指南**，从零开始跑通同一套流程

## 快速开始

### 我是学员，想自己复现

```
1. 下载本仓库（Code → Download ZIP，或 releases/ 目录下的发布包）
2. 打开 09-student-guide/quick-start.md
3. 按 09-student-guide/reproduction-guide.md 逐步操作
4. 用 09-student-guide/operation-checklist.md 自检
```

### 我是讲师，想直接开课

```
1. 打开 02-slides/workbuddy-beginner-training.pptx
2. 对照 03-instructor-guide/full-speaking-script.md 逐页讲
3. 用 03-instructor-guide/timing-guide.md 控时（总计 60 分钟）
4. 课堂提问参考 03-instructor-guide/classroom-faq.md
```

### 我想改造成自己公司的版本

```
1. 替换 04-demo-inputs/ 下的材料为你的业务场景
2. 06-prompts/ 里的提示词结构可直接沿用，只改业务名词
3. 按 11-quality-assurance/privacy-review.md 做脱敏检查
4. 遵守 LICENSE-CONTENT（CC BY 4.0）保留署名
```

## 课程核心：一个完整闭环

```
提出任务
  → 让 WorkBuddy 补充提问
    → 回答关键问题
      → 获得可执行方案
        → 准备并上传附件
          → 分步骤执行任务
            → 查看生成结果
              → 提出修改要求
                → 人工检查
                  → 导出最终文件
```

**最关键的一步是第 2 步。** 大多数人失败，是因为跳过它，直接要最终文件。

## 目录结构

| 目录 | 内容 |
|---|---|
| [`01-course-overview/`](./01-course-overview/) | 课程目标、教学对象、议程、学习成果 |
| [`02-slides/`](./02-slides/) | 教学 PPT（PPTX/PDF）、逐页大纲、页面备注 |
| [`03-instructor-guide/`](./03-instructor-guide/) | 讲师手册、完整讲稿、控时表、课堂 FAQ |
| [`04-demo-inputs/`](./04-demo-inputs/) | Demo 输入：文档、数据表、模板 |
| [`05-demo-outputs/`](./05-demo-outputs/) | Demo 输出：WorkBuddy 原始版 / 人工审核版 / 预期结果 |
| [`06-prompts/`](./06-prompts/) | 8 份可复制提示词 |
| [`07-screenshots/`](./07-screenshots/) | 原始截图 / 标注截图 / 合成图 + 索引 |
| [`08-design-assets/`](./08-design-assets/) | 封面、社交分享图、PPT 母版、截图模板 |
| [`09-student-guide/`](./09-student-guide/) | 学员快速入门、复现指南、操作清单、FAQ |
| [`10-agent-tasks/`](./10-agent-tasks/) | 项目任务清单、任务卡、验收标准 |
| [`11-quality-assurance/`](./11-quality-assurance/) | 内容/截图/PPT/隐私/发布 五类检查清单 |
| [`12-video-backup/`](./12-video-backup/) | Demo 录屏备份与录制说明 |
| [`tools/`](./tools/) | 截图批处理等自动化脚本（MIT） |
| [`releases/`](./releases/) | 打包发布物 |

## 统一 Demo 案例

**青屿茶研市场部 · 2026 年 7 月月度复盘**

| 维度 | 设定 |
|---|---|
| 市场活动 | 5 项（夏日新品试饮、小红书达人种草、微信社群优惠、商场快闪、校园联名） |
| 渠道 | 5 个（小红书、抖音、微信公众号、微信社群、线下门店） |
| 数据周期 | 2026 年 7 月 |
| 参会人员 | 6 人 |
| 会议原始记录 | 约 2000 字 |
| 最终汇报 PPT | 10–12 页 |

**核心指标口径（全项目统一）：**

```
转化率 = 有效线索数量 ÷ 线索数量 × 100%
```

| 渠道 | 曝光量 | 线索数量 | 有效线索 | 转化率 |
|---|---:|---:|---:|---:|
| 小红书 | 185,000 | 1,280 | 436 | 34.1% |
| 抖音 | 262,000 | 1,560 | 390 | 25.0% |
| 微信公众号 | 58,000 | 620 | 279 | 45.0% |
| 微信社群 | 32,000 | 510 | 306 | 60.0% |
| 线下门店 | 46,000 | 740 | 333 | 45.0% |
| **合计** | **583,000** | **4,710** | **1,744** | **约 37.0%** |

> 表中所有汇总值与转化率在 `04-demo-inputs/spreadsheets/` 的 xlsx 中由公式计算得出，不是硬编码文本。

## 课程议程（60 分钟 · 49 页 PPT）

> 本教程为 **功能导向** 结构：先讲 WorkBuddy 各功能是什么 / 干什么 / 怎么用 / 在哪里，
> 用统一案例（青屿茶研）演示每个功能的真实使用流程，再加进阶技巧，Demo 结果仅顺带展示。

| 章节 | 内容 | 页数 | 时长 |
|---|---|---:|---:|
| 封面 + 目标 | 课程目标：认知 + 操作 + 两个进阶习惯 | 2 | 3 min |
| 一 | 为什么学习 WorkBuddy（本质 / 效率 / 未来，困境对应功能，发布会式收益） | 4 | 4 min |
| 二 | 认识界面与起步（实际界面截图 / 抽象结构 / 工作模式 / 确认积分 / 下达任务 / 上下文） | 8 | 10 min |
| 三 | **核心功能逐个看**（10 大功能 + Demo 使用流程：文档 / 数据 / PPT / 设计 / 连接器…） | 19 | 24 min |
| 四 | 进阶技巧（先方案后执行 / 出错五步纠正 / 提示词四要素 / 四查清单） | 7 | 9 min |
| 五 | Demo 结果顺带展示（五类产出一览 / 可视化结果 / 一条命令复现） | 5 | 5 min |
| 六 | 检查、复现与寄语（三动作 / GitHub / 寄语） | 4 | 5 min |

## 授权

| 内容类型 | 许可证 |
|---|---|
| 教学内容（PPT、PDF、讲师稿、README、提示词说明、截图讲解、学员指南） | [CC BY 4.0](./LICENSE-CONTENT) |
| 代码与自动化脚本（`tools/`） | [MIT](./LICENSE-CODE) |

**关于 WorkBuddy 产品名称与界面：**

- WorkBuddy 产品名称、界面、标识和相关商标归原权利人所有。
- 本项目为独立制作的公开教学材料，不代表产品官方立场。
- 本项目不暗示与产品方存在商业合作或官方授权关系。

第三方素材授权见 [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md)。

## 相关文档

- [English Introduction](./README_EN.md)
- [更新日志](./CHANGELOG.md)
- [贡献指南](./CONTRIBUTING.md)
- [安全与隐私](./SECURITY.md)

---

<div align="center">

**初入职场，不必要求自己立刻掌握所有工具，也不必一次写出完美的指令。**

**真正重要的是理解工作的本质，知道目标是什么，知道结果应该如何判断。**

</div>
