# 第三方素材授权清单

本文件登记本项目使用的全部第三方素材，包括字体、图标、图片、模板、代码库。

**登记原则：宁可多记，不可漏记。** 任何无法说明来源与授权方式的素材，一律不得进入本仓库。

---

## 一、软件依赖（代码类）

| 名称 | 版本 | 来源 | 作者 / 维护方 | 授权方式 | 使用位置 |
|---|---|---|---|---|---|
| Pillow | ≥10.0 | https://pypi.org/project/pillow/ | Jeffrey A. Clark and contributors | HPND License | `tools/build_composites.py`、`tools/annotate.py` |
| python-docx | ≥1.1 | https://pypi.org/project/python-docx/ | Steve Canny | MIT License | `tools/gen_documents.py`、`tools/check_privacy.py` |
| openpyxl | ≥3.1 | https://pypi.org/project/openpyxl/ | Eric Gazoni, Charlie Clark | MIT License | `tools/gen_spreadsheets.py` |
| python-pptx | ≥1.0 | https://pypi.org/project/python-pptx/ | Steve Canny | MIT License | `tools/gen_slides.py` |

> 上述依赖仅在**构建期**使用，学员使用本教学材料时无需安装。

---

## 二、字体

| 名称 | 来源 | 授权方式 | 使用位置 | 是否随仓库分发 |
|---|---|---|---|---|
| 微软雅黑 (Microsoft YaHei) | Windows 系统内置 | 随 Windows 授权，**不可再分发** | PPT 中文正文与标题 | ❌ 否 |
| 思源黑体 (Source Han Sans) | https://github.com/adobe-fonts/source-han-sans | SIL Open Font License 1.1 | PPT 中文备选字体 | ❌ 否（请自行下载） |
| Aptos | Microsoft 365 内置 | 随 Microsoft 365 授权，**不可再分发** | PPT 英文与数字 | ❌ 否 |
| Arial | 系统内置 | 系统授权，**不可再分发** | PPT 英文与数字备选 | ❌ 否 |

> ⚠️ **重要：本仓库不分发任何字体文件。**
> PPT 使用系统字体名称引用。若你的环境缺少微软雅黑，PPT 会回退到思源黑体或系统默认黑体，可能出现轻微排版差异。
> 需要完全一致的排版效果时，请优先使用 `02-slides/workbuddy-beginner-training.pdf`。

---

## 三、图标与图形

| 名称 | 来源 | 作者 | 授权方式 | 使用位置 |
|---|---|---|---|---|
| 全部图标与装饰图形 | 本项目原创 | WorkBuddy Training Project | CC BY 4.0 | `08-design-assets/icons/`、PPT 页面 |
| 章节页几何背景 | 本项目原创（脚本生成） | WorkBuddy Training Project | CC BY 4.0 | PPT 章节页 |
| 封面视觉 | 本项目原创（脚本生成） | WorkBuddy Training Project | CC BY 4.0 | `08-design-assets/cover/` |
| 社交分享图 | 本项目原创（脚本生成） | WorkBuddy Training Project | CC BY 4.0 | `08-design-assets/social-preview/` |

> 本项目**未使用**任何图标库（Font Awesome / Iconfont / Noun Project 等）。
> 所有视觉元素由 `tools/` 下脚本以纯几何方式生成，避免授权争议。

---

## 四、图片与照片

| 名称 | 来源 | 授权方式 | 使用位置 |
|---|---|---|---|
| — | — | — | 本项目未使用任何摄影素材 |

> 本项目**不使用**任何图库照片（Unsplash / Pexels / 视觉中国 / 摄图网等）。
> 教学页面视觉全部由几何图形、文字排版和产品操作截图构成。

---

## 五、产品界面截图

| 名称 | 来源 | 权利归属 | 使用依据 | 使用位置 |
|---|---|---|---|---|
| WorkBuddy 软件界面截图 | 本项目自行采集 | 界面、标识、商标归 WorkBuddy 产品权利人所有 | 教学说明性合理使用 | `07-screenshots/` 全部、PPT 第五至七章 |

**声明：**

1. WorkBuddy 产品名称、界面外观、标识和相关商标归其原权利人所有。
2. 本项目为独立制作的公开教学材料，截图仅用于说明操作步骤。
3. 本项目不代表产品官方立场，不暗示与产品方存在商业合作或官方授权关系。
4. 若产品权利人认为本项目的截图使用不当，请通过 Issue 或 SECURITY.md 中的渠道联系，我们将在确认后移除或替换。

---

## 六、Demo 内容原创性声明

| 内容 | 原创性 |
|---|---|
| 品牌「青屿茶研 QINGYU TEA LAB」 | 本项目虚构，与任何真实企业无关 |
| 6 名参会人员及岗位 | 本项目虚构 |
| 5 项市场活动 | 本项目虚构 |
| 全部渠道曝光与线索数据 | 本项目构造，非真实业务数据 |
| 会议原始记录约 2000 字 | 本项目原创撰写 |
| 5 类文档模板 | 本项目原创设计 |
| 8 份提示词 | 本项目原创撰写 |
| 全部讲师稿与学员指南 | 本项目原创撰写 |

---

## 七、如何新增登记

向本仓库添加任何第三方素材时，必须在本文件对应章节新增一行，并填齐五个字段：

```
| 素材名称 | 来源 URL | 作者 | 授权方式 | 使用位置 |
```

缺任何一项，PR 不予合并。详见 [CONTRIBUTING.md](./CONTRIBUTING.md) 红线二。

---

_最后更新：2026-08-09_
