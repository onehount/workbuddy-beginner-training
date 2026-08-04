# 贡献指南

感谢你愿意改进这套教学材料。

---

## 在提交之前

本项目是**公开教学材料**，不是产品代码库。因此有三条硬性红线，任何 PR 违反其一都会被直接关闭：

### 🚫 红线一：不得引入真实信息

- 不得出现任何真实企业名称、Logo、商标
- 不得出现真实员工姓名、手机号、邮箱、地址、工号
- 不得出现真实的账号、Token、密码、API Key、内网地址
- 不得出现可识别的真实文件路径（如 `C:\Users\张三\...`）
- 截图中不得出现真实的历史任务、聊天记录、通知弹窗

### 🚫 红线二：不得引入未授权素材

任何图片、字体、图标、模板，必须能说明来源与授权方式，并登记到 [`THIRD_PARTY_NOTICES.md`](./THIRD_PARTY_NOTICES.md)。

### 🚫 红线三：不得破坏数据一致性

本项目所有数字构成一个闭环。改动任何一个数字，必须同步更新：

```
04-demo-inputs/spreadsheets/*.xlsx
05-demo-outputs/**/*.docx
02-slides/workbuddy-beginner-training.pptx
03-instructor-guide/full-speaking-script.md
README.md 的数据表
```

统一口径不可更改：`转化率 = 有效线索数量 ÷ 线索数量 × 100%`

---

## 我可以贡献什么

| 类型 | 欢迎程度 | 说明 |
|---|---|---|
| 错别字、语病、链接失效修正 | ⭐⭐⭐ | 直接提 PR |
| 讲师稿表达优化 | ⭐⭐⭐ | 保持 60 分钟总时长 |
| FAQ 补充（真实课堂遇到的问题） | ⭐⭐⭐ | 请注明提问场景 |
| 截图批处理脚本改进 | ⭐⭐ | 需附测试用例 |
| 新增业务场景案例 | ⭐⭐ | 请先开 Issue 讨论 |
| 更换 PPT 视觉风格 | ⭐ | 需保持科技风与投影可读性 |
| 增加新的 Demo 类型 | ⭐ | v1.0 范围已冻结，请开 Issue |

---

## 提交流程

```bash
# 1. Fork 并克隆
git clone https://github.com/<your-name>/workbuddy-beginner-training.git
cd workbuddy-beginner-training

# 2. 建分支，分支名说明意图
git checkout -b fix/typo-in-chapter-5
git checkout -b docs/add-classroom-faq
git checkout -b feat/screenshot-tool-batch-mode

# 3. 修改并自检
python tools/check_repo.py          # 仓库完整性检查
python tools/check_privacy.py       # 隐私红线扫描

# 4. 提交
git commit -m "fix: 修正第五章第 3 页的转化率口径描述"

# 5. 推送并开 PR
git push origin fix/typo-in-chapter-5
```

### Commit 信息规范

```
<type>: <简要说明>

type 取值：
  feat     新增内容
  fix      修正错误
  docs     文档改动
  style    排版、格式，不影响含义
  refactor 结构调整
  chore    构建、脚本、工具
```

---

## 文件命名规范

本项目统一使用：

```
序号_内容名称_版本_日期.扩展名

示例：
  01_案例背景说明_v1.0_20260805.docx
  03_会议纪要_WorkBuddy原始版_v1.0_20260806.docx
  04_会议纪要_人工审核版_v1.0_20260806.docx
```

截图统一使用：

```
章节-步骤_类型_英文说明.png

示例：
  S05-03_RAW_upload-meeting-file.png
  S05-03_ANNOTATED_upload-meeting-file.png
  S05-03_COMPOSITE_upload-meeting-file.png
```

**禁止出现的命名：** `final`、`final2`、`最新版`、`真正最终版`、`新建文件`、`副本`

---

## PR 检查清单

提交前请逐项确认：

- [ ] 没有引入任何真实企业 / 人员 / 联系方式信息
- [ ] 没有引入未登记授权的第三方素材
- [ ] 改动的数字已在所有相关文件中同步
- [ ] 文件命名符合规范，未出现禁用词
- [ ] 相对链接均有效（`python tools/check_repo.py` 通过）
- [ ] 如改动 PPT，页数仍在 38–50 页，60 分钟可讲完
- [ ] 如改动讲师稿，`timing-guide.md` 已同步

---

## 授权约定

提交 PR 即表示你同意：

- 你贡献的**教学内容**以 [CC BY 4.0](./LICENSE-CONTENT) 授权
- 你贡献的**代码**以 [MIT](./LICENSE-CODE) 授权
- 你拥有所提交内容的合法权利

---

## 行为准则

就事论事，对内容不对人。教学材料的目标是让第一次接触 WorkBuddy 的人能看懂、能复现 —— 任何让材料变得更难懂的改动，无论多"专业"，都不会被合并。
