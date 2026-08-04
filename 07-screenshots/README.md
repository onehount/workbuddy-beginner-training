# 截图素材（可选）

> **本目录为可选素材。** 课程 PPT（`02-slides/`）已使用数据可视化图表和文字内容替代截图占位。
> 如果你希望为学员提供更直观的 WorkBuddy 操作参考，可按以下说明采集并加工截图。

## 目录结构

```
07-screenshots/
├── screenshot-spec.json    # 截图规范（JSON，含每张的编号/标题/高亮区域）
├── screenshot-pipeline.py  # 批处理工具（annotate / composite）
├── raw/                    # 原始截图（学员自行采集，1920×1080 PNG）
├── annotated/              # 加工后：原图 + 红色定位框
└── composite/              # 加工后：合成展示图（带 Step 编号和说明）
```

## 如何采集

1. 在 WorkBuddy 中按 `screenshot-spec.json` 中的步骤逐一操作
2. 每步完成后按 **PrtSc** 或截图工具截取全屏
3. 保存为 `raw/S05-01_create-task.png` 等（文件名须与 spec 中 `raw` 字段一致）
4. 分辨率建议 **1920×1080**，格式 **PNG**

## 如何加工

```bash
# 进入项目根目录
cd workbuddy-beginner-training

# 加标注（红色定位框）
python tools/screenshot_pipeline.py annotate

# 生成合成展示图
python tools/screenshot_pipeline.py composite
```

输出分别存入 `annotated/` 和 `composite/`。

## 截图清单

| 编号 | 步骤 | 高亮区域 | 说明 |
|------|------|---------|------|
| S05-01 | 创建任务 | 任务名称输入框 | 命名「7月市场复盘会议纪要」 |
| S05-02 | 上传材料 | 附件上传区域 | 原始记录 + 模板 |
| S05-03 | 输入约束 | 对话框文本区 | 不得虚构、区分三类 |
| S05-04 | 先出结构 | WB 回复区域 | 结构 + 待确认列表 |
| S05-05 | 待确认事项 | 列表区域 | 5 项未决内容 |
| S05-06 | 行动事项 | 表格区域 | 6 项含负责人/截止 |
| S05-07 | 生成纪要 | 文件预览/下载区 | 正式文档 |
| S05-08 | 最终结果 | 打开的文档 | 完整纪要 |

## 注意事项

- ⚠️ 截图中不得出现真实个人信息、账号、密码、企业内部路径
- ⚠️ 如使用了真实数据，请先脱敏再提交到仓库
- 本工具链基于 Pillow，依赖已在项目 venv 中安装
