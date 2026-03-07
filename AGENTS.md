# 仓库规范

## 规则优先级与维护边界
- 本文件用于维护“必须执行”的硬规则与验收门槛。
- `README.md` 仅保留快速入口、命令示例与最小说明，不再维护长篇硬规则。
- 若 `README.md` 与本文件存在冲突，以本文件为准。

## 项目结构与模块组织
本仓库以模板与数据驱动为主（不是传统应用代码仓库）。

- `README.md`：主流程与报告约束说明。
- `template.docx`：最终交付所需的 Word 模板。
- `*第四章数据.xlsx`：第4章市场分析的数据源文件。
- `autofile/<疾病名>/`：生成产物目录（不存在时需自动创建）。
- `autofile/<疾病名>/figures/`：全部导出的图表文件（如 `fig_1_1.png`、`fig_4_1.png`）。
- `autofile/<疾病名>/ch04_codex_extract.json`：由 Codex 会话生成的第四章标准化结构化数据。
- `autofile/<疾病名>/backup/`：覆盖写入前的时间戳备份目录。

所有生成文件必须写入 `autofile/` 内，禁止写到仓库根目录。

## 构建、测试与开发命令
当前未配置构建系统。请使用轻量 PowerShell 校验命令：

```powershell
# 创建输出根目录
New-Item -ItemType Directory -Force autofile

# 查看某病种的预期输出
Get-ChildItem "autofile/<疾病名>" -Recurse

# 快速检查关键文件是否存在
@("00_evidence.txt","manifest_text.csv","manifest_fig.csv","ch04_agg_tables.xlsx") |
  ForEach-Object { Test-Path "autofile/<疾病名>/$_" }
```

## 编码风格与命名约定
- 文本文件统一使用 UTF-8 编码，并保留中文医学术语。
- 文件命名需简洁且可复现：
  - 章节文本：`ch01.txt` ... `ch07.txt`
  - 图表文件：`fig_<chapter>_<index>.png`
  - 最终交付：`《<疾病名>疾病市场分析报告》_final.docx`
- 各类清单文件（manifest）需保持机器可读（`.csv`、稳定字段名、避免临时重命名）。

## 测试与验证规范
当前没有自动化测试框架。请将验证视为发布闸门：

- 确认 `autofile/<疾病名>/` 下必需文件齐全。
- 确认第4章图表均来源于 Excel 数据，并且标签一致。
- 确认覆盖写入会在 `backup/` 下生成备份。
- 打开最终 `.docx`，检查目录、标题层级、图注与数据来源行是否正确渲染。

## 提交与合并请求规范
当前目录没有可用 Git 历史，后续请按以下标准执行：

- 提交信息格式：`type(scope): summary`（示例：`docs(readme): clarify chapter-4 outputs`）。
- 每次提交应聚焦单一变更主题（模板改动、数据改动、示例产物尽量分开）。
- PR 需包含：变更目的、修改文件列表、已执行验证步骤、版式改动前后截图。
- 有任务/问题编号时请在 PR 中关联。

## Codex 优先 AI 策略（强制）
- 本仓库中的“AI主导”定义为：**由当前 Codex 会话模型主导写作与语义决策**。
- Python 脚本仅做辅助：数据整理、图表渲染、docx 装配、QA 校验、备份与清单处理。
- 禁止新增或恢复任何需要 API 凭证的脚本侧外部模型调用。
- 正常工作流中不得要求设置 `OPENAI_API_KEY`。
- 禁止恢复 `--gen-mode ai` 或其他脚本内 LLM 生成路径。
- 默认协作方式：用户仅提供疾病名与第4章 Excel 数据；Codex 负责写作，并负责将第4章 Excel 提取为 `ch04_codex_extract.json`；脚本负责装配与校验。

## 图表 QA 工作流（强制）
- 图表生成后，必须先执行脚本侧 QA（确定性规则，见 `qa_check.txt`）。
- 在最终交付前，必须由 **Codex 会话** 对关键机制图做视觉复核（重点是 `fig_2_3.png`）。
- 脚本与 Codex 职责拆分如下：
  - 脚本：结构化规则检查（标题/来源一致性、语义约束、几何与可读性规则等）。
  - Codex：图像级临床逻辑复核，并给出最终通过/驳回判断。
- 注意：Python 脚本无法直接调用 Codex；视觉复核必须在 Codex 会话中显式触发。
