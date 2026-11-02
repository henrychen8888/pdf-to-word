---
name: pdf-to-word
description: 将 PDF 转换为可编辑的 Word (.docx) 文档，支持单文件、批量目录、页范围截取，并在转换后自动核验文字覆盖率（不依赖段落顺序，可识别丢字与内容重复）。当用户提出「PDF 转 Word」「pdf 转 docx」「把这个 PDF 变成可编辑的 Word」「批量转换这些 PDF」「PDF 里的表格要能编辑」「扫描件 PDF 能不能转 Word」等需求时使用本技能。
agent_created: true
display_name: PDF 转 Word
description_zh: PDF 转 Word（双引擎 + 文字覆盖率核验）
description_en: Convert PDF to editable Word (.docx) with dual engines and post-conversion coverage verification
version: 1.0.0
---

# PDF 转 Word

把 PDF 转成可编辑的 .docx，并在转换后**用数字告诉用户转得怎么样**，而不是只交一个文件。

## 核心原则

转换结果必须经过覆盖率核验才能交付。只看「文件生成了」不够——PDF 转 Word 最常见的失败不是报错，而是**静默丢内容**：文件正常打开、页数看着对、里面少了一大半字。

## 最快路径

```bash
python scripts/pdf2word.py "<PDF 路径或目录>" -o "<输出目录>" --report report.json
```

- 不带 `-o` 时输出到源文件同目录
- 目录输入会递归处理其中所有 `.pdf`
- `--pages 1-10` 只转指定页范围
- 每份文件都会打印：`PASS/WARN/FAIL` + 覆盖率 + 页数 + 表格数 + 图片数

脚本退出码：`0` 全部成功 / `1` 有文件失败 / `2` 环境缺依赖。

## 工作流

1. **先确认依赖**（通常已就绪）：
   ```bash
   python -c "import pdf2docx, pymupdf, docx; print('ok')"
   ```
   缺失时：`python -m pip install pdf2docx PyMuPDF python-docx`

2. **执行转换**，同时输出 JSON 报告：
   ```bash
   python scripts/pdf2word.py "C:/path/文档.pdf" -o "C:/path/out" --report "C:/path/rep.json"
   ```

3. **读报告下结论，再决定要不要多说一句**：
   - `PASS`（覆盖率 ≥98%）→ 直接交付路径
   - `WARN`（95%~98%）→ 交付并提示抽查
   - `FAIL`（<95%）→ 不要只交文件，先说明哪里可能缺，建议复核
   - `OCR`（扫描件）→ 见下方「扫描件」

4. **需要更高质量时**用 `--engine best`：两引擎都跑，按「覆盖率优先、2% 以内比真表格结构」择优，耗时约翻倍。

## 引擎与取舍

| 引擎 | 文字覆盖率（实测） | 表格结构（30 页样本） | 速度 | 适用 |
|---|---|---|---|---|
| `pdf2docx`（默认） | 98.8%~98.9% | 58 个表，其中真表格 24 个 | 36~54s | 常规文档、报告、教程 |
| `layout`（自研兜底） | 100.0% | 14 个表，结构弱 | 33~37s | 文本框密集、pdf2docx 丢字时 |

`--engine auto`（默认）：先跑 `pdf2docx`，覆盖率 ≥95% 直接采用；低于 95% 自动换 `layout` 重转并取优。

择优不是只看覆盖率：pdf2docx 会把代码块、URL 退化成大量 1×1 表格（某 47 页样本 139 个表里有 138 个是 1×1），所以比较时只认「至少 2×2」的真表格。覆盖率差距在 2 个百分点以内时，优先保住表格结构。

`layout` 引擎（`scripts/layout_engine.py`）按 PyMuPDF 的阅读顺序重建流式文档，保留字号、颜色、加粗；图片按原始坐标浮动定位。文字更完整，但表格识别弱于 pdf2docx。

## 质量核验（scripts/qc.py）

```bash
python scripts/qc.py 原文件.pdf 产物.docx
```

指标是**字符多重集覆盖率**：统计产物中出现且数量不超过原文的字符占比。

必须用多重集而不是「连续片段匹配」：
- 按顺序切片匹配会把表格重排、分栏重排误判成丢字（实测同一文件被误判为 0%）
- 按 `id(cell._tc)` 去重统计表格会因 lxml 代理对象 id 复用而严重低估（实测 60 格只读到 11 格）

同时输出 `extra`（产物中原文没有的字符），非 0 说明存在重复或异常插入。

## 必须如实告知的局限

**PPT / 设计稿型 PDF 转 Word 会损失版式，这是格式本质差异，不是可修复的 bug。**

PDF 是绝对定位格式，Word 是流式格式。当源 PDF 是「整页背景图 + 叠加文本框 + 色块图形」时：

- 文字可以完整保留（覆盖率可达 100%）
- 矢量色块、连接线、SmartArt 类图形无法还原
- 版面会重排，页数可能变化（实测 30 页 → 36 页）
- pdf2docx 会把大量文本框退化成 1×1 表格

遇到这类文档，主动说明「文字和表格可编辑，版式需重新排版」，不要让用户误以为是转换失败。

普通文档（报告、论文、合同、教程）版式还原良好。

## 扫描件

脚本会检测文字层：每页不足约 30 个字符即判定为扫描件，标记 `OCR` 并告警。

此时不要直接交付转换结果——产物基本是空壳。正确做法：

1. 先做 OCR 得到带文字层的 PDF，再转 Word
2. 或直接告知用户需要 OCR，说明本技能只做格式转换、不含识别能力
3. 不要在没有文字层的情况下声称转换成功

## 常见故障

| 现象 | 原因 | 处理 |
|---|---|---|
| `ERROR: 无可用引擎` | 缺 pdf2docx | `pip install pdf2docx` |
| `PDF 已加密` | 有密码保护 | 先解密，或「打印为 PDF」重新导出 |
| 覆盖率低但文件能开 | 特殊字体 / 复杂版面 | `--engine layout` 或 `--engine best` |
| 产物页数暴增 | 图片被内嵌撑开版面 | 用 `pdf2docx`；`layout` 已改为浮动定位 |
| 命令无输出直接退出 | 路径含空格未加引号 | 路径整体加双引号 |

## 脚本清单

- `scripts/pdf2word.py` — 主入口，含引擎调度、择优、质量自检、JSON 报告
- `scripts/layout_engine.py` — 自研版面引擎，文字零丢失优先
- `scripts/qc.py` — 独立质量核验，可对任意 docx 做覆盖率体检

细节与实测数据见 `references/quality-and-pitfalls.md`。
