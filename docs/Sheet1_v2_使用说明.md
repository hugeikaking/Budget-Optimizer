# M1/M2 Sheet1 适配版（v2）使用说明

> **仅用于脱敏测试工作簿，尚未在公司真实 Excel 环境验证。不要直接针对正式预算表运行。** 此 GitHub 仓库公开，禁止上传真实公司预算、客户价格、凭证或结果文件。

## 修复内容

旧版宏 `RefreshBudgetPreview` 固定要求 `Budget_Data`、`Budget_Panel`、`Edit_Plan`、`Base26` 四张工作表，只有 `Sheet1` 时触发 `Subscript out of range`。

新版源码为 `src/BudgetOptimizer_Sheet1_v2.bas`，支持：
- 同一 .xlsm 内数据为 `Sheet1`：自动创建 `Budget_Panel`、`Edit_Plan`、`Base26`、`H40_List`。
- 独立工具 .xlsm + 数据 .xlsx：先运行 `ImportBudgetSheet1`，选取包含 `Sheet1` 的工作簿；程序只读打开该文件、复制到工具中的 `Budget_Data`，关闭原文件不保存。
- 从 Set QTY 的 X:AQ 或 Sales Amt 的 DW:EP 区块**前 25 行**识别连续的 20 个滚动月份表头（如 `2026_11`、`2027_1`）；不再假设固定年份。
- 以 S 列 H40 分组，以 `Edit_Plan` 作候选改动，生成临时 `Budget_Trial`，保留原始 `Sheet1` / `Budget_Data` 不变。
- 26F 全年基准输入在 `Base26` 的 B:E 列；试算期间按本次表内 26F 差额修正。**尚不支持滚动文件之间的历史月份自动结转。**

## 最推荐的首次测试（独立工具工作簿）

1. 在公司 Excel 中新建空白工作簿，另存为 **`Budget_Optimizer.xlsm`**（须允许宏）。
2. 打开 VBA 编辑器（开发工具 → Visual Basic）；在该 .xlsm 导入 `BudgetOptimizer_Sheet1_v2.bas`。
3. **若已导入旧版模块，可以保留，但以后只运行新版下述宏**；不要运行旧 `RefreshBudgetPreview`。
4. 先运行 `ImportBudgetSheet1`，选择复制粘贴后保存的测试 Excel（其中数据标签为 `Sheet1`），**只使用脱敏测试数据**。
5. 导入完成后会自动创建面板，打开 `Budget_Panel`，在 `B3` 选择 H40，`E3` 输入关注年度（例如 2027）。
6. 到 `Base26` 按 H40 填写完整 26F 的 Demand Amt、Demand Qty、Sales Amt、Sales Qty（若有基准数据可填；没有可先空着）。
7. 运行 `RefreshBudgetSheet1Preview`。查看 `Budget_Panel` 的 26F 表内、27F 表内四项指标、26F 基准修正与 YoY，以及 `Budget_Trial` 数据。
8. 可选：在 `Edit_Plan` 输入修改行（**预算原表里的行号**）、起止月份（`YYYY_MM`）、Pcs/Set、S/P；第一列填 `Y` 启用，再运行刷新。只改变 `Budget_Trial`。

## 当前限制

- 目前是**手动刷新**，不是单元格变化后自动计算；没有自动寻找最优数字。
- Demand Qty、Demand Amt、Sales Qty 按已确认的模拟计算规则计算；Sales Amt 优先读取试算 Sheet 的 **DW:EP Excel 公式结果**。正式表中必须进一步验证各项指标口径。
- `S` 列必须确实是 H40，且年月表头在前 25 行，数值块宽度与 X:AQ、AR:BK、BL:CE、CF:CY、CZ:DS、DW:EP 一致。
- `ImportBudgetSheet1` 读取外部文件中的 `Sheet1`；若已存在 `Budget_Data` 会在确认后替换，但保留 `Base26`、`Edit_Plan`。输入计划需自行检查是否仍与新文件行号吻合。
- 还未实现原表回写、标记 C、改动单元格标黄、差异审批、自动优化、跨月份滚动版本历史同步。
- 若源数据使用外部链接或跨工作表公式，复制后的计算结果需要检查。
- 切忌将真实公司预算表或含业务数据的 .xlsm 上传至公开 GitHub。

## 运行时报错怎么反馈

把弹窗内的错误号码和文字，以及「`Sheet1` 或 `Budget_Data` 是否存在」告诉我即可。不要发真实的公司数据或客户明细。
