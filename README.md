# 预算优化 · Budget Optimizer

公司内部 Excel VBA 预算试算工具。公开仓库**只含通用代码及说明，不得上传真实公司预算或客户明细**。

## 当前推荐：V3 两页式预算

- 源码：[src/BudgetOptimizer_V3.bas](src/BudgetOptimizer_V3.bas)
- 步骤与验收：[docs/V3_两页操作与验收.md](docs/V3_两页操作与验收.md)
- 之前已经过公司 Excel 验证的 M1/M2 版本保存在 [src/BudgetOptimizer_Sheet1_v2.bas](src/BudgetOptimizer_Sheet1_v2.bas)，可回退。

### Budget_Overview（第一页）
- 26F TTL Set QTY **仅手动输入一个全年值**；27F 按 B 列机种唯一编号分别配置目标。
- 自动建议 27F 12 个月 Set QTY：4–9 月逐渐上升、10 月下降、11 月至次年 1 月再上升、2 月下降、3 月回升，月度台数取两位有效数字的整数。
- 同一个机种下所有 H40 行的同月台数必须一致；机种 Set QTY 总量去重，Sales Amt 继续按所有 H40 业务行汇总。
- 生成试算后，月度台数直接链接 Budget_Trial Excel 公式，TTL Sales Amt 与月度图表联动。

### Budget_Detail（第二页）
- 原版 H40 四项指标及 Edit_Plan 试算。
- Pcs/Set 非负整数、S/P 最多五位小数、Share 和 Comp Ratio 是 0–100 的整数。
- 所有数据调整均在 Budget_Trial 中执行，不直接修改 Sheet1 / Budget_Data。

**27F 是 2027/04–2028/03；26F 是 2026/04–2027/03，不是自然年。**

## 当前状态

V2 的原始四项指标、财年汇总、Pcs/Set 和 S/P 试算等已经在公司 Excel 人工核对通过。V3 新版**已经提交，但未完成公司 Excel 验收**。其自动月度分配、两页联动、Set QTY 图表和金额重算都需要按说明逐项核对，不应直接用于正式预算数据输出。正式确认/回写、变更审批、标黄导出及跨滚动文件差额结转仍未实现。

## 本地与隐私

用户 Windows 项目目录：`E:\WaPJTS\预算表优化`。公司数据只留在用户本地，勿将公司真实数据、业务价格和工作簿上传 GitHub。
