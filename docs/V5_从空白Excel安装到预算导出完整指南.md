# Budget Optimizer V5｜从空白 Excel 到预算导出（Windows 操作手册）

> 适用版本：2026-11-24 已实机验收的 V4 核心预算试算、V5 界面、简化导出与导出按钮。以下按**新空白工作簿**开始。原始公司预算保留在本地，禁止上传公开 GitHub。新装版本先测试、确认后再转正式使用。

## 0. 准备材料

- Windows 桌面版 Excel，允许启用本地 VBA 宏。
- 一份独立的原预算文件（.xlsx/.xlsm），里面必须有名称为 **Sheet1** 的工作表；列布局需符合已验收的版本。
- GitHub 代码（使用各页面的 **Raw** 下载，不要复制整个网页）：  
  1. [BudgetOptimizer_V4.bas](../src/BudgetOptimizer_V4.bas)：计算引擎、导入、初始化和月度计划、Trial A 列 C 标记。
  2. [BudgetOptimizer_V5_UI.bas](../src/BudgetOptimizer_V5_UI.bas)：Overview / Detail 可视化布局、目标和实时金额区。
  3. [BudgetOptimizer_V5_ExportSimple.bas](../src/BudgetOptimizer_V5_ExportSimple.bas)：固定数值、改动标黄、行标 C 的简化导出。
  4. [BudgetOptimizer_V5_ExportButton.bas](../src/BudgetOptimizer_V5_ExportButton.bas)：Overview 一键导出按钮。
  5. [ThisWorkbook_V4_event.txt](../src/ThisWorkbook_V4_event.txt)：需要粘贴到 ThisWorkbook 的 Excel 事件代码。

**不需要**旧的 `C_BudgetV4Events.cls` 类模块，也不需要旧 `BudgetOptimizer_V3.bas`、旧 `BudgetOptimizer_V5_Export.bas`（保留公式的复杂导出）。避免同时导入同名公共模块。

## 1. 创建新工具

1. 关闭不相关的 Excel 工作簿。
2. 新建空白工作簿，**先另存为** `Budget_Optimizer_V5_Dev.xlsm`。选择“Excel 启用宏的工作簿 (*.xlsm)”，不要保存为 .xlsx。
3. 按 **Alt+F11** 进入 VBA 编辑器 → **File → Import File**，依次导入上面四个 .bas。
4. 左侧 **Microsoft Excel Objects → ThisWorkbook**，将 `ThisWorkbook_V4_event.txt` 中 `Workbook_SheetChange` 代码粘贴进去；不要把它当普通模块导入。一个工作簿中同名事件过程只能有一份。
5. 菜单 **Debug → Compile VBAProject**：成功时通常没有提示。保存（Ctrl+S），回到 Excel（Alt+Q）。
6. 确保 Excel 启用宏；只对信任的本地工具文件启用。

## 2. 导入原始预算

1. Alt+F8 → 运行 **`ImportBudgetSheet1`**。
2. 在文件选择窗口选择**单独保存、包含 Sheet1 的原预算测试文件**；程序以只读方式打开源文件，将 `Sheet1` 复制到当前工具的 `Budget_Data`。
3. 导入时会建立基础工作表，如 `Budget_Detail`、`Edit_Plan`、`Base26`、H40 列表和 `Budget_Overview`。
4. Alt+F8 → 运行 **`SetupBudgetOptimizerV4`** 一次。此时创建/更新 Detail 的 H40 编辑区；若 Trial 缺失，会初始化 `Budget_Trial`。
5. Alt+F8 → 运行 **`SetupBudgetOptimizerV5UI`** 一次，为 Detail 添加仪表板布局并给 Overview 添加年度金额 Target / Gap。
6. Alt+F8 → 运行 **`SetupBudgetV5ExportButton`** 一次，为 Overview 添加 **EXPORT BUDGET** 按钮。**不要重新运行这些 Setup 作为日常修改操作**。
7. 鼠标拖动底部标签，将 **Budget_Overview** 和 **Budget_Detail** 放到最前、紧挨着；其他后台表放后面。Ctrl+S 保存。

## 3. Overview（第一页）：年度目标、机种台数

- `B5`：**26F TTL Set QTY**，只手动填写年度总台数，不要求录入 26F 每机种每月历史。
- `B6`：**27F TTL Set QTY 目标**，手动填整数。
- `E9`：**27F TTL Sales Amt Target**，手动填写金额；是金额目标，和 B6 台数目标分开。
- `E6`：**27F TTL Sales Amt Current**，来自 Trial 的实时金额（绿色）。
- `E10`：**Sales Amt Gap = E6-E9**，负数代表不足，正数代表超出。
- 第 **40 行起** A 列显示 Model，B 列可选输入每机种 27F 年度目标，E:P 为 27F 12 个月 Set QTY（4 月至次年 3 月）。
- 目标输入后，Alt+F8 运行一次 **`GenerateSeasonalSetQtyPlanV3`**，自动生成机种分月建议，并建立与 Trial 的联动。默认季节性为 4–9 月上行、10 月回落、11–1 月再上行、2 月回落、3 月回升。
- 以后直接编辑 E:P 的月度 Set QTY 即可；同 B 列机种编号的不同 H40 行采用相同月度 Set QTY，数字会在 Trial 中实时变化。
- 黄色 = 人工目标/编辑区域；绿色 = Current Sales Amt。

**提醒**：该季节性建议采用两位有效数字取整，月度总数与手动年度目标可能存在少量偏差；请检查 B7/B9。

## 4. Detail（第二页）：按 H40 试算

- `B3` 下拉选择 H40，下面的四项指标和第 23 行起的原预算行明细自动加载。
- `E6:E9`：选定 H40 的四项目标（Demand Amt / Demand Qty / Sales Amt / Sales Qty），不同 H40 分别保存和恢复。
- `D6:D9`：选定 H40 的四项 27F 当前试算；`F6:F9`：与手动目标的差额。
- 第 23 行起：A = **Source row**（原预算对应行号），B = Model，C = H40，D:E = 生效起止月（YYYY_MM）；**F** = New Pcs/Set，**G** = New S/P；**H** = New Share，**I** = New Comp Ratio；J:M = 27F 第一个月的原值参考（注意各个月份原值可能不同）。
- Pcs/Set 输入非负整数；S/P 非负且最多 5 位小数；Share 和 Comp Ratio 均为 0–100 整数，必须先把 `R3` 设为 **YES** 才允许填写 H:I。
- 修改 F:I 后自动刷新 Trial，留在 Detail 页面，不用反复运行刷新宏。参数为空表示保持原值；需要恢复原预算时清空修改参数，或者按实际情况移除相关 Edit_Plan 修改。
- `Budget_Overview!E6` 和 `Budget_Detail!E11` 的 TTL Sales Amt 应保持一致。

## 5. Trial 与导出

- `Budget_Trial` = 原始数据 + Overview 的月度 Set QTY + Detail 的参数调整。
- **A 列 C**：该行 X:DS 五个预算输入块的任意月度值与 `Budget_Data` 不同，则显示 **C**；全部恢复原值后会自动取消。
- Overview 顶部点击 **EXPORT BUDGET**：首先显示“修改行数、修改单元格数”，确认后进入简化导出宏 `ExportBudgetV5SimpleCopy`。
- 选择新的 **.xlsx** 文件名（不能覆盖已存在文件）。
- 输出只有 `Sheet1`：保留原预算布局/主要格式；变化的输入值写成 Trial 的最终值并标**黄色**，发生变化的行 A 列写**固定的 C**；其他公式均转为**静态计算值**，DW:EP 的 Sales Amt 写为 Trial 最新计算的静态数值。不带 Change_Log。
- 你可以把必要的调整数字/标记复制回公司正式上传模板；粘贴时使用“选择性粘贴 → 值”或按公司要求操作，**切勿用导出文件直接覆盖正式上传模板的公式**。

## 6. 初次验收清单

1. Setup 后两个主页面紧邻显示，B3 可切换 H40，目标能分别记忆。
2. Overview E:P 的 Set QTY 能联动 Trial；Detail F/G 改动能联动 Trial 和两页 TTL Sales Amt。
3. Share / Comp Ratio 未解锁不能修改，解锁后生效。
4. Trial 改过的行 A 列出现 C；恢复原值并确认该行其他输入也无差异后，C 消失。
5. 导出新 .xlsx，确认 A 列固定 C、实际变化的 X:DS 单元格标黄，未变化的不新增黄色，DW:EP Sales Amt 与 Trial 相应单元格一致。
6. 原始 Budget_Data 保持不变；保存工具 .xlsm 后关闭并重新打开，事件仍自动生效。

## 7. 常见故障

- **Compile 没有提示**：通常代表编译通过。
- **Alt+F8 找不到 `RefreshBudgetSheet1Preview`**：该过程有可选参数，不一定显示在宏列表；可用 Alt+F11 → Ctrl+G 输入 `RefreshBudgetSheet1Preview False` 然后 Enter。
- **误导入 .cls 后 Compile 报 `VERSION 1.0 CLASS`**：删除错误普通模块，按本指南仅导入四个 .bas + 粘贴事件代码。
- **出现 `ref` 变量未定义**：不要使用旧 `BudgetOptimizer_V5_Export.bas`，本指南采用 `ExportSimple`。
- **修改后无自动更新**：检查本文件已启用宏、ThisWorkbook 事件代码存在，并确认 Excel `Application.EnableEvents` 没被其他宏关闭。
- **导出的文件没有黄色/行 C**：先检查 Trial 与原始表对应输入单元格是否真正不同；确认执行的是新简化导出宏。
- **重新更换滚动月份源文件**：当前工具尚未完整实现跨滚动窗口的目标/调整自动迁移。导入新源前应备份并单独核对 Edit_Plan 的行号、月份和机种键，避免旧计划应用错行。

## 8. 版本保存建议

- 首次全链路验收通过后保存 `Budget_Optimizer_V5_Stable.xlsm`，不随意改代码。
- 后续功能试验只在 `Budget_Optimizer_V5_Dev.xlsm` 进行。
- 公开 GitHub 仅保存代码与**不含真实业务数据**的说明；预算文件及客户信息只留在公司授权存储环境。
