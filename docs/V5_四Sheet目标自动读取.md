# V5｜从四张下载表直接读取 2027F H40 目标

## 用户提供的源表格式（已由图片确认）

四个 Sheet 分别表示 Demand Amt、Demand Qty、Sales Amt、Sales Qty。每个 Sheet 的主要列均相同：

- B 列：H40（五位 H40 可能有尾部空格）
- F 列：2027F 目标
- 第 2 行：年份标题（F2 = 2027F）
- 第 3 行：Total，不导入
- 第 4 行起：按 H40 逐行排列，直到 B 列数据结束

不使用 A 列 Customer Account 来匹配，也不读 C/D/E 历史年度数据。G 列百分比不作为目标。**各 Sheet 的实际名称由用户在配置表填写**，程序不根据图片中的不完整 Sheet 标签臆测名称。

## 安装（首次一次；以后重复运行导入）

1. 保存和备份当前已修改的 `.xlsm` 工具。请在备份的测试副本先运行。
2. VBA 编辑器 Alt+F11 → File → Import File → 导入 [BudgetOptimizer_V5_Target4Sheets.bas](../src/BudgetOptimizer_V5_Target4Sheets.bas)。**不删除原 V4、V5_UI、ExportSimple、导出按钮，也不更改 ThisWorkbook**。
3. Debug → Compile VBAProject。Alt+F8 运行一次 `SetupBudgetV5FourSheetTargets`。
4. 新增的 `H40_Target_Source` 里，在 **B2:B5** 填写四个源 Sheet 的**准确名称**：
   - B2 = Demand Amt 所在 Sheet
   - B3 = Demand Qty 所在 Sheet
   - B4 = Sales Amt 所在 Sheet
   - B5 = Sales Qty 所在 Sheet
5. 上述 Setup 同时在 `Budget_Detail!G6:G9` 设置 `D6/E6` 至 `D9/E9` 的 **0% 整数百分比**。E 为空或 0 显示空白。

## 每次更新目标

1. 保留工具 `.xlsm`；无需重装任何 VBA。
2. Alt+F8 → `ImportBudgetV5FourSheetTargets`。
3. 选择公司最新下载的**包含四个目标 Sheet 的 Excel 文件**。可与预算 Sheet1 同文件，只要求四个 Sheet 的名字和结构保持一致。
4. 程序只读打开源文件、不更新外部链接，逐 Sheet 校验 F2 为 2027F；按 B 列 H40 读取 F 列数值；匹配工具中的 `H40_List`。
5. 弹窗显示四项各自读到的 H40 个数、当前预算之外的记录数、空白目标个数。确认 Yes 后更新 `H40_Targets` 对应 B:E。
6. 切换 `Budget_Detail!B3` 会自动读取各 H40 最新四项目标到 E6:E9。当前正在选中的 H40 也会同步刷新屏幕值，避免下次切换把刚导入的目标覆盖。

## 保留与异常规则

- H40 编码 `Trim` 去掉前后空格；也处理常见 NBSP 空格；不修改 Budget_Data 的原始 H40。
- 空白 F 值：不覆盖已有目标；F=0 是合法的明确目标值；非数字、负数或 Excel 错误的 F 值：程序中止本批导入。
- 同一源 Sheet 里，若出现**两个已经存在于本预算 H40_List 的相同 H40**，中止，防止错覆盖。
- 不在当前预算 H40_List 的额外行：跳过并计数。其它未导入的旧目标保留，不被清除。
- 读完全部四张表并检查无误后才进入确认与写入；目标值写到既有 `H40_Targets`，不修改 `Edit_Plan`、`Budget_Data`、`Budget_Trial`、`Budget_Overview` 月份台数。
- 本版**不自动覆盖最新的 Budget_Data**，也不迁移跨版本的旧 Edit_Plan 行号；这两件事涉及错行风险，必须独立验证。
- 所有真实公司预算只在本地 Excel 内处理；代码可以在公开 GitHub，无需上传包含真实数字的预算。
- 功能尚未在公司 Excel 实机验收。

## 验收

用公司允许的两个测试 H40（五位带空格与六位）先核对：四张 Sheet 中的 F 列 → `H40_Targets` 对应一行 B:E → `Budget_Detail!E6:E9` → G6:G9 0% 完成率。核对后再全量导入，检查现有预算改动没有丢失。
