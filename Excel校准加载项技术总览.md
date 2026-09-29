# Excel 校准加载项技术总览

> 更新时间：2026-09-29  
> 文档定位：项目技术事实、运行方式、验证证据和后续技术计划的统一入口。  
> 状态说明：本文把“当前已实现/已验证”和“规划/待实现”分开描述；需求文档中的目标不能直接视为已交付能力。

## 1. 文档范围与来源

本总览基于项目原有技术文档、功能需求、重构设计、迭代需求、README、实机回归清单和同步规范整理而成。原始 Markdown 已清理，当前项目根目录只保留本技术总览和《Excel校准加载项需求汇总.md》两份统一文档。

## 2. 项目定位

这是一个面向 Excel 校准记录表的 Windows VSTO/COM 加载项。系统读取工作簿结构，识别校准模板和校准项区域，匹配本地或远端模板规则，解析标准值、误差、量程、技术要求及公式依赖，在当前工作簿状态下生成符合约束的测量值，并批量写回测量值单元格。

系统由以下工作流组成：

```text
工作簿快照
  -> 模板/字段识别
  -> 指纹和本地模板匹配
  -> 规则编辑、保存和同步
  -> 当前公式与参数解析
  -> 候选测量值生成
  -> 公式/技术要求验证
  -> Excel 事务写回和回滚
```

核心设计目标：降低人工填数和复核成本；沉淀可复用模板规则；本地优先保证现场可用；在不破坏公式、格式和原始值的前提下完成可追踪写入。

## 3. 技术栈和环境

| 类别 | 选型 |
| --- | --- |
| 语言/框架 | C#、.NET Framework 4.8 |
| Excel 集成 | VSTO、Excel COM Interop |
| UI | WinForms、VSTO Ribbon、CustomTaskPane |
| 本地存储 | SQLite，`System.Data.SQLite.Core` |
| JSON | `Newtonsoft.Json` |
| 本地接口 | `HttpListener` loopback HTTP 服务 |
| 测试 | MSTest、Core 纯逻辑测试 |
| IDE/构建 | Visual Studio 2022、MSBuild |
| 远端模板服务 | 生产 `https://www.wzglpt.top/api/excel-templates` |

运行前置条件：Windows 桌面版 Excel、Visual Studio 2022 Office/SharePoint 开发工具、.NET Framework 4.8 targeting pack、Visual Studio Tools for Office Runtime。

## 4. 解决方案结构

```text
ExcelCalibrationAddin.sln
├─ src
│  ├─ ExcelCalibrationAddin.Contracts
│  ├─ ExcelCalibrationAddin.Core
│  ├─ ExcelCalibrationAddin.Host
│  ├─ ExcelCalibrationAddin.Vsto
│  ├─ ExcelCalibrationAddin.LocalServer
│  └─ ExcelCalibrationAddin.AutomationBridge
├─ tests
│  └─ ExcelCalibrationAddin.Core.Tests
├─ tools
│  ├─ verify.ps1
│  ├─ install_debug_addin.ps1
│  ├─ reset_registration.ps1
│  ├─ start_local_backend.ps1
│  └─ create_excel_regression_workbook.mjs
└─ outputs\excel-regression
```

### 4.1 Contracts

纯模型和跨层契约，主要包括：

- `WorkbookSnapshot`、`SheetSnapshot`、`CellMeta`：工作簿、工作表、单元格快照。
- `TemplateFingerprint`：模板稳定身份和兼容匹配信息。
- `RecognizedField`、`TemplateRegionMapping`：字段和区域映射。
- `MeasurementRule`：识别、保存、加载和生成共同使用的规则载体。
- `MeasurementGenerationInput/Result`：生成器输入输出。
- `TemplateFieldDefinition`：表头、区域、单位、公式和条件分支定义。
- 样例数据、同步状态、生命周期和诊断模型。

当前模型覆盖面较完整，但 `Models.cs` 仍偏大，后续应按工作簿、模板、生成、同步和诊断领域拆分。

### 4.2 Core

不依赖 Excel COM 和 UI 的可测试核心：

- `MeasurementValueGenerator`、`MeasurementSeriesGenerator`：单点和多点候选值生成。
- `TemplateFingerprintBuilder`：从稳定结构生成指纹。
- `FieldMatcher`：字段和校准区域识别。
- `LocalTemplateRuleCacheRepository`：SQLite 缓存、匹配、冲突和 pending upload。
- `TemplateSyncClient`：远端 HTTP 客户端。
- `ConfigurationLoader`、`GenerationConfigurationStore`：配置默认值、迁移和克隆。
- `ParameterValueParser`、`NumberFormatInterpreter`：数值、百分比、单位和格式解析。

### 4.3 Host

业务用例、流程编排和 Excel 适配之间的连接层：

- `PluginBootstrapper`：组装配置、缓存、识别、同步、保存和生成依赖。
- `PluginWorkflowOrchestrator`、`VstoAddinFacade`：对 VSTO 暴露稳定入口。
- `TemplateRecognitionUseCase`、`SyncTemplateUseCase`、`GenerateMeasurementUseCase`：识别、同步、生成主流程。
- `MeasurementRuleDraftBuilder`、`MeasurementRuleParameterResolver`、`MeasurementRuleStructureAnalyzer`：规则草稿、参数和公式结构处理。
- `TemplateFieldDefinitionBuilder/Matcher`：模板字段定义和匹配。
- `ExcelInteropSnapshotProvider`、`ExcelInteropWriter`：快照和批量写入。

### 4.4 Vsto

Excel 加载项入口和交互：

- `ThisAddIn`：生命周期、Excel 事件、任务窗格和生成状态。
- `CalibrationRibbon`：Ribbon 命令。
- `CalibrationTaskPaneControl`：校准项、字段区域、模板操作和生成交互。
- `RandomGenerationConfigurationDialog`：随机配置。
- `TemplateLibraryManagerDialog`：模板库、同步和诊断包。
- `CloudLoginDialog`、`DailySyncScheduler`：云端登录和每日同步。
- `ExcelGenerationTransaction`：Excel 状态保护、提交和回滚。

### 4.5 LocalServer 和 AutomationBridge

`ExcelCalibrationAddin.LocalServer` 复用 SQLite，提供本机模板 HTTP 服务：

- `GET /api/excel-templates/health`
- `GET /api/excel-templates/list`
- `POST /api/excel-templates/match`
- `POST /api/excel-templates/save`

`AutomationBridge` 为影刀等外部自动化程序提供转发入口；VSTO 内嵌 `YingdaoAutomationServer` 监听 `127.0.0.1:30771`，提供 `health/status/generate`。

## 5. 核心流程

### 5.1 启动

```text
ThisAddIn_Startup
  -> 读取 appsettings.json
  -> ConfigurationLoader 恢复默认值/迁移旧配置
  -> PluginBootstrapper 组装依赖
  -> 初始化 Ribbon、TaskPane、事件路由和同步调度
```

业务对象在启动阶段组装；VSTO 负责生命周期和事件转发，业务规则放在 Host/Core。

### 5.2 识别和匹配

1. 以当前活动工作表打印区域为主要识别范围。
2. 捕获文字、原始值、显示值、公式、R1C1 公式、数字格式、合并区域和多级表头。
3. 识别校准项块、字段、数据行和行级映射。
4. 生成稳定指纹，忽略测量值、标准值、误差结果、技术要求当前结果和动态单位文本；工作簿未保存时从 Excel 实时读取包含空白格在内的完整合并布局，避免仅改变小数位后结构指纹漂移。
5. 本地精确/兼容匹配；本地没有强匹配时使用同步结果中的候选信息，不在生成路径访问远端。
6. 返回 `RecognitionAndSyncResult`，任务窗格展示匹配模板、识别字段和可编辑规则。

指纹只表达稳定结构：工作表身份、标题层级、字段相对位置、校准项区域、合并关系和必要公式结构。边框、颜色、行高、列宽不是正式强匹配依据。

### 5.3 模板草稿、保存和同步

识别结果经过 `MeasurementRuleDraftBuilder` 和 `TemplateFieldDefinitionBuilder` 固化为规则草稿。保存前清理当前工作簿瞬时测量值，校验指纹、字段、公式、映射和必填区域。

- 远端保存成功：以远端回传数据更新本地 SQLite。
- 远端不可用：写入本地 pending upload，不能丢失用户编辑。
- 同步状态和业务状态分开：业务状态为启用/停用/废止，同步状态为待上传/冲突/失败。
- 冲突由用户选择保留本地、使用远端或另存；废止模板保留状态，不静默删除。

### 5.4 生成和写回

```text
读取已匹配规则
  -> Excel 重算
  -> 定向批量读取标准值、技术要求、量程、误差/结论公式及依赖
  -> 解析行级约束
  -> 校验生成前置条件
  -> 内存中生成候选序列
  -> 用同一规则验证公式和技术要求
  -> 确定可写测量值单元格
  -> 批量写回
  -> Excel 重算并验证
```

生成只写测量值父表头下的数字列，不写 AVG、误差和结论；标准值为空的行保持不动；停用和非数值项目跳过。一个校准项是原子提交单位，任一关联行失败则整项恢复原值。

当前实现已包含预览工作流、格式/分辨力读取、趋势误差处理和公式验证拆分文件；“识别只保存位置、生成读取当前公式和结果”的完整重构仍属于目标设计，不能仅凭文档视为全部完成。

### 5.5 工作簿和公式依赖

- 识别扫描当前工作表打印区域；公式引用打印区域外的当前工作簿单元格时，按地址定向读取。
- 依赖范围包括当前工作簿其他工作表和隐藏工作表；不读取外部工作簿、网络文件或其他外部源。
- 支持普通引用、连续区域和多个不连续区域的内部模型。
- 静态命名区域可按实际区域读取；结构化引用、外部名称、`INDIRECT`、`OFFSET` 等不确定表达式必须记录原因并阻止对应生成。
- 生成前重算并重新读取，确保人工修改的标准值、量程和技术要求生效。

## 6. 识别和规则模型

### 6.1 校准项身份

无子标题时一级标题就是校准项；有子标题时子标题是校准项，一级标题是分组上下文。分组、名称和完整标题路径分开保存，同名项目允许重复。

同名实例重定位顺序：目标区域完全一致、列布局和高度一致、分组路径一致、规范化名称一致、实例顺序兜底。气体/通道/组分是附加上下文，不作为唯一身份。

### 6.2 按需字段识别

处理顺序为：校准项区域 -> 测量值 -> 误差 -> 技术要求 -> 公式解析 -> 由引用补充标准值、平均值和量程。平均值、量程、不确定度和结论不是所有项目的固定必填项；必填性由公式和项目类型决定。

误差搜索严格限制在当前校准项边界内，结合标题别名、通用词、单位行、公式引用、连续性、位置和数字格式。字段候选按语义优先级互斥：表头显式包含“技术要求/允许误差/MPE/限值”等技术要求语义时，即使同时包含 `%FS` 等单位，也不得进入误差候选；没有技术要求语义的独立 `%FS` 表头仍可作为误差候选。一个项目可以有多个不连续误差区域，内部作为一个逻辑字段保存。

### 6.3 合并单元格和行级映射

合并单元格按逻辑区域处理，不把合并覆盖的非左上角单元格误判为可写测量格。典型映射为：

```text
校准项
  -> 标准值行
     -> 标准值
     -> 测量值单元格组
     -> 平均值
     -> 误差
     -> 技术要求/MPE
     -> 不确定度
     -> 结论
```

技术要求合并区域按覆盖行分配；标准值为空的行不生成。目标需求还要求同一校准项多标准值序列保留顺序、重复测量子组和跨行上下文。

### 6.4 动态双行表头和多误差/MPE

第一行名称、第二行单位的父子层级和相对位置属于稳定结构，当前显示文字和动态单位结果不作为唯一身份。最新设计允许同一校准项拥有多个误差/MPE/结论约束：附属约束保留在同一 `MeasurementRule`，不得拆成独立校准项。横向或纵向布局都需按行/列重叠、距离、边界和表头关系唯一配对；数量不一致或并列歧义必须识别失败。

## 7. 生成模型和配置

### 7.1 行级统一规则

目标重构把生成规则统一为“当前行公式、技术要求、单位/量纲、精度和依赖值”的组合。项目名称只用于展示和诊断，不能决定算法类别。支持普通数值、`±`、`<`、`≤`、`>`、`≥`、百分比、`%FS/FS`、自定义格式和简单 `IF` 条件；不支持或歧义公式必须失败关闭。

生成和验证消费同一个行级规则对象。若公式依赖候选测量值，可在受控 Excel 试算中临时写入候选、重算、读取结果并无论成功失败都恢复；涉及宏、异步函数、外链或不可恢复副作用时使用隔离副本或拒绝。

### 7.2 配置层级

```text
全局配置
  -> 模板配置（逐参数覆盖或使用全局）
    -> 校准项配置（逐参数覆盖）
```

目标模型使用稀疏覆盖，区分“未设置”和“显式设置为默认值”，并保留参数来源。配置包括方向、正负允许误差、误差占比、重复测量波动、多标准值趋势、响应时间、分布、结果计算方式、标准值来源、测量值邻近比例和校准项误差倍数关系等。

精度优先从每个目标格当前数字格式解析；空白但有格式仍使用该格式；无格式或无法解析时回退 1 位小数。配置不能放宽工作簿当前技术要求。

### 7.3 跨校准项关联

同标准值自动关联按模板从上到下选择第一个兼容参照项；目标测量值可受邻近比例约束。用户也可显式指定 A -> B 的正数随机倍数区间，按标准值点抽取、同点重复测量复用，保留误差正负号。邻近模式和倍数模式互斥；关系环、单位不兼容、来源缺点和无可行交集都应在生成前诊断并阻止目标项写入。

## 8. 其他功能模块

### 8.1 样例数据

`SampleDataUseCase` 从已识别测量值区域采集数值，按模板指纹保存多个版本；选择对话框支持全选和逐项选择，版本列表支持查看和删除。样例仅用于分布参考，不能直接回填；非数值项目不采集。模板删除时默认清理样例，废止/覆盖时保留并标记来源版本。

### 8.2 多区域复制粘贴

`MultiAreaPositionTemplateStore` 保存多个不相连区域相对锚点的行列偏移、尺寸和公式，不保存业务值。目标可选择多个匹配区域或只选起始单元格；空间不足、列错位和结构不匹配必须报错。功能独立于随机数生成，并保持 Excel 撤销和异常回滚能力。

### 8.3 影刀和云端

内嵌接口：

- `GET http://127.0.0.1:30771/api/yingdao/health`
- `GET http://127.0.0.1:30771/api/yingdao/status`
- `POST http://127.0.0.1:30771/api/yingdao/generate`

配置 `Automation.Token` 后要求 `X-Excel-Calibration-Token`；带 `Origin` 的浏览器请求拒绝；生成用 `SemaphoreSlim(1,1)` 串行。影刀可通过 `tools/yingdao_excel.py` 等待加载项、工作簿、模板和生成规则就绪，默认最多约 60 秒。任务完成后按协议回传管理系统 `workbench/sync`。

云端登录通过 `CloudLoginDialog`，Token 使用 DPAPI `CurrentUser` 加密保存为 `%LOCALAPPDATA%\ExcelCalibrationAddin\cloud-session.dat`。同步失败只轻提示和记日志，不阻断本地生成。

## 9. 配置、数据和安全

- `src/ExcelCalibrationAddin.Vsto/appsettings.json`：后端、缓存、自动化端口、Token 和生成默认值。
- SQLite 默认位于 `%LOCALAPPDATA%\ExcelCalibrationAddin\cache.db`。
- 本机缓存、注册信息、`bin/obj` 和环境配置不通过 Git 同步。
- LocalServer 和影刀服务只监听 loopback；POST 必须为 `application/json`。LocalServer 请求体上限为 5 MB，影刀内嵌服务当前上限为 1 MB（需求约定为 5 MB，尚未对齐）。
- 不向客户端回传内部堆栈、路径、密码或 Token；日志不记录 Token/密码明文。
- 诊断包包含日志、指纹、识别 JSON、匹配结果和必要环境信息；日常日志保存摘要，不保存完整单元格快照。

## 10. 测试与验证

自动化测试覆盖指纹稳定性、偏移/表头/公式变化、字段和单位解析、随机数边界/方向/趋势、重复性、最大误差、模板缓存、同步、冲突、pending upload、行映射、合并单元格、公式验证和预览工作流。2026-09-29 实际执行结果为 202 项测试中 201 项通过、1 项失败；失败为附属 MAX-MIN 约束未限制生成测量值离散范围，不能将当前生成链路标为全绿。

构建和测试命令：

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\verify.ps1
dotnet test .\tests\ExcelCalibrationAddin.Core.Tests\ExcelCalibrationAddin.Core.Tests.csproj --configuration Debug
```

VSTO Release 安装包使用 `tools\publish_latest_addin.ps1` 生成。脚本会准备本机 ClickOnce Bootstrapper 的 `.NETFramework,Version=v4.8` 和 `Microsoft.VSTORuntime.4.0` 元数据，下载并校验 Microsoft Corporation 签名的 VSTO Runtime，再运行带 `/warnaserror` 的 Publish。2026-09-29 已验证 `outputs\vsto-publish-final\setup.exe`、`ExcelStandaloneComAddin.Vsto.vsto` 和 `Application Files` 均生成，发布日志无 warning/error；该结果只证明打包链路，不替代 Excel 实机安装和 EXCEL-01～08 回归。

真实 Excel 回归资产为 `outputs\excel-regression\Excel加载项回归样本.xlsx`，覆盖合并表头、多校准项、条件公式和无关工作表。必须验证：单工作簿全流程、多工作簿切换、打印区域、合并单元格、多级表头、条件分支、多项目、模板重载、异常回滚和关闭无残留进程。

重构指定的基准文件为 `C:\Users\LiMing\Desktop\五合一.xlsx`，须使用副本测试，记录原件哈希、公式/格式摘要和耗时；原件不得写入。另有 `紫外全项目.xlsx` 用于同一校准项多误差/MPE案例，均需在实际可访问时完成 Excel 回归。

## 11. 当前状态、维护热点和风险

已具备的主要能力：Core 生成、指纹、配置、SQLite 缓存、同步客户端、本地服务、样例数据骨架、多区域位置模板、影刀服务/桥接和大量纯逻辑测试。当前构建 0 错误但有 1 个 `MSB3178` 警告；测试 201/202 通过。VSTO 与 Excel COM、影刀和云端行为仍必须以真实环境回归为准。

重点大文件：

- `MeasurementRuleDraftBuilder.cs`：约 2178 行。
- `CalibrationTaskPaneControl.cs`：约 1652 行。
- `GenerateMeasurementUseCase.cs`：约 1240 行。
- `MeasurementRuleParameterResolver.cs`：约 1105 行。
- `ThisAddIn.cs`：约 879 行。

新增业务不应继续堆入这些文件。目标拆分目录包括 `Host/Recognition`、`Host/Templates`、`Host/Generation`、`Host/ExcelInterop`、`Vsto/TaskPane` 和 VSTO 事件/状态控制器。

主要技术风险：任意 Excel 公式无法安全本地求值；受控试算可能触发事件、外链或宏副作用；旧模板可能依赖已废弃的静态 MPE；`%FS` 和单位转换容易发生数量级错误；跨校准项关联可能出现环或无解；真实 Excel、撤销、COM 释放和任务窗格状态仍有回归风险。

## 12. 开发和交付规范

文件规模目标：普通业务类 300 行以内，复杂编排 500 行以内，UI 控件 500 行以内；超过 800 行原则上只允许缺陷修复或迁出逻辑。方法目标 40 行以内，超过 80 行必须拆分。

职责边界：UI 只展示和收集输入；Host 编排用例；Core 承载纯逻辑；Excel COM 只在适配层；生成不访问远端、不完整识别、不全表扫描、不逐格 COM 访问。

双机开发流程：从 GitHub `main` 拉取，功能分支使用 `feat/...`；开始开发前 `git pull --ff-only`；结束时只提交实际修改并 push；冲突使用 stash/rebase，禁止用 `reset --hard` 覆盖工作。项目仓库与 `project3-instrument-management` 仓库独立维护。

安装/更新当前 Debug 加载项可运行 `tools\install_debug_addin.ps1` 或项目根目录 `手动更新加载项.bat`；这些脚本会关闭 Excel，运行前须保存工作簿。

## 13. 后续技术路线

1. 完成真实 Excel 回归和失败资产留存。
2. 将 COM 缺陷转为 Host/Core 自动化测试。
3. 冻结行级规则、公式支持清单、失败粒度和性能基准。
4. 拆出合并单元格、行映射、技术要求解析、生成计划和公式验证。
5. 完成多气体、样例数据、多区域复制粘贴和影刀/云端链路验收。
6. 完成多误差/MPE约束全链路、稀疏配置继承和跨项关联。
7. 继续治理 UI 状态、本地服务生命周期、诊断导出和大文件规模。
