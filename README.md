# Excel COM Add-in

This directory is the isolated Excel add-in workspace for this project. It contains only the Excel COM/VSTO add-in source code and its supporting C# class libraries.

## Scope

- `src/ExcelCalibrationAddin.Vsto`: Excel COM/VSTO entry point, ribbon, task pane, dialogs, and Excel interop wiring.
- `src/ExcelCalibrationAddin.Host`: add-in workflow orchestration and Excel-facing adapters.
- `src/ExcelCalibrationAddin.Core`: random value generation, template recognition support, local cache, and configuration services used by the add-in.
- `src/ExcelCalibrationAddin.Contracts`: DTOs, domain models, and shared contracts for the add-in.
- `src/ExcelCalibrationAddin.LocalServer`: loopback-only HTTP template service backed by the same SQLite cache.
- `tests/ExcelCalibrationAddin.Core.Tests`: pure logic, workflow, persistence, and HTTP-client regression tests.
- `tools`: build, install, reset, and verification scripts for this isolated add-in workspace.

This folder intentionally does not include the Python desktop application or packaged executables. The `outputs\excel-regression` directory contains the maintained Excel regression sample for this add-in.

## Requirements

- Windows desktop Excel.
- Visual Studio 2022 with Office/SharePoint development tools.
- .NET Framework 4.8 targeting pack.
- Visual Studio Tools for Office Runtime.

## Build

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\verify.ps1
```

or build the solution directly:

```powershell
& (.\tools\find_vs_msbuild.ps1) .\ExcelCalibrationAddin.sln /t:Restore /p:Configuration=Debug /p:Platform="Any CPU"
& (.\tools\find_vs_msbuild.ps1) .\ExcelCalibrationAddin.sln /t:Build /p:Configuration=Debug /p:Platform="Any CPU"
```

Run the pure logic tests with:

```powershell
dotnet test .\tests\ExcelCalibrationAddin.Core.Tests\ExcelCalibrationAddin.Core.Tests.csproj --configuration Debug
```

## Excel Regression

The maintained real-Excel regression assets are:

- `Excel实机回归清单.md`: execution steps, expected results, and issue-record requirements.
- `outputs\excel-regression\Excel加载项回归样本.xlsx`: representative workbook covering merged headers, multiple calibration items, conditional formulas, and an unrelated worksheet.
- `tools\create_excel_regression_workbook.mjs`: reproducible workbook builder used by the development environment.

Automated tests do not replace this checklist. Workbook switching, task-pane state, undo, rollback, and COM release must still be verified in desktop Excel.

## Local Template Backend

The add-in can use a local HTTP backend backed by the same SQLite database as the VSTO client:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\start_local_backend.ps1
```

生产环境默认 endpoint: `https://www.wzglpt.top/api/excel-templates`。加载项启动后会从本机 DPAPI 会话存储读取云端 Token；首次使用请在 Excel“校准助手”功能区点击“云端登录”，使用管理系统账号登录。

本地开发仍可在 `src\ExcelCalibrationAddin.Vsto\appsettings.json` 中将 `Backend.BaseUrl` 改为 `http://localhost:3002`。

The bundled server only listens on an HTTP loopback address. POST endpoints require
`application/json`, reject browser-originated requests, and limit request bodies to 5 MB.
It is intended for the add-in's `HttpClient`, not for direct browser access.

Supported routes:

- `GET /api/excel-templates/health`
- `GET /api/excel-templates/list`
- `POST /api/excel-templates/match`
- `POST /api/excel-templates/save`

The add-in performs a background template synchronization on startup when the last successful sync is older than one day. The template library window also provides manual synchronization and diagnostic package export.

## Shadow Knife Automation

影刀调用 Excel 加载项的本机接口：

- `GET http://127.0.0.1:30771/api/yingdao/health`
- `GET http://127.0.0.1:30771/api/yingdao/status`
- `POST http://127.0.0.1:30771/api/yingdao/generate`

影刀推荐调用 `tools\yingdao_excel.py` 中的模块 `EXCELjzx` 函数
`generate_after_excel_open`。函数会由模块监视加载项状态，每隔
`interval_seconds` 秒检查一次，最多检查 `retries` 次；只有确认加载项已加载、目标工作簿已打开、模板已匹配且存在可生成规则后，才调用生成接口。
默认参数为 `retries=60`、`interval_seconds=1`，即最多等待约 60 秒。

当 `Automation.Token` 非空时，请在请求头发送 `X-Excel-Calibration-Token`。影刀任务完成后可将结果回传管理系统：
`POST https://www.wzglpt.top/api/shadow-knife-linkage/workbench/sync`，并在 `X-Shadow-Knife-Key` 中发送服务端配置的 `SHADOW_KNIFE_WEBHOOK_KEY`。

The database path comes from `Cache.SqliteFile` in `src\ExcelCalibrationAddin.Vsto\appsettings.json`; by default it is `%LOCALAPPDATA%\ExcelCalibrationAddin\cache.db`.

## Install Current Debug Add-in

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\install_debug_addin.ps1
```

The script automatically closes all Excel processes, removes this project's old VSTO
registrations and ClickOnce cache entries, cleans/restores/builds the solution, then waits
for VSTO installation and verifies the current Excel registration. Any unsaved Excel work
will be lost, so save it before running the script.

```text
src\ExcelCalibrationAddin.Vsto\bin\Debug\ExcelStandaloneComAddin.Vsto.vsto
```

## Reset Registration

If Excel keeps an old ClickOnce/VSTO registration, run (it closes Excel automatically):

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\reset_registration.ps1
```

## Manual Update After Each Change

After completing any project modification, double-click the project-root batch file:

```text
手动更新加载项.bat
```

It closes Excel, removes this project's old registrations and ClickOnce cache, performs a
clean restore/build, installs the current VSTO manifest, and verifies the Excel registration.
It may discard unsaved Excel work, so save workbooks before running it.
