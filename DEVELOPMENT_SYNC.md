# 两台电脑同步开发规范

本项目与 `project3-instrument-management` 保持两个独立 Git 仓库。两台电脑都从 GitHub 的 `main` 分支获取代码；电脑时间不同步不会影响 Git 的提交顺序，提交历史以父子关系和提交哈希为准。

## 第一次使用另一台电脑

```powershell
git clone https://github.com/lm1285/ExcelCalibrationAddin.git
cd ExcelCalibrationAddin
git switch main
```

不要复制 `.git`、`bin`、`obj`、`.tmp-*` 或本机缓存。按 README 安装 Visual Studio、Office 开发工具和 .NET Framework 4.8 后再构建。

## 每次开始开发

确保当前没有未保存的修改，然后执行：

```powershell
git switch main
git pull --ff-only origin main
```

如需同时开发多个功能，从最新 `main` 建立功能分支：

```powershell
git switch -c feat/简短功能名
```

## 每次结束开发

```powershell
git status
git add <实际修改的文件>
git commit -m "说明本次完成的功能"
git push -u origin <当前分支>
```

如果只使用 `main`，最后一条命令可以改为 `git push origin main`。提交后再切换电脑，另一台电脑先执行本节“每次开始开发”。

## 出现冲突或无法快进

不要使用 `reset --hard` 覆盖工作。先保留修改并查看状态：

```powershell
git status
git stash push -u -m "临时保存"
git pull --rebase origin main
git stash pop
```

解决冲突后执行 `git add`、`git rebase --continue`，通过测试后再 push。已经 push 的提交不要改写历史。

## 本机文件边界

`src/ExcelCalibrationAddin.Vsto/appsettings.json`、SQLite 缓存、Excel 注册信息和构建输出属于本机环境，不作为两台电脑之间的进度同步内容。云端模板和业务数据由管理系统后端保存，不通过 Git 同步。

## 管理系统仓库

管理系统使用独立仓库 `https://github.com/lm1285/project3-instrument-management.git`，流程完全相同，只需把 clone 地址和目录替换为该仓库。两个仓库应分别提交、分别 push；不要把一个仓库嵌套到另一个仓库中。
