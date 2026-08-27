# PROJECT_STATE.md — Excel 格式模板快速生成工具

> **这份文件的作用**：任何时候、任何电脑上开一个新的 Trae 会话，让 AI 先读这个文件就能快速接上项目。
>
> 读完此文件，AI 应该能：理解项目现状、知道哪些坑踩过、不重复造轮子、直接开始下一个任务。

---

## 0. 快速导航

| 要知道什么 | 看哪一节 |
|------------|----------|
| 项目是什么、能干嘛 | §1 项目概览 |
| 现在版本、哪些功能已完成 | §2 当前状态 |
| 踩过哪些坑、怎么修的 | §3 已修复 Bug |
| 关键设计决策 & 为什么 | §4 架构决策 |
| 有哪些已知限制 | §5 已知问题 |
| 下一步建议做什么 | §6 待办清单 |
| 怎么构建/发布 | §7 构建发布 |
| 给新 AI 的快速上手提示 | §8 交接提示 |

---

## 1. 项目概览

**名称**：Excel 格式模板快速生成工具  
**版本**：v1.0.0（2026-08-27）  
**一句话描述**：基于 ClosedXML 的 Excel 邮件合并工具 — 数据源 xlsx + 模板 xlsx → 批量填充占位符 `{列名}` → 生成一个工作簿或多个工作簿  
**技术栈**：C# 12 / .NET 8 / WPF / ClosedXML 0.102 / Newtonsoft.Json 13  
**许可证**：MIT  
**目标用户**：需要批量生成个性化 Excel 文档的办公人员（员工信息卡、合同模板、批量报表等）  
**输出模式**：
- **多文件（MultiFilePerRow）**：每行数据 → 一个独立 .xlsx 文件
- **单文件（SingleFileMultiSheet）**：所有行 → 一个 .xlsx 文件，每行一个 Sheet

---

## 2. 当前状态

### ✅ 已完成功能

| 模块 | 文件 | 状态 |
|------|------|------|
| 数据源加载 + 预览 | `Services/DataSourceService.cs` + `MainWindow.xaml` | ✅ 支持选 Sheet、自定义标题行号、DataGrid 动态列 |
| 数据源校验 | `Services/ValidationService.cs` | ✅ 检测合并单元格、空标题行、列名重复等 |
| 模板扫描 + 占位符识别 | `Services/TemplateService.cs` | ✅ `{列名}` 正则匹配、高亮显示、匹配状态提示 |
| 试生成预览 | `Services/GenerationEngine.cs::GeneratePreview` | ✅ 生成后读回填充值展示在 DataGrid |
| 全部生成 | `Services/GenerationEngine.cs::GenerateAsync` | ✅ 多文件/单文件、进度报告、失败重试 |
| 命名规则 | `Services/NamingService.cs` | ✅ 主命名列 + 次命名列 + 分隔符 + 重名自动追加序号 |
| 自动日期子文件夹 | `GenerationEngine.cs::GenerateAsync` | ✅ `输出结果yyyyMMdd_HHmmss` |
| 格式完整保留 | `GenerationEngine.cs::CopySheetContentAndStyle` | ✅ ClosedXML `Range.CopyTo` 一次性复制值+样式+合并区域 |
| 公式转值 | `GenerationEngine.cs::ConvertFormulasToValues` | ✅ 扫描所有公式单元格，保留计算值，清空公式表达式 |
| 模板编辑器（外部 Excel/WPS） | `Views/TemplateEditorWindow.*` | ✅ 方案 B：点列名复制剪贴板 + 用 Excel/WPS 打开 + FileSystemWatcher 自动重载 |
| 四步向导 UI | `MainWindow.xaml` | ✅ 数据源 → 模板 → 预览 → 生成（渐进解锁） |
| 文件名预览实时刷新 | `MainViewModel.cs::RefreshFileNamePreview` | ✅ 单文件显示"合并结果_xxx.xlsx | Sheet:..."、多文件显示每行文件名 |
| 保存文件夹浏览 | `MainViewModel.cs::BrowseFolder` | ✅ OpenFileDialog 选文件取路径 |
| self-contained 单文件发布 | csproj + publish 参数 | ✅ 免 .NET 运行时，71MB |
| WinRAR 自解压安装 | `_installer_stage/setup.bat` | ✅ ProgramFiles 安装 + 桌面/开始菜单快捷方式 + 注册表卸载 |
| 异常捕获 + 日志 | `App.xaml.cs` + DispatcherUnhandledException | ✅ 启动崩溃 → `%TEMP%\ExcelMailMerge_startup_error.txt`、运行时崩溃 → `%TEMP%\ExcelMailMerge_runtime_error.txt` |
| 命令行测试 | `App.xaml.cs` | ✅ `--gen-test-files` 生成测试数据、`--smoke-test` 端到端冒烟测试 |

### 📦 已产出物

| 位置 | 说明 |
|------|------|
| `publish/ExcelMailMerge.exe` | self-contained 单文件（71MB） |
| `Excel格式模板快速生成工具_v1.0.0_安装版.exe` | WinRAR 自解压安装程序（64.72MB） |
| `Excel格式模板快速生成工具_v1.0.0.zip` | 便携版（65.46MB） |
| `Archive_v1.0.0_*/` | 源码存档（Git 备份） |
| `ExcelMailMerge_github/` | GitHub 仓库源码（29 文件 / 5039 行） |

---

## 3. 已修复 Bug（**重要：踩过的坑不要重复踩**）

### BUG-01：WPF 绑定死锁导致 Release 版崩溃
- **现象**：Debug 版正常，Release 版启动即崩溃，提示 `System.InvalidOperationException: Cannot find non-neutral culture related to 'en-us'`
- **根因**：csproj 加了 `<InvariantGlobalization>true</InvariantGlobalization>`（为了"精简多语言"），**但 WPF 不支持 invariant globalization 模式**。self-contained publish 时 WPF 的 `PresentationCore.resources.dll` 被一起裁掉，`XmlLanguage.GetSpecificCulture()` 崩溃
- **修复**：**删掉** `<InvariantGlobalization>true</InvariantGlobalization>`。多语言资源总共才 ~1MB，不值得优化掉
- **教训**：任何 .NET WPF 项目 **不可** 启用 invariant globalization

### BUG-02：冒烟测试死锁
- **现象**：`Task.Run(engine.GenerateAsync(...)).Result` 在主线程调用时永远卡住
- **根因**：`GenerateAsync` 内部有 `await Task.Yield()`，捕获了 WPF Dispatcher SynchronizationContext，`.Result` 在主线程请求时死锁
- **修复**：用 `Task.Run(() => engine.GenerateAsync(...)).GetAwaiter().GetResult()` 把整个调用包在线程池（无 SyncContext）上
- **教训**：WPF 应用里 `await` + `.Result` 会死锁。要么用 `Task.Run` 隔离，要么全 async/await

### BUG-03：DataGrid 显示 `(Collection)`
- **现象**：数据源预览 DataGrid 里每个单元格显示 `(Collection)` 而不是实际值
- **根因**：`DataRow.Values` 是 Dictionary，AutoGenerateColumns 无法展开显示
- **修复**：给 DataRow 加索引器 `public object this[string col] => Values.TryGetValue(col, out var v) ? v : ""`，代码动态生成 DataGrid 列

### BUG-04：ProgressBar TwoWay 绑定只读属性
- **现象**：生成阶段弹异常 `Cannot perform TwoWay or OneWayToSource binding on read-only property 'ProgressPercent'`
- **根因**：`ProgressPercent` 是 `=>` 表达式体只读属性，但 WPF `ProgressBar.Value` 默认 `Mode=TwoWay`
- **修复**：XAML 里显式加 `Mode=OneWay`

### BUG-05：右/下边框丢失
- **现象**：FillTemplate 后，模板最右列的右边框和最末行的下边框丢失
- **根因**：① 手动逐格复制 Style + 合并单元格冲突；② `RangeUsed.LastColumn()` 不包含"有边框无值"的单元格
- **修复**：① 用 `srcRange.CopyTo(dst.FirstCell())` 一次性复制 ClosedXML Range（比手动循环可靠 N 倍）；② LastRow/LastCol 各 +1 扩展
- **教训**：ClosedXML 的 `Range.CopyTo` 是首选，永远优先于手动逐格循环

### BUG-06：单文件模式 Sheet 名带模板名后缀
- **现象**：用户选"张三"作为 Sheet 名，实际生成"张三_信息卡"（"信息卡"是模板 Sheet 名）
- **根因**：`NamingService.EnsureUniqueSheetName` 硬编码把模板 Sheet 名拼到用户命名后面
- **修复**：直接用用户选的命名，重名时只加 `_2`、`_3`

### BUG-07：单文件模式文件名预览错误
- **现象**：预览显示用户数据拼接的文件名（"张三_李四.xlsx"），实际生成"合并结果_时间戳.xlsx"
- **根因**：Preview 逻辑是拍脑袋编的，没读 GenerationEngine 的真实命名规则
- **修复**：Preview 必须忠实反映 GenerationEngine.cs 里硬编码的实际规则
- **教训**：Preview 代码和实际逻辑 **必须** 从同一个数据源派生，不能写两份

### BUG-08：XAML ElementStyle TargetType 错误
- **现象**：`Style(typeof(TextBox))` 赋给 `DataGridTextColumn.ElementStyle` 抛 `TextBox TargetType 与 TextBlock 类型不匹配`
- **根因**：ElementStyle 是"显示模式"样式，目标是 `TextBlock`；只有 `EditingElementStyle` 才是 `TextBox`
- **修复**：`Style(typeof(TextBlock))`

---

## 4. 架构决策

### 4.1 为什么选 ClosedXML 而不是 NPOI / EPPlus？

| 库 | 许可证 | 说明 |
|----|--------|------|
| **ClosedXML** | MIT | ✅ 完全免费商用、API 现代、xlsx only |
| NPOI | Apache 2.0 | 维护不活跃，0.21.0 后多年没发布 |
| EPPlus | Polyform Noncommercial | 5+ 版本商业公司需付费 |
| DocumentFormat.OpenXml | MIT | OpenXML 原生态，但 API 极其底层，写起来像 XML DOM |

**最终选 ClosedXML 0.102**：MIT + 活跃维护 + API 优雅 + 与 Office/WPS 兼容性好。

### 4.2 为什么 WPF 而不是 WinForms / Electron？

| 维度 | WPF | WinForms | Electron |
|------|-----|----------|----------|
| 现代 UI 能力 | ⭐⭐⭐⭐⭐ | ⭐⭐ | ⭐⭐⭐⭐⭐ |
| 包体积 | 71MB self-contained | ~50MB | ~150MB+ |
| 原生 .NET 生态 | ✅ | ✅ | ❌ |
| Office/WPS 兼容性 | ✅（不需要装 Office） | ✅ | ✅ |
| 技术门槛 | 中 | 低 | 高（前端栈） |

**选 WPF + .NET 8**：UI 够现代、原生 .NET 便于用 ClosedXML、self-contained 单文件方便分发。

### 4.3 模板编辑器为什么用"外部 Excel/WPS"而不是高仿 DataGrid？

**放弃方案**：ContentControl + Grid 逐格高仿渲染（已写过 600+ 行，最后废弃）

| 高仿方案的硬伤 | 外部编辑方案的优势 |
|----------------|------------------|
| DataGrid 不支持 RowSpan/ColSpan | Excel/WPS 自己渲染 = 100% 正确 |
| 行高列宽字体颜色每一项都要大量代码 | 完整格式预览（图片/图表/条件格式/冻结窗格全部支持） |
| 合并单元格边框被覆盖丢失 | ClosedXML CopyTo 一次性复制，不丢任何东西 |
| 代码量 ~800 行 | 代码量 ~200 行 |
| 用户体验差（"内容大面积不显示"） | 用户习惯的 Excel/WPS 界面 |

**当前方案**：
1. 左侧数据源列名 → 点击复制 `{列名}` 到剪贴板（最稳定，零 Win32 互操作）
2. 右侧按钮「📄 用 Excel/WPS 打开模板」→ 用户在 Excel 里 Ctrl+V 粘贴
3. FileSystemWatcher 监听模板文件变更 → 自动重载预览

### 4.4 依赖注入：为什么不引第三方 DI 容器？

- 服务只有 7 个，层级简单（`GenerationEngine → NamingService + TemplateService`）
- 用 `ServiceLocator.cs` 静态类手写，7 行代码搞定
- 引一个 Microsoft.Extensions.DependencyInjection 带来 ~30 个传递依赖，没必要

### 4.5 self-contained + single-file + EnableCompressionInSingleFile

这三个参数组合决定了分发方式：
- `--self-contained true`：用户机器**不需要装 .NET 运行时**（关键卖点）
- `PublishSingleFile=true`：单文件，不是一堆 DLL
- `EnableCompressionInSingleFile=true`：exe 内部压缩，从 100MB → 71MB

**代价**：启动慢 ~1-2 秒（自解压），进程常驻内存 250MB+。对于"双击即用"的桌面应用完全可接受。

---

## 5. 已知问题（**不要去改"修复"它们，是特性限制**）

| # | 问题 | 原因 | 建议 |
|---|------|------|------|
| 1 | 多 Sheet 模板修改只能在 Excel/WPS 里做 | 我们的模板编辑器极简版不支持切 Sheet 编辑 | 直接用外部 Excel/WPS，让 FileSystemWatcher 自动重载 |
| 2 | 占位符 `{列名}` 区分大小写 + trim | 设计决策：精确匹配，不做模糊 | 用户写模板时注意列名大小写和空格 |
| 3 | `.xls`（老格式 Office 2003）不支持 | ClosedXML 只支持 `.xlsx` | 提示用户另存为 .xlsx |
| 4 | 图片/图表/批注/条件格式 | 用 `Range.CopyTo` 全部保留 | 正常工作，不需要额外代码 |
| 5 | 公式自动转值，不保留 | 设计决策（用户要求："公式不保留，直接输入公式计算的值"） | 用户知道这是预期行为 |
| 6 | 数据源不支持合并单元格 | 设计决策 + ClosedXML 读合并会把所有子格值合并到左上角 | 校验阶段自动检测并提示用户修正 |
| 7 | 设置文件存 `%LOCALAPPDATA%\ExcelMailMerge\settings.json` | 硬编码路径 | 如需改位置需改 SettingsService.cs |

---

## 6. 待办清单（按优先级排序）

> 以下是合理的下一阶段方向，**不是全部要做**。AI 接手后应该问用户想先做什么，而不是全部开搞。

### P0（高价值）
- [ ] **图标/Logo**：csproj 里 `<ApplicationIcon />` 是空的，需要一个 .ico 文件（README 里提到的"快速生成工具"图标）
- [ ] **版本号统一**：csproj 里 `Version/AssemblyVersion/FileVersion` 都是 1.0.0，需要建立 semver 升级机制
- [ ] **测试数据生成命令** `--gen-test-files` 依赖 ClosedXML 字体测量 AdjustToContents，可能失败 → 已用固定列宽规避，稳定

### P1（中价值）
- [ ] **更多命名规则选项**：当前只有"主+次+分隔符"，可加"日期时间"、"前缀/后缀"、"自动流水号"
- [ ] **输出文件冲突策略**：当前"自动追加日期子文件夹"，可加"覆盖 / 跳过 / 追加序号"选项
- [ ] **模板编辑器增强**：左侧列表支持搜索过滤（当前有 SearchBox 但逻辑未验证）
- [ ] **生成进度取消**：GenerationEngine.GenerateAsync 没传 CancellationToken，用户中途无法取消

### P2（低价值 / 锦上添花）
- [ ] **多语言 UI**：当前全中文硬编码（InvariantGlobalization 已删，有条件做 i18n）
- [ ] **主题切换**：当前只有亮色主题（#1A2540 深蓝色顶栏 + #D4A256 金色按钮）
- [ ] **命令行完整支持**：除了 smoke-test，可以加 `--generate <dataSource> <template> <outputDir>` 脚本模式
- [ ] **PDF 导出**：模板填充后可一键导出 PDF（需加 ClosedXML → Excel COM → PDF 的链路，用户机器需装 Office/WPS）
- [ ] **生成完成后自动打开输出文件夹**（Windows Explorer）

### P3（重构 / 清理）
- [ ] 删掉 `项目设计方案.md`（已完成使命），或保留作为架构文档
- [ ] `Services/` 里 `SettingsService` 只被 MainViewModel 用 → 可以考虑合并
- [ ] `Views/TemplateEditorWindow` 当前功能非常轻（只打开外部程序 + 监听变更）→ 考虑合并到 MainWindow 第 2 步的"预览 + 打开"按钮里，省掉一个独立窗口

---

## 7. 构建发布

```powershell
# === 开发 ===
cd ExcelMailMerge_github
dotnet restore
dotnet run                                    # 启动 WPF 界面
dotnet run -- --gen-test-files                # 生成测试数据
dotnet run -- --smoke-test                    # 端到端冒烟测试

# === 发布 ===
dotnet publish -c Release -r win-x64 `
  --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:EnableCompressionInSingleFile=true `
  -o ./publish

# === WinRAR 自解压安装 ===
RAR a -sfx -z"sfx.cfg" -ep1 -r -m5 `
  "Excel格式模板快速生成工具_v1.0.0_安装版.exe" `
  "Excel格式模板快速生成工具\"
# sfx.cfg 内容：Setup=setup.bat / TempMode / Silent=2 / Overwrite=1
```

**必须用参数组合**：PublishSingleFile + IncludeNativeLibrariesForSelfExtract + EnableCompressionInSingleFile  
**不能加**：InvariantGlobalization=true（会破坏 WPF 绑定）

---

## 8. 交接提示（给新 AI）

```
你好，我是这个项目的 AI 助手。
项目：Excel 格式模板快速生成工具
版本：v1.0.0（2026-08-27）

请先做以下事情：
1. 读 ExcelMailMerge.csproj 确认依赖版本
2. 跑 dotnet build 确认能编译（0 错误）
3. 跑 dotnet run -- --smoke-test 确认端到端通过
4. 看 Services/GenerationEngine.cs 的 GenerateAsync 和 CopySheetContentAndStyle
   （这两个是核心，理解它们就能理解整个填充机制）
5. 看 MainWindow.xaml 的四步向导布局
6. 看 PROJECT_STATE.md 的 §3 已修复 Bug（避免重复踩坑）

准备好后，问用户："接下来想先做什么？"
```

---

## 9. 快速定位指南

| 想做什么 | 从哪里看起 |
|----------|------------|
| 改/加命名规则 | `Services/NamingService.cs` |
| 改输出格式（多文件/单文件） | `Models/GenerationModels.cs` → `OutputMode` 枚举 |
| 改文件命名前缀/后缀/日期格式 | `Services/GenerationEngine.cs::GenerateAsync` |
| 改占位符正则 | `Services/TemplateService.cs`（约第 40 行 `PlaceholderRegex`） |
| 改 UI 布局 | `MainWindow.xaml`（四步 Wizard 在同一个文件里） |
| 改颜色/主题 | `MainWindow.xaml` 顶部 `<Window.Resources>`（PrimaryBtn/AccentBtn 等） |
| 改数据源校验规则 | `Services/ValidationService.cs` |
| 改设置持久化 | `Services/SettingsService.cs` + `Models/AppSettings.cs` |
| 读/写 xlsx 单元格值/样式 | ClosedXML API（`IXLWorksheet.Cell(row, col)`、`IXLCell.Style`） |
| 调试崩溃 | 看 `%TEMP%\ExcelMailMerge_startup_error.txt` 和 `runtime_error.txt` |
