# Excel 格式模板快速生成工具

一款基于 WPF + ClosedXML 的 Windows 桌面应用，实现类似"邮件合并"的 Excel 批量填充功能。支持 **Office Excel** 和 **WPS 电子表格** 的 `.xlsx` / `.xls` / `.et` 文件。

> 当前版本：**v1.1.2** · .NET 8 · ClosedXML 0.102 · [⬇ 下载最新版](https://github.com/mylxnet/excelmailmerge/releases/latest)

---

## ✨ 功能特性

### 核心功能
- **数据源解析**：加载 `.xlsx` / `.xls` / `.et` 文件（后两者自动转换为 `.xlsx`），自动识别列名、预览前 5 行数据，支持自定义标题行号
- **模板占位符**：模板中用 `{列名}` 标记占位符（如 `{姓名}`、`{工号}`），程序自动扫描匹配
- **两种输出模式**：
  - **多文件模式**：每行数据 → 一个独立工作簿（`.xlsx`）
  - **单文件模式**：全部行数据 → 一个工作簿，每行一个 Sheet
- **试生成预览**：生成全部前先试填第一条数据，让用户确认效果
- **智能命名规则**：主命名列 + 次命名列 + 自定义分隔符，重名自动追加序号
- **自动日期文件夹**：输出到 `输出结果yyyyMMdd_HHmmss` 子文件夹，永不覆盖

### 格式保真
- ✅ 保留模板行高、列宽、合并单元格、字体、颜色、边框、数字格式
- ✅ 生成后公式自动转为计算值（不保留公式表达式）
- ✅ 多 Sheet 模板完整支持，每个 Sheet 独立处理占位符

### 模板编辑
- 上传已有模板文件（支持 `.xlsx` / `.xls` / `.et`），不新建空白模板
- 左侧点数据源列名 → 自动复制 `{列名}` 到剪贴板
- 「📄 用 Excel/WPS 打开模板」按钮 → 获得完整格式预览，保存后程序自动重载
- 占位符自动高亮显示（黄色），预览区不显示行号列

### 稳定性与进程管理
- **单实例运行**：重复启动不会多开，自动唤起已有窗口（命名 `Mutex` + 命名管道）
- **启动清理残留进程**：启动前自动结束同名残留进程，避免文件被占用
- **生成中锁定窗口**：生成过程中「关闭」按钮禁用，避免误操作中断
- **大数据量保护**：单文件模式下若 `模板Sheet数 × 数据行数 > 200`，校验阶段直接报错并阻断生成
- **界面不卡死**：生成逻辑在线程池执行，进度条与实时日志即时刷新

---

## 🖥 界面预览

四步渐进式向导：

```
① 数据源  ──→  ② 模板  ──→  ③ 预览  ──→  ④ 生成
   |                              |              |
   ├ 加载 xlsx/xls/et             ├ 试生成一条    ├ 选择保存位置
   ├ 预览表格内容                 ├ 查看填充效果  ├ 选择输出模式
   └ 自动校验                     └ 确认无误     └ 选择命名规则
```

---

## 📦 安装方式

**下载地址**：https://github.com/mylxnet/excelmailmerge/releases/latest

### 方式一：自解压安装（推荐）
下载 `ExcelMailMerge_v1.1.2_Setup.exe`，双击按向导安装（自动创建桌面 / 开始菜单快捷方式）。

### 方式二：便携版
下载 `ExcelMailMerge_v1.1.2_Portable.zip`，解压后双击 `ExcelMailMerge.exe` 直接运行，无需安装。

> 📌 GitHub 上的附件名为英文（`..._Setup.exe` / `..._Portable.zip`），与本地中文原名 `Excel格式模板快速生成工具_v1.1.2_安装版.exe` / `..._v1.1.2.zip` 一一对应。

> **无需安装 .NET 运行时** — 自包含单文件（Self-contained Publish），大小约 **71 MB**。
>
> ⚠️ 若使用 `.xls` / `.et`（WPS）格式，需本机已安装 **Microsoft Excel 或 WPS 表格**（程序通过 COM 调用它们完成格式转换）；`.xlsx` 无此要求。

---

## 🛠 从源码构建

### 环境要求
- Windows 10/11 或 Windows Server 2019+
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)

### 构建命令
```powershell
# 克隆项目后进入目录
cd ExcelMailMerge_github

# 还原依赖
dotnet restore

# 开发模式运行
dotnet run

# 发布为 self-contained 单文件（win-x64）
dotnet publish -c Release -r win-x64 --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:EnableCompressionInSingleFile=true `
  -o ./publish
```

> ⚠️ 不要启用 `<InvariantGlobalization>true</InvariantGlobalization>`：WPF 不支持该模式，会导致 Release 版启动崩溃。

### 命令行测试
```powershell
# 生成测试数据（数据源 + 模板）
dotnet run -- --gen-test-files

# 端到端冒烟测试
dotnet run -- --smoke-test
```

---

## 📁 项目结构

```
ExcelMailMerge_github/
├── App.xaml / App.xaml.cs          # WPF 入口 + 单实例检测 + 残留进程清理 + 全局异常处理 + 命令行测试
├── MainWindow.xaml / MainWindow.xaml.cs   # 主界面（四步向导）
├── ExcelMailMerge.csproj          # 项目配置（.NET 8 / ClosedXML / Newtonsoft.Json / app.ico）
├── ServiceLocator.cs               # 简易 DI 服务定位器
├── app.ico                         # 应用图标
├── nuget.config                    # 包源配置
│
├── Models/                         # 数据模型
│   ├── DataSourceModels.cs         # ParsedDataSource / DataColumn / DataRow
│   ├── GenerationModels.cs         # GenerationProgress（INotifyPropertyChanged）/ GenerationResult
│   ├── TemplateAndValidationModels.cs  # ScannedTemplate / PlaceholderInfo
│   ├── AppSettings.cs              # 应用设置
│   └── WizardModels.cs             # 向导步骤枚举
│
├── Services/                       # 核心业务
│   ├── DataSourceService.cs        # Excel 数据源解析
│   ├── TemplateService.cs          # 模板扫描 + 占位符识别
│   ├── GenerationEngine.cs         # 填充引擎（核心：替换占位符 + 复制样式）
│   ├── FileFormatConverter.cs      # .xls/.et → .xlsx 转换（COM 互操作）
│   ├── NamingService.cs            # 文件/Sheet 命名规则 + 重名处理
│   ├── ValidationService.cs        # 数据源校验 + Sheet 数上限校验
│   └── SettingsService.cs          # 设置持久化
│
├── ViewModels/MainViewModel.cs     # MVVM 主视图模型
├── Views/TemplateEditorWindow.*    # 模板编辑窗口（外部 Excel/WPS 打开 + 自动重载）
├── Converters/CommonConverters.cs  # XAML 值转换器（占位符高亮 / InverseBool 等）
└── Helpers/                        # ObservableObject / RelayCommand / AsyncRelayCommand
```

---

## 🔧 技术栈

| 类别 | 技术 | 许可证 |
|------|------|--------|
| 语言 | C# 12 / .NET 8 | MIT |
| UI 框架 | WPF | MIT（.NET 基金会） |
| Excel 操作 | [ClosedXML 0.102](https://github.com/ClosedXML/ClosedXML) | MIT |
| JSON | Newtonsoft.Json 13 | MIT |
| 格式兼容 | COM 互操作（WPS / Microsoft Excel） | — |
| 架构 | MVVM（CommunityToolkit.Mvvm 风格，手动实现 ObservableObject） | — |

> 选择 ClosedXML 而非 NPOI / EPPlus 的原因：**完全免费商用（MIT）**、API 现代、与 Excel/WPS 兼容性好。

---

## 📄 许可证

**MIT License** — 可自由使用、修改、分发，包括商业用途。

---

## 🐛 问题反馈

如果程序崩溃或异常：

1. 查看 `%TEMP%\ExcelMailMerge_startup_error.txt`（启动崩溃）或 `%TEMP%\ExcelMailMerge_runtime_error.txt`（运行时崩溃）
2. 附上异常类型、消息和堆栈
3. 说明你的操作步骤（加载了什么数据源 / 什么模板 / 什么输出模式）

---

## 📝 更新日志

### v1.1.2（2026-09-17）
- 🐛 修复单文件模式行探查死循环（循环边界被循环体修改），每行耗时从 100+ 秒恢复正常
- 🐛 修复进度条不动 / 实时日志不刷新（`INotifyPropertyChanged` + `ObservableCollection`）
- 🐛 修复单实例锁（改为重试获取），解决"连开两次后进程数为 0"
- ⚡ 单文件模式性能优化：模板只打开一次
- ✅ 新增单文件模式 Sheet 数上限校验（`模板Sheet数 × 数据行数 > 200` 时报错阻断）
- ✅ 启动时自动清理残留进程，避免文件被占用
- ✅ 生成过程中禁用「关闭」按钮
- 🔧 数据源区按钮顺序调整（「浏览…」在左、「刷新Sheet」在右）
- 📄 新增测试报告 `TEST_REPORT.md`；csproj 版本号统一为 1.1.2

### v1.1.1（2026-09-17）
- 🔧 标题栏在署名前加版本号（`v1.1.1 「by Mr lin」`）

### v1.1.0（2026-09-17）
- ✅ 新增 `.xls` / `.et`（WPS）格式支持，自动转换为 `.xlsx`
- ✅ 新增应用图标 `app.ico`
- ✅ 新增单实例运行（重复启动唤起已有窗口）
- 🐛 修复输出列宽与模板不一致的问题
- 🔧 模板编辑器预览去掉左侧行号列；主界面不出现整页滚动条
- 🔧 数据源 / 模板路径不再记忆；校验失败不在选择模板时弹窗
- 🔧 标题栏加入署名「by Mr lin」

### v1.0.0（2026-08-27）
- 🎉 首个正式版本
- ✅ 四步渐进式向导 UI
- ✅ 多文件 / 单文件两种输出模式
- ✅ 占位符 `{列名}` 自动识别和高亮
- ✅ 试生成预览功能
- ✅ 文件 / Sheet 命名规则（主命名列 + 次命名列 + 分隔符 + 重名自动序号）
- ✅ 输出目录自动创建日期子文件夹
- ✅ 格式完整保留（行高、列宽、合并、字体、颜色、边框、数字格式）
- ✅ 公式自动转值，不保留公式表达式
- ✅ 支持 Office Excel 和 WPS 电子表格
- ✅ self-contained 单文件发布，用户机器无需 .NET 运行时
- ✅ WinRAR 自解压安装程序（安装到 ProgramFiles，自动创建桌面/开始菜单快捷方式）
