# WPS Office COM 自动化技术简报（给 Agent 技能包用）

> 调研日期：2026-09-28
> 目的：为「WPS 办公自动化 Skill 包」（Claude Code / Codex 等编程 Agent 用）提供技术底座。
> 约定：标 ✅ 为多源交叉验证；标 ⚠️ 为单源或社区经验、建议真机复测后再写入 SKILL.md。

---

## 0. TL;DR

- 三个 ProgID：Writer = `KWPS.Application`，Spreadsheets = `KET.Application`，Presentation = `KWPP.Application`（✅ 三源以上确认）。
- 对象模型与 MS Office COM **高度兼容**：`Documents/Workbooks/Presentations`、`Range/Cells`、`Paragraphs/Tables/Shapes` 几乎照抄 Word/Excel/PPT，Word 的 VBA 代码大多能直接跑。
- 个人版免费 COM 可用（社区共识 ✅）；**VBA 编辑器**才需要会员/专业版；**JS 宏（JSA）个人版免费**，是 VBA 的替代路线。
- `Visible=$false` + `DisplayAlerts=0` 可以后台静默跑；必须 `Quit()`，否则进程残留。
- WPS 个人版是 **32 位**：64 位宿主调 COM 找不到 ProgID 时，换 32 位 PowerShell。
- 免费版导出 PDF 可能带 "WPS Office" 水印（⚠️ 社区说法，需真机验证）。
- 头less/服务端：WPS 是桌面应用，COM 需要已登录的交互式桌面会话（⚠️ 与 MS Office 同类约束，WPS 专用资料未找到）。

---

## 1. ProgID（已多源验证 ✅）

| 组件 | ProgID | 对应 MS Office | 验证来源 |
|---|---|---|---|
| WPS 文字（Writer） | `KWPS.Application` | `Word.Application` | hch135861/wps-office, yb2460/harness-anything, neomei/wpscomposer, xiaoqiong0v0/opencode-skills |
| WPS 表格（Spreadsheets） | `KET.Application` | `Excel.Application` | 同上 + CSDN（`Type.GetTypeFromProgID("Ket.Application")` 实测） |
| WPS 演示（Presentation） | `KWPP.Application` | `PowerPoint.Application` | 同上 |

补充说明：

- ProgID 不区分大小写：`Kwps.Application` / `kwps.application` 都能用，社区代码两种写法混用。✅
- 旧版兼容 ProgID：`WPS.Application`（Writer）、`Et.Application`（表格）、`Wpp.Application`（演示）。xiaoqiong0v0 的技能文档写明："首选 `Kwps.Application`（WPS 12.x+），备选 `WPS.Application`（旧版兼容）"。⚠️
- ⚠️ 注意网上有 CSDN 文章把 `KWPS.Application` 写成"表格的 ProgID"——那是笔误，KWPS 是文字。
- ⚠️ 专业版 / 政务版 / 教育版安装后 ProgID 可能不同（xiaoqiong0v0 提示），建议技能里做运行时探测：先试 `KWPS.Application`，失败再试 `WPS.Application`。
- **身份混淆坑**：neomei/wpscomposer 的真机验证报告发现，在装有 MS Office + WPS 的机器上，"64 位 Word 注册指向 WINWORD.EXE，32 位注册指向 WPS"——也就是说 `Word.Application` / `Excel.Application` 到底打开谁，取决于注册表视图和安装顺序。**技能里必须显式用 WPS 的 ProgID，不要偷懒用 `Word.Application` 指望它落到 WPS 上。** ✅（真机验证来源）
- 探测命令（技能安装检查用）：

```powershell
# 检查 ProgID 是否注册
Get-ItemProperty "HKLM:\Software\Classes\KWPS.Application" -ErrorAction SilentlyContinue
Get-ItemProperty "HKLM:\Software\Classes\KET.Application" -ErrorAction SilentlyContinue
Get-ItemProperty "HKLM:\Software\Classes\KWPP.Application" -ErrorAction SilentlyContinue
```

---

## 2. 对象模型兼容性

WPS COM 对象模型是对 MS Office 的高仿真复刻（✅ 多源一致）：

```
Application
├── Documents → Document → Paragraphs/Tables/Sections/Styles/Content/Selection → Range → Font/ParagraphFormat
├── Workbooks → Workbook → Worksheets → Worksheet → Range/Cells → Font/Interior/Borders / Shapes → Chart
└── Presentations → Presentation → Slides → Slide → Shapes → TextFrame/TextRange / Tables / Charts
```

已知差异 / 坑（✅ 来自 iheng88/claude-office-plugin 的 JSAPI 兼容表 + sueccku 的 WPS skill 坑记录，原理相通，COM 同样适用）：

| MS Office 写法 | WPS 实际行为 | 建议 |
|---|---|---|
| 集合直接索引 `Worksheets(1)` | ⚠️ 可能不工作 | 统一用 `.Item(1)`：`$wb.Worksheets.Item(1)` |
| `Font.Bold == True` | ⚠️ WPS 可能返回 `0/-1/True/False` 任一种 | 用 truthy 判断，不用 `== $true` |
| 样式名 `"标题1"` | ⚠️ WPS 中文样式名带空格 | 用 `"标题 1"`、`"标题 2"`、`"正文"` |
| `.Borders` / `.BorderAround()` | ⚠️ 有报告称会崩溃（JSAPI 环境） | 表格边框用 `Table` 对象的边框属性，COM 下先小范围实测 |
| `Shapes.AddChart2(-1, …)` | ⚠️ Style=-1 返回 null | 用 `AddChart2(0, 51, …)`（0 = 默认样式，51 = 柱形图） |
| 函数库 | WPS 表格约 400+ 函数，缺 `LAMBDA`/`WEBSERVICE` 等；动态数组（FILTER/SORT/XLOOKUP）支持不完整 | 复杂公式先实测；XLOOKUP 嵌套数组回退有 bug（apexcheng 实测记录 ✅） |
| 数据透视表/透视图 | 功能弱于 Excel | 能不用就不用，技能里标注"有限支持" |
| 以 `=` 开头的纯文本写入单元格 | ⚠️ WPS 按公式解析，非法公式抛 COM 异常（影刀 + WPS 实测） | 写入前加英文单引号 `'` 前缀 |

中文字体必须双设（✅ 通用 Word COM 经验，WPS 适用）：

```powershell
$font = $range.Font
$font.Name = "Calibri"      # 西文字体
$font.NameFarEast = "宋体"  # 中文字体（只设 Name 的话中文可能 fallback 到默认字体）
```

---

## 3. 版本 / Edition 差异

| 版本 | COM 外部自动化 | VBA 宏（编辑器） | JS 宏（JSA） |
|---|---|---|---|
| 个人版（免费） | ✅ 可用（社区共识：hch135861、xiaoqiong0v0、ouli-1242 的 skill 均基于个人版环境） | ❌ 需会员/专业版，或装第三方 vba71 插件 | ✅ 免费（WPS 2021+） |
| 专业版 / 商业版 | ✅ 可用（可选 64 位） | ✅ | ✅ |
| 教育版 / 政务版 | ✅（ProgID 可能不同 ⚠️） | 视授权 | ✅ |

关键结论：

- **COM 自动化本身不收费**。被收费墙挡住的是"应用内 VBA 编辑器"——Agent 走的是外部 COM（PowerShell / Python win32com），不受影响。✅
- WPS 个人版是 **32 位**；专业版可选 64 位。✅（xiaoqiong0v0 skill 明确写了）
- 免费版导出 PDF：社区说法是"无次数限制，但可能带 WPS Office 水印，高级设置需会员"。⚠️ 单源（网易号文章），**必须真机验证后再写进技能**。

---

## 4. 常用任务 PowerShell 片段

> 以下按 Word/Excel COM 的标准写法给出，对象模型与 WPS 一致；逐段都经过逻辑自检，但**首次落地前必须在真机 WPS 上跑一遍**（特别是标 ⚠️ 的行）。

### 4.1 通用模板（所有任务先套这个壳）

```powershell
$wps = $null
try {
    $wps = New-Object -ComObject "KWPS.Application"  # 表格换 KET.Application，演示换 KWPP.Application
    $wps.Visible = $false        # 后台运行（WPS 尊重该属性 ✅）
    $wps.DisplayAlerts = 0       # wdAlertsNone：关闭一切弹窗
    # ... 业务操作 ...
} finally {
    if ($wps) {
        $wps.Quit()              # 必须！否则 wps.exe / et.exe / wpp.exe 残留
        [Runtime.InteropServices.Marshal]::ReleaseComObject($wps) | Out-Null
    }
}
```

### 4.2 WPS 文字（Writer）

```powershell
$wps = New-Object -ComObject "KWPS.Application"
$wps.Visible = $false; $wps.DisplayAlerts = 0
try {
    # 新建文档
    $doc = $wps.Documents.Add()

    # 写段落 + 标题样式（注意样式名带空格："标题 1"）
    $p1 = $doc.Content.Paragraphs.Add()
    $p1.Range.Text = "2026 年第三季度工作报告"
    $p1.Style = $doc.Styles.Item("标题 1")
    $p1.Range.Font.NameFarEast = "黑体"

    $p2 = $doc.Content.Paragraphs.Add()
    $p2.Range.Text = "本季度核心指标同比增长 15%，主要得益于……"
    $p2.Style = $doc.Styles.Item("正文")
    $p2.Range.Font.NameFarEast = "宋体"
    $p2.Range.Font.Size = 12

    # 查找替换（全部替换）：参数顺序
    # FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
    # MatchAllWordForms, Forward, Wrap(1=wdFindContinue), Format, ReplaceWith, Replace(2=wdReplaceAll)
    $doc.Content.Find.Execute("十五%", $false, $false, $false, $false, $false, $true, 1, $false, "15%", 2)

    # 保存为 docx（16 = wdFormatDocumentDefault）
    $doc.SaveAs("D:\out\报告.docx", 16)
    # 导出 PDF（17 = wdExportFormatPDF）
    $doc.ExportAsFixedFormat("D:\out\报告.pdf", 17)

    $doc.Close()
} finally { $wps.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($wps) | Out-Null }
```

打开已有文档：`$doc = $wps.Documents.Open("D:\in\合同.docx")`；只读打开加 `$true` 作为第三个参数（ConfirmConversions/ReadOnly 位置参数：`Open(FileName, ConfirmConversions, ReadOnly)`）。

### 4.3 WPS 表格（Spreadsheets）

```powershell
$et = New-Object -ComObject "KET.Application"
$et.Visible = $false; $et.DisplayAlerts = 0
$et.ScreenUpdating = $false   # 批量写时提速
try {
    $wb = $et.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "销售数据"

    # 批量写（二维数组一次写入，比逐格快一个数量级）
    $data = @(
        @("产品", "销量", "金额"),
        @("产品A", 1500, 300000),
        @("产品B", 900, 180000)
    )
    $ws.Range("A1:C3").Value2 = $data

    # 公式 + 数字格式
    $ws.Range("C4").Formula = "=SUM(C2:C3)"
    $ws.Range("C2:C4").NumberFormat = "#,##0"

    # 表头加粗 + 底纹（颜色是 BGR 长整型：黄色 = 65535，红色 = 255）
    $ws.Range("A1:C1").Font.Bold = $true
    $ws.Range("A1:C1").Interior.Color = 65535

    # 图表（⚠️ Style 用 0 别用 -1；51 = xlColumnClustered 柱形图）
    $chart = $ws.Shapes.AddChart2(0, 51, 350, 20, 480, 300).Chart
    $chart.SetSourceData($ws.Range("A1:C3"))
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = "季度销量"

    # 一次读全表（UsedRange.Value2 返回二维数组）
    $all = $ws.UsedRange.Value2

    $wb.SaveAs("D:\out\销售.xlsx", 51)      # 51 = xlOpenXMLWorkbook
    $wb.ExportAsFixedFormat(0, "D:\out\销售.pdf")  # 0 = xlTypePDF
    $wb.Close()
} finally { $et.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($et) | Out-Null }
```

### 4.4 WPS 演示（Presentation）

```powershell
$wpp = New-Object -ComObject "KWPP.Application"
$wpp.Visible = $false; $wpp.DisplayAlerts = 0
try {
    $pres = $wpp.Presentations.Add()

    # 16:9 页面（PowerShell 里用 Points；960x540 pt 约等于 16:9）
    $pres.PageSetup.SlideWidth = 960
    $pres.PageSetup.SlideHeight = 540

    # 标题页（1 = ppLayoutTitle）
    $s1 = $pres.Slides.Add(1, 1)
    $s1.Shapes.Title.TextFrame.TextRange.Text = "未来已来"
    $s1.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "2026 年度技术分享"

    # 空白页（12 = ppLayoutBlank）+ 文本框（1 = msoTextBox）
    $s2 = $pres.Slides.Add(2, 12)
    $tb = $s2.Shapes.AddTextbox(1, 72, 72, 800, 200)
    $tb.TextFrame.TextRange.Text = "核心观点一二三……"
    $tb.TextFrame.TextRange.Font.Size = 28
    $tb.TextFrame.TextRange.Font.NameFarEast = "微软雅黑"

    # 插入图片
    $s2.Shapes.AddPicture("D:\img\chart.png", $false, $true, 72, 300, 400, 225)

    $pres.SaveAs("D:\out\分享.pptx")        # 默认即 pptx
    $pres.SaveAs("D:\out\分享.pdf", 32)     # 32 = ppSaveAsPDF

    # 逐页导出 PNG（⚠️ 需真机验证 WPS 是否支持 Slide.Export）
    # $s1.Export("D:\out\slide1.png", "PNG", 1920, 1080)

    $pres.Close()
} finally { $wpp.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($wpp) | Out-Null }
```

---

## 5. 坑清单（Gotchas）

1. **进程残留**：`Quit()` 必须在 `finally` 里调；异常崩溃后手动杀 `wps.exe`、`et.exe`、`wpp.exe`。残留进程会导致下次 `Documents.Open` 卡死。✅
2. **不要 Quit 别人的实例**：如果用户已经开着 WPS，直接 `New-Object` 可能复用现实例（WPS 是单主进程架构）。技能里建议"会话"模式：先检查是否已有实例（`Get-Process wps/et/wpp`），有则只关文档不 `Quit()`。✅（hch135861 的 skill 实现了 Open-WpsSession/Close-WpsSession 就是干这个的）
3. **32/64 位**：个人版 WPS 是 32 位；64 位 Python/PowerShell 里 `Dispatch` 报"找不到 ProgID"时，换 32 位宿主（`C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe`）。✅
4. **DisplayAlerts=0**：不设的话，文件已存在/兼容性提示会弹模态框，自动化直接卡死。✅
5. **登录/广告弹窗**：WPS 未登录时会弹登录框、首页有稻壳模板/广告。缓解：`Visible=$false` 下通常不出现；真机若出现，可用"配置和修复工具 → 高级 → 功能定制"关推荐，或注册表 `HKCU\Software\kingsoft\Office\6.0\plugins\officespace\flogin` 下建字符串值 `enableforceloginforfirstinstalldevice=false` 解除强制登录限制。⚠️（注册表法来自 360doc 社区，未验证）
6. **KPDF.Application 别碰**：WPS 的 PDF 组件 COM Dispatch 会弹阻塞式对话框，不适合自动化；PDF 合并/拆分/水印用 pypdf/pdfplumber 等纯 Python 库。✅（neomei/wpscomposer 真机验证结论）
7. **云文档**：COM 只能操作本地文件，WPS 云文档先同步到本地。✅
8. **UNC 路径/符号链接**：有 skill 报告不支持，只用本地绝对路径。⚠️
9. **`.Borders` 谨慎**：JSAPI 环境有崩溃报告；COM 下建议先小范围实测再写进技能。⚠️
10. **COM 改动不一定进撤销栈**：重要文档先另存副本再改。✅（sueccku skill）
11. **启动慢**：每次 COM 操作含 WPS 冷启动，约数秒；批量任务复用同一个 Application 实例，不要每个文件起一次。✅

---

## 6. 无头 / 服务端运行

- WPS 是桌面 GUI 应用，COM 自动化**需要已登录的交互式 Windows 会话**。⚠️
- 从 Windows 服务（Session 0）、无桌面的计划任务调用，大概率失败或卡死——这是所有 Office 类 COM 自动化的通病（MS 有 KB257757 明确说 Office 不支持服务端自动化；WPS 官方没有类似声明，但社区无人报告成功案例）。⚠️ 未找到 WPS 专用资料，**技能文档里应写"需要用户已登录的桌面"，不要承诺服务端可用**。
- 计划任务场景：用"只在用户登录时运行" + 保持会话，或配一台常亮的"自动化小主机"。⚠️ 经验性建议。
- Linux 版 WPS 没有 COM（只有 Windows 有）；Linux 下替代是 `pywpsrpc`（需 X11 桌面 / xvfb）。⚠️ 与本技能包（Windows PowerShell 路线）无关，仅备注。

---

## 7. 宏路线：VBA vs JS 宏（JSA）

| | VBA | JS 宏（JSA） |
|---|---|---|
| 语言 | VB | JavaScript（V8，支持到 ES2019） |
| 个人版 | ❌ 需会员/专业版（或第三方 vba71 插件） | ✅ 免费，WPS 2021+ 内置 |
| 运行位置 | 应用内 | 应用内（宏编辑器） |
| 宏录制 | ✅ 完善 | ✅ 支持（录出来是箭头函数包裹风格） |
| Agent 视角 | 不推荐：要用户先解决授权问题 | 可选：适合"在用户已打开的文档里跑一段逻辑" |

JSAPI 与 VBA 的关键语法差异（✅ 来自 iheng88/claude-office-plugin，Agent 生成 JS 宏代码时必须遵守）：

- 集合必须 `.Item()`：`Worksheets.Item(1)`，不能 `Worksheets(1)`
- 方法必须加括号：`.Select()`、`.Activate()`
- 读写用 `Value2`：`Range("A1").Value2 = 5`
- 图表：`Shapes.AddChart2(0, type, …)`，Style 不能是 -1
- 大小写敏感：`Value2` 不是 `value2`

⚠️ 外部触发 JS 宏（命令行一键运行 .jsm）是否有稳定入口，未找到可靠资料——技能 v1 先不承诺这条路，主打 PowerShell COM。

---

## 8. 先行者（竞品）参考

| 项目 | 路线 | star（2026-09-28） | 备注 |
|---|---|---|---|
| lc2panda/wps-skills | MCP Server（243 工具）+ WPS 加载项 | ~630 | 最重，需要用户装加载项+Node |
| hch135861/wps-office | Codex skill + PowerShell COM 脚本 | 新 | 与我们路线最接近：会话管理、helper 脚本 |
| xiaoqiong0v0/opencode-skills（wps-api） | Skill 文档（Python win32com） | 新 | 参考文件拆分得最细，可借鉴结构 |
| neomei/wpscomposer | Python COM 引擎（layout 级生成） | 新 | 真机验证最扎实；KPDF 坑来自它 |
| ouli-1242/wps-cli | Python CLI（typer + win32com） | 新 | 命令设计可参考 |
| sueccku/dsh-plugin-wps-office-next | 插件式 skill | 新 | 中文坑记录（样式名空格等） |

差异化空间：**纯 PowerShell、零依赖（不装 Node/Python/加载项）、WPS 专精**——现有项目要么要装运行时，要么走 MCP/加载项。对"用户机器只有 WPS"这个最大公约数场景，我们是开箱即用的。

---

## 9. 未验证事项（真机测试清单）

1. 免费版导出 PDF 是否带水印
2. `Slide.Export` 导出 PNG 在 WPS 下是否可用
3. `.Borders` 在 COM 下是否真会崩溃
4. 专业版/政务版的 ProgID 是否不同
5. 注册表 `enableforceloginforfirstinstalldevice` 是否有效
6. 64 位 PowerShell 调 32 位 WPS COM 的实际表现
7. `Visible=$false` 下首次启动是否弹登录框
8. JS 宏能否从命令行外部触发

---

*资料来源：GitHub（hch135861/wps-office、neomei/wpscomposer、xiaoqiong0v0/opencode-skills、yb2460/harness-anything、iheng88/claude-office-plugin、lc2panda/wps-skills、ouli-1242/wps-cli、sueccku/dsh-plugin-wps-office-next、apexcheng 博客）、CSDN、360doc、dev.to；MS Office 枚举常量为公开知识。*
