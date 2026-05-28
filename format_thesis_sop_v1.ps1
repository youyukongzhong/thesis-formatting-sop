<#

.SYNOPSIS
论文格式修改系统 v1 — 模板脚本

.DESCRIPTION
这是"论文修改系统 SOP v1"的自动化演示模板。
它展示了整套系统的执行流程：文档体检 → 样式库 → 全局框架 → 正文处理 → 图表文献 → 终检。

使用方法：
    1. 打开 PowerShell（管理员非必需，但建议）
    2. 把论文 docx 和格式要求文件放到 workspace/ 目录下
    3. 运行脚本，传入论文路径和格式要求路径：
       .\format_thesis_sop_v1.ps1 -ThesisPath "workspace\你的论文.docx" -RequirementsPath "workspace\格式要求.doc"
    4. 交付包会生成在 workspace/交付_时间戳/ 下

前置条件：
    - 已安装 Microsoft Word（脚本通过 COM 操作 Word）
    - 论文主文件必须是 docx（doc 可转但不在 doc 上长期施工）
    - 有学校格式要求文件（doc/docx/pdf/图片均可，脚本只复制留档，规则需手动填入下方）

注意：
    - 本脚本是模板，里面的格式规则（字体、字号、行距等）需要按你的学校要求修改
    - 搜索"【待修改】"标记的位置，把示例规则替换成你的实际要求

兼容性：
    - 本脚本不依赖 Claude Code 或任何 AI 工具，纯 PowerShell，可在任何 Windows 机器上运行
    - 配合 AI 使用时，建议把 01_SOP.md 给 AI 阅读，让 AI 理解流程后再执行
#>

param(
    [Parameter(Mandatory = $true, HelpMessage = "论文 docx 文件的路径")]
    [string]$ThesisPath,

    [Parameter(Mandatory = $true, HelpMessage = "学校格式要求文件的路径")]
    [string]$RequirementsPath
)

$ErrorActionPreference = "Stop"

# ============================================================
# 工具函数
# ============================================================

function Join-Child([string]$base, [string]$child) {
    return [System.IO.Path]::Combine($base, $child)
}

function Cm([double]$value) {
    return $value * 28.3464566929
}

function Safe-Text([string]$text) {
    if ($null -eq $text) { return "" }
    return ($text -replace "`r", "").Trim()
}

function Ensure-ParagraphStyle($doc, [string]$name) {
    try {
        return $doc.Styles.Item($name)
    } catch {
        return $doc.Styles.Add($name, 1)
    }
}

function Set-ParagraphStyle($style, [string]$eastAsia, [string]$ascii, [double]$size, [int]$bold,
    [int]$alignment, [double]$before, [double]$after, [int]$lineRule, [double]$lineSpacing,
    [double]$firstIndent, [double]$leftIndent) {
    $style.Font.NameFarEast = $eastAsia
    $style.Font.NameAscii = $ascii
    $style.Font.NameOther = $ascii
    $style.Font.Size = $size
    $style.Font.Bold = $bold
    $style.ParagraphFormat.Alignment = $alignment
    $style.ParagraphFormat.SpaceBefore = $before
    $style.ParagraphFormat.SpaceAfter = $after
    $style.ParagraphFormat.LineSpacingRule = $lineRule
    if ($lineSpacing -gt 0) { $style.ParagraphFormat.LineSpacing = $lineSpacing }
    $style.ParagraphFormat.FirstLineIndent = $firstIndent
    $style.ParagraphFormat.LeftIndent = $leftIndent
}

function Apply-DirectFormat($range, [string]$eastAsia, [string]$ascii, [double]$size, [int]$bold,
    [int]$alignment, [double]$before, [double]$after, [int]$lineRule, [double]$lineSpacing,
    [double]$firstIndent, [double]$leftIndent) {
    $range.Font.NameFarEast = $eastAsia
    $range.Font.NameAscii = $ascii
    $range.Font.NameOther = $ascii
    $range.Font.Size = $size
    $range.Font.Bold = $bold
    $range.ParagraphFormat.Alignment = $alignment
    $range.ParagraphFormat.SpaceBefore = $before
    $range.ParagraphFormat.SpaceAfter = $after
    $range.ParagraphFormat.LineSpacingRule = $lineRule
    if ($lineSpacing -gt 0) { $range.ParagraphFormat.LineSpacing = $lineSpacing }
    $range.ParagraphFormat.FirstLineIndent = $firstIndent
    $range.ParagraphFormat.LeftIndent = $leftIndent
}

function In-Table($range) {
    try {
        return [bool]$range.Information(12)
    } catch {
        return $false
    }
}

function Find-Paragraph($doc, [scriptblock]$predicate) {
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        $txt = Safe-Text $p.Range.Text
        if (& $predicate $txt $p $i) { return $p }
    }
    return $null
}

function Clear-HeaderFooter($section) {
    foreach ($idx in 1, 2, 3) {
        $section.Headers.Item($idx).LinkToPrevious = $false
        $section.Footers.Item($idx).LinkToPrevious = $false
        $section.Headers.Item($idx).Range.Text = ""
        $section.Footers.Item($idx).Range.Text = ""
    }
}

function Set-HeaderText($section, [string]$text, [string]$eastAsia, [string]$ascii) {
    foreach ($idx in 1, 2, 3) {
        $header = $section.Headers.Item($idx)
        $header.LinkToPrevious = $false
        $header.Range.Text = $text
        $header.Range.ParagraphFormat.Alignment = 1
        $header.Range.Font.NameFarEast = $eastAsia
        $header.Range.Font.NameOther = $ascii
        $header.Range.Font.Size = 10.5
    }
}

function Set-HeaderStyleRef($doc, $section) {
    foreach ($idx in 1, 2, 3) {
        $header = $section.Headers.Item($idx)
        $header.LinkToPrevious = $false
        $header.Range.Text = ""
        $header.Range.ParagraphFormat.Alignment = 1
        $header.Range.Font.NameFarEast = "宋体"
        $header.Range.Font.NameOther = "Times New Roman"
        $header.Range.Font.Size = 10.5
        $fieldRange = $header.Range
        $fieldRange.Collapse(1)
        $doc.Fields.Add($fieldRange, -1, 'STYLEREF "标题 1"', $true) | Out-Null
    }
}

function Set-CenteredPageNumber($section, [int]$numberStyle, [int]$startAt) {
    foreach ($idx in 1, 2, 3) {
        $footer = $section.Footers.Item($idx)
        $footer.LinkToPrevious = $false
        $footer.Range.Text = ""
        $footer.Range.ParagraphFormat.Alignment = 1
        try {
            $footer.PageNumbers.RestartNumberingAtSection = $true
            $footer.PageNumbers.StartingNumber = $startAt
            $footer.PageNumbers.NumberStyle = $numberStyle
        } catch {}
        $fieldCode = "PAGE"
        if ($numberStyle -eq 1) { $fieldCode = "PAGE \* ROMAN" }
        $fieldRange = $footer.Range
        $fieldRange.Collapse(1)
        $section.Range.Document.Fields.Add($fieldRange, -1, $fieldCode, $true) | Out-Null
    }
}

# ============================================================
# 第 1 步：验证输入，准备输出目录
# ============================================================

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$workspace = Join-Child $root "workspace"

if (-not (Test-Path -LiteralPath $ThesisPath)) { throw "未找到论文主文件：$ThesisPath" }
if (-not (Test-Path -LiteralPath $RequirementsPath)) { throw "未找到格式要求文件：$RequirementsPath" }

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outDir = Join-Child $workspace "交付_$stamp"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$thesisFileName = [System.IO.Path]::GetFileNameWithoutExtension($ThesisPath)
$backup = Join-Child $outDir "01_原稿备份_${thesisFileName}.docx"
$final = Join-Child $outDir "02_最终稿_${thesisFileName}.docx"
$reportPath = Join-Child $outDir "03_修改报告.md"
$todoPath = Join-Child $outDir "04_待确认问题.md"

Copy-Item -LiteralPath $ThesisPath -Destination $backup -Force
Copy-Item -LiteralPath $ThesisPath -Destination $final -Force


$word = $null
$doc = $null
$diagnosis = New-Object System.Collections.Generic.List[string]
$changes = New-Object System.Collections.Generic.List[string]
$manual = New-Object System.Collections.Generic.List[string]

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $doc = $word.Documents.Open((Resolve-Path -LiteralPath $final).Path, $false, $false)

    # ---- 3. 文档体检 ----
    $diagnosis.Add("原稿为 docx，可作为正式工作格式。")
    $diagnosis.Add("原稿初始状态：$($doc.Sections.Count) 个分节，$($doc.TablesOfContents.Count) 个自动目录，$($doc.Tables.Count) 个表格，$($doc.InlineShapes.Count + $doc.Shapes.Count) 个图形/图片对象。")

    if ($doc.TablesOfContents.Count -eq 0) { $diagnosis.Add("未发现自动目录。") }
    if ($doc.Sections.Count -eq 1) { $diagnosis.Add("仅有 1 个分节，无法满足封面/目录/正文页码体系要求。") }

    # 检查前置部分（封面、声明、摘要）
    $missingFrontMatter = @("封面", "声明", "摘要", "Abstract")
    foreach ($label in $missingFrontMatter) {
        $found = Find-Paragraph $doc { param($txt, $p, $i) $txt -eq $label -or $txt -like "$label*" }
        if ($null -eq $found) { $manual.Add("未在原稿中发现《$label》部分，需要作者补齐或确认是否另有单独文件。") }
    }

    # ---- 4. 建立样式库 ----
    $styleNormal = $doc.Styles.Item(-1)
    $styleH1 = $doc.Styles.Item(-2)
    $styleH2 = $doc.Styles.Item(-3)
    $styleH3 = $doc.Styles.Item(-4)
    $styleH4 = $doc.Styles.Item(-5)
    $styleRef = Ensure-ParagraphStyle $doc "论文参考文献"
    $styleTableCaption = Ensure-ParagraphStyle $doc "论文表题"
    $styleFigureCaption = Ensure-ParagraphStyle $doc "论文图题"
    $styleTitle = Ensure-ParagraphStyle $doc "论文标题页"
    $styleTocTitle = Ensure-ParagraphStyle $doc "论文目录题头"

    # 【待修改】以下样式参数按你的学校要求修改
    Set-ParagraphStyle $styleNormal "宋体" "Times New Roman" 12 0 3 0 0 4 20 24 0
    Set-ParagraphStyle $styleH1 "黑体" "Times New Roman" 16 -1 1 24 18 0 0 0 0
    Set-ParagraphStyle $styleH2 "黑体" "Times New Roman" 14 -1 0 24 6 4 20 0 24
    Set-ParagraphStyle $styleH3 "黑体" "Times New Roman" 14 -1 0 12 6 4 20 0 24
    Set-ParagraphStyle $styleH4 "黑体" "Times New Roman" 12 -1 0 12 6 4 20 0 24
    Set-ParagraphStyle $styleRef "宋体" "Times New Roman" 10.5 0 3 3 0 4 16 0 0
    Set-ParagraphStyle $styleTableCaption "宋体" "Times New Roman" 10.5 0 1 12 6 0 0 0 0
    Set-ParagraphStyle $styleFigureCaption "宋体" "Times New Roman" 10.5 0 1 6 12 0 0 0 0
    Set-ParagraphStyle $styleTitle "楷体_GB2312" "Times New Roman" 18 -1 1 0 0 0 0 0 0
    Set-ParagraphStyle $styleTocTitle "黑体" "Times New Roman" 16 -1 1 24 18 0 0 0 0
    Apply-DirectFormat $doc.Content "宋体" "Times New Roman" 12 0 3 0 0 4 20 24 0
    $changes.Add("已建立/规范化正文、四级标题、目录题头、图题、表题、参考文献等样式。")

    # ---- 5. 重建全局框架 ----
    $state = "body"
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        $txt = Safe-Text $p.Range.Text
        if ($txt.Length -eq 0 -or (In-Table $p.Range)) { continue }

        if ($i -eq 1) {
            $p.Range.Style = $styleTitle
            Apply-DirectFormat $p.Range "楷体_GB2312" "Times New Roman" 18 -1 1 0 0 0 0 0 0
            continue
        }

        if ($txt -eq "参考文献") { $state = "references" }

        # 【待修改】以下标题匹配规则按你的论文实际标题体系调整
        if ($txt -match "^第[一二三四五六七八九十0-9]+章\s*.+" -or
            $txt -eq "参考文献" -or $txt -eq "致谢") {
            $p.Range.Style = $styleH1
            Apply-DirectFormat $p.Range "黑体" "Times New Roman" 16 -1 1 24 18 0 0 0 0
            if ($txt -ne "绪论") { $p.Format.PageBreakBefore = $true } else { $p.Format.PageBreakBefore = $false }
        } elseif ($txt -match "^[一二三四五六七八九十]+、") {
            $p.Range.Style = $styleH2
            Apply-DirectFormat $p.Range "黑体" "Times New Roman" 14 -1 0 24 6 4 20 0 24
        } elseif ($txt -match "^（[一二三四五六七八九十]+）") {
            $p.Range.Style = $styleH3
            Apply-DirectFormat $p.Range "黑体" "Times New Roman" 14 -1 0 12 6 4 20 0 24
        } elseif ($txt -match "^[0-9]+[.．]\s*\S+" -or $txt -match "^\([0-9]+\)\s*\S+") {
            $p.Range.Style = $styleH4
            Apply-DirectFormat $p.Range "黑体" "Times New Roman" 12 -1 0 12 6 4 20 0 24
        } elseif ($state -eq "references") {
            $p.Range.Style = $styleRef
            Apply-DirectFormat $p.Range "宋体" "Times New Roman" 10.5 0 3 3 0 4 16 0 0
        } else {
            $p.Range.Style = $styleNormal
        }
    }
    $changes.Add("已按标题层级重设全文章节样式。")

    # 自动目录处理
    $intro = Find-Paragraph $doc { param($txt, $p, $i) $txt -eq "绪论" -or $txt -like "第一章*" }
    if ($null -ne $intro -and $doc.TablesOfContents.Count -eq 0) {
        $intro.Range.Select()
        $word.Selection.Collapse(1) | Out-Null
        $word.Selection.InsertBreak(2) | Out-Null
        $word.Selection.TypeText("目录")
        $word.Selection.TypeParagraph()
        $tocRange = $word.Selection.Range
        $doc.TablesOfContents.Add($tocRange, $true, 1, 3) | Out-Null
        $doc.TablesOfContents.Item(1).Update()
        $tocEnd = $doc.TablesOfContents.Item(1).Range.End
        $afterToc = $doc.Range($tocEnd, $tocEnd)
        $afterToc.InsertBreak(2) | Out-Null
        $tocTitle = Find-Paragraph $doc { param($txt, $p, $i) $txt -eq "目录" }
        if ($null -ne $tocTitle) {
            $tocTitle.Range.Style = $styleTocTitle
            Apply-DirectFormat $tocTitle.Range "黑体" "Times New Roman" 16 -1 1 24 18 0 0 0 0
        }
        $changes.Add("已在标题页后新增自动目录，目录层级为 1-3 级。")
    } elseif ($doc.TablesOfContents.Count -gt 0) {
        $doc.TablesOfContents.Item(1).Update()
        $changes.Add("已更新原有自动目录。")
    } else {
        $manual.Add("未能定位章节起始，未自动插入目录。")
    }

    # 页面设置
    for ($s = 1; $s -le $doc.Sections.Count; $s++) {
        $sec = $doc.Sections.Item($s)
        $sec.PageSetup.PaperSize = 7
        $sec.PageSetup.HeaderDistance = Cm 2.8
        $sec.PageSetup.FooterDistance = Cm 2.8
    }

    # 页眉页脚
    if ($doc.Sections.Count -ge 1) {
        Clear-HeaderFooter $doc.Sections.Item(1)
    }
    if ($doc.Sections.Count -ge 2) {
        Set-HeaderText $doc.Sections.Item(2) "目录" "宋体" "Times New Roman"
        Set-CenteredPageNumber $doc.Sections.Item(2) 1 1
    }
    if ($doc.Sections.Count -ge 3) {
        for ($s = 3; $s -le $doc.Sections.Count; $s++) {
            Set-HeaderStyleRef $doc $doc.Sections.Item($s)
            Set-CenteredPageNumber $doc.Sections.Item($s) 0 1
        }
    } else {
        $manual.Add("因文档分节不足，正文页码体系可能需要人工复核。")
    }
    $changes.Add("已按三段式建立页眉页脚和页码分区。")

    # ---- 6. 图表处理 ----
    for ($t = 1; $t -le $doc.Tables.Count; $t++) {
        $table = $doc.Tables.Item($t)
        $caption = $doc.Range([Math]::Max(0, $table.Range.Start - 120), $table.Range.Start)
        $captionText = Safe-Text $caption.Text
        if ($captionText -notmatch "表[0-9一二三四五六七八九十]") {
            $manual.Add("第 $t 个表格上方标题未见规范表序，需要确认是否改为《表x.x 题名》。")
        }

        $beforePara = $null
        for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
            $p = $doc.Paragraphs.Item($i)
            if ($p.Range.End -le $table.Range.Start) { $beforePara = $p } else { break }
        }
        if ($null -ne $beforePara) {
            $bt = Safe-Text $beforePara.Range.Text
            if ($bt.Length -gt 0) {
                $beforePara.Range.Style = $styleTableCaption
                Apply-DirectFormat $beforePara.Range "宋体" "Times New Roman" 10.5 0 1 12 6 0 0 0 0
            }
        }

        $table.Range.Font.NameFarEast = "宋体"
        $table.Range.Font.NameOther = "Times New Roman"
        $table.Range.Font.Size = 10.5
        $table.Range.ParagraphFormat.Alignment = 1
        $table.Range.ParagraphFormat.LineSpacingRule = 0
        $table.Range.ParagraphFormat.SpaceBefore = 3
        $table.Range.ParagraphFormat.SpaceAfter = 3
        foreach ($row in $table.Rows) { $row.Cells.VerticalAlignment = 1 }

        $table.Borders.Enable = 0
        $table.Borders.InsideLineStyle = 0
        $table.Borders.OutsideLineStyle = 0
        $table.Borders.Item(-1).LineStyle = 1
        $table.Borders.Item(-1).LineWidth = 12
        $table.Borders.Item(-3).LineStyle = 1
        $table.Borders.Item(-3).LineWidth = 12
        if ($table.Rows.Count -ge 1) {
            $table.Rows.Item(1).Borders.Item(-3).LineStyle = 1
            $table.Rows.Item(1).Borders.Item(-3).LineWidth = 8
        }
    }
    if ($doc.Tables.Count -gt 0) {
        $changes.Add("已将 $($doc.Tables.Count) 个表格整理为三线表样式。")
    }
    if (($doc.InlineShapes.Count + $doc.Shapes.Count) -eq 0) {
        $diagnosis.Add("未发现图片或图形对象。")
    }

    # ---- 7. 更新域 ----
    foreach ($field in $doc.Fields) {
        try { $field.Update() | Out-Null } catch {}
    }
    if ($doc.TablesOfContents.Count -gt 0) {
        $doc.TablesOfContents.Item(1).Update()
    }

    $doc.Save()

    # ---- 8. 生成交付包（合并报告） ----
    $diagnosisText = ($diagnosis | ForEach-Object { "- $_" }) -join [Environment]::NewLine
    $changesText = ($changes | ForEach-Object { "- $_" }) -join [Environment]::NewLine
    $manualText = ($manual | Select-Object -Unique | ForEach-Object { "- $_" }) -join [Environment]::NewLine

    $report = @(
        "# 修改报告",
        "",
        "依据格式：$(Split-Path -Leaf $RequirementsPath)",
        "",
        "## 文档体检",
        "",
        $diagnosisText,
        "",
        "## 已修改",
        "",
        $changesText,
        "",
        "## 仍未自动处理的项目",
        "",
        "- 格式规则需对照学校规范原文复核（以下为本次使用的示例规则，请自行验证准确性）：",
        "  - 章标题：黑体三号加粗，居中，单倍行距，段前 24 磅，段后 18 磅",
        "  - 一级标题：黑体四号加粗，段前 24 磅，段后 6 磅，左缩进 2 字符",
        "  - 正文：小四宋体，两端对齐，首行缩进 2 字符，固定 20 磅行距",
        "  - 页眉：五号宋体居中，距边界 2.8 厘米",
        "  - 页码：目录罗马数字，正文阿拉伯数字",
        "- 封面、声明、摘要等前置部分是否完整。",
        "- 文献编号、正文引用与真实来源。",
        "- 目录页码是否与实际一致（建议更新域后人工目检）。",
        "",
        "## 声明",
        "",
        "- 本次未改写论文核心论证、未补写章节内容、未虚构数据或文献。",
        "- 已在副本上操作，原稿完整保留。",
        "- 最终提交前需结合完整论文材料复核。"
    ) -join [Environment]::NewLine
    Set-Content -LiteralPath $reportPath -Value $report -Encoding UTF8

    $todo = @(
        "# 待确认问题",
        "",
        if ($manualText.Length -gt 0) { $manualText } else { "- 无自动标记问题。" }
    ) -join [Environment]::NewLine
    Set-Content -LiteralPath $todoPath -Value $todo -Encoding UTF8

    [pscustomobject]@{
        OutputDirectory = $outDir
        Backup = $backup
        FinalDraft = $final
        Report = $reportPath
        Todo = $todoPath
    } | Format-List
} finally {
    if ($null -ne $doc) {
        try { $doc.Close($false) | Out-Null } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    }
    if ($null -ne $word) {
        try { $word.Quit() | Out-Null } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}