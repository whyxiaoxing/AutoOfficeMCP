# 生成全工具冒烟测试的基础夹具。
# 只需要执行一次；run-all.ps1 会在夹具缺失时自动调用。
param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$fixtureDir = Join-Path $PSScriptRoot "fixtures"
New-Item -ItemType Directory -Path $fixtureDir -Force | Out-Null

$excelPath = Join-Path $fixtureDir "base.xlsx"
$excelCopyPath = Join-Path $fixtureDir "base-copy.xlsx"
$wordPath = Join-Path $fixtureDir "base.docx"
$wordCopyPath = Join-Path $fixtureDir "base-copy.docx"
$pptPath = Join-Path $fixtureDir "base.pptx"
$pptCopyPath = Join-Path $fixtureDir "base-copy.pptx"
$imagePath = Join-Path $fixtureDir "image.png"
$csvPath = Join-Path $fixtureDir "data.csv"

foreach ($path in @($excelPath, $excelCopyPath, $wordPath, $wordCopyPath, $pptPath, $pptCopyPath, $imagePath, $csvPath)) {
    if ($Force -and (Test-Path -LiteralPath $path)) {
        Remove-Item -LiteralPath $path -Force
    }
}

function Release-ComObject {
    param($Object)
    if ($null -ne $Object) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) | Out-Null } catch { }
    }
}

# ---------- image / csv ----------
if (-not (Test-Path -LiteralPath $imagePath)) {
    try {
        Add-Type -AssemblyName System.Drawing
        $bitmap = New-Object System.Drawing.Bitmap 32, 32
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([System.Drawing.Color]::CornflowerBlue)
        $bitmap.Save($imagePath, [System.Drawing.Imaging.ImageFormat]::Png)
        $graphics.Dispose()
        $bitmap.Dispose()
    }
    catch {
        # 兜底：1x1 PNG。
        $png = [Convert]::FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2nXQAAAAASUVORK5CYII=")
        [System.IO.File]::WriteAllBytes($imagePath, $png)
    }
}
if (-not (Test-Path -LiteralPath $csvPath)) {
    @(
        "Name,Q1,Q2",
        "Alpha,10,20",
        "Beta,30,40",
        "Gamma,50,60"
    ) | Set-Content -LiteralPath $csvPath -Encoding utf8NoBOM
}

# ---------- Excel ----------
if ($Force -or -not (Test-Path -LiteralPath $excelPath)) {
    $excel = $null
    $workbook = $null
    try {
        if (Test-Path -LiteralPath $excelPath) { Remove-Item -LiteralPath $excelPath -Force }
        if (Test-Path -LiteralPath $excelCopyPath) { Remove-Item -LiteralPath $excelCopyPath -Force }

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()

        $data = $workbook.Worksheets.Item(1)
        $data.Name = "Data"
        $headers = @("Name", "Q1", "Q2", "Total")
        for ($column = 1; $column -le $headers.Count; $column++) {
            $data.Cells.Item(1, $column).Value2 = $headers[$column - 1]
        }
        for ($row = 1; $row -le 10; $row++) {
            $data.Cells.Item($row + 1, 1).Value2 = "Item$row"
            $data.Cells.Item($row + 1, 2).Value2 = "$($row * 10)"
            $data.Cells.Item($row + 1, 3).Value2 = "$(100 - $row)"
            $data.Cells.Item($row + 1, 4).Formula = "=B$($row + 1)+C$($row + 1)"
        }

        $calc = $workbook.Worksheets.Add()
        $calc.Name = "Calc"
        $calc.Cells.Item(1, 1).Value2 = "A"
        $calc.Cells.Item(1, 2).Value2 = "B"
        $calc.Cells.Item(1, 3).Value2 = "Sum"
        for ($row = 2; $row -le 8; $row++) {
            $calc.Cells.Item($row, 1).Value2 = "$row"
            $calc.Cells.Item($row, 2).Value2 = "$($row * 2)"
            $calc.Cells.Item($row, 3).Formula = "=A$row+B$row"
        }

        $format = $workbook.Worksheets.Add()
        $format.Name = "Format"
        $format.Cells.Item(1, 1).Value2 = "Merged"
        $format.Cells.Item(3, 1).Value2 = "Alpha"
        $format.Cells.Item(3, 2).Value2 = "Beta"
        $format.Cells.Item(3, 3).Value2 = "Gamma"
        try { $format.Range("A9:B9").Merge() | Out-Null } catch { }
        try { $format.Range("A1").AddComment("Test comment") | Out-Null } catch { }
        try {
            $table = $format.ListObjects.Add(1, $format.Range("A3:C7"), $null, 1)
            $table.Name = "TestTable"
        }
        catch { }

        try { $null = $workbook.Names.Add("TestRange", "=Data!`$A`$1:`$C`$3") } catch { }
        try {
            $chartObject = $calc.Shapes.AddChart2(201, 51, 200, 10, 320, 200)
            $chartObject.Chart.SetSourceData($calc.Range("A1:B6"))
            $chartObject.Chart.HasTitle = $true
            $chartObject.Chart.ChartTitle.Text = "TestChart"
            $chartObject.Name = "TestChart"
            $chartObject.Chart.ChartTitle.Font.Size = 10
        }
        catch { }

        $workbook.SaveAs($excelPath)
        $workbook.SaveCopyAs($excelCopyPath)
        $workbook.Close($false)
        Release-ComObject $workbook
        $workbook = $null
        $null = $excel.Quit()
        Release-ComObject $excel
        $excel = $null
        Write-Host "FIXTURE excel: $excelPath"
    }
    finally {
        if ($null -ne $workbook) { try { $null = $workbook.Close($false) } catch { }; Release-ComObject $workbook }
        if ($null -ne $excel) { try { $null = $excel.Quit() } catch { }; Release-ComObject $excel }
    }
}

# ---------- Word ----------
if ($Force -or -not (Test-Path -LiteralPath $wordPath)) {
    $word = $null
    $document = $null
    try {
        if (Test-Path -LiteralPath $wordPath) { Remove-Item -LiteralPath $wordPath -Force }
        if (Test-Path -LiteralPath $wordCopyPath) { Remove-Item -LiteralPath $wordCopyPath -Force }

        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $document = $word.Documents.Add()
        $document.Content.Text = "AutoOfficeMCP All Tools`rHeading One`rAlpha Beta Gamma`rDelta Epsilon`rZeta Eta Theta`rFinal paragraph`r1. Manual Heading"

        $document.Paragraphs.Item(2).Range.Style = $document.Styles.Item(-2)
        try {
            $document.TrackRevisions = $true
            $document.Content.InsertAfter(" Revision line")
            $document.TrackRevisions = $false
        }
        catch { }
        try {
            $textBox = $document.Shapes.AddTextbox(1, 100, 100, 200, 50)
            $textBox.Name = "TestShape"
            $textBox.TextFrame.TextRange.Text = "Test shape"
        }
        catch { }
        try { $document.Paragraphs.Item(4).Range.InsertBreak(7) } catch { }
        try {
            $table = $document.Tables.Add($document.Paragraphs.Item(5).Range, 3, 3)
            for ($row = 1; $row -le 3; $row++) {
                for ($column = 1; $column -le 3; $column++) {
                    $table.Cell($row, $column).Range.Text = "R$row" + "C$column"
                }
            }
        }
        catch { }
        try { $null = $document.Comments.Add($document.Paragraphs.Item(3).Range, "Test comment") } catch { }
        try { $document.Sections.Item(1).Headers.Item(1).Range.Text = "AutoOfficeMCP Header" } catch { }
        try { $document.Sections.Item(1).Footers.Item(1).Range.Text = "AutoOfficeMCP Footer" } catch { }
        try { $null = $document.InlineShapes.AddPicture($imagePath) } catch { }
        try { $null = $document.Bookmarks.Add("TestBookmark", $document.Paragraphs.Item(3).Range) } catch { }

        $document.SaveAs($wordPath)
        $document.Close(0)
        Release-ComObject $document
        $document = $null
        $null = $word.Quit()
        Release-ComObject $word
        $word = $null
        Copy-Item -LiteralPath $wordPath -Destination $wordCopyPath -Force
        Write-Host "FIXTURE word: $wordPath"
    }
    finally {
        if ($null -ne $document) { try { $null = $document.Close(0) } catch { }; Release-ComObject $document }
        if ($null -ne $word) { try { $null = $word.Quit() } catch { }; Release-ComObject $word }
    }
}

# ---------- PowerPoint ----------
if ($Force -or -not (Test-Path -LiteralPath $pptPath)) {
    $ppt = $null
    $presentation = $null
    try {
        if (Test-Path -LiteralPath $pptPath) { Remove-Item -LiteralPath $pptPath -Force }
        if (Test-Path -LiteralPath $pptCopyPath) { Remove-Item -LiteralPath $pptCopyPath -Force }

        $ppt = New-Object -ComObject PowerPoint.Application
        $ppt.Visible = -1
        $presentation = $ppt.Presentations.Add(-1)
        $slide1 = $presentation.Slides.Add(1, 1)
        $slide1.Shapes.Title.TextFrame.TextRange.Text = "AutoOfficeMCP All Tools"
        try { $slide1.Shapes.Item(2).TextFrame.TextRange.Text = "Smoke test fixtures" } catch { }
        try {
            $rect1 = $slide1.Shapes.AddShape(1, 40, 380, 80, 40)
            $rect2 = $slide1.Shapes.AddShape(1, 140, 380, 80, 40)
            $rect1.Name = "GroupRect1"
            $rect2.Name = "GroupRect2"
            $rect3 = $slide1.Shapes.AddShape(1, 240, 380, 80, 40)
            $rect4 = $slide1.Shapes.AddShape(1, 340, 380, 80, 40)
            $rect3.Name = "GroupRect3"
            $rect4.Name = "GroupRect4"
            $groupRange = $slide1.Shapes.Range([object[]]@("GroupRect3", "GroupRect4"))
            $group = $groupRange.Group()
            $group.Name = "TestGroup"
        }
        catch { }

        $slide2 = $presentation.Slides.Add(2, 2)
        try { $slide2.Shapes.Title.TextFrame.TextRange.Text = "Content" } catch { }
        try { $slide2.Shapes.Item(2).TextFrame.TextRange.Text = "Alpha Beta Gamma" } catch { }
        try {
            $table = $slide2.Shapes.AddTable(3, 3, 80, 150, 400, 100)
            for ($row = 1; $row -le 3; $row++) {
                for ($column = 1; $column -le 3; $column++) {
                    $table.Cell($row, $column).Shape.TextFrame.TextRange.Text = "R$row" + "C$column"
                }
            }
        }
        catch { }
        try {
            $chart = $slide2.Shapes.AddChart2(-1, 51, 80, 280, 400, 200)
            $chart.Chart.ChartTitle.Text = "TestChart"
        }
        catch { }
        try { $null = $slide2.Shapes.AddPicture($imagePath, $false, $true, 520, 100, 120, 120) } catch { }
        try { $slide1.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = "AutoOfficeMCP speaker notes" } catch { }

        $presentation.SaveAs($pptPath)
        $presentation.SaveCopyAs($pptCopyPath)
        $presentation.Close()
        Release-ComObject $presentation
        $presentation = $null
        $null = $ppt.Quit()
        Release-ComObject $ppt
        $ppt = $null
        Write-Host "FIXTURE ppt: $pptPath"
    }
    finally {
        if ($null -ne $presentation) { try { $presentation.Close() } catch { }; Release-ComObject $presentation }
        if ($null -ne $ppt) { try { $null = $ppt.Quit() } catch { }; Release-ComObject $ppt }
    }
}

Write-Host "FIXTURES READY: $fixtureDir"
