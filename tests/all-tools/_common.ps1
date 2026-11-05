# 全工具冒烟测试共享模块。
# 设计目标：
# 1) 不启动/停止 server，也不改 config.json；默认连接 http://127.0.0.1:8010；
# 2) 每个部分使用自己的 fixture 文件名，支持多个脚本并发执行；
# 3) 每次工具调用前重置 fixture，避免破坏性工具影响后续用例；
# 4) PASS=工具成功；GATED=校验/安全/状态门禁拒绝；FAIL=内部错误或协议失败。

$script:SuiteRoot = $PSScriptRoot

$script:AllowlistedErrorCodes = @(
    "APP_NOT_RUNNING",
    "DOC_NOT_OPEN",
    "FILE_NOT_FOUND",
    "INVALID_RANGE",
    "INVALID_PARAMS",
    "SHEET_NOT_FOUND",
    "SHEET_PROTECTED",
    "DOC_PROTECTED",
    "READ_ONLY",
    "PARTIAL_APPLY",
    "CANCELLED",
    "VBA_DISABLED",
    "CLIPBOARD_DISABLED",
    "OFFICE_BUSY",
    "SERVER_BUSY",
    "EXECUTOR_NOT_AVAILABLE",
    "NOT_SUPPORTED",
    "UNSUPPORTED"
)

function Get-AuthHeaders {
    param([string]$Token)
    $headers = @{}
    if (-not [string]::IsNullOrWhiteSpace($Token)) {
        $headers["Authorization"] = "Bearer $Token"
    }
    return $headers
}

function Invoke-McpRequest {
    param(
        [string]$BaseUrl,
        [object]$Body,
        [string]$SessionId,
        [string]$Token
    )
    $headers = Get-AuthHeaders -Token $Token
    if (-not [string]::IsNullOrWhiteSpace($SessionId)) {
        $headers["Mcp-Session-Id"] = $SessionId
    }
    $json = $Body | ConvertTo-Json -Depth 40 -Compress
    $uri = "$BaseUrl/mcp"
    $response = Invoke-WebRequest -Uri $uri -Method Post -ContentType "application/json" `
        -Headers $headers -Body $json -UseBasicParsing -SkipHttpErrorCheck -TimeoutSec 180
    if ($response.StatusCode -lt 200 -or $response.StatusCode -ge 300) {
        throw "HTTP $($response.StatusCode)：$($response.Content)"
    }
    $newSession = $null
    try {
        $headerValue = $response.Headers["Mcp-Session-Id"]
        if ($headerValue) { $newSession = [string]($headerValue | Select-Object -First 1) }
    }
    catch {
    }
    $parsed = $null
    if (-not [string]::IsNullOrWhiteSpace([string]$response.Content)) {
        $parsed = $response.Content | ConvertFrom-Json
    }
    return [pscustomobject]@{ Json = $parsed; SessionId = $newSession }
}

function Invoke-McpNotify {
    param(
        [string]$BaseUrl,
        [object]$Body,
        [string]$SessionId,
        [string]$Token
    )
    $headers = Get-AuthHeaders -Token $Token
    if (-not [string]::IsNullOrWhiteSpace($SessionId)) {
        $headers["Mcp-Session-Id"] = $SessionId
    }
    $uri = "$BaseUrl/mcp"
    $null = Invoke-WebRequest -Uri $uri -Method Post -ContentType "application/json" `
        -Headers $headers -Body ($Body | ConvertTo-Json -Depth 20 -Compress) -UseBasicParsing `
        -SkipHttpErrorCheck -TimeoutSec 60
}

function Initialize-McpSession {
    param([string]$BaseUrl, [string]$Token)
    $result = Invoke-McpRequest -BaseUrl $BaseUrl -Token $Token -Body @{
        jsonrpc = "2.0"
        id = 1
        method = "initialize"
        params = @{ protocolVersion = "2024-11-05" }
    }
    if ($null -eq $result.Json -or $null -eq $result.Json.result) {
        throw "initialize 未返回有效结果：$BaseUrl"
    }
    if ([string]::IsNullOrWhiteSpace($result.SessionId)) {
        throw "initialize 未返回 Mcp-Session-Id：$BaseUrl"
    }
    Invoke-McpNotify -BaseUrl $BaseUrl -SessionId $result.SessionId -Token $Token -Body @{
        jsonrpc = "2.0"
        method = "notifications/initialized"
        params = @{}
    }
    return $result.SessionId
}

function Get-McpTools {
    param([string]$BaseUrl, [string]$SessionId, [string]$Token)
    $result = Invoke-McpRequest -BaseUrl $BaseUrl -SessionId $SessionId -Token $Token -Body @{
        jsonrpc = "2.0"
        id = 2
        method = "tools/list"
        params = @{}
    }
    if ($null -eq $result.Json -or $null -eq $result.Json.result) {
        throw "tools/list 未返回有效结果：$BaseUrl"
    }
    return @($result.Json.result.tools)
}

function Invoke-McpTool {
    param(
        [string]$BaseUrl,
        [string]$SessionId,
        [string]$Token,
        [string]$Name,
        [object]$Arguments
    )
    $result = Invoke-McpRequest -BaseUrl $BaseUrl -SessionId $SessionId -Token $Token -Body @{
        jsonrpc = "2.0"
        id = 100
        method = "tools/call"
        params = @{ name = $Name; arguments = $Arguments }
    }
    return $result.Json
}

function Get-ServerHealth {
    param([string]$BaseUrl, [string]$Token)
    $response = Invoke-WebRequest -Uri "$BaseUrl/health" -Headers (Get-AuthHeaders -Token $Token) `
        -UseBasicParsing -SkipHttpErrorCheck -TimeoutSec 15
    if ($response.StatusCode -ne 200) {
        throw "server /health 返回 HTTP $($response.StatusCode)：$($response.Content)"
    }
    return $response.Content | ConvertFrom-Json
}

function Get-ToolResponseInfo {
    param($Response)
    $isError = $false
    $text = ""
    try {
        $isError = [bool]$Response.result.isError
        $text = [string]$Response.result.content[0].text
    }
    catch {
    }
    $data = $null
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        $lines = @($text -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($lines.Count -gt 0) {
            try { $data = $lines[$lines.Count - 1] | ConvertFrom-Json } catch { }
        }
    }
    $errorCode = ""
    if ($null -ne $data -and $null -ne $data.PSObject.Properties["errorCode"]) {
        $errorCode = [string]$data.errorCode
    }
    $firstLine = ($text -split "`r?`n")[0]
    return [pscustomobject]@{
        IsError = $isError
        ErrorCode = $errorCode
        Data = $data
        Text = $text
        Message = $firstLine
    }
}

function Get-SuiteContext {
    param(
        [string]$App,
        [int]$Part,
        [int]$Parts,
        [string]$BaseUrl,
        [string]$Token
    )
    $health = Get-ServerHealth -BaseUrl $BaseUrl -Token $Token
    $workspace = [string]$health.workspace
    if ([string]::IsNullOrWhiteSpace($workspace) -or $workspace -eq "不可用") {
        throw "server /health 未返回可用 workspace：$workspace"
    }
    if (-not (Test-Path -LiteralPath $workspace)) {
        throw "workspace 不存在：$workspace"
    }

    $fixtures = Join-Path $script:SuiteRoot "fixtures"
    $prefix = "all-tools-$App-$Part"
    $extension = if ($App -eq "excel") { "xlsx" } elseif ($App -eq "word") { "docx" } else { "pptx" }
    $context = @{
        App = $App
        Part = $Part
        Parts = $Parts
        BaseUrl = $BaseUrl
        Token = $Token
        Workspace = $workspace
        Fixtures = $fixtures
        Prefix = $prefix
        Fixture = "$prefix.$extension"
        CopyFixture = "$prefix-copy.$extension"
        Image = "$prefix.png"
        Csv = "$prefix.csv"
        OutputPrefix = "$prefix-out"
        OutputPath = $null
        OutputPathXlsx = "$prefix-out.xlsx"
        OutputPathCsv = "$prefix-out.csv"
        OutputDirectory = Join-Path $workspace "$prefix-images"
        Sheet1 = "Data"
        Sheet2 = "Calc"
        Range = "Data!A1:C5"
        ChartName = "TestChart"
        WordStyleName = "Normal"
        PptSlideCount = 2
        PptFirstShapeId = 2
        PptSecondShapeId = 3
        PptTableShapeId = 4
        PptPictureShapeId = 7
        PptGroupableShapeIds = @(4, 5)
        PptGroupShapeId = 6
        PptOrder = @(2, 1)
    }
    $context.OutputPath = Join-Path $workspace "$prefix-out.txt"
    return $context
}

function Reset-Fixture {
    param([hashtable]$Context)
    $workspace = $Context.Workspace
    $prefix = $Context.Prefix

    # 只清理本部分的输出文件，不碰 fixture 本体。
    Get-ChildItem -LiteralPath $workspace -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "$($Context.OutputPrefix)*" } |
        Remove-Item -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $Context.OutputDirectory) {
        Remove-Item -LiteralPath $Context.OutputDirectory -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory -Path $Context.OutputDirectory -Force | Out-Null

    $baseName = if ($Context.App -eq "excel") { "base.xlsx" } elseif ($Context.App -eq "word") { "base.docx" } else { "base.pptx" }
    $copyName = if ($Context.App -eq "excel") { "base-copy.xlsx" } elseif ($Context.App -eq "word") { "base-copy.docx" } else { "base-copy.pptx" }
    $sourceFixture = Join-Path $Context.Fixtures $baseName
    $targetFixture = Join-Path $workspace $Context.Fixture
    $copied = $false
    for ($attempt = 0; $attempt -lt 10; $attempt++) {
        try {
            Copy-Item -LiteralPath $sourceFixture -Destination $targetFixture -Force
            $copied = $true
            break
        }
        catch {
            Start-Sleep -Milliseconds 300
        }
    }
    if (-not $copied) {
        throw "无法覆盖 fixture（可能仍被 Office 占用）：$targetFixture"
    }
    Copy-Item -LiteralPath (Join-Path $Context.Fixtures $copyName) -Destination (Join-Path $workspace $Context.CopyFixture) -Force
    Copy-Item -LiteralPath (Join-Path $Context.Fixtures "image.png") -Destination (Join-Path $workspace $Context.Image) -Force
    Copy-Item -LiteralPath (Join-Path $Context.Fixtures "data.csv") -Destination (Join-Path $workspace $Context.Csv) -Force

    if ($Context.App -eq "ppt" -or $Context.App -eq "common") {
        Copy-Item -LiteralPath (Join-Path $Context.Fixtures "base.xlsx") -Destination (Join-Path $workspace "$prefix.xlsx") -Force
        Copy-Item -LiteralPath (Join-Path $Context.Fixtures "base.docx") -Destination (Join-Path $workspace "$prefix.docx") -Force
        Copy-Item -LiteralPath (Join-Path $Context.Fixtures "base.pptx") -Destination (Join-Path $workspace "$prefix.pptx") -Force
    }
    if ($Context.App -eq "common") {
        Copy-Item -LiteralPath (Join-Path $Context.Fixtures "base-copy.pptx") -Destination (Join-Path $workspace "$prefix-copy.pptx") -Force
        $Context.Fixture = "$prefix.pptx"
        $Context.CopyFixture = "$prefix-copy.pptx"
    }
}

function Get-ArgumentValue {
    param(
        [string]$Name,
        [string]$Type,
        [string]$Description,
        [hashtable]$Context
    )
    $lower = $Name.ToLowerInvariant()
    $app = [string]$Context.App

    switch ($lower) {
        "file" { return $Context.Fixture }
        "files" { return @($Context.Fixture, $Context.CopyFixture) }
        "filepath" { if ($app -eq "word") { return $Context.Image } else { return $Context.Fixture } }
        "sourcepath" { return $Context.CopyFixture }
        "originalpath" { return $Context.Fixture }
        "revisedpath" { return $Context.CopyFixture }
        "outputpath" { return $Context.OutputPath }
        "outputdir" { return $Context.OutputDirectory }
        "outputdirectory" { return $Context.OutputDirectory }
        "saveas" { return $Context.OutputPath }
        "path" { return $Context.OutputPath }
        "source" { if ($app -eq "ppt") { return $Context.Image } else { return $Context.Fixture } }
        "excelfile" { return "$($Context.Prefix).xlsx" }
        "sourcefile" { return $Context.CopyFixture }
        "document" { return $Context.Fixture }
        "range" { if ($app -eq "word") { return "all" } else { return $Context.Range } }
        "sourcerange" { return "Data!A1:B3" }
        "targetrange" { return "Data!D1" }
        "datarange" { return "Data!A1:B5" }
        "targetcell" { return "Data!E1" }
        "changingcell" { return "Calc!A2" }
        "sheetname" { return $Context.Sheet1 }
        "sheet1" { return $Context.Sheet1 }
        "sheet2" { return $Context.Sheet2 }
        "sourcename" { return $Context.Sheet1 }
        "oldname" { return $Context.Sheet1 }
        "newname" { return "RenamedSheet" }
        "name" { if ($app -eq "word") { return $Context.WordStyleName } else { return $Context.Sheet1 } }
        "styleName" { return $Context.WordStyleName }
        "stylename" { return $Context.WordStyleName }
        "chartname" { return $Context.ChartName }
        "tablename" { return "TestTable2" }
        "themeFile" { return "test.thmx" }
        "themefile" { return "test.thmx" }
        "type" {
            if ($app -eq "excel") { return "row" }
            if ($app -eq "word") { return "text" }
            return "summary"
        }
        "target" { return "row" }
        "direction" { return "tableToText" }
        "operation" { return "trim" }
        "action" { return "list" }
        "mode" { return "values" }
        "format" {
            if ($app -eq "excel") { return "csv" }
            if ($app -eq "word") { return "txt" }
            return "png"
        }
        "position" {
            if ($app -eq "excel") { return 2 }
            if ($app -eq "ppt") { return "front" }
            return "end"
        }
        "layout" { return "titleAndContent" }
        "transitiontype" { return "fade" }
        "charttype" { return "column" }
        "scope" { return "all" }
        "find" { return "Alpha" }
        "replace" { return "Beta" }
        "findtext" { return "Alpha" }
        "replacetext" { return "Beta" }
        "query" { return "Alpha" }
        "pattern" { return "\d+" }
        "text" { return "AutoOfficeMCP Test" }
        "content" { return "AutoOfficeMCP Test" }
        "title" { return "AutoOfficeMCP Test" }
        "notes" { return "AutoOfficeMCP speaker notes" }
        "code" { return "Sub Test()`r`nEnd Sub" }
        "equation" { return "a^2+b^2=c^2" }
        "formula" { return "=1+1" }
        "formula1" { return "=1" }
        "formula2" { return "=2" }
        "address" { return "https://example.com" }
        "displaytext" { return "Example" }
        "subaddress" { return "Data!A1" }
        "delimiter" { return "," }
        "separator" { return "tab" }
        "encoding" { return "utf8bom" }
        "password" { return "test" }
        "criteria" { return ">0" }
        "order" {
            if ($app -eq "ppt") { return $Context.PptOrder }
            return "asc"
        }
        "paragraphindex" { return 2 }
        "paragraphindexes" { return @(1, 2) }
        "page" { return 1 }
        "slideindex" { return 0 }
        "tableindex" { return 1 }
        "shapeid" { return $Context.PptFirstShapeId }
        "shapeids" { return @($Context.PptFirstShapeId, $Context.PptSecondShapeId) }
        "commentindex" { return 1 }
        "commentindexes" { return @(1) }
        "revisionindex" { return 1 }
        "index" { return 1 }
        "row" { return 1 }
        "col" { return 1 }
        "rows" { return 2 }
        "cols" { return 2 }
        "startrow" { return 1 }
        "startcol" { return 1 }
        "endrow" { return 2 }
        "endcol" { return 2 }
        "fromindex" { return 1 }
        "toindex" { return 2 }
        "column" { return 1 }
        "goalvalue" { return 100 }
        "width" { return 960 }
        "height" { return 540 }
        "x" { return 10 }
        "y" { return 10 }
        "fontsize" { return 12 }
        "enabled" { return $true }
        "hidden" { return $true }
        "protect" { return $true }
        "hasheader" { return $true }
        "filldown" { return $true }
        "createSheet" { return $true }
        "createsheet" { return $true }
        "open" { return $true }
        "readonly" { return $true }
        "savechanges" { return $true }
        "overwrite" { return $true }
        "includeeformulas" { return $true }
        "includeformulas" { return $true }
        "infertypes" { return $true }
        "dryrun" { return $true }
        "confirm" { return $true }
        "commands" {
            return @(@{ tool = "word_page_map"; arguments = @{ file = $Context.Fixture } })
        }
        "items" {
            return @(@{ app = "word"; tool = "word_page_map"; file = $Context.Fixture; savePolicy = "none" })
        }
        "slides" {
            return @(@{ title = "AutoOfficeMCP Test"; content = "Body"; layout = "titleAndContent" })
        }
        "data" { return @(@("Name", "Value"), @("A", 1), @("B", 2)) }
        "values" { return @("A", "B", "C") }
        "mapping" { return @(@{ oldName = "Data"; newName = "DataRenamed" }) }
        "sheets" { return @("Data", "Calc") }
        "rowfields" { return @("Name") }
        "columnfields" { return @() }
        "valuefields" { return @("Q1") }
        "sortkeys" { return @(@{ column = "Name"; order = "asc" }) }
        "sortcolumn" { return 2 }
    }

    switch ($Type) {
        "integer" { return 1 }
        "number" { return 1 }
        "boolean" { return $true }
        "array" { return @() }
        "object" { return @{} }
        default { return "test" }
    }
}

function Get-ToolOverrides {
    param([hashtable]$Context)
    $overrides = @{}
    $app = [string]$Context.App

    if ($app -eq "excel") {
        $overrides["excel_analyze_data"] = @{ type = "summary"; sourceRange = "Data!A1:C11"; targetRange = "H1" }
        $overrides["excel_clean_data"] = @{ range = "Data!A1:C11"; operation = "trim" }
        $overrides["excel_filter_data"] = @{ range = "Data!A1:C11"; column = 1; criteria = ">0" }
        $overrides["excel_paste_special"] = @{ sourceRange = "Data!A1:B3"; targetRange = "Data!E1"; mode = "values" }
        $overrides["excel_print_setup"] = @{ action = "get"; sheetName = "Data" }
        $overrides["excel_protect_sheet"] = @{ sheetName = "Data"; protect = $true; password = "test" }
        $overrides["excel_protect_workbook"] = @{ protect = $true; password = "test" }
        $overrides["excel_table_manage"] = @{ action = "create"; sheetName = "Data"; range = "Data!A1:C5"; tableName = "TestTable2" }
        $overrides["excel_transform_data"] = @{ operation = "transpose"; sourceRange = "Data!A1:B3"; targetRange = "Data!E1" }
        $overrides["excel_conditional_format"] = @{ action = "add"; range = "Data!A1:C5"; rule = "highlight"; condition = "0" }
        $overrides["excel_comment_manage"] = @{ action = "add"; range = "Data!A1"; text = "Test comment" }
        $overrides["excel_hyperlink_manage"] = @{ action = "set"; range = "Data!E2"; address = "https://example.com"; displayText = "Example" }
        $overrides["excel_named_range_manage"] = @{ action = "add"; name = "TestRange2"; refersTo = "=Data!`$A`$1:`$C`$3" }
        $overrides["excel_batch_rename_sheets"] = @{ mode = "regex"; pattern = "^Data`$"; replacement = "DataRenamed"; dryRun = $true }
        $overrides["excel_create_pivot_table"] = @{ sourceRange = "Data!A1:C11"; valueFields = @("Q1"); rowFields = @("Name"); targetSheet = "Pivot" }
        $overrides["excel_create_workbook"] = @{ path = $Context.OutputPathXlsx; sheets = @("Data", "Calc"); open = $true }
        $overrides["excel_export"] = @{ format = "csv"; path = $Context.OutputPathCsv; range = "Data!A1:C5"; overwrite = $true }
        $overrides["excel_import"] = @{ format = "csv"; path = $Context.Csv; targetRange = "Import!A1"; createSheet = $true }
        $overrides["excel_save_as"] = @{ path = $Context.OutputPathXlsx }
        $overrides["excel_run_vba"] = @{ code = "Sub Test()`r`nEnd Sub"; confirm = $true }
        $overrides["excel_delete_sheet"] = @{ name = "Calc"; confirm = $true }
        $overrides["excel_delete_row_col"] = @{ type = "row"; position = 2; confirm = $true }
        $overrides["excel_hide_row_col"] = @{ type = "row"; position = 2 }
        $overrides["excel_insert_row_col"] = @{ type = "row"; position = 2 }
        $overrides["excel_goal_seek"] = @{ targetCell = "Calc!C2"; goalValue = 30; changingCell = "Calc!A2" }
        $overrides["excel_camera_picture"] = @{ sourceRange = "Data!A1:B2"; targetCell = "Format!E1" }
        $overrides["excel_create_chart"] = @{ dataRange = "Data!A1:B5"; type = "column"; position = "Data!H2" }
        $overrides["excel_delete_chart"] = @{ chartName = $Context.ChartName; sheetName = "Calc"; confirm = $true }
        $overrides["excel_update_chart"] = @{ chartName = $Context.ChartName; sheetName = "Calc"; title = "Updated Chart" }
        $overrides["excel_diff_sheets"] = @{ sheet1 = "Data"; sheet2 = "Calc" }
        $overrides["excel_regex_extract"] = @{ range = "Data!A1:A5"; pattern = "\d+" }
        $overrides["excel_rmb_upper"] = @{ range = "Data!B1:B3" }
        $overrides["excel_search_all"] = @{ query = "Alpha" }
        $overrides["excel_find_replace"] = @{ find = "Alpha"; replace = "Beta" }
        $overrides["excel_write_data"] = @{ targetRange = "Data!E1"; data = @(@("A", "B"), @(1, 2)) }
        $overrides["excel_read_range"] = @{ range = "Data!A1:C5" }
        $overrides["excel_apply_formula"] = @{ targetRange = "Calc!C2"; formula = "=A2+B2"; fillDown = $true }
        $overrides["excel_clear_range"] = @{ range = "Data!E1:E2" }
        $overrides["excel_format_range"] = @{ range = "Data!A1:C1"; bold = $true }
        $overrides["excel_merge_cells"] = @{ range = "Data!E1:F1" }
        $overrides["excel_copy_format"] = @{ sourceRange = "Data!A1"; targetRange = "Data!E1" }
        $overrides["excel_autofit"] = @{ range = "Data!A1:C5" }
        $overrides["excel_remove_duplicates"] = @{ range = "Data!A1:C5"; hasHeader = $true }
        $overrides["excel_sort_data"] = @{ range = "Data!A1:C5"; hasHeader = $true; sortKeys = @(@{ column = "Name"; order = "asc" }) }
        $overrides["excel_create_sheet"] = @{ name = "NewSheet" }
        $overrides["excel_copy_sheet"] = @{ sourceName = "Data"; newName = "DataCopy" }
        $overrides["excel_rename_sheet"] = @{ oldName = "Calc"; newName = "CalcRenamed" }
        $overrides["excel_generate_report"] = @{ sourceRange = "Data!A1:C5"; targetSheet = "Report" }
        $overrides["excel_precision_check"] = @{ range = "Calc!A1:C5" }
        $overrides["excel_evaluate"] = @{ formula = "=1+1" }
        $overrides["excel_query_files"] = @{ files = @($Context.Fixture, $Context.CopyFixture); backend = "openxml" }
        $overrides["excel_profile_files"] = @{ files = @($Context.Fixture, $Context.CopyFixture); backend = "openxml" }
        $overrides["excel_sparkline_manage"] = @{ action = "add"; targetRange = "Data!D2:D5"; sourceRange = "Data!A2:C5"; type = "line" }
        $overrides["excel_freeze_and_view"] = @{ freezePane = "A2" }
        $overrides["excel_cross_sheet_consolidate"] = @{ targetSheet = "Summary"; range = "A1:C5" }
        $overrides["excel_data_validation_manage"] = @{ action = "set"; range = "Data!E1"; type = "list"; values = @("A", "B") }
    }
    elseif ($app -eq "word") {
        $overrides["word_apply_as_revision"] = @{ paragraphIndex = 2; original = "Heading One"; corrected = "Heading One Updated" }
        $overrides["word_apply_style"] = @{ styleName = $Context.WordStyleName; paragraphIndex = 2 }
        $overrides["word_beautify"] = @{ margins = @{ top = 2.5; bottom = 2.5; left = 2.5; right = 2.5 } }
        $overrides["word_comment_add"] = @{ text = "Test comment"; paragraphIndex = 2 }
        $overrides["word_comment_delete"] = @{ commentIndexes = @(1); confirm = $true }
        $overrides["word_comment_resolve"] = @{ commentIndex = 1 }
        $overrides["word_compare_documents"] = @{ originalPath = $Context.Fixture; revisedPath = $Context.CopyFixture; outputPath = "$($Context.OutputPrefix)-compare.docx" }
        $overrides["word_content_controls"] = @{ action = "list" }
        $overrides["word_delete_text"] = @{ range = "selection" }
        $overrides["word_doc_merge"] = @{ sourcePath = $Context.CopyFixture }
        $overrides["word_export"] = @{ path = "$($Context.OutputPrefix)-export.txt"; format = "txt"; confirm = $true }
        $overrides["word_extract_images"] = @{ outputDirectory = $Context.OutputDirectory }
        $overrides["word_find_and_select"] = @{ query = "Alpha" }
        $overrides["word_footnote_manage"] = @{ action = "add"; paragraphIndex = 2; text = "Test footnote" }
        $overrides["word_format_text"] = @{ paragraphIndex = 2; bold = $true }
        $overrides["word_get_paragraph_info"] = @{ paragraphIndex = 2 }
        $overrides["word_header_footer"] = @{ headerText = "AutoOfficeMCP Header"; footerText = "AutoOfficeMCP Footer" }
        $overrides["word_goto_page"] = @{ page = 1 }
        $overrides["word_image_layout"] = @{ index = 1; collection = "inline"; wrap = "inline" }
        $overrides["word_image_resize"] = @{ index = 1; collection = "inline"; widthPt = 120 }
        $overrides["word_insert_equation"] = @{ equation = "a^2+b^2=c^2" }
        $overrides["word_insert_image"] = @{ filePath = $Context.Image; paragraphIndex = 2 }
        $overrides["word_insert_table"] = @{ rows = 2; cols = 2; paragraphIndex = 2 }
        $overrides["word_insert_text"] = @{ content = "AutoOfficeMCP Test"; paragraphIndex = 2 }
        $overrides["word_list_paragraphs"] = @{ maxCount = 50 }
        $overrides["word_md_export"] = @{ path = "$($Context.OutputPrefix)-export.md" }
        $overrides["word_md_import"] = @{ markdown = "# AutoOfficeMCP`r`n`r`nTest paragraph." }
        $overrides["word_replace_text"] = @{ find = "Alpha"; replace = "Beta" }
        $overrides["word_revision_accept"] = @{ revisionIndex = 1 }
        $overrides["word_revision_reject"] = @{ revisionIndex = 1 }
        $overrides["word_set_paragraph_format"] = @{ paragraphIndex = 2; bold = $true }
        $overrides["word_style_modify"] = @{ name = $Context.WordStyleName; bold = $true }
        $overrides["word_table_add_row_col"] = @{ tableIndex = 1; rowCount = 1; colCount = 0 }
        $overrides["word_table_delete_row_col"] = @{ tableIndex = 1; target = "row"; index = 2; confirm = $true }
        $overrides["word_table_format"] = @{ tableIndex = 1; borderStyle = "single"; borderColor = "#808080"; autofit = "window"; repeatHeaderRow = $true }
        $overrides["word_table_formula"] = @{ tableIndex = 1; row = 2; col = 3; formula = "=SUM(ABOVE)" }
        $overrides["word_table_merge_cells"] = @{ tableIndex = 1; startRow = 1; startCol = 1; endRow = 1; endCol = 2; confirm = $true }
        $overrides["word_table_read"] = @{ tableIndex = 1 }
        $overrides["word_table_set_cell"] = @{ tableIndex = 1; row = 2; col = 2; text = "Cell" }
        $overrides["word_table_set_col"] = @{ tableIndex = 1; col = 2; values = @("A", "B", "C") }
        $overrides["word_table_set_row"] = @{ tableIndex = 1; row = 2; values = @("A", "B", "C") }
        $overrides["word_track_changes_toggle"] = @{ enabled = $true }
        $overrides["word_watermark"] = @{ action = "add"; text = "AutoOfficeMCP" }
        $overrides["word_section_manage"] = @{ action = "list" }
        $overrides["word_set_page_setup"] = @{ topMarginCm = 2.5; bottomMarginCm = 2.5; leftMarginCm = 2.5; rightMarginCm = 2.5; paperSize = "A4"; orientation = "portrait" }
        $overrides["word_vertical_text"] = @{ action = "get"; target = "shape"; shapeIndex = 1 }
    }
    elseif ($app -eq "ppt") {
        $overrides["ppt_add_animation"] = @{ slideIndex = 0; effect = "fadeIn"; targetShapes = "all" }
        $overrides["ppt_add_speaker_notes"] = @{ slideIndex = 0; notes = "AutoOfficeMCP speaker notes" }
        $overrides["ppt_align_shapes"] = @{ slideIndex = 0; align = "left" }
        $overrides["ppt_apply_theme"] = @{ themeFile = "test.thmx" }
        $overrides["ppt_apply_transition"] = @{ transitionType = "fade"; scope = "all" }
        $overrides["ppt_clear_slide"] = @{ slideIndex = 2; confirm = $true }
        $overrides["ppt_copy_slide_from_file"] = @{ sourceFile = $Context.CopyFixture; sourceSlideIndex = 1; insertAfter = 0 }
        $overrides["ppt_create_slides"] = @{ slides = @(@{ title = "AutoOfficeMCP Test"; content = "Body"; layout = "titleAndContent" }) }
        $overrides["ppt_crop_image"] = @{ shapeId = $Context.PptPictureShapeId; slideIndex = 2; cropLeft = 0.05; cropTop = 0.05; cropRight = 0.05; cropBottom = 0.05 }
        $overrides["ppt_delete_animation"] = @{ slideIndex = 0; animationIndex = 1; confirm = $true }
        $overrides["ppt_delete_shape"] = @{ shapeId = $Context.PptFirstShapeId; slideIndex = 0; confirm = $true }
        $overrides["ppt_delete_slide"] = @{ slideIndex = 2; confirm = $true }
        $overrides["ppt_duplicate_slide"] = @{ slideIndex = 0 }
        $overrides["ppt_execute_vba"] = @{ code = "Sub Test()`r`nEnd Sub"; confirm = $true }
        $overrides["ppt_export_images"] = @{ scope = "all"; outputDir = $Context.OutputDirectory; format = "png" }
        $overrides["ppt_export_pdf"] = @{ path = "$($Context.OutputPrefix)-export.pdf"; confirm = $true }
        $overrides["ppt_format_slide"] = @{ slideIndex = 0 }
        $overrides["ppt_format_table"] = @{ shapeId = $Context.PptTableShapeId; slideIndex = 2; headerBold = $true; headerFillColor = "#1F4E79" }
        $overrides["ppt_format_text"] = @{ slideIndex = 0; shapeId = $Context.PptFirstShapeId; bold = $true }
        $overrides["ppt_get_slide_content"] = @{ slideIndex = 0; includeShapes = $true }
        $overrides["ppt_get_table"] = @{ shapeId = $Context.PptTableShapeId; slideIndex = 2 }
        $overrides["ppt_get_thumbnail"] = @{ slideIndex = 0 }
        $overrides["ppt_get_transition"] = @{ slideIndex = 0 }
        $overrides["ppt_group_shapes"] = @{ slideIndex = 1; shapeIds = $Context.PptGroupableShapeIds }
        $overrides["ppt_hide_slide"] = @{ slideIndex = 2; hidden = $true }
        $overrides["ppt_insert_chart"] = @{ slideIndex = 0; chartType = "column"; data = @(@("Month", "Value"), @("Jan", 100), @("Feb", 120)); title = "Test Chart" }
        $overrides["ppt_insert_chart_from_excel"] = @{ slideIndex = 0; excelFile = "$($Context.Prefix).xlsx"; range = "Data!A1:B5" }
        $overrides["ppt_insert_image"] = @{ slideIndex = 0; source = $Context.Image }
        $overrides["ppt_insert_shape"] = @{ slideIndex = 0; shapeType = "rectangle"; x = 10; y = 10; width = 100; height = 50 }
        $overrides["ppt_insert_slide"] = @{ position = "end"; layout = "titleAndContent"; title = "AutoOfficeMCP Test"; content = "Body" }
        $overrides["ppt_insert_table"] = @{ slideIndex = 0; rows = 2; cols = 2; data = @(@("A", "B"), @("1", "2")) }
        $overrides["ppt_insert_table_from_excel"] = @{ slideIndex = 0; excelFile = "$($Context.Prefix).xlsx"; range = "Data!A1:B3" }
        $overrides["ppt_insert_text"] = @{ slideIndex = 0; content = "AutoOfficeMCP Test" }
        $overrides["ppt_list_animations"] = @{ slideIndex = 0 }
        $overrides["ppt_list_shapes"] = @{ slideIndex = 0 }
        $overrides["ppt_move_resize_shape"] = @{ slideIndex = 0; shapeId = $Context.PptFirstShapeId; x = 20; y = 20; width = 120; height = 60 }
        $overrides["ppt_move_slide"] = @{ fromIndex = 1; toIndex = 2 }
        $overrides["ppt_reorder_slides"] = @{ order = $Context.PptOrder; confirm = $true }
        $overrides["ppt_replace_image"] = @{ shapeId = $Context.PptPictureShapeId; slideIndex = 2; source = $Context.Image }
        $overrides["ppt_set_bullet"] = @{ slideIndex = 0; shapeId = $Context.PptFirstShapeId; enabled = $true }
        $overrides["ppt_set_image_format"] = @{ shapeId = $Context.PptPictureShapeId; slideIndex = 2; borderColor = "#FFFFFF"; borderWidth = 2 }
        $overrides["ppt_set_paragraph"] = @{ slideIndex = 0; shapeId = $Context.PptFirstShapeId; align = "left" }
        $overrides["ppt_set_shape_format"] = @{ slideIndex = 0; shapeId = $Context.PptFirstShapeId; fillColor = "#2E86DE" }
        $overrides["ppt_set_shape_text"] = @{ slideIndex = 0; shapeId = $Context.PptFirstShapeId; content = "AutoOfficeMCP Test" }
        $overrides["ppt_set_slide_layout"] = @{ slideIndex = 0; layout = "titleAndContent" }
        $overrides["ppt_set_slide_size"] = @{ width = 960; height = 540; confirm = $true }
        $overrides["ppt_set_table_cell"] = @{ shapeId = $Context.PptTableShapeId; slideIndex = 2; row = 1; col = 1; text = "Cell" }
        $overrides["ppt_set_zorder"] = @{ slideIndex = 0; shapeId = $Context.PptFirstShapeId; position = "front" }
        $overrides["ppt_ungroup_shapes"] = @{ slideIndex = 1; shapeId = $Context.PptGroupShapeId }
        $overrides["ppt_validate_layout"] = @{ slideIndex = 0 }
    }
    elseif ($app -eq "common") {
        $overrides["office_open_document"] = @{ path = $Context.Fixture; app = "ppt"; readOnly = $false }
        $overrides["office_close_document"] = @{ document = $Context.Fixture; app = "ppt"; saveChanges = $true; confirm = $true }
        $overrides["office_save_document"] = @{ document = $Context.Fixture; app = "ppt"; saveAs = "$($Context.OutputPrefix)-saveas.pptx" }
        $overrides["office_save_all"] = @{}
        $overrides["office_activate_document"] = @{ document = $Context.Fixture; app = "ppt" }
        $overrides["office_batch"] = @{ confirm = $true; commands = @(@{ tool = "ppt_get_presentation_info"; arguments = @{ file = $Context.Fixture } }) }
        $overrides["office_batch_files"] = @{ confirm = $true; items = @(@{ app = "ppt"; tool = "ppt_get_presentation_info"; file = $Context.CopyFixture; savePolicy = "none" }) }
        $overrides["office_undo"] = @{ steps = 1 }
        $overrides["office_undo_list"] = @{ limit = 10 }
        $overrides["office_undo_last"] = @{}
        $overrides["office_stub_probe"] = @{}
        $overrides["office_get_window_info"] = @{ app = "ppt" }
        $overrides["office_list_documents"] = @{}
        $overrides["office_pool_status"] = @{}
        $overrides["office_screenshot"] = @{ app = "ppt" }
    }
    return $overrides
}

function Build-ToolArguments {
    param(
        [string]$Tool,
        [object]$Schema,
        [hashtable]$Context,
        [hashtable]$Overrides
    )
    $toolArgs = @{}
    if ($Overrides.ContainsKey($Tool)) {
        foreach ($key in $Overrides[$Tool].Keys) {
            $toolArgs[$key] = $Overrides[$Tool][$key]
        }
    }

    $properties = @{}
    if ($null -ne $Schema -and $null -ne $Schema.properties) {
        foreach ($property in $Schema.properties.PSObject.Properties) {
            $properties[$property.Name] = $property.Value
        }
    }
    $required = @()
    if ($null -ne $Schema -and $null -ne $Schema.required) {
        $required = @($Schema.required)
    }

    foreach ($name in $required) {
        if ($toolArgs.ContainsKey($name)) { continue }
        $definition = $properties[$name]
        $type = if ($null -ne $definition -and $null -ne $definition.type) { [string]$definition.type } else { "string" }
        $description = if ($null -ne $definition -and $null -ne $definition.description) { [string]$definition.description } else { "" }
        $toolArgs[$name] = Get-ArgumentValue -Name $name -Type $type -Description $description -Context $Context
    }

    # 对可选但决定测试有效性的路径参数补齐。
    foreach ($name in @(
        "file", "files", "filePath", "sourcePath", "originalPath", "revisedPath",
        "outputPath", "outputDir", "outputDirectory", "path", "saveAs", "source",
        "excelFile", "sourceFile", "range", "sourceRange", "targetRange", "dataRange",
        "targetCell", "sheetName", "sheet1", "sheet2", "slideIndex", "paragraphIndex",
        "tableIndex"
    )) {
        if ($toolArgs.ContainsKey($name)) { continue }
        if (-not $properties.ContainsKey($name)) { continue }
        $definition = $properties[$name]
        $type = if ($null -ne $definition -and $null -ne $definition.type) { [string]$definition.type } else { "string" }
        $description = if ($null -ne $definition -and $null -ne $definition.description) { [string]$definition.description } else { "" }
        $toolArgs[$name] = Get-ArgumentValue -Name $name -Type $type -Description $description -Context $Context
    }

    # Excel/Word/PPT 直连工具依赖隐式 file（描述符通常不声明该参数）。
    if ($Context.App -in @("excel", "word", "ppt") -and -not $toolArgs.ContainsKey("file")) {
        $toolArgs["file"] = $Context.Fixture
    }
    if (-not $toolArgs.ContainsKey("confirm")) {
        $toolArgs["confirm"] = $true
    }
    return $toolArgs
}

function Open-PptFixture {
    param([string]$BaseUrl, [string]$SessionId, [string]$Token, [hashtable]$Context)
    try {
        $null = Invoke-McpTool -BaseUrl $BaseUrl -SessionId $SessionId -Token $Token `
            -Name "office_open_document" -Arguments @{ path = $Context.Fixture; app = "ppt"; readOnly = $false }
    }
    catch {
    }
}

function Close-PptFixture {
    param([string]$BaseUrl, [string]$SessionId, [string]$Token, [hashtable]$Context)
    try {
        $null = Invoke-McpTool -BaseUrl $BaseUrl -SessionId $SessionId -Token $Token `
            -Name "office_close_document" -Arguments @{
                document = $Context.Fixture
                app = "ppt"
                saveChanges = $true
                confirm = $true
            }
    }
    catch {
    }
}

function Prepare-AppContext {
    param(
        [string]$BaseUrl,
        [string]$SessionId,
        [string]$Token,
        [hashtable]$Context
    )
    if ($Context.App -eq "ppt") {
        try {
            $info = Invoke-McpTool -BaseUrl $BaseUrl -SessionId $SessionId -Token $Token `
                -Name "ppt_get_presentation_info" -Arguments @{ file = $Context.Fixture }
            $infoData = (Get-ToolResponseInfo -Response $info).Data
            if ($null -ne $infoData -and [int]$infoData.slideCount -ge 1) {
                $Context.PptSlideCount = [int]$infoData.slideCount
            }
            $slides = @(1, 2)
            $firstSlideShapeIds = @()
            $groupableIds = @()
            $groupShapeId = $Context.PptGroupShapeId
            $tableId = $Context.PptTableShapeId
            $pictureId = $Context.PptPictureShapeId
            foreach ($slideIndex in $slides) {
                $shapesResponse = Invoke-McpTool -BaseUrl $BaseUrl -SessionId $SessionId -Token $Token `
                    -Name "ppt_list_shapes" -Arguments @{ file = $Context.Fixture; slideIndex = $slideIndex }
                $shapesData = (Get-ToolResponseInfo -Response $shapesResponse).Data
                if ($null -eq $shapesData) { continue }
                foreach ($shape in @($shapesData.shapes)) {
                    $id = [int]$shape.id
                    if ($id -le 0) { continue }
                    if ($slideIndex -eq 1) { $firstSlideShapeIds += $id }
                    if ([string]$shape.name -like "GroupRect*") { $groupableIds += $id }
                    if ([string]$shape.name -eq "TestGroup" -or [string]$shape.type -like "*Group*") { $groupShapeId = $id }
                    if ([bool]$shape.hasTable) { $tableId = $id }
                    if ([string]$shape.type -like "*Picture*" -or [string]$shape.name -like "*Picture*") { $pictureId = $id }
                }
            }
            if ($firstSlideShapeIds.Count -ge 1) { $Context.PptFirstShapeId = $firstSlideShapeIds[0] }
            if ($firstSlideShapeIds.Count -ge 2) { $Context.PptSecondShapeId = $firstSlideShapeIds[1] }
            if ($groupableIds.Count -ge 2) { $Context.PptGroupableShapeIds = $groupableIds }
            $Context.PptGroupShapeId = $groupShapeId
            $Context.PptTableShapeId = $tableId
            $Context.PptPictureShapeId = $pictureId
            if ($Context.PptSlideCount -ge 2) {
                $Context.PptOrder = @(2) + @(1..($Context.PptSlideCount - 1) | Where-Object { $_ -ne 2 })
            }
            else {
                $Context.PptOrder = @(1)
            }
        }
        catch {
            Write-Host "[prepare] PPT 形状探测失败，使用默认 ID：$($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    elseif ($Context.App -eq "word") {
        try {
            $stylesResponse = Invoke-McpTool -BaseUrl $BaseUrl -SessionId $SessionId -Token $Token `
                -Name "word_styles_list" -Arguments @{ file = $Context.Fixture; type = "paragraph" }
            $stylesData = (Get-ToolResponseInfo -Response $stylesResponse).Data
            if ($null -ne $stylesData) {
                $styles = @($stylesData.styles)
                $normal = @($styles | Where-Object { [string]$_.name -eq "Normal" -or [string]$_.name -eq "正文" })[0]
                if ($null -eq $normal) {
                    $normal = @($styles | Where-Object { [int]$_.inUseCount -gt 0 })[0]
                }
                if ($null -eq $normal) { $normal = @($styles)[0] }
                if ($null -ne $normal -and -not [string]::IsNullOrWhiteSpace([string]$normal.name)) {
                    $Context.WordStyleName = [string]$normal.name
                }
            }
        }
        catch {
            Write-Host "[prepare] Word 样式探测失败，使用默认样式名。" -ForegroundColor Yellow
        }
    }
}

function Invoke-AllToolsPart {
    param(
        [string]$App,
        [int]$Part,
        [int]$Parts,
        [string]$BaseUrl,
        [string]$Token,
        [int]$Limit = 0,
        [switch]$KeepFixtures
    )
    $context = Get-SuiteContext -App $App -Part $Part -Parts $Parts -BaseUrl $BaseUrl -Token $Token
    $sessionId = Initialize-McpSession -BaseUrl $BaseUrl -Token $Token
    $allTools = Get-McpTools -BaseUrl $BaseUrl -SessionId $sessionId -Token $Token

    $prefix = switch ($App) {
        "excel" { "excel_" }
        "word" { "word_" }
        "ppt" { "ppt_" }
        "common" { "office_" }
        default { "" }
    }
    $tools = @($allTools | Where-Object { [string]$_.name -like "$prefix*" } | Sort-Object { [string]$_.name })
    $selected = @()
    for ($index = 0; $index -lt $tools.Count; $index++) {
        if (($index % $Parts) -eq ($Part - 1)) {
            $selected += $tools[$index]
        }
    }
    if ($Limit -gt 0 -and $selected.Count -gt $Limit) {
        $selected = @($selected | Select-Object -First $Limit)
    }

    Reset-Fixture -Context $context
    if ($App -eq "ppt") { Open-PptFixture -BaseUrl $BaseUrl -SessionId $sessionId -Token $Token -Context $context }
    Prepare-AppContext -BaseUrl $BaseUrl -SessionId $sessionId -Token $Token -Context $context
    if ($App -eq "ppt") { Close-PptFixture -BaseUrl $BaseUrl -SessionId $sessionId -Token $Token -Context $context }
    $overrides = Get-ToolOverrides -Context $context

    $records = New-Object System.Collections.ArrayList
    $pass = 0
    $gated = 0
    $fail = 0

    foreach ($tool in $selected) {
        $name = [string]$tool.name
        $status = "FAIL"
        $errorCode = ""
        $message = ""
        $detail = ""
        $commonNeedsDocument = $App -eq "common" -and $name -notin @(
            "office_open_document",
            "office_list_documents", "office_pool_status", "office_stub_probe"
        )
        $pptNeedsPresentation = $App -eq "ppt"
        try {
            if ($pptNeedsPresentation) {
                Close-PptFixture -BaseUrl $BaseUrl -SessionId $sessionId -Token $Token -Context $context
            }
            if ($App -eq "common") {
                Close-PptFixture -BaseUrl $BaseUrl -SessionId $sessionId -Token $Token -Context $context
            }
            Reset-Fixture -Context $context
            if ($pptNeedsPresentation) {
                Open-PptFixture -BaseUrl $BaseUrl -SessionId $sessionId -Token $Token -Context $context
            }
            if ($commonNeedsDocument) {
                try {
                    $null = Invoke-McpTool -BaseUrl $BaseUrl -SessionId $sessionId -Token $Token `
                        -Name "office_open_document" -Arguments @{ path = $context.Fixture; app = "ppt"; readOnly = $false }
                }
                catch {
                }
            }
            try {
                $arguments = Build-ToolArguments -Tool $name -Schema $tool.inputSchema -Context $context -Overrides $overrides
                $response = Invoke-McpTool -BaseUrl $BaseUrl -SessionId $sessionId -Token $Token -Name $name -Arguments $arguments
                $info = Get-ToolResponseInfo -Response $response
                $errorCode = $info.ErrorCode
                $message = $info.Message
                $detail = $info.Text
                if (-not $info.IsError) {
                    $status = "PASS"
                    $pass++
                }
                elseif ($script:AllowlistedErrorCodes -contains $info.ErrorCode) {
                    $status = "GATED"
                    $gated++
                }
                else {
                    $status = "FAIL"
                    $fail++
                }
            }
            finally {
                if ($commonNeedsDocument -and $name -ne "office_close_document") {
                    try {
                        $null = Invoke-McpTool -BaseUrl $BaseUrl -SessionId $sessionId -Token $Token `
                            -Name "office_close_document" -Arguments @{
                                document = $context.Fixture
                                app = "ppt"
                                saveChanges = $true
                                confirm = $true
                            }
                    }
                    catch {
                    }
                }
                if ($pptNeedsPresentation) {
                    Close-PptFixture -BaseUrl $BaseUrl -SessionId $sessionId -Token $Token -Context $context
                }
            }
        }
        catch {
            $status = "FAIL"
            $message = $_.Exception.Message
            $detail = $message
            $fail++
        }
        $color = if ($status -eq "PASS") { "Green" } elseif ($status -eq "GATED") { "Yellow" } else { "Red" }
        Write-Host ("[{0}] {1} {2} {3}" -f $status, $App, $name, $errorCode) -ForegroundColor $color
        [void]$records.Add([pscustomobject]@{
            tool = $name
            status = $status
            errorCode = $errorCode
            message = $message
            detail = $detail
        })
    }

    $resultDir = Join-Path $script:SuiteRoot "results"
    New-Item -ItemType Directory -Path $resultDir -Force | Out-Null
    $reportPath = Join-Path $resultDir ("{0}-part{1}.json" -f $App, $Part)
    $report = [pscustomobject]@{
        app = $App
        part = $Part
        parts = $Parts
        total = $records.Count
        pass = $pass
        gated = $gated
        fail = $fail
        generatedAt = (Get-Date).ToString("o")
        tools = $records
    }
    $report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $reportPath -Encoding utf8NoBOM

    Write-Host ""
    Write-Host ("SUMMARY {0} part {1}/{2}: total={3} pass={4} gated={5} fail={6}" -f $App, $Part, $Parts, $records.Count, $pass, $gated, $fail) `
        -ForegroundColor $(if ($fail -eq 0) { "Green" } else { "Red" })
    Write-Host "REPORT $reportPath"
    return [pscustomobject]@{ Pass = $pass; Gated = $gated; Fail = $fail; Total = $records.Count; Report = $reportPath }
}
