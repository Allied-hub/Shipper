param(
    [Parameter(Mandatory = $true)]
    [string]$PayloadPath,

    [Parameter(Mandatory = $true)]
    [string]$TemplatePath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$WorkbookStructurePassword = $env:WORKBOOK_STRUCTURE_PASSWORD,

    [switch]$Visible
)

$ErrorActionPreference = "Stop"
trap {
    Write-Error ("Linea {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message)
    break
}

$HeaderAliases = @{
    "DRAWING #" = "DWG #"
    "DRAWING NO" = "DWG #"
    "DRAWING NO." = "DWG #"
    "DWG" = "DWG #"
    "WT" = "WT."
    "WEIGHT" = "WT."
}

$FieldHeaders = [ordered]@{
    qty = "QTY"
    mark = "MARK"
    desc = "DESCRIPTION"
    pitch = "PITCH"
    part = "PART"
    punch = "PUNCH"
    dwg = "DWG #"
    color = "COLOR"
    length = "LENGTH"
    wt = "WT."
}

$KnownWorkbookStructurePasswords = @(
    $WorkbookStructurePassword,
    "a0:aap"
) | Where-Object { -not [string]::IsNullOrEmpty($_) } | Select-Object -Unique

$AppliedWorkbookStructurePassword = $null

function Normalize-Header {
    param($Value)
    if ($null -eq $Value) { return "" }
    $text = ([string]$Value).Trim().ToUpperInvariant() -replace "\s+", " "
    if ($text.EndsWith(":")) {
        $text = $text.Substring(0, $text.Length - 1).Trim()
    }
    if ($HeaderAliases.ContainsKey($text)) {
        return $HeaderAliases[$text]
    }
    return $text
}

function To-Int {
    param($Value)
    if ($Value -is [System.Array]) {
        $Value = $Value[0]
    }
    return [int]$Value
}

function To-ExcelValue {
    param($Value)
    if ($Value -is [System.Array]) {
        if ($Value.Count -eq 0) { return "" }
        $Value = $Value[0]
    }
    if ($null -eq $Value) { return "" }
    return $Value
}

function Get-MergeAreaOrNull {
    param($Cell)
    try {
        if ($Cell.MergeCells) {
            $area = $Cell.MergeArea
            if ($area -is [System.Array]) {
                $area = $area[0]
            }
            return $area
        }
    } catch {
        return $null
    }
    return $null
}

function Set-RangeValue {
    param($Range, $Value)
    $flags = [Reflection.BindingFlags]::SetProperty
    if ($Value -is [string]) {
        try {
            $Range.GetType().InvokeMember("NumberFormat", $flags, $null, $Range, @("@")) | Out-Null
        } catch {}
    }
    try {
        $Range.GetType().InvokeMember("Value2", $flags, $null, $Range, @($Value)) | Out-Null
    } catch {
        $Range.GetType().InvokeMember("Value", $flags, $null, $Range, @($Value)) | Out-Null
    }
}

function Set-CellSafe {
    param($Worksheet, [int]$Row, [int]$Column, $Value)
    $Value = To-ExcelValue $Value
    $cell = $Worksheet.Cells.Item($Row, $Column)
    $mergeArea = Get-MergeAreaOrNull $cell
    try {
        Set-RangeValue $cell $Value
    } catch {
        try {
            if ($null -ne $mergeArea) {
                Set-RangeValue ($mergeArea.Cells.Item(1)) $Value
            } else {
                throw
            }
        } catch {
            $valueType = if ($null -eq $Value) { "null" } else { $Value.GetType().FullName }
            $cellType = if ($null -eq $cell) { "null" } else { $cell.GetType().FullName }
            $mergeType = if ($null -eq $mergeArea) { "null" } else { $mergeArea.GetType().FullName }
            throw "Set-CellSafe fallo en '$($Worksheet.Name)' R$Row C$Column valor=[$Value] tipo=[$valueType] cell=[$cellType] merge=[$mergeType]: $($_.Exception.Message)"
        }
    }
}

function Get-ValueCellAfterLabel {
    param($Cell)
    $mergeArea = Get-MergeAreaOrNull $Cell
    if ($null -ne $mergeArea) {
        return @{
            Row = To-Int $mergeArea.Row
            Column = (To-Int $mergeArea.Column) + (To-Int $mergeArea.Columns.Count)
        }
    }
    return @{
        Row = To-Int $Cell.Row
        Column = (To-Int $Cell.Column) + 1
    }
}

function Find-HeaderRow {
    param($Worksheet)
    $used = $Worksheet.UsedRange
    $maxCol = [Math]::Max(60, (To-Int $used.Column) + (To-Int $used.Columns.Count) - 1)
    for ($row = 1; $row -le 40; $row++) {
        for ($col = 1; $col -le $maxCol; $col++) {
            if ((Normalize-Header $Worksheet.Cells.Item($row, $col).Value2) -eq "QTY") {
                return $row
            }
        }
    }
    throw "No se encontro fila de encabezados QTY en '$($Worksheet.Name)'"
}

function Get-ColumnMap {
    param($Worksheet, [int]$HeaderRow)
    $used = $Worksheet.UsedRange
    $maxCol = [Math]::Max(60, (To-Int $used.Column) + (To-Int $used.Columns.Count) - 1)
    $map = @{}
    for ($col = 1; $col -le $maxCol; $col++) {
        $header = Normalize-Header $Worksheet.Cells.Item($HeaderRow, $col).Value2
        if ($header -and -not $map.ContainsKey($header)) {
            $map[$header] = $col
        }
    }
    return $map
}

function Get-HeaderMergeSpecs {
    param($Worksheet, [int]$HeaderRow)
    $used = $Worksheet.UsedRange
    $maxCol = [Math]::Max(60, (To-Int $used.Column) + (To-Int $used.Columns.Count) - 1)
    $specs = @{}

    for ($col = 1; $col -le $maxCol; $col++) {
        $cell = $Worksheet.Cells.Item($HeaderRow, $col)
        $header = Normalize-Header $cell.Value2
        if (-not $header -or $specs.ContainsKey($header)) { continue }

        $mergeArea = Get-MergeAreaOrNull $cell
        if ($null -ne $mergeArea) {
            $startCol = To-Int $mergeArea.Column
            $endCol = $startCol + (To-Int $mergeArea.Columns.Count) - 1
        } else {
            $startCol = $col
            $endCol = $col
        }

        $specs[$header] = @{
            Start = $startCol
            End = $endCol
        }
    }

    return $specs
}

function Get-LastHeaderColumn {
    param($Columns, $MergeSpecs)
    $lastCol = 1

    foreach ($col in $Columns.Values) {
        $lastCol = [Math]::Max($lastCol, (To-Int $col))
    }
    foreach ($spec in $MergeSpecs.Values) {
        $lastCol = [Math]::Max($lastCol, (To-Int $spec.End))
    }

    return $lastCol
}

function Merge-CellSpan {
    param($Worksheet, [int]$Row, [int]$StartCol, [int]$EndCol)
    if ($EndCol -le $StartCol) { return }

    $range = $Worksheet.Range($Worksheet.Cells.Item($Row, $StartCol), $Worksheet.Cells.Item($Row, $EndCol))
    try {
        $range.Merge() | Out-Null
    } catch {
        throw "No se pudo fusionar '$($Worksheet.Name)' R$Row C$StartCol:C$EndCol`: $($_.Exception.Message)"
    }
}

function Apply-FieldMerges {
    param($Worksheet, [int]$Row, $MergeSpecs)

    foreach ($header in $FieldHeaders.Values) {
        if (-not $MergeSpecs.ContainsKey($header)) { continue }
        $spec = $MergeSpecs[$header]
        Merge-CellSpan $Worksheet $Row (To-Int $spec.Start) (To-Int $spec.End)
    }
}

function Get-ObjectProperty {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object.PSObject.Properties.Name -contains $Name) {
        return $Object.$Name
    }
    return $null
}

function Clear-RowMerges {
    param($Worksheet, [int]$Row, [int]$LastCol)
    $range = $Worksheet.Range($Worksheet.Cells.Item($Row, 1), $Worksheet.Cells.Item($Row, $LastCol))
    try { $range.UnMerge() | Out-Null } catch {}
}

function Copy-RowFormat {
    param($Worksheet, [int]$TemplateRow, [int]$TargetRow, [int]$LastCol)
    if ($TemplateRow -le 0 -or $TargetRow -le 0) { return }

    $source = $Worksheet.Range($Worksheet.Cells.Item($TemplateRow, 1), $Worksheet.Cells.Item($TemplateRow, $LastCol))
    $target = $Worksheet.Range($Worksheet.Cells.Item($TargetRow, 1), $Worksheet.Cells.Item($TargetRow, $LastCol))
    try {
        $source.Copy() | Out-Null
        # -4122 = xlPasteFormats
        $target.PasteSpecial(-4122) | Out-Null
    } catch {
        try {
            $source.Copy($target) | Out-Null
            $target.ClearContents() | Out-Null
        } catch {}
    }
    try { $Worksheet.Rows.Item($TargetRow).RowHeight = $Worksheet.Rows.Item($TemplateRow).RowHeight } catch {}
    try { $Worksheet.Application.CutCopyMode = $false } catch {}
}

function Find-WeightRow {
    param($Worksheet, [int]$HeaderRow)
    $used = $Worksheet.UsedRange
    $maxRow = [Math]::Max($HeaderRow + 20, (To-Int $used.Row) + (To-Int $used.Rows.Count) - 1)
    $maxCol = [Math]::Max(30, (To-Int $used.Column) + (To-Int $used.Columns.Count) - 1)

    for ($row = $HeaderRow + 2; $row -le $maxRow; $row++) {
        for ($col = 1; $col -le $maxCol; $col++) {
            $text = Normalize-Header $Worksheet.Cells.Item($row, $col).Value2
            if ($text -like "*WEIGHT*") {
                return $row
            }
        }
    }
    return 0
}

function Get-PieceRowCount {
    param($Piece)
    $count = 1
    if ($Piece.PSObject.Properties.Name -contains "detalles" -and $null -ne $Piece.detalles) {
        $count += @($Piece.detalles).Count
    } elseif ($Piece.PSObject.Properties.Name -contains "detalle" -and $null -ne $Piece.detalle -and [string]$Piece.detalle -ne "") {
        $count += 1
    }
    return $count
}

function Get-NeededDataRows {
    param($Piezas)
    $rows = 0
    foreach ($piece in @($Piezas)) {
        $rows += Get-PieceRowCount $piece
    }
    return $rows
}

function Ensure-DataCapacity {
    param($Worksheet, [int]$DataStart, [int]$WeightRow, [int]$NeededRows)
    if ($WeightRow -le 0) { return $WeightRow }

    $capacity = $WeightRow - $DataStart - 1
    if ($NeededRows -le $capacity) { return $WeightRow }

    $extraRows = $NeededRows - $capacity
    for ($i = 0; $i -lt $extraRows; $i++) {
        $Worksheet.Rows.Item($WeightRow).Insert() | Out-Null
    }
    return $WeightRow + $extraRows
}

function Find-DetailTemplateRow {
    param($Worksheet, [int]$DataStart, [int]$WeightRow, [int]$LastCol)
    if ($WeightRow -le $DataStart) { return 0 }

    for ($row = $DataStart; $row -lt $WeightRow; $row++) {
        for ($col = 1; $col -le $LastCol; $col++) {
            $cell = $Worksheet.Cells.Item($row, $col)
            $mergeArea = Get-MergeAreaOrNull $cell
            if ($null -eq $mergeArea) { continue }
            $mergeStart = To-Int $mergeArea.Column
            $mergeEnd = $mergeStart + (To-Int $mergeArea.Columns.Count) - 1
            if ($mergeStart -eq 1 -and $mergeEnd -ge $LastCol) {
                return $row
            }
        }
    }
    return $DataStart
}

function Write-Header {
    param($Worksheet, $Encabezado, [int]$ShipperNumber, [int]$TotalShippers)
    $used = $Worksheet.UsedRange
    $maxCol = [Math]::Max(20, (To-Int $used.Column) + (To-Int $used.Columns.Count) - 1)

    for ($row = 1; $row -le 7; $row++) {
        for ($col = 1; $col -le $maxCol; $col++) {
            $cell = $Worksheet.Cells.Item($row, $col)
            $text = Normalize-Header $cell.Value2
            if (-not $text) { continue }

            $target = Get-ValueCellAfterLabel $cell
            if ($text -like "*SHIPPER NUMBER*") {
                Set-CellSafe $Worksheet $target.Row $target.Column $ShipperNumber
            } elseif ($text -eq "OF") {
                Set-CellSafe $Worksheet $target.Row $target.Column $TotalShippers
            } elseif ($text -like "*JOB NUMBER*" -and $null -ne (Get-ObjectProperty $Encabezado "job")) {
                Set-CellSafe $Worksheet $target.Row $target.Column (Get-ObjectProperty $Encabezado "job")
            } elseif ($text -like "*ISSUE DATE*" -and $null -ne (Get-ObjectProperty $Encabezado "fecha")) {
                Set-CellSafe $Worksheet $target.Row $target.Column (Get-ObjectProperty $Encabezado "fecha")
            } elseif ($text -like "*BUILDING NUMBER*" -and $null -ne (Get-ObjectProperty $Encabezado "edificio")) {
                Set-CellSafe $Worksheet $target.Row $target.Column (Get-ObjectProperty $Encabezado "edificio")
            } elseif ($text -like "*BLDG DESCRIP*" -and $null -ne (Get-ObjectProperty $Encabezado "descrip")) {
                Set-CellSafe $Worksheet $target.Row $target.Column (Get-ObjectProperty $Encabezado "descrip")
            } elseif ($text -like "*CUSTOMER*" -and $null -ne (Get-ObjectProperty $Encabezado "cliente")) {
                Set-CellSafe $Worksheet $target.Row $target.Column (Get-ObjectProperty $Encabezado "cliente")
            }
        }
    }
}

function Write-RawRows {
    param($Worksheet, $RawRows)

    $rawRows = @($RawRows)
    if ($rawRows.Count -eq 0) { return $false }

    $headerRow = Find-HeaderRow $Worksheet
    $columns = Get-ColumnMap $Worksheet $headerRow
    $mergeSpecs = Get-HeaderMergeSpecs $Worksheet $headerRow
    $lastHeaderCol = Get-LastHeaderColumn $columns $mergeSpecs
    $dataStart = $headerRow + 2
    $weightRow = Find-WeightRow $Worksheet $headerRow
    $pieceTemplateRow = $dataStart
    $detailTemplateRow = Find-DetailTemplateRow $Worksheet $dataStart $weightRow $lastHeaderCol
    $blankTemplateRow = $headerRow + 1

    $targetEndRow = $headerRow + $rawRows.Count - 1
    if ($weightRow -gt 0 -and $targetEndRow -gt $weightRow) {
        $extraRows = $targetEndRow - $weightRow
        for ($i = 0; $i -lt $extraRows; $i++) {
            $Worksheet.Rows.Item($weightRow).Insert() | Out-Null
        }
        $weightRow += $extraRows
    }
    if ($weightRow -le 0) {
        $weightRow = $targetEndRow
    }

    $used = $Worksheet.UsedRange
    $lastCol = [Math]::Max($lastHeaderCol, (To-Int $used.Column) + (To-Int $used.Columns.Count) - 1)
    $clearEndRow = [Math]::Max($targetEndRow, $weightRow)
    $clearRange = $Worksheet.Range($Worksheet.Cells.Item($headerRow, 1), $Worksheet.Cells.Item($clearEndRow, $lastCol))
    $clearRange.UnMerge() | Out-Null
    $clearRange.ClearContents() | Out-Null

    for ($i = 0; $i -lt $rawRows.Count; $i++) {
        $rowData = $rawRows[$i]
        $row = $headerRow + $i
        $kind = [string]$rowData.kind

        $templateRow = $pieceTemplateRow
        if ($kind -eq "header") {
            $templateRow = $headerRow
        } elseif ($kind -eq "blank") {
            $templateRow = $blankTemplateRow
        } elseif ($kind -eq "detail") {
            $templateRow = $detailTemplateRow
        } elseif ($kind -eq "weight") {
            $templateRow = $weightRow
        }

        Copy-RowFormat $Worksheet $templateRow $row $lastCol
        Clear-RowMerges $Worksheet $row $lastCol

        $col = 1
        foreach ($cellData in @($rowData.cells)) {
            $span = if ($null -ne $cellData.span) { To-Int $cellData.span } else { 1 }
            $value = $cellData.value
            $endCol = $col + $span - 1
            Merge-CellSpan $Worksheet $row $col $endCol
            Set-CellSafe $Worksheet $row $col $value
            $col = $endCol + 1
        }
    }

    return $true
}

function Clear-WorksheetData {
    param($Worksheet)
    try {
        $headerRow = Find-HeaderRow $Worksheet
    } catch {
        return
    }

    $dataStart = $headerRow + 2
    $used = $Worksheet.UsedRange
    $lastRow = [Math]::Max($dataStart, (To-Int $used.Row) + (To-Int $used.Rows.Count) - 1)
    $lastCol = [Math]::Max(30, (To-Int $used.Column) + (To-Int $used.Columns.Count) - 1)

    $range = $Worksheet.Range($Worksheet.Cells.Item($dataStart, 1), $Worksheet.Cells.Item($lastRow, $lastCol))
    $range.ClearContents() | Out-Null
}

function Write-Pieces {
    param($Worksheet, $Piezas, $Peso)

    $headerRow = Find-HeaderRow $Worksheet
    $columns = Get-ColumnMap $Worksheet $headerRow
    $mergeSpecs = Get-HeaderMergeSpecs $Worksheet $headerRow
    $dataStart = $headerRow + 2
    $lastHeaderCol = Get-LastHeaderColumn $columns $mergeSpecs
    $weightRow = Find-WeightRow $Worksheet $headerRow
    $pieceTemplateRow = $dataStart
    $detailTemplateRow = Find-DetailTemplateRow $Worksheet $dataStart $weightRow $lastHeaderCol

    $neededRows = Get-NeededDataRows $Piezas
    $weightRow = Ensure-DataCapacity $Worksheet $dataStart $weightRow $neededRows

    $used = $Worksheet.UsedRange
    $lastCol = [Math]::Max($lastHeaderCol, (To-Int $used.Column) + (To-Int $used.Columns.Count) - 1)
    $clearEndRow = if ($weightRow -gt $dataStart) { $weightRow - 1 } else { [Math]::Max($dataStart, (To-Int $used.Row) + (To-Int $used.Rows.Count) - 1) }
    $dataRange = $Worksheet.Range($Worksheet.Cells.Item($dataStart, 1), $Worksheet.Cells.Item($clearEndRow, $lastCol))
    $dataRange.UnMerge() | Out-Null
    $dataRange.ClearContents() | Out-Null

    $detailCol = if ($columns.ContainsKey("QTY")) { To-Int $columns["QTY"] } else { 1 }
    $row = $dataStart

    foreach ($piece in @($Piezas)) {
        Copy-RowFormat $Worksheet $pieceTemplateRow $row $lastCol
        Clear-RowMerges $Worksheet $row $lastCol
        Apply-FieldMerges $Worksheet $row $mergeSpecs

        foreach ($field in $FieldHeaders.Keys) {
            $header = $FieldHeaders[$field]
            if ($columns.ContainsKey($header)) {
                $value = $piece.$field
                if ($null -eq $value) { $value = "" }
                Set-CellSafe $Worksheet $row (To-Int $columns[$header]) $value
            }
        }
        $row++

        $detalles = @()
        if ($piece.PSObject.Properties.Name -contains "detalles" -and $null -ne $piece.detalles) {
            $detalles = @($piece.detalles)
        } elseif ($piece.PSObject.Properties.Name -contains "detalle" -and $null -ne $piece.detalle -and [string]$piece.detalle -ne "") {
            $detalles = @($piece.detalle)
        }

        foreach ($detalle in $detalles) {
            Copy-RowFormat $Worksheet $detailTemplateRow $row $lastCol
            Clear-RowMerges $Worksheet $row $lastCol
            Merge-CellSpan $Worksheet $row $detailCol $lastHeaderCol
            Set-CellSafe $Worksheet $row $detailCol $detalle
            $detailRange = $Worksheet.Range($Worksheet.Cells.Item($row, $detailCol), $Worksheet.Cells.Item($row, $lastHeaderCol))
            try { $detailRange.NumberFormat = "@" } catch {}
            try { $detailRange.HorizontalAlignment = -4131 } catch {}
            try { $detailRange.WrapText = $false } catch {}
            $row++
        }
    }

    if ($null -ne $Peso -and [string]$Peso -ne "") {
        if ($weightRow -gt 0) {
            $row = $weightRow
        } else {
            $row++
        }
        $wtCol = if ($columns.ContainsKey("WT.")) { To-Int $columns["WT."] } else { 10 }
        if ($mergeSpecs.ContainsKey("LENGTH")) {
            $labelSpec = $mergeSpecs["LENGTH"]
            $labelCol = To-Int $labelSpec.Start
            Merge-CellSpan $Worksheet $row (To-Int $labelSpec.Start) (To-Int $labelSpec.End)
        } else {
            $labelCol = [Math]::Max(1, $wtCol - 2)
        }
        if ($mergeSpecs.ContainsKey("WT.")) {
            $wtSpec = $mergeSpecs["WT."]
            $wtCol = To-Int $wtSpec.Start
            Merge-CellSpan $Worksheet $row (To-Int $wtSpec.Start) (To-Int $wtSpec.End)
        }
        $existingLabel = $Worksheet.Cells.Item($row, $labelCol).Text
        if ($null -eq $existingLabel -or [string]$existingLabel -eq "") {
            Set-CellSafe $Worksheet $row $labelCol "PAGE WEIGHT:"
        }
        Set-CellSafe $Worksheet $row $wtCol $Peso
    }
}

function Get-WorksheetOrClone {
    param($Workbook, $SheetData, [string]$StructurePassword, $AfterWorksheet = $null)

    $sheetName = [string]$SheetData.tab_macro
    try {
        return $Workbook.Worksheets.Item($sheetName)
    } catch {}

    $templateName = $sheetName
    if ($SheetData.PSObject.Properties.Name -contains "template_tab_macro" -and $null -ne $SheetData.template_tab_macro) {
        $templateName = [string]$SheetData.template_tab_macro
    }

    try {
        $template = $Workbook.Worksheets.Item($templateName)
    } catch {
        throw "No existe pestana '$sheetName' ni plantilla '$templateName'"
    }

    if ([bool]$Workbook.ProtectStructure) {
        $passwordCandidates = @($StructurePassword) + $script:KnownWorkbookStructurePasswords
        $passwordCandidates = $passwordCandidates | Where-Object { -not [string]::IsNullOrEmpty($_) } | Select-Object -Unique
        if ($passwordCandidates.Count -eq 0) {
            throw "No se puede crear la pestana '$sheetName' porque la estructura del workbook esta protegida. Define WORKBOOK_STRUCTURE_PASSWORD para permitir crear pestanas cuando una seccion supera 38 filas."
        }

        $lastError = $null
        foreach ($password in $passwordCandidates) {
            try {
                $Workbook.Unprotect($password) | Out-Null
                if (-not [bool]$Workbook.ProtectStructure) {
                    $script:AppliedWorkbookStructurePassword = $password
                    break
                }
            } catch {
                $lastError = $_.Exception.Message
            }
        }

        if ([bool]$Workbook.ProtectStructure) {
            throw "No se pudo desproteger la estructura del workbook para crear la pestana '$sheetName'. Revisa WORKBOOK_STRUCTURE_PASSWORD: $lastError"
        }
    }

    $sheetCount = To-Int $Workbook.Worksheets.Count
    $template.Copy([System.Reflection.Missing]::Value, $Workbook.Worksheets.Item($sheetCount)) | Out-Null
    $created = $Workbook.Worksheets.Item($sheetCount + 1)
    try {
        if ([string]$created.Name -ne $sheetName) {
            $created.Name = $sheetName
        }
    } catch {
        try {
            $created = $Workbook.Worksheets.Item($sheetName)
        } catch {
            throw "No se pudo nombrar la pestana creada como '$sheetName': $($_.Exception.Message)"
        }
    }

    # Move the cloned tab to come right after $AfterWorksheet so it appears
    # in payload order instead of at the end of the workbook.
    if ($null -ne $AfterWorksheet) {
        try {
            $created.Move([System.Reflection.Missing]::Value, $AfterWorksheet) | Out-Null
        } catch {}
    }

    Clear-WorksheetData $created
    return $created
}

function Remove-WorksheetIfExists {
    param($Workbook, [string]$SheetName)

    try {
        $worksheet = $Workbook.Worksheets.Item($SheetName)
    } catch {
        return
    }

    if ((To-Int $Workbook.Worksheets.Count) -le 1) { return }

    try {
        $worksheet.Delete() | Out-Null
    } catch {
        throw "No se pudo eliminar la pestana '$SheetName': $($_.Exception.Message)"
    }
}

if (-not (Test-Path -LiteralPath $PayloadPath)) {
    throw "No existe payload: $PayloadPath"
}
if (-not (Test-Path -LiteralPath $TemplatePath)) {
    throw "No existe template .xls: $TemplatePath"
}
$templateInfo = Get-Item -LiteralPath $TemplatePath
if ($templateInfo.Length -le 0) {
    throw "El template .xls esta vacio: $TemplatePath"
}
if ([IO.Path]::GetExtension($OutputPath).ToLowerInvariant() -ne ".xls") {
    throw "El output debe terminar en .xls: $OutputPath"
}
if ((Test-Path -LiteralPath $OutputPath) -and ((Resolve-Path -LiteralPath $OutputPath).Path -eq (Resolve-Path -LiteralPath $TemplatePath).Path)) {
    throw "OutputPath no puede ser el mismo archivo que TemplatePath"
}

$payload = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
$outputDir = Split-Path -Parent $OutputPath
if ($outputDir) {
    New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
}

Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath -Force
try { Unblock-File -LiteralPath $OutputPath -ErrorAction SilentlyContinue } catch {}

$excel = $null
$workbook = $null
$workbookStructureWasProtected = $false

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Open($OutputPath, 0, $false, 5, "", "", $true)
    $workbookStructureWasProtected = [bool]$workbook.ProtectStructure
    Write-Output "Se conservan las pestanas Screws/Screw."

    $targetSheets = @{}
    foreach ($sheetData in @($payload.sheets)) {
        $targetSheets[[string]$sheetData.tab_macro] = $true
    }

    foreach ($sheetName in @("Cold Form Members (ZEE) (2)", "Cold Form Members (ZEE) (3)")) {
        if (-not $targetSheets.ContainsKey($sheetName)) {
            try {
                $worksheet = $workbook.Worksheets.Item($sheetName)
                Clear-WorksheetData $worksheet
            } catch {}
        }
    }

    $prevWorksheet = $null
    foreach ($sheetData in @($payload.sheets)) {
        $sheetName = [string]$sheetData.tab_macro
        $worksheet = Get-WorksheetOrClone $workbook $sheetData $WorkbookStructurePassword $prevWorksheet
        Write-Header $worksheet $sheetData.encabezado (To-Int $sheetData.shipper_number) (To-Int $sheetData.total_shippers)
        $wroteRawRows = $false
        if ($sheetData.PSObject.Properties.Name -contains "raw_rows" -and $null -ne $sheetData.raw_rows) {
            $wroteRawRows = Write-RawRows $worksheet $sheetData.raw_rows
        }
        if (-not $wroteRawRows) {
            Write-Pieces $worksheet $sheetData.piezas $sheetData.peso
        }
        $prevWorksheet = $worksheet
    }

    if ($null -ne $payload.job_number) {
        try {
            $cover = $workbook.Worksheets.Item("Cover")
            $used = $cover.UsedRange
            $maxRow = (To-Int $used.Row) + (To-Int $used.Rows.Count) - 1
            $maxCol = (To-Int $used.Column) + (To-Int $used.Columns.Count) - 1
            for ($row = 1; $row -le $maxRow; $row++) {
                for ($col = 1; $col -le $maxCol; $col++) {
                    $cell = $cover.Cells.Item($row, $col)
                    $text = Normalize-Header $cell.Value2
                    $target = Get-ValueCellAfterLabel $cell
                    if ($text -like "*JOB NUMBER*") {
                        Set-CellSafe $cover $target.Row $target.Column $payload.job_number
                    } elseif ($text -like "*TOTAL NUMBER*") {
                        Set-CellSafe $cover $target.Row $target.Column (To-Int $payload.files_processed)
                    }
                }
            }
        } catch {
            Write-Warning "No se pudo actualizar Cover: $($_.Exception.Message)"
        }
    }

    if ($workbookStructureWasProtected -and -not [string]::IsNullOrEmpty($AppliedWorkbookStructurePassword) -and -not [bool]$workbook.ProtectStructure) {
        $workbook.Protect($AppliedWorkbookStructurePassword, $true, $false) | Out-Null
    }

    $workbook.Save()
    Write-Output "Guardado: $OutputPath"
} finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($null -ne $excel) {
        try { $excel.EnableEvents = $true } catch {}
        try { $excel.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
