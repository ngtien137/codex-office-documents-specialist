[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ActionsPath,
    [string]$Path,
    [string]$OutPath,
    [string]$PdfPath,
    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-JsonValue {
    param(
        [Parameter(Mandatory = $false)]$Object,
        [Parameter(Mandatory = $true)][string]$Name,
        $Default = $null
    )

    if ($null -eq $Object) {
        return $Default
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $Default
    }

    return $property.Value
}

function Resolve-AbsolutePath {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    if ([System.IO.Path]::IsPathRooted($Value)) {
        return [System.IO.Path]::GetFullPath($Value)
    }

    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Value))
}

function Ensure-ParentDirectory {
    param([string]$TargetPath)

    if ([string]::IsNullOrWhiteSpace($TargetPath)) {
        return
    }

    $parent = Split-Path -Parent $TargetPath
    if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent | Out-Null
    }
}

function Move-SelectionToEnd {
    param($Selection)
    [void]$Selection.EndKey(6)
}

function Set-SelectionStyle {
    param(
        $Selection,
        $Style
    )

    if ($null -eq $Style -or [string]::IsNullOrWhiteSpace([string]$Style)) {
        return
    }

    $Selection.Style = $Style
}

function Get-HeadingStyleId {
    param([int]$Level)

    $normalized = [Math]::Max(1, [Math]::Min(9, $Level))
    return -1 - $normalized
}

function Add-Paragraph {
    param(
        $Selection,
        [string]$Text,
        $Style = $null
    )

    Move-SelectionToEnd -Selection $Selection
    Set-SelectionStyle -Selection $Selection -Style $Style
    $Selection.TypeText($Text)
    $Selection.TypeParagraph()
}

function Add-Heading {
    param(
        $Selection,
        [string]$Text,
        [int]$Level
    )

    Move-SelectionToEnd -Selection $Selection
    $Selection.Style = Get-HeadingStyleId -Level $Level
    $Selection.TypeText($Text)
    $Selection.TypeParagraph()
}

function Add-Table {
    param(
        $Document,
        $Selection,
        $Action
    )

    $rows = @(Get-JsonValue -Object $Action -Name "rows" -Default @())
    if ($rows.Count -eq 0) {
        throw "insert_table requires a non-empty rows array."
    }

    $columnCount = 0
    foreach ($row in $rows) {
        $columnCount = [Math]::Max($columnCount, @($row).Count)
    }

    Move-SelectionToEnd -Selection $Selection
    $table = $Document.Tables.Add($Selection.Range, $rows.Count, $columnCount)

    $style = Get-JsonValue -Object $Action -Name "style"
    if (-not [string]::IsNullOrWhiteSpace([string]$style)) {
        $table.Style = $style
    }

    $headerRow = [bool](Get-JsonValue -Object $Action -Name "headerRow" -Default $false)
    if ($headerRow) {
        $table.Rows.Item(1).Range.Bold = 1
    }

    for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
        $currentRow = @($rows[$rowIndex])
        for ($columnIndex = 0; $columnIndex -lt $currentRow.Count; $columnIndex++) {
            $table.Cell($rowIndex + 1, $columnIndex + 1).Range.Text = [string]$currentRow[$columnIndex]
        }
    }

    switch (([string](Get-JsonValue -Object $Action -Name "autoFit" -Default "content")).ToLowerInvariant()) {
        "fixed" { [void]$table.AutoFitBehavior(0) }
        "window" { [void]$table.AutoFitBehavior(2) }
        default { [void]$table.AutoFitBehavior(1) }
    }

    Move-SelectionToEnd -Selection $Selection
    $Selection.TypeParagraph()
}

function Replace-Text {
    param(
        $Document,
        $Action
    )

    $findText = [string](Get-JsonValue -Object $Action -Name "find")
    if ([string]::IsNullOrWhiteSpace($findText)) {
        throw "replace_text requires a non-empty find value."
    }

    $replaceText = [string](Get-JsonValue -Object $Action -Name "replace" -Default "")
    $matchCase = [bool](Get-JsonValue -Object $Action -Name "matchCase" -Default $false)
    $wholeWord = [bool](Get-JsonValue -Object $Action -Name "wholeWord" -Default $false)
    $range = $Document.Content
    $find = $range.Find
    $find.ClearFormatting()
    $find.Replacement.ClearFormatting()

    [void]$find.Execute(
        $findText,
        $matchCase,
        $wholeWord,
        $false,
        $false,
        $false,
        $true,
        1,
        $false,
        $replaceText,
        2
    )
}

function Find-FirstRange {
    param(
        $Document,
        [string]$FindText,
        [bool]$MatchCase,
        [bool]$WholeWord
    )

    $range = $Document.Content
    $find = $range.Find
    $find.ClearFormatting()

    $found = $find.Execute(
        $FindText,
        $MatchCase,
        $WholeWord,
        $false,
        $false,
        $false,
        $true,
        1,
        $false
    )

    if ($found) {
        return $range
    }

    return $null
}

function Add-Comment {
    param(
        $Document,
        $Action
    )

    $findText = [string](Get-JsonValue -Object $Action -Name "find")
    $commentText = [string](Get-JsonValue -Object $Action -Name "comment")
    if ([string]::IsNullOrWhiteSpace($findText) -or [string]::IsNullOrWhiteSpace($commentText)) {
        throw "add_comment requires both find and comment values."
    }

    $matchCase = [bool](Get-JsonValue -Object $Action -Name "matchCase" -Default $false)
    $wholeWord = [bool](Get-JsonValue -Object $Action -Name "wholeWord" -Default $false)
    $range = Find-FirstRange -Document $Document -FindText $findText -MatchCase $matchCase -WholeWord $wholeWord
    if ($null -eq $range) {
        throw "Could not find text for comment: $findText"
    }

    [void]$Document.Comments.Add($range, $commentText)
}

function Set-MarginsCm {
    param(
        $Word,
        $Document,
        $Action
    )

    $pageSetup = $Document.PageSetup
    foreach ($field in @("Top", "Bottom", "Left", "Right")) {
        $value = Get-JsonValue -Object $Action -Name $field.ToLowerInvariant()
        if ($null -ne $value) {
            $pageSetup."${field}Margin" = $Word.CentimetersToPoints([double]$value)
        }
    }
}

function Set-Orientation {
    param(
        $Document,
        [string]$Orientation
    )

    $value = switch ($Orientation.ToLowerInvariant()) {
        "landscape" { 1 }
        default { 0 }
    }

    foreach ($section in $Document.Sections) {
        $section.PageSetup.Orientation = $value
    }
}

function Set-HeaderFooterText {
    param(
        $Document,
        [ValidateSet("Header", "Footer")][string]$Kind,
        [string]$Text
    )

    foreach ($section in $Document.Sections) {
        $collection = if ($Kind -eq "Header") { $section.Headers } else { $section.Footers }
        $item = $collection.Item(1)
        if ($null -ne $item) {
            if ($section.Index -gt 1) {
                $item.LinkToPrevious = $false
            }
            $item.Range.Text = $Text
        }
    }
}

function Add-PageNumbers {
    param(
        $Document,
        $Action
    )

    $alignmentName = ([string](Get-JsonValue -Object $Action -Name "alignment" -Default "right")).ToLowerInvariant()
    $alignment = switch ($alignmentName) {
        "left" { 0 }
        "center" { 1 }
        default { 2 }
    }

    $firstPage = [bool](Get-JsonValue -Object $Action -Name "firstPage" -Default $true)
    foreach ($section in $Document.Sections) {
        [void]$section.Footers.Item(1).PageNumbers.Add($alignment, $firstPage)
    }
}

function Ensure-Toc {
    param(
        $Document,
        $Action
    )

    if ($Document.TablesOfContents.Count -gt 0) {
        return
    }

    $upper = [int](Get-JsonValue -Object $Action -Name "upperHeadingLevel" -Default 1)
    $lower = [int](Get-JsonValue -Object $Action -Name "lowerHeadingLevel" -Default 3)
    $range = $Document.Range(0, 0)
    [void]$Document.TablesOfContents.Add($range, $true, $upper, $lower)
}

function Update-Toc {
    param($Document)

    foreach ($toc in $Document.TablesOfContents) {
        [void]$toc.Update()
    }
}

function Update-Fields {
    param($Document)
    [void]$Document.Fields.Update()
}

function Set-TrackRevisions {
    param(
        $Document,
        [bool]$Enabled
    )

    $Document.TrackRevisions = $Enabled
}

function Accept-AllRevisions {
    param($Document)
    if ($Document.Revisions.Count -gt 0) {
        [void]$Document.Revisions.AcceptAll()
    }
}

function Reject-AllRevisions {
    param($Document)
    if ($Document.Revisions.Count -gt 0) {
        [void]$Document.Revisions.RejectAll()
    }
}

$actionsPathResolved = Resolve-AbsolutePath -Value $ActionsPath
if (-not (Test-Path -LiteralPath $actionsPathResolved)) {
    throw "Actions file not found: $actionsPathResolved"
}

$config = Get-Content -Raw -LiteralPath $actionsPathResolved | ConvertFrom-Json
$sourcePath = Resolve-AbsolutePath -Value $(if ($PSBoundParameters.ContainsKey("Path")) { $Path } else { Get-JsonValue -Object $config -Name "source" })
$saveTarget = Resolve-AbsolutePath -Value $(if ($PSBoundParameters.ContainsKey("OutPath")) { $OutPath } else { Get-JsonValue -Object $config -Name "saveAs" })
$pdfTarget = Resolve-AbsolutePath -Value $(if ($PSBoundParameters.ContainsKey("PdfPath")) { $PdfPath } else { Get-JsonValue -Object $config -Name "exportPdf" })
$visibleFlag = if ($PSBoundParameters.ContainsKey("Visible")) { $Visible.IsPresent } else { [bool](Get-JsonValue -Object $config -Name "visible" -Default $false) }
$actionList = @(Get-JsonValue -Object $config -Name "actions" -Default @())

if (-not $sourcePath -and -not $saveTarget) {
    throw "New documents require saveAs or -OutPath."
}

if ($sourcePath -and -not (Test-Path -LiteralPath $sourcePath)) {
    throw "Source document not found: $sourcePath"
}

Ensure-ParentDirectory -TargetPath $saveTarget
Ensure-ParentDirectory -TargetPath $pdfTarget

$word = $null
$document = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $visibleFlag
    $word.DisplayAlerts = 0

    if ($sourcePath) {
        $document = $word.Documents.Open($sourcePath)
        Write-Output "Opened document: $sourcePath"
    }
    else {
        $document = $word.Documents.Add()
        Write-Output "Created new document"
    }

    $selection = $word.Selection

    foreach ($action in $actionList) {
        $type = ([string](Get-JsonValue -Object $action -Name "type")).ToLowerInvariant()
        switch ($type) {
            "replace_text" {
                Replace-Text -Document $document -Action $action
                Write-Output "Applied replace_text"
            }
            "append_heading" {
                Add-Heading -Selection $selection -Text ([string](Get-JsonValue -Object $action -Name "text" -Default "")) -Level ([int](Get-JsonValue -Object $action -Name "level" -Default 1))
                Write-Output "Applied append_heading"
            }
            "append_paragraph" {
                Add-Paragraph -Selection $selection -Text ([string](Get-JsonValue -Object $action -Name "text" -Default "")) -Style (Get-JsonValue -Object $action -Name "style")
                Write-Output "Applied append_paragraph"
            }
            "append_page_break" {
                Move-SelectionToEnd -Selection $selection
                [void]$selection.InsertBreak(7)
                Write-Output "Applied append_page_break"
            }
            "insert_table" {
                Add-Table -Document $document -Selection $selection -Action $action
                Write-Output "Applied insert_table"
            }
            "add_comment" {
                Add-Comment -Document $document -Action $action
                Write-Output "Applied add_comment"
            }
            "set_margins_cm" {
                Set-MarginsCm -Word $word -Document $document -Action $action
                Write-Output "Applied set_margins_cm"
            }
            "set_orientation" {
                Set-Orientation -Document $document -Orientation ([string](Get-JsonValue -Object $action -Name "orientation" -Default "portrait"))
                Write-Output "Applied set_orientation"
            }
            "set_header_text" {
                Set-HeaderFooterText -Document $document -Kind Header -Text ([string](Get-JsonValue -Object $action -Name "text" -Default ""))
                Write-Output "Applied set_header_text"
            }
            "set_footer_text" {
                Set-HeaderFooterText -Document $document -Kind Footer -Text ([string](Get-JsonValue -Object $action -Name "text" -Default ""))
                Write-Output "Applied set_footer_text"
            }
            "add_page_numbers" {
                Add-PageNumbers -Document $document -Action $action
                Write-Output "Applied add_page_numbers"
            }
            "create_toc" {
                Ensure-Toc -Document $document -Action $action
                Write-Output "Applied create_toc"
            }
            "update_toc" {
                Update-Toc -Document $document
                Write-Output "Applied update_toc"
            }
            "update_fields" {
                Update-Fields -Document $document
                Write-Output "Applied update_fields"
            }
            "set_track_revisions" {
                Set-TrackRevisions -Document $document -Enabled ([bool](Get-JsonValue -Object $action -Name "enabled" -Default $true))
                Write-Output "Applied set_track_revisions"
            }
            "accept_all_revisions" {
                Accept-AllRevisions -Document $document
                Write-Output "Applied accept_all_revisions"
            }
            "reject_all_revisions" {
                Reject-AllRevisions -Document $document
                Write-Output "Applied reject_all_revisions"
            }
            default {
                throw "Unsupported action type: $type"
            }
        }
    }

    if ($saveTarget) {
        if ($sourcePath -and ($saveTarget -ieq $sourcePath)) {
            [void]$document.Save()
            Write-Output "Saved document in place: $saveTarget"
        }
        else {
            [void]$document.SaveAs2($saveTarget)
            Write-Output "Saved document copy: $saveTarget"
        }
    }
    elseif ($sourcePath) {
        [void]$document.Save()
        Write-Output "Saved document in place: $sourcePath"
    }

    if ($pdfTarget) {
        [void]$document.ExportAsFixedFormat($pdfTarget, 17)
        Write-Output "Exported PDF: $pdfTarget"
    }
}
finally {
    if ($null -ne $document) {
        [void]$document.Close()
    }
    if ($null -ne $word) {
        [void]$word.Quit()
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
