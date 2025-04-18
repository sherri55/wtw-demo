<#
.SYNOPSIS
    Export Renderings, Layouts, and Placeholder Settings per-language to CSV.
.DESCRIPTION
    Processes items under specified roots and, for each existing language version,
    emits one row with columns: ID, Name, Path, Language, then all other fields
    (excluding those in $excludedFields). Offers the CSV for download.
#>

# 1. Configuration
$roots = @(
    "/sitecore/Layout/Layouts/WTW",
    "/sitecore/Layout/Renderings/WTW",
    "/sitecore/Layout/Placeholder Settings/Feature",
    "/sitecore/Layout/Placeholder Settings/Foundation"
)
$excludedFields = @("ID", "Name", "Path")   # fields to omit from dynamic columns
$tempFileName = "WTW_Renderings_Layouts_Placeholders_ByLanguage.csv"

# 2. Retrieve all items under each root
$allItems = foreach ($root in $roots) {
    $ri = Get-Item -Path "master:$root" -ErrorAction SilentlyContinue
    if ($ri) { $ri | Get-ChildItem -Recurse }
}

# 3. Collect all dynamic field names
$allDynamicFields = New-Object System.Collections.Generic.HashSet[string]
foreach ($item in $allItems) {
    foreach ($field in $item.Fields) {
        if ($excludedFields -notcontains $field.Name) {
            [void]$allDynamicFields.Add($field.Name)
        }
    }
}
$allDynamicFields = $allDynamicFields | Sort-Object

# 4. Get all languages in master DB
$database = Get-Database "master"
$languages = [Sitecore.Globalization.Language]::GetLanguages($database) 

# 5. Build export data
$data = foreach ($item in $allItems) {
    foreach ($lang in $languages) {
        $itemLang = Get-Item -Path $item.ID -Language $lang.Name -Version Latest -ErrorAction SilentlyContinue
        if ($itemLang -and $itemLang.Versions.Count -gt 0) {
            $rec = New-Object PSObject
            # fixed columns
            $rec | Add-Member NoteProperty ID       ($item.ID.ToString())
            $rec | Add-Member NoteProperty Name     ($item.Name)
            $rec | Add-Member NoteProperty Path     ($item.Paths.FullPath)
            $rec | Add-Member NoteProperty Language ($lang.Name)
            # dynamic fields
            foreach ($fn in $allDynamicFields) {
                $rec | Add-Member NoteProperty $fn ($itemLang[$fn])
            }
            $rec
        }
    }
}

# 6. Export to CSV & Download
if ($data.Count -gt 0) {
    $csvContent = $data | ConvertTo-Csv -NoTypeInformation
    $stream = New-Object System.IO.MemoryStream
    $writer = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::UTF8)
    $writer.Write($csvContent -join "`n"); $writer.Flush(); $stream.Position = 0
    Out-Download -Name $tempFileName -InputObject $stream
    Write-Host "Export ready:" $tempFileName -ForegroundColor Green
}
else {
    Write-Host "No items or language versions found." -ForegroundColor Yellow
}
