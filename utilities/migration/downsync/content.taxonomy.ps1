<#
.SYNOPSIS
    Export Taxonomy items (sitecore/content/WTW/Taxonomy) per-language to CSV.
.DESCRIPTION
    Retrieves every item under /sitecore/content/WTW/Taxonomy, and for each
    language version that exists, outputs one row with columns:
    ID, Name, Path, Language, then all other fields.
#>

# 1. Configuration
$taxonomyRoot = "/sitecore/content/WTW/Taxonomy"
$excludedFields = @("ID", "Name", "Path")
$tempFileName = "WTW_Taxonomy_ByLanguage.csv"

# 2. Retrieve all taxonomy items
$taxonomyItems = Get-Item -Path "master:$taxonomyRoot" | Get-ChildItem -Recurse

# 3. Collect dynamic fields
$allDynamicFields = New-Object System.Collections.Generic.HashSet[string]
foreach ($item in $taxonomyItems) {
    foreach ($field in $item.Fields) {
        if ($excludedFields -notcontains $field.Name) {
            [void]$allDynamicFields.Add($field.Name)
        }
    }
}
$allDynamicFields = $allDynamicFields.ToArray() | Sort-Object

# 4. Get languages
$availableLanguages = [Sitecore.Data.Managers.LanguageManager]::GetLanguages([Sitecore.Context]::ContentDatabase)

# 5. Build export data
$data = foreach ($item in $taxonomyItems) {
    foreach ($lang in $availableLanguages) {
        $itemLang = Get-Item -Path $item.ID -Language $lang.Name -Version Latest -ErrorAction SilentlyContinue
        if ($itemLang -and $itemLang.Versions.Count -gt 0) {
            $rec = New-Object PSObject
            $rec | Add-Member NoteProperty ID       ($item.ID.ToString())
            $rec | Add-Member NoteProperty Name     ($item.Name)
            $rec | Add-Member NoteProperty Path     ($item.Paths.FullPath)
            $rec | Add-Member NoteProperty Language ($lang.Name)
            foreach ($fn in $allDynamicFields) {
                $rec | Add-Member NoteProperty $fn ($itemLang[$fn])
            }
            $rec
        }
    }
}

# 6. Export & Download
if ($data.Count -gt 0) {
    $csvContent = $data | ConvertTo-Csv -NoTypeInformation
    $stream = New-Object System.IO.MemoryStream
    $writer = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::UTF8)
    $writer.Write($csvContent -join "`n"); $writer.Flush(); $stream.Position = 0
    Out-Download -Name $tempFileName -InputObject $stream
    Write-Host "Export ready:" $tempFileName -ForegroundColor Green
}
else {
    Write-Host "No taxonomy items or language versions found." -ForegroundColor Yellow
}
