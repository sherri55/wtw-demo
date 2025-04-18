<#
.SYNOPSIS
    Export templates under /sitecore/templates/WTW per-language to CSV.
.DESCRIPTION
    Retrieves every template under the specified folder, then for each language version that exists,
    outputs one row with columns: ID, Name, Path, Language, TemplateID, TemplateName, then all other fields
    (excluding the ones in $excludedFields). The CSV is streamed to the browser for download.
.NOTES
    Test in a non‑production environment and adjust $excludedFields as needed.
#>

# 1. Configuration
$templatesRoot = "/sitecore/templates/WTW"
$excludedFields = @(
    "ID", "Name", "Path", "Template", "__Created", "__Created by",
    "__Updated", "__Owner", "__Revision", "__Renderings"
)
$tempFileName = "WTW_Templates_ByLanguage.csv"

# 2. Retrieve all template items
$templates = Get-Item -Path "master:$templatesRoot" | Get-ChildItem -Recurse

# 3. Build union of all dynamic field names
$allDynamicFields = New-Object System.Collections.Generic.HashSet[string]
foreach ($template in $templates) {
    foreach ($field in $template.Fields) {
        if ($excludedFields -notcontains $field.Name) {
            [void]$allDynamicFields.Add($field.Name)
        }
    }
}
$allDynamicFieldsArray = $allDynamicFields.ToArray() | Sort-Object

# 4. Get all languages in master database
$database = Get-Database "master"
$languages = [Sitecore.Globalization.Language]::GetLanguages($database)

# 5. Build export data
$data = foreach ($template in $templates) {
    foreach ($lang in $languages) {
        # load this template in the given language
        $templLang = Get-Item -Path $template.ID -Language $lang.Name -Version Latest -ErrorAction SilentlyContinue
        if ($templLang -and $templLang.Versions.Count -gt 0) {
            # prepare one record per-language
            $rec = New-Object PSObject

            # fixed columns
            $rec | Add-Member NoteProperty -Name "ID"           -Value $template.ID.ToString()
            $rec | Add-Member NoteProperty -Name "Name"         -Value $templLang.Name
            $rec | Add-Member NoteProperty -Name "Path"         -Value $template.Paths.FullPath
            $rec | Add-Member NoteProperty -Name "Language"     -Value $lang.Name
            $rec | Add-Member NoteProperty -Name "TemplateID"   -Value $template.TemplateID.ToString()

            # resolve base template name
            $baseTpl = Get-Item -Path "master:$($template.TemplateID)" -ErrorAction SilentlyContinue
            $tplName = if ($baseTpl) { $baseTpl.Name } else { "" }
            $rec | Add-Member NoteProperty -Name "TemplateName" -Value $tplName

            # dynamic fields
            foreach ($fn in $allDynamicFieldsArray) {
                $rec | Add-Member NoteProperty -Name $fn -Value $templLang[$fn]
            }

            $rec
        }
    }
}

# 6. Convert to CSV & Download
if ($data.Count -gt 0) {
    $csvContent = $data | ConvertTo-Csv -NoTypeInformation

    $stream = New-Object System.IO.MemoryStream
    $writer = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::UTF8)
    $writer.Write($csvContent -join "`n")
    $writer.Flush()
    $stream.Position = 0

    Out-Download -Name $tempFileName -InputObject $stream
    Write-Host "Export completed. File ready for download: $tempFileName" -ForegroundColor Green
}
else {
    Write-Host "No templates or language versions found to export." -ForegroundColor Yellow
}
