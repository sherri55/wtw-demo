# Parameters
$rootItemPath = '/sitecore/content/BTLaw/Home/Insights/Alerts'
$tempFileName = 'Insight-AlertsData.csv'
$languages = @('en', 'ja-jp', 'zh-cn')  # Specify the languages to process
$filterDate = Get-Date '2025-02-14'  # Filter items updated after this date

# Load the root item
$rootItem = Get-Item -Path $rootItemPath

if (-Not $rootItem) {
    Write-Host 'Root item not found at path: $rootItemPath' -ForegroundColor Red
    return
}

# Initialize a collection for CSV data
$data = @()

# Process each language
foreach ($language in $languages) {
    Write-Host "Processing language: $language" -ForegroundColor Cyan

    # Get child items and filter by __Updated field
    $childItems = Get-ChildItem -Path $rootItem.FullPath -Language $language -Recurse | Where-Object {
        $updatedField = $_.Fields['__Updated']
        if ($updatedField) {
            try {
                $parsedDate = [datetime]::ParseExact($updatedField, 'yyyyMMddTHHmmssZ', $null)
                return $parsedDate -gt $filterDate
            }
            catch {
                return $false
            }
        }
        return $false
    }

    # Process filtered items
    foreach ($item in $childItems) {
        $data += [PSCustomObject]@{
            Id                                 = $item.Id.Guid
            Name                               = $item.Name
            ParentPath                         = $item.Parent.Paths.FullPath
            Language                           = $language
            TemplateName                       = $item.TemplateName
            'Field.__Icon'                     = $item['__Icon']
            'Field.__Sortorder'                = $item['__Sortorder']
            'Field.__Display Name'             = $item['__Display Name']
            'Field.__Enable item fallback'     = $item['__Enable item fallback']
            'Field.__Enforce version presence' = $item['__Enforce version presence']
            'Field.__Unpublish'                = $item['__Unpublish']
            'Field.__Never Publish'            = $item['__Never Publish']
            'Field.__Help link'                = $item['__Help link']
            'Field.__Long description'         = $item['__Long description']
            'Field.__Short description'        = $item['__Short description']
            'Field.Title'                      = $item['Title']
            'Field.SubTitle'                   = $item['Sub Headline']
            'Field.Summary'                    = $item['Highlights']
            'Field.Description'                = $item['Main Content']
            'Field.Date'                       = $item['Date']
            'Field.Image'                      = $item['Detail Image']
            'Field.ThumbnailImage'             = $item['Listing Image']
            'Field.RelatedPractices'           = $item['Related Practices']
            'Field.RelatedIndustries'          = $item['Related Sectors']
            'Field.RelatedLocations'           = $item['Locations']
            'Field.RelatedPeople'              = $item['Professionals']
            'Field.RelatedInsights'            = $item['Related Articles']
            'Field.SxaTags'                    = $item['Popular Tags']
            'Field.MediaType'                  = $item['Types']
            'Field.MetaTitle'                  = $item['Browser Title']
            'Field.MetaDescription'            = $item['Meta Description']
            'Field.MetaKeywords'               = $item['Meta Keywords']
            'Field.OpenGraphTitle'             = $item['Browser Title']
            'Field.OpenGraphDescription'       = $item['Meta Description']
            'Field.OpenGraphImageUrl'          = $item['Meta Image']
            'Field.OpenGraphType'              = 'article'
        }
    }
}

# Check if data exists
if ($data.Count -gt 0) {
    # Convert to CSV with proper handling for commas
    $csvContent = $data | ConvertTo-Csv -NoTypeInformation

    # Write CSV to a memory stream
    $stream = New-Object System.IO.MemoryStream
    $writer = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::UTF8)
    $writer.Write($csvContent -join "`n")
    $writer.Flush()
    $stream.Position = 0  # Reset stream position

    # Provide CSV for download
    Out-Download -Name $tempFileName -InputObject $stream
    Write-Host 'Export completed. File ready for download: $tempFileName' -ForegroundColor Green
}
else {
    Write-Host 'No child items found to export.' -ForegroundColor Yellow
}
