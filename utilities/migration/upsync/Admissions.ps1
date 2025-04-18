# Parameters
$mediaItemPath = '/sitecore/media library/Migration/AdmissionsData'  # Path to the JSON file in the Media Library
$templateName = '/sitecore/templates/Foundation/BTLaw/BaseContent/_enum'  # Template to use for item creation

Write-Host "Starting the script..." -ForegroundColor Cyan

# Load the Media Library item
Write-Host "Attempting to load media item from path: $mediaItemPath" -ForegroundColor Cyan
$mediaItem = Get-Item -Path $mediaItemPath

if (-Not $mediaItem) {
    Write-Host "Error: Media item not found at path: $mediaItemPath" -ForegroundColor Red
    return
}
else {
    Write-Host "Successfully loaded media item." -ForegroundColor Green
}

# Get the binary stream of the media item
Write-Host "Reading binary stream from the media item..." -ForegroundColor Cyan
[System.IO.Stream]$body = $mediaItem.Fields['Blob'].GetBlobStream()
try {
    $contents = New-Object byte[] $body.Length
    $body.Read($contents, 0, $body.Length) | Out-Null
}
finally {
    $body.Close()
    Write-Host "Binary stream read successfully." -ForegroundColor Green
}

# Convert CSV content to PowerShell objects
Write-Host "Converting binary content to CSV..." -ForegroundColor Cyan
$csvContent = [System.Text.Encoding]::UTF8.GetString($contents)
$data = $csvContent | ConvertFrom-Csv

# Ensure data exists
if ($null -eq $data -or $data.Count -eq 0) {
    Write-Host "Warning: No data found in the CSV file." -ForegroundColor Yellow
    return
}
else {
    Write-Host "Successfully loaded and parsed CSV data. Processing records..." -ForegroundColor Green
}

# Process each row and create/update items
foreach ($record in $data) {
    $language = $record.Language
    $itemName = $record.Name -replace '-', ' '
    $itemId = $record.Id
    $parentPath = $record.ParentPath -replace '/sitecore/content/BTLaw/Global/Data/Admissions', '/sitecore/content/BTLaw/BTLaw/Global Data/Admissions'
    $itemPath = "$parentPath/$itemName"

    Write-Host "Processing record: $itemName (ID: $itemId, Language: $language)" -ForegroundColor Cyan

    # Check if item already exists
    if (Test-Path -Path $itemPath) {
        Write-Host "Item exists at path: $itemPath. Checking language version..." -ForegroundColor Yellow
        $defaultItem = Get-Item -Path "$parentPath/$itemName"
        $item = Get-Item -Path "$parentPath/$itemName" -Language $language
        if ($null -eq $item) {
            Write-Host "Language version not found. Creating new language version for: $itemName" -ForegroundColor Cyan
            $item = Add-ItemVersion -Item $defaultItem -TargetLanguage $language -DoNotCopyFields
        }
        else {
            Write-Host "Language version already exists for: $itemName" -ForegroundColor Green
        }
    }
    else {
        Write-Host "Item does not exist. Creating new item: $itemName at path: $parentPath" -ForegroundColor Cyan
        $item = New-Item -Path $parentPath -Name $itemName -ItemType $templateName -ForceId $itemId -ErrorAction Stop
    }

    # Set field values
    if ($null -ne $item) {
        Write-Host "Updating fields for item: $itemName" -ForegroundColor Cyan
        $item.Editing.BeginEdit()
        try {
            foreach ($field in $record.PSObject.Properties) {
                $fieldName = $field.Name -replace 'Field.', ''
                $fieldValue = $field.Value

                # Ensure the field name is valid and not null
                if (-not [string]::IsNullOrWhiteSpace($fieldName)) {
                    $item[$fieldName] = $fieldValue
                }
            }
            $item.Editing.EndEdit()
            Write-Host "Successfully updated fields for item: $itemName (ID: $itemId)" -ForegroundColor Green
        }
        catch {
            $item.Editing.CancelEdit()
            Write-Host "Error: Failed to update fields for item: $itemName (ID: $itemId)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Error: Failed to create or locate item: $itemName" -ForegroundColor Red
    }
}

Write-Host "Script execution completed." -ForegroundColor Cyan
