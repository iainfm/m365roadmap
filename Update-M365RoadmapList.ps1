# Powershell script to download the Microsoft 365 Roadmap items and populate/update a Sharepoint list with them.
# Status is reported to a Teams channel.
# Downloader is no longer based on the downloader/creator in the ff-efd53b.js file used by Microsoft's web page.
# iainfm
# Jan 2025

# Updated April 2025 due to Microsoft providing a better way to directly download the CSV/JSON information from the website.

# Requirements: Access to the certificate thumbprint, client ID, and tenant ID for the SharePoint site. PnP.PowerShell module installed

# Variable definitions

# Sharepoint details
$tenantId = '<sharepoint tenant id>'
$thumbPrint = '<auth cert thumbprint>'
$clientId = '<app registration client id>'
$siteUrl = "<sharepoint site URL containing list to update>"
$ListName = '<sharepoint list name'

# Teams webhook URL and headers
$webhookUrl = '<webhook URL for Teams channel posting>'
$webhookHeaders = @{ 'Content-Type' = 'application/json' }

# Microsoft's API endpoint for the roadmap data
$apiEndpoint = "https://www.microsoft.com/releasecommunications/api/v2/m365?responseFormat=json"

# Logs
$transcriptPath = 'C:\Logs\M365Roadmap'
$transcriptName = "Transcript_$(get-date -format yyyy-MM-dd-HHmm).txt"

# Functions
function Send-TeamsMessage {
    Param ( [string]$Message )

    # Adaptive card payload
    $adaptiveCardPayload = @{
        type = 'message'
        attachments = @(
            @{
                contentType = 'application/vnd.microsoft.card.adaptive'
                content = @{
                    "$schema" = 'http://adaptivecards.io/schemas/adaptive-card.json'
                    type = 'AdaptiveCard'
                    version = '1.4'
                    msteams = @{
                        width = "Full"
                    }
                    body = @(
                        @{
                            type = 'TextBlock'
                            text = "$Message"
                            size = 'Medium'
                            weight = 'Default'
                            wrap = $true
                        }
                    )
                }
            }
        )
    } | ConvertTo-Json -Depth 10

    # Post the message
    try {
        Invoke-RestMethod -Uri $webhookUrl -Method Post -Headers $webhookHeaders -Body $adaptiveCardPayload | Out-Null
    }
    catch {
        Write-Warning "Posting to Teams failed. Error $_"
    }
}

# Start transcript logging
Start-Transcript -Path (Join-Path -Path $transcriptPath -ChildPath $transcriptName)

# Step 1 - download and build the data into a usable format

try {

    # Download the roadmap data from the website
    $roadmapData = (Invoke-RestMethod -Uri $apiEndpoint).Value
} catch {

    Write-Host "Unable to get Roadmap data from DOM: $_"
    Exit 1

}

# Check we got data back
if ($roadmapData.Count -eq 0) {
    Write-Host "No data received from the API. Exiting."
    Stop-Transcript
    Exit 1
}

# Build the roadmap data, based on what the website provides as a CSV file
$jsonData = @()

foreach ($r in $roadmapData) {

    $featureId = $r.id
    $description = $r.title
    $details = $r.description
    $status = $r.status
    $moreInfo = $r.moreInfoUrls[0] # Only take the first one in the array in case there are many.

    $tagsProduct = ""
    foreach ($tag in $r.products) {
        $tagsProduct += $tag + ", "
    }
    try {
        $tagsProduct = $tagsProduct.TrimEnd(", ")
    } catch { }

    $tagsReleasePhase = ""
    foreach ($tag in $r.releaseRings) {
        $tagsReleasePhase += $tag + ", "
    }
    try {
        $tagsReleasePhase = $tagsReleasePhase.TrimEnd(", ")
    } catch { }

    $tagsPlatform = ""
    foreach ($tag in $r.platforms) {
        $tagsPlatform += $tag + ", "
    }
    try {
        $tagsPlatform = $tagsPlatform.TrimEnd(", ")
    } catch { }

    $tagsCloudInstance = ""
    foreach ($tag in $r.cloudInstances) {
        $tagsCloudInstance += $tag + ", "
    }
    try {
        $tagsCloudInstance = $tagsCloudInstance.TrimEnd(", ")
    } catch { }

    $addedToRoadmap = $r.created
    $lastModified = $r.modified
    $preview = $r.previewAvailabilityDate
    $release = $r.generalAvailabilityDate

    # This builds the data into a table using the same column names as the CSV file that Microsoft provides when the data is downloaded manually
    $jsonData += [PSCustomObject]@{
        "Feature ID" = $featureId
        Description = $description
        Details = $details
        Status = $status
        "More Info" = $moreInfo
        "Tags - Product" = $tagsProduct
        "Tags - Release phase" = $tagsReleasePhase
        "Tags - Platform" = $tagsPlatform
        "Tags - Cloud instance" = $tagsCloudInstance
        "Added to Roadmap" = $addedToRoadmap
        "Last Modified" = $lastModified
        Preview = $preview
        Release = $release
    }

}

# We should now have the data for importing into the Sharepoint list

# Step 2 - Update the Sharepoint list using the data received

# Connect to Sharepoint
Import-Module Pnp.Powershell

try {
    Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenantId -Thumbprint $thumbPrint
}
catch {
    Write-Error "Could not connect to Sharepoint: $_"
    Stop-Transcript
    Exit 1
}

# The field mapping is as follows:

# JSON                    ->   Sharepoint
# "Feature ID"            ->   field_0
# Description             ->   Title
# Details                 ->   field_2
# Status                  ->   field_3
# More Info               ->   field_4
# Products                ->   field_5
# ReleasePhase            ->   field_6
# Platforms               ->   field_7
# CloudInstances          ->   field_8
# "Added to Roadmap"      ->   field_9
# "Last Modified"         ->   field_10
# Preview                 ->   field_11
# Release                 ->   field_12

# Download all the existing list items
$ListItems = Get-PnPListItem -List $ListName -Fields ID, "field_0", "Title", "field_2", "field_3", "field_4", "field_5", "field_6", "field_7", "field_8", "field_9", "field_10", "field_11", "field_12"

# Create an array to store and process the list item field values
$ListItemsArray = @()

# Loop through each item in the list and add it to the array
foreach ($Item in $ListItems) {
    $ListItemsArray += $Item.FieldValues
}

# Loop through each record in the JSON data, compare it with the List data. If there's a difference, update or add the list item from the JSON
$progressCounter = 0
$itemsUpdated = 0
$itemsAdded = 0
$itemsFailed = 0

foreach ($record in $jsonData) {

    # Show progress
    Write-Progress -Activity "Updating Sharepoint list" -Status "Processing record $progressCounter of $($jsonData.count)" -PercentComplete (100 * $progressCounter / $jsonData.count)
    $progressCounter++

    # Find the List record with the matching Feature ID
    $listItem = $ListItemsArray.Where{$_.field_0 -eq $record."Feature ID"}

    # Initialist an empty hash table for the updates and a string to store what fields will be updated
    $updates = @{}
    $toUpdate = ""

    # Get the Sharepoint ID of the list item (this is the List's primary key, not the Feature ID)
    $id = $listItem.ID

    if ($listItem) {

        # Compare the data of each field between the JSON and the Sharepoint list

        # 'Description' / 'Title' field
        if (($record.Description -ne $listItem.Title) -and (($record.Description -ne "" ) -and ($null -ne $listItem.Title ))) {
            $updates += @{'Title' = $record.Description}
            $toUpdate += "Title, "
        }

        # 'Details' field
        if (($record.Details -ne $listItem.field_2) -and (($record.Details -ne "" ) -and ($null -ne $listItem.field_2 ))) {
            $updates += @{'field_2' = $record.Details}
            $toUpdate += "Details, "
        }

        # 'Status' field
        if (($record.Status -ne $listItem.field_3) -and (($record.Status -ne "" ) -and ($null -ne $listItem.field_3 ))) {
            $updates += @{'field_3' = $record.Status}
            $toUpdate += "Status, "
        }

        # 'More Info' field
        if (($record.'More Info' -ne $listItem.field_4) -and (($record.'More Info' -ne "" ) -and ($null -ne $listItem.field_4 ))) {
            $updates += @{'field_4' = $record.'More Info'}
            $toUpdate += "More Info, "
        }

        # 'Tags - Product field'
        if (($record.'Tags - Product' -ne $listItem.field_5) -and (($record.'Tags - Product' -ne "" ) -and ($null -ne $listItem.field_5 ))) {
            $updates += @{'field_5' = $record.'Tags - Product'}
            $toUpdate += "Tags - Product, "
        }

        # 'Tags - Release phase field'
        if (($record.'Tags - Release phase' -ne $listItem.field_6) -and (($record.'Tags - Release phase' -ne "" ) -and ($null -ne $listItem.field_6 ))) {
            $updates += @{'field_6' = $record.'Tags - Release phase'}
            $toUpdate += "Tags - Release phase, "
        }

        # 'Tags - platform field'
        if (($record.'Tags - platform' -ne $listItem.field_7) -and (($record.'Tags - platform' -ne "" ) -and ($null -ne $listItem.field_7 ))) {
            $updates += @{'field_7' = $record.'Tags - platform'}
            $toUpdate += "Tags - platform, "
        }

        # 'Tags - Cloud instance field'
        if (($record.'Tags - Cloud instance' -ne $listItem.field_8) -and (($record.'Tags - Cloud instance' -ne "" ) -and ($null -ne $listItem.field_8 ))) {
            $updates += @{'field_8' = $record.'Tags - Cloud instance'}
            $toUpdate += "Tags - Cloud instance, "
        }

        # 'Added to Roadmap field' - this is a date field. For some reason Sharepoint often thinks they are an hour apart, so only update if
        # the difference is greater than 0 days
        if (($record.'Added to Roadmap' -ne "" ) -and ($null -ne $listItem.field_9 )) {
            if ((New-TimeSpan -start $record.'Added to Roadmap' -end $listItem.field_9).days -gt 0) {
                $updates += @{'field_9' = $record.'Added to Roadmap'}
                $toUpdate += "Added to Roadmap, "
            }
        }

        # 'Last Modified field' - this is another date field
        if (($record.'Last Modified' -ne "" ) -and ($null -ne $listItem.field_10 )) {
            if ((New-TimeSpan -start $record.'Last Modified' -end $listItem.field_10).days -gt 0) {
                $updates += @{'field_10' = $record.'Last Modified'}
                $toUpdate += "Last Modified, "
            }
        }

        # 'Preview' field
        if (($record.Preview -ne $listItem.field_11) -and (($record.Preview -ne "" ) -and ($null -ne $listItem.field_11 ))) {
            $updates += @{'field_11' = $record.Preview}
            $toUpdate += "Preview, "
        }

        # 'Release' field
        if (($record.Release -ne $listItem.field_12) -and (($record.Release -ne "" ) -and ($null -ne $listItem.field_12 ))) {
            $updates += @{'field_12' = $record.Release}
            $toUpdate += "Release"
        }

        try {
            $toUpdate = $toUpdate.TrimEnd(", ")
        } catch { }

        # If there are updates, update the list item
        if ($updates.count -ne 0) {
            # Update the list
            Write-Output "Updating Feature ID $($record.'Feature ID') with $toUpdate"
            try {
                Set-PnPListItem -List $ListName -Identity $id -Values $updates | Out-Null
                $itemsUpdated++
            }
            catch {
                Write-Error $_
                $itemsFailed++
            }
        }
    }
    else {
        # This is a new record - add the item to the Sharepoint list
        Write-Warning "Adding new item - Feature ID $($record.'Feature ID')"
        $updates += @{'Title' = $record.Description}
        $updates += @{'field_0' = $record.'Feature ID'}
        $updates += @{'field_2' = $record.Details}
        $updates += @{'field_3' = $record.Status}
        $updates += @{'field_4' = $record.'More Info'}
        $updates += @{'field_5' = $record.'Tags - Product'}
        $updates += @{'field_6' = $record.'Tags - Release phase'}
        $updates += @{'field_7' = $record.'Tags - platform'}
        $updates += @{'field_8' = $record.'Tags - Cloud instance'}
        $updates += @{'field_9' = (Get-Date($record.'Added to Roadmap') -Format "MM/dd/yyyy HH:mm:ss")}
        $updates += @{'field_10' = (Get-Date($record.'Last Modified') -Format "MM/dd/yyyy HH:mm:ss")}
        $updates += @{'field_11' = $record.Preview}
        $updates += @{'field_12' = $record.Release}

        try {
            Add-PnPListItem -List $ListName -Values $updates | Out-Null
            $itemsAdded++
        }
        catch {
            Write-Error $_
            $itemsFailed++
        }
    }
}

Write-Output ''
Write-Output "Items updated: $itemsUpdated"
Write-Output "Items added: $itemsAdded"
Write-Output "Items failed: $itemsFailed"

Send-TeamsMessage -Message "Sharepoint list has been updated.

Items updated: $itemsUpdated

Items added: $itemsAdded

Items failed: $itemsFailed"

Stop-Transcript
