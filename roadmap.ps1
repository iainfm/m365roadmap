# Based on the downloader/creator in ff-efd53b.js, reworked for PowerShell

$apiEndpoint = "https://www.microsoft.com/msonecloudapi/roadmap/features/"
# $apiEndpoint = "/en-us/microsoft-365/roadmap/assets/test/mocks/features.json"

try {
    $response = Invoke-WebRequest -Uri $apiEndpoint # Comment this line
    $content = $response.content.replace('\n', ' ') # and this one for testing
    # $content = Get-Content .\data.json            # Uncomment this line for speedier testing purposes
    $roadmapData = $content | ConvertFrom-Json -AsHashtable


} catch {
    Write-Host "Unable to get Roadmap data from DOM: $_"
}

# Build the roadmap data, based on what the website provides as a CSV file
$csvData = @()

foreach ($r in $roadmapData) {

    $featureId = $r.id
    $description = $r.title
    $details = $r.description
    $status = $r.status
    $moreInfo = $r.moreInfoLink

    $tagsProduct = ""
    foreach ($tag in $r.tagsContainer.products) {
        $tagsProduct += $tag.tagName + ", "
    }
    $tagsProduct = $tagsProduct.TrimEnd(", ")

    $tagsReleasePhase = ""
    foreach ($tag in $r.tagsContainer.releasePhase) {
        $tagsReleasePhase += $tag.tagName + ", "
    }
    $tagsReleasePhase = $tagsRelease.TrimEnd(", ")

    $tagsPlatform = ""
    foreach ($tag in $r.tagsContainer.platforms) {
        $tagsPlatform += $tag.tagName + ", "
    }
    $tagsPlatform = $releasePlatforms.TrimEnd(", ")

    $tagsCloudInstance = ""
    foreach ($tag in $r.tagsContainer.cloudInstances) {
        $tagsCloudInstance += $tag.tagName + ", "
    }
    $tagsCloudInstance = $tagsCloudInstance.TrimEnd(", ")

    $addedToRoadmap = $r.created
    $lastModified = $r.modified
    $preview = $r.publicPreviewDate
    $release = $r.publicDisclosureAvailabilityDate

    $csvData += [PSCustomObject]@{
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

$csvData | Export-Csv -Path .\roadmap.csv -NoTypeInformation