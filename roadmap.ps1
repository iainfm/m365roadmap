# Based on the downloader/creator in ff-efd53b.js, reworked for PowerShell

$apiEndpoint = "https://www.microsoft.com/msonecloudapi/roadmap/features/"

try {

    $roadmapData = Invoke-RestMethod -Uri $apiEndpoint # Comment this line

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
    try {
        $tagsProduct = $tagsProduct.TrimEnd(", ")
    } catch { }

    $tagsReleasePhase = ""
    foreach ($tag in $r.tagsContainer.releasePhase) {
        $tagsReleasePhase += $tag.tagName + ", "
    }
    try {
        $tagsReleasePhase = $tagsRelease.TrimEnd(", ")
    } catch { }

    $tagsPlatform = ""
    foreach ($tag in $r.tagsContainer.platforms) {
        $tagsPlatform += $tag.tagName + ", "
    }
    try {
        $tagsPlatform = $releasePlatforms.TrimEnd(", ")
    } catch { }

    $tagsCloudInstance = ""
    foreach ($tag in $r.tagsContainer.cloudInstances) {
        $tagsCloudInstance += $tag.tagName + ", "
    }
    try {
        $tagsCloudInstance = $tagsCloudInstance.TrimEnd(", ")
    } catch { }

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

$csvData | Export-Csv -Path .\roadmap.csv -NoTypeInformation -Encoding utf8BOM

# ToDo: Fix this so that special characters are properly encoded
$csvData | ConvertTo-Csv -NoTypeInformation |  select-string "security measures. We"