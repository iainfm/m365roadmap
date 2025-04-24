# m365roadmap
Meddling with the Microsoft 365 Roadmap

## Update April 2025
Microsoft has changed their API endpoint, as published at https://www.microsoft.com/en-gb/microsoft-365/roadmap.

This has made accessing the M365 Roadmap data much easier. The old v1 API lives on at https://www.microsoft.com/releasecommunications/api/v1/m365 while the new v2 API is at https://roadmap-api.azurewebsites.net/api/features.

The new API appears to take a responseFormat parameter. Valid options include json and csv.
