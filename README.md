# Excel-Optimizer



This is my attempt at solving some headaches that come with modernizing Excel ahead of cloud migrations.

Untested at scale, use at your own risk.

There are three folders.

One contains the link finder script. Run this to find connected workbooks and worksheets.

One contains an issue finder. This checks for common issues.

The third is still a work in progress, but will work to resolve the issues in the second folder. Ideally, this will make its way into a small desktop app or CLI program 


Installing PreReqs


Install-Package Microsoft.Office.Interop.Excel

Register-PackageSource -Name NuGet -Location https://api.nuget.org/v3/index.json -ProviderName NuGet

