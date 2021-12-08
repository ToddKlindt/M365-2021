Get-Command -Noun *sitescript*,*sitedesign*,*sitetemplate* -Module pnp.powershell,Microsoft.Online.SharePoint.PowerShell

Get-PnPSiteScriptFromWeb -IncludeAll -Lists "Lists/Customer Tracking" | Tee-Object -Variable SiteScriptFromWeb

Get-PnPSiteScriptFromList -Url "https://m365x995492.sharepoint.com/sites/M365Test1/Lists/Customer%20Tracking" | Tee-Object -Variable NewSiteScript

$SiteScript = Get-Content 'C:\Users\todd_\Todd Klindt Consulting LLC\Conference Presentations - Topics\Site Scripts and Site Designs\Site Script Demo.json' -Raw
Add-PnPSiteScript -Title "Site Script Demo" -Description "Site Script Demo" -Content $SiteScript
Add-PnPSiteDesign -Title "Site Script Demo" -SiteScriptIds 1b0d9d60-8683-49cf-969a-37aee7af9829 -Description "1b0d9d60-8683-49cf-969a-37aee7af9829" -WebTemplate 64
