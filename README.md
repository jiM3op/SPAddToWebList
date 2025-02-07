This is a small Console App, that we internally use for the following purposes:

- Create a Navigatiion Item in a List called "AllSites" whenever we create a new SharePoint Site Collection (as part of our Powershell Script)
- Run daily through Scheduler with Prameter "check" to check if a SiteCollection is no longer available and delete the corresponding navigation Item

  This is .NetFramework 4.72 - Using Microsoft.SharePoint.Client.
