PARAM ($PastDays = 7)
#************************************************
# Office365UsageReports.ps1
# Version 1.0
# Date: 5-23-2018
# Author: Tim Springston [MSFT]
# Description: This script will use a previously created AAD application context with sufficient authorization to
# connect and pull all Office 365 reports. Results are placed into CSV files for review, one per report, all in the 
# same directory.
#************************************************
cls
#Reference links.
#Prereqs to using this script https://docs.microsoft.com/en-us/azure/active-directory/active-directory-reporting-api-prerequisites-azure-portal
#Office 365 Reporting web service https://msdn.microsoft.com/en-us/library/office/jj984325.aspx
#Mailbox Usage Reports https://msdn.microsoft.com/en-us/library/office/jj984325.aspx 

#Reports to collect.
$O365Reports = @(
    'getEmailActivityUserDetail';
    'getEmailActivityCounts';
    'getEmailActivityUserCounts';
    'getEmailAppUsageUserDetail';
    'getEmailAppUsageAppsUserCounts';
    'getEmailAppUsageUserCounts';
    'getEmailAppUsageVersionsUserCounts';
    'getMailboxUsageDetail';
    'getMailboxUsageMailboxCounts';
    'getMailboxUsageQuotaStatusMailboxCounts';
    'getMailboxUsageStorage';
    'getOffice365ActivationsUserDetail';
    'getOffice365ActivationCounts';
    'getOffice365ActivationsUserCounts';
    'getOffice365ActiveUserDetail';
    'getOffice365ActiveUserCounts';
    'getOffice365ServicesUserCounts';
    'getOffice365GroupsActivityDetail';
    'getOffice365GroupsActivityCounts';
    'getOffice365GroupsActivityGroupCounts';
    'getOffice365GroupsActivityStorage';
    'getOffice365GroupsActivityFileCounts';
    'getOneDriveActivityUserDetail';
    'getOneDriveActivityUserCounts';
    'getOneDriveActivityFileCounts';
    'getOneDriveUsageAccountDetail';
    'getOneDriveUsageAccountCounts';
    'getOneDriveUsageFileCounts';
    'getOneDriveUsageStorage';
    'getSharePointActivityUserDetail';
    'getSharePointActivityFileCounts';
    'getSharePointActivityUserCounts';
    'getSharePointActivityPages';
    'getSharePointSiteUsageDetail';
    'getSharePointSiteUsageFileCounts';
    'getSharePointSiteUsageSiteCounts';
    'getSharePointSiteUsageStorage';
    'getSharePointSiteUsagePages';
    'getSkypeForBusinessActivityUserDetail';
    'getSkypeForBusinessActivityCounts';
    'getSkypeForBusinessActivityUserCounts';
    'getSkypeForBusinessDeviceUsageUserDetail';
    'getSkypeForBusinessDeviceUsageDistributionUserCounts';
    'getSkypeForBusinessDeviceUsageUserCounts';
    'getSkypeForBusinessOrganizerActivityCounts';
    'getSkypeForBusinessOrganizerActivityUserCounts';
    'getSkypeForBusinessOrganizerActivityMinuteCounts';
    'getSkypeForBusinessParticipantActivityCounts';
    'getSkypeForBusinessParticipantActivityUserCounts';
    'getSkypeForBusinessParticipantActivityMinuteCounts';
    'getSkypeForBusinessPeerToPeerActivityCounts';
    'getSkypeForBusinessPeerToPeerActivityUserCounts';
    'getSkypeForBusinessPeerToPeerActivityMinuteCounts';
    'getteamsDeviceUsageUserDetail';
    'getteamsDeviceUsageUserCounts';
    'getteamsDeviceUsagedistributionUserCounts';
    'getteamsUserActivityUserDetail';
    'getteamsUserActivityCounts';
    'getteamsUserActivityUserCounts';
    'getYammerActivityUserDetail';
    'getYammerActivityCounts';
    'getYammerActivityUserCounts';
    'getYammerDeviceUsageUserDetail';
    'getYammerDeviceUsageDistributionUserCounts';
    'getYammerDeviceUsageUserCounts';
    'getYammerGroupsActivityDetail';
    'getYammerGroupsActivityGroupCounts';
    'getYammerGroupsActivityCounts'
)

#Application registration context under which the report will be collected.
$ClientID       = "GUID"             #Should be a ~35 character string
$ClientSecret   = "secret"           #Should be a ~44 character string 
$loginURL       = "https://login.windows.net"
$tenantdomain   = "domainname"            # For example, contoso.onmicrosoft.com

function GetO365Report      ($url, $reportname, $tenantname, $days) {
Write-Host "Collecting Office 365 report "  $reportname "..."
# Get an Oauth 2 access token based on client id, secret and tenant domain
#Request Graph API Token and build request header.
$body       = @{grant_type="client_credentials";client_id=$ClientID;client_secret=$ClientSecret}
$BearerToken      = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body
$headerParams = @{'Authorization'="$($oauth.token_type)", "$($oauth.access_token)"}
$RESTEndpoint = "https://graph.microsoft.com" #Graph REST endpoint

#Get Bearer token for REST endpoint
$body       = @{grant_type="client_credentials";resource=$RESTEndpoint;client_id=$ClientID;client_secret=$ClientSecret}
$oauth      = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body

#Get access token for REST endpoint
$body       = @{grant_type="client_credentials";resource=$RESTEndpoint;client_id=$ClientID;client_secret=$ClientSecret}
$headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}

$url = $RESTEndpoint + "/v1.0/reports/" + $reportname + "(period='d$days')" 
$AuditOutputCSV = $Pwd.Path + "\" + (($tenantdomain.Split('.')[0]) + "_" + $reportname + "_AuditReport.csv")
try {
    $myReport = (Invoke-WebRequest -Headers $headerParams -Uri $url -UseBasicParsing -outfile $AuditOutputCSV ) 
    #Write-host "Report $reportname can be found at $AuditOutputCSV."
    }
    catch {[Net.WebException] $_.Exception.ToString() | FL}
}


foreach ($Report in $O365Reports)		
    {
    GetO365Report $Report $Report $Tenantdomain $pastdays
    }
$Path = $PWd.path + "\"
Write-host "All reports are located at $Path."
