<# 
Script 	: To fetch all the changeset from given date and related PBIs/ADOcards from mention date.

Output: In CSV format

changesetId	|author	        |createdDate	            |workItem_ID	    |workItemType	|workItem_Title
---------------------------------------------------------------------------------------------------------------------------------
2153	    |Praveen     	|2023-04-07T17:04:58.04Z	|13782	            |Bug	        |Incorrect field in Paymode-X output



Author 	: Praveen Kumar Sharma
version : 1.0
Date 	: 14/04/2023
#>

#Install-Module TfsCmdlets
Connect-TfsTeamProjectCollection '<TFS URL>'
Get-TfsTeamProject
Connect-TfsTeamProject "<Branch Name>" #D365FO

#region : Procees to create base64 Authorisation
[string]$user = <AzureDevOpsUsername> #"AzureDevOpsPATTom"
[string]$token = <Personal Access Token>  #"i6mesjn5d37k52xfx6vt54udarfjtswq"
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $user,$token)))
#endregion 

#region Variables to configure, URL, Project Collection, Date, OutputFile path
$projectCollection = <BranchName> #"D365FO"
$fromDate = (Get-Date).AddDays(-99).ToString("M/d/yyyy")
$baseUrl = '<TFS URL>' #"https://devops.visualstudio.com"
$ChangeSetbaseUrl = "$($baseUrl)/_apis/tfvc/changesets"
$changeSetHistoryUrl = "{0}?api-version=1.0&searchCriteria.itemPath=$/{1}&searchCriteria.fromDate={2}" -f $ChangeSetbaseUrl, $projectCollection, $fromDate
$FileDate = (Get-date).ToString("Mdyyyy_hhmmss")
$outputPath = "C:\temp\Changeset_$($FileDate).csv"
$finalTable = @()
#endregion

$param = @{
            uri =  $changeSetHistoryUrl
            Method = "Get"
            ContentType = "application/json"
            Header = @{Authorization=("Basic {0}" -f $base64AuthInfo)}
            }
Write-Host "Fetching all the changeSets! Pleasse Wait..." -ForegroundColor Yellow
$Header = Invoke-webRequest @param | Select-Object Content, Headers
$Content_1 = $Header.Content | ConvertFrom-Json
$Content_1.value  | %{ 
                    [dateTime]$date = $($($_.createdDate) -split "T")[0]
                    if($date -gt $fromDate)
                        {
                        $changeSetDetail = $_
                        $WIUrl = "{0}/_apis/tfvc/changesets/{1}/workItems" -f $($baseUrl), $($changeSetDetail.changesetId)
                        $WI = Invoke-WebRequest -Uri $WIUrl -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} | Select-Object Content 

                        $WIDetails = $WI.Content | ConvertFrom-Json
                        $WIDetails.Value | % {
                                $WItemp = $_
                                $finalTable += $changeSetDetail | Select-object changesetId, @{n="author"; e={$($changeSetDetail.author).displayName}}, createdDate, @{n="workItem_ID"; e={$($WItemp.ID)}}, @{n="workItemType"; e={$($WItemp.workItemType)}},@{n="workItem_Title"; e={$($WItemp.title)}} #| Export-Csv -Path $outputPath -Append -NoTypeInformation
                                }
                        }
                    }
$link = $($Header.Headers.Link) -replace "&lt;",""


#followUp Links
While($link -notlike ""  )
    {
    $Contents = Invoke-webRequest -uri $($link) -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} | Select-Object Content, Headers
    $link = ""
    $link = $($Contents.Headers.Link) -replace "&lt;",""
    $Content = $Contents.Content | ConvertFrom-Json
    $Content.value  | %{ 
                        [dateTime]$date = $($($_.createdDate) -split "T")[0]
                        if($date -gt $fromDate)
                            {
                            $changeSetDetail = $_
                            $WIUrl = "{0}/_apis/tfvc/changesets/{1}/workItems" -f $($baseUrl), $($changeSetDetail.changesetId)
                            $WI = Invoke-WebRequest -Uri $WIUrl -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} | Select-Object Content 
                            $WIDetails = $WI.Content | ConvertFrom-Json
                            $WIDetails.Value | % {
                                    $WItemp = $_
                                    $finalTable += $changeSetDetail | Select-object changesetId, @{n="author"; e={$($changeSetDetail.author).displayName}}, createdDate, @{n="workItem_ID"; e={$($WItemp.ID)}}, @{n="workItemType"; e={$($WItemp.workItemType)}},@{n="workItem_Title"; e={$($WItemp.title)}} #| Export-Csv -Path $outputPath -Append -NoTypeInformation
                                    }
                            }
                        }
    
    }
$finalTable | Export-Csv -Path $outputPath -Append -NoTypeInformation
Write-Host "All the changeSet is Exported to $($outputPath)" -ForegroundColor Green