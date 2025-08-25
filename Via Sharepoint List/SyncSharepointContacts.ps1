
$csvLogging = "C:\LOG_Sync_FuncNum_Delta.txt"

#Graph
$thumb = ""
$applicationID =""
$tenantID =""

#Sharepoint 
$contactListeSiteCollecionID = ""
$contactListeID = ""
$contactListSiteCollectionUrl = "url"
$contactListSharePoint = "Contact List"

$folderName = "Custom Sharepoint Contacts"
$e5groupId = ""

$targetType = "Group" 
$groupId = "" 




function WriteLog{    
    Param ([string]$logString)
    $dateTime = "[{0:dd/MM/yy} {0:HH:mm:ss}]" -f (Get-Date)
    # If File not exists use Add-Content to create it and add content
    if (-not (Test-Path -Path $csvLogging)) {Add-Content -Path $csvLogging -Value "Start CAL Contacts FuncNum Logging"}
    Add-content $csvLogging -value "$datetime $logString"
}

function CheckE5{
    Param ([string]$userIdForCheck_email)
    $searchQuery = "userPrincipalName:$userIdForCheck_email"
    $checkE5 = Get-MgGroupMember -GroupId $e5groupId -search $searchQuery -ConsistencyLevel eventual -ErrorAction Stop -all
    if($checkE5){
        return $true
    } else {
        return $false
    }   
}

$benchmark = [System.Diagnostics.Stopwatch]::StartNew()

WriteLog('-------------------------------------------')
WriteLog("[INIT] Script start run, group or user: $groupId $TestUserId ")
WriteLog('-------------------------------------------')

try{
    Connect-MgGraph -TenantId $tenantID -ClientID $applicationID -CertificateThumbprint $thumb -ErrorAction Stop

    [array]$Data = Get-MgSiteListItem -SiteId $contactListeSiteCollecionID -ListId $contactListeID -ExpandProperty "Fields" -ErrorAction Stop -all

    $contacts = [System.Collections.Generic.List[Object]]::new()
    foreach ($item in $data) {
       $fields = $item.Fields.AdditionalProperties
       if ($fields.Status -ne $null) {
       $contact = [PSCustomObject]@{
           id             = $item.Id
           givenName      = $fields.field_4
           surname        = $fields.field_3
           businessPhones = $fields.Title
           officeLocation = $fields.field_5
           status         = if ($fields.Status -eq $null) { "" } else { $fields.Status }
       }
       $contacts.Add($contact)
        }
    }
    

 if ($contacts.count -eq 0) {
        WriteLog('[INFO] Problem with list or item retrieval. Either list is empty or items could not be retrieved')
        
    }
   
    

    if ($targetType -eq "Group") {
        $members = Get-MgGroupMember -GroupId $groupId -ConsistencyLevel eventual -ErrorAction Stop -all
        if ($members.count -eq 0) {
            WriteLog('[WARNING] Problem with retrieving group or group is empty')
            return
        }
    } elseif ($targetType -eq "User") {
            $members = $TestUserId
        if ($members.count -eq 0) {
        WriteLog('[WARNING] Problem with retrieving user or user is empty')
        return
         }
        }
        

    
} catch {
    WriteLog("[ERROR] Error/Exception in List and Member retrieval: $($_.Exception.Message)")
    WriteLog("[ERROR] $($_.Exception.StackTrace)")
}

if ($contacts.count -ne 0) {

        foreach($member in $members) {
            try {
                $userId = if ($targetType -eq "User") { $TestUserId } else { $member.Id }
                $memberDets = Get-MgUser -UserId $userId
                $memberId = $memberDets.UserPrincipalName

                $existingFolderSkip = Get-MgUserContactFolder -UserId $memberId | Where-Object { $_.DisplayName -eq $folderName }
            } catch {
                WriteLog("[ERROR] Error retrieving member and folder: $($_.Exception.Message)")
            }

            $hasE5 = CheckE5($memberId)

            if($hasE5){
                if($existingFolderSkip){
                    $folderId = $existingFolderSkip.Id
                    try{
                        $allUserContacts = @()
                        $cs = Get-MgUserContactFolderContact -userid $memberId -ContactFolderId $folderId -All -ErrorAction Stop 
                        $allUserContacts += $cs
                    } catch {
                        WriteLog("[ERROR] Couldnt retrieve contacts for $($memberId): $($_.Exception.Message)")
                    }
            
                    foreach($contact in $contacts){
                        $params = @{
                            givenName = $contact.givenName
                            surname = $contact.surname
                            businessPhones = @($contact.businessPhones)
                            officeLocation = $contact.OfficeLocation
                            categories = @($folderName)
                            personalNotes = $contact.id
                        }
            
                        try{
                            $existingContact = $allUserContacts | Where-Object{ $_.PersonalNotes -eq $contact.id  }

                            if($existingContact){
                                $exists = $true
                            } else {
                                $exists = $false
                            }                                                 
                        } catch {
                            WriteLog("[ERROR] Error with filter method: $($_.Exception.Message)")
                        }
                        
                        if($exists){
                            if($contact.Status -eq "Update") {
                                try{
                                    $updateid = [String]$existingContact.Id
                                                                
                                    Update-MgUserContactFolderContact -Userid $memberId -ContactFolderId $folderId -ContactId $updateid -ErrorAction Stop  -BodyParameter $params
                                    $updateSPItem = $true
                                    #WriteLog("[INFO] Update contact: $($existingContact.givenName) $($existingContact.surname) for $($memberId)")
                                } catch {
                                    WriteLog("[ERROR] Couldnt update contact: $($contact.givenName) $($contact.surname) for $($memberId): $($_.Exception.Message)")
                                }
                            } elseif($contact.Status -eq "Delete") {
                                try{
                                    $deleteTag = "$($existingContact.givenName) $($existingContact.surname)"
                                    $deleteId = [String]$existingContact.Id
                                    Remove-MgUserContactFolderContact -UserId $memberId -ContactFolderId $folderId -ContactId $deleteId -ErrorAction Stop
                                    $updateSPItem = $true
                                    #WriteLog("[INFO] Deleted contact: $deleteTag for $($memberId)")
                                } catch {
                                    WriteLog("[ERROR] Couldnt delete contact: $deleteTag for $($memberId): $($_.Exception.Message)")
                                }
                            }
                        } else {
                            if($contact.Status -eq "New"){
                                try{
                                    New-MgUserContactFolderContact -UserId $memberId -ContactFolderId $folderId -BodyParameter $params
                                    $updateSPItem = $true
                                    #WriteLog("[INFO] Created new contact: $($contact.givenName) $($contact.surname) for $($memberId)")
                                } catch {
                                    WriteLog("[ERROR] Couldnt create contact: $($contact.givenName) $($contact.surname) for $($memberId): $($_.Exception.Message)")
                                }
                            }
                        }

                        
                    }
                    WriteLog("[DONE] Func Numbers synced for $($memberDets.DisplayName), $($memberId)") 
                } else {
                    WriteLog("[DONE] Skipped $($memberDets.DisplayName), $($memberId)") 
                }
            } else {
                WriteLog("[DONE] Skipped $($memberDets.DisplayName), $($memberId), is not part of the E5 group.")
                continue
            }
        }


            foreach ($contact in $contacts) {
                try {
                    if ($contact.Status -eq "Delete") {
                        Remove-MgSiteListItem -SiteId $contactListeSiteCollecionID -ListId $contactListeID -ListItemId $contact.id
                        WriteLog("[INFO] SharePoint Item Deleted: $($contact.id) for $($contact.surname) $($contact.givenName)")
                    } elseif ($contact.Status -eq "New" -or $contact.Status -eq "Update") {
                        $siteURL = Get-MgSite -SiteId $contactListeSiteCollecionID

                        $body = @{
                            fields = @{
                                Status = ""
                            }
                        } | ConvertTo-Json -Depth 3

                        Invoke-MgGraphRequest -Method PATCH `
                            -Uri "https://graph.microsoft.com/v1.0/sites/$($siteURL.Id)/lists/$($contactListeID)/items/$($contact.id)" `
                            -Body $body

                        WriteLog("[INFO] SharePoint Item Updated: $($contact.id)  $($contact.surname) $($contact.givenName)")
                    }
                } catch {
                    WriteLog("[ERROR] Couldn't update SharePoint Item: $($contact.id)  $($contact.surname) $($contact.givenName)")
                }
            }


} else {
    WriteLog('[DONE] Skipped ALL, there was nothing to do ')
    
}
Disconnect-MgGraph

$benchmark.Stop()
WriteLog("[Benchmark] $($benchmark.ElapsedMilliseconds) ms")
WriteLog('-------------------------------------------')
