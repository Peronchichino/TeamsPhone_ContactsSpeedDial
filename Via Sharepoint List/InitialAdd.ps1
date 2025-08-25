
$csvLogging = "C:\LOG_InitialContactAdd_FuncNum.txt"

#Graph
$thumb = ""
$applicationID =""
$tenantID =""

#Sharepoint 
$contactListeSiteCollecionID = ""
$contactListeID = ""
$contactListSiteCollectionUrl = "url"
$contactListSharePoint = "Contact List"

#KontaktFolderName und E5Gruppe
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
WriteLog("[INIT] Script start run, group or user: $groupId $userId")
WriteLog('-------------------------------------------')

try {
    Connect-MgGraph -TenantId $tenantID -ClientID $applicationID -CertificateThumbprint $thumb -ErrorAction Stop

    [array]$Data = Get-MgSiteListItem -SiteId $contactListeSiteCollecionID -ListId $contactListeID -ExpandProperty "Fields" -ErrorAction Stop -all

    $contacts = [System.Collections.Generic.List[Object]]::new()
    foreach ($item in $data) {
       $fields = $item.Fields.AdditionalProperties
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


    if ($contacts.count -eq 0) {
        WriteLog('[WARNING] Problem with list or item retrieval. Either list is empty or items could not be retrieved')
        return
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


foreach($member in $members) {
    try {
        $userId = if ($targetType -eq "User") { $TestUserId } else { $member.Id }
        $memberDets = Get-MgUser -UserId $userId
        $memberId = $memberDets.UserPrincipalName

        $existingFolderSkip = Get-MgUserContactFolder -UserId $memberId | Where-Object { $_.DisplayName -eq $folderName }
    } catch {
        WriteLog("[ERROR] Error retrieving member and folder: $($_.Exception.Message)")
    }



    $hasE5 = CheckE5($memberId);

    if($hasE5){
        if($existingFolderSkip){
            #nothing, skip user
            $logmsg = "[INFO] Skipped $($memberDets.DisplayName), $($memberId)"
            WriteLog($logmsg) 
        } else {
            try{
                New-MgUserContactFolder -userid $memberId -DisplayName $folderName
                $folder = Get-MgUserContactFolder -userid $memberId | Where-Object { $_.DisplayName -eq $folderName }
                $folderid = $folder.id
            } catch {
                WriteLog("[ERROR] Error creating contact folder and setting ID: $($_.Exception.Message)")
            }
    
            foreach($contact in $contacts){
                $params = @{
                    givenName = $contact.givenName
                    surname = $contact.surname
                    businessPhones = @($contact.businessPhones)
                    categories = @($folderName)
                    officeLocation = $contact.OfficeLocation
                    personalNotes = $contact.id
                }
        
                try{
                    New-MgUserContactFolderContact -userId $memberId -ContactFolderId $folderid -BodyParameter $params
                } catch {
                    WriteLog("[ERROR] Couldnt create contact: $($contact.givenName) $($contact.surname) for $($memberId): $($_.Exception.Message)")
                }
                
            }
        
            WriteLog("[INFO] Func Numbers added for $($memberDets.DisplayName), $($memberId)") 
        }
    }else {
        continue;
    }

}

Disconnect-MgGraph

$benchmark.Stop()
WriteLog("[Benchmark] $($benchmark.ElapsedMilliseconds) ms")
WriteLog('-------------------------------------------')
