$csvLogging = "C:\LOG_Sync_Users.txt"
$thumb = ""
$applicationID =""
$tenantID =""
$folderName = "Custom Contacts"
$groupId = ""
$e5groupId = ""

$date = Get-Date -Format "yyyyMMdd"


if ((Get-Date).DayOfWeek -eq "Monday") {
    $yesterday = (Get-Date).AddDays(-3).ToString("yyyyMMdd")  # Letzter Freitag
} else {
    $yesterday = (Get-Date).AddDays(-1).ToString("yyyyMMdd")  # Normaler Vortag
}



$csvFilePath = "C:\Import\added_user_"+$date+".csv"
$csvFilePathDel = "C:\Import\deleted_user_"+$date+".csv"

$csvFilePath_yesterday = "C:\Import\added_user_"+$yesterday+".csv"
$csvFilePathDel_yesterday = "C:\Import\deleted_user_"+$yesterday+".csv"

$csvFilePath_yesterday.ToString() #making sure it really is a string
$csvFilePathDel_yesterday.ToString()

function WriteLog{    
    Param ([string]$logString)
    $dateTime = "[{0:dd/MM/yy} {0:HH:mm:ss}]" -f (Get-Date)
    if (-not (Test-Path -Path $csvLogging)) {Add-Content -Path $csvLogging -Value "Start CAL Contacts Users Logging"}
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
WriteLog("[INIT] Script start run, group: $groupId")
WriteLog('-------------------------------------------')

try{
    Connect-MgGraph -TenantId $tenantID -ClientID $applicationID -CertificateThumbprint $thumb -ErrorAction Stop

    $members = Get-MgGroupMember -GroupId $groupId -ConsistencyLevel eventual -all -ErrorAction stop

    if($members.count -eq 0){
        WriteLog('[WARNING] Problem with retrieving group or group is empty')
        return
    }


        $contacts = Import-Csv -Path $csvFilePath -Delimiter ";" -Header @("PersonalNotes", "givenName", "surname", "email", "businessPhones", "mobilePhone", "Department")

        if($contacts -eq $null){
            WriteLog('[INFO] No contacts to be added')
        }



        $contactsDel = Import-Csv -Path $csvFilePathDel -Delimiter ";" -Header @("PersonalNotes", "givenName", "surname", "email", "businessPhones", "mobilePhone", "Department")

        if($contactsDel -eq $null){
            WriteLog('[INFO] No contacts to be deleted')
        }


    try{
        $bool = Test-Path -Path $csvFilePath_yesterday
        [System.IO.File]::Exists($csvFilePath_yesterday) #double check
        if($bool){
            Remove-Item -Path $csvFilePath_yesterday -ErrorAction Stop
        }

        $bool = Test-Path -Path $csvFilePathDel_yesterday
        [System.IO.File]::Exists($csvFilePathDel_yesterday)
        if($bool){
            Remove-Item -Path $csvFilePathDel_yesterday -ErrorAction Stop
        }
    } catch {
        WriteLog("[Error] $($_.Exception.Message)")
    }

} catch {
    WriteLog("[ERROR] Error/Exception in List and Member retrieval: $($_.Exception.Message)")
    WriteLog("[ERROR] $($_.Exception.StackTrace)")
}

if ($contactsDel -or $contacts){

    foreach($member in $members){
            try{
                $memberDets = Get-MgUser -userid $member.Id
                $memberId = $memberDets.UserPrincipalName
                $existingFolderSkip = Get-MgUserContactFolder -userid $memberId | Where-Object { $_.DisplayName -eq $folderName }
            } catch {
                WriteLog("[ERROR] Error retrieving member and folder: $($_.Exception.Message)")
            }

            $hasE5 = CheckE5($memberId);

            if($hasE5){

            if($existingFolderSkip){
                    $folderId = $existingFolderSkip.Id
                    try{
                        $existingContacts = Get-MgUserContactFolderContact -userid $memberId -ContactFolderId $folderId -All
                    } catch {
                        WriteLog("[ERROR] Couldnt retrieve contacts for $($memberId): $($_.Exception.Message)")
                    }
            
                    foreach($contactDel in $contactsDel){
                        try{
                            $existingContact = $existingContacts | Where-Object {$_.PersonalNotes -eq $contactDel.PersonalNotes}
            
                            if($existingContact){
                                Remove-MgUserContactFolderContact -userid $memberId -ContactFolderId $folderId -ContactId $existingContact.Id -ErrorAction Stop
                                #$msg = "[INFO] Deleted contact $($existingContact.PersonalNotes) for $($memberId)"
                            }
                        } catch{
                            WriteLog("[ERROR] Couldnt delete contact $($existingContact.PersonalNotes) for $($memberId): $($_.Exception.Message)")
                        }
                        
                    }
            
                    foreach($contact in $contacts){
                        $params = @{
                            givenName = $contact.givenName
                            surname = $contact.surname
                            businessPhones = @($contact.businessPhones)
                            mobilePhone = $contact.mobilePhone
                            Department = $contact.Department
                            categories = @($folderName)
                            PersonalNotes = $contact.PersonalNotes
                        }
                        try {
                            #$msg = "[INFO] Added contact $($contact.givenName) $($contact.Surname) $($contact.mobilePhone) to $($memberId)"
                            New-MgUserContactFolderContact -userid $memberId -ContactFolderId $folderId -BodyParameter $params -ErrorAction Stop
                        } catch {
                            WriteLog("[ERROR] Couldnt add contact $($existingContact.PersonalNotes) for $($memberId): $($_.Exception.Message)")
                        }
                    }
            
                    WriteLog("[INFO] Contacts synced for $($memberDets.DisplayName), $($memberId)")
            
                } else {
                    #nothing, skip user
                    $logmsg = "[INFO] Skipped $($memberDets.DisplayName), $($memberId)"
                    WriteLog($logmsg) 
                }
            } else {
                continue;
            }

            

        }

    } else {

    WriteLog('[INFO] no contacts to process')
}


Disconnect-MgGraph

$benchmark.Stop()
WriteLog("[Benchmark] $($benchmark.ElapsedMilliseconds) ms")
WriteLog('-------------------------------------------')
