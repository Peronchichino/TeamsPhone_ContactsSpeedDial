$csvLogging = "C:\LOG_InitialContactAdd_Users.txt"
$thumb = ""
$applicationID = ""
$tenantID =""
$folderName = "Custom Contacts"
$groupId = ""
$csvFilePath = "C:\Import\contacts.csv"

$e5groupId = "" //e5 license check

                                                                                                                                                      



function WriteLog{    
    Param ([string]$logString)
    $dateTime = "[{0:dd/MM/yy} {0:HH:mm:ss}]" -f (Get-Date)
    # If File not exists use Add-Content to create it and add content
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
	

    $members = Get-MgGroupMember -GroupId $groupId -ConsistencyLevel eventual -all

    if($members.count -eq 0){
        WriteLog('[WARNING] Problem with retrieving group or group is empty')
        return
    }

    try{
        $importedContacts = Import-Csv -Path $csvFilePath -Delimiter ";" -Header @("PersonalNotes", "GivenName", "Surname", "email", "BusinessPhones", "MobilePhone", "Department")
        
        if($importedContacts -eq $null){
            WriteLog('[WARNING] No contacts to be added')
        }
    } catch [System.IO.FileNotFoundException]{
        WriteLog("[ERROR] $($_.Exception.FileName)") 
    }
} catch {
    WriteLog("[ERROR] Error/Exception in List and Member retrieval: $($_.Exception.Message)")
    WriteLog("[ERROR] $($_.Exception.StackTrace)")
}








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
    
            foreach($contact in $importedContacts){
                $params = @{
                    GivenName = $contact.GivenName
                    Surname = $contact.Surname
                    BusinessPhones = @($contact.BusinessPhones)
                    MobilePhone = $contact.MobilePhone
                    Department = $contact.Department
                    Categories = @($folderName)
                    PersonalNotes = $contact.PersonalNotes
                }
                try{
                    New-MgUserContactFolderContact -UserId $memberId -ContactFolderId $folderid -BodyParameter $params
                } catch {
                    WriteLog("[ERROR] Couldnt create contact: $($contact.PersonalNotes) for $($memberId): $($_.Exception.Message)")
                }
            }
    
            WriteLog("[INFO] Contacts added for $($memberDets.DisplayName), $($memberId)") 
        }
    }else {
        continue;
    }


}

Disconnect-MgGraph

$benchmark.Stop()
WriteLog("[Benchmark] $($benchmark.ElapsedMilliseconds) ms")
WriteLog('-------------------------------------------')


