$csvLogging = "C:\Service\TeamsPhone_Contacts_Sync\Logging\LOG_InitialContactAdd_Users.txt"
$thumb = ""
$applicationID =""
$tenantID =""
$folderName = "CAL Kontakte"
$groupId = ""
$csvFilePath = "C:\Service\TeamsPhone_Contacts_Sync\Import\contacts.csv"

$e5groupId = ""
                                                                                                                                                 
$batchSize = 20


function WriteLog{    
    Param ([string]$logString)
    $dateTime = "[{0:dd/MM/yy} {0:HH:mm:ss}]" -f (Get-Date)

    if (-not (Test-Path -Path $csvLogging)) {Add-Content -Path $csvLogging -Value "Start CAL Contacts Users Logging"}
    Add-content $csvLogging -value "$datetime $logString"
}

function IsE5Member{
    param($userPrincipalName)
    return $e5MembersTable.ContainsKey($userPrincipalName)
}

$benchmark = [System.Diagnostics.Stopwatch]::StartNew()

WriteLog('-------------------------------------------')
WriteLog("[INIT] Script start run, group or user: $groupId $userId")
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

$userIdTable = @{}
foreach($member in $members){
    try{
        $userId = $member.Id
        $memberDets = Get-MgUser -userid $userId
        $userIdTable[$member.Id] = $memberDets.UserPrincipalName
        #Write-Host $userIdTable[$member.Id]
    } catch {
        WriteLog("[ERROR] Error retrieving member: $($_.Exception.Message)")
    }
}

try{
    $e5GroupMembers = Get-MgGroupMember -GroupId $e5groupId -ConsistencyLevel eventual -all | Where-Object {$_.OdataType -eq "#microsoft.graph.user"}
    $e5MembersTable = @{}
    foreach($member in $e5GroupMembers){
        $e5MembersTable[$member.UserPrincipalName] = $true
    }
} catch {
    WriteLog("[ERROR] Error/Exception in e5 Group member extraction: $($_.Exception.Message)")
    WriteLog("[ERROR] $($_.Exception.StackTrace)")
}

foreach($member in $members){
    try{
        $memberId = $userIdTable[$member.Id]
        Write-Host $memberId
        if(-not $memberId){
            WriteLog("[WARNING] Skipping member with missing UPN: $($member.Id)")
            continue
        }

        if(IsE5Member $memberId){
            #Write-Host "Loop e5 func call"
            continue
        } 
    
        $existingFolderSkip = Get-MgUserContactFolder -userid $memberId | Where-Object { $_.DisplayName -eq $folderName }
    
        if($existingFolder){
            WriteLog("[INFO] SKipped $($memberId), folder exists")
            #Write-Host "Skipped user existing folder"
            continue
        } else {
            New-MgUserContactFolder -UserId $memberId -DisplayName $folderName
            $folder = Get-MgUserContactFolder -UserId $memberId | Where-Object {$_.DisplayName -eq $folderName }
            $folderId = $folder.Id

            try{
                for($i = 0; $i -lt $contacts.count; $i+=20){
                    $batchContacts = $contacts[$i..([math]::Min($i + $batchSize - 1, $contacts.Count - 1))]
                    
                    $batchRequests = @()
                    $requestId = 1
                
                    #for each contact (temp var) in contacts ready for batching aka max 20 pieces
                    foreach($contact in $batchContacts){
                        # request
                        $req = @{
                            id = "$counter"
                            method = "POST"
                            url = "/users/$memberId/contactFolders/$folderId/contacts"
                            body = @{
                                GivenName = $contact.GivenName
                                Surname = $contact.Surname
                                BusinessPhones = @($contact.BusinessPhones)
                                MobilePhone = $contact.MobilePhone
                                Department = $contact.Department
                                Categories = @($folderName)
                                PersonalNotes = $contact.PersonalNotes
                            }
                            headers = @{
                                "Content-Type" = "application/json"
                            }
                        }

                        $batchRequests += $req
                        $counter++
                        
                    }

                    # batch body
                    $batchBody = @{requests = $batchRequests} | ConvertTo-Json -Depth 5

                    #$batchBody | Out-File "C:\Users\82816\Log Files\batchbody.json"

                    # send the shit
                    try{
                        $response = Invoke-MgGraphRequest -Method POST -uri "https://graph.microsoft.com/v1.0/`$batch" -body $batchBody -ErrorAction stop

                        foreach ($resp in $response.responses) {
                            if ($resp.status -ge 200 -and $resp.status -lt 300) {
                                Write-Host "Contact added successfully (Request ID: $($resp.id))"
                                
                            } else {
                                #Write-Warning "Failed to add contact (Request ID: $($resp.id)). Status: $($resp.status)"
                                WriteLog("[WARNING] Failed to add contact (Request ID:  $($resp.id)). Status: $($resp.status))")
                            }
                        }


                    } catch {
                        WriteLog("[ERROR] Error/Exception in batch requests loop: $($_.Exception.Message)")
                        WriteLog("[ERROR] $($_.Exception.StackTrace)")
                    }
                
                }
            } catch {
                WriteLog("[ERROR] Error/Exception in batch requests: $($_.Exception.Message)")
                WriteLog("[ERROR] $($_.Exception.StackTrace)")
            }

            WriteLog("[INFO] Func Numbers added for $($memberId)")
        }
    
    } catch {
        WriteLog("[ERROR] Error/Exception in member loop $($memberId): $($_.Exception.Message)")
        WriteLog("[ERROR] $($_.Exception.StackTrace)")
    }


}

Disconnect-MgGraph

$benchmark.Stop()
$t = $benchmark.Elapsed.ToString("dd\.hh\:mm\:ss\.ff")
WriteLog("[Benchmark] $($t) (DD/HH/MM/SS/MS)")
WriteLog('-------------------------------------------')
