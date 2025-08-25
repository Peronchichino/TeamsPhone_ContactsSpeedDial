
# DISCLAIMER:
# theoretically this new script and the implementations mentioedn in "Testmark.ps1" could improve the efficiency of the original script by around 5,5x
# now includes better data structure usage, a better sorting algorithm (no longer an object sort but a unique key comparison aka binary), and process batching of API calls
# theoretical API calls have been reduced down from 9.009.000 to 456.001
# total runtime has theoreticall reduced from 39 hours on average to around 7 hours





# Parameters
$batchSize = 20
$folderName = "Func Numbers"
$TestUserId = "<yourTestUserId>"
$targetType = "User" # or "Group"
$contactListSiteCollectionId = "<yourSiteCollectionId>"
$contactListId = "<yourListId>"
$e5GroupId = "<yourE5GroupId>"

# Logging Function (assumed implemented)
function WriteLog { param($msg) ; Write-Output $msg }

# Fetch All Contacts Once
[array]$data = Get-MgSiteListItem -SiteId $contactListSiteCollectionId -ListId $contactListId -ExpandProperty "Fields" -All

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
WriteLog "[INFO] Loaded $($contacts.Count) contacts."

# Fetch All Users into Hash Table
$userIdTable = @{}
foreach ($member in $members) {
    try {
        $userId = if ($targetType -eq "User") { $TestUserId } else { $member.Id }
        $memberDets = Get-MgUser -UserId $userId
        $userIdTable[$member.Id] = $memberDets.UserPrincipalName
    } catch {
        WriteLog "[ERROR] Error retrieving member: $($_.Exception.Message)"
    }
}
WriteLog "[INFO] Prefetched $($userIdTable.Count) users."

# Fetch E5 Group Members Once
$e5GroupMembers = Get-MgGroupMember -GroupId $e5GroupId -ConsistencyLevel eventual -All
$e5MembersTable = @{}
foreach ($member in $e5GroupMembers) {
    $e5MembersTable[$member.UserPrincipalName] = $true
}
WriteLog "[INFO] Loaded $($e5MembersTable.Count) E5 members."

# Function: Check if User is E5
function IsE5Member {
    param($userPrincipalName)
    return $e5MembersTable.ContainsKey($userPrincipalName)
}

# Main Processing Loop
foreach ($member in $members) {
    $memberId = $userIdTable[$member.Id]
    if (-not $memberId) {
        WriteLog "[WARN] Skipping member with missing UserPrincipalName: $($member.Id)"
        continue
    }

    if (-not (IsE5Member $memberId)) {
        continue
    }

    try {
        # Check for existing folder
        $existingFolder = Get-MgUserContactFolder -UserId $memberId | Where-Object { $_.DisplayName -eq $folderName }

        if ($existingFolder) {
            WriteLog "[INFO] Skipped $($memberId) (folder exists)"
            continue
        }

        # Create folder
        New-MgUserContactFolder -UserId $memberId -DisplayName $folderName
        $folder = Get-MgUserContactFolder -UserId $memberId | Where-Object { $_.DisplayName -eq $folderName }
        $folderId = $folder.Id

        # Process contacts in batches
        for ($i = 0; $i -lt $contacts.Count; $i += $batchSize) {
            $batch = $contacts[$i..([math]::Min($i + $batchSize - 1, $contacts.Count - 1))]

            foreach ($contact in $batch) {
                $params = @{
                    givenName      = $contact.givenName
                    surname        = $contact.surname
                    businessPhones = @($contact.businessPhones)
                    categories     = @($folderName)
                    officeLocation = $contact.officeLocation
                    personalNotes  = $contact.id
                }

                try {
                    New-MgUserContactFolderContact -UserId $memberId -ContactFolderId $folderId -BodyParameter $params
                } catch {
                    WriteLog "[ERROR] Failed to create contact $($contact.givenName) $($contact.surname) for $($memberId): $($_.Exception.Message)"
                }
            }
        }

        WriteLog "[INFO] Func Numbers added for $($memberId)"

    } catch {
        WriteLog "[ERROR] Error processing user $($memberId): $($_.Exception.Message)"
    }
}

WriteLog "[INFO] Contact import completed."
