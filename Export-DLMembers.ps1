# -------------------
# ✅ Authentication Parameters (Please Modify)
# -------------------
$tenantId     = ""
$clientId     = ""
$clientSecret = ""
$secureSecret = ConvertTo-SecureString $clientSecret -AsPlainText -Force

# Load Modules
Import-Module MSAL.PS -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

# App-only Token
$token = Get-MsalToken -ClientId $clientId -TenantId $tenantId -ClientSecret $secureSecret -Scopes "https://graph.microsoft.com/.default"
$headers = @{ Authorization = "Bearer $($token.AccessToken)" }

# Create Export Folder
$basePath = "C:\MailExports\DL"
New-Item -ItemType Directory -Path $basePath -Force | Out-Null

# Excel Export Function
function Export-ToExcel {
    param ($data, $path, $groupName)
    if ([string]::IsNullOrWhiteSpace($groupName)) { $groupName = "Unnamed" }
    $safeName = $groupName -replace '[\\/:*?"<>|]', '_'
    $fullPath = Join-Path $path "$safeName.xlsx"
    $data | Export-Excel -Path $fullPath -WorksheetName "Members" -AutoSize -TableName "Members"
    Write-Host "✔ Exported: $safeName → $fullPath"
}

# Get Group Members
function Get-DirectMembers {
    param ($groupId, $headers)
    $members = @()
    $memberUrl = "https://graph.microsoft.com/v1.0/groups/$groupId/transitiveMembers"
    do {
        $resp = Invoke-RestMethod -Uri $memberUrl -Headers $headers
        foreach ($m in $resp.value) {
            $type = if ($m.'@odata.type' -eq '#microsoft.graph.user') {
                if ($m.userType -eq 'Guest') { 'Guest' } else { 'User' }
            } elseif ($m.'@odata.type' -eq '#microsoft.graph.group') {
                'Group'
            } else {
                'Other'
            }
            $members += [PSCustomObject]@{
                DisplayName = $m.displayName
                Email       = if ($m.mail) { $m.mail } else { $m.userPrincipalName }
                Type        = $type
            }
        }
        $memberUrl = $resp.'@odata.nextLink'
    } while ($memberUrl)
    return $members
}

# Get All Distribution Lists
function Get-DLGroups {
    $uri = 'https://graph.microsoft.com/v1.0/groups?$select=id,displayName,mailEnabled,securityEnabled,groupTypes'
    $groups = @()
    do {
        $res = Invoke-RestMethod -Uri $uri -Headers $headers
        $groups += $res.value
        $uri = $res.'@odata.nextLink'
    } while ($uri)
    return $groups | Where-Object {
        $_.mailEnabled -eq $true -and $_.securityEnabled -eq $false -and ($_.groupTypes -eq $null -or $_.groupTypes.Count -eq 0)
    }
}

$dlGroups = Get-DLGroups

foreach ($group in $dlGroups) {
    $members = Get-DirectMembers -groupId $group.id -headers $headers
    $guests = $members | Where-Object { $_.Type -eq "Guest" }
    $nonGuests = $members | Where-Object { $_.Type -ne "Guest" }

    Export-ToExcel -data $nonGuests -path $basePath -groupName $group.displayName
    if ($guests.Count -gt 0) {
        Export-ToExcel -data $guests -path "$basePath\Guests" -groupName $group.displayName
    }
}

Write-Host "`n✅ Completed: All DL lists and members (guests separated) have been exported."
