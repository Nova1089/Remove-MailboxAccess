<#
This script removes access to a mailbox from a provided list of users.
#>

# functions
function Initialize-ColorScheme
{
    $script:successColor = "Green"
    $script:infoColor = "DarkCyan"
}

function Show-Introduction
{
    Write-Host "This script removes access to a mailbox from a list of users." -ForegroundColor $infoColor
    Read-Host "Press Enter to continue"
}

function Use-Module($moduleName)
{    
    $keepGoing = -not(Test-ModuleInstalled $moduleName)
    while ($keepGoing)
    {
        Prompt-InstallModule($moduleName)
        Test-SessionPrivileges
        Install-Module $moduleName

        if ((Test-ModuleInstalled $moduleName) -eq $true)
        {
            Write-Host "Importing module..."
            Import-Module $moduleName
            $keepGoing = $false
        }
    }
}

function Test-ModuleInstalled($moduleName)
{    
    $module = Get-Module -Name $moduleName -ListAvailable
    return ($null -ne $module)
}

function Prompt-InstallModule($moduleName)
{
    do 
    {
        Write-Host "$moduleName module is required."
        $confirmInstall = Read-Host -Prompt "Would you like to install it? (y/n)"
    }
    while ($confirmInstall -inotmatch "(?<!\S)y(?!\S)") # regex matches a y but allows spaces
}

function Test-SessionPrivileges
{
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentSessionIsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($currentSessionIsAdmin -ne $true)
    {
        Throw "Please run script with admin privileges. 
            1. Open Powershell as admin.
            2. CD into script directory.
            3. Run .\scriptname.ps1"
    }
}

function TryConnect-ExchangeOnline
{
    $connectionStatus = Get-ConnectionInformation -ErrorAction SilentlyContinue

    while ($null -eq $connectionStatus)
    {
        Write-Host "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ErrorAction SilentlyContinue
        $connectionStatus = Get-ConnectionInformation

        if ($null -eq $connectionStatus)
        {
            Read-Host -Prompt "Failed to connect to Exchange Online. Press Enter to try again"
        }
    }
}

function Prompt-MailboxIdentifier
{
    do
    {
        $mailboxIdentifier = Read-Host "Enter name or email of mailbox"        
        if ([string]::IsNullOrWhiteSpace($mailboxIdentifier)) { continue }
        $mailbox = TryGet-Mailbox -mailboxIdentifier $mailboxIdentifier -tellWhenFound
    }
    while ($null -eq $mailbox)

    return $mailbox.UserPrincipalName
}

function TryGet-Mailbox($mailboxIdentifier, [switch]$tellWhenFound)
{
    $mailbox = Get-EXOMailbox -Identity $mailboxIdentifier -ErrorAction SilentlyContinue

    if ($null -eq $mailbox)
    {
        Write-Warning "User or mailbox not found: $mailboxIdentifier."
        return $null
    }

    if ($tellWhenFound)
    {
        Write-Host "Mailbox found!" -ForegroundColor $successColor

        [PSCustomObject]@{
            DisplayName = $mailbox.DisplayName
            UPN = $mailbox.UserPrincipalName
            Type = $mailbox.RecipientTypeDetails
        } | Format-Table | Out-Host
    }
    return $mailbox
}

function Prompt-UserListInputMethod
{
    Write-Host ("Choose user input method:`n" +
                "[1] Provide text file. (Users listed by full name or email, separated by new line.)`n" +
                "[2] Enter user list manually.")
    do
    {
        $choice = Read-Host
        
        if ($choice -inotmatch '^\s*[12]\s*$') # regex matches a 1 or 2 but allows whitespace
        {
            Write-Warning "Please enter 1 or 2."
        }
    }
    while ($choice -notmatch '^\s*[12]\s*$') # regex matches a 1 or 2 but allows whitespace

    return [int]$choice
}

function Get-UserFromTXT
{
    do
    {
        $path = Read-Host "Enter path to txt file. (i.e. C:\UserList.txt)"
        $path = $path.Trim('"') # trims quotes off path
        $userList = Get-Content -Path $path -ErrorAction SilentlyContinue 

        if ($null -eq $userList)
        {
            Write-Warning "File not found or contents are empty."
            $keepGoing = $true
        }
        else
        {
            Write-Host "User list found!" -ForegroundColor $successColor
            $keepGoing = $false
        }
    }
    while ($keepGoing)

    return $userList
}

function Get-UsersFromTXT
{
    do 
    {
        $path = Read-Host "Enter path to txt file. (i.e. C:\UserList.txt)"
        $path = $path.Trim('"') # trims quotes off path
        $userList = Get-Content -Path $path -ErrorAction SilentlyContinue 
        
        if ($null -eq $userList)
        {
            Write-Warning "File not found or contents are empty."
            $keepGoing = $true
            continue
        }
        else
        {
            Write-Host "User list found!" -ForegroundColor $successColor
            $keepGoing = $false
        }

        $finalUserList = New-Object -TypeName System.Collections.Generic.List[string]
        $noToAll = $false
        $i = 0
        :userListLoop foreach ($user in $userList)
        {
            if ([string]::IsNullOrWhiteSpace($user)) { continue }

            $i++
            Write-Progress -Activity "Looking up users..." -Status "$i users checked."

            $user = $user.Trim()
            $userFound = $null -ne (TryGet-Mailbox $user)
            $validUser = (($userFound) -or ($user -imatch '^S-\d(-\d+){6}$')) # regex matches an SID
            
            if ($validUser)
            {
                $finalUserList.Add($user)
            }
            else
            {                
                if ($noToAll)
                {
                    continue
                }
                else
                {
                    $fixFile = Prompt-YesOrNo -question "Would you like to fix the file and try again?" -includeNoToAll

                    switch ($fixFile)
                    {
                        'Y' 
                        { 
                            $keepGoing = $true
                            break userListLoop
                        }
                        'N' 
                        { 
                            $keepGoing = $false
                            continue
                        }
                        'L' # no to all
                        {
                            $keepGoing = $false
                            $noToAll = $true
                            continue
                        }
                    }
                }
            }
        }
    }
    while ($keepGoing)

    return $finalUserList
}

function Prompt-YesOrNo($question, [switch]$includeYesToAll, [switch]$includeNoToAll)
{
    $prompt = ("$question`n" + 
                "[Y] Yes  [N] No")
    
    if ($includeYesToAll -and $includeNoToAll)
    {
        $prompt += "  [A] Yes to All  [L] No to All"

        $response = Read-HostAndValidate -prompt $prompt -regex '^\s*[ynal]\s*$' -warning "Please enter y, n, a, or l."
    }
    elseif($includeYesToAll)
    {
        $prompt += "  [A] Yes to All"

        $response = Read-HostAndValidate -prompt $prompt -regex '^\s*[yna]\s*$' -warning "Please enter y, n, or a."
    }
    elseif($includeNoToAll)
    {
        $prompt += "  [L] No to All"

        $response = Read-HostAndValidate -prompt $prompt -regex '^\s*[ynl]\s*$' -warning "Please enter y, n, or l."        
    }
    else
    {
        $response = Read-HostAndValidate -prompt $prompt -regex '^\s*[yn]\s*$' -warning "Please enter y or n." 
    }

    return $response.Trim().ToUpper()
}

function Read-HostAndValidate($prompt, $regex, $warning)
{
    Write-Host $prompt

    do
    {
        $response = Read-Host

        if ($response -inotmatch $regex)
        {
            Write-Warning $warning
        }
    }
    while ($response -inotmatch $regex)

    return $response
}

function Get-UsersManually
{
    $userList = New-Object -TypeName System.Collections.Generic.List[string]

    while ($true)
    {
        $response = Read-Host "Enter a user (full name or email) or type `"done`""
        if ($response -imatch '(?<!\S)done(?!\S)') { break } # regex matches the word done but allows spaces
        if ($null -eq (TryGet-Mailbox $response -tellWhenFound)) { continue }
        $userList.Add($response)
    }

    return $userList
}

function Prompt-PermissionsToRemove
{
    do
    {
        $removeFullAccess = Prompt-YesOrNo "Remove `"full access`" (read and manage) from all users?"
        $removeSendAs = Prompt-YesOrNo "Remove `"send as`" access? from all users?"

        if (($removeFullAccess -eq "N") -and ($removeSendAs -eq "N"))
        {
            Write-Warning "No permissions were selected to remove."
            $keepGoing = $true
        }
        else
        {
            $keepGoing = $false
        }
    }
    while ($keepGoing)

    return [PSCustomObject]@{
        removeFullAccess = $removeFullAccess
        removeSendAs = $removeSendAs
    }
}

function Remove-AccessToMailbox($mailboxIdentifier, $userList, $removeFullAccess, $removeSendAs, [switch]$logChanges)
{
    Read-Host "Press Enter to begin processing"

    if ($logChanges)
    {
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $timeStamp = New-TimeStamp
        $path = "$desktopPath\Remove Mailbox Access Logs $timeStamp.csv"
    }
    
    $i = 0
    foreach ($user in $userList)
    {        
        $fullAccessRemoved = $false
        $sendAsRemoved = $false
        
        Write-Progress -Activity "Removing access to mailbox..." -Status "$i users removed"

        if ($removeFullAccess -eq "Y")
        {
            Remove-MailboxPermission -Identity $mailboxIdentifier -User $user -AccessRights FullAccess -Confirm:$false -WarningAction SilentlyContinue
            $fullAccessRemoved = $true
        }

        if ($removeSendAs -eq "Y")
        {
            Remove-RecipientPermission -Identity $mailboxIdentifier -Trustee $user -AccessRights SendAs -Confirm:$false -WarningAction SilentlyContinue
            $sendAsRemoved = $true
        }

        if ($logChanges)
        {
            [PSCustomObject]@{
                Mailbox = $mailboxIdentifier
                Delegate = $user
                FullAccessRemoved = $fullAccessRemoved
                SendAsRemoved = $sendAsRemoved
            } | Export-Csv -Path $path -Append -NoTypeInformation
        }

        $i++
    }
    Write-Progress -Activity "Removing access to mailbox..." -Status "$i users removed."
    Write-Host "Finished removing access from $i users. (If they had access to begin with.)" -ForegroundColor $successColor

    if ($logChanges)
    {
        Write-Host "Changes logged to $path" -ForegroundColor $successColor
        Write-Host "Please note: The logs indicate if an attempt was made to remove the access, not if they had access to begin with." -ForegroundColor $infoColor
    }
}

function New-TimeStamp
{
    return (Get-Date -Format yyyy-MM-dd-hh-mm).ToString()
}

# main
Initialize-ColorScheme
Show-Introduction
Use-Module("ExchangeOnlineManagement")
TryConnect-ExchangeOnline
$mailboxIdentifier = Prompt-MailboxIdentifier
$userListInputMethod = Prompt-UserListInputMethod
switch ($userListInputMethod)
{
    1 { $userList = Get-UsersFromTXT }
    2 { $userList = Get-UsersManually }
}

$permissionSelections = Prompt-PermissionsToRemove
Remove-AccessToMailbox -mailboxIdentifier $mailboxIdentifier -userList $userList -removeFullAccess $permissionSelections.removeFullAccess -removeSendAs $permissionSelections.removeSendAs -logChanges
Read-Host -Prompt "Press Enter to exit"