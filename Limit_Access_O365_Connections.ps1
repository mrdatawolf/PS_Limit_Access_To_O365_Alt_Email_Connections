param (
    [string]$email,
    [string]$password,
    [string]$scriptRoot
)

if (-not $scriptRoot) {
    $scriptRoot = $PSScriptRoot
}
$envFilePath = Join-Path -Path $scriptRoot -ChildPath ".env"
if (Test-Path $envFilePath) {
    Get-Content $envFilePath | ForEach-Object {
        if ($_ -match "^\s*([^#][^=]+?)\s*=\s*(.*?)\s*$") {
            [System.Environment]::SetEnvironmentVariable($matches[1], $matches[2])
        }
    }
} else {
    Write-Host "You need a .env setup before running this!" -ForegroundColor Red
    Pause
    exit
}
$baseFolderPath = [System.Environment]::GetEnvironmentVariable("BaseFolderPath")
$aEmailsFilePath = [System.Environment]::GetEnvironmentVariable("AEmailsFilePath")
$uEmailsFilePath = [System.Environment]::GetEnvironmentVariable("UEmailsFilePath")
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    if ($null -ne $email -and $null -ne $password) {
        $arguments = "-File `"$($myinvocation.MyCommand.Definition)`" -email `"$email`" -password `"$password`" -scriptRoot `"$scriptRoot`""
        Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs
    } else {
        $arguments = "-File `"$($myinvocation.MyCommand.Definition)`" -scriptRoot `"$scriptRoot`""
        Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs
    }
    exit
}
function Test-ModuleInstallation {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ModuleName
    )

    if (!(Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "The $ModuleName module is not installed. Installing..." -ForegroundColor Yellow
        Install-Module -Name $ModuleName -Force
        
        return $false
    } else {
        Write-Host "Importing $ModuleName..." -ForegroundColor Green
        Import-Module $ModuleName
    }

    return $true
}

function Set-UserData {
    param (
        [Parameter(Mandatory=$true)]
        [string]$BaseFolderPath,
        [string]$username,
        [SecureString]$securePassword,
        [array]$userEmailExcelData,
        [bool]$skipIMAP,
        [bool]$skipPop,
        [bool]$skipOWA,
        [bool]$skipSMTP
    )
    try {
        $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $securePassword
        Connect-AzureAD -Credential $credential 2>$null
        Connect-ExchangeOnline -Credential $credential -ShowBanner:$false
    }
    catch {
        Write-Host "Username/password login failed.  We will try again with a user prompted login" -ForegroundColor Yellow
        Connect-AzureAD
        Connect-ExchangeOnline
    }
    $users = Get-AzureADUser -All $true
    foreach ($user in $users) {
        if ($null -ne $user) {
            $userData = $userEmailExcelData | Where-Object { $_.username -eq $user.UserPrincipalName } 
            if ($null -eq $userData) {
                Write-Host "No data found for user $($user.UserPrincipalName)"
                continue
            }
            $mailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            if ($null -eq $mailbox) {
                Write-Host "No mailbox found for user $($user.UserPrincipalName)"
                continue
            }
            $skipIMAP = ($userData.IMAP_override -eq "true" -or $userData.IMAP_override -eq $true)
            $skipPop = ($userData.POP_override -eq "true" -or $userData.POP_override -eq $true)
            $skipOWA = ($userData.OWA_override -eq "true" -or $userData.OWA_override -eq $true)
            $skipSMTP = ($userData.SMTP_override -eq "true" -or $userData.SMTP_override -eq $true)
            if($false -ne $skipImap) {
                Set-CASMailbox -Identity $user.UserPrincipalName -ImapEnabled $false
            }
            if($false -ne $skipPop) {
                Set-CASMailbox -Identity $user.UserPrincipalName -PopEnabled $false
            }
            if($false -ne $skipOwa) {
                Set-CASMailbox -Identity $user.UserPrincipalName -OwaEnabled $false
            }
            if($false -ne $skipOwa) {
                Set-CASMailbox -Identity $user.UserPrincipalName -SmtpClientAuthenticationDisabled $true
            }
        }
    }
    Disconnect-AzureAD
    Disconnect-ExchangeOnline -Confirm:$false
}
$modules = @("ImportExcel", "AzureAD.Standard.Preview", "ExchangeOnlineManagement")
foreach ($module in $modules) {
    $result = Test-ModuleInstallation -ModuleName $module
    if (-not $result) {
        Write-Host "Please restart the script after installing the required modules." -ForegroundColor Red
        exit
    }
}
Write-Host "All required modules are installed and imported."
$userEmailExcelData = Import-Excel -Path $uEmailsFilePath
if ($email -and $password) {
    $securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force
    Write-Host "Running once. Current admin email: $email"
    Set-UserData -BaseFolderPath $baseFolderPath -username $email -securePassword $securePassword -userEmailExcelData $userEmailExcelData
} else {
    try {
        $excelData = Import-Excel -Path $aEmailsFilePath
    } Catch {
        Write-Host "Failed to find the admin list here: $aEmailsFilePath"
        Pause
    }
    foreach ($row in $excelData) {
        if ($row.automate -eq 1) {
            $email = $row.Email
            $securePassword = ConvertTo-SecureString -String $row.Password -AsPlainText -Force
            Write-Host "Current admin email: $email"
            Set-UserData -BaseFolderPath $baseFolderPath -username $email -securePassword $securePassword -userEmailExcelData $userEmailExcelData
        }
    } 
}
Write-Host "Completed."
Pause