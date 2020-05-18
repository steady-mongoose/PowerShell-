    $credential = Get-Credential
    if ($cred -eq $null) {$cred = Get-Credential $credential}
    Write-Output "Connecting to Office 365..."
    Connect-MsolService -Credential $cred
    Set-ExecutionPolicy 'RemoteSigned' -Scope Process -Confirm:$false
    Set-ExecutionPolicy 'RemoteSigned' -Scope CurrentUser -Confirm:$false
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $Credential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -AllowClobber -DisableNameChecking