<#  Modified by Josh McMullin on 12/30/19
    This script adds a new user to AD as well the following attributes:
    Automatically derives the username based on the first, last & middle initial
    (first 7 characters of last name + first character of first name unless
    already in AD and then does the first 6 characters of first name + first character of
    first name + middle initial)
    Address information including street, PO Box, City, State, Zip
    E-mail address
    Phone
    Changes the UPN, Proxy Addresses
#>

# Note the data boxes pop up behind PowerShell ISE for some reason.
# Working on fixing where the pop up box outputs to
# Note this version of the script creates the username as first initial and last name

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }


$opt = (Get-Host).PrivateData
$opt.WarningBackgroundColor = "red"
$opt.WarningForegroundColor = "white"
$opt.ErrorBackgroundColor = "green"
$opt.ErrorForegroundColor = "white"

Set-ExecutionPolicy RemoteSigned -Force -Scope CurrentUser
Write-Host "Enter your domain admin credentials." -ForegroundColor red
$UserCredential = Get-Credential

# Connect to Office 365
Write-Host "Enter Office 365 admin credentials." -ForegroundColor red
try
{
    Get-MsolDomain -ErrorAction Stop > $null
}
catch 
{
    $credential = Get-Credential
    if ($cred -eq $null) {$cred = Get-Credential $credential}
    Write-Output "Connecting to Office 365..."
    Connect-MsolService -Credential $cred
    Set-ExecutionPolicy 'RemoteSigned' -Scope Process -Confirm:$false
    Set-ExecutionPolicy 'RemoteSigned' -Scope CurrentUser -Confirm:$false
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $Credential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -AllowClobber -DisableNameChecking
}


# Enter Unique Employee Values
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
#$StreetAddress = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Address")
$City = [Microsoft.VisualBasic.Interaction]::InputBox("Enter City")
$State = [Microsoft.VisualBasic.Interaction]::InputBox("Enter State")
$PostCode = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Zip")
$Country = "US"
$Company = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Company Name")
$DNSRoot = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Domain" , "It needs to match their Office365 email domain")


$unique = $false
While($unique -eq $false){
	# Acquiring name data
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
	$GivenName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new users first name", "New User Tool - First Name", "First")
	$Initial = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new users Middle Initial", "New User Tool - Middle Initail", "Middle Initial")
	$SurName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new users last name", "New User Tool - Last Name", "Last")

	#Process that derives the username from the Given, Initial & Surname
	$SAMAccountName = $GivenName.Substring(0,1) + $Surname.Substring(0,[System.Math]::Min(20, $Surname.Length))
	Write-Verbose "$samaccountname" -Verbose


	If((Get-ADUser -Filter "samaccountname -eq '$samaccountname'" -ea Silentlycontinue)){
		Write-Warning "user $samaccountname already exists, please choose a different name" 
	}
	Else {
		Write-Verbose "$samaccountname does not exist." -Verbose
		$unique = $true
	}			
}


# Converts the samaccountname to lower case
$SAMAccountLower = $SAMAccountName.ToLower()

#Creates the display name
$DisplayName = $GivenName + " " + $Surname

#Acquires more data
#$EmpID = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new users Employee ID", "New User Tool - Employee ID", "1234")
$Title = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new users Title", "New User Tool - Title", "Clerk I")
$Office = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new users workplace, Cottonwood or Field Staff", "New User Tool - Office", "Cottonwood")
$Department = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new users Department", "New User Tool - Department", "Department")
$Phone = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new users phone. 507-423-6262 for main office, mobile # for field staff", "New User Tool - Phone", "555-555-1212")
Write-Host "Select the employee's manager" -ForegroundColor red
$Manager = Get-ADUser -Filter {enabled -eq $true} | Select-Object SamAccountName | Out-GridView -PassThru

# Process that creates email address
$Mail = $SAMAccountLower.ToLower() + "@" + $DNSRoot

# Process that creates other field data that needs to filled in for Exchange Online & Signatures
$ProxyAddress1 = "SMTP:" + $Mail
$UserPrincipalName = $Mail
$Description = $Department + " - " + $title

# Setting OU that Account will Reside in
# Suggest using search filter in pop-up "ou=user" to return user ou's
Write-Host "Select the employee's OU. Suggest using search filter in pop-up ou=user to return user ou's" -ForegroundColor red
$SelectOU = Get-ADOrganizationalUnit -Filter * | Select-Object -Property DistinguishedName | Out-GridView -PassThru | Select-Object -ExpandProperty DistinguishedName
Get-ADUser -filter {samAccountName -eq $SamAccountLower} | Move-ADObject -TargetPath $SelectOU

# set default password
$defpassword = (ConvertTo-SecureString "SomePassword" -AsPlainText -force)

#This portion displays a summary of all the data that the user has entered
[System.Windows.Forms.MessageBox]::show("Verify the following is correct:
The new user             $DisplayName       will be created with the following attributes:
Full Name:               $GivenName $Initial $Surname $Creds
Username:                $SAMAccountLower
Department/Title:        $Description
Office Location:         $Office
Phone:                   $Phone    
Email Address is:        $Mail
Manager is:              $Manager
OU is:                   $SelectOU
 
OK will continue and add the above information to the Active Directory
 
OK to Continue." , "AD New User", 1)

$splat = @{
Path = $SelectOU
SamAccountName = $SamAccountLower
GivenName = $GivenName
Initial = $Initial
Surname = $Surname
Name = $DisplayName
DisplayName = $DisplayName
EmailAddress = $Mail
UserPrincipalName = $Mail
Title = $title
Description = $Description
Enabled = $true
ChangePasswordAtLogon = $true
PasswordNeverExpires  = $false
AccountPassword = $defpassword
#EmployeeID = $EmpID
OfficePhone = $Phone
Office = $Office
Department = $Department
Manager = $Manager
#StreetAddress = $StreetAddress
City = $City
State = $State
PostalCode = $PostCode 
Company = $Company
OtherAttributes = @{proxyAddresses = ($ProxyAddress1)}
}

New-ADUser @splat -Verbose
Set-ADUser $SAMAccountLower
Set-ADUser $SAMAccountLower -add @{Co = $Country}

# Sync to Azure
$session = New-PSSession -cn "your domain controller" -Credential $UserCredential
Invoke-Command -ComputerName "your domain controller" -ScriptBlock {
    Import-Module adsync
    Start-ADSyncSyncCycle -PolicyType Delta
}

# Pause script for 5 minutes
Start-Sleep -Seconds 300

# Get available license options
# License user's mailbox 
$User = Get-MsolUser -All -UnlicensedUsersOnly | Out-GridView -Title 'Select a user' -OutputMode Single | Select-Object -ExpandProperty UserPrincipalName 
$OfficeLicenses = Get-MsolAccountSku | Out-GridView -Title 'Select a license plan' -OutputMode Single | Select-Object -ExpandProperty AccountSkuId
Set-MsolUser -UserPrincipalName $User -UsageLocation US
Set-MsolUserLicense -UserPrincipalName $User -AddLicenses $OfficeLicenses

# Sync to Azure Again
$session = New-PSSession -cn "your domain controller" -Credential $UserCredential
Invoke-Command -ComputerName "your domain controller" -ScriptBlock {
    Import-Module adsync
    Start-ADSyncSyncCycle -PolicyType Delta
}
