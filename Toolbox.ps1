<#
	.SYNOPSIS
	    A function to connect to Office 365 for administrative purposes. Imports the module from Office 365 for Msol cmdlets

	.EXAMPLE
	    Connect-o365
#>
function Connect-O365{
	$o365cred = Get-Credential username@domain.onmicrosoft.com
	$session365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell/" -Credential $o365cred -Authentication Basic -AllowRedirection 
	Import-Module (Import-PSSession $session365 -AllowClobber) -Global
}
function Disconnect-ExchangeOnline {
    Get-PSSession | ?{$_.ComputerName -like "*outlook.com"} | Remove-PSSession
}
function Get-Distro {
    [CmdLetBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [string]$User
    )

    $user_dn = (get-mailbox $user).distinguishedname
    foreach ($group in get-distributiongroup -resultsize unlimited)
    {
        if ((get-distributiongroupmember $group.identity | select -expand distinguishedname) -contains $user_dn)
        {
            $group.name 
        }
    }
}
