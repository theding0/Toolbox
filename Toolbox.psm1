<#
	.SYNOPSIS
	    A function to connect to Office 365 for administrative purposes. Imports the module from Office 365 for Msol cmdlets

	.EXAMPLE
	    Connect-o365
#>
function Connect-O365{
	$o365cred = Get-Credential username@domain.onmicrosoft.com
	$session365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell/" -Credential $o365cred -Authentication Basic -AllowRedirection 
    Connect-MsolService -Credential $o365cred
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
    BEGIN {

    $user_dn = (Get-Mailbox $user).distinguishedname
    $Group = @(Get-DistributionGroup -ResultSize Unlimited)
    $Distros = @()
    
    } # End BEGIN block
        
    PROCESS {

        foreach ($g in $Group)
        {
            $i = 1
            Write-Progress -Activity "Collecting distribution groups" -Status "Checking $g" -PercentComplete ($i/$Group.Count)
            if ((Get-DistributionGroupMember $g.identity | select -expand distinguishedname) -contains $user_dn)
            {
                
                $Distros += $g
            
            } # End identity check
            $i++
        } # End foreach $Group
    } #End PROCESS block

    END {

    Write-Output $Distros


    } # End END block
} # End Get-Distro function
