<#
	.SYNOPSIS
	    A function to connect to Office 365 for administrative purposes. Imports the module from Office 365 for Msol cmdlets
    .PARAMETER Username
        The login id/e-mail used for the administrator account

    .PARAMETER Password
        Password used for the administrative account.

        NOTE: Passwords with special characters such as '$' must be enclosed in single quotes

	.EXAMPLE
        Connect-O365

        Generates a password prompt for Office 365 credentials

    .EXAMPLE
        Connect-o365 bingo@bango.onmicrosoft.com

        Generates a credential prompt with username filled in

    .EXAMPLE
        Connect-o365 bingo@bango.onmicrosoft.com 'Pas$word'

        Connect to Office 365 by passing credentials in as parameters

#>
function Connect-O365 {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$false)]
        [String]$Username,
        [Parameter(Position=1,Mandatory=$false)]
        [String]$Password,
        [Parameter(Position=2,Mandatory=$false,ValueFromPipeline=$true)]
        [PSCredential]$Cred  

    )
    BEGIN
    {
       
       if ((!$Password) -and (!$Cred)) { $Cred = Get-Credential $Username }
       else {
       
        $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
        $Cred = New-Object System.Management.Automation.PSCredential($Username,$secpasswd)

       }
    }
    PROCESS
    {
        
        $session365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/﻿" -Credential $cred -Authentication Basic -AllowRedirection 

        Connect-MsolService -Credential $cred
    }
    END
    {
	    Import-Module (Import-PSSession $session365 -AllowClobber) -Global
    }
}
function Disconnect-O365 {
    Get-PSSession | ?{$_.ComputerName -like "*office365.com"} | Remove-PSSession
}
<#
.NAME
    Get-Distro

.SYNOPSIS
    Provide a list of distribution groups a user is a member of

.DESCRIPTION
    Checks every distribution group for a member with the same distinguished name as the user given

.PARAMETERS User
    Specify the user

.OUTPUTS
    Array of Group objects

.EXAMPLE
    Get-Distro djones
#>
function Get-Distro {
    
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True)]
        [string[]]$User
    )
    BEGIN {

        $Group = @(Get-DistributionGroup -ResultSize Unlimited)
        $Distros = @()




    } # End BEGIN block


    PROCESS 
    {
        
        for($ucount=0; $ucount -lt $User.Count; $ucount++) {
            $Distros += New-object -TypeName psobject -Property @{Name = $((Get-Mailbox $User[$ucount]).displayname);MemberOf = @() -split ","} 
            $user_dn = $((Get-Mailbox $User[$ucount]).distinguishedname)
            for($i=0; $i -lt $Group.Count; $i++) {
                Write-Progress -Activity "Collecting distribution groups for $((Get-Mailbox ($User[$ucount])).displayname)" -Status "Checking $($Group[$i].DisplayName)" -PercentComplete ($i / $Group.Count * 100)
                if ((Get-DistributionGroupMember $Group[$i].identity | select -expand distinguishedname) -contains $user_dn) {
                    $Distros[$ucount].MemberOf += $Group[$i]
                } # End identity check
            } # End group/progress FOR
        } # End user count FOR
    } # End PROCESS block



    END {

    $Distros | ft -AutoSize -Wrap

    } # End END block
} # End Get-Distro function

<#
.SYNOPSIS
    Command prompt Telnet replacement

.DESCRIPTION
    Creates a socket object and attempts to connect that object to the specified server on the specified port

.PARAMETER Server
    Multiple servers can be tested by comma seperated IP addresses 

.PARAMETER Port
    Port number to be tested

.EXAMPLE
    Get-Telnet 10.150.1.10 25

    Test a single server for a connection on SMTP port

.EXAMPLE
    Get-Telnet 10.150.1.10,10.150.1.12 80

    Check for open HTTP port on multiple servers

.OUTPUT
    PSObject Array

#>
Function Get-Telnet{   Param (
        [CmdletBinding()]
        [Parameter(ValueFromPipeline=$true)]
        [String[]]$Commands = @("username"),
        [string]$RemoteHost = @(),
        [string]$Port = "23",
        [int]$WaitTime = 1000
        
    )
    #Attach to the remote device, setup streaming requirements
    $Socket = New-Object System.Net.Sockets.TcpClient($RemoteHost, $Port)
    if ($Socket)
    {   
        $Stream = $Socket.GetStream()
        $Writer = New-Object System.IO.StreamWriter($Stream)
        $Buffer = New-Object System.Byte[] 1024 
        $Encoding = New-Object System.Text.AsciiEncoding

        #Now start issuing the commands
        ForEach ($Command in $Commands) 
        {   
            $Writer.WriteLine($Command) 
            $Writer.Flush()
            Start-Sleep -Milliseconds $WaitTime
        }
        #All commands issued, but since the last command is usually going to be
        #the longest let's wait a little longer for it to finish
        Start-Sleep -Milliseconds ($WaitTime * 4)
        $Result = @()
        #Save all the results
        While($Stream.DataAvailable) 
        {   
            $Read = $Stream.Read($Buffer, 0, 1024) 
            $Result += ($Encoding.GetString($Buffer, 0, $Read))
        }
    }
    else     
    {   
        $Result = "Unable to connect to host: $($RemoteHost):$Port"
    }
    
    $Result 
}

#region BRANDED REPORTS

<#
	.SYNOPSIS
	    Create a branded HTML report with all mailboxes and their associated e-mail addresses
    .PARAMETER Identity
        Alias for mailboxes to be included in the report. If not specified, a report will be generated including every mailbox in the organization. Multiple identities is supported.

    .PARAMETER Path
        Path to export the HTML report. If not specified, the report will be generated in the current users Documents folder inside a folder named "Powershell Reports"

    .EXAMPLE
        Create-EmailAddressReport

        Creates an email address report for all mailboxes at the default location

	.EXAMPLE
        Create-EmailAddressReport -Path C:\emailreport.html

        Generates a report showing all mailboxes and their associated email addresses at the root of C:\

    .EXAMPLE
        Create-EmailAddressReport jmetcalf

        Creates a report for a single user


    .EXAMPLE
        Create-EmailAddressReport jmetcalf,afasl

        Creates an email address report for multiple specific users at the root of C:\

#>
function Create-EmailAddressReport {
    [CmdletBinding()]
    param (
    [Parameter(ValueFromPipeline=$true)]
    [String[]]$Identity,
    [String]$Path
    )
    BEGIN {
    $MCount = 0
    $css = Get-Content "$PSScriptRoot\style.css"	
    $image = "$PSScriptRoot\Logo.png"
    $html = "<html><head><title>$((Get-MsolCompanyInformation).DisplayName) - E-mail Address Report</title><style>$css</style></head><body><img class='logo' src=$image alt='Salvus TG'><table border='1'><tr><th>Mailbox</th><th>Associated E-mail Addresses</th></tr>"
    if ($Identity -eq "") { $Mailbox = Get-Mailbox }
    elseif ($Identity.Count -gt 1) {
        foreach ($i in $Identity) {
            $Mailbox += (Get-Mailbox $i)
        }
    }
    else { $Mailbox = Get-Mailbox $Identity }
    if (!$Path) { $Path = [Environment]::GetFolderPath("MyDocuments") + "\Powershell Reports" }
    if (!(Test-Path $Path)) { New-Item $Path -ItemType Directory }

    } # END BEGIN BLOCK
    PROCESS {
        foreach ($m in $Mailbox) {
        Write-Progress -Activity "Generating report" -PercentComplete ($i / $Mailbox.Count * 100)
            if ($m.DisplayName -match 'Discovery Search') { }
            else { $html = $html + "<tr><td>$($m.DisplayName)</td><td>" 
                foreach ($e in $m.EmailAddresses) {
                    if (($e -match "X400") -or ($e -match "X500")) { }
                    else { $html = $html + "$($e -replace "smtp:", " ")<br>" }
                }
            }
            $html = $html + "</td></tr>"
            $i++
        } # End Foreach statement
        $html = $html + "</table></body></html>"
    } # END PROCESS BLOCK
        
    END {
        $html | out-file $Path\Email_Address_Report.html
        Invoke-Item $Path\Email_Address_Report.html
    } # END END BLOCK
}
function Get-MailboxAccessReport {
[CmdletBinding()]
    param (
    [Parameter(ValueFromPipeline=$true)]
    [String[]]$Identity,
    [String]$Path
    )
BEGIN{
$Mailbox = Get-Mailbox
    $css = Get-Content "$PSScriptRoot\style.css"
    $image = "$PSScriptRoot\Logo.png"
    $html = "<html><head><title>$((Get-MsolCompanyInformation).DisplayName) - Mailbox Access Report</title><style>$css</style></head><body><img class='logo' src=$image alt='Salvus TG'><table border='1'><tr><th>Mailbox</th><th>User's With Full Access</th></tr>"
    if ($Identity -eq "") { $Mailbox = Get-Mailbox }
    elseif ($Identity.Count -gt 1) {
        foreach ($i in $Identity) {
            $Mailbox += (Get-Mailbox $i)
        }
    }
    else { $Mailbox = Get-Mailbox $Identity }
    if (!$Path) { $Path = [Environment]::GetFolderPath("MyDocuments") + "\Powershell Reports" }
    if (!(Test-Path $Path)) { New-Item $Path -ItemType Directory }

    } # END BEGIN BLOCK
    PROCESS {
        foreach ($m in $Mailbox) {
        
            if ($m.DisplayName -match 'Discovery Search') { }
            else { $html = $html + "<tr><td>$($m.DisplayName)</td><td>" 
                foreach ($a in (Get-MailboxPermission $m.Id)) {
                    if (($a.User -match "X400") -or ($a.User -match "X500") -or ($a.User -match "admin") -or ($a.User -match "NT AUTHORITY") -or ($a.User -match "Management") -or ($a.User -match "Exchange") -or ($a.User -match "Folder") -or ($a.User -match "Delegated") -or ($a.User -match "S-1-5") -or ($a.User -match "Managed")) { }
                    else { $html = $html + "$($a.User)<br>" }
                }
            }
            $html = $html + "</td></tr>"
            $i++
        } # End Foreach statement
        $html = $html + "</table></body></html>"
    } # END PROCESS BLOCK
        
    END {
        $html | out-file $Path\Mailbox_Access_Report.html
        Invoke-Item $Path\Mailbox_Access_Report.html
    } # END END BLOCK
} # End Get-MailboxAccessReport 

#endregion BRANDED REPORTS

#region PVS

<#
.SYNOPSIS
Creates a new PVS user in office 365 with the option to apply a license

.DESCRIPTION
The New-PVSUser cmdlet will create a new user in Office 365. Depending on the parameters, you can select an existing user by UPN and it will add the new user to all of the distribution groups the copied user is a part of (provided you use the license switch).

NOTE: There MUST be an available license for the -Licensed switch to work. Hopefully I'll find a way to automatically purchase a license if one doesn't exist, but it eludes me at the moment. The ultimate way for full usage of this tool is to add a licence before running New-PVSUser

TIP: All the parameters are positional, so you could just type everything out without typing parameter names like -Alias. See 'Get-Help New-PVSUser -Examples' for an example.

.PARAMETER Alias
Alias for the user. PVS uses the first initial last name scheme

.PARAMETER FirstName
First name of the new user

.PARAMETER LastName
Last name of the new user

.PARAMETER Copied
UPN for the user you are copying distribution groups from. This needs to be full e-mail address

.PARAMETER Exchange
Switch that will apply the Exchange Online Plan to the user.

.PARAMETER Lync
Switch that will apply the Lync license.

.EXAMPLE
New-PVSUser -Alias bbango -FirstName Bingo -LastName Bango -Copied khileman@propertyvaluationservices.net -Licensed
This will create a user named Bingo Bango and copy all distribution groups that Kent Hileman is a member off

.EXAMPLE
New-PVSUser -Alias bbango -FirstName Bingo -LastName Bango
This will create a non-licensed user and will not copy distribution groups

.EXAMPLE
New-PVSUser bbango Bingo Bango khileman@propertyvaluationservices.net -Licenced
This is an example of just using the positional parameters while not typing out -Alias and the like

#>
function New-PVSUser {
    [CmdLetBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [string]$Alias,
        [Parameter(Position=1,Mandatory=$true)]
        [string]$FirstName,
        [Parameter(Position=2,Mandatory=$true)]
        [string]$LastName,
        [Parameter(Position=3,Mandatory=$false)]
        [string]$Copied,
        [switch]$Exchange,
        [switch]$Lync
    )
    $Params = @{
    FirstName = $FirstName
    LastName = $LastName
    DisplayName = "$FirstName $LastName"
    UserPrincipalName = "$Alias@propertyvaluationservices.net"
    LicenseAssignment = 'propertyvaluation:EXCHANGESTANDARD'
    UsageLocation = 'US'
    }
    if ($Exchange) {

        # Create new user with exchange license

        New-MsolUser @params
        Write-Verbose "Provisioning mailbox and copying groups from $Copied"
        while($(Get-MsolUser -UserPrincipalName "$Alias@propertyvaluationservices.net").OverallProvisioningStatus -ne "Success") {
            Write-Verbose "Provisioning mailbox... Status: $((Get-MsolUser -UserPrincipalName "$Alias@propertyvaluationservices.net").OverallProvisioningStatus)"
            start-sleep -Seconds 5
        }
        Write-Verbose "Provisioning complete. Copying groups from $Copied"
        start-sleep -Seconds 10

        # Copy distribution group membership

        $copied_dn = (get-mailbox $Copied).distinguishedname
        $new_dn = (get-mailbox $Alias).distinguishedname
        $groups = Get-DistributionGroup -ResultSize unlimited

        foreach ($group in $groups) {
            if ((get-distributiongroupmember $group.identity | select -expand distinguishedname) -contains $copied_dn) {
                Add-DistributionGroupMember -Identity $group.identity -Member $Alias
                Write-Verbose "Added to $($group.identity)"
                }
            }
        }
    else {
        Write-Verbose "Creating user"
        # Create user without license
        New-MsolUser -FirstName $FirstName -LastName $LastName -DisplayName "$FirstName $LastName" -UserPrincipalName "$Alias@propertyvaluationservices.net" -UsageLocation US
        Write-Verbose "User creation complete"

    }
    Write-Verbose "Changing user password to Welcome1"
    Set-MsolUserPassword -UserPrincipalName "$Alias@propertyvaluationservices.net" -NewPassword Welcome1
    Write-Verbose "Password is now Welcome1"

    if($Lync) {
        Set-MsolUserLicense -UserPrincipalName "$Alias@propertyvaluationservices.net" -AddLicenses "propertyvaluation:MCOSTANDARD"
        Write-Verbose "Added lync license"
    }

}


<#
.SYNOPSIS
Automates all processes involved with terminating a PVS user in office 365

.DESCRIPTION
The Term-PVSUser command will work on whoever you enter into the -Alias parameter. It will change their password to P@ssw0rd!, forward e-mail to the supervisor, give the supervisor full access to the mailbox, remove the user from all distribution groups they are a part of, and hide from the GAL

.PARAMETER Alias
Alias for the user. PVS uses the first initial last name scheme.

.PARAMETER Alias
First intial and last name of the user you are terminating

.PARAMETER SupervisorUpn
E-mail address of the supervisor

.EXAMPLE
Term-PVSUser -Alias bbango -SupervisorUpn vparker@propertyvaluationservices.net

#>
function Term-PVSUser {
    [CmdLetBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [string]$Alias,
        [Parameter(Position=1,Mandatory=$true)]
        [string]$SupervisorUpn
    )
    # Confirmation window
    $title = "Terminate User"
    $message = "You have chosen to terminate $Alias and e-mail will be forwarded to $SupervisorUpn. Continue? (Default is 'No')"
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Terminate $Alias"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Cancel's script."
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $message, $options, 1)
    switch($result) {
        0 { 

            # Change o365 password     
            Set-MsolUserPassword -UserPrincipalName "$Alias@propertyvaluationservices.net" -NewPassword P@ssw0rd! -ForceChangePassword $false
            Write-Verbose "Password changed to P@ssw0rd!" 

            # Forward e-mail to supervisor
            Set-Mailbox -Identity $Alias -DeliverToMailboxAndForward $true -ForwardingSmtpAddress $SupervisorUpn
            Write-Verbose "E-mail has been forwarded to $SupervisorUpn" 

            # Give supervisor full access to mailbox
            Add-MailboxPermission -Identity $Alias -User $SupervisorUpn -AccessRights FullAccess -InheritanceType All 
            Write-Verbose "$SupervisorUpn has been given full access to $Alias 's mailbox"

            # Remove from all distribution groups
            $user_dn = (get-mailbox $Alias).distinguishedname
            foreach ($group in Get-DistributionGroup -resultsize unlimited){
                if ((Get-DistributionGroupMember $group.identity | select -expand distinguishedname) -contains $user_dn){
                ForEach-Object { 
                    Remove-DistributionGroupMember -Identity $group.name -Member $Alias -Confirm:$false
                    Write-Verbose "$Alias has been removed from $($group.name)" 
                } 
            }   
        } 
        # Hide from GAL
            Set-Mailbox -Identity $Alias -HiddenFromAddressListsEnabled $true
            Write-Verbose "$Alias has been hidden from the Exchange GAL"
        }

        1 {
            Write-Verbose "Termination Canceled"
        }
    }
}
#endregion PVS