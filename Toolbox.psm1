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
        $ucount = 0
        $Distros = @()




    } # End BEGIN block


    PROCESS {

        foreach ($u in $User) {

            $Distros += New-object -TypeName psobject -Property @{
                Name = $((Get-Mailbox $u).displayname)
                MemberOf = @() -split ","

            } # End property hash

            $i = 0
            $user_dn = $((Get-Mailbox $u).distinguishedname)

            foreach ($g in $Group) {


                Write-Progress -Activity "Collecting distribution groups for $((Get-Mailbox ($u)).displayname)" -Status "Checking $g" -PercentComplete ($i / $Group.Count * 100)

                if ((Get-DistributionGroupMember $g.identity | select -expand distinguishedname) -contains $user_dn) {

                    $Distros[$ucount].MemberOf += $g

                 } # End identity check

                    $i++

             } # End foreach $Group
             $ucount++

    } # End foreach $User

} #End PROCESS block

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
Function Get-Telnet
{   Param (
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
