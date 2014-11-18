<#
	.SYNOPSIS
	    A function to connect to Office 365 for administrative purposes. Imports the module from Office 365 for Msol cmdlets

	.EXAMPLE
	    Connect-o365
#>
function Connect-O365{
    [CmdletBinding()]
    param(
	$cred

    )
    BEGIN
    {

        $cred = Get-Credential

    }
    PROCESS
    {
	    $session365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell/" -Credential $cred -Authentication Basic -AllowRedirection
        Connect-MsolService -Credential $cred
    }
    END
    {
	    Import-Module (Import-PSSession $session365 -AllowClobber) -Global
    }
}
function Disconnect-O365 {
    Get-PSSession | ?{$_.ComputerName -like "*outlook.com"} | Remove-PSSession
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
    [CmdLetBinding()]
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
                MemberOf = @()

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

    Write-Output $Distros 



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
function Get-Telnet {
    param(
    [Parameter(Position=0,Mandatory=$true)]
    [string[]]$Server,
    [Parameter(Position=1,Mandatory=$true)]
    [int]$Port
    )
    BEGIN {
    $Results = @()
    }

    PROCESS {
    foreach ($s in $Server) {
                        
            # Test and enumerate servers with TCP port open
            # Create a Net.Sockets.TcpClient object to use for checking for open TCP ports
            $Socket = New-Object Net.Sockets.TcpClient

            # Suppress error messages
            $ErrorActionPreference = 'SilentlyContinue'

            # Try to connect
            $Socket.Connect($s, $Port)

            # Make error messages visible again
            $ErrorActionPreference = 'Continue'

            # Determine if we are connected
            $obj = New-Object -TypeName PSObject -Property @{
                Name = $s
                Connected = $Socket.Connected
            }    
            
            $Results += $obj
            $Socket.Dispose()
        } # End foreach $s in $server
    } # End PROCESS block

    END {
        Write-Output $Results
    }

} # End Get-Telnet function