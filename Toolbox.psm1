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
       
       if ((!$Password) -and (!$Cred) -and ($Username)) { $Cred = Get-Credential -Message $Username }
       elseif (!$Username) { $cred = Get-Credential }
       else {
       
        $secpasswd = ConvertTo-SecureString -String $Password -AsPlainText -Force
        $Cred = New-Object -TypeName System.Management.Automation.PSCredential($Username,$secpasswd)

       }
    }
    PROCESS
    {
        
        $session365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/﻿" -Credential $cred -Authentication Basic -AllowRedirection 

        Connect-MsolService -Credential $cred
    }
    END
    {
	    Import-Module -ModuleInfo (Import-PSSession -Session $session365 -AllowClobber) -Global
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
                if ((Get-DistributionGroupMember $Group[$i].identity | Select-Object -expand distinguishedname) -contains $user_dn) {
                    $Distros[$ucount].MemberOf += $Group[$i]
                } # End identity check
            } # End group/progress FOR
        } # End user count FOR
    } # End PROCESS block



    END {

    $Distros | Format-Table -AutoSize -Wrap

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
    $Socket = New-Object -TypeName System.Net.Sockets.TcpClient($RemoteHost, $Port)
    if ($Socket)
    {   
        $Stream = $Socket.GetStream()
        $Writer = New-Object -TypeName System.IO.StreamWriter($Stream)
        $Buffer = New-Object -TypeName System.Byte[] 1024 
        $Encoding = New-Object -TypeName System.Text.AsciiEncoding

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
function Publish-EmailAddressReport {
    [CmdletBinding()]
    param (
    [Parameter(ValueFromPipeline=$true)]
    [String[]]$Identity,
    [String]$Path
    )
    BEGIN {
    $MCount = 0
    $css = @"
    table { border:0px; border-collapse:collapse ; width:100%} 
td {vertical-align:top; padding:0px; }


td, table, select, input, textarea{
	font-family:tahoma;
	font-size:11px;
	vertical-align:top;
	line-height:14px;
	color:#363636
}
IMG.logo {
    display: block;
    margin-left: auto;
    margin-right: auto;
    padding-bottom: 25px;
}
tr { page-break-inside:avoid; page-break-after:auto }
th { background-color: DarkOrange;color: white; }

form { margin:0px; padding:0px}
body { margin:0px; padding:0px; background:#FFFFFF}

ul{margin:0px; padding:0px; list-style:none}
ul li { background:url(images/marker.gif) no-repeat 0 8px; padding-left:21px; line-height:18px}
ul li a{text-decoration:none; color:#363636}
ul li a:hover{color:#ED2024}

b { text-transform:uppercase; color:#ED2024; font-weight:100}
b a{text-decoration:none; color:#ED2024}
b a:hover{ text-decoration:underline}
span {color:#ED2024}
span a{text-decoration:none; color:#ED2024}
span a:hover{ text-decoration:underline}


.header {height:164px}
.header1 {width:218px; padding-top:14px}
.header1 a{color:#363636; text-decoration:none}
.header1 a:hover{ color:#ED2024}
.header2 {background:url(images/bg_header.jpg) no-repeat top; height:246px}
.header3 {background:url(images/line_header.jpg) no-repeat top left; height:22px}

.style {width:766px; height:900px}
.style1 {width:277px; padding:18px 27px 0 0}
.style2 {width:382px; padding-top:14px}
.style3 {width:416px; padding:14px 27px 0 0}
.style4 {width:240px; padding-top:18px}
.style5 {width:402px; padding:14px 27px 0 0}
.style6 {width:247px; padding-top:14px}
.style7 {width:253px; padding:14px 27px 0 0}
.style8 {width:394px; padding-top:14px}
.style9 {width:446px; padding:14px 22px 0 0}
.style10 {width:210px; padding-top:14px}
.style11 {width:315px; padding:14px 26px 0 0}
.style12 {width:334px; padding-top:14px}
.style13 {width:254px; padding:14px 25px 0 0}
.style14 {width:393px; padding-top:14px}
.style15 {width:215px; padding:18px 25px 0 0}
.style16 {width:435px; padding-top:14px}
.style17 {width:710px; padding-top:14px}


.bg1{background:url(images/bg_line1.gif) repeat-y; width:1px; height:100%}
.bg2{background:#363636; height:4px}
.bg3{background:url(images/bg_line2.gif) no-repeat bottom}
.bg4 {background:url(images/bg_line3.gif) repeat-x; margin:0 74px 0 53px; width:auto}
.bg5 {background:url(images/bg_line3.gif) repeat-x; margin:0 55px 0 53px; width:auto}


.footer {height:84px; vertical-align: middle; text-align:center}
.footer span a{ text-decoration:none; color:#ED2024}
.footer span a:hover{color:#363636}
.footer a{ text-decoration:none; color:#363636}
.footer a:hover{color:#ED2024}


.form input {
	width:201px;
	height:26px;
	padding:5px 0 0 6px;
	line-height:13px;
	border:solid 1px #363636;
}

.form textarea {
	width:217px;
	height:149px;
	overflow: auto;
	padding:4px 0 0 6px;
	border:solid 1px #363636;
}
"@	
    $image = "$PSScriptRoot\Logo.png"
    $html = "<html><head><title>$((Get-MsolCompanyInformation).DisplayName) - E-mail Address Report</title><style>$css</style></head><body><img class='logo' src=$image alt='Salvus TG'>Report generated: $(Get-Date)<table border='1'><tr><th>Mailbox</th><th>Associated E-mail Addresses</th></tr>"
    if ($Identity -eq "") { $Mailbox = Get-Mailbox }
    elseif ($Identity.Count -gt 1) {
        foreach ($i in $Identity) {
            $Mailbox += (Get-Mailbox $i)
        }
    }
    else { $Mailbox = Get-Mailbox $Identity }
    if (!$Path) { $Path = [Environment]::GetFolderPath("MyDocuments") + "\Powershell Reports" }
    if (!(Test-Path -Path $Path)) { New-Item -Path $Path -ItemType Directory }

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
        $html | Out-File -FilePath $Path\Email_Address_Report.html
        Invoke-Item -Path $Path\Email_Address_Report.html
    } # END END BLOCK
}