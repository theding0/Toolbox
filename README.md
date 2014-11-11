#Office 365/Exchange Toolbox

General utilities for maintaining Office 365 and Exchange Online


## Connect-O365
Creates PSSession to Office 365 and imports module supplied for global use of MSOnline cmdlets

## Disconnect-O365
Closes all PSSessions connected to *.outlook.com for cleanup

## Get-Distro:
Retrieves a list of distribution group a user is a member of

This command can be used with pipeline input, although doing a mass audit will be time consuming in it's current form.
Output of the command can also be sent through the pipeline for official reporting

    Get-Distro <alias>						- Returns specific user's distribution group membership
	Get-Mailbox | Get-Distro				- Full audit
	

     