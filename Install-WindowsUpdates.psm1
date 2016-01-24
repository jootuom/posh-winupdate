Function Install-WindowsUpdates {
	<#
		.SYNOPSIS
		Install Windows updates
		.DESCRIPTION
		Install Windows updates from the specified source.
		.NOTES
		This interface is documented here:
		https://msdn.microsoft.com/en-us/library/windows/desktop/aa386854(v=vs.85).aspx
		
		Update filter values are listed here:
		https://msdn.microsoft.com/en-us/library/windows/desktop/aa386526(v=vs.85).aspx
	#>
	[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact="High")]
	Param(
		[Parameter()]
		[switch] $AutoReboot,
		
		[Parameter()]
		[ValidateSet("Default","WSUS","Microsoft")]
		[string] $Source = "Default",
		
		[Parameter()]
		[string] $SearchFilter = "IsInstalled=0 and Type='Software'"
	)

	Set-Variable -Name "ErrorActionPreference" -Value "Stop" -Scope "Script"

	$session = New-Object -ComObject Microsoft.Update.Session

	switch($Source) {
		"Default"   { $updatesource = 0 }
		"WSUS"      { $updatesource = 1 }
		"Microsoft" { $updatesource = 2 }
		default     { $updatesource = 0 }
	}

	Function Format-Date {
		Param(
			$delta
		)
		[ref] $ref = $null
		$rem = [math]::divrem($delta.Seconds, 60, $ref)
		"{0} minutes {1} seconds" -f $rem, $delta.Seconds
	}

	# --- Search for updates
	Write-Verbose "Searching for updates..."
	$start = Get-Date

	$searcher = $session.CreateUpdateSearcher()
	$searcher.ServerSelection = $updatesource

	$updates = $searcher.Search($SearchFilter).Updates

	Write-Verbose ("Search took {0}" -f (Format-Date ((Get-Date) - $start)))

	if ($updates.Count -eq 0) {
		Write-Output "No updates to install."
		Exit
	} else {
		Write-Output ("Found {0} update(s):" -f $updates.Count)
		$updates | % { Write-Output $_.Title }
	}

	if (!$PSCmdlet.ShouldProcess("{0} update(s)" -f $updates.Count)) { Exit }

	# --- Download updates
	Write-Verbose "Downloading updates..."
	$start = Get-Date
		
	$downloader = $session.CreateUpdateDownloader()
	$downloader.Updates = $updates
	$downloader.Priority = 3

	$results = $downloader.Download()

	Write-Verbose ("Download took {0}" -f (Format-Date ((Get-Date) - $start)))

	# --- Install updates
	Write-Verbose "Installing updates..."
	$start = Get-Date

	$installer = $session.CreateUpdateInstaller()
	$installer.Updates = $updates
	$installer.AllowSourcePrompts = $false

	$results = $installer.Install()

	Write-Verbose ("Install took {0}" -f (Format-Date ((Get-Date) - $start)))

	if ($results.rebootRequired -and $AutoReboot) { Restart-Computer -Force }
	elseif ($results.rebootRequired) { Write-Host "A reboot is required." }
}