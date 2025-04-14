function Resolve-ImplicitUninstall {
	# Adapted from https://www.alkanesolutions.co.uk/2022/09/20/use-powershell-to-calculate-bit-flags/?doing_wp_cron=1696459713.4683570861816406250000

	param(
		$OfferFlags
	)

	[flags()] 
	enum SMS_ApplicationAssignment_OfferFlags
	{
		None	= 0
		PREDEPLOY = 1
		ONDEMAND = 2
		ENABLEPROCESSTERMINATION = 4
		ALLOWUSERSTOREPAIRAPP = 8
		RELATIVESCHEDULE = 16
		HIGHIMPACTDEPLOYMENT = 32
		IMPLICITUNINSTALL = 64
	}
	
	[SMS_ApplicationAssignment_OfferFlags]$flags = $OfferFlags

	Write-Verbose "Checking if the app has implicit uninstall enabled"
	if ($flags.HasFlag([SMS_ApplicationAssignment_OfferFlags]::IMPLICITUNINSTALL)) {
		Write-Verbose "Implicit uninstall was enabled"
		return $true
	}else{
		Write-Verbose "Implicit uninstall was disabled"
		return $false
	}
}
