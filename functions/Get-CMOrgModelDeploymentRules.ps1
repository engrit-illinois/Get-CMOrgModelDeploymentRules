function Get-CMOrgModelDeploymentRules {

	[CmdletBinding(SupportsShouldProcess)]
	param(
		[switch]$Json,
		[switch]$ISOnly,
		[switch]$IncludeCollectionsWithNoDeployments,
		[string[]]$CollectionQueries = @("UIUC-ENGR-Deploy*","UIUC-ENGR-IS Deploy*","UIUC-ENGR-Uninstall*","UIUC-ENGR-IS Uninstall*","UIUC-ENGR-IS Maint Window*"),
		[switch]$Test,
		[string[]]$TestCollectionQueries = @("UIUC-ENGR-Deploy A*","UIUC-ENGR-IS Deploy A*","UIUC-ENGR-Uninstall*","UIUC-ENGR-IS Uninstall*","UIUC-ENGR-IS Maint Window*"),
		[switch]$RandomCollections,
		[int]$RandomCollectionsNum = 10,
		[string]$Prefix = $DEFAULT_PREFIX,
		[string]$SiteCode=$DEFAULT_SITE_CODE,
		[string]$Provider=$DEFAULT_PROVIDER,
		[string]$CMPSModulePath="$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
	)

	begin{
		$myPWD = $PWD.Path
		Connect-ToMECM -SiteCode $SiteCode -Provider $Provider -CMPSModulePath $CMPSModulePath
	}

	process{
		$Collections = Get-Collections @PSBoundParameters -CollectionQueries $CollectionQueries -TestCollectionQueries $TestCollectionQueries
		$Collections = Process-Collections $Collections
	} # End of process block

	end{
		Write-Host "Done!"
		Set-Location $myPWD
		$Output = $Collections
		if($Json){
			$Output = Get-Json $Collections
		}
		$Output
	}
}
Get-CMOrgModelDeploymentRules -Test -RandomCollections -RandomCollectionsNum 3