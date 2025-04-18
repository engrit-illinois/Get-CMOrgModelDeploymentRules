function Get-CMOrgModelApps {

	[CmdletBinding()]
	param(
		[switch]$Json,
		[switch]$ISOnly,
        [switch]$AppNamesOnly,
		[string[]]$CollectionQueries = @("UIUC-ENGR-Deploy*","UIUC-ENGR-IS Deploy*","UIUC-ENGR-Uninstall*","UIUC-ENGR-IS Uninstall*"),
		[switch]$Test,
		[string[]]$TestCollectionQueries = @("UIUC-ENGR-Deploy A*","UIUC-ENGR-IS Deploy A*","UIUC-ENGR-Uninstall*","UIUC-ENGR-IS Uninstall*"),
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

    process {
        $Collections = Get-Collections @PSBoundParameters -CollectionQueries $CollectionQueries -TestCollectionQueries $TestCollectionQueries
        $Collections = Process-Collections -Collections $Collections -ISOnly:$ISOnly
        $Apps = $Collections | ForEach-Object {
            (Get-CMDeployment -CollectionName $_.Name).ApplicationName
        }
        # First removing duplicates (e.g. Available + Required)
        $Apps = $Apps | Sort-Object -Unique
        if($AppNamesOnly){
            $Apps = $Apps | ForEach-Object {
                $_ -replace "UIUC-ENGR-",""
            }
            # Sorting again to sort without the UIUC-ENGR- in the app names
            $Apps = $Apps | Sort-Object
        }
    }

    end {
        Write-Host "Done!"
		Set-Location $myPWD
		$Output = $Apps
		if($Json){
			$Output = Get-Json $Apps
		}
		$Output
    }
}