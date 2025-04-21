function Get-Collections {
	[CmdletBinding()]
	param (
		[switch]$Json,
		[switch]$ISOnly,
		[switch]$IncludeCollectionsWithNoDeployments,
		[switch]$AppNamesOnly,
		[string[]]$CollectionQueries,
		[switch]$Test,
		[string[]]$TestCollectionQueries,
		[switch]$RandomCollections,
		[int]$RandomCollectionsNum,
		[string]$Prefix = $DEFAULT_PREFIX,
		[string]$SiteCode=$DEFAULT_SITE_CODE,
		[string]$Provider=$DEFAULT_PROVIDER,
		[string]$CMPSModulePath
	)
	Write-Host "Getting all collections... (note: this takes a while)"
	if($Test) { $CollectionQueries = $TestCollectionQueries }
	$Collections = $CollectionQueries | ForEach-Object {
		Get-CMDeviceCollection -Name $_
	} | Sort-Object -Property "Name"
	if($RandomCollections) {
		$Collections = $Collections | Get-Random -Count $RandomCollectionsNum
	}
	$Collections
}
