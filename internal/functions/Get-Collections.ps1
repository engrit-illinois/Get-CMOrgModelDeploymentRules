function Get-Collections {
	[CmdletBinding()]
	param (
		[Parameter()]
		[switch]
		$Test,
		[Parameter()]
		[switch]
		$RandomCollections,
		[Parameter()]
		[Int]
		$RandomCollectionsNum,
		[string[]]$CollectionQueries,
		[string[]]$TestCollectionQueries,
		[switch]$ISOnly
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
