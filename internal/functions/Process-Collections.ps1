function Process-Collections
{
	[CmdletBinding()]
	param (
		[Object[]]$Collections,
		[switch]$ISOnly,
		[switch]$Formatted
	)
	Write-Host "Processing collections..."
	$Collections = $Collections | ForEach-Object {
		$Collection = $_
		Write-Host "Processing $($Collection.Name)..."
		
		$Collection = Get-MembershipRules $Collection
		
		# If -ISOnly was specified, ignore collections which do not have any membership rules which reference IS device collections
		if(
			($ISOnly) -and
			(-not (Test-IncludesIsCollections $Collection))
		) {
			# This weakly only filters by Exclude and Include membership rules, because we don't have easily identifiable conventions via Direct or Query-based membership rules
			Write-Verbose "ISOnly flag was declared, but no Include or Exclude membership rules were found on $($Collection.Name) referencing `"UIUC-ENGR-IS*`" collections. Skipping collection."
			return
		}
		if($Formatted){
			Write-Verbose "Getting the Deployments for $($Collection.Name)..."
			$Deployments = Get-CMDeployment -CollectionName $Collection.Name | Sort-Object "ApplicationName"
			
			$Deployments = Process-Deployments $Deployments
			
			# Store deployments in collection object, now that we're done modifying them
			$Collection | Add-Member -NotePropertyName "_Deployments" -NotePropertyValue $Deployments

			$Collection | Add-Member -NotePropertyName "_CollectionNameF" -NotePropertyValue "📂&nbsp;$($Collection.Name) \\ \\🛡️&nbsp;$($Collection.LimitToCollectionName)"
			
			# Join relevant properties from all deployments into a formatted version and store that
			$Collection | Add-Member -NotePropertyName "_ContentNamesF" -NotePropertyValue ($Deployments._ContentNameF -join " \\")
			$Collection | Add-Member -NotePropertyName "_PurposesF" -NotePropertyValue ($Deployments._PurposeF -join " \\")
			
			$Collection | Add-Member -NotePropertyName "_ActionsF" -NotePropertyValue ($Deployments._Content._ActionF -join " \\")
			$Collection | Add-Member -NotePropertyName "_SupersedencesF" -NotePropertyValue ($Deployments._Content._SupersedenceF -join " \\")
			$Collection | Add-Member -NotePropertyName "_ImplicitUninstallsF" -NotePropertyValue ($Deployments._Content._ImplicitUninstallF -join " \\")
			$Collection | Add-Member -NotePropertyName "_OverridesF" -NotePropertyValue ($Deployments._Content._OverrideF -join " \\")
			$Collection | Add-Member -NotePropertyName "_CommentsF" -NotePropertyValue ($Deployments._Content._CommentF -join " \\")
		}
		
		$Collection
	}
	
	$Collections
}
