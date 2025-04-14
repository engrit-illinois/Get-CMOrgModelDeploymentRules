function Get-MembershipRules($Collection) {
	Write-Verbose "Getting Membership Rules for $($Collection.Name)..."
	$MembershipRulesF = ""
	
	$IncludeMembershipRules = (Get-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection.Name).RuleName | Sort
	if($IncludeMembershipRules) {
		$bullet = "✔️&nbsp;"
		$IncludeMembershipRulesF = $bullet + ($IncludeMembershipRules -join " \\$bullet")
		$MembershipRulesF = $IncludeMembershipRulesF
	}
	
	$ExcludeMembershipRules = (Get-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection.Name).RuleName | Sort
	if($ExcludeMembershipRules) {
		$bullet = "❌&nbsp;"
		$ExcludeMembershipRulesF = $bullet + ($ExcludeMembershipRules -join " \\$bullet")
		$MembershipRulesF = $MembershipRulesF + " \\" + $ExcludeMembershipRulesF
	}
	
	$QueryMembershipRules = (Get-CMDeviceCollectionQueryMembershipRule -CollectionName $Collection.Name).RuleName | Sort
	if($QueryMembershipRules) {
		$bullet = "🔍&nbsp;"
		$QueryMembershipRulesF = $bullet + ($QueryMembershipRules -join " \\$bullet")
		$MembershipRulesF = $MembershipRulesF + " \\" + $QueryMembershipRulesF
	}
	
	$DirectMembershipRules = (Get-CMDeviceCollectionDirectMembershipRule -CollectionName $Collection.Name).RuleName | Sort
	if($DirectMembershipRules) {
		$bullet = "🖥️&nbsp;"
		$DirectMembershipRulesF = $bullet + ($DirectMembershipRules -join " \\$bullet")
		$MembershipRulesF = $MembershipRulesF + " \\" + $DirectMembershipRulesF
	}
	
	$MembershipRulesF = $MembershipRulesF.Trim(" \\")
	
	$Collection | Add-Member -NotePropertyName "_IncludeMembershipRules" -NotePropertyValue $IncludeMembershipRules
	$Collection | Add-Member -NotePropertyName "_IncludeMembershipRulesF" -NotePropertyValue $IncludeMembershipRulesF
	$Collection | Add-Member -NotePropertyName "_ExcludeMembershipRules" -NotePropertyValue $ExcludeMembershipRules
	$Collection | Add-Member -NotePropertyName "_ExcludeMembershipRulesF" -NotePropertyValue $ExcludeMembershipRulesF
	$Collection | Add-Member -NotePropertyName "_QueryMembershipRules" -NotePropertyValue $QueryMembershipRules
	$Collection | Add-Member -NotePropertyName "_QueryMembershipRulesF" -NotePropertyValue $QueryMembershipRulesF
	$Collection | Add-Member -NotePropertyName "_DirectMembershipRules" -NotePropertyValue $DirectMembershipRules
	$Collection | Add-Member -NotePropertyName "_DirectMembershipRulesF" -NotePropertyValue $DirectMembershipRulesF
	$Collection | Add-Member -NotePropertyName "_MembershipRulesF" -NotePropertyValue $MembershipRulesF
	
	$Collection
}
