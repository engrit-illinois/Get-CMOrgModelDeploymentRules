function Test-IncludesIsCollections($Collection) {
	if(
		($Collection._IncludeMembershipRules | Where-Object {$_ -like "UIUC-ENGR-IS*"}) -or
		($Collection._ExcludeMembershipRules | Where-Object {$_ -like "UIUC-ENGR-IS*"})
	) {
		return $true
	}
	return $false
}
