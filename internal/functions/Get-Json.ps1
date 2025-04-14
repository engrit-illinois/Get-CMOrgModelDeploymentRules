function Get-Json($Collections) {
	# Give custom names to the relevant properties for friendlier table headings
	$Format = @(
		@{ Name = "Deployment Collection"; Expression = { $_._CollectionNameF } },
		@{ Name = "Membership Rules"; Expression = { $_._MembershipRulesF } },
		@{ Name = "Deployed Content"; Expression = { $_._ContentNamesF } },
		@{ Name = "Purpose"; Expression = { $_._PurposesF } },
		@{ Name = "Action"; Expression = { $_._ActionsF } },
		@{ Name = "Supersedence"; Expression = { $_._SupersedencesF } },
		@{ Name = "ImplicitUninstall"; Expression = { $_._ImplicitUninstallsF } },
		@{ Name = "Overrides"; Expression = { $_._OverridesF } },
		@{ Name = "Comment"; Expression = { $_._CommentsF } }
	)
	
	$Collections | Select-Object $Format | ConvertTo-Json
}
