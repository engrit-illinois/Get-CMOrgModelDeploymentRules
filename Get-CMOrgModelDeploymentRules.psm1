$DEFAULT_PREFIX = "UIUC-ENGR-"
$DEFAULT_SITE_CODE = "MP0"
$DEFAULT_PROVIDER = "sccmcas.ad.uillinois.edu"

# https://learn.microsoft.com/en-us/mem/configmgr/develop/reference/apps/sms_deploymentsummary-server-wmi-class
$FEATURETYPES = @(
	[PSCustomObject]@{ Id = 1; Meaning = "Application"; Icon = "💠" },
	[PSCustomObject]@{ Id = 2; Meaning = "Program"; Icon = "📺" },
	[PSCustomObject]@{ Id = 3; Meaning = "MobileProgram"; Icon = "📱" },
	[PSCustomObject]@{ Id = 4; Meaning = "Script"; Icon = "🧾" },
	[PSCustomObject]@{ Id = 5; Meaning = "SoftwareUpdate"; Icon = "🩹" },
	[PSCustomObject]@{ Id = 6; Meaning = "Baseline"; Icon = "🎸" },
	[PSCustomObject]@{ Id = 7; Meaning = "TaskSequence"; Icon = "📋" },
	[PSCustomObject]@{ Id = 8; Meaning = "ContentDistribution"; Icon = "🚚" },
	[PSCustomObject]@{ Id = 9; Meaning = "DistributionPointGroup"; Icon = "🗄️" },
	[PSCustomObject]@{ Id = 10; Meaning = "DistributionPointHealth"; Icon = "⛑️" },
	[PSCustomObject]@{ Id = 11; Meaning = "ConfigurationPolicy"; Icon = "📖" },
	[PSCustomObject]@{ Id = 28; Meaning = "AbstractConfigurationItem"; Icon = "🏷️" }
)

# https://learn.microsoft.com/en-us/mem/configmgr/develop/reference/apps/sms_appdeploymentassetdetails-server-wmi-class
# Note: "DeploymentIntent" (or "intent") is also sometimes known as "purpose"
$DEPLOYMENTINTENTS = @(
	[PSCustomObject]@{ Id = 1; Meaning = "Required"; Icon = "💡" },
	[PSCustomObject]@{ Id = 2; Meaning = "Available"; Icon = "🔒" },
	[PSCustomObject]@{ Id = 3; Meaning = "Simulate"; Icon = "🤖" }
)

# https://learn.microsoft.com/en-us/mem/configmgr/develop/reference/apps/sms_applicationassignment-server-wmi-class
# Note: "offertype" is also sometimes known as "intent" or "purpose"
$OFFERTYPEIDS = @(
	[PSCustomObject]@{ Id = 0; Meaning = "REQUIRED" },
	[PSCustomObject]@{ Id = 2; Meaning = "AVAILABLE" }
)

# https://learn.microsoft.com/en-us/mem/configmgr/develop/reference/compliance/sms_ciassignmentbaseclass-server-wmi-class
# Note: "DesiredConfigType" (or "configtype") is also sometimes known as "action"
$DESIREDCONFIGTYPES = @(
	[PSCustomObject]@{ Id = 1; Meaning = "REQUIRED"; FriendlyMeaning = "Install"; Icon = "💾" },
	[PSCustomObject]@{ Id = 2; Meaning = "NOT_ALLOWED"; FriendlyMeaning = "Uninstall"; Icon = "🗑" }
)

function Connect-ToMECM {
	[CmdletBinding(SupportsShouldProcess)]
	param(
		[string]$SiteCode,
		[string]$Provider,
		[string]$CMPSModulePath
	)

	Write-Host "Preparing connection to MECM..."
	$initParams = @{}
	if($null -eq (Get-Module ConfigurationManager)) {
		# The ConfigurationManager Powershell module switched filepaths at some point around CB 18##
		# So you may need to modify this to match your local environment
		Import-Module $CMPSModulePath @initParams -Scope Global
	}
	if(($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue))) {
		New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $Provider @initParams
	}
	Set-Location "$($SiteCode):\" @initParams
	Write-Host "Done prepping connection to MECM."
}

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

function Test-IncludesIsCollections($Collection) {
	if(
		($Collection._IncludeMembershipRules | Where-Object {$_ -like "UIUC-ENGR-IS*"}) -or
		($Collection._ExcludeMembershipRules | Where-Object {$_ -like "UIUC-ENGR-IS*"})
	) {
		return $true
	}
	return $false
}

function Get-Collections {
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

function Get-AppDeploymentContent($Deployment) {
	# Get app and full app deployment objects to gather more detailed info
	$Content = Get-CMApplication -Fast -Name $Deployment.ApplicationName
	$ContentDeployment = Get-CMApplicationDeployment -DeploymentId $Deployment.DeploymentId
	
	# Store formatted version of Comment
	$Comment = $Content.LocalizedDescription
	$CommentF = "💭&nbsp;None"
	if($Comment) { $CommentF = "💬&nbsp;$($Comment)" }
	$Content | Add-Member -NotePropertyName "_CommentF" -NotePropertyValue $CommentF
	
	# Store formatted version of Action
	$Action = $DESIREDCONFIGTYPES | Where { $_.Id -eq $ContentDeployment.DesiredConfigType }
	$Content | Add-Member -NotePropertyName "_ActionF" -NotePropertyValue ($Action.Icon + "&nbsp;" + $Action.FriendlyMeaning)
	
	# Store formatted version of Supersedence
	$Supersedence = $ContentDeployment.UpdateSupersedence
	$SupersedenceF = "👑&nbsp;❌&nbsp;Disabled"
	if($Supersedence) { $SupersedenceF = "👑&nbsp;✔️&nbsp;Enabled" }
	$Content | Add-Member -NotePropertyName "_SupersedenceF" -NotePropertyValue $SupersedenceF
	
	# Store formatted version of ImplicitUninstall
	$ImplicitUninstall = Resolve-ImplicitUninstall -OfferFlags $ContentDeployment.OfferFlags
	$ImplicitUninstallF = "🚯&nbsp;Not Implicit"
	if($ImplicitUninstall) { $ImplicitUninstallF = "🚮&nbsp;Implicit" }
	$Content | Add-Member -NotePropertyName "_ImplicitUninstallF" -NotePropertyValue $ImplicitUninstallF

	# Store formatted version of OverrideServiceWindows
	$OverrideServiceWindows = $ContentDeployment.OverrideServiceWindows
	$OverrideServiceWindowsF = "▶️&nbsp;🕑"
	if($OverrideServiceWindows) { $OverrideServiceWindowsF = "▶️&nbsp;⚠️" }
	$Content | Add-Member -NotePropertyName "_OverrideServiceWindowsF" -NotePropertyValue $OverrideServiceWindowsF
	
	$RebootOutsideOfServiceWindows = $ContentDeployment.RebootOutsideOfServiceWindows
	$RebootOutsideOfServiceWindowsF = "🔃&nbsp;🕑"
	if($RebootOutsideOfServiceWindows) { $RebootOutsideOfServiceWindowsF = "🔃&nbsp;🚨" }
	$Content | Add-Member -NotePropertyName "_RebootOutsideOfServiceWindowsF" -NotePropertyValue $RebootOutsideOfServiceWindowsF
	
	$Content | Add-Member -NotePropertyName "_OverrideF" -NotePropertyValue ($OverrideServiceWindowsF + "&nbsp;,&nbsp;" + $RebootOutsideOfServiceWindowsF)
	
	$Content
}

function Get-TaskSequenceDeploymentContent($Deployment) {
	# Get TS and full TS deployment objects to gather more detailed info
	$Content = Get-CMTaskSequence -Fast -Name $Deployment.ApplicationName
	#$ContentDeployment = Get-CMTaskSequenceDeployment -DeploymentId $Deployment.DeploymentId
	# Everything we care about we already have in $deployment
	
	# Store formatted version of Comment
	$Comment = $Content.LocalizedTaskSequenceDescription
	$CommentF = "💭&nbsp;None"
	if($Comment) { $CommentF = "💬&nbsp;$($Comment)" }
	$Content | Add-Member -NotePropertyName "_CommentF" -NotePropertyValue $CommentF
	
	$Content
}

function Process-Deployments($Deployments) {
	
	if($null -eq $Deployments){
		Write-Verbose "Found no deployments."
		
		if($IncludeCollectionsWithNoDeployments) {
			Write-Verbose "-IncludeCollectionsWithNoDeployments was specified. Creating dummy deployment for output."
			# Create a dummy deployment with dummy content so it ends up in the final array
			# Populate all properties with dummy info so that rows with multiple deployments will have the same number of items in every cell
			$Deployments = @{
				_ContentNameF = "❓&nbsp;No deployments!"
				_PurposeF = "N/A"
				_Content = [PSCustomObject]@{}
			}
		}
		else {
			Write-Verbose "Skipping $($Collection.Name) because it has no Deployments"
		}
	}
	else{
		$DeploymentsCount = $Deployments.count
		Write-Verbose "Found $($Deployments.Count) Deployments."
		
		Write-Verbose "Processing Deployments..."
		$Deployments = $Deployments | ForEach-Object {
			$Deployment = $_
			$ContentName = $Deployment.ApplicationName
			Write-Verbose "Processing `"$contentName`"..."
			
			# Store formatted version of deployment/content name
			# As far as I can tell, this should be identical to whatever name is stored in the actual deployed content, and
			# it's just easier to reference from the deployment data
			$ContentNameIcon = $FEATURETYPES | Where { $_.Id -eq $Deployment.FeatureType } | Select -ExpandProperty "Icon"
			if(-not $ContentNameIcon) { $ContentNameIcon = "❓" }
			$Deployment | Add-Member -NotePropertyName "_ContentNameF" -NotePropertyValue "$($ContentNameIcon)&nbsp;$($ContentName)"
			
			# Store formatted version of the deployment purpose
			$Purpose = $DEPLOYMENTINTENTS | Where { $_.Id -eq $Deployment.DeploymentIntent }
			$Deployment | Add-Member -NotePropertyName "_PurposeF" -NotePropertyValue ($Purpose.Icon + "&nbsp;" + $Purpose.Meaning)
			
			switch($Deployment.FeatureType) {
				# For applications get some additional info
				1 {
					Write-Verbose "Content is an application."
					$Content = Get-AppDeploymentContent $Deployment
				}
				# For task sequences
				7 {
					Write-Verbose "Content is a task sequence."
					$Content = Get-TaskSequenceDeploymentContent $Deployment
				}
				# For any other content type
				Default {
					# Make a dummy content object so we have somewhere to attach the misc. info.
					$Content = [PSCustomObject]@{
						Name = $ContentName
					}
				}
			}
			
			$Deployment | Add-Member -NotePropertyName "_Content" -NotePropertyValue $Content
			
			$Deployment
		}
	}
	
	# Fill in any missing formatted info so that rows with multiple deployments will have the same number of items in every cell
	$Deployments = $Deployments | ForEach-Object {
		$Deployment = $_
		$Content = $Deployment._Content
		if(-not $Content._ActionF) { $Content | Add-Member -NotePropertyName "_ActionF" -NotePropertyValue "N/A" }
		if(-not $Content._SupersedenceF) { $Content | Add-Member -NotePropertyName "_SupersedenceF" -NotePropertyValue "N/A" }
		if(-not $Content._ImplicitUninstallF) { $Content | Add-Member -NotePropertyName "_ImplicitUninstallF" -NotePropertyValue "N/A" }
		if(-not $Content._OverrideF) { $Content | Add-Member -NotePropertyName "_OverrideF" -NotePropertyValue "N/A" }
		if(-not $Content._CommentF) { $Content | Add-Member -NotePropertyName "_CommentF" -NotePropertyValue "N/A" }
		
		$Deployment._Content = $Content
		$Deployment
	}
	
	$Deployments
}

function Process-Collections($Collections) {
	Write-Host "Processing collections..."
	$Collections = $Collections | ForEach-Object {
		$Collection = $_
		Write-Host "Processing $($Collection.Name)..."
		
		$Collection | Add-Member -NotePropertyName "_CollectionNameF" -NotePropertyValue "📂&nbsp;$($Collection.Name) \\ \\🛡️&nbsp;$($Collection.LimitToCollectionName)"
		
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
		
		Write-Verbose "Getting the Deployments for $($Collection.Name)..."
		$Deployments = Get-CMDeployment -CollectionName $Collection.Name | Sort "ApplicationName"
		
		$Deployments = Process-Deployments $Deployments
		
		# Store deployments in collection object, now that we're done modying them
		$Collection | Add-Member -NotePropertyName "_Deployments" -NotePropertyValue $Deployments
		
		# Join relevant properties from all deployments into a formatted version and store that
		$Collection | Add-Member -NotePropertyName "_ContentNamesF" -NotePropertyValue ($Deployments._ContentNameF -join " \\")
		$Collection | Add-Member -NotePropertyName "_PurposesF" -NotePropertyValue ($Deployments._PurposeF -join " \\")
		
		$Collection | Add-Member -NotePropertyName "_ActionsF" -NotePropertyValue ($Deployments._Content._ActionF -join " \\")
		$Collection | Add-Member -NotePropertyName "_SupersedencesF" -NotePropertyValue ($Deployments._Content._SupersedenceF -join " \\")
		$Collection | Add-Member -NotePropertyName "_ImplicitUninstallsF" -NotePropertyValue ($Deployments._Content._ImplicitUninstallF -join " \\")
		$Collection | Add-Member -NotePropertyName "_OverridesF" -NotePropertyValue ($Deployments._Content._OverrideF -join " \\")
		$Collection | Add-Member -NotePropertyName "_CommentsF" -NotePropertyValue ($Deployments._Content._CommentF -join " \\")
		
		$Collection
	}
	
	$Collections
}

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
		$Collections = Get-Collections
		$Collections = Process-Collections $Collections
	} # End of process block

	end{
		Write-Host "Done!"
		$Output = $Collections
		if($Json){
			$Output = Get-Json $Collections
		}
		$Output
	}
}

Export-ModuleMember Get-CMOrgModelDeploymentRules