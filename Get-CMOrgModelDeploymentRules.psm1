$DEFAULT_PREFIX = "UIUC-ENGR-"
$DEFAULT_SITE_CODE = "MP0"
$DEFAULT_PROVIDER = "sccmcas.ad.uillinois.edu"
$DEFAULT_COMMENT = "None"

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

function Translate-FeatureType($featureType) {
	# https://learn.microsoft.com/en-us/mem/configmgr/develop/reference/apps/sms_deploymentsummary-server-wmi-class
	$featureTypes = @(
		[PSCustomObject]@{ Id = 1; Meaning = "Application" },
		[PSCustomObject]@{ Id = 2; Meaning = "Program" },
		[PSCustomObject]@{ Id = 3; Meaning = "MobileProgram" },
		[PSCustomObject]@{ Id = 4; Meaning = "Script" },
		[PSCustomObject]@{ Id = 5; Meaning = "SoftwareUpdate" },
		[PSCustomObject]@{ Id = 6; Meaning = "Baseline" },
		[PSCustomObject]@{ Id = 7; Meaning = "TaskSequence" },
		[PSCustomObject]@{ Id = 8; Meaning = "ContentDistribution" },
		[PSCustomObject]@{ Id = 9; Meaning = "DistributionPointGroup" },
		[PSCustomObject]@{ Id = 10; Meaning = "DistributionPointHealth" },
		[PSCustomObject]@{ Id = 11; Meaning = "ConfigurationPolicy" },
		[PSCustomObject]@{ Id = 28; Meaning = "AbstractConfigurationItem" }
	)
	
	$result = $featureTypes | Where { $_.Id -eq $featureType } | Select -ExpandProperty "Meaning"
	if(-not $result) { $result = "UNKNOWN" }
	
	$result
}

function Translate-DeploymentIntent($deploymentIntent) {
	# https://learn.microsoft.com/en-us/mem/configmgr/develop/reference/apps/sms_appdeploymentassetdetails-server-wmi-class
	# Note: "DeploymentIntent" (or "intent") is also sometimes known as "purpose"
	$deploymentIntents = @(
		[PSCustomObject]@{ Id = 1; Meaning = "Required" },
		[PSCustomObject]@{ Id = 2; Meaning = "Available" },
		[PSCustomObject]@{ Id = 3; Meaning = "Simulate" }
	)
	
	$result = $deploymentIntents | Where { $_.Id -eq $deploymentIntent } | Select -ExpandProperty "Meaning"
	if(-not $result) { $result = "UNKNOWN" }
	
	$result
}

# OfferTypeID is the application-only version of the DeploymentIntent property, which is available on the more generic deployment object
<#
function Translate-OfferTypeID($offerTypeId) {
	# https://learn.microsoft.com/en-us/mem/configmgr/develop/reference/apps/sms_applicationassignment-server-wmi-class
	# Note: "offertype" is also sometimes known as "intent" or "purpose"
	$offerTypeIds = @(
		[PSCustomObject]@{ Id = 0; Meaning = "REQUIRED" },
		[PSCustomObject]@{ Id = 2; Meaning = "AVAILABLE" }
	)
	
	$result = $offerTypeIds | Where { $_.Id -eq $offerTypeId } | Select -ExpandProperty "Meaning"
	if(-not $result) { $result = "UNKNOWN" }
	
	$result
}
#>

function Translate-DesiredConfigType($desiredConfigType) {
	# https://learn.microsoft.com/en-us/mem/configmgr/develop/reference/compliance/sms_ciassignmentbaseclass-server-wmi-class
	# Note: "DesiredConfigType" (or "configtype") is also sometimes known as "action"
	$desiredConfigTypes = @(
		[PSCustomObject]@{ Id = 1; Meaning = "REQUIRED" }, # a.k.a. "INSTALL"
		[PSCustomObject]@{ Id = 2; Meaning = "NOT_ALLOWED" } # a.k.a. "UNINSTALL"
	)
	
	$result = $desiredConfigTypes | Where { $_.Id -eq $desiredConfigType } | Select -ExpandProperty "Meaning"
	if(-not $result) { $result = "UNKNOWN" }
	
	$result
}

function Translate-Action($desiredConfigType) {
	# Just translating DesiredConfigType again to make it more recognizable
	$actions = @(
		[PSCustomObject]@{ Id = 1; Meaning = "Install" },
		[PSCustomObject]@{ Id = "REQUIRED"; Meaning = "Install" },
		[PSCustomObject]@{ Id = 2; Meaning = "Uninstall" },
		[PSCustomObject]@{ Id = "NOT_ALLOWED"; Meaning = "Uninstall" }
	)
	
	$result = $actions | Where { $_.Id -eq $desiredConfigType } | Select -ExpandProperty "Meaning"
	if(-not $result) { $result = "UNKNOWN" }
	
	$result
}

function Build-ArrayObject {
	
	[CmdletBinding()]
	param(
		$Collection,
		$IncludeMembershipRules,
		$ExcludeMembershipRules,
		$QueryMembershipRules,
		$DirectMembershipRules,
		$Deployments
	)
	
	$CollectionName = $Collection.Name
	Write-Verbose "Building the custom array for $CollectionName..."
	
	$LimitingCollection = $Collection.LimitToCollectionName
	$CollectionNameFormatted = $CollectionName
	if($null -ne $LimitingCollection) {
		$CollectionNameFormatted = "📂$($CollectionName)" + " \\ \\️🛡️$($LimitingCollection)"
	}
	
	# Certain fields may be arrays of values, if the collection has multiple deployments.
	# Save a formatted version of those which are more readable after being JSON-ified:
	Write-Verbose "IncludeMembershipRules = $IncludeMembershipRules"
	if($null -ne $IncludeMembershipRules) {
		$bullet = "✔️&nbsp;"
		$IncludeMembershipRulesFormatted = $bullet + ($IncludeMembershipRules -join " \\$bullet")
	}
	Write-Verbose "IncludeMembershipRulesFormatted = $IncludeMembershipRulesFormatted"
	
	Write-Verbose "ExcludeMembershipRules = $ExcludeMembershipRules"
	if($null -ne $ExcludeMembershipRules) {
		$bullet = "❌&nbsp;"
		$ExcludeMembershipRulesFormatted = $bullet + ($ExcludeMembershipRules -join " \\$bullet")
	}
	Write-Verbose "ExcludeMembershipRulesFormatted = $ExcludeMembershipRulesFormatted"
	
	$IncludeAndExcludeMembershipRulesFormatted = $IncludeMembershipRulesFormatted + " \\" + $ExcludeMembershipRulesFormatted
	# In case one or the other is null:
	$IncludeAndExcludeMembershipRulesFormatted = $IncludeAndExcludeMembershipRulesFormatted.Trim(" \\")
	Write-Verbose "IncludeAndExcludeMembershipRulesFormatted = $IncludeAndExcludeMembershipRulesFormatted"
	
	Write-Verbose "QueryMembershipRules = $QueryMembershipRules"
	Write-Verbose "DirectMembershipRules = $DirectMembershipRules"
	
	$Contents = $Deployments._Content | Sort-Object "ApplicationName"
	
	$ContentNames = $Contents._NameWithType
	Write-Verbose "ContentNames = $ContentNames"
	if($null -ne $ContentNames) {
		$contentNamesBulleted = $ContentNames | ForEach-Object {
			$name = $_
			$name = $name.Replace("Application;;","💠&nbsp;")
			$name = $name.Replace("TaskSequence;;","📋&nbsp;")
			$name = $name.Replace("None;;","❓&nbsp;")
			# For unhandled content types
			$name = $name -replace "^.*;;","🤷&nbsp;"
			$name
		}
		$ContentNamesFormatted = $contentNamesBulleted -join " \\"
	}
	Write-Verbose "ContentNamesFormatted = $ContentNamesFormatted"
	
	$Comments = $Contents._Comment
	Write-Verbose "Comments = $Comments"
	if($null -ne $Comments) {
		$commentsBulleted = $Comments | ForEach-Object {
			$comment = $_
			if($comment -eq "N/A") {
				# Leave as is
			}
			elseif($comment -eq $DEFAULT_COMMENT) {
				$comment = "💭&nbsp;$($comment)"
			}
			else {
				$comment = "💬&nbsp;$($comment)"
			}
			$comment
		}
		$CommentsFormatted = $commentsBulleted -join " \\"
	}
	Write-Verbose "CommentsFormatted = $CommentsFormatted"

	$Purposes = $Contents._Purpose
	Write-Verbose "Purposes = $Purposes"
	if($null -ne $Purposes) {
		$PurposesFormatted = $Purposes -join " \\"
		$PurposesFormatted = $PurposesFormatted.Replace("Available","💡&nbsp;Available")
		$PurposesFormatted = $PurposesFormatted.Replace("Required","🔒&nbsp;Required")
	}
	Write-Verbose "PurposesFormatted = $PurposesFormatted"
		
	$Actions = $Contents._Action
	Write-Verbose "Actions = $Actions"
	if($null -ne $Actions) {
		$ActionsFormatted = $Actions -join " \\"
		$ActionsFormatted = $ActionsFormatted.Replace("Install","💾&nbsp;Install")
		$ActionsFormatted = $ActionsFormatted.Replace("Uninstall","🗑&nbsp;Uninstall")
	}
	Write-Verbose "ActionsFormatted = $ActionsFormatted"
	
	$ImplicitUninstalls = $Contents._ImplicitUninstall
	Write-Verbose "ImplicitUninstalls = $ImplicitUninstalls"
	if($null -ne $ImplicitUninstalls) {
		$ImplicitUninstallsFormatted = $ImplicitUninstalls -join " \\"
		$ImplicitUninstallsFormatted = $ImplicitUninstallsFormatted.Replace("True","🚮&nbsp;Implicit")
		$ImplicitUninstallsFormatted = $ImplicitUninstallsFormatted.Replace("False","🚯&nbsp;Not Implicit")
	}
	Write-Verbose "ImplicitUninstallsFormatted = $ImplicitUninstallsFormatted"
	
	$Supersedences = $Contents._Supersedence
	Write-Verbose "Supersedences = $Supersedences"
	if($null -ne $Supersedences) {
		$SupersedencesFormatted = $Supersedences -join " \\"
		$SupersedencesFormatted = $SupersedencesFormatted.Replace("True","👑&nbsp;✔️&nbsp;Enabled")
		$SupersedencesFormatted = $SupersedencesFormatted.Replace("False","👑&nbsp;❌&nbsp;Disabled")
	}
	Write-Verbose "SupersedencesFormatted = $SupersedencesFormatted"
	
	$OverrideServiceWindowses = $Contents._OverrideServiceWindows
	Write-Verbose "OverrideServiceWindowses = $OverrideServiceWindowses"
	
	$RebootOutsideOfServiceWindowses = $Contents._RebootOutsideOfServiceWindows
	Write-Verbose "RebootOutsideOfServiceWindowses = $RebootOutsideOfServiceWindowses"
	
	$StartTimes = $Contents._StartTime
	Write-Verbose "StartTimes = $StartTimes"
	
	## Not sure how to handle Direct, Exclude, or Query rules yet. Will deal with them later
	[PSCustomObject]@{
		CollectionName					= $CollectionName
		LimitingCollection				= $LimitingCollection
		CollectionNameF					= $CollectionNameFormatted
		IncludeMembershipRules			= $IncludeMembershipRules
		IncludeMembershipRulesF			= $IncludeMembershipRulesFormatted
		ExcludeMembershipRules			= $ExcludeMembershipRules
		ExcludeMembershipRulesF			= $ExcludeMembershipRulesFormatted
		MembershipRulesF				= $IncludeAndExcludeMembershipRulesFormatted
		QueryMembershipRules			= $QueryMembershipRules
		DirectMembershipRules			= $DirectMembershipRules
		ContentNames					= $ContentNames
		ContentNamesF					= $ContentNamesFormatted
		Comments						= $Comments
		CommentsF						= $CommentsFormatted
		Purposes						= $Purposes
		PurposesF						= $PurposesFormatted
		Actions							= $Actions
		ActionsF						= $ActionsFormatted
		ImplicitUninstalls				= $ImplicitUninstalls
		ImplicitUninstallsF				= $ImplicitUninstallsFormatted
		Supersedences					= $Supersedences
		SupersedencesF					= $SupersedencesFormatted
		OverrideServiceWindowses		= $OverrideServiceWindowses
		RebootOutsideOfServiceWindowses = $RebootOutsideOfServiceWindowses
		StartTimes						= $StartTimes
	}
}

function Get-CMOrgModelDeploymentRules{

	[CmdletBinding(SupportsShouldProcess)]
	param(
		[switch]$Json,
		[switch]$ISOnly,
		[switch]$IncludeCollectionsWithNoDeployments,
		[string[]]$CollectionQueries = @("UIUC-ENGR-Deploy*","UIUC-ENGR-IS Deploy*","UIUC-ENGR-IS Uninstall*","UIUC-ENGR-IS Uninstall*"),
		[switch]$Test,
		[string[]]$TestCollectionQueries = @("UIUC-ENGR-IS Deploy A*","UIUC-ENGR-IS Uninstall A*"),
		[switch]$RandomCollections,
		[int]$RandomCollectionsNum = 10,
		[switch]$NoProgressBar,
		[string]$Prefix = $DEFAULT_PREFIX,
		[string]$SiteCode=$DEFAULT_SITE_CODE,
		[string]$Provider=$DEFAULT_PROVIDER,
		[string]$CMPSModulePath="$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
	)
	begin{
		$myPWD = $PWD.Path

		try {
			Connect-ToMECM -SiteCode $SiteCode -Provider $Provider -CMPSModulePath $CMPSModulePath
			
			## Declare the empty arraylist
			$output = New-Object System.Collections.ArrayList

			$myPSStyle = $PSStyle.Progress.View
			$PSStyle.Progress.View = 'Classic'

			if($Test){
				$TestType = $Test.GetType()
				Write-Verbose "Test type is $TestType"
			}
			
		} catch {
			Write-Host $_
		}
	}

	process{
		try{
			Write-Host "Getting all collections... (note: this takes a while)"
			
			if($Test) { $CollectionQueries = $TestCollectionQueries }
			$DeployCollections = $CollectionQueries | ForEach-Object {
				Get-CMDeviceCollection -Name $_
			} | Sort-Object -Property "Name"
			
			if($RandomCollections) {
				$DeployCollections = $DeployCollections | Get-Random -Count $RandomCollectionsNum
			}

			$TotalCollections = $DeployCollections.Count									# Using for progress bar
			$PercentComplete = 0															# Using for progress bar
			$CollectionCount = 0															# Using for progress bar
			if($NoProgressBar){
				Write-Host "Processing Membership Rules..."
			}
			foreach($Collection in $DeployCollections){
				if($NoProgressBar){
					Write-Host "Processing $($Collection.Name)..."
				}else{
					Write-Progress -Activity "Processing Membership Rules..." -Status "$($PercentComplete)% Processing $($Collection.Name)..." -PercentComplete $PercentComplete
					$CollectionCount++														  # Using for progress bar
					$PercentComplete = [int](($CollectionCount / $TotalCollections) * 100)	  # Using for progress bar
				}

				Write-Verbose "Initializing the Membership Rules array as empty"
				$MembershipRules = $null
				Write-Verbose "Getting Direct Membership Rules for $($Collection.Name)..."
				$DirectMembershipRules = (Get-CMDeviceCollectionDirectMembershipRule -CollectionName $Collection.Name).RuleName
				Write-Verbose "Getting Exclude Membership Rules for $($Collection.Name)..."
				$ExcludeMembershipRules = (Get-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection.Name).RuleName
				Write-Verbose "Getting Include Membership Rules for $($Collection.Name)..."
				$IncludeMembershipRules = (Get-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection.Name).RuleName
				Write-Verbose "Getting Query Membership Rules for $($Collection.Name)..."
				$QueryMembershipRules = (Get-CMDeviceCollectionQueryMembershipRule -CollectionName $Collection.Name).RuleName
				
				# If -ISOnly was specified, ignore collections which do not have any membership rules which reference IS device collections
				if(
					($ISOnly) -and
					(-not ($ExcludeMembershipRules | Where-Object {$_ -like "UIUC-ENGR-IS*"}) ) -and
					(-not ($IncludeMembershipRules | Where-Object {$_ -like "UIUC-ENGR-IS*"}) )
				) {
					# This weakly only filters by Exclude and Include membership rules, because we don't have easily identifiable conventions via Direct or Query-based membership rules
					Write-Verbose "ISOnly flag was declared, but no Include or Exclude membership rules were found on $($Collection.Name) referencing `"UIUC-ENGR-IS*`" collections."
				} else {
					
					Write-Verbose "Getting the Deployments for $($Collection.Name)..."
					$AllDeployments = Get-CMDeployment -CollectionName $Collection.Name

					if($null -eq $AllDeployments){
						if($IncludeCollectionsWithNoDeployments) {
							Write-Verbose "Collection has no deployments. Creating dummy deployment for output."
							# Create a dummy deployment with dummy content so it ends up in the final array
							# Populate all properties with dummy info so that rows with multiple deployments will have the same number of items in every cell
							$FilteredDeployments = @{
								ApplicationName = "No deployments!"
								_Content = @{
									_FeatureType = "None"
									_NameWithType = "None;;No deployments!"
									_Comment = "N/A"
									_Purpose = "N/A"
									_Action = "N/A"
									_ImplicitUninstall = "N/A"
									_Supersedence = "N/A"
									_OverrideServiceWindows = "N/A"
									_RebootOutsideOfServiceWindows = "N/A"
									_StartTime = "N/A"
								}
							}
						}
						else {
							Write-Verbose "Skipping $($Collection.Name) because it has no Deployments"
						}
					}else{
						Write-Verbose "Getting Deployment info..."
						$FilteredDeployments = $AllDeployments | ForEach-Object {
							$deployment = $_
							$deploymentId = $deployment.DeploymentId
							$contentName = $deployment.ApplicationName
							$purpose = $deployment.DeploymentIntent
							$purposeTranslated = Translate-DeploymentIntent $purpose
							$featureType = Translate-FeatureType $deployment.FeatureType
							$nameWithType = "$($featureType);;$contentName"
							
							Write-Verbose "    Getting content info for `"$contentName`"..."
							
							if($featureType -eq "Application") {
								Write-Verbose "Content is an application."
								
								# Get app info
								$content = Get-CMApplication -Fast -Name $contentName
								$comment = $content.LocalizedDescription
								if(-not $comment) { $comment = $DEFAULT_COMMENT }
								
								# Get app deployment info
								$contentDeployment = Get-CMApplicationDeployment -DeploymentId $deploymentId
								$action = $contentDeployment.DesiredConfigType
								$actionTranslated = Translate-Action $action
								$supersedence = $contentDeployment.UpdateSupersedence
								$implicitUninstall = Resolve-ImplicitUninstall -OfferFlags $contentDeployment.OfferFlags
								$overrideServiceWindows = $contentDeployment.OverrideServiceWindows
								$rebootOutsideOfServiceWindows = $contentDeployment.RebootOutsideOfServiceWindows
								$startTime = $contentDeployment.StartTime
								# Lazy hack to account for us not being on GMT
								if($startTime) { $startTimeShifted = $startTime.AddHours(5) }
							}
							elseif($featureType -eq "TaskSequence") {
								Write-Verbose "Content is a task sequence."
								
								# Get TS info
								$content = Get-CMTaskSequence -Fast -Name $contentName
								$comment = $content.LocalizedTaskSequenceDescription
								if(-not $comment) { $comment = $DEFAULT_COMMENT }
								
								# Populate all properties with dummy info so that rows with multiple deployments will have the same number of items in every cell
								$action = "N/A"
								$actionTranslated = $action
								$implicitUninstall = "N/A"
								$supersedence = "N/A"
								$overrideServiceWindows = "N/A"
								$rebootOutsideOfServiceWindows = "N/A"
								$startTime = "N/A"
								
								# Get TS deployment info
								#$contentDeployment = Get-CMTaskSequenceDeployment -DeploymentId $deploymentId
								# Everything we care about we already have in $deployment
							}
							else {
								# Make a dummy content object so we have somewhere to attach the misc. info.
								$content = [PSCustomObject]@{
									Name = $contentName
								}
								
								$action = "N/A"
								$actionTranslated = $action
								$implicitUninstall = "N/A"
								$supersedence = "N/A"
								$overrideServiceWindows = "N/A"
								$rebootOutsideOfServiceWindows = "N/A"
								$startTime = "N/A"
							}
							
							if($content) {
								if($content -eq "IGNORED") {
									Write-Verbose "Skipping deployment because it's not an application nor a task sequence."
								}
								else {
									$content | Add-Member -NotePropertyName "_FeatureType" -NotePropertyValue $featureType
									$content | Add-Member -NotePropertyName "_NameWithType" -NotePropertyValue $nameWithType
									$content | Add-Member -NotePropertyName "_Comment" -NotePropertyValue $comment
									$content | Add-Member -NotePropertyName "_Purpose" -NotePropertyValue $purposeTranslated
									$content | Add-Member -NotePropertyName "_Action" -NotePropertyValue $actionTranslated
									$content | Add-Member -NotePropertyName "_ImplicitUninstall" -NotePropertyValue $implicitUninstall
									$content | Add-Member -NotePropertyName "_Supersedence" -NotePropertyValue $supersedence
									$content | Add-Member -NotePropertyName "_OverrideServiceWindows" -NotePropertyValue $overrideServiceWindows
									$content | Add-Member -NotePropertyName "_RebootOutsideOfServiceWindows" -NotePropertyValue $rebootOutsideOfServiceWindows
									$content | Add-Member -NotePropertyName "_StartTime" -NotePropertyValue $startTimeShifted
									
									Write-Verbose "Name = $contentName"
									Write-Verbose "NameWithType = $nameWithType"
									Write-Verbose "Comment = $comment"
									Write-Verbose "Purpose = $purpose"
									Write-Verbose "PurposeTranslated = $purposeTranslated"
									Write-Verbose "Action = $action"
									Write-Verbose "ActionTranslated = $actionTranslated"
									Write-Verbose "ImplicitUninstall = $implicitUninstall"
									Write-Verbose "Supersedence = $supersedence"
									Write-Verbose "OverrideServiceWindows = $overrideServiceWindows"
									Write-Verbose "RebootOutsideOfServiceWindows = $rebootOutsideOfServiceWindows"
									Write-Verbose "StartTime = $startTime"
									Write-Verbose "StartTimeShifted = $startTimeShifted"
									
									$deployment | Add-Member -NotePropertyName "_Content" -NotePropertyValue $content
									$deployment
								}
							}
							else {
								Write-Verbose "Skipping deployment because its content could not be retrieved."
							}
						}
					}
					
					# Using all of the gathered info, build an object suitable to represent one row of the table in which we want to view the data
					$MembershipRules = Build-ArrayObject -Collection $Collection -IncludeMembershipRules $IncludeMembershipRules -ExcludeMembershipRules $ExcludeMembershipRules -QueryMembershipRules $QueryMembershipRules -DirectMembershipRules $DirectMembershipRules -Deployments $FilteredDeployments
					Write-Verbose ($MembershipRules | Format-List | Out-String)
					
					Write-Verbose "Adding object to the output array."
					$output.Add($MembershipRules) | Out-Null
				}
			}
		} catch {
			Write-Host $_
		} finally {
			$PSStyle.Progress.View = $myPSStyle
			Set-Location $myPWD
		}
	}

	end{
		Write-Host "Done!"
		#$PSStyle.Progress.View = $myPSStyle
		#Set-Location $myPWD
		if($Json){
			return $output | ConvertTo-Json
		}else{
			return $output
		}
	}
}
Export-ModuleMember Get-CMOrgModelDeploymentRules
