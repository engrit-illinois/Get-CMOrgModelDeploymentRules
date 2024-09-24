$DEFAULT_PREFIX = "UIUC-ENGR-"
$DEFAULT_SITE_CODE = "MP0"
$DEFAULT_PROVIDER = "sccmcas.ad.uillinois.edu"

function Connect-ToMECM {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$Prefix,
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
        $ApplicationDeployment
    )

    [flags()] 
    enum SMS_ApplicationAssignment_OfferFlags
    {
        None    = 0
        PREDEPLOY = 1
        ONDEMAND = 2
        ENABLEPROCESSTERMINATION = 4
        ALLOWUSERSTOREPAIRAPP = 8
        RELATIVESCHEDULE = 16
        HIGHIMPACTDEPLOYMENT = 32
        IMPLICITUNINSTALL = 64
    }
	
	$ApplicationDeployment | ForEach-Object {
		[SMS_ApplicationAssignment_OfferFlags]$assignmentFlags = $_.OfferFlags

		Write-Verbose "Checking if the app has implicit uninstall enabled"
		if ($assignmentFlags.HasFlag([SMS_ApplicationAssignment_OfferFlags]::IMPLICITUNINSTALL)) {
			Write-Verbose "Implicit uninstall was enabled"
			return $true
		}else{
			Write-Verbose "Implicit uninstall was disabled"
			return $false
		}
	}
}

function Build-ArrayObject {
    
    [CmdletBinding()]
    param(
        $Collection,
        $Application,
        $AppDeployment,
        $Action,
        $DirectMembershipRules,
        $ExcludeMembershipRules,
        $IncludeMembershipRules,
        $QueryMembershipRules,
        $Purpose
    )
    
	$ImplicitUninstall = Resolve-ImplicitUninstall -ApplicationDeployment $AppDeployment
	
	# Certain fields may be arrays of values, if the collection has multiple deployments.
	# Save a formatted version of those which are more readable after being JSON-ified:
	if($null -ne $IncludeMembershipRules.RuleName) {
		$IncludeMembershipRulesFormatted = "🔹" + ($IncludeMembershipRules -join " \\🔹")
	}
	
	if($null -ne $AppDeployment.ApplicationName) {
		$NameFormatted = "🔹" + ($AppDeployment.ApplicationName -join " \\🔹")
	}
	
	if($null -ne $Action) {
		$ActionFormatted = $Action -join " \\"
		$ActionFormatted = $ActionFormatted.Replace("INSTALL","📦INSTALL")
		$ActionFormatted = $ActionFormatted.Replace("UNINSTALL","🗑️UNINSTALL")
	}
	
	if($null -ne $Application.LocalizedDescription) {
		if($Application.LocalizedDescription -ne "") {
			$CommentsFormatted = "🔹" + ($Application.LocalizedDescription -join " \\🔹")
		}
	}
	
	if($null -ne $Purpose) {
		$PurposeFormatted = $Purpose -join " \\"
		$PurposeFormatted = $PurposeFormatted.Replace("AVAILABLE","💡AVAILABLE")
		$PurposeFormatted = $PurposeFormatted.Replace("REQUIRED","🔒REQUIRED")
	}
	
	if($null -ne $AppDeployment.UpdateSupersedence) {
		$SupersedenceFormatted = $AppDeployment.UpdateSupersedence -join " \\"
		$SupersedenceFormatted = $SupersedenceFormatted.Replace("True","👑✔️Enabled")
		$SupersedenceFormatted = $SupersedenceFormatted.Replace("False","👑❌Disabled")
	}
	
	if($null -ne $ImplicitUninstall) {
		$ImplicitUninstallFormatted = $ImplicitUninstall -join " \\"
		$ImplicitUninstallFormatted = $ImplicitUninstallFormatted.Replace("True","🚮Implicit")
		$ImplicitUninstallFormatted = $ImplicitUninstallFormatted.Replace("False","🚯Not Implicit")
	}

    $CollectionName = $Collection.Name
    $Name = $AppDeployment.ApplicationName
    ### Lazy hack to account for us not being on GMT
    $DeploymentStartTime = $($AppDeployment.StartTime.AddHours(5))
    $OverrideServiceWindows = $AppDeployment.OverrideServiceWindows
    $RebootOutsideOfServiceWindows = $AppDeployment.RebootOutsideOfServiceWindows
    $Supersedence = $AppDeployment.UpdateSupersedence
    $Comments = $Application.LocalizedDescription

    Write-Verbose "Building the custom array for $CollectionName..."
    Write-Verbose "Name = $Name"
    Write-Verbose "NameFormatted = $NameFormatted"
    Write-Verbose "DeploymentStartTime = $DeploymentStartTime"
    Write-Verbose "Action = $Action"
    Write-Verbose "ActionFormatted = $ActionFormatted"
    Write-Verbose "DirectMembershipRules = $DirectMembershipRules"
    Write-Verbose "ExcludeMembershipRules = $ExcludeMembershipRules"
    Write-Verbose "IncludeMembershipRules = $IncludeMembershipRules"
    Write-Verbose "IncludeMembershipRulesFormatted = $IncludeMembershipRulesFormatted"
    Write-Verbose "QueryMembershipRules = $QueryMembershipRules"
    Write-Verbose "OverrideServiceWindows = $OverrideServiceWindows"
    Write-Verbose "RebootOutsideOfServiceWindows = $RebootOutsideOfServiceWindows"
    Write-Verbose "Purpose = $Purpose"
    Write-Verbose "PurposeFormatted = $PurposeFormatted"
    Write-Verbose "Supersedence = $Supersedence"
    Write-Verbose "SupersedenceFormatted = $SupersedenceFormatted"
    Write-Verbose "ImplicitUninstall = $ImplicitUninstall"
    Write-Verbose "ImplicitUninstallFormatted = $ImplicitUninstallFormatted"
    Write-Verbose "Comments = $Comments"
    Write-Verbose "CommentsFormatted = $CommentsFormatted"
	
    ## Not sure how to handle Direct, Exclude, or Query rules yet. Will deal with them later
    [PSCustomObject]@{
        CollectionName                  = $CollectionName
        Name                            = $Name
		NameFormatted                   = $NameFormatted
        DeploymentStartTime             = $DeploymentStartTime
        Action                          = $Action
		ActionFormatted                 = $ActionFormatted
        DirectMembershipRules           = $DirectMembershipRules
        ExcludeMembershipRules          = $ExcludeMembershipRules
        IncludeMembershipRules          = $IncludeMembershipRules
		IncludeMembershipRulesFormatted = $IncludeMembershipRulesFormatted
        QueryMembershipRules            = $QueryMembershipRules
        OverrideServiceWindows          = $OverrideServiceWindows
        RebootOutsideOfServiceWindows   = $RebootOutsideOfServiceWindows
        Purpose                         = $Purpose
		PurposeFormatted                = $PurposeFormatted
        Supersedence                    = $Supersedence
		SupersedenceFormatted           = $SupersedenceFormatted
        ImplicitUninstall               = $ImplicitUninstall
		ImplicitUninstallFormatted      = $ImplicitUninstallFormatted
        Comments                        = $Comments
		CommentsFormatted               = $CommentsFormatted
    }
}

function Get-CMOrgModelDeploymentRules{

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$Json,
        [switch]$ISOnly,
        $Test,                                                                         # Shuffles then shrinks the array this runs on for faster testing
        [switch]$NoProgressBar,
        [string]$Prefix = $DEFAULT_PREFIX,
		[string]$SiteCode=$DEFAULT_SITE_CODE,
		[string]$Provider=$DEFAULT_PROVIDER,
        [string]$CMPSModulePath="$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
    )
    begin{
        $myPWD = $PWD.Path

        try {
            Connect-ToMECM -Prefix $Prefix -SiteCode $SiteCode -Provider $Provider -CMPSModulePath $CMPSModulePath
            
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
            if($Test -and ($TestType -like "*string*")){
                $DeployCollections = Get-CMDeviceCollection -Name $Test
            } else {
                $DeployCollections = @(Get-CMDeviceCollection -Name "UIUC-ENGR-Deploy*") + @(Get-CMDeviceCollection -Name "UIUC-ENGR-IS Deploy*")
				$DeployCollections = $DeployCollections | Sort-Object -Property Name
            }

            if($Test -and ($TestType -like "Int*")){
                $DeployCollections = $DeployCollections | Get-Random -Count $Test
            }

            $TotalCollections = $DeployCollections.Count                                    # Using for progress bar
            $PercentComplete = 0                                                            # Using for progress bar
            $CollectionCount = 0                                                            # Using for progress bar
            if($NoProgressBar){
                Write-Host "Processing Membership Rules..."
            }
            foreach($Collection in $DeployCollections){
                if($NoProgressBar){
                    Write-Host "Processing $($Collection.Name)..."
                }else{
                    Write-Progress -Activity "Processing Membership Rules..." -Status "$($PercentComplete)% Processing $($Collection.Name)..." -PercentComplete $PercentComplete
                    $CollectionCount++                                                          # Using for progress bar
                    $PercentComplete = [int](($CollectionCount / $TotalCollections) * 100)      # Using for progress bar
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
                
                Write-Verbose "Getting the Application Deployment for $($Collection.Name)..."
                $AppDeployment = Get-CMApplicationDeployment -Collection $Collection
                Write-Verbose "Getting other Deployment info for $($Collection.Name)..."
                $DeploySummary = Get-CMDeployment -CollectionName $Collection.Name
                

                if($null -eq $AppDeployment){
                    if($NoProgressBar){
                        Write-Host "Skipping $($Collection.Name) because it has no App deployment"
                    }
                }else{
                    Write-Verbose "Getting Application info..."
                    Write-Verbose "Initializing the empty Applications array"
                    $Applications = New-Object System.Collections.ArrayList
                    foreach($AppD in $AppDeployment){
                        Write-Verbose "Getting Application info for $($AppD.ApplicationName)"
                        $Application = Get-CMApplication -Fast -Name $AppD.ApplicationName
                        Write-Verbose "Adding to the Applications array."
                        $Applications.Add($Application) | Out-Null
                    }

                    ## Install or Uninstall
                    Write-Verbose "Determining whether this is Install or Uninstall"
                    $Action = switch($DeploySummary.DesiredConfigType){
                        1 { "INSTALL" }
                        2 { "UNINSTALL" }
                        Default { "UNKNOWN" }
                    }
                
                    ## Available or Required
                    Write-Verbose "Determining whether this is Required or Available"
                    $Purpose = switch($AppDeployment.OfferTypeID){
                        0 { "REQUIRED" }
                        2 { "AVAILABLE" }
                        Default { "UNKNOWN" }
                    }

                    if(
                        ($ISOnly) -and
                        (-not ($ExcludeMembershipRules | Where-Object {$_.RuleName -like "UIUC-ENGR-IS*"}) ) -and
                        (-not ($IncludeMembershipRules | Where-Object {$_.RuleName -like "UIUC-ENGR-IS*"}) )
                    ) {
                        # This weakly only filters by Exclude and Include membership rules, because we don't have easily identifiable conventions via Direct or Query-based membership rules
                        Write-Verbose "ISOnly flag was declared, but no Include or Exclude membership rules were found on $($Collection.Name) referencing `"UIUC-ENGR-IS*`" collections."
                    } else {
                        $MembershipRules = Build-ArrayObject -Collection $Collection -Application $Applications -AppDeployment $AppDeployment -Action $Action -DirectMembershipRules $DirectMembershipRules -ExcludeMembershipRules $ExcludeMembershipRules -IncludeMembershipRules $IncludeMembershipRules -QueryMembershipRules $QueryMembershipRules -Purpose $Purpose
                        Write-Verbose "Adding to the function output array."
                        Write-Verbose ($MembershipRules | Format-List | Out-String)
                        $output.Add($MembershipRules) | Out-Null
                    }
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