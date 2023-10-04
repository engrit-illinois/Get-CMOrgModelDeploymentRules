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
        $DeploymentType
    )

    if($Application.Count -le 1){
        $Comments = $Application.LocalizedDescription
    }else{
        Write-Verbose "Building comment array for $($Collection.Name)"
        $Comments = New-Object System.Collections.ArrayList
        foreach($App in @($Application)){
            $Comments.Add($App.LocalizedDescription) | Out-Null
        }
    }

    Write-Verbose "Building the custom array for $($Collection.Name)..."
    Write-Verbose "Name = $($AppDeployment.ApplicationName)"
    Write-Verbose "DeploymentStartTime = $($AppDeployment.StartTime.AddHours(5))"
    Write-Verbose "Action = $Action"
    Write-Verbose "DirectMembershipRules = $($DirectMembershipRules.RuleName)"
    Write-Verbose "ExcludeMembershipRules = $($ExcludeMembershipRules.RuleName)"
    Write-Verbose "IncludeMembershipRules = $($IncludeMembershipRules.RuleName)"
    Write-Verbose "QueryMembershipRules = $($QueryMembershipRules.RuleName)"
    Write-Verbose "DeploymentType = $DeploymentType"
    Write-Verbose "Supersedence = $($AppDeployment.UpdateSupersedence)"
    Write-Verbose "Comments = $($Comments)"
    ## Not sure how to handle Direct, Exclude, or Query rules yet. Will deal with them later
    $output = [PSCustomObject]@{
        CollectionName                  = $Collection.Name
        Name                            = $AppDeployment.ApplicationName
        ### Lazy hack to account for us not being on GMT
        DeploymentStartTime             = $AppDeployment.StartTime.AddHours(5)
        Action                          = $Action
        DirectMembershipRules           = $DirectMembershipRules.RuleName
        ExcludeMembershipRules          = $ExcludeMembershipRules.RuleName
        IncludeMembershipRules          = $IncludeMembershipRules.RuleName
        QueryMembershipRules            = $QueryMembershipRules.RuleName
        OverrideServiceWindows          = $AppDeployment.OverrideServiceWindows
        RebootOutsideOfServiceWindows   = $AppDeployment.RebootOutsideOfServiceWindows
        DeploymentType                  = $DeploymentType
        Supersedence                    = $AppDeployment.UpdateSupersedence
        Comments                        = $Comments
    }
    $output
}

function Get-CMOrgModelDeploymentRules{

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$Json,
        [switch]$ISOnly,
        [int]$Test,                                                                         # Shuffles then shrinks the array this runs on for faster testing
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
        } catch {
            Write-Host $_
        }
    }

    process{
        try{
            Write-Host "Getting all collections... (note: this takes a while)"
            $DeployCollections = @(Get-CMDeviceCollection -Name "UIUC-ENGR-Deploy*") + @(Get-CMDeviceCollection -Name "UIUC-ENGR-IS Deploy*")
            $DeployCollections = $DeployCollections | Sort-Object -Property Name

            if($Test){
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
                $DirectMembershipRules = Get-CMDeviceCollectionDirectMembershipRule -CollectionName $Collection.Name
                Write-Verbose "Getting Exclude Membership Rules for $($Collection.Name)..."
                $ExcludeMembershipRules = Get-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection.Name
                Write-Verbose "Getting Include Membership Rules for $($Collection.Name)..."
                $IncludeMembershipRules = Get-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection.Name
                Write-Verbose "Getting Query Membership Rules for $($Collection.Name)..."
                $QueryMembershipRules = Get-CMDeviceCollectionQueryMembershipRule -CollectionName $Collection.Name
                
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
                    switch($DeploySummary.DesiredConfigType){
                        1 { $Action = "INSTALL" }
                        2 { $Action = "UNINSTALL" }
                        Default {
                            "UNKNOWN"
                        }
                    }
                
                    ## Available or Required
                    Write-Verbose "Determining whether this is Required or Available"
                    $OfferType = ($AppDeployment).OfferTypeID
                    switch ($OfferType){
                        0 { $DeploymentType = "REQUIRED" }
                        2 { $DeploymentType = "AVAILABLE" }
                        Default {
                            "UNKNOWN"
                        }
                    }

                    if(
                        ($ISOnly) -and
                        (-not ($ExcludeMembershipRules | Where-Object {$_.RuleName -like "UIUC-ENGR-IS*"}) ) -and
                        (-not ($IncludeMembershipRules | Where-Object {$_.RuleName -like "UIUC-ENGR-IS*"}) )
                    ) {
                        # This weakly only filters by Exclude and Include membership rules, because we don't have easily identifiable conventions via Direct or Query-based membership rules
                        Write-Verbose "ISOnly flag was declared, but no Include or Exclude membership rules were found on $($Collection.Name) referencing `"UIUC-ENGR-IS*`" collections."
                    } else {
                        $MembershipRules = Build-ArrayObject -Collection $Collection -Application $Applications -AppDeployment $AppDeployment -Action $Action -DirectMembershipRules $DirectMembershipRules -ExcludeMembershipRules $ExcludeMembershipRules -IncludeMembershipRules $IncludeMembershipRules -QueryMembershipRules $QueryMembershipRules -DeploymentType $DeploymentType
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