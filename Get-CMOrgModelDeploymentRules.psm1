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
        $AppDeployment,
        $Action,
        $DirectMembershipRules,
        $ExcludeMembershipRules,
        $IncludeMembershipRules,
        $QueryMembershipRules,
        $DeploymentType
    )
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
    
    ## Not sure how to handle Direct, Exclude, or Query rules yet. Will deal with them later
    $output = [PSCustomObject]@{
        CollectionName              = $Collection.Name
        Name                        = $AppDeployment.ApplicationName
        ### Lazy hack to account for us not being on GMT
        DeploymentStartTime         = $AppDeployment.StartTime.AddHours(5)
        Action                      = $Action
        DirectMembershipRules       = $DirectMembershipRules.RuleName
        ExcludeMembershipRules      = $ExcludeMembershipRules.RuleName
        IncludeMembershipRules      = $IncludeMembershipRules.RuleName
        QueryMembershipRules        = $QueryMembershipRules.RuleName
        DeploymentType              = $DeploymentType
        Supersedence                = $AppDeployment.UpdateSupersedence
    }
    $output
}

function Get-CMOrgModelDeploymentRules{

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$Json,
        [switch]$ISOnly,
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
        } catch {
            Write-Host $_
        }
    }

    process{
        try{
            Write-Host "Getting all collections... (note: this takes a while)"
            $DeployCollections = @(Get-CMDeviceCollection -Name "UIUC-ENGR-Deploy*") + @(Get-CMDeviceCollection -Name "UIUC-ENGR-IS Deploy*")

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
                        $MembershipRules = Build-ArrayObject -Collection $Collection -AppDeployment $AppDeployment -Action $Action -DirectMembershipRules $DirectMembershipRules -ExcludeMembershipRules $ExcludeMembershipRules -IncludeMembershipRules $IncludeMembershipRules -QueryMembershipRules $QueryMembershipRules -DeploymentType $DeploymentType
                        Write-Verbose "Adding to the function output array."
                        $output.Add($MembershipRules) | Out-Null
                    }
                }
            }
            $output = $output | Sort-Object -Property CollectionName
        } catch {
            Write-Host $_
        } finally {
            Write-Host "Operation Cancelled by User"
            Set-Location $myPWD
        }
    }

    end{
        Write-Host "Done!"
        Set-Location $myPWD
        if($Json){
            return $output | ConvertTo-Json
        }else{
            return $output
        }
    }
}
Export-ModuleMember Get-CMOrgModelDeploymentRules