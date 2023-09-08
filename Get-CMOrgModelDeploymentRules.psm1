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

function Get-CMOrgModelDeploymentRules{

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$Json,
        [string]$Prefix="UIUC-ENGR-",
		[string]$SiteCode="MP0",
		[string]$Provider="sccmcas.ad.uillinois.edu",
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
            $DeviceCollections = Get-CMDeviceCollection
            
            Write-Host "Filtering collections..."
            $DeployCollections = $DeviceCollections | Where-Object -Property Name -Like "UIUC-ENGR-IS Deploy*"
            $DeployCollections += $DeviceCollections | Where-Object -Property Name -Like "UIUC-ENGR-Deploy*"

            Write-Host "Getting Membership Rules..."
            foreach($Collection in $DeployCollections){
                Write-Host "Processing $($Collection.Name)..."
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
            
                if($null -eq $AppDeployment){
                    Write-Host "$($Collection.Name) has no App deployment, so skipping!"
                }else{
                    ## Not sure how to handle Direct, Exclude, or Query rules yet. Will deal with them later
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
                    $MembershipRules = [PSCustomObject]@{
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
                    Write-Verbose "Adding details for $($Collection.Name) to the output..."
                    $output.Add($MembershipRules) | Out-Null
                }
            }
        } catch {
            Write-Host $_
        } finally {
            Set-Location $myPWD
        }
    }

    end{
        Set-Location $myPWD
        if($Json){
            return $output | ConvertTo-Json
        }else{
            return $output
        }
    }
}
Export-ModuleMember Get-CMOrgModelDeploymentRules