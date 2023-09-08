Function Get-CMOrgModelDeploymentCollections{

    Prep-MECM
     
    Write-Verbose "Getting all collections... (note: this takes a while)"
    $DeviceCollections = Get-CMDeviceCollection
     
    Write-Verbose "Filtering collections..."
    $DeployCollections = $DeviceCollections | Where-Object -Property Name -Like "UIUC-ENGR-IS Deploy*"
    $DeployCollections += $DeviceCollections | Where-Object -Property Name -Like "UIUC-ENGR-Deploy*"
     
    ## Declare the empty arraylist
    $output = New-Object System.Collections.ArrayList
     
    Write-Verbose "Getting Membership Rules..."
    foreach($Collection in $DeployCollections){
        Write-Verbose "Processing $($Collection.Name)..."
        $DirectMembershipRules = Get-CMDeviceCollectionDirectMembershipRule -CollectionName $Collection.Name
        $ExcludeMembershipRules = Get-CMDeviceCollectionExcludeMembershipRule -CollectionName $Collection.Name
        $IncludeMembershipRules = Get-CMDeviceCollectionIncludeMembershipRule -CollectionName $Collection.Name
        $QueryMembershipRules = Get-CMDeviceCollectionQueryMembershipRule -CollectionName $Collection.Name
         
        $AppDeployment = Get-CMApplicationDeployment -Collection $Collection
        $DeploySummary = Get-CMDeployment -CollectionName $Collection.Name
     
        ## Install or Uninstall
        switch($DeploySummary.DesiredConfigType){
            1 { $Action = "INSTALL" }
            2 { $Action = "UNINSTALL" }
            Default {
                "UNKNOWN"
            }
        }
     
        ## Available or Required
        $OfferType = ($AppDeployment).OfferTypeID
        switch ($OfferType){
            0 { $DeploymentType = "REQUIRED" }
            2 { $DeploymentType = "AVAILABLE" }
            Default {
                "UNKNOWN"
            }
        }
     
        ## Not sure how to handle Direct, Exclude, or Query rules yet. Will deal with them later
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
        $output.Add($MembershipRules) | Out-Null
    }
   
    return $output
}
