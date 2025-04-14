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
		Write-Verbose "Found $($Deployments.Count) Deployments."
		
		Write-Verbose "Processing Deployments..."
		$Deployments = $Deployments | ForEach-Object {
			$Deployment = $_
			$ContentName = $Deployment.ApplicationName
			Write-Verbose "Processing `"$contentName`"..."
			
			# Store formatted version of deployment/content name
			# As far as I can tell, this should be identical to whatever name is stored in the actual deployed content, and
			# it's just easier to reference from the deployment data
			$ContentNameIcon = $FEATURETYPES | Where-Object { $_.Id -eq $Deployment.FeatureType } | Select -ExpandProperty "Icon"
			if(-not $ContentNameIcon) { $ContentNameIcon = "❓" }
			$Deployment | Add-Member -NotePropertyName "_ContentNameF" -NotePropertyValue "$($ContentNameIcon)&nbsp;$($ContentName)"
			
			# Store formatted version of the deployment purpose
			$Purpose = $DEPLOYMENTINTENTS | Where-Object { $_.Id -eq $Deployment.DeploymentIntent }
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
