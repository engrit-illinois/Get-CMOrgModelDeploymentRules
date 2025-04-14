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
	$Action = $DESIREDCONFIGTYPES | Where-Object { $_.Id -eq $ContentDeployment.DesiredConfigType }
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
