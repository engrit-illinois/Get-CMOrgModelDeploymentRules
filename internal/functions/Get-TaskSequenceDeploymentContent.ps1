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
