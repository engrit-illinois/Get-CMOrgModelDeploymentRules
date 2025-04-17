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