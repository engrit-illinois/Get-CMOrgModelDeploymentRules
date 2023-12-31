# Get-CMOrgModelDeploymentRules

# Description
A documentation function which queries all relevant deployment collections in Engineering IT's MECM environment so we know what apps are being deployed where. Assumes you are using the [`New-CMOrgModelDeploymentCollection`](https://github.com/engrit-illinois/New-CMOrgModelDeploymentCollection) cmdlet to build your deployment collections.

# Syntax
```powershell
Get-CMOrgModelDeploymentRules
    [-Json]
    [-ISOnly]
    [-NoProgressBar]
```

# Examples
This command queries all the collections and outputs the result in JSON.
```powershell
Get-CMOrgModelDeploymentRules -Json
```

This command queries all collections and outputs only results with Instructional device collections inside in JSON.
```powershell
Get-CMOrgModelDeploymentRules -ISOnly -Json
```

# Parameters
### -Json
Returns an output in the JSON format

### -ISOnly
Returns only results with Instructional device collections inside

Note: This simple filtering is just done on Inclusion and Exclusion rules on Device collections by name. This switch will not catch Direct adds or queries.

### -NoProgressBar
By default, this command uses a progress bar to show progress in processing membership rules. Specifying the `-NoProgressBar` switch will disable the progress bar and instead show each collection's name as it is being processed. Useful if you want to see all of the processed collections, but can be messy.