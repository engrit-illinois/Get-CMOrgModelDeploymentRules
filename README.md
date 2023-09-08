# Get-CMOrgModelDeploymentRules

# Description
This function queries all relevant deployment collections in Engineering IT's MECM environment so we know what apps are being deployed where.

# Syntax
```powershell
Get-CMOrgModelDeploymentRules
    [-Json]
    [-ISOnly]
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