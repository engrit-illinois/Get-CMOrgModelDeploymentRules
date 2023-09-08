# Get-CMOrgModelDeploymentRules

# Description
This function queries all relevant deployment collections in Engineering IT's MECM environment so we know what apps are being deployed where.

# Syntax
```powershell
Get-CMOrgModelDeploymentRules
    [-Json]
```

# Examples
This command queries all the collections and outputs the result in JSON
```powershell
Get-CMOrgModelDeploymentRules -Json
```
