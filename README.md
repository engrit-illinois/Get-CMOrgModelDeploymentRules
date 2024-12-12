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
When specified, output is returned in JSON format.  

### -ISOnly
When specified, only deployment collections matching the given queries which also include Instructional Services device collections are returned.  
Note: This simple filtering is just done on Include and Exclude rules, based on the naming convention of the included/excluded collections. This will not catch collections containing IS devices due to Direct Query-based membership rules.  

### -IncludeCollectionsWithNoDeployments
When specified, deployment collections which match the given queries are returned regardless of whether there are actually any deployments to the matching collections.  
Useful to find collections intended to be deployment collections, but where the deployment is missing. Often this just means the app that previous deployed was retired and the collection was never removed.  
By default collections which have no deployments are omitted from the output.  

### -NoProgressBar
By default, this command uses a progress bar to show progress in processing membership rules. Specifying the `-NoProgressBar` switch will disable the progress bar and instead show each collection's name as it is being processed. Useful if you want to see all of the processed collections, but can be messy.  

### -CollectionQueries [string[]]
String array of wildcard queries defining which collections should be polled.  
Default is `@("UIUC-ENGR-Deploy*","UIUC-ENGR-IS Deploy*","UIUC-ENGR-IS Uninstall*","UIUC-ENGR-IS Uninstall*")`.  

### -Test
When specified the queries defined by `-TestCollectionQueries` are used, instead of those defined by `-CollectionQueries`.  

### -TestCollectionQueries [string[]]
Alternate set of deployment collection queries to be used for faster testing.  
Default is `@("UIUC-ENGR-IS Deploy A*","UIUC-ENGR-IS Uninstall A*")`.  

### -RandomCollections
When specified, only a random set of deployment collections matching the given queries are processed, shortening the time to test, while still giving a hopefully representative, unique sample.  
The number of collections in the randomly chosen set is defined by `-RandomCollectionsNum`.  

### -RandomCollectionsNum [int]
The number of collections to randomly select, when `-RandomCollections` is specified.  