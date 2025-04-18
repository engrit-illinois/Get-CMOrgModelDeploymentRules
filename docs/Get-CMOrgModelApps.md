---
external help file: MECMOrgModelDeployments-help.xml
Module Name: MECMOrgModelDeployments
online version:
schema: 2.0.0
---

# Get-CMOrgModelApps

## SYNOPSIS
A documentation function which queries all relevant deployment collections in Engineering IT's MECM environment so we know what apps are being deployed. Assumes you are using the [`New-CMOrgModelDeploymentCollection`](https://github.com/engrit-illinois/New-CMOrgModelDeploymentCollection) cmdlet to build your deployment collections.

## SYNTAX

```
Get-CMOrgModelApps [-Json] [-ISOnly] [-AppNamesOnly] [[-CollectionQueries] <String[]>] [-Test]
 [[-TestCollectionQueries] <String[]>] [-RandomCollections] [[-RandomCollectionsNum] <Int32>]
 [[-Prefix] <String>] [[-SiteCode] <String>] [[-Provider] <String>] [[-CMPSModulePath] <String>]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
A documentation function which queries all relevant deployment collections in Engineering IT's MECM environment so we know what apps are being deployed. Assumes you are using the [`New-CMOrgModelDeploymentCollection`](https://github.com/engrit-illinois/New-CMOrgModelDeploymentCollection) cmdlet to build your deployment collections.

## EXAMPLES

### Example 1
```powershell
PS C:\> Get-CMOrgModelApps
```

Returns a list of all deployed applications in the Deployment Collection Model.

## PARAMETERS

### -AppNamesOnly
Switch parameter that when specified, removes the Prefix from the returned Application list. This can help improve readability of the resulting output.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CMPSModulePath
Path to the MECM Powershell Module file. By default points to the default install location.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CollectionQueries
String array of queries to use to find deployment collections.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ISOnly
Switch parameter if you only want Deployment Collections for Instruction.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Json
Switch parameter if you want a JSON output instead of a powershell output. Essentially does `ConvertTo-Json` but a little fancier.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Prefix
Defines the prefix used to find your collections.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Provider
Defines the provider, usually the CAS, that you will be querying.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RandomCollections
Switch parameter for if you want a random selection of collections instead of all of them. Useful if you're just doing some testing or want a random sampling, since getting all collections can take a long time.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RandomCollectionsNum
Int parameter that defines how many random collections to sample when using the `-RandomCollections` parameter. By default set to 10.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SiteCode
Defines the site code.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Test
Switch parameter to indicate you are doing a test.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TestCollectionQueries
String array of queries to use to find deployment collections. This is used instead of `$CollectionQueries` when `-Test` is specified.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### None

## OUTPUTS

### System.Object
## NOTES

## RELATED LINKS
