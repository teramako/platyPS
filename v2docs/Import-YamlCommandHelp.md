---
document type: cmdlet
external help file: Microsoft.PowerShell.PlatyPS.dll-Help.xml
HelpUri: ''
Locale: en-US
Module Name: Microsoft.PowerShell.PlatyPS
ms.date: 07/10/2024
PlatyPS schema version: 2024-05-01
title: Import-YamlCommandHelp
---

# Import-YamlCommandHelp

## SYNOPSIS

Imports Yaml help content into **CommandHelp** objects.

## SYNTAX

### Path (Default)

```
Import-YamlCommandHelp [-AsDictionary] [-Path] <string[]> [<CommonParameters>]
```

### LiteralPath

```
Import-YamlCommandHelp [-AsDictionary] -LiteralPath <string[]> [<CommonParameters>]
```

## ALIASES

## DESCRIPTION

The command imports Yaml files containing command help and creates **CommandHelp** objects. The
**CommandHelp** object is a structured representation of the help content that can be used to export
to different formats.

## EXAMPLES

### Example 1 - Import all cmdlet Yaml files in a folder

The following example import cmdlet Markdown files in a folder and converts them to **CommandHelp**
objects. These objects can be used to export to different formats.

```powershell
$chArray = Import-MarkdownCommandHelp -Path .\v1\*-*.yml
$chArray.title
```

```Output
Compare-CommandHelp
ConvertTo-CommandHelp
Export-MamlCommandHelp
Export-MarkdownCommandHelp
Export-MarkdownModuleFile
Export-YamlCommandHelp
Export-YamlModuleFile
Import-MamlHelp
Import-MarkdownCommandHelp
Import-MarkdownModuleFile
Import-YamlCommandHelp
Import-YamlModuleFile
New-MarkdownCommandHelp
Test-MarkdownCommandHelp
```

## PARAMETERS

### -AsDictionary

By default this cmdlet returns **CommandHelp** objects. When you use this parameter. the cmdlet
returns the same information as an **OrderedDictionary** object.

```yaml
Type: System.Management.Automation.SwitchParameter
DefaultValue: ''
SupportsWildcards: false
ParameterValue: []
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -LiteralPath

The path to the Yaml command file. Specifies a path to one or more locations. The value of LiteralPath
is used exactly as it's typed. No characters are interpreted as wildcards. If the path includes
escape characters, enclose it in single quotation marks. Single quotation marks tell PowerShell not
to interpret any characters as escape sequences.

For more information, see
[about_Quoting_Rules](/powershell/module/microsoft.powershell.core/about/about_quoting_rules).

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: false
ParameterValue: []
Aliases: []
ParameterSets:
- Name: LiteralPath
  Position: Named
  IsRequired: true
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Path

The path to the Yaml command file. Specifies a path to one or more locations.

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: true
ParameterValue: []
Aliases:
- FullName
ParameterSets:
- Name: Path
  Position: 0
  IsRequired: true
  ValueFromPipeline: true
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.String

## OUTPUTS

## System.Collections.Specialized.OrderedDictionary

This cmdlet returns an **OrderedDictionary** object when you use the **AsDictionary** parameter.

### Microsoft.PowerShell.PlatyPS.Model.CommandHelp

By default, this cmdlet returns **CommandHelp** objects.

## NOTES

## RELATED LINKS

[Export-YamlCommandHelp](Export-YamlCommandHelp.md)