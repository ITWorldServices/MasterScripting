$file = Export-TransportRuleCollection -ExportLegacyRules
[System.IO.File]::WriteAllBytes('C:\Support\LegacyRules.xml', $file.FileData)