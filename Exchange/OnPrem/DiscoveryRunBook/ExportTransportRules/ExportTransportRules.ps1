$file = Export-TransportRuleCollection
[System.IO.File]::WriteAllBytes('C:\Support\Rules.xml', $file.FileData)