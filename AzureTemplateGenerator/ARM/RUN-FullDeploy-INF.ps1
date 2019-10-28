Write-Output "Executing Full Deployment"
Connect-AzAccount
$sPOL = Get-ChildItem "$($PSScriptRoot)\00-POL\" | where-object {$_.extension -eq ".ps1"} | % { & $_.FullName }
$sRG = Get-ChildItem "$($PSScriptRoot)\01-RG\" | where-object {$_.extension -eq ".ps1"} | % { & $_.FullName }