Write-Output "Executing Full Infrastructure Deployment"
Connect-AzAccount
$sPOL = Get-ChildItem "$($PSScriptRoot)\00-POL\" | where-object {$_.extension -eq ".ps1"} | % { & $_.FullName }
$sLAW = Get-ChildItem "$($PSScriptRoot)\01-LAW\" | where-object {$_.extension -eq ".ps1"} | % { & $_.FullName }
$sRG = Get-ChildItem "$($PSScriptRoot)\02-RG\" | where-object {$_.extension -eq ".ps1"} | % { & $_.FullName }
