#requires -version 5
<#
.SYNOPSIS
    Generate ARM Templates from Excel
.DESCRIPTION
    GoC tool to Generate ARM Templates for Azure Network Infrastructure from specially designed Excel file
.PARAMETER filename
    Specifies the Excel file containing the Azure parameters. Must be formatted according to spec.
.PARAMETER outputfolder
    Specifies the folder where ARM template .JSON files should be created.
.INPUTS
    None.
.OUTPUTS
    None.
.NOTES
    Version:        1.0
    Author:         Ken Morency - Transport Canada
    Creation Date:  Friday September 13, 2019
    Purpose/Change: Initial script development

    MIT License

    Copyright (c) 2019 Government of Canada - Gouvernement du Canada

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
.EXAMPLE
    .PS> .\GenerateTemplates.ps1 -filename "\NetworkInfrastructure.xlsx" -outputfolder "\TEMPLATES\"
.EXAMPLE
    .PS> .\GenerateTemplates.ps1 "\NetworkInfrastructure.xlsx" "\TEMPLATES\"
.LINK
    https://github.com/KenMorencyTC/TC-Cloud-Team-Automation-Scripting
#>

param (
    [String] $filename = "$($PSScriptRoot)\TCNetworkInfrastructure.xlsx",
    [String] $outputfolder = "$($PSScriptRoot)\ARM\"
) 

#region INIT

$sScriptVersion = "1.0"
$sScriptName = "GenerateARMTemplates"

$sNSGOutPath = "$($outputfolder)01-NSG\"
$sRTOutPath = "$($outputfolder)02-RT\"
$sVNETOutPath = "$($outputfolder)03-VNET\"
$sPEEROutPath = "$($outputfolder)04-PEER\"
#$sAGWOutPath = "$($outputfolder)05-AGW\"
$sFWOutPath = "$($outputfolder)06-FW\"
#$sSAOutPath = "$($outputfolder)07-SA\"
#$sLAWOutPath = "$($outputfolder)08-LAW\"
#$sRSVBKPOutPath = "$($outputfolder)09-RSVBKP\"
#$sRSVASROutPath = "$($outputfolder)10-RSVASR\"

Write-Output "$($sScriptName) $($sScriptVersion)"

#Check/Install Import-Excel Module
$LoadedModules = Get-Module | Select-Object Name
if (!$LoadedModules -like "*ImportExcel*") {Install-Module -Name "ImportExcel" -Scope CurrentUser}

Write-Output "Generating ARM templates from $($filename)"

#endregion INIT

#region Functions

Function parseNSGs() {
    param ($curNSG) 
        $sOutput = '{
            "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentTemplate.json#",
            "contentVersion": "1.0.0.0",
            "parameters": {},
            "variables": {},
            "resources": [
            {
                "apiVersion": "2018-10-01",
                "type": "Microsoft.Network/networkSecurityGroups",
                "name": "' + $curNSG.NSGNAME + '",
                "location": "' + $curNSG.LOCATION + '",
                "properties": {
                    "securityRules": ['
        $iCount = 0
        $aRULESET = $curNSG.RULESET.Split(",")
        foreach ($curRULESET in $sRULESET) {
            $bUseRule = $false
            foreach ($sRule in $aRULESET) {
                if ($sRule -eq $curRULESET.RULESET) {
                    $bUseRule = $true
                }
            }
            if ($bUseRule -eq $true) {
                if ($iCount -gt 0) {
                    $sOutput += ','
                }
                $sOutput += '{
                    "name": "' + $curRULESET.RULENAME + '",
                    "properties": {
                        "description": "' + $curRULESET.DESCRIPTION + '",
                        "type": "' + $curRULESET.TYPE + '",
                        "protocol": "' + $curRULESET.PROTOCOL + '",
                        "sourcePortRange": "' + $curRULESET.SOURCE + '",
                        "destinationPortRange": "' + $curRULESET.DESTINATION + '",
                        "sourceAddressPrefix": "' + $curRULESET.SOURCEPREFIX + '",
                        "destinationAddressPrefix": "' + $curRULESET.DESTINATIONPREFIX + '",
                        "access": "' + $curRULESET.ACCESS + '",
                        "priority": ' + $curRULESET.PRIORITY + ',
                        "direction": "' + $curRULESET.DIRECTION + '"
                    }
                    }'
                $iCount += 1
            }
        }
            
        $sOutput += ']}}]}'
        return $sOutput
}
Function parseRTs() {
    param ($curRT)
    #Set route table template header JSON
    $sOutput = '{
        "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentTemplate.json#",
        "contentVersion": "1.0.0.0",
        "parameters": {},
        "variables": {},
        "resources": [
        {
            "apiVersion": "2018-10-01",
            "type": "Microsoft.Network/routeTables",
            "name": "' + $curRT.ROUTETABLENAME + '",
            "location": "' + $curRT.LOCATION + '",
            "properties": {
                "routes": ['
    $iCount = 0
    foreach ($curROUTESET in $sROUTESET) {
        #Check if current route is global for the region or name matched. (Ex. CACN or CORE-CACN-INTERNAL-RT)
        if ($curROUTESET.ROUTESET -eq $curRT.REGION -OR $curROUTESET.ROUTESET -eq $curRT.ROUTETABLENAME) {
            if ($iCount -gt 0) {
                $sOutput += ',' #Add comma prefix if not 1st object.
            }
            #Generate route template JSON.
            $sOutput += '{
                "name": "' + $curROUTESET.NAME + '",
                "properties": {
                    "addressPrefix": "' + $curROUTESET.ADDRESSPREFIX + '",
                    "nextHopType": "' + $curROUTESET.NEXTHOPTYPE + '",
                    "nextHopIpAddress": "' + $curROUTESET.NEXTHOPIPADDRESS + '"
                }}'
            $iCount += 1
        }
    }
    #Set route table template footer JSON
    $sOutput += ']}}]}'
    return $sOutput
}
Function parseVNETs() {
    param ($curVNET) 
    #TODO Add support for multiple address prefixes
    #Set vnet template header JSON
    $sOutput = '{
        "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentTemplate.json#",
        "contentVersion": "1.0.0.0",
        "parameters": {},
        "variables": {},
        "resources": [
        {
            "apiVersion": "2018-10-01",
            "type": "Microsoft.Network/virtualNetworks",
            "name": "' + $curVNET.VIRTUALNETWORKNAME + '",
            "location": "' + $curVNET.LOCATION + '",
            "properties": {
                "addressSpace": {
                    "addressPrefixes": [
                        "' + $curVNET.VNETADDRESSPREFIX + '"
                    ]
                    },
                "subnets": ['
    $iCount = 0
    foreach ($curSUBNET in $sSUBNET) {
        #Check if current subnet VNET is name matched with host VNET. (Ex. CORE-CACN-INTERNAL-VNET)
        if ($curVNET.VIRTUALNETWORKNAME -eq $curSUBNET.VIRTUALNETWORKNAME) {
            if ($iCount -gt 0) {
                $sOutput += ',' #Add comma prefix if not 1st object.
            }
            #Generate subnet template JSON.
            $sOutput += '{
                "name": "' + $curSUBNET.SUBNETNAME + '",
                "properties": {
                    "addressPrefix": "' + $curSUBNET.SUBNETPREFIX + '",
                    "networkSecurityGroup": {
                    "id": "[resourceId(''' + $curSUBNET.RESOURCEGROUPNAME + ''', ''Microsoft.Network/networkSecurityGroups/'', ''' + $curVNET.NSGNAME + ''')]"
                    },
                    "routeTable": {
                    "id": "[resourceId(''' + $curSUBNET.RESOURCEGROUPNAME + ''', ''Microsoft.Network/routeTables'', ''' + $curVNET.ROUTETABLENAME + ''')]"
                    }
                }
                }'
            $iCount += 1
        }
    }
    #Set route table template footer JSON
    $sOutput += ']}}]}'
    return $sOutput
}
Function parsePEERs() {
    param ($curVNET,$curSUB)
    #Set vnet template header JSON
    $sOutput = '{
        "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentTemplate.json#",
        "contentVersion": "1.0.0.0",
        "parameters": {},
        "variables": {},
        "resources": ['
    $iCount = 0
    foreach ($curPEER in $sPEER) {
        #TODO Determine why VIRTUALNETWORKNAME returning with trailing space when none exists in excel. Fix below is temporary.
        $sVNETName = [string]$curVNET.VIRTUALNETWORKNAME
        $sPEERName = [string]$curPEER.VIRTUALNETWORKNAME
        $sVNETName = $sVNETName.Trim()
        $sPEERName = $sPEERName.Trim()
        #Check if current vnet is name matched with host vnet peer. (Ex. CORE-CACN-INTERNAL-VNET)
        if ($sVNETName -eq $sPEERName) {
            $sSUBID = GetSubscriptionID($curPEER.REMOTESUB)
            #Write-Host $sPEERName' '$sVNETName
            if ($iCount -gt 0) {
                $sOutput += ',' #Add comma prefix if not 1st object.
            }
            #Generate vnet peering template JSON.
            $sOutput += '{
                "apiVersion": "2017-06-01",
                "type": "Microsoft.Network/virtualNetworks/virtualNetworkPeerings",
                "name": "' + $sVNETName + '/peering-to-' + $curPEER.REMOTERESOURCEGROUPNAME + '",
                "location": "' + $curVNET.LOCATION + '",
                "dependsOn": [
                "[resourceid(''Microsoft.Network/virtualNetworks'', ''' + $sVNETName + ''')]"
            ],
                "properties": {
                "allowVirtualNetworkAccess": ' + $curPEER.ALLOWVIRTUALNETWORKACCESS + ',
                "allowForwardedTraffic": ' + $curPEER.ALLOWFORWARDEDTRAFFIC + ',
                "allowGatewayTransit": ' + $curPEER.ALLOWGATEWAYTRANSIT + ',
                "useRemoteGateways": ' + $curPEER.USEREMOTEGATEWAYS + ',
                "remoteVirtualNetwork": {
                    "id": "[resourceId(''' + $sSUBID + ''', ''' + $curPEER.REMOTERESOURCEGROUPNAME + 
                        ''', ''Microsoft.Network/virtualNetworks'', ''' + $curPEER.REMOTEVIRTUALNETWORKNAME + ''')]"
                }
                }
            }'

            $iCount += 1
        }
    }
    #Set route table template footer JSON
    $sOutput += ']}'
    
    #Each vnet peer host has it's own template thus requiring the write action here.
    New-Item -ItemType Directory -Force -Path $sPEEROutPath | Out-Null
    $sOutJsonName = "$($sPEEROutPath)ARM-$($sVNETName)-PEER.json"
    $sOutput | Out-File $sOutJsonName -Force

    return ($iCount - 1)
}
Function parseFWs() {
    param ($curFW)
    #Set firewall template header JSON
    $sOutput = '{
        "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentTemplate.json#",
        "contentVersion": "1.0.0.0",
        "parameters": {},
        "variables": {},
        "resources": [
        {
            "apiVersion": "2018-10-01",
            "type": "Microsoft.Network/publicIPAddresses",
            "name": "' + $curFW.PUBLICIPNAME + '",
            "location": "' + $curFW.LOCATION + '",
            "sku": {
                "name": "Standard"
            },
            "properties": {
              "publicIPAllocationMethod": "Static",
              "publicIPAddressVersion": "IPv4"
            }
          },
          {
            "apiVersion": "2019-04-01",
            "type": "Microsoft.Network/azureFirewalls",
            "name": "' + $curFW.FIREWALLNAME + '",
            "location": "' + $curFW.LOCATION + '",
            "zones": [' + $curFW.AVAILABILITYZONES + '],        
            "dependsOn": [
            "[resourceId(''Microsoft.Network/publicIPAddresses/'', ''' + $curFW.PUBLICIPNAME + ''')]"
            ],
            "properties": {
                "ipConfigurations": [
                {
                  "name": "IpConf",
                  "properties" : {
                    "subnet": {
                      "id": "[resourceId(''Microsoft.Network/virtualNetworks/subnets'',''' + $curFW.VIRTUALNETWORKNAME + ''', ''AZUREFIREWALLSUBNET'')]"
                    },
                    "PublicIPAddress": {
                      "id": "[resourceId(''Microsoft.Network/publicIPAddresses'',''' + $curFW.PUBLICIPNAME + ''')]"
                    }
                  }
                }
              ]'
    
    #LOOP THROUGH 3 TYPES OF RULES: applicationRuleCollections networkRuleCollections natRuleCollections
    $sOutput += ',"applicationRuleCollections": ['
    $iCount = 0
    foreach ($curFWRULE in $sFWAPP) {
        #Check if current rule is name matched. (Ex. CORE-CACN-PRIM-AZF)
        if ($curFWRULE.FIREWALLNAME -eq $curFW.FIREWALLNAME) {
            if ($iCount -gt 0) {
                $sOutput += ',' #Add comma prefix if not 1st object.
            }
            #Generate firewall rule template JSON.
            $sOutput += '
            {
              "name": "' + $curFWRULE.NAME + '",
              "properties": {
                "priority": ' + $curFWRULE.PRIORITY + ',
                "action": { 
                    "type": "' + $curFWRULE.ACTION + '" 
                },
                "rules": [
                  {
                    "name": "' + $curFWRULE.RULENAME + '",
                    "description": "' + $curFWRULE.DESCRIPTION + '",
                    "sourceAddresses": [ ' + $curFWRULE.SOURCEADDRESSES + '],
                    "protocols": [{' + $curFWRULE.PROTOCOLS + '}],
                    "targetFqdns": [' + $curFWRULE.TARGETFQDNS + '],
                    "fqdnTags": [' + $curFWRULE.FQDNTAGS + ']
                  }
                ]
              }
            }'
            $iCount += 1
        }
    }
    $sOutput += ']
    ,"natRuleCollections": ['
    $iCount = 0
    foreach ($curFWRULE in $sFWNAT) {
        #Check if current rule is name matched. (Ex. CORE-CACN-PRIM-AZF)
        if ($curFWRULE.FIREWALLNAME -eq $curFW.FIREWALLNAME) {
            if ($iCount -gt 0) {
                $sOutput += ',' #Add comma prefix if not 1st object.
            }
            #Generate firewall rule template JSON.
            $sOutput += '
            {
              "name": "' + $curFWRULE.NAME + '",
              "properties": {
                "priority": ' + $curFWRULE.PRIORITY + ',
                "action": { 
                    "type": "' + $curFWRULE.ACTION + '" 
                },
                "rules": [
                  {
                    "name": "' + $curFWRULE.RULENAME + '",
                    "description": "' + $curFWRULE.DESCRIPTION + '",
                    "sourceAddresses": [' + $curFWRULE.SOURCEADDRESSES + '],
                    "destinationAddresses": [' + $curFWRULE.DESTINATIONADDRESSES + '],
                    "destinationPorts": [' + $curFWRULE.DESTINATIONPORTS + '],
                    "protocols": [' + $curFWRULE.PROTOCOLS + '],
                    "translatedAddress": ["' + $curFWRULE.TRANSLATEDADDRESS + '"],
                    "translatedPort": ["' + $curFWRULE.TRANSLATEDPORT + '"]
                  }
                ]
              }
            }'
            $iCount += 1
        }
    }
    $sOutput += ']
    ,"networkRuleCollections": ['
    $iCount = 0
    foreach ($curFWRULE in $sFWNET) {
        #Check if current rule is name matched. (Ex. CORE-CACN-PRIM-AZF)
        if ($curFWRULE.FIREWALLNAME -eq $curFW.FIREWALLNAME) {
            if ($iCount -gt 0) {
                $sOutput += ',' #Add comma prefix if not 1st object.
            }
            #Generate firewall rule template JSON.
            $sOutput += '
            {
              "name": "' + $curFWRULE.NAME + '",
              "properties": {
                "priority": ' + $curFWRULE.PRIORITY + ',
                "action": { 
                    "type": "' + $curFWRULE.ACTION + '" 
                },
                "rules": [
                  {
                    "name": "' + $curFWRULE.RULENAME + '",
                    "description": "' + $curFWRULE.DESCRIPTION + '",
                    "sourceAddresses": [' + $curFWRULE.SOURCEADDRESSES + '],
                    "destinationAddresses": [' + $curFWRULE.DESTINATIONADDRESSES + '],
                    "destinationPorts": [' + $curFWRULE.DESTINATIONPORTS + '],
                    "protocols": [' + $curFWRULE.PROTOCOLS + ']
                  }
                ]
              }
            }'
            $iCount += 1
        }
    }
    $sOutput += ']'
    #Set firewall template footer JSON
    $sOutput += '}}]}'
    return $sOutput
}
Function GetSubscriptionID {
    param ($sSubName)
    #Retrieve the Subscription ID from the SUB worksheet based on the Subscription name.
    $sSUBID = "ERROR"
    foreach ($curSub in $sSUB) {
        if ($curSub.NAME -eq $sSubName) {
            $sSUBID = $curSUB.ID
        }
    } 
    return $sSUBID
}

#endregion Functions

#region Import-Excel
#Subscriptions
$sSUB = Import-Excel $filename "SUB"
"$($sSUB.Count) subscription(s) found."

#Network Security Groups
$sNSG = Import-Excel $filename "NSG" -DataOnly -ErrorAction SilentlyContinue
"$($sNSG.Count) network security group(s) found."

#Route Tables
$sRT = Import-Excel $filename "RT" -DataOnly -ErrorAction SilentlyContinue
"$($sRT.Count) route table(s) found."

#Virtual Networks
$sVNET = Import-Excel $filename "VNET" -DataOnly -ErrorAction SilentlyContinue
"$($sVNET.Count) virtual network(s) found."

#Firewalls
$sFW = Import-Excel $filename "FW" -DataOnly -ErrorAction SilentlyContinue
"$($sFW.Count) firewall(s) found."

#Application Gateways
$sAG = Import-Excel $filename "AG" -DataOnly -ErrorAction SilentlyContinue
"$($sAG.Count) application gateway(s) found."

#Network Security Group Rule Sets
$sRULESET = Import-Excel $filename "NSGRULES" -DataOnly -ErrorAction SilentlyContinue
"$($sRULESET.Count) rule(s) found."

#Route Table Entries
$sROUTESET = Import-Excel $filename "ROUTES" -DataOnly -ErrorAction SilentlyContinue
"$($sROUTESET.Count) route(s) found."

#Virtual Network Subnets
$sSUBNET = Import-Excel $filename "SUBNET" -DataOnly -ErrorAction SilentlyContinue
"$($sSUBNET.Count) subnet(s) found."

#Firewall App Rules
$sFWAPP = Import-Excel $filename "FWAPP" -DataOnly -ErrorAction SilentlyContinue
"$($sFWAPP.Count) firewall app rule(s) found."

#Firewall Nat Rules
$sFWNAT = Import-Excel $filename "FWNAT" -DataOnly -ErrorAction SilentlyContinue
"$($sFWNAT.Count) firewall nat rule(s) found."

#Firewall Network Rules
$sFWNET = Import-Excel $filename "FWNET" -DataOnly -ErrorAction SilentlyContinue
"$($sFWNET.Count) firewall network rule(s) found."

#Virtual Network Peerings
$sPEER = Import-Excel $filename "PEER" -DataOnly -ErrorAction SilentlyContinue
"$($sPEER.Count) virtual network peer(s) found."
#endregion Import-Excel

foreach ($curSUB in $sSUB) {
    Write-Output "Processing $($curSUB.NAME) $($curSUB.ID)"

    $iNSG = 0
    foreach ($curNSG in $sNSG) {
        if ($curNSG.SUB -eq $curSUB.NAME) {
            # EXCELHEADERS: SUB REGION LOCATION NSGNAME RESOURCEGROUPNAME RULESET
            New-Item -ItemType Directory -Force -Path $sNSGOutPath | Out-Null
            $sOutJsonName = "$($sNSGOutPath)ARM-$($curNSG.NSGNAME).json"
            $sNSGOutput = parseNSGs($curNSG)
            $sNSGOutput | Out-File $sOutJsonName -Force
            $iNSG += 1
        }
    }
    Write-Output "Generated $($iNSG) network security group templates" 

    $iRT = 0
    foreach ($curRT in $sRT) {
        if ($curRT.SUB -eq $curSUB.NAME) {
            # EXCELHEADERS: SUB REGION LOCATION ROUTETABLENAME RESOURCEGROUPNAME
            New-Item -ItemType Directory -Force -Path $sRTOutPath | Out-Null
            $sOutJsonName = "$($sRTOutPath)ARM-$($curRT.ROUTETABLENAME).json"
            $sRTOutput = parseRTs($curRT)
            $sRTOutput | Out-File $sOutJsonName -Force
            $iRT += 1
        }
    }
    Write-Output "Generated $($iRT) route table templates"

    $iVNET = 0
    $iPEERs = 0
    foreach ($curVNET in $sVNET) {
        if ($curVNET.SUB -eq $curSUB.NAME) {
            # EXCELHEADERS: SUB REGION LOCATION VIRTUALNETWORKNAME RESOURCEGROUPNAME ROUTETABLENAME NSGNAME
            New-Item -ItemType Directory -Force -Path $sVNETOutPath | Out-Null
            $sOutJsonName = "$($sVNETOutPath)ARM-$($curVNET.VIRTUALNETWORKNAME).json"
            $sVNETOutput = parseVNETs($curVNET)
            $sVNETOutput | Out-File $sOutJsonName -Force
            
            $iPEERs += parsePEERs $curVNET $curSUB
            $iVNET += 1
        }
    }
    Write-Output "Generated $($iVNET) virtual network templates"
    Write-Output "Generated $($iPEERs) virtual network peering templates"

    $iFW = 0
    foreach ($curFW in $sFW) {
        if ($curFW.SUB -eq $curSUB.NAME) {
            # EXCELHEADERS: SUB REGION LOCATION FIREWALLNAME RESOURCEGROUPNAME VIRTUALNETWORKNAME PUBLICIPNAME IPADDRESS
            New-Item -ItemType Directory -Force -Path $sFWOutPath | Out-Null
            $sOutJsonName = "$($sFWOutPath)ARM-$($curFW.FIREWALLNAME).json"
            $sFWOutput = parseFWs($curFW)
            $sFWOutput | Out-File $sOutJsonName -Force
            $iFW += 1
        }
    }
    Write-Output "Generated $($iFW) firewall templates"

    foreach ($curAG in $sAG) {
        if ($curAG.SUB -eq $curSUB.NAME) {
            # EXCELHEADERS: SUB REGION LOCATION APPLICATIONGATEWAYNAME RESOURCEGROUPNAME VIRTUALNETWORKNAME SUBNETNAME
            
        }
    }
}
    
exit