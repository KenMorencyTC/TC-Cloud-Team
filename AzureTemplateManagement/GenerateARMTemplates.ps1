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

    TODOs:
    1. Add empty value handling procedure(s).
    2. Add support for VM/Service resources.
    3. Refactor Excel imports for efficiency.
    4. Refactor Sheet processing for efficiency.
    5. Complete integration of Core Network Infrastructure ASR/Backup Policies and Log Auditing Workspaces.
.EXAMPLE
    .PS> .\GenerateTemplates.ps1 -filename "\NetworkInfrastructure.xlsx" -outputfolder "\TEMPLATES\"
.EXAMPLE
    .PS> .\GenerateTemplates.ps1 "\NetworkInfrastructure.xlsx" "\TEMPLATES\"
.LINK
    https://github.com/KenMorencyTC/TC-Cloud-Team-Automation-Scripting
#>

param (
    [String] $filename = "$($PSScriptRoot)\AzureNetworkInfrastructure.xlsx",
    [String] $outputfolder = "$($PSScriptRoot)\ARM\"
) 

#region INIT

$sScriptVersion = "1.0"
$sScriptName = "GenerateARMTemplates"

#$sPOLOutPath = "$($outputfolder)00-POL\"
$sLAWOutPath = "$($outputfolder)01-LAW\"
$sRGOutPath = "$($outputfolder)02-RG\"
$sNSGOutPath = "$($outputfolder)03-NSG\"
$sRTOutPath = "$($outputfolder)04-RT\"
$sVNETOutPath = "$($outputfolder)05-VNET\"
$sPEEROutPath = "$($outputfolder)06-PEER\"
$sAGOutPath = "$($outputfolder)07-AG\"
$sFWOutPath = "$($outputfolder)08-FW\"
$sSAOutPath = "$($outputfolder)09-SA\"
$sRSVOutPath = "$($outputfolder)10-RSV\"
#$sBASTOutPath = "$($outputfolder)12-BAST\"
#$sKVOutPath = "$($outputfolder)13-KV\"
#$sDDOSOutPath = "$($outputfolder)14-DDOS\"

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
            $sREMOTESUBID = GetSubscriptionID($curPEER.REMOTESUB)
            #Write-Output $sPEERName' '$sVNETName
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
                    "id": "[resourceId(''' + $sREMOTESUBID + ''', ''' + $curPEER.REMOTERESOURCEGROUPNAME + 
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
    #$sOutJsonName = "$($sPEEROutPath)ARM-$($sVNETName)-PEER.json"
    $sOutFileName = "ARM-$($sVNETName)-PEER.json"
    $sOutJsonName = $($sPEEROutPath) + $sOutFileName
    $sOutPSName = "$($sPEEROutPath)RUN-$($sVNETName)-PEER.ps1"
    $sOutput | Out-File $sOutJsonName -Force
    $sPSOutput = BuildPowerShell $curVNET "$($sVNETName) Peering" $sOutFileName
    $sPSOutput | Out-File $sOutPSName -Force

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
Function parseSAs() {
    param ($curSA)
    #Set storage account template header JSON
    $sOutput = '{
        "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentTemplate.json#",
        "contentVersion": "1.0.0.0",
        "parameters": {},
        "variables": {},
        "resources": [
          {
            "type": "Microsoft.Storage/storageAccounts",
            "name": "' + $curSA.STORAGEACCOUNTNAME + '",
            "location": "' + $curSA.LOCATION + '",
            "apiVersion": "2018-07-01",
            "sku": {
              "name": "' + $curSA.SKU + '"
            },
            "kind": "' + $curSA.KIND + '",
            "properties": {}
          }
        ],
        "outputs": {
          "storageAccountName": {
            "type": "string",
            "value": "' + $curSA.STORAGEACCOUNTNAME + '"
          }
        }
      }'
    return $sOutput
}
Function parseAGs() {
    param ($curAG)
    #"resourceGroup": "' + $curFW.RESOURCEGROUPNAME + '",
    #"subscriptionId": "' + $sSUBID + '",
    #Set application gateway template header JSON
    foreach ($curLAW in $sLAW) {
      if ($curLAW.OMSWORKSPACENAME -eq $curAG.LAWNAME) {
        $sLAWOutput = $curLAW.RESOURCEGROUPNAME
      }
    }

    $sOutput = '{
        "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentTemplate.json#",
        "contentVersion": "1.0.0.0",
        "parameters": {
        },
        "variables": {
          "subnetRef": "[resourceId(''Microsoft.Network/virtualNetworks/subnets'', ''' + $curAG.VIRTUALNETWORKNAME + ''', ''' + $curAG.SUBNETNAME + ''')]"
        },
        "resources": [
          {
            "apiVersion":"2018-08-01",
            "type":"Microsoft.Network/publicIPAddresses",
            "name":"' + $curAG.PUBLICIPNAME + '",
            "location":"' + $curAG.LOCATION + '",
            "sku":{
               "name":"Standard"
            },
            "properties":{
               "publicIPAllocationMethod":"Static"
            }
         },
               {
            "type": "Microsoft.Network/publicIPAddresses/providers/diagnosticSettings",
            "name": "[concat(''' + $curAG.PUBLICIPNAME + ''', ''/microsoft.insights/'', ''' + $curAG.LAWNAME + ''')]",
            "apiVersion": "2017-05-01-preview",
            "properties": {
              "name": "[concat(''' + $curAG.PUBLICIPNAME + ''', ''/microsoft.insights/'', ''' + $curAG.LAWNAME + ''')]",
              "workspaceId": "[resourceId(''' + $sLAWOutput + ''',''microsoft.operationalinsights/workspaces/'', ''' + $curAG.LAWNAME + ''')]",
              "logs": [
                {
                  "category": "DDoSProtectionNotifications",
                  "enabled": true,
                  "retentionPolicy": {
                    "days": 0,
                    "enabled": false
                  }
                },
                {
                  "category": "DDoSMitigationReports",
                  "enabled": true,
                  "retentionPolicy": {
                    "days": 0,
                    "enabled": false
                  }
                }
              ],
              "metrics": [
                {
                  "category": "AllMetrics",
                  "enabled": true,
                  "retentionPolicy": {
                    "days": 0,
                    "enabled": false
                  }
                }
              ]
            },
            "dependsOn": [
              "[resourceId(''Microsoft.Network/publicIPAddresses'', ''' + $curAG.PUBLICIPNAME + ''')]"
            ]
          },
          {
            "apiVersion":"2018-08-01",
            "name":"' + $curAG.APPLICATIONGATEWAYNAME + '",
            "type":"Microsoft.Network/applicationGateways",
            "location":"' + $curAG.LOCATION + '",
            "dependsOn":[
              "' + $curAG.PUBLICIPNAME + '"
            ],
            "properties":{
               "sku":{
                  "name":"Standard_v2",
                  "tier":"Standard_v2"
               },
               "autoscaleConfiguration":{
                  "minCapacity":2
               },
               "gatewayIPConfigurations":[
                  {
                     "name":"appGatewayIpConfig",
                     "properties":{
                        "subnet":{
                           "id":"[variables(''subnetRef'')]"
                        }
                     }
                  }
               ],
               "frontendIPConfigurations":[
                {
                   "name":"appGatewayFrontendIP",
                   "properties":{
                      "PublicIPAddress":{
                         "id":"[resourceid(''Microsoft.Network/publicIPAddresses'', ''' + $curAG.PUBLICIPNAME + ''')]"
                      }
                   }
                }
             ],
               "frontendPorts":[
                  {
                     "name":"appGatewayFrontendPort",
                     "properties":{
                        "Port":80
                     }
                  }
               ],
               "backendAddressPools":[
                  {
                     "name":"appGatewayBackendPool",
                     "properties":{
                        "BackendAddresses":[{
                          "IpAddress": "10.0.0.4"
                        },
                        {
                          "IpAddress": "10.0.0.5"
                        }]
                     }
                  }
               ],
               "backendHttpSettingsCollection":[
                  {
                     "name":"appGatewayBackendHttpSettings",
                     "properties":{
                        "Port":80,
                        "Protocol":"Http",
                        "CookieBasedAffinity":"Disabled"
                     }
                  }
               ],
               "httpListeners":[
                  {
                     "name":"appGatewayHttpListener",
                     "properties":{
                        "FrontendIpConfiguration":{
                           "Id":"[resourceId(''Microsoft.Network/applicationGateways/frontendIPConfigurations'', ''' + $curAG.APPLICATIONGATEWAYNAME + ''', ''appGatewayFrontendIP'')]"
                        },
                        "FrontendPort":{
                           "Id":"[resourceId(''Microsoft.Network/applicationGateways/frontendPorts'', ''' + $curAG.APPLICATIONGATEWAYNAME + ''', ''appGatewayFrontendPort'')]"
                        },
                        "Protocol":"Http"
                     }
                  }
               ],
               "requestRoutingRules":[
                  {
                     "Name":"rule1",
                   "properties": {
                     "RuleType": "Basic",
                     "httpListener": {
                       "id": "[resourceId(''Microsoft.Network/applicationGateways/httpListeners'', ''' + $curAG.APPLICATIONGATEWAYNAME + ''', ''appGatewayHttpListener'')]"
                     },
                     "backendAddressPool": {
                       "id": "[resourceId(''Microsoft.Network/applicationGateways/backendAddressPools'', ''' + $curAG.APPLICATIONGATEWAYNAME + ''', ''appGatewayBackendPool'')]"
                     },
                     "backendHttpSettings": {
                       "id": "[resourceId(''Microsoft.Network/applicationGateways/backendHttpSettingsCollection'', ''' + $curAG.APPLICATIONGATEWAYNAME + ''', ''appGatewayBackendHttpSettings'')]"
                     }
                   }
                  }
               ]
            }
         },
         {
          "apiVersion": "2017-05-01-preview",
          "type": "Microsoft.Network/applicationgateways/providers/diagnosticSettings",
          "name": "[concat(''' + $curAG.APPLICATIONGATEWAYNAME + ''', ''/microsoft.insights/'', ''' + $curAG.LAWNAME + ''')]",
          "properties": {
            "workspaceId": "[resourceId(''' + $sLAWOutput + ''', ''microsoft.operationalinsights/workspaces/'', ''' + $curAG.LAWNAME + ''')]",
            "logs": [
              {
                "category": "ApplicationGatewayAccessLog",
                "enabled": true,
                "retentionPolicy": {
                  "days": 0,
                  "enabled": false
                }
              },
              {
                "category": "ApplicationGatewayPerformanceLog",
                "enabled": true,
                "retentionPolicy": {
                  "days": 0,
                  "enabled": false
                }
              },
              {
                "category": "ApplicationGatewayFirewallLog",
                "enabled": true,
                "retentionPolicy": {
                  "days": 0,
                  "enabled": false
                }
              }           
            ],
            "metrics": [
              {
                "category": "AllMetrics",
                "enabled": true,
                "retentionPolicy": {
                  "days": 0,
                  "enabled": false
                }
              }
            ]
          }
        }
        ]
      }'

    return $sOutput
}
Function parseRSVs() {
    param ($curRSV) 
    $sOutput = '
{
  "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentTemplate.json#",
  "contentVersion": "1.0.0.0",
  "parameters": {},
  "variables": {},
  "resources": [
    {
      "type": "Microsoft.RecoveryServices/vaults",
      "apiVersion": "2018-01-10",
      "name": "' + $curRSV.VAULTNAME + '",
      "location": "' + $curRSV.LOCATION + '",
      "sku": {' + $curRSV.SKU + '},
      "properties": {}
    },'
        #Write-Output $curRSV.ASR
        if ($curRSV.ASR) {
          $iASRCount = 0
          $sOutput += '
  {
    "type": "Microsoft.RecoveryServices/vaults/providers/diagnosticSettings",
    "name": "[concat(''' + $curRSV.VAULTNAME + ''', ''/microsoft.insights/'', ''' + $curRSV.LAWNAME + ''')]",
    "apiVersion": "2017-05-01-preview",
    "properties": {
      "name": "[concat(''' + $curRSV.VAULTNAME + ''', ''/microsoft.insights/'', ''' + $curRSV.LAWNAME + ''')]",          
      "workspaceId": "[resourceId(''' + $curRSV.RESOURCEGROUPNAME + ''',''microsoft.operationalinsights/workspaces/'', ''' + $curRSV.LAWNAME + ''')]",
      "logs": ['
          $tASR = $curRSV.ASR | Out-String
          $aASR = $tASR.Split(",")
          foreach ($curASR in $aASR) {
            #Write-Output $curASR.ID
            #$bUseASR = $false
            foreach ($cASR in $sASR) {
              #Write-Output "$($cASR) - $($curASR.POLICYID)"
              $bUseASR = $false
              if ($cASR.POLICYID -eq $curASR) {
                  $bUseASR = $true
              }
              #Write-Output $bUseASR
              if ($bUseASR -eq $true) {
                if ($iASRCount -gt 0) {
                    $sOutput += ','
                }
                $sOutput += '
        {
          "category": "' + $cASR.CATEGORY + '",
          "enabled": ' + $cASR.ENABLED + ',
          "retentionPolicy": {' + $cASR.RETENTIONPOLICY + '}
        }'
                $iASRCount += 1
              }
            }
          }
          $sOutput += '
      ],
      "metrics": [' + $curRSV.METRICS + ']
          },
          "dependsOn": [
              "[resourceId(''Microsoft.RecoveryServices/vaults/'', ''' + $curRSV.VAULTNAME + ''')]"
          ]
    }'
        }
        if ($curRSV.BCK) {
          if ($iASRCount -gt 0) {$sOutput += ","}
          $iBCKCount = 0
          $sOutput += '
    {
      "type": "Microsoft.RecoveryServices/vaults/backupPolicies",
      "name": "[concat(''' + $curRSV.VAULTNAME + ''', ''' + $curRSV.POLICYNAME + ''')]",
      "apiVersion": "2016-06-01",
      "location": "' + $curRSV.LOCATION + '",
      "dependsOn": [
      "[concat(''Microsoft.RecoveryServices/vaults/'', ''' + $curRSV.VAULTNAME + ''')]"
      ],
      "properties": {'
          $tBCK = $curRSV.BCK | Out-String
          $aBCK = $tBCK.Split(",")
          foreach ($curBCK in $aBCK) {
            $bUseBCK = $false
            foreach ($cBCK in $sBCK) {
              #Write-Host $curBCK"-"$cBCK.POLICYID
              if ($cBCK.POLICYID -eq $curBCK) {
                  $bUseBCK = $true
                  if ($bUseBCK -eq $true) {
                    if ($iBCKCount -gt 0) {
                        $sOutput += ','
                    }
                    $sOutput += '
        "backupManagementType": "AzureIaasVM",
        "instantRpRetentionRangeInDays": "' + $cBCK.INSTANTRETDAYS + '",
        "schedulePolicy": {
          "scheduleRunFrequency": "Daily",
          "scheduleRunDays": null,
          "scheduleRunTimes": [' + $cBCK.SCHEDULERUNTIMES + '],
          "schedulePolicyType": "SimpleSchedulePolicy"
        },
        "retentionPolicy": {
          "dailySchedule": {
            "retentionTimes": [' + $cBCK.SCHEDULERUNTIMES + '],
            "retentionDuration": {
              "count": "' + $cBCK.RETDAILY + '",
              "durationType": "Days"
            }
          },
          "weeklySchedule": {
            "daysOfTheWeek": [' + $cBCK.DAYS + '],
            "retentionTimes": [' + $cBCK.SCHEDULERUNTIMES + '],
            "retentionDuration": {
              "count": "' + $cBCK.RETWEEKLY + '",
              "durationType": "Weeks"
            }
          },
          "monthlySchedule": {
            "retentionScheduleFormatType": "Daily",
            "retentionScheduleDaily": {
              "daysOfTheMonth": [
                {
                  "date": 1,
                  "isLast": false
                }
              ]
            },
            "retentionScheduleWeekly": null,
            "retentionTimes": [' + $cBCK.SCHEDULERUNTIMES + '],
            "retentionDuration": {
              "count": "' + $cBCK.RETMONTHLY + '",
              "durationType": "Months"
            }
          },
          "yearlySchedule": {
            "retentionScheduleFormatType": "Daily",
            "monthsOfYear": [' + $cBCK.MONTHS + '],
            "retentionScheduleDaily": {
              "daysOfTheMonth": [
                {
                  "date": 1,
                  "isLast": false
                }
              ]
            },
            "retentionScheduleWeekly": null,
            "retentionTimes": [' + $cBCK.SCHEDULERUNTIMES + '],
            "retentionDuration": {
              "count": "' + $cBCK.RETYEARLY + '",
              "durationType": "Years"
            }
          },
          "retentionPolicyType": "LongTermRetentionPolicy"
        }
      }
    }'
                    $iBCKCount += 1
              }
            }
            }
          }
        }

        $sOutput += '
    ]
  }'
        return $sOutput
}
Function parseRGs() {
  param ($curRG)
  #Set resource group template header JSON
  $sOutput = '{
    "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentTemplate.json#",
    "contentVersion": "1.0.0.1",
    "parameters": {},
    "variables": {},
    "resources": [
        {
            "apiVersion": "2018-05-01",
            "type": "Microsoft.Resources/resourceGroups",
            "location": "' + $curRG.LOCATION + '",
            "name": "' + $curRG.RESOURCEGROUPNAME + '",
            "properties": {},
            "tags": {
                "Solution-name": "' + $curRG.TSOLUTIONNAME + '",
                "Environment": "' + $curRG.SUB + '",
                "Solution-version": "' + $curRG.TSOLUTIONVERSION + '",
                "Cost-Center": "' + $curRG.TCOSTCENTER + '",
                "Business-Unit": "' + $curRG.TBUSINESSUNIT + '",
                "Project-Name": "' + $curRG.TPROJECTNAME + '",
                "Tech-Owner": "' + $curRG.TTECHOWNER + '",
                "Exp-Date": "' + $curRG.TEXPDATE + '",
                "Sensitivity": "' + $curRG.TSENSITIVITY + '",
                "BusinessValueRating": "' + $curRG.TVALUERATING + '",
                "TechContact": "' + $curRG.TTECHCONTACT + '"
            }
        },
        {
            "type": "Microsoft.Authorization/policyAssignments",
            "name": "' + $curRG.RESOURCEGROUPNAME + ' Tag Policy Assignment v2",
            "location": "CandaCentral",
            "apiVersion": "2018-05-01",
            "dependsOn": [
                "[resourceId(''Microsoft.Resources/resourceGroups/'', ''' + $curRG.RESOURCEGROUPNAME + ''')]"
            ],
            "properties": {
                "scope": "[concat(subscription().id, ''/resourceGroups/'', ''' + $curRG.RESOURCEGROUPNAME + ''')]",
                "policyDefinitionId": "[resourceId(''Microsoft.Authorization/policyDefinitions'', ''Tag Resources in core'')]",
                "Solution-name": "' + $curRG.TSOLUTIONNAME + '",
                "Environment": "' + $curRG.SUB + '",
                "Solution-version": "' + $curRG.TSOLUTIONVERSION + '",
                "Cost-Center": "' + $curRG.TCOSTCENTER + '",
                "Business-Unit": "' + $curRG.TBUSINESSUNIT + '",
                "Project-Name": "' + $curRG.TPROJECTNAME + '",
                "Tech-Owner": "' + $curRG.TTECHOWNER + '",
                "Exp-Date": "' + $curRG.TEXPDATE + '",
                "Sensitivity": "' + $curRG.TSENSITIVITY + '",
                "BusinessValueRating": "' + $curRG.TVALUERATING + '",
                "TechContact": "' + $curRG.TTECHCONTACT + '"
            }
        },
        {
            "type": "Microsoft.Authorization/policyAssignments",
            "name": "' + $curRG.RESOURCEGROUPNAME + ' PBMM Policy for GoC",
            "location": "' + $curRG.LOCATION + '",
            "apiVersion": "2018-05-01",
            "dependsOn": [
                "[resourceId(''Microsoft.Resources/resourceGroups/'', ''' + $curRG.RESOURCEGROUPNAME + ''')]"
            ],
            "properties": {
                "scope": "[concat(subscription().id, ''/resourceGroups/'', ''' + $curRG.RESOURCEGROUPNAME + ''')]",
                "policyDefinitionId": "/providers/Microsoft.Authorization/policySetDefinitions/89c6cddc-1c73-4ac1-b19c-54d1a15a42f2"
                "assignIdentity": ""
            }
        },
        {
            "type": "Microsoft.Authorization/policyAssignments",
            "name": "' + $curRG.RESOURCEGROUPNAME + ' Custom Ops Policy",
            "location": "' + $curRG.LOCATION + '",
            "apiVersion": "2018-05-01",
            "dependsOn": [
                "[resourceId(''Microsoft.Resources/resourceGroups/'', ''' + $curRG.RESOURCEGROUPNAME + ''')]"
            ],
            "properties": {
                "scope": "[concat(subscription().id, ''/resourceGroups/'', ''' + $curRG.RESOURCEGROUPNAME + ''')]",
                "policyDefinitionId": "[resourceId(''Microsoft.Authorization/policyDefinitions'', ''GC PBMM Policy Set'')]",
                "loganalytics": "/subscriptions/2a4beb17-e62e-4ded-9455-be61e92329e1/resourcegroups/core01-ops01-rgp/providers/microsoft.operationalinsights/workspaces/core01-law-main"
            }
        }
    ],
    "outputs": {}
}'
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
function Format-Json([Parameter(Mandatory, ValueFromPipeline)][String] $json) {
  $indent = 0;
  ($json -Split '\n' |
    % {
      if ($_ -match '[\}\]]') {
        # This line contains  ] or }, decrement the indentation level
        $indent--
      }
      $line = (' ' * $indent * 2) + $_.TrimStart().Replace(':  ', ': ')
      if ($_ -match '[\{\[]') {
        # This line contains [ or {, increment the indentation level
        $indent++
      }
      $line
  }) -Join "`n"
}
Function BuildPowerShell {
  param ($curObj,$curName,$curFile,$curArgs)
  $sSubID = GetSubscriptionID($curObj.SUB)
  $sOutput = 'Set-AzContext -Subscriptionid "' + $sSubID + '"

  Write-Output "Deploying ' + $curName + ' in ' + $curObj.RESOURCEGROUPNAME + '"
  New-AzResourceGroupDeployment -ResourceGroupName "' + $curObj.RESOURCEGROUPNAME + '" -TemplateFile "$($PSScriptRoot)\' + $curFile + '" '

  if ($curArgs) {$sOutput += $curArgs}

  return $sOutput
}
Function BuildRGPowerShell {
  param ($curObj,$curName,$curFile,$curArgs)
  $sSubID = GetSubscriptionID($curObj.SUB)
  $sOutput = 'Set-AzContext -Subscriptionid "' + $sSubID + '"

  Write-Output "Deploying ' + $curName + ' in ' + $curObj.SUB + '"
  New-AzDeployment -Name "' + $curObj.RESOURCEGROUPNAME + '" -Location "' + $curObj.LOCATION + '" -TemplateFile "$($PSScriptRoot)\' + $curFile + '" '

  if ($curArgs) {$sOutput += $curArgs}

  return $sOutput
}

#endregion Functions

#region Import-Excel
#Subscriptions
$sSUB = Import-Excel $filename "SUB"
"$($sSUB.Count) subscription(s) found."

#Log Analytics Workspaces
$sLAW = Import-Excel $filename "LAW" -DataOnly -ErrorAction SilentlyContinue
"$($sLAW.Count) log analytics workspace(s) found."

#Resource Groups
$sRG = Import-Excel $filename "RG" -DataOnly -ErrorAction SilentlyContinue
"$($sRG.Count) resource group(s) found."

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

#Storage Accounts
$sSA = Import-Excel $filename "SA" -DataOnly -ErrorAction SilentlyContinue
"$($sSA.Count) storage account(s) found."

#Application Gateways
$sAG = Import-Excel $filename "AG" -DataOnly -ErrorAction SilentlyContinue
"$($sAG.Count) application gateway(s) found."

#Recovery Service Vaults
$sRSV = Import-Excel $filename "RSV" -DataOnly -ErrorAction SilentlyContinue
"$($sRSV.Count) recovery service vault(s) found."

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

#Backup Profile
$sBCK = Import-Excel $filename "BCK" -DataOnly -ErrorAction SilentlyContinue
"$($sBCK.Count) backup profile(s) found."

#Azure Site Recovery
$sASR = Import-Excel $filename "ASR" -DataOnly -ErrorAction SilentlyContinue
"$($sASR.Count) azure site recovery profile(s) found."


#endregion Import-Excel

foreach ($curSUB in $sSUB) {
    Write-Output "Processing $($curSUB.NAME) $($curSUB.ID)"

    $iLAW = 0
    foreach ($curLAW in $sLAW) {
        if ($curLAW.SUB -eq $curSUB.NAME) {
            # EXCELHEADERS: SUB REGION LOCATION OMSWORKSPACENAME RESOURCEGROUPNAME SKU
            New-Item -ItemType Directory -Force -Path $sLAWOutPath | Out-Null
            $sOutFileName = "ARM-$($curLAW.OMSWORKSPACENAME).json"
            $sOutJsonName = $($sLAWOutPath) + $sOutFileName
            $sOutPSName = "$($sLAWOutPath)RUN-$($curLAW.OMSWORKSPACENAME).ps1"
            $sLAWOutput = Get-Content -Path "$($PSScriptRoot)\resources\ARM-LAW.json"
            $sLAWOutput | Out-File $sOutJsonName -Force
            $sArgs = '-omsworkspacename "' + $curLAW.OMSWORKSPACENAME + '" -omsworkspaceregion "' + $curLAW.OMSWORKSPACENAME + '"'
            $sLAWPSOutput = BuildPowerShell $curLAW $curLAW.OMSWORKSPACENAME $sOutFileName $sArgs
            $sLAWPSOutput | Out-File $sOutPSName -Force
            $iLAW += 1
        }
    }
    Write-Output "Generated $($iLAW) log analytics workspace templates"

    $iRG = 0
    foreach ($curRG in $sRG) {
        if ($curRG.SUB -eq $curSUB.NAME) {
            # EXCELHEADERS: SUB REGION LOCATION RESOURCEGROUPNAME TSOLUTIONNAME TSOLUTIONVERSION TCOSTCENTER TBUSINESSUNIT TPROJECTNAME TTECHOWNER TSENSITIVITY TVALUERATING TTECHCONTACT LAWNAME
            New-Item -ItemType Directory -Force -Path $sRGOutPath | Out-Null
            $sOutFileName = "ARM-$($curRG.RESOURCEGROUPNAME).json"
            $sOutJsonName = $($sRGOutPath) + $sOutFileName
            $sOutPSName = "$($sRGOutPath)RUN-$($curRG.RESOURCEGROUPNAME).ps1"
            $sRGOutput = parseRGs($curRG)
            $sRGOutput | Out-File $sOutJsonName -Force
            $sRGPSOutput = BuildRGPowerShell $curRG $curRG.RESOURCEGROUPNAME $sOutFileName
            $sRGPSOutput | Out-File $sOutPSName -Force
            $iRG += 1
        }
    }
    Write-Output "Generated $($iRG) resource group templates"

    $iNSG = 0
    foreach ($curNSG in $sNSG) {
        if ($curNSG.SUB -eq $curSUB.NAME) {
            # EXCELHEADERS: SUB REGION LOCATION NSGNAME RESOURCEGROUPNAME RULESET
            New-Item -ItemType Directory -Force -Path $sNSGOutPath | Out-Null
            $sOutFileName = "ARM-$($curNSG.NSGNAME).json"
            $sOutJsonName = $($sNSGOutPath) + $sOutFileName
            $sOutPSName = "$($sNSGOutPath)RUN-$($curNSG.NSGNAME).ps1"
            $sNSGOutput = parseNSGs($curNSG)
            $sNSGOutput | Out-File $sOutJsonName -Force
            $sNSGPSOutput = BuildPowerShell $curNSG $curNSG.NSGNAME $sOutFileName
            $sNSGPSOutput | Out-File $sOutPSName -Force
            $iNSG += 1
        }
    }
    Write-Output "Generated $($iNSG) network security group templates." 

    $iRT = 0
    foreach ($curRT in $sRT) {
        if ($curRT.SUB -eq $curSUB.NAME) {
            # EXCELHEADERS: SUB REGION LOCATION ROUTETABLENAME RESOURCEGROUPNAME
            New-Item -ItemType Directory -Force -Path $sRTOutPath | Out-Null
            $sOutFileName = "ARM-$($curRT.ROUTETABLENAME).json"
            $sOutJsonName = $($sRTOutPath) + $sOutFileName
            $sOutPSName = "$($sRTOutPath)RUN-$($curRT.ROUTETABLENAME).ps1"
            $sRTOutput = parseRTs($curRT)
            $sRTOutput | Out-File $sOutJsonName -Force
            $sRTPSOutput = BuildPowerShell $curRT $curRT.ROUTETABLENAME $sOutFileName
            $sRTPSOutput | Out-File $sOutPSName -Force
            $iRT += 1
        }
    }
    Write-Output "Generated $($iRT) route table templates."

    $iVNET = 0
    $iPEERs = 0
    foreach ($curVNET in $sVNET) {
        if ($curVNET.SUB -eq $curSUB.NAME) {
            # EXCELHEADERS: SUB REGION LOCATION VIRTUALNETWORKNAME RESOURCEGROUPNAME ROUTETABLENAME NSGNAME
            New-Item -ItemType Directory -Force -Path $sVNETOutPath | Out-Null
            $sOutFileName = "ARM-$($curVNET.VIRTUALNETWORKNAME).json"
            $sOutJsonName = $($sRTOutPath) + $sOutFileName
            $sOutPSName = "$($sVNETOutPath)RUN-$($curVNET.VIRTUALNETWORKNAME).ps1"
            $sVNETOutput = parseVNETs($curVNET)
            $sVNETOutput | Out-File $sOutJsonName -Force
            $sVNETPSOutput = BuildPowerShell $curVNET $curVNET.VIRTUALNETWORKNAME $sOutFileName
            $sVNETPSOutput | Out-File $sOutPSName -Force
            
            $iPEERs += parsePEERs $curVNET $curSUB
            $iVNET += 1
        }
    }
    Write-Output "Generated $($iVNET) virtual network templates."
    Write-Output "Generated $($iPEERs) virtual network peering templates."

    $iAG = 0
    foreach ($curAG in $sAG) {
        if ($curAG.SUB -eq $curSUB.NAME) {
            # EXCELHEADERS: SUB REGION LOCATION APPLICATIONGATEWAYNAME RESOURCEGROUPNAME VIRTUALNETWORKNAME SUBNETNAME
            New-Item -ItemType Directory -Force -Path $sAGOutPath | Out-Null
            $sOutFileName = "ARM-$($curAG.APPLICATIONGATEWAYNAME).json"
            $sOutJsonName = $($sAGOutPath) + $sOutFileName
            $sOutPSName = "$($sAGOutPath)RUN-$($curAG.APPLICATIONGATEWAYNAME).ps1"
            $sAGOutput = parseAGs($curAG)
            $sAGOutput | Out-File $sOutJsonName -Force
            $sAGPSOutput = BuildPowerShell $curAG $curAG.APPLICATIONGATEWAYNAME $sOutFileName
            $sAGPSOutput | Out-File $sOutPSName -Force
            $iAG += 1
        }
    }
    Write-Output "Generated $($iAG) application gateway templates."

    $iFW = 0
    foreach ($curFW in $sFW) {
        if ($curFW.SUB -eq $curSUB.NAME) {
            # EXCELHEADERS: SUB REGION LOCATION FIREWALLNAME RESOURCEGROUPNAME VIRTUALNETWORKNAME PUBLICIPNAME IPADDRESS
            New-Item -ItemType Directory -Force -Path $sFWOutPath | Out-Null
            $sOutFileName = "ARM-$($curFW.FIREWALLNAME).json"
            $sOutJsonName = $($sFWOutPath) + $sOutFileName
            $sOutPSName = "$($sFWOutPath)RUN-$($curFW.FIREWALLNAME).ps1"
            $sFWOutput = parseFWs($curFW)
            $sFWOutput | Out-File $sOutJsonName -Force
            $sFWPSOutput = BuildPowerShell $curFW $curFW.FIREWALLNAME $sOutFileName
            $sFWPSOutput | Out-File $sOutPSName -Force
            $iFW += 1
        }
    }
    Write-Output "Generated $($iFW) firewall templates."

    $iRSV = 0
    foreach ($curRSV in $sRSV) {
        if ($curRSV.SUB -eq $curSUB.NAME) {
            # EXCELHEADERS: SUB REGION LOCATION VAULTNAME TYPE RESOURCEGROUPNAME SKU METRICS ASR BCK LAWNAME
            New-Item -ItemType Directory -Force -Path $sRSVOutPath | Out-Null
            $sOutFileName = "ARM-$($curRSV.VAULTNAME).json"
            $sOutJsonName = $($sRSVOutPath) + $sOutFileName
            $sOutPSName = "$($sRSVOutPath)RUN-$($curRSV.VAULTNAME).ps1"
            $sRSVOutput = parseRSVs($curRSV)
            $sRSVOutput | Format-Json | Out-File $sOutJsonName -Force
            $sRSVPSOutput = BuildPowerShell $curRSV $curRSV.VAULTNAME $sOutFileName
            $sRSVPSOutput | Format-Json | Out-File $sOutPSName -Force
            $iRSV += 1
        }
    }
    Write-Output "Generated $($iRSV) recovery service vault templates."

    $iSA = 0
    foreach ($curSA in $sSA) {
        if ($curSA.SUB -eq $curSUB.NAME) {
            # EXCELHEADERS: SUB REGION LOCATION STORAGEACCOUNTNAME SKU RESOURCEGROUPNAME DESCRIPTION
            New-Item -ItemType Directory -Force -Path $sSAOutPath | Out-Null
            $sOutFileName = "ARM-$($curSA.STORAGEACCOUNTNAME).json"
            $sOutJsonName = $($sSAOutPath) + $sOutFileName
            $sOutPSName = "$($sSAOutPath)RUN-$($curSA.STORAGEACCOUNTNAME).ps1"
            $sSAOutput = parseSAs($curSA)
            $sSAOutput | Out-File $sOutJsonName -Force
            $sSAPSOutput = BuildPowerShell $curSA $curSA.STORAGEACCOUNTNAME $sOutFileName
            $sSAPSOutput | Out-File $sOutPSName -Force
            $iSA += 1
        }
    }
    Write-Output "Generated $($iSA) storage account templates" 



    <#
    $i = 0
    foreach ($cur in $s) {
        if ($cur.SUB -eq $curSUB.NAME) {
            # EXCELHEADERS: SUB REGION LOCATION RESOURCEGROUPNAME
            New-Item -ItemType Directory -Force -Path $sOutPath | Out-Null
            $sOutFileName = "ARM-$($cur.).json"
            $sOutJsonName = $($sOutPath) + $sOutFileName
            $sOutPSName = "$($sOutPath)RUN-$($cur.).ps1"
            $sOutput = parses($cur)
            $sOutput | Out-File $sOutJsonName -Force
            $sPSOutput = BuildPowerShell $cur $cur. $sOutFileName
            $sPSOutput | Out-File $sOutPSName -Force
            $i += 1
        }
    }
    Write-Output "Generated $($i)  templates"
    #>

}
    
exit