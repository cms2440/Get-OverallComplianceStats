#Domains being struck
$domainFQDNs = @"
acc.accroot.ds.af.smil.mil
afmc.ds.af.smil.mil
aetc.aetcroot.ds.af.smil.mil
amchub.amc.ds.af.smil.mil
usafe.usaferoot.ds.af.smil.mil
eielson.pacaf.ds.af.smil.mil
misawa.pacaf.ds.af.smil.mil
andersen.pacaf.ds.af.smil.mil
diego.pacaf.ds.af.smil.mil
elmendorf.pacaf.ds.af.smil.mil
hickam.pacaf.ds.af.smil.mil
kadena.pacaf.ds.af.smil.mil
kunsan.pacaf.ds.af.smil.mil
osan.pacaf.ds.af.smil.mil
nosc.pacaf.ds.af.smil.mil
yokota.pacaf.ds.af.smil.mil
"@.split("`n") | foreach {$_.trim()}

#Rows tracked in spreadsheet
$accountTypes = @"
Admin Accounts
AdminLevel Groups
BaseLevel Groups
Computers
Managed Service Accounts
Role Accounts
Service Accounts
Users
"@.Split("`n") | foreach {$_.trim()}

#Make dynamic variables for the Object Types
foreach ($type in $accountTypes) {
    New-Variable $($type.replace(" ","_") + "_Total") -Value 0 -Force
    New-Variable $($type.replace(" ","_") + "_Compliant") -Value 0 -Force
    }
$tmpArr = @()

:main foreach ($Domain in $domainFQDNs) {
    #Error handling if we cant reach domain
    try {
        $StatsPath = gci "\\$domain\netlogon\NonCompliance" -Filter "*Compliance Stats.csv" -EA Stop | select -ExpandProperty fullname
        }
    catch {
        foreach ($type in $accountTypes) {
            $tmpArr += New-Object PSObject -Property ([ordered]@{
                Domain = $domain.split(".")[0]
                "Object Type" = $type
                Total = "CouldNotContact"
                Compliant = "CouldNotContact"
                "Percent Compliant" = "N/A"
                })
            }
        continue main
        }
    #Gather Stats from noncompliance share Compliance Stats.csv
    $csv = Import-Csv $StatsPath
    $domainStats = $csv | where base -like "* Overall"
    if (!$domainStats) {$domainStats = $csv | where base -notlike "*NOS"}

    #Gather numbers for each Object Type
    foreach ($type in $accountTypes) {
        $typeStat = $domainStats | where "object Type" -eq $type
        if (!$typeStat) {continue}

        [int]$tmpTotal = ($typeStat | select -ExpandProperty total).replace(",","")
        [int]$tmpCompliant = ($typeStat | select -ExpandProperty compliant).replace(",","")
        
        $tmpTotal += (Get-Variable $($type.replace(" ","_") + "_Total")).Value
        $tmpCompliant += (Get-Variable $($type.replace(" ","_") + "_Compliant")).Value

        Set-Variable $($type.replace(" ","_") + "_Total") -Value $tmpTotal
        Set-Variable $($type.replace(" ","_") + "_Compliant") -Value $tmpCompliant
        }

    #Add domain's overall stats
    $tmpArr += $domainStats | select @{n="Domain";e={$_.base.replace(" Overall","")}},"Object Type",Total,Compliant,"Percent Compliant"

    }

#Put into final array, putting Overall on top
$FinalArr = @()
foreach ($type in $accountTypes) {
    $tempPercent = (Get-Variable $($type.replace(" ","_") + "_Compliant")).Value / (Get-Variable $($type.replace(" ","_") + "_Total")).Value
    $FinalArr += New-Object PSObject -Property ([ordered]@{
        Domain = "Overall"
        "Object Type" = $type
        Total = "{0:N0}" -f (Get-Variable $($type.replace(" ","_") + "_Total")).Value
        Compliant = "{0:N0}" -f (Get-Variable $($type.replace(" ","_") + "_Compliant")).Value
        "Percent Compliant" = [string]([math]::Round($tempPercent * 100,2)) + "%"
        })
    }

$FinalArr += $tmpArr
$desktop = [environment]::GetFolderPath("Desktop")
$timestamp = Get-Date -Format "yyyyMMdd"
$FinalArr | Export-Csv -NoTypeInformation "$desktop\OverallComplianceStats_$timestamp.csv"

<#
catch {
    $_.exception.message | Write-Host -ForegroundColor red
    $_.InvocationInfo.PositionMessage | Write-Host -ForegroundColor red
    }
    #>
