Param (
[Parameter(Mandatory=$true)][string]$tpidInput,
[Parameter(Mandatory=$true)][string]$bearer_token
)

$targetMode = $true
$ids = Get-Content $tpidInput
$currentUTCtime = (Get-Date).ToUniversalTime()
$startDate = $currentUTCtime.AddDays(-30)

$outfileC = "C:\temp\Consolidated-LynxReport{0}.csv" -f $currentUTCtime.tostring("dd-MM-yyyy-hh-mm-ss")  
$newcsv = {} | Select-Object "OrgName","TPID","Customer","Tenant","ID","IsGov", "E3", "E5", "E5 Sec","MDCA","MDI","AADP2","MDO-U","MDE-U","MDCA-U","MDI-U","AADP2-U","MDO-P","MDE-P","MDI-P","MDCA-P","AADP2-P"| Export-Csv $outfileC
$csvfileC = Import-Csv $outfileC

function get-licenses($tenantID, $bearerToken){
$method = "GET"
$headers = @{Authorization = "Bearer $bearer_token"} 
$url = "https://lynx.office.net/api/LynxStorage/TenantSubscriptions?statusFilters%5B%5D%3DActive&omsTenantId={0}&includeInformationWorkerSubscriptions=false" -f $tenantID

$resLic = Invoke-RestMethod -Uri $url -Method $method -Headers $headers 

return $resLic

}

function get-tenants($tpid, $bearerToken){
    $method = "POST"
    $headers = @{Authorization = "Bearer $bearer_token"}     
    $tpidarray = @($tpid) 

    $payload = @{
        "TopParentOrgIds" = $tpidarray
        "ShowDeletedTenants" = $false
    }

    $payload = $payload | ConvertTo-Json
    $urlt = "https://lynx.office.net/api/Search/Tenants?SearchTerm=*"
    $resTenants = Invoke-RestMethod -uri $urlt -Method $method -Headers $headers -Body $payload -ContentType "application/json"
    return $resTenants 
    }
    
function get-usagestats ($tenantID, $bearerToken, $workload, $et, $startDate, $enddate){
    $method = "GET"
    $headers = @{Authorization = "Bearer $bearerToken"} 
    
    $url = "https://lynx.office.net/api/ApplicationUsage/AllUpHistory?omsTenantId={0}&workloads%5B0%5D={1}&startDate={2}&endDate={3}&usageType=RL28&entityType={4}" -f $tenantID,$workload,$startDate,$enddate,$et

    try{
        $results = Invoke-RestMethod -Uri $url -Method $method -Headers $headers -ContentType "application/json" -UseBasicParsing 
        return $results
    } catch {
        "Error pulling {0} Data, could be no vaild results: {1}" -f $workload,$resLic.statuscode | Write-Host 
        return $null
    }

}

function get-usagePercent($licNum,$usageNum){
    $percent = 0
    if($licNum -eq 0){
        return 0
    }
    try {
        $percent = $usageNum/$licNum
        $a = [math]::Round($percent,2)
        return $a
    } catch {
        return 0
    } 
}

foreach($id in $ids){
    Write-Host "TPID: $id------------------------------------------------------------"
    $et = $workload = ""
    $E3 = $E5 = $E5Sec = $AADP2 = $MDCA = $MDI = $MDATPAverage = $OATPAverage = $AADPAverage = 0
    #
    # Customer Name and Tenants
    $tenants = get-tenants -tpid $id -bearerToken $bearer_token
    
    
    foreach ($cxtenant in $tenants.Results.Document){
        $csvfileC.TPID = $id
        $csvfileC.Customer = $cxtenant.MSSalesTopParentOrgName
        $csvfileC.Tenant = $cxtenant.DefaultDomain
        $csvfileC.OrgName = $cxtenant.Name
        $csvfileC.Id = $cxtenant.OmsTenantId
        $csvfileC.IsGov = $cxtenant.IsGov

    
        Write-Host $cxtenant.Name
    
        $tenLics = get-licenses -tenantID $cxtenant.OmsTenantId -bearerToken $bearer_token
    
        # Extract Licensing Data
        foreach($lic in $tenLics){
            if($lic.StateName = "Active" -and $lic.SubscriptionEndDate -gt  $currentUTCtime){
                switch -wildcard ( $lic.OfferProductName )
                    {
                        '* 365 *3*' { $E3 += $lic.IncludedQuantity   }
                        '* 365 *5*' { $E5 += $lic.IncludedQuantity   }
                        'ENTERPRISE MOBILITY + SECURITY*'{ $E5Sec += $lic.IncludedQuantity     }
                        'MICROSOFT DEFENDER FOR CLOUD APPS*' { $MDCA += $lic.IncludedQuantity   }
                        'MICROSOFT DEFENDER FOR IDENTITY*' { $MDI += $lic.IncludedQuantity   }
                        'AZURE ACTIVE DIRECTORY PREMIUM P2*'{ $AADP2 += $lic.IncludedQuantity     }
                    }
            }
        }
        $csvfileC.E3 = $E3
        $csvfileC.E5 = $E5
        $csvfileC.'E5 Sec' = $E5Sec
        $csvfileC.MDCA = $MDCA
        $csvfileC.MDI = $MDI
        $csvfileC.AADP2 = $AADP2

        if ($E5Sec -gt $E5) { 
            $secLic = $E5Sec
        } else { 
            $secLic =$E5
        }

        if ($MDI -gt $secLic) { 
            $mdiLic = $MDI
        } else { 
            $mdiLic = $secLic
        }
    
        $1 = get-usagestats -tenantID $cxtenant.OmsTenantId -bearerToken $bearer_token -et "Device" -startDate $startDate -enddate $currentUTCtime -workload "MDATP"
        $MDATPAverage = $1.Usage.MDATP |  ForEach-Object {$_.Usage} | Measure-Object -Average
        $MDEAvg = [math]::Round($MDATPAverage.Average)
        $csvfileC.'MDE-U' = $MDEAvg
        $percentUsage = get-usagePercent -licNum $secLic -usageNum $MDEAvg
        $csvfileC.'MDE-P' = $percentUsage 
    
        $2 = get-usagestats -tenantID $cxtenant.OmsTenantId -bearerToken $bearer_token -et "User" -startDate $startDate -enddate $currentUTCtime -workload "OATP"
        $OATPAverage = $2.Usage.OATP |  ForEach-Object {$_.Usage} | Measure-Object -Average
        $MDOAvg = [math]::Round($OATPAverage.Average)
        $csvfileC.'MDO-U' = $MDOAvg
        $percentUsage = get-usagePercent -licNum $secLic -usageNum $MDOAvg
        $csvfileC.'MDO-P' = $percentUsage 

        $2 = get-usagestats -tenantID $cxtenant.OmsTenantId -bearerToken $bearer_token -et "User" -startDate $startDate -enddate $currentUTCtime -workload "MCAS"
        $MCASAverage = $2.Usage.MCAS |  ForEach-Object {$_.Usage} | Measure-Object -Average
        $MCASAvg = [math]::Round($MCASAverage.Average)
        $csvfileC.'MDCA-U' = $MCASAvg
        $percentUsage = get-usagePercent -licNum $secLic -usageNum $MCASAvg
        $csvfileC.'MDCA-P' = $percentUsage 

        $2 = get-usagestats -tenantID $cxtenant.OmsTenantId -bearerToken $bearer_token -et "User" -startDate $startDate -enddate $currentUTCtime -workload "AATP"
        $AATPAverage = $2.Usage.AATP |  ForEach-Object {$_.Usage} | Measure-Object -Average
        $MDIAvg = [math]::Round($AATPAverage.Average)
        $csvfileC.'MDI-U' = $MDIAvg
        $percentUsage = get-usagePercent -licNum $mdiLic -usageNum $MDIAvg
        $csvfileC.'MDI-P' = $percentUsage 
    
        $3 = get-usagestats -tenantID $cxtenant.OmsTenantId -bearerToken $bearer_token -et "User" -startDate $startDate -enddate $currentUTCtime -workload "AADP"
        $AADPAverage = $3.Usage.AADP |  ForEach-Object {$_.Usage} | Measure-Object -Average
        $AADPAvg = [math]::Round($AADPAverage.Average)
        $csvfileC.'AADP2-U' = $AADPAvg
        $percentUsage = get-usagePercent -licNum $secLic -usageNum $AADPAvg
        $csvfileC.'AADP2-P' = $percentUsage 
        
        if($targetMode){
            if($E3 -eq 0 -and $E5 -eq 0){
                $E3 = $E5 = $E5Sec = $secLic = $AADP2 = $MDCA = $MDI = $MDATPAverage = $OATPAverage = $AADPAverage = 0 
            } else {
                $csvfileC | Export-Csv $outfileC -Append 
            }
        } else {
            $csvfileC | Export-Csv $outfileC -Append 
        }
        
        $E3 = $E5 = $E5Sec = $secLic = $AADP2 = $MDCA = $MDI = $MDATPAverage = $OATPAverage = $AADPAverage = 0
    }
    

    $cleanFileC =  Import-Csv $outfileC | Where-Object 'Customer' -ne ''
    $cleanFileC | Export-Csv $outfileC
    
}





