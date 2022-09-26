Param (
[Parameter(Mandatory=$true)][string]$tpidInput,
[Parameter(Mandatory=$true)][string]$bearer_token
)

$ids = Get-Content $tpidInput
$currentUTCtime = (Get-Date).ToUniversalTime()
$startDate = $currentUTCtime.AddDays(-30)

$outfileC = "C:\temp\Consolidated-LynxReport{0}.csv" -f $currentUTCtime.tostring("dd-MM-yyyy-hh-mm-ss")  
$newcsv = {} | Select-Object "Customer","TPID","OrgName","Tenant","ID","IsGov", "E3", "E5", "E5 Sec","MDCA","MDI","AADP2","MDO-U","MDE-U","MDCA-U","MDI-U","AADP2-U"| Export-Csv $outfileC
$csvfileC = Import-Csv $outfileC

    foreach($id in $ids){
        Write-Host "TPID: $id------------------------------------------------------------"
        $et = $workload = ""
        $E3 = $E5 = $E5Sec = $AADP2 = $MDCA = $MDI = $MDATPAverage = $OATPAverage = $AADPAverage = 0
        $outfile = "C:\temp\{0}-LynxReport{1}.csv" -f $id,$currentUTCtime.tostring("dd-MM-yyyy-hh-mm-ss")  
        $newcsv = {} | Select-Object "Customer","OrgName","Tenant","ID","IsGov", "E3", "E5", "E5 Sec","MDCA","MDI","AADP2","MDO-U","MDE-U","MDCA-U","MDI-U","AADP2-U"| Export-Csv $outfile
        $csvfile = Import-Csv $outfile
        
        
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
            # $urlt = "https://lynx.office.net/api/LynxStorage/Customer?tpid={0}&PageSize=25" -f $tpid
        
        
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
        
        #
        # Customer Name and Tenants
        $tenants = get-tenants -tpid $id -bearerToken $bearer_token
        
        
        foreach ($cxtenant in $tenants.Results.Document){
            $csvfileC.TPID = $id
            $csvfileC.Customer = $csvfile.Customer = $cxtenant.MSSalesTopParentOrgName
            $csvfileC.Tenant = $csvfile.Tenant = $cxtenant.DefaultDomain
            $csvfileC.OrgName = $csvfile.OrgName = $cxtenant.Name
            $csvfileC.Id = $csvfile.Id = $cxtenant.OmsTenantId
            $csvfileC.IsGov = $csvfile.IsGov = $cxtenant.IsGov
        
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
            $csvfileC.E3 = $csvfile.E3 = $E3
            $csvfileC.E5 = $csvfile.E5 = $E5
            $csvfileC.'E5 Sec' = $csvfile.'E5 Sec' = $E5Sec
            $csvfileC.MDCA = $csvfile.MDCA = $MDCA
            $csvfileC.MDI = $csvfile.MDI = $MDI
            $csvfileC.AADP2 = $csvfile.AADP2 = $AADP2
        
            $1 = get-usagestats -tenantID $cxtenant.OmsTenantId -bearerToken $bearer_token -et "Device" -startDate $startDate -enddate $currentUTCtime -workload "MDATP"
            $MDATPAverage = $1.Usage.MDATP |  ForEach-Object {$_.Usage} | Measure-Object -Average
            $csvfileC.'MDE-U' = $csvfile.'MDE-U' = $MDATPAverage.Average
        
            $2 = get-usagestats -tenantID $cxtenant.OmsTenantId -bearerToken $bearer_token -et "User" -startDate $startDate -enddate $currentUTCtime -workload "OATP"
            $OATPAverage = $2.Usage.OATP |  ForEach-Object {$_.Usage} | Measure-Object -Average
            $csvfileC.'MDO-U' = $csvfile.'MDO-U' = $OATPAverage.Average

            $2 = get-usagestats -tenantID $cxtenant.OmsTenantId -bearerToken $bearer_token -et "User" -startDate $startDate -enddate $currentUTCtime -workload "MCAS"
            $OATPAverage = $2.Usage.MCAS |  ForEach-Object {$_.Usage} | Measure-Object -Average
            $csvfileC.'MDCA-U' = $csvfile.'MDCA-U' = $OATPAverage.Average

            $2 = get-usagestats -tenantID $cxtenant.OmsTenantId -bearerToken $bearer_token -et "User" -startDate $startDate -enddate $currentUTCtime -workload "AATP"
            $OATPAverage = $2.Usage.AATP |  ForEach-Object {$_.Usage} | Measure-Object -Average
            $csvfileC.'MDI-U' = $csvfile.'MDI-U' = $OATPAverage.Average
        
            $3 = get-usagestats -tenantID $cxtenant.OmsTenantId -bearerToken $bearer_token -et "User" -startDate $startDate -enddate $currentUTCtime -workload "AADP"
            $AADPAverage = $3.Usage.AADP |  ForEach-Object {$_.Usage} | Measure-Object -Average
            $csvfileC.'AADP2-U' = $csvfile.'AADP2-U' = $AADPAverage.Average
            
            $csvfile | Export-Csv $outfile -Append
            $csvfileC | Export-Csv $outfileC -Append
            $E3 = $E5 = $E5Sec = $AADP2 = $MDCA = $MDI = $MDATPAverage = $OATPAverage = $AADPAverage = 0
        }
        
        #Clean Up CSV
        $cleanFile =  Import-Csv $outfile | Where-Object 'Customer' -ne ''
        $cleanFile | Export-Csv $outfile

        $cleanFileC =  Import-Csv $outfileC | Where-Object 'Customer' -ne ''
        $cleanFileC | Export-Csv $outfileC

        
}





