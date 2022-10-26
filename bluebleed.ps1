Param (
[Parameter(Mandatory=$true)][string]$tpidInput,
[Parameter(Mandatory=$true)][string]$bearer_token
)

$ids = Get-Content $tpidInput
$currentUTCtime = (Get-Date).ToUniversalTime()
$outfileC = "C:\temp\BlueBleed-{0}.csv" -f $currentUTCtime.tostring("dd-MM-yyyy-hh-mm-ss")  

$newcsv = {} | Select-Object "TPID","Customer","Domain","Messages"| Export-Csv $outfileC
$csvfileC = Import-Csv $outfileC

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
    
function get-messages ($tenantID, $bearerToken){
    $method = "GET"
    $headers = @{Authorization = "Bearer $bearerToken"} 
    
    $url = "https://lynx.office.net/api/Tenant/Messages?omsTenantId={0}" -f $tenantID

    try{
        $results = Invoke-RestMethod -Uri $url -Method $method -Headers $headers -ContentType "application/json" -UseBasicParsing 
        return $results
    } catch {
        "Error pulling {0} Data, could be no vaild results: {1}" -f $workload,$results.statuscode | Write-Host 
        return $null
    }

}

 "TPID,Customer,Domain,Messages" | Write-Host  

foreach($id in $ids){
    $tpid=$id
    $tenants = get-tenants -tpid $tpid -bearerToken $bearer_token
     
    foreach ($cxtenant in $tenants.Results.Document){
        $messages = get-messages -tenantID $cxtenant.OmsTenantId -bearerToken $bearer_token

        foreach ($message in $messages){
            if(($message.Id -eq "MC442057") -or ($message.Id -eq "MC442048")){

                "{0},{1},{2},{3}" -f $tpid,$cxtenant.Name,$cxtenant.DefaultDomain,$message.Id | Write-Host 
                $csvfileC.TPID = $tpid
                $csvfileC.Customer = $cxtenant.Name
                $csvfileC.Domain = $cxtenant.DefaultDomain
                $csvfileC.Messages = $message.Id
                $csvfileC | Export-Csv $outfileC -Append 
            }  
        }
    }
    $cleanFileC =  Import-Csv $outfileC | Where-Object 'Customer' -ne ''
    $cleanFileC | Export-Csv $outfileC
}





