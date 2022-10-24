Param (
[Parameter(Mandatory=$true)][string]$tpidInput,
[Parameter(Mandatory=$true)][string]$bearer_token
)

$targetMode = $true
$ids = Get-Content $tpidInput
$currentUTCtime = (Get-Date).ToUniversalTime()
$startDate = $currentUTCtime.AddDays(-30)

$outfileC = "C:\temp\BlueBleed-LynxReport-{0}.csv" -f $currentUTCtime.tostring("dd-MM-yyyy-hh-mm-ss")  
$newcsv = {} | Select-Object "OrgName","TPID","Customer","Tenant","ID","IsGov","MC442057","MC442048"| Export-Csv $outfileC
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

foreach($id in $ids){
    Write-Host "TPID: $id------------------------------------------------------------"
    #
    # Customer Name and Tenants
    $tenants = get-tenants -tpid $id -bearerToken $bearer_token
    
    
    foreach ($cxtenant in $tenants.Results.Document){
        $messages = get-messages -tenantID $cxtenant.OmsTenantId -bearerToken $bearer_token

        foreach ($message in $messages){
            if($message.Id -eq "MC442057" -or $message.Id -eq "MC442048"){
                $csvfileC.TPID = $id
                $csvfileC.Customer = $cxtenant.MSSalesTopParentOrgName
                $csvfileC.Tenant = $cxtenant.DefaultDomain
                $csvfileC.OrgName = $cxtenant.Name
                $csvfileC.Id = $cxtenant.OmsTenantId
                $csvfileC.IsGov = $cxtenant.IsGov
                Write-Host $cxtenant.Name
                Write-Host $cxtenant.DefaultDomain
                Write-Host $message.Id
                if($message.Id -eq "MC442057"){
                    $csvfileC.MC442057 = "X"
                }
                if($message.Id -eq "MC442048"){
                    $csvfileC.MC442048 = "X"
                }
            }
        }

    $csvfileC | Export-Csv $outfileC -Append 

    }
    

    $cleanFileC =  Import-Csv $outfileC | Where-Object 'Customer' -ne ''
    $cleanFileC | Export-Csv $outfileC
    
}





