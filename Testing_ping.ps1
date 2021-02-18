$input_test = Read-Host "Enter path to .csv file Containing only HOST + IP file..."
$item_for_test = Import-Csv $input_test -Delimiter ";"

#set the objects to work on them later on
$vol2 = $item_for_test | Select-Object  hostname, ipaddress

#loop for araychecking from the .csv file
foreach ($vol1 in $vol2) {
    if (Test-Connection $vol1.IPAddress -Count 2 -Quiet) {
        Write-Host $vol1.HostName + $vol1.IPAddress "is pingable" -ForegroundColor Green
    }else {
        Write-Host $vol1.HostName + $vol1.IPAddress "is not pingable" -ForegroundColor Red
    }
}
Write-Host "ping complete"


#### needs more optimisation -- :) 



