$something = "FILE_TO_SCAN" 
$collection = Import-Excel -Path $something | Select-Object Link
$collection1 = New-Object -TypeName "System.Collections.ArrayList"

foreach ($item in $collection) {
   $result =  "{0:N2}" -f ((Get-ChildItem –force $item.Link –Recurse -ErrorAction SilentlyContinue| Measure-Object Length -s).sum / 1Gb)
   Write-Host $item.Link + "   "  $result
   $collection1.Add($result)
}

$Exportlist = @()
for ( $i = 0; $i -lt $collection.Length; $i++){
    $obj = New-Object psobject
    $obj = Add-Member -InputObject $obj Link $collection[$i] -PassThru
    $obj = Add-Member -InputObject $obj Size  $collection1[$i] -PassThru
    $Exportlist += $obj
}

$Exportlist | Export-Csv -Path "PATH_TO_SAVE" -NoTypeInformation -Force -Delimiter ";"


