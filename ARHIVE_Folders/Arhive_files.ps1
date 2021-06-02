$something = "C:\Users\mkm6si\Desktop\Powershell\SGHIFS01_to_be_Archive_folders.xlsx" 
$collection = Import-Excel -Path $something | Select-Object Link




foreach ($item in $collection) {
   $result =  "{0:N2}" -f ((Get-ChildItem -path $item.Link -Recurse -ErrorAction SilentlyContinue -Force |  Measure-Object -property length -sum ).sum /1MB) + " MB" 
   Write-Host $item.Link + "   " + $result 
}