#if -eq not true return to list with object and find object


$get_all_user = Get-ADUser -Filter 'Name -notlike "*-t*"' | Select-Object Name, GivenName, Surname
$collectionUSER = New-Object -TypeName "System.Collections.ArrayList" 
$collectionFAIL = New-Object -TypeName "System.Collections.ArrayList" 

#check all user for admin
foreach ($user in $get_all_user) {

   $user1 = $user.Name + "*"
   $userNEW = Get-ADUser -Filter 'Name -like $user1' #changing the string for filter in AD

   if ($userNEW.Count -gt 1) {
    #if the filter result is greather then 1 (means has _admin or -t something to) check all results for surname and givenname    

      foreach ($adUSER in $userNEW){
         if (($adUSER.GivenName -ne $user.GivenName) -or ($adUSER.Surname -ne $user.Surname)) {
            Write-Host $adUser.Name.ToString()
            $collectionUSER.Add($user)
            $collectionFAIL.Add($adUSER)
         }   
      }

   }
}

$Exportlist = @()
for($i = 0; $i -lt $collectionUSER.Count; $i++){
    $obj = New-Object psobject
    $obj = Add-Member -InputObject $obj Checked_USER $collectionUSER.Name[$i] -PassThru
    $obj = Add-Member -InputObject $obj Checked_NAME $collectionUSER.GivenName[$i] -PassThru
    $obj = Add-Member -InputObject $obj Checked_SURNAME $collectionUSER.Surname[$i] -PassThru
    $obj = Add-Member -InputObject $obj FAIL_USER $collectionFAIL.Name[$i] -PassThru
    $obj = Add-Member -InputObject $obj FAIL_NAME $collectionFAIL.GivenName[$i] -PassThru
    $obj = Add-Member -InputObject $obj FAIL_SURRNAME $collectionFAIL.Surname[$i] -PassThru
    $Exportlist += $obj
}

$Exportlist | Export-Csv -Path "C:\Users\mkm6si\OneDrive - SEG Automotive\Desktop\Powershell\CHECK_ADMIN\names.csv" -NoTypeInformation -Force -Delimiter ";"