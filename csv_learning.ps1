#getting/changing user prop in AD 

$importingad = Import-Csv -Path "C:\Users\mkm6si\Downloads\Outlook_Umbenennung_Arbeitsort.csv" -Delimiter ";" 

foreach ($user in $importingad) {
    #find the users 
    $user_1 = $user.Email
    $user_2 = Get-ADUser -Filter {mail -eq $user_1} -Properties mail
    if ($user_2) {
        Write-Host "$user_1"
        Set-ADUser $user_2 -Office $user.'chang to new location'
    }else {
        Write-Host "$user_2 not found."
    }
   
    
}