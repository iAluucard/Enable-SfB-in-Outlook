####################################################
##                                                ##
##          SEG Automotive   eForms               ##
##                                                ##
####################################################
#if you happen to have some suggestions please contact me 
#for enabling eForms, please test this 


#adding the .net Forms to the script 
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$war_message = [System.Windows.Forms.MessageBox]::Show("!!!!  Please close all your work on Outlook and Outlook then confirm with a click on the button  !!!! THIS SCRIPT WILL DO CORE CHANGES ON OUTLOOK  !!!!", "Outlook configuration", "OKCancel", "Warning")
$outlook = Get-Process -Name "*outlook"
$war_message 

#based on answer the block of code will be execuded 
switch ($war_message) {
    
    "OK"{

        Start-Process -FilePath "C:\Program Files (x86)\eForms\MLAGENT.exe"
        Start-Sleep -Seconds 5

            if ($null -eq $outlook) {

                    Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\CrashingAddinList" -Name * #for removing all chached problems for addons in outlook
                    Get-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\CrashingAddinList" | Out-Null
                    Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems" -Name * #clearing all disabled elements in addons
                    Get-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems" | Out-Null
                    Get-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList" | Set-ItemProperty -Name MLOCAI02.Connect -Value 1

                Start-Process -FilePath "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"
                Start-Sleep -Seconds 10
                try {
                     Get-Item "HKCU:\Software\Microsoft\Office\Outlook\Addins\MLOCAI02.Connect" | Set-ItemProperty -Name LoadBehavior -Value 3
                }
                catch {
                    
                }        
            }
            else
            {

            Stop-Process -Name "*outlook"
            Start-Sleep -Seconds 5
                Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\CrashingAddinList" -Name * #for removing all chached problems for addons in outlook
                Get-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\CrashingAddinList" | Out-Null
                Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems" -Name * #clearing all disabled elements in addons
                Get-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems" | Out-Null
                Get-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList" | Set-ItemProperty -Name MLOCAI02.Connect -Value 1
            Start-Process -FilePath "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"

            try {
                Get-Item "HKCU:\Software\Microsoft\Office\Outlook\Addins\MLOCAI02.Connect" | Set-ItemProperty -Name LoadBehavior -Value 3
            }
            catch {
                
            }
            
            Start-Sleep -Seconds 10

            }         

            Stop-Process -Name "*outlook"
            Start-Sleep -Seconds 5
            Start-Process -FilePath "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"

    }

    "Cancel" { 
       
    }
}
Start-Sleep -Seconds 10
[System.Windows.Forms.MessageBox]::Show("Reenabling the eForms addon is done.", "Ok")




    


