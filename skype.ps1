####################################################
##                                                ##
##          SEG Automotive Skype for Bussines     ##
##                 addon in outlook               ##
##                                                ##
####################################################
#if you happen to have some suggestions please contact me 
#for enabling eForms, please test this 


#adding the .net Forms to the script 
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$war_message = [System.Windows.Forms.MessageBox]::Show("!!!!  Please close all your work on Outlook and Outlook then confirm with a click on the button  !!!! THIS SCRIPT WILL DO CORE CHANGES ON OUTLOOK  !!!!", "Outlook configuration", "OKCancel", "Warning")
$kill_outlook = Get-Process -Name "*outlook" 

switch ($war_message) {
    "OK" { 
            Stop-Process $kill_outlook

                    Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\CrashingAddinList" -Name * #for removing all chached problems for addons in outlook
                    Get-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\CrashingAddinList" | Out-Null
                    Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems" -Name * #clearing all disabled elements in addons
                    Get-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems" | Out-Null

            Get-Item "HKCU:\Software\Microsoft\Office\16.0\Outlook\Addins\UCAddin.LyncAddin.1" | Set-ItemProperty -Name (Default) -Value 0
            Get-Item "HKCU:\Software\Microsoft\Office\16.0\Outlook\Addins\UCAddin.LyncAddin.1" | Set-ItemProperty -Name LoadBehavior -Value 3
     }
    
    "Cancel"{

    }
}

Start-Sleep -Seconds 10
[System.Windows.Forms.MessageBox]::Show("Reenabling the Skype addon is done.", "Ok")