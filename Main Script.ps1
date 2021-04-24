<#Welcome to the Automation of SharePoint Modern application using PnP Powershell (Project: ZNA LN2SP Migration)#>
#################################################################################################################################################################################################

function Colour-Combination {
$TitleColour =  -ForegroundColor DarkMagenta -BackgroundColor White
$AlertColour =  -ForegroundColor DarkRed -BackgroundColor White
$SuccessColour =  -ForegroundColor DarkRed -BackgroundColor White}

function CSVInput_MenuList {
#Write-Host  `t "[1] To Clone a Modern Site completely using PnP PowerShell."
Write-Host "[1]  To Create Site collection of Modern Team Site in SharePoint Online."
Write-Host "[2]  To Clone a Modern Site completely using PnP PowerShell."
Write-Host "[3]  To Copy Sites Page(s) from one site to another site."
Write-Host "[4]  To Add/Remove Site Collection App Catalog (Apps For SharePoint) List for Modern Site of SharePoint Online."
Write-Host "[5]  To Enable/Disable Social Option from Modern Page of SharePoint Online."
Write-Host "[6]  To Enable/Disable Comments Option from Modern Page of SharePoint Online."
Write-Host "[7]  To Enable/Disable Banner Option from Modern Page of SharePoint Online."
Write-Host "[8]  To Enable/Disable Quick Launch or Left Navigation from Modern Site of SharePoint Online."
Write-Host "[9]  To Enable/Disable Flow or Power Automate Option from Modern Site of SharePoint Online."
Write-Host "[10] To Enable/Disable AppViews or PowerApps Option from Modern Site of SharePoint Online."
Write-Host "[11] All In One."
Write-Host "[12] To Quit and Exit"
}

function Run-this-Script-with-Admin-access {
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
#exit $LASTEXITCODE}}
}}

function Title {
Write-Host "`n Welcome to the Automation of SharePoint Modern application using PnP Powershell (Project: ZNA LN2SP Migration) " -ForegroundColor DarkMagenta -BackgroundColor White
}

function SDK-CSOM {
# Paths to SDK. #Load SharePoint CSOM Assemblies. Please verify location on your computer.
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll" 
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.Office.Client.Policy.dll"
}

function Log-File-Directory {
# Log File Directory
$LogDir = $PSScriptRoot + "\LogFile"
$LogFileName = "DefaultLogFile_$(Get-Date -format dd-MM-yyyy)_$((Get-Date -format HH:mm:ss).Replace(":","-")).log"
$LogFile = Join-path $LogDir $LogFileName
$Module = $PSScriptRoot + "\DS_PowerShell_Function_Library.psm1"
Import-Module $Module
}

$Create_PnP_Sites_CSV = $PSScriptRoot + "\Scripts\Create_PnP_Sites_CSV\Create_PnP_Sites_CSV.ps1"
$Clone_PnP_Site_CSV = $PSScriptRoot + "\Scripts\Clone_PnP_Site_CSV\Clone_PnP_Site_CSV.ps1"
$Copy_Sites_Pages_CSV = $PSScriptRoot + "\Scripts\Copy_Sites_Pages_CSV\Copy_PnP_Sites_Pages_CSV.ps1"
$Apps_For_SharePoint_CSV = $PSScriptRoot + "\Scripts\Apps_For_SharePoint_CSV\Apps_For_SharePoint_CSV.ps1"
$Social_Option_CSV = $PSScriptRoot + "\Scripts\Social_Option_CSV\Social_Option_CSV.ps1"
$Comments_Option_CSV = $PSScriptRoot + "\Scripts\Comments_Option_CSV\Comments_Option_CSV.ps1"
$Banner_Option_CSV = $PSScriptRoot + "\Scripts\Banner_Option_CSV\Banner_Option_CSV.ps1"
$Quick_Launch_Option_CSV = $PSScriptRoot + "\Scripts\Quick_Launch_CSV\Quick_Launch_CSV.ps1"
$Flow_Option_CSV = $PSScriptRoot + "\Scripts\PowerAutomate_Option_CSV\Flow_Option_CSV.ps1"
$PowerApps_Option_CSV = $PSScriptRoot + "\Scripts\PowerApps_Option_CSV\PowerApps_Option_CSV.ps1"
$AIO_CSV = $PSScriptRoot + "\Scripts\AIO\AIO_CSV.ps1"

function Main-Menu {
  cls
  $Start_Date = Get-Date
     do {
        $userMenuChoice = 0
        while ( $userMenuChoice -lt 1 -or $userMenuChoice -gt 12) 
        {cls
        Write-Host "`n ============= Working on Modern Site using PnP PowerShell. Please choose any of the following options ============= "  -ForegroundColor DarkMagenta -BackgroundColor White
        Write-Host "Following options are depend on CSV Input."-ForegroundColor Green
        CSVInput_MenuList
        $userMenuChoice = Read-Host -Prompt "Which task do you want to perform from the above list? Please enter the task number"
      
        switch ($userMenuChoice)
        {
      
         '1' {
             CLS
             Clear-Host
             Import-Module $Create_PnP_Sites_CSV
             return
             }
                    
         '2' {
             CLS
             Clear-Host
             Import-Module $Clone_PnP_Site_CSV
             return
             }
      
         '3' {
             CLS
             Clear-Host
             Import-Module $Copy_Sites_Pages_CSV
             return
             }
      
         '4' {
             CLS
             Clear-Host
             Import-Module $Apps_For_SharePoint_CSV
             return
             }
      
         '5' {
             CLS
             Clear-Host
             Import-Module $Social_Option_CSV
             return
           }
      
         '6' {
             CLS
             Clear-Host
             Import-Module $Comments_Option_CSV
             return
           }
      
         '7' {
             CLS
             Clear-Host
             Import-Module $Banner_Option_CSV
             return
           }
      
         '8' {
             CLS
             Clear-Host
             Import-Module $Quick_Launch_Option_CSV
             return
           }
      
         '9' {
             CLS
             Clear-Host
             Import-Module $Flow_Option_CSV
             return
           }
      
         '10' {
             CLS
             Clear-Host
             Import-Module $PowerApps_Option_CSV
             return
           }
      
         '11' {
             CLS
             Clear-Host
             Import-Module $AIO_CSV
             return
           }
      
         <#'12' {
              return
              Write-Host "You chose option 3"
              Connect-PnPOnline -Url https://demoperform1.sharepoint.com/sites/OrgSite -UseWebLogin
              $Measures = Measure-PnPList Org
              $Measures.TotalFileSize
              }#>
      
        #default {Write-Host "Nothing or Wrong selection"}
        }
        }
        $End_Date = Get-Date
        Write-Host "`n"
        Write-Host "Job Started Date and Time : $Start_Date"
        Write-Host "Job End Date and Time : $End_Date"
        Write-Host "`n"
        #pause
        #To show the status of execution until and unless we press any key. 
        Write-Host -NoNewLine 'Press any key to close this application';
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');Write-Host "`n"
        } 
#while ($userMenuChoice -ne 12)
until ($userMenuChoice -eq '12')
}

function Start_Script {
 do
   {$PrerequisiteUserInput = Read-Host -Prompt "Are you going to run this application for the first time on this computer?[Yes/No/Exit]"
    switch ($PrerequisiteUserInput)
      {
      'Yes' {
            #cls
            Write-Host "`n Prerequisite-Installation to the work on SharePoint Modern application using PnP Powershell" -ForegroundColor DarkMagenta -BackgroundColor White
            Run-this-Script-with-Admin-access
            Write-Host "Please run the below code via Powershell before click on Enter."
            Write-Host "Uninstall-Module -Name SharePointPnPPowerShellOnline -AllVersions -Force"
            Write-Host "Install-Module -Name SharePointPnPPowerShellOnline -RequiredVersion 3.23.2007.1"
            Write-Host "Get-Module SharePointPnPPowerShell* -ListAvailable | Select-Object Name,Version | Sort-Object Version -Descending"
            Write-Host "Additionally, you should run the "Prerequisite-Installation" via PowerShell with Admin right"
            Write-Host "You should be added in "Site Collection Administration" for every site along with App Catalog."
            pause
            Main-Menu
            return
            } 
       'No' {
            cls
            Main-Menu
            return
            } 
      'Exit'{
            return
            } 
      }
    #Write-Color "Nothing or Wrong input" -Color Red
    #pause 
   }
until ($PrerequisiteUserInput -eq 'No')
}

Title
Write-Host "You should required SharePointPnPPowerShellOnline Version 3.23.2007.1 to run this application and " -NonewLine -ForegroundColor Red
Write-Host "If you are using this application for the first time on your computer then you have to install some configuration otherwise it will not run smoothly/correctly."-ForegroundColor Red
SDK-CSOM
Log-File-Directory
Start_Script

