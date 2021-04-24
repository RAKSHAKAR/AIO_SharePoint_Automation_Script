

cls
Write-Host "`n Welcome to the Create a Modern Team Site using CSV via PnP PowerShell (Project: ZNA LN2SP Migration)" -ForegroundColor DarkMagenta -BackgroundColor White
Write-Host "`n"
Write-Host "###################################################################################################################################"
Write-Host "This script is using Flow_Option.csv. So Please update Flow_Option.csv before starting this script." -ForegroundColor Red
Write-Host "###################################################################################################################################"
<#Write-Host "Press any key to Start this script once you have ready with above mentioned details." -ForegroundColor Green
Write-Host "`n"
#To show the status of execution until and unless we press any key. 
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
#>
Pause
<################################################################>
Write-Host "Starting to Add/Remove Flow Option from Modern Page of SharePoint Online using PnP PowerShell" -ForegroundColor Green
Write-Host "###############################################################################################"

#Input File Location
$File = $PSScriptRoot + "\Flow_Option.csv"
$csv = Import-Csv $File

#Output File Location
$output = $PSScriptRoot + "\Flow_Option_Result.csv"
#$csvOutput= Export-Csv $output
#write-output "Tenant_URL,Destination_URL,Page_Name,Disable_Flow,Status,Start_Date,End_Date" | Out-File -FilePath $output -Append -Encoding ascii
#$Added= "List Added"     

foreach ($row in $csv)
{
        $Tenant_URL = $row.Tenant_URL
        $Destination_URL=$row.Destination_URL
        $Disable_Flow=$row.Disable_Flow

   $Start_Date = Get-Date        
   #Login to Tenant_URL
   Connect-PnPOnline -Url $Tenant_URL -UseWebLogin
   #Login to Destination_URL
   Connect-PnPOnline -Url $Destination_URL -UseWebLogin
   
   If ($Disable_Flow -eq "Yes") 
        {
        #Add Social button to the Mordern Page
        Set-PnPWeb -DisablePowerAutomate:$true
        Write-Host "Flow or power Automate Option is successfully removed!"
        $Status="Flow or power Automate Option is successfully removed!"
        }
        elseif ($Disable_Flow -eq "No") 
        {
        #Remove Social button to the Mordern Page
        Set-PnPWeb -DisablePowerAutomate:$false
        Write-Host "Flow or power Automate Option is successfully added!"
        $Status="Flow or power Automate Option is successfully added!"
        }
    Disconnect-PnPOnline
    $End_Date = Get-Date
    Write-Output "$Tenant_URL,$Destination_URL,$Page_Name,$Disable_Flow,$Status,$Start_Date,$End_Date" | Out-File -FilePath $Output -Append -Encoding ascii
}


 $UserInput = Read-Host -Prompt 'Do you need this execution report in Excel?[Y/N]'
    switch ($UserInput)
      {
      'Y' {
            # Specify the path to the Excel file and the WorkSheet Name
            $FilePath = $PSScriptRoot + "\Flow_Option_Result.csv"

            # Create an Object Excel.Application using Com interface
            $objExcel = New-Object -ComObject Excel.Application

            # Disable the 'visible' property so the document won't open in excel
            $objExcel.Visible = $true

            # Open the Excel file and save it in $WorkBook
            $WorkBook = $objExcel.Workbooks.Open($FilePath)
            
            Main-Menu
            return
          }
      'N' {
            #return
            Main-Menu
            return
           }
      }
