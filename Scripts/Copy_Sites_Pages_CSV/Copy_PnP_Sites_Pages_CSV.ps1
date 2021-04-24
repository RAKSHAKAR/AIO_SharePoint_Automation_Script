

cls
Write-Host "`n Welcome to the Create a Modern Team Site using CSV via PnP PowerShell (Project: ZNA LN2SP Migration)" -ForegroundColor DarkMagenta -BackgroundColor White
Write-Host "`n"
Write-Host "###########################################################################################################"
Write-Host "This script is using Copy_Site_Pages.csv. So Please update Copy_Site_Pages.csv before starting this script." -ForegroundColor Red
Write-Host "###########################################################################################################"
<#Write-Host "Press any key to Start this script once you have ready with above mentioned details." -ForegroundColor Green
Write-Host "`n"
#To show the status of execution until and unless we press any key. 
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
#>
Pause
<################################################################>
Write-Host "Starting to Copy a Modern Team Site Pages using PnP PowerShell" -ForegroundColor Green
Write-Host "##############################################################"

#Input File Location
$File = $PSScriptRoot + "\Copy_Site_Pages.csv"
$csv = Import-Csv $File

#Output File Location
$output = $PSScriptRoot + "\Copy_Site_Pages_Result.csv"
#$csvOutput= Export-Csv $output
#write-output "Tenant_URL,Source_URL,Page_Name,Destination_URL,Status,Start_Date, End_Date" | Out-File -FilePath $output -Append -Encoding ascii 
#$Added= "List Added"     


foreach ($row in $csv)
{
        $Tenant_URL = $row.Tenant_URL
        $Page_Name=$row.Page_Name
        $Source_URL=$row.Source_URL
        $Destination_URL=$row.Destination_URL
    
    $Start_Date = Get-Date        
    #Login to Tenant_URL
    Connect-PnPOnline -Url $Tenant_URL -UseWebLogin
    #Login to Source_URL
    Connect-PnPOnline -Url $Source_URL -UseWebLogin

    #Copy Pages(s) from Source_URL
    $tempFile = [System.IO.Path]::GetTempFileName();
    Export-PnPClientSidePage -Force -Identity $Page_Name -Out $tempFile

    #Apply Pages(s) to Destination_URL
    Connect-PnPOnline -Url $Destination_URL -UseWebLogin
    Apply-PnPProvisioningTemplate -Path $tempFile
    Write-Host "$Page_Name Page is successfully copied!"
    sleep 10
    
    $Status="$Page_Name Page is successfully copied!"
    Disconnect-PnPOnline
    $End_Date = Get-Date
    Write-Output "$Tenant_URL,$Source_URL,$Page_Name,$Destination_URL,$Status,$Start_Date, $End_Date" | Out-File -FilePath $Output -Append -Encoding ascii
}

 

 $UserInput = Read-Host -Prompt 'Do you need this execution report in Excel?[Y/N]'
    switch ($UserInput)
      {
      'Y' {
            # Specify the path to the Excel file and the WorkSheet Name
            $FilePath = $PSScriptRoot + "\Copy_Site_Pages_Result.csv"

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
