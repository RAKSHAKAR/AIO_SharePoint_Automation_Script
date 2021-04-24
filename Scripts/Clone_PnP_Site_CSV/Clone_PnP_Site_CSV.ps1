

cls
Write-Host "`n Welcome to the Clone a Modern Team Site using CSV via PnP PowerShell (Project: ZNA LN2SP Migration)" -ForegroundColor DarkMagenta -BackgroundColor White
Write-Host "`n"
Write-Host "###################################################################################################"
Write-Host "This script is using Clone_Site.csv. So Please update Clone_Site.csv before starting this script." -ForegroundColor Red
Write-Host "###################################################################################################"
<#Write-Host "Press any key to Start this script once you have ready with above mentioned details." -ForegroundColor Green
Write-Host "`n"
#To show the status of execution until and unless we press any key. 
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
#>
Pause
Write-Host "`n"
<################################################################>
Write-Host "Starting to Clone a Modern Team Site using PnP PowerShell" -ForegroundColor Green
Write-Host "##########################################################"

#Input File Location
$File = $PSScriptRoot + "\Clone_Site.csv"
$csv = Import-Csv $File
$xml = $PSScriptRoot + "\MyApplications.xml"

#Output File Location
$output = $PSScriptRoot + "\Clone_Site_Result.csv"
#$csvOutput= Export-Csv $output
#write-output "Tenant_URL,Source_URL,Destination_URL,Status,Start_Date,End_Date" | Out-File -FilePath $output -Append -Encoding ascii 
#$Added= "List Added"     


foreach ($row in $csv)
{
        $Tenant_URL = $row.Tenant_URL
        $Source_URL=$row.Source_URL
        $Destination_URL=$row.Destination_URL
    
    $Start_Date = Get-Date        
    #Login to Tenant_URL
    Connect-PnPOnline -Url $Tenant_URL -UseWebLogin
    #Login to Source_URL
    Connect-PnPOnline -Url $Source_URL -UseWebLogin
    Set-PnPTenantCdnEnabled -CdnType Both -Enable $true

    #Get the Template of the above Modern Source site using PnP Powershell script
    Write-Host "Getting the Site Template from $Source_URL" -ForegroundColor Green
    Get-PnPProvisioningTemplate -Out "$xml" -IncludeAllClientSidePages
    
    #After creating the above Team site now try to Apply the extracted PnP Template only after Connecting it to that new Modern Target Site by running the scripts
    Connect-PnPOnline -Url $Destination_URL -UseWebLogin
        
    Write-Host "Applying site template to $Destination_URL" -ForegroundColor Green
    Apply-PnPProvisioningTemplate -Path "$xml"
    Add-PnPSiteCollectionAppCatalog -Site $Destination_URL
    Set-PnPTenantCdnEnabled -CdnType Both -Enable $true

    Write-Host "`n"
    write-host "Cloning is completed successfully!" -foregroundcolor Yellow
    $Status="Cloning is completed successfully!"
    Disconnect-PnPOnline

    $End_Date = Get-Date
    Write-Output "$TenantURL,$Source_URL,$Destination_URL,$Status,$Start_Date,$End_Date" | Out-File -FilePath $Output -Append -Encoding ascii
}
    

 $UserInput = Read-Host -Prompt 'Do you need this execution report in Excel?[Y/N]'
    switch ($UserInput)
      {
      'Y' {
            # Specify the path to the Excel file and the WorkSheet Name
            $FilePath = $PSScriptRoot + "\Clone_Site_Result.csv"
            
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
