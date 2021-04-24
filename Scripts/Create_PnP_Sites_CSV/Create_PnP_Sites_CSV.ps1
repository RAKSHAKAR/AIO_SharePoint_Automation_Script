

cls
Write-Host "`n Welcome to the Create a Modern Team Site using CSV via PnP PowerShell (Project: ZNA LN2SP Migration)" -ForegroundColor DarkMagenta -BackgroundColor White
Write-Host "`n"
Write-Host "###################################################################################################"
Write-Host "This script is using Create_Site.csv. So Please update Create_Site.csv before starting this script." -ForegroundColor Red
Write-Host "###################################################################################################"
<#Write-Host "Press any key to Start this script once you have ready with above mentioned details." -ForegroundColor Green
Write-Host "`n"
#To show the status of execution until and unless we press any key. 
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
#>
Pause
<################################################################>
Write-Host "Starting to Create a Modern Team Site using PnP PowerShell" -ForegroundColor Green
Write-Host "##########################################################"

#Input File Location
$File = $PSScriptRoot + "\Create_Site.csv"
$csv = Import-Csv $File

#Output File Location
$output = $PSScriptRoot + "\Create_Site_Result.csv"
#$csvOutput= Export-Csv $output
#write-output "Tenant_URL,Title,Alias,Source_URL,Status,Start_Date, End_Date" | Out-File -FilePath $output -Append -Encoding ascii 
#$Added= "List Added"     


foreach ($row in $csv)
{
        $TenantURL = $row.Tenant_URL
        $Source_URL=$row.Source_URL
        $Title=$row.Title
        $Alias=$row.Alias
    
    $Start_Date = Get-Date        
    #Login to Tenant_URL
    Connect-SPOService -URL $TenantURL

    #Check if the site collection exists already
    $SiteExists = Get-SPOSite | Where {$_.URL -eq $Source_URL}
     
    #Check if the site exists in the recycle bin
    $SiteExistsInRecycleBin = Get-SPODeletedSite | where {$_.url -eq $Source_URL}
      
    If($SiteExists -ne $null)
    {
        write-host "$($Source_URL) - Website exists already!" -foregroundcolor Yellow
        $Status="$($Source_URL) - Website exists already!"
        pause
    }
    elseIf($SiteExistsInRecycleBin -ne $null)
    {
        write-host "$($Source_URL) - Website exists in the recycle bin!" -foregroundcolor Yellow
        $Status="$($Source_URL) - Website exists in the recycle bin!" 
        pause              
    }
    else
    {
        #sharepoint online create modern site collection powershell
        #New-SPOSite -Url $URL -title $Title -Owner $Owner -StorageQuota $StorageQuota -NoWait -ResourceQuota $ResourceQuota -Template $Template
        
        Disconnect-SPOService
        Connect-PnPOnline -Url $Tenanturl -UseWebLogin
        New-PnPSite -Type TeamSite -Title $Title -Alias $Alias -IsPublic 
        write-host "$($Source_URL) - Website Created Successfully!" -foregroundcolor Green
        $Status="$($Source_URL) - Website Created Successfully!" 
        #$Error=$_.Exception.Message
        Disconnect-PnPOnline
    }
        $End_Date = Get-Date
        Write-Output "$TenantURL,$Title,$Alias,$Source_URL,$Status,$Start_Date,$End_Date" | Out-File -FilePath $Output -Append -Encoding ascii
}

 

 $UserInput = Read-Host -Prompt 'Do you need this execution report in Excel?[Y/N]'
    switch ($UserInput)
      {
      'Y' {
            # Specify the path to the Excel file and the WorkSheet Name
            $FilePath = $PSScriptRoot + "\Create_Site_Result.csv"

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
