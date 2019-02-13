###### SiteSetup ######

###### Provisioning Site columns,Site Content Types,Lists, Adding Site Column To Content Types,Adding Content Type to List etc. #######

$date= Get-Date -format MMddyyyyHHmmss  

#Creates a Log File with the Log of Execution of the scripts

start-transcript -path .\Log_$date.txt

#  User import module if powershell commands are not supported by default in your system

import-module ("D:\My Files\Binaries\SPOnline\Modules\SharePointPnPPowerShellOnline\SharePointPnPPowerShellOnline.psd1")
        Add-Type -Path ("D:\My Files\Binaries\SPOnline\Modules\SharePointPnPPowerShellOnline\SharePointPnP.PowerShell.Online.Commands.dll")
        Add-Type -Path ("D:\My Files\Binaries\SPOnline\Modules\SharePointPnPPowerShellOnline\Microsoft.SharePoint.Client.dll")
        Add-Type -Path ("D:\My Files\Binaries\SPOnline\Modules\SharepointPnPPowerShellOnline\Microsoft.SharePoint.Client.Runtime.dll")

#Provide the path where your XML File exists for Reading the XML Schema

$XMLFile = Read-Host 'Please Enter the Path of ur xml (.xml)'

#Gets the XML File and reads all the data in the XML File

[xml]$xmlData=Get-Content $XMLFile 

#Connecting to the online site and getting user Credentitals

$siteURL=Read-Host "Enter URL" #Connect to site

$credentials= Get-Credential

Connect-PnPOnline -Url $siteURL -Credentials $credentials

#Enter Program  

#Function for Creating Pages

function CreatePages()
{

#Loop Begins for Creating Publishing Pages

foreach($Page in $xmlData.Inputs.Pages.Page)
{
$PageTitle=$Page.Title

$PageName=$Page.Name

$Layout=$Page.Layout

$PageLocation=$Page.Location

try
{

#Add Publishing Pages to the Site

Add-PnPPublishingPage -PageName $PageName -Title $PageTitle -PageTemplateName $Layout -FolderPath $PageLocation -ErrorAction Stop
Write-Host $PageName "created in" $PageLocation
}
catch
{
Write-Host $PageTitle "Already Exists"
}
}

}



function Initiate()  
{  
     write-host -ForegroundColor Green "Initiating the script"   
 
 
     # Call the required functions
       
     
     CreatePages
    
     Disconnect-PnPOnline # Disconnecting from the server
      
     write-host -ForegroundColor Green "Completed!!!!"   
}  
Initiate  
Stop-Transcript