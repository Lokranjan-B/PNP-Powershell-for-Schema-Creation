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

#Function for Creating Groups

function CreateGroupandAddPermissions()
{

   foreach($Group in $xmlData.Inputs.Groups.Group)
   {
    
    $GroupName=$Group.Name
    $GroupDescription=$Group.Description
    $GroupOwner = $Group.Owner
    $UserPermission =$Group.PermissionLevel
    #Check if Group already exists
    try
    {
        Get-PnPGroup -Identity $GroupName -ErrorAction Stop
        Write-Host "Group Already Exists"
    }
    catch
    {

        #Create the new group 
        New-PnPGroup -Title $GroupName -Description $GroupDescription -Owner $GroupOwner -AllowMembersEditMembership

        #Add Permission level to the group

        Write-Host "Added the new group "$GroupName

        Set-PnPGroupPermissions -Identity $GroupName -AddRole $UserPermission

        Write-Host "Added the Permssion Level to" $GroupName
    }
   }
}




function Initiate()  
{  
     write-host -ForegroundColor Green "Initiating the script"   
 
 
     # Call the required functions
       
     
     CreateGroupandAddPermissions
    
     Disconnect-PnPOnline # Disconnecting from the server
      
     write-host -ForegroundColor Green "Completed!!!!"   
}  
Initiate  
Stop-Transcript