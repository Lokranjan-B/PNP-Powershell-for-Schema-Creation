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

# Gets the Lists XML node

[System.Xml.XmlElement]$lists = $xmlData.Inputs.Lists 

 #Function To Create Lists

function CreateLists() 
{  
    write-host -ForegroundColor Green "Creating Lists"   
 
 # Loop through List XML node
      
    foreach($list in $lists.List)  
    {  
        # Get List node parameters
        
        $listTitle=$list.Title  
        
        $listURL=$list.URL  
        
        $listTemplate=$list.Template 
        
        # Get the list object 
          
        $getList=Get-PnPList -Identity $listURL  

        # Check if list exists  
        
        if($getList)  
        {  
           write-host -ForegroundColor Magenta $listURL " - List already exists" 
		}   
        else  
        {  
           # Create new list  

           write-host -ForegroundColor Magenta "Creating list: " $listURL  
		   
           New-PnPList -Title $listTitle -Url $listURL -Template $listTemplate
		   
        }          
    }
    }

    function Initiate()  
{  
     write-host -ForegroundColor Green "Initiating the script"   
 
 
     # Calling the required functions
       
     CreateLists

     Disconnect-PnPOnline # Disconnecting from the server
      
     write-host -ForegroundColor Green "Completed!!!!"   
}  
Initiate  
Stop-Transcript