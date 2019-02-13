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

#Function for Creating SubFolders

function CreateSubFolder()
{

#Looping Starts for creating Site Columns

foreach($ParentFolder in $xmlData.Inputs.Folders.Folder)
{

$ParentFolderLibrary=$ParentFolder.Library

$ParentFolderName=$ParentFolder.Name

$ParentFolderExistsUrl=$ParentFolderLibrary+"/"+$parentFolderName

$getParentLibrary=Get-PnPList -Identity $ParentFolderLibrary  # Check if Library exists

 if($getParentLibrary)  
        {  
           
           try
           {

                #Checking if folder Exists

                $ParentFolderExisting=Get-PnPFolder -Url $ParentFolderExistsUrl -ErrorAction Stop

                foreach($SubFolder in $ParentFolder.SubFolders)
                {
                    $SubFolderName = $SubFolder.Name

                    $SubFolderLocation=$SubFolder.Folder

                    #Checking if sub folder Exists

                    $SubfolderExistsUrl = $ParentFolderExistsUrl+"/"+$SubFolderName

                    try
                    {

                    $getSubFolderExists=Get-PnPFolder -Url $SubfolderExistsUrl -ErrorAction Stop

                    Write-Host $SubFolderName "already exists in" $SubFolderLocation

                    }
                    catch
                    {
                    
                    #Creating SubFolder inside Folder (Currently written only for single branch)

                    Add-PnPFolder -Name $SubFolderName -Folder $ParentFolderExistsUrl

                    Write-Host -ForegroundColor Green $SubFolderName "folder created in" $SubFolderLocation

                    }

                }

           }
           catch
           {

           Write-Host -ForegroundColor Red $ParentFolderName "Does Not Exist in" $ParentFolderLibrary 
           
           }

		}   

        else  
        {  
           
           write-host -ForegroundColor Magenta  $ParentFolderLibrary " Does Not Exist "
		   
		   
        } 

}

}



function Initiate()  
{  
     write-host -ForegroundColor Green "Initiating the script"   
 
 
     # Call the required functions
       
     
     CreateSubFolder
    
     Disconnect-PnPOnline # Disconnecting from the server
      
     write-host -ForegroundColor Green "Completed!!!!"   
}  
Initiate  
Stop-Transcript