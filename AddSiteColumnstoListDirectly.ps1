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

#Adding Site Columns directly to the list

function AddSiteColumnsToLists()
{
     
     #Get PnP Context

     $clientContext=Get-PnPContext

     #Get PnP Web

     $web = $clientContext.Web
     #Load Web and get the site values
     $clientContext.Load($web)
     $clientContext.ExecuteQuery()
    #Get the Listname  and Loading the List 
    foreach($listex in $xmlData.Inputs.Lists.List)
    {
    
     #getting the list from XmlData
     $listname=$listex.Title
     Write-Host "Adding ExistingSitecolumns to" $listname

     

     #loading the list
     $list=$clientContext.Web.Lists.GetByTitle($listname)
     $clientContext.Load($list)
     $clientContext.ExecuteQuery()
   
   #Getting the Site and List Fields

    foreach($field in $listex.Fields.Field )
    {

     #getting the fieldnames from the Xmldata
     $fieldname=$field.Name
     
    
 #loading the list fields

     $listfields=$list.Fields
     $clientContext.Load($listfields)

    $sitefields = $web.Fields 
    $clientContext.Load($sitefields) 
    $clientContext.ExecuteQuery() 

    $webfield=$sitefields.GetByInternalNameOrTitle($fieldname)
    $clientContext.Load($webfield) 
    $clientContext.ExecuteQuery()
    #Check if the Column already exists in the List mentioned
    try
    {
    $listfield=$listfields.GetByInternalNameOrTitle($fieldname)
    $clientContext.Load($listfield)  
    $clientContext.ExecuteQuery()
    Write-Host -ForegroundColor Magenta "Field" $fieldname "already exists in list"
    }
    #Add the Columns to the List and add them to all items view 
     catch
    {
   # $list.Fields.AddFieldAsXml($webfield.SchemaXml,$true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
    $list.Fields.Add($webfield)
    $list.Update()
    $clientContext.ExecuteQuery() 
 
     #adding the columns to default "All Items"View
      $views = $list.Views
      $clientContext.Load($views)
       $clientContext.ExecuteQuery() 
    #Add to all items view for List
      
      if($list.BaseTemplate -eq '100')
      {
      $view = $views.GetByTitle("All Items")
      $viewFields = $view.ViewFields  
      $viewFields.Add($fieldname)      
      $view.Update()
      $clientContext.ExecuteQuery() 
      }
      #Add to all items view for Library
      else
      {
      $view=$views.GetByTitle("All Documents")
      $viewFields = $view.ViewFields  
      $viewFields.Add($fieldname)      
      $view.Update()
      $clientContext.ExecuteQuery()
      }
      
   Write-Host  $fieldname "Field Added" "to" $listname
    }
    }
    }

}

function Initiate()  
{  
     write-host -ForegroundColor Green "Initiating the script"   
 
 
     # Call the required functions
       
     
     AddSiteColumnsToLists
    
     Disconnect-PnPOnline # Disconnecting from the server
      
     write-host -ForegroundColor Green "Completed!!!!"   
}  
Initiate  
Stop-Transcript