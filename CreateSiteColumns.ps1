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

#Function To Create SiteColumns

function CreateSiteColumns() 
   {

  #gets the xml elements of the Site Columns Described in the XML

    [System.Xml.XmlElement]$sitecolumns=$xmlData.Inputs.SiteColumns

    #Beginning of the Loop to Create Site Column
    
    foreach($xml in $xmlData.Inputs.SiteColumns.Field)

   {
   # Checks if the Site Column Already Exists
        try
        {

        $SCexists=Get-PnPField -Identity $xml.Name -ErrorAction Stop

        if($SCexists)
        {
        Write-Host $xml.Name "already Exists"

        }

        }
        catch
        {

        #Section for creating Column Types except lookup or Calculated or taxonomy column

	    if($xml.Type -ne "Lookup" -and $xml.Type -ne "Calculated" -and $xml.Type -ne "TaxonomyFieldType") #Create all Site Columns except Lookup
        {

        #getting the OuterXML of the Column

        $listxml=$xml.OuterXML 

        #Generating ID for the Column that is to be created since ID is a key parameter for creating the column

        $listfield=$listxml.Replace('Type=','ID="{'+[guid]::NewGuid()+'}" Type=')

        Add-PnPFieldFromXml -FieldXml $listfield

        Write-Host $xml.Name "Column Created"
       
        }

        #Section for creating lookup column
        
        elseif($xml.Type -eq "Lookup")
        {
        #Get the LookUp List Metadata

                    $lookupList=Get-PnPList -Identity $xml.List

                    #Get the LookUp List GUID
                
                    $lookupListID=$lookupList.Id

                    #Get Web of the Particular Site
                
                    $web=Get-PnPWeb

                    #Getting Web Id of the Site
                
                    $webID=$web.Id

                    #Getting the Column's outer xml
                
                    $lookupxml= $xml.OuterXml

                    #Dynamically placing the Look Up list ID
                
                    $lookupidreplace=$lookupxml.Replace('List="'+$xml.List+'"','List="{'+$lookupListID+'}"')

                    #Dynamically placing the Source ID
                
                    $sourceidreplace=$lookupidreplace.Replace('SourceID="{}"','SourceID="{'+$webID+'}"')

                    #Dynamically placing the Web ID (Source and Web Id are the same)
                
                    $webidreplace=$sourceidreplace.Replace('WebId=""','WebId="'+$webID+'"')

                    #Generating ID for the Column that is to be created since ID is a key parameter for creating the column
                
                    $fieldIDreplace=$webidreplace.Replace('Type=','ID="{'+[guid]::NewGuid()+'}" Type=')

                    #Creating the Look Up Column
                
                    Add-PnPFieldFromXml -FieldXml $fieldIDreplace
                
                    Write-Host $xml.Name "Look Up Column Created"
                
           }

           
        #Section for creating calculated column

           elseif($xml.Type -eq "Calculated")
           {
                #Get Web of the Particular Site

                    $webCal=Get-PnPWeb

                    #Getting Web Id of the Site

                
                    $webCalID=$webCal.Id

                     #Getting the Column's outer xml


                    $calcXml= $xml.OuterXml

                    #Dynamically placing the Source ID

                    $calcIdReplace=$calcXml.Replace('SourceID="{}"','SourceID="{'+$webCalID+'}"')

                    #Generating ID for the Column that is to be created since ID is a key parameter for creating the column

                    $CalcIDGUIDReplace=$calcIdReplace.Replace('DisplayName=','ID="{'+[guid]::NewGuid()+'}" DisplayName=')
                 
                    #Mapping the ID of the Source Field of the Calculated Column
                    
                    foreach($fieldref in $xml.FieldRefs.FieldRef)
                    { 
                    $fieldRefName=$fieldref.Name
                    $fieldrefId=Get-PnPField -Identity $fieldRefName
                    
                    $CalcIDGUIDReplace=$CalcIDGUIDReplace.Replace('FieldRef Name="'+$fieldref.Name+'" ID="{}"','FieldRef Name="'+$fieldref.Name+'" ID="{'+ $fieldrefId.Id +'}"')

                    }
                    
                    Add-PnPFieldFromXml -FieldXml $CalcIDGUIDReplace
                
                    Write-Host $xml.Name "Calculated Column Created"

            }

            
        #Section for ManagedMetaData column

            elseif($xml.Type -eq "TaxonomyFieldType")
            {

            $TaxXml=$xml.OuterXml

            #Get Web of the Particular Site
            
            $TaxWeb=Get-PnPWeb

            #Getting Web Id of the Site
            
            $TaxWebID=$TaxWeb.Id

            #Mapping the ID of the Source Field of the Calculated Column

            $TaxSourceIdreplace=$TaxXml.Replace('SourceID="{}"','SourceID="{'+$TaxWebID+'}"')

            #Dynamically placing the Web ID (Source and Web Id are the same)

            $TaxXmlWebIdReplace=$TaxSourceIdreplace.Replace('WebId=""','WebId="'+$TaxWebID+'"')

            #Generating ID for the Column that is to be created since ID is a key parameter for creating the column

            $TaxIDreplace=$TaxSourceIdreplace.Replace('Type=','ID="{'+[guid]::NewGuid()+'}" Type=')

            Add-PnPFieldFromXml -FieldXml $TaxIDreplace

            Write-Host $xml.Name "Managed Metadata Column Created"
        }
        }
        }
        }



function Initiate()  
{  
     write-host -ForegroundColor Green "Initiating the script"   
 
 
     # Calling the required functions
       
     CreateSiteColumns

     Disconnect-PnPOnline # Disconnecting from the server
      
     write-host -ForegroundColor Green "Completed!!!!"   
}  
Initiate  
Stop-Transcript