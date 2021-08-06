###### SiteSetup ######

###### Provisioning Site columns,Site Content Types,Lists, Adding Site Column To Content Types,Adding Content Type to List #######



$date= Get-Date -format MMddyyyyHHmmss  

start-transcript -path .\Log_$date.txt #Create Log File    

Import-Module -Name SharePointPnPPowerShellOnline

$XMLFile = Read-Host 'Please Enter the Path of ur xml (.xml)'

[xml]$xmlData=Get-Content $XMLFile  # Get content from XML file 



#Enter Program  
[System.Xml.XmlElement]$lists = $xmlData.Inputs.Lists # Lists node

[System.Xml.XmlElement]$sitecolumns=$xmlData.Inputs.SiteColumns # Lists node for Site Columns

  Function Get-ScriptDirectory {

    Split-Path $script:MyInvocation.MyCommand.Path
}

function CreateLists()  #Function To Create Lists
{  
    write-host -ForegroundColor Green "Creating Lists"   
 
      
    foreach($list in $lists.List)  # Loop through List XML node
    {  
        $listTitle=$list.Title  # Get List node parameters
        
        $listURL=$list.URL  
        
        $listTemplate=$list.Template  
          
        $getList=Get-PnPList -Identity $listURL  # Get the list object

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
    
    function CreateSiteColumns() #Function To Create SiteColumns
   {
    [System.Xml.XmlElement]$sitecolumns=$xmlData.Inputs.SiteColumns

    
    foreach($xml in $xmlData.Inputs.SiteColumns.Field)

   {
   
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

           if($xml.Type -ne "Lookup" -and $xml.Type -ne "Calculated" -and $xml.Type -ne "TaxonomyFieldType") #Create all Site Columns except Lookup
        {

        
        $listxml=$xml.OuterXML #$xmlData.Inputs.SiteColumns.Field[$i].OuterXml

        $listfield=$listxml.Replace('Type=','ID="{'+[guid]::NewGuid()+'}" Type=')

        Add-PnPFieldFromXml -FieldXml $listfield

        Write-Host $xml.Name "Column Created"
       
        }
        elseif($xml.Type -eq "Lookup")
        {
                    $lookupList=Get-PnPList -Identity $xml.List
                
                    $lookupListID=$lookupList.Id

                    $Placing
                
                    $web=Get-PnPWeb
                
                    $webID=$web.Id
                
                    $lookupxml= $xml.OuterXml
                
                    $lookupidreplace=$lookupxml.Replace('List="'+$xml.List+'"','List="{'+$lookupListID+'}"')
                
                    $sourceidreplace=$lookupidreplace.Replace('SourceID="{}"','SourceID="{'+$webID+'}"')
                
                    $webidreplace=$sourceidreplace.Replace('WebId=""','WebId="'+$webID+'"')
                
                    $fieldIDreplace=$webidreplace.Replace('Type=','ID="{'+[guid]::NewGuid()+'}" Type=')
                
                    Add-PnPFieldFromXml -FieldXml $fieldIDreplace
                
                    Write-Host $xml.Name "Look Up Column Created"
                
           }
           elseif($xml.Type -eq "Calculated")
           {
                
                    $webCal=Get-PnPWeb
                
                    $webCalID=$webCal.Id

                    $calcXml= $xml.OuterXml

                    $calcIdReplace=$calcXml.Replace('SourceID="{}"','SourceID="{'+$webCalID+'}"')

                    $CalcIDGUIDReplace=$calcIdReplace.Replace('DisplayName=','ID="{'+[guid]::NewGuid()+'}" DisplayName=')
                 
                    
                    foreach($fieldref in $xml.FieldRefs.FieldRef)
                    { 
                    $fieldRefName=$fieldref.Name
                    $fieldrefId=Get-PnPField -Identity $fieldRefName
                    
                    $CalcIDGUIDReplace=$CalcIDGUIDReplace.Replace('FieldRef Name="'+$fieldref.Name+'" ID="{}"','FieldRef Name="'+$fieldref.Name+'" ID="{'+ $fieldrefId.Id +'}"')

                    }
                    
                    Add-PnPFieldFromXml -FieldXml $CalcIDGUIDReplace
                
                    Write-Host $xml.Name "Calculated Column Created"

            }

            elseif($xml.Type -eq "TaxonomyFieldType")
            {

            $TaxXml=$xml.OuterXml
            
            $TaxWeb=Get-PnPWeb
            
            $TaxWebID=$TaxWeb.Id

            $TaxSourceIdreplace=$TaxXml.Replace('SourceID="{}"','SourceID="{'+$TaxWebID+'}"')

            $TaxXmlWebIdReplace=$TaxSourceIdreplace.Replace('WebId=""','WebId="'+$TaxWebID+'"')

            $TaxIDreplace=$TaxSourceIdreplace.Replace('Type=','ID="{'+[guid]::NewGuid()+'}" Type=')

            #Write-Host $TaxXmlWebIdReplace

            Add-PnPFieldFromXml -FieldXml $TaxIDreplace

            Write-Host $xml.Name "Managed Metadata Column Created"
        }
        }
        }
        }
        
       



  

function CreateContentTypes()  #Function To Create Content types

{

foreach($contentType in $xmlData.Inputs.ContentTypes.ContentType)

{

$contentTypeExists=Get-PnPContentType -Identity $contentType.Name 

if($contentTypeExists) #Check if ContentType Already Exits
{
Write-Host -ForegroundColor Magenta $contentType.Name "already Exists"
}

else #Create if ContentType doesn't already Exits
{

$contentName=$contentType.Name

$contentDescription=$contentType.Description

$parentName=$contentType.Parent

$contentTypeID=$contentType.ContentTypeID

$parentCT= Get-PnPContentType -Identity $parentName

$contentTypeGroup=$contentType.Group


Add-PnPContentType -Name $contentName -Description $contentDescription -Group $contentTypeGroup -ContentTypeId $contentTypeID

#Add-PnPContentType -Name $contentName -Description $contentDescription -Group $contentTypeGroup -ParentContentType $parentCT

Write-Host $contentName "Content type Creation Completed"
}
}
}


function AddSiteColumnsToContentType() #Function To Add Site Column To Content types

{

foreach($ctName in $xmlData.Inputs.ContentTypes.ContentType)

{

$preexistingCT=$ctName.Name

$exists=Get-PnPContentType -Identity $preexistingCT

if($exists -ne $null)

{

foreach($ctField in $ctName.Fields.Field)

{

$ctColumnname=$ctField.Name

Write-Host "Adding Column" $ctColumnname  "in" $preexistingCT 

Add-PnPFieldToContentType -Field $ctColumnname -ContentType $exists

}
}
}
}

function AddContentTypetoLists() #Function To Add  Content types to Lists
{

foreach($ctList in $xmlData.Inputs.Lists.List)

{

$listName=$ctList.URL

Write-Host "Adding Content type to" $listName



foreach($contentTypes in $ctList.SiteContentTypes)

{

$contentTypeFlag=$contentTypes.ContentType

Write-Host $contentTypeFlag

foreach($contentType in $contentTypes.ContentType)

{

$ListCTType=$contentType.Name

$ListCTexists=Get-PnPContentType -Identity $ListCTType

if($ListCTexists -ne $null)

{

Add-PnPContentTypeToList -List $listName  -ContentType $ListCTexists

Write-Host "Completed Addition of Content Type" $contentTypes "To" $listName
}
}
}
}
}

function AddSiteColumnsToLists()
{
     
     $clientContext=Get-PnPContext
     $web = $clientContext.Web
     $clientContext.Load($web)
     $clientContext.ExecuteQuery()
    
    foreach($listex in $xmlData.Inputs.Lists.List)
    {
    
     #getting the list from XmlData
     $listname=$listex.Title
     Write-Host "Adding ExistingSitecolumns to" $listname

     

     #loading the list
     $list=$clientContext.Web.Lists.GetByTitle($listname)
     $clientContext.Load($list)
     $clientContext.ExecuteQuery()
   
   

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
    
    try
    {
    $listfield=$listfields.GetByInternalNameOrTitle($fieldname)
    $clientContext.Load($listfield)  
    $clientContext.ExecuteQuery()
    Write-Host -ForegroundColor Magenta "Field" $fieldname "already exists in list"
    }

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

      
      if($list.BaseTemplate -eq '100')
      {
      $view = $views.GetByTitle("All Items")
      $viewFields = $view.ViewFields  
      $viewFields.Add($fieldname)      
      $view.Update()
      $clientContext.ExecuteQuery() 
      }
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

function CreateFolder()
{

foreach($Folder in $xmlData.Inputs.Folders.Folder)
{

$FolderLibrary=$Folder.Library

$FolderName=$Folder.Name

$FolderExistsUrl=$FolderLibrary+"/"+$FolderName

$getLibrary=Get-PnPList -Identity $FolderLibrary  # Check if Library exists

if($getLibrary)  
        {  
           
           try
           {

                $folderexisting=Get-PnPFolder -Url $FolderExistsUrl -ErrorAction Stop

                Write-Host -ForegroundColor Magenta $FolderName "Folder already exists in" $FolderLibrary

           }
           catch
           {

           Add-PnPFolder -Name $FolderName -Folder $FolderLibrary

           Write-Host -ForegroundColor Green $FolderName "folder created in" $FolderLibrary

           }

             }   

        else  
        {  
           
           write-host -ForegroundColor Magenta  $FolderLibrary " Does Not Exist "
                
                
        } 

}


}


function CreateSubFolder()
{


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

                $ParentFolderExisting=Get-PnPFolder -Url $ParentFolderExistsUrl -ErrorAction Stop

                foreach($SubFolder in $ParentFolder.SubFolders)
                {
                    $SubFolderName = $SubFolder.Name

                    $SubFolderLocation=$SubFolder.Folder

                    $SubfolderExistsUrl = $ParentFolderExistsUrl+"/"+$SubFolderName

                    try
                    {

                    $getSubFolderExists=Get-PnPFolder -Url $SubfolderExistsUrl -ErrorAction Stop

                    Write-Host $SubFolderName "already exists in" $SubFolderLocation

                    }
                    catch
                    {
                    
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

function CreatePages()
{

foreach($Page in $xmlData.Inputs.Pages.Page)
{
$PageTitle=$Page.Title

$PageName=$Page.Name

$Layout=$Page.Layout

$PageLocation=$Page.Location

try
{
Add-PnPPublishingPage -PageName $PageName -Title $PageTitle -PageTemplateName $Layout -FolderPath $PageLocation -ErrorAction Stop
Write-Host $PageName "created in" $PageLocation
}
catch
{
Write-Host $PageTitle "Already Exists"
}
}

}

function GetAllRepositoryItems()
{

$camlQuery = @"
<View Scope="RecursiveAll">
    <Query>
        <OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy>
    </Query>
    <ViewFields>
        <FieldRef Name='DocAuthor'/>
        <FieldRef Name='DocumentOwner'/>
    </ViewFields>
    <RowLimit Paged="TRUE">5000</RowLimit>
</View>
"@

$items = Get-PnPListItem -List "Repository" -Query $camlQuery

Write-Host "All non folders loaded for Repository"

Set-PnPListItem -List "Repository" -Identity "99" -Values @{"DocAuthorName" = "20"}

#foreach($item in $items)
#   {
  #  Write-Host $item["ID"]
   # $DocumentTitle=$item["Title"]
    #Write-Host $DocumentTitle
    #$DocAuthor=$item["DocAuthor"].LookupValue
    #Write-Host $DocAuthor
    #Set-PnPListItem -List "Repository" -Identity $item["ID"] -Values @{"DocAuthorName" = $DocAuthor;"ContentStatus" = "Approved"}
    #}

}


function SetCreatedByModifiedBy()
{

$camlQuery = @"
<View Scope="RecursiveAll">
    <Query>
        <Where>
        <And>
       <Neq>
      <FieldRef Name="ContentType"/>
        <Value Type="Text">Folder</Value>
      </Neq>
      <Eq>
            <FieldRef Name='ContentStatus'  />
            <Value Type='Text'>Approved</Value>
        </Eq>
      </And>
      </Where>
        <OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy>
    </Query>
    <ViewFields>
        <FieldRef Name='Title'/>
        <FieldRef Name='ID'/>
        <FieldRef Name='FileRef'/>
        <FieldRef Name='Opco'/>
        <FieldRef Name='IsApprovedMailSent'/>
        <FieldRef Name='Created'/>
        <FieldRef Name='Modified'/>
        <FieldRef Name='Author'/>
        <FieldRef Name='Editor'/>
    </ViewFields>
   <RowLimit Paged="TRUE">1000</RowLimit>
</View>
"@

$Listitems = Get-PnPListItem -List "Repository" -Query $camlQuery 

foreach($Items in $Listitems)
{
        #$FileReference=""

        #$fileNewURL=""

        #$filred=""

       
            #$FileReference=$Items["FileRef"]
            #$filred=$FileReference.Split("/")
            #$fileNewURL="/"+$filred["1"]+"/"+$filred["2"]+"/"+$filred["3"]+"/"+$filred["4"]+"/Robi/"+$filred["5"]
            #Write-Host $fileNewURL
            #Move-PnPFile -ServerRelativeUrl $FileReference -TargetUrl $fileNewURL -Force

        Write-Host $items["Created"]
        Write-Host $items["Author"].LookupValue
        write-Host $items["Modified"]
        Write-Host $items["Editor"]
        $clientcontext=Get-PnPcontext
        $web=$clientcontext.Web
        #$username="oneaxiatakm@axtcloud.onmicrosoft.com"
        #$user=$web.EnsureUser($username)
        #$clientcontext.load($user)
        $list=$web.Lists.GetByTitle("Repository")
        $ListItem=$List.GetItemById($Items["ID"])
        $clientcontext.Load($ListItem)
        
        #$ListItem["Author"]=$items["Author"].LookupId
        #$ListItem["Editor"]=$items["Author"].LookupId
        #$ListItem["IsApprovedMailSent"]="True"
        $ListItem["ApprovedBy"]=$items["Editor"].LookupValue
        $ListItem["ApprovedDate"]=$items["Modified"]
        $ListItem.Update()
        $clientcontext.ExecuteQuery()
        Write-Host "Completed Override for" $items["Title"]
        
        
          

}
}

function SetConfidentiality()
{

$camlQuery = @"
<View Scope="RecursiveAll">
    <Query>
        <Where>
       <Neq>
      <FieldRef Name="ContentType"/>
        <Value Type="Text">Folder</Value>
      </Neq>
      </Where>
        <OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy>
    </Query>
    <ViewFields>
        <FieldRef Name='Title'/>
        <FieldRef Name='ID'/>
        <FieldRef Name='Confidentiality'/>
    </ViewFields>
   <RowLimit Paged="TRUE">5000</RowLimit>
</View>
"@

$Listitems = Get-PnPListItem -List "Repository" -Query $camlQuery

foreach($Items in $Listitems)
{

        Write-Host $items["Title"]
        Write-Host $items["ID"]
        Write-Host $Items["Confidentiality"]

        if($Items["Confidentiality"] -eq "A3")
        {

        $clientcontext=Get-PnPcontext
        $web=$clientcontext.Web
        #$username="oneaxiatakm@axtcloud.onmicrosoft.com"
        #$user=$web.EnsureUser($username)
        #$clientcontext.load($user)
        $list=$web.Lists.GetByTitle("Repository")
        $ListItem=$List.GetItemById($Items["ID"])
        $clientcontext.Load($ListItem)
        #$ListItem["Author"]=$user
        #$ListItem["Editor"]=$user
        Write-Host $Listitem["Confidentiality"]
        $Listitem["Confidentiality"]="Public"
        $ListItem.Update()
        $clientcontext.ExecuteQuery()
        Write-Host "Completed Override for" $items["Title"]
        }

        
        if($Items["Confidentiality"] -eq "A2")
        {

        $clientcontext=Get-PnPcontext
        $web=$clientcontext.Web
        #$username="oneaxiatakm@axtcloud.onmicrosoft.com"
        #$user=$web.EnsureUser($username)
        #$clientcontext.load($user)
        $list=$web.Lists.GetByTitle("Repository")
        $ListItem=$List.GetItemById($Items["ID"])
        $clientcontext.Load($ListItem)
        #$ListItem["Author"]=$user
        #$ListItem["Editor"]=$user
        Write-Host $Listitem["Confidentiality"]
        $Listitem["Confidentiality"]="Restricted to Function or Theme"
        $ListItem.Update()
       $clientcontext.ExecuteQuery()
        Write-Host "Completed Override for" $items["Title"]
        }

        if($Items["Confidentiality"] -eq "A1")
        {

        $clientcontext=Get-PnPcontext
        $web=$clientcontext.Web
        #$username="oneaxiatakm@axtcloud.onmicrosoft.com"
        #$user=$web.EnsureUser($username)
        #$clientcontext.load($user)
        $list=$web.Lists.GetByTitle("Repository")
        $ListItem=$List.GetItemById($Items["ID"])
        $clientcontext.Load($ListItem)
        #$ListItem["Author"]=$user
        #$ListItem["Editor"]=$user
        Write-Host $Listitem["Confidentiality"]
        $Listitem["Confidentiality"]="Restricted to Senior Management"
        $ListItem.Update()
        $clientcontext.ExecuteQuery()
        Write-Host "Completed Override for" $items["Title"]
        }
        
          

}
}

function UserProfilePropertyUpdate()
{
    $UserName='i:0#.f|membership|lokranjan6@portalknow.onmicrosoft.com'
    #Get-PnPUserProfileProperty -Account $UserName
    #Set-PnPUserProfileProperty -Account $UserName -Property 'EMail' -Value 'Priscilla.Amos@cognizant.com'
    Set-PnPUserProfileProperty -Account $UserName -Property 'Email' -Value 'Lokranjan06@PortalKnow.onmicrosoft.com'
}

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

function AddUserstoGroupbyXML()
{
    foreach($User in $xmlData.Inputs.Groups.Group.User)
    {

        $UserGroup = $User.Group
        $UserEmail =$User.Email
        try
        {
        Add-PnPUserToGroup -Identity $UserGroup -EmailAddress $UserEmail -ErrorAction Stop
        }
        catch
        {
          Write-Host $UserGroup
        }
    }
}

function ImportDatafromExcel()
{
    $csvFile = Read-Host 'Please Enter the Path of ur Excel (.csv)'
    $Data = Import-Csv $csvFile
    foreach($listdata in $Data)
    {
       Add-PnPListItem -List "ImportData" -ContentType "Item" -Values @{"Title" = $listdata.Title; "ImportValues"=$listdata.Value}
       Write-Host "List Data Value" $listdata.Value "added to ImportData List"
    }

    #Get-PnPUserProfileProperty -Account 'abhay.nidhi@ncell.axiata.com'
  
  }

  function getusersfromdl
  {
    
    #Get-DistributionGroup -Identity "SP_Celcom" | Format-List
  }

  function deletelistitems()
  {
    $DataFunction="Strategy"
    $camlQuery = @"
<View Scope="RecursiveAll">
    <Query>
        <Where>
  <Eq>
    <FieldRef Name="Author" LookupId="True" />
    <Value Type="Lookup">311</Value>
  </Eq>
</Where>
        <OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy>
    </Query>
    <ViewFields>
        <FieldRef Name='Title'/>
        <FieldRef Name='ID'/>
        <FieldRef Name='Function'/>
    </ViewFields>
   <RowLimit Paged="TRUE">5000</RowLimit>
</View>
"@

$Listitems = Get-PnPListItem -List "Axiata KM Discussion" -Query $camlQuery

foreach($Items in $Listitems)
{
    $varfunc=$Items["Function"].LookupValue
    $varitemId=$Items["ID"]

    if($varfunc -eq "strategy")
    {
    Remove-PnPListItem -List "Axiata KM Discussion" -Identity $varitemId -Force
    }
}

  }

  function UploadFiletoSharepoint()
  {
    $path="https://axtcloud.sharepoint.com/sites/oneaxtkmsit/Repository/Technology/Public/Future%20Computer.jpg"
    
    Add-PnPFile -Path $path -Folder "Repository" -Values @{Modified="1/1/2016"}
  }

  function CopyFolder()
{

    $DocumentSource="Shared Documents/A2"
    $TargetSource="/sites/Teacher/Shared Documents"

    try
    {
    Copy-PnPFile -SourceUrl $DocumentSource -TargetUrl $TargetSource  -OverwriteIfAlreadyExists -Force
    Write-Host "Action Completed for Copying"
    }
    catch
    {
        Write-Host "Error Occured During Copy"
    }
}

function RenamecolumsExcel()

{
    $csvFile = Read-Host 'Please Enter the Path of ur Excel (.csv)'
    $Data = Import-Csv $csvFile
    foreach($listdata in $Data)
    {
        Write-Host "Internal Name "$listdata.InternalColumnName"Display Name "$listdata.DisplayName
        try
        {
        Set-PnPField -List "NameChangeList" -Identity $listdata.InternalColumnName -Values @{Title=$listdata.DisplayName} -ErrorAction Stop
        }
        catch
        {
        Write-Host "Error occured"
        }
    }
  
  }

  function CreateWebPartPages()
  {
    $csvFile = "C:\Users\rk901\Documents\PagesWebparts.csv"
    $Data = Import-Csv $csvFile
    foreach($listdata in $Data)
    {
    $PageName=$listdata.PageName
    $WebpartName=$listdata.WebPartName
    Write-Host $PageName
    Write-Host $WebpartName
    Add-PnPClientSidePage -Name $PageName -Publish
    Add-PnPClientSideWebPart -Page $PageName -Component $WebpartName
    }

  }


  function CreateListColumns()
  {
    [System.Xml.XmlElement]$Listcolumns=$xmlData.Inputs.Lists


    foreach($List in $Listcolumns.List)
    {

        $ListName=$List.URL
        Write-Host $ListName

        foreach($Fields in $List.AddListColumns.Field)

        {

            $FieldFlag=$Fields.OuterXML

        try
        {

        $SCexists=Get-PnPField -Identity $Fields.Name -List $ListName  -ErrorAction Stop

        
        Write-Host $Fields.Name "already Exists"

        

        }
        catch
        {
         if($Fields.Type -ne "Lookup" -and $Fields.Type -ne "Calculated" -and $Fields.Type -ne "TaxonomyFieldType") #Create all Site Columns except Lookup
        {

        $listxml=$Fields.OuterXML #$xmlData.Inputs.SiteColumns.Field[$i].OuterXml

        $listfield=$listxml.Replace('Type=','ID="{'+[guid]::NewGuid()+'}" Type=')

        Add-PnPFieldFromXml -FieldXml $listfield -List $ListName

        Write-Host $Fields.Name "Column Created"
       
        }
        elseif($Fields.Type -eq "Lookup")
        {
                    $lookupList=Get-PnPList -Identity $Fields.List

                    $listtoaddin=Get-PnPList -Identity $ListName

                    $listtoaddinID=$listtoaddin.Id
                
                    $lookupListID=$lookupList.Id

                    Write-Host $lookupList.Id
                
                    $web=Get-PnPWeb
                
                    $webID=$web.Id
                
                    $lookupxml= $Fields.OuterXml
                
                    $lookupidreplace=$lookupxml.Replace('List="'+$Fields.List+'"','List="{'+$lookupListID+'}"')
                
                    $sourceidreplace=$lookupidreplace.Replace('SourceID="{}"','SourceID="{'+$webID+'}"')
                
                    $webidreplace=$sourceidreplace.Replace('WebId=""','WebId="'+$webID+'"')
                
                    $fieldIDreplace=$webidreplace.Replace('Type=','ID="{'+[guid]::NewGuid()+'}" Type=')
                
                    Add-PnPFieldFromXml -FieldXml $fieldIDreplace -List $listtoaddinID

                    #Add-PnPField -Type Lookup -List $ListName -DisplayName $Fields.DisplayName -InternalName $Fields.Name -FieldOptions
                
                    Write-Host $Fields.Name "Look Up Column Created"
                
           }
           elseif($xml.Type -eq "Calculated")
           {
                
                    $webCal=Get-PnPWeb
                
                    $webCalID=$webCal.Id

                    $calcXml= $Fields.OuterXml

                    $calcIdReplace=$calcXml.Replace('SourceID="{}"','SourceID="{'+$webCalID+'}"')

                    $CalcIDGUIDReplace=$calcIdReplace.Replace('DisplayName=','ID="{'+[guid]::NewGuid()+'}" DisplayName=')
                 
                    
                    foreach($fieldref in $Fields.FieldRefs.FieldRef)
                    { 
                    $fieldRefName=$fieldref.Name
                    $fieldrefId=Get-PnPField -Identity $fieldRefName
                    
                    $CalcIDGUIDReplace=$CalcIDGUIDReplace.Replace('FieldRef Name="'+$fieldref.Name+'" ID="{}"','FieldRef Name="'+$fieldref.Name+'" ID="{'+ $fieldrefId.Id +'}"')

                    }
                    
                    Add-PnPFieldFromXml -FieldXml $CalcIDGUIDReplace -List $ListName
                
                    Write-Host $Fields.Name "Calculated Column Created"

            }

            elseif($Fields.Type -eq "TaxonomyFieldType")
            {

            $TaxXml=$Fields.OuterXml
            
            $TaxWeb=Get-PnPWeb
            
            $TaxWebID=$TaxWeb.Id

            $TaxSourceIdreplace=$TaxXml.Replace('SourceID="{}"','SourceID="{'+$TaxWebID+'}"')

            $TaxXmlWebIdReplace=$TaxSourceIdreplace.Replace('WebId=""','WebId="'+$TaxWebID+'"')

            $TaxIDreplace=$TaxSourceIdreplace.Replace('Type=','ID="{'+[guid]::NewGuid()+'}" Type=')

            #Write-Host $TaxXmlWebIdReplace

            Add-PnPFieldFromXml -FieldXml $TaxIDreplace -List $ListName

            Write-Host $Fields.Name "Managed Metadata Column Created"
        }
        }

    
    }

     }   
   }

  function CreateListViews()
{

     $Context = Get-PnPContext

foreach($ViewName in $xmlData.Inputs.Views.View)

{

$ViewsName=$ViewName.Name
$ListName=$ViewName.ListName
$CamlQuery=$ViewName.CamlQuery.InnerXml

Write-Host $ViewsName
Write-Host $ListName

$AddingFields=""

foreach($ViewField in $ViewName.Fields.Field)

{

$ViewFieldName=$ViewField.Name

if($AddingFields -eq "")
{
  $AddingFields=$ViewFieldName
  Add-PnPView -List $ListName -Title $ViewsName -Fields $AddingFields -Query $CamlQuery
  
}

else
{
    $AddingFields=$ViewFieldName
    $ListView  =  Get-PnPView -List $ListName -Identity $ViewsName -ErrorAction Stop
    if($ListView.ViewFields -notcontains $AddingFields)
    {
        #Add Column to View
        $ListView.ViewFields.Add($AddingFields)
        $ListView.Update()
        $Context.ExecuteQuery()
        Write-host -f Green "Column '$AddingFields' Added to View '$ViewsName'!"
    }
}

}
}

}

   function DeleteSiteColumns()
   {
    
    [System.Xml.XmlElement]$sitecolumns=$xmlData.Inputs.SiteColumns

    
    foreach($xml in $xmlData.Inputs.SiteColumns.Field)

   {
   
        try
        {

        $SCexists=Get-PnPField -Identity $xml.Name -ErrorAction Stop
        Remove-PnPField -Identity $xml.Name -Force
        }
        catch
        {
            Write-Host "Field Does Not Exist"
        }

   }
   }

   function DeleteContentTypes()
   {
    foreach($contentType in $xmlData.Inputs.ContentTypes.ContentType)
    {


#$contentTypeExists=Get-PnPContentType -Identity $contentType.Name 

#Write-Host -ForegroundColor Magenta $contentType.Name "already Exists"
Remove-PnPContentType -Identity $contentType.Name -Force

}
   }

   function DeleteLists()
   {

    foreach($list in $lists.List)  # Loop through List XML node
    {  
        $listTitle=$list.Title  # Get List node parameters
        
        $listURL=$list.URL  
        
        $listTemplate=$list.Template  
          
        $getList=Get-PnPList -Identity $listURL  # Get the list object

        # Check if list exists  
        if($getList)  
        {  
           write-host -ForegroundColor Magenta $listURL " - List already exists" 
             Remove-PnPList -Identity $listURL -Force
             }   
        else  
        {  
           # Create new list  
           write-host -ForegroundColor Magenta "Non existing list: " $listURL
                
        }          
    }
    
   }


   function RemovalofcolumsExcel()

{
    $csvFile = "C:\Users\rk901\Documents\RemovalDCColumns.csv"
    $Data = Import-Csv $csvFile
    foreach($listdata in $Data)
    {
        
        try
        {
        Remove-PnPField -Identity $listdata.RemovalColumns -List "Document Control" -Force
        Write-Host "Internal Name "$listdata.RemovalColumns
        }
        catch
        {
        Write-Host "Error occured"
        }
    }
  
  }


  function RemovalofcolumsExcelContentType()

{
    $csvFile = "C:\Users\rk901\Documents\RemovalColumns.csv"
    $Data = Import-Csv $csvFile
    foreach($listdata in $Data)
    {
        
        try
        {
        Remove-PnPFieldFromContentType -Field $listdata.RemovalColumns -ContentType "	Paper Distribution Record"
        }
        catch
        {
        Write-Host "Error occured"
        }
    }
  
  }


    function AdditionofcolumsExcel()

{
    $csvFile = $csvFile = Read-Host 'Please Enter the Path of ur Excel (.csv)'
    $Data = Import-Csv $csvFile
    foreach($listdata in $Data)
    {
        
        try
        {
         Add-PnPField -Type Number -InternalName $listdata.RemovalColumns  -DisplayName $listdata.DispalyNames -Group "Document Control Columns"
         Write-Host $listdata.RemovalColumns "is Created"
        }
        catch
        {
        Write-Host "Error occured"
        }
    }
  
  }

  function RemovalofcolumsfromContentType()

{
    $csvFile = "C:\Users\ri251\Documents\DCSchema\Site URL.csv"
    $Data = Import-Csv $csvFile
    foreach($listdata in $Data)
    {
        Remove-PnPFieldFromContentType -Field $listdata.RemovalColumns -ContentType "Specification Control"
        Write-Host "Internal Name "$listdata.RemovalColumns
    }
  
  }
  
  
  function LoopSiteUrl()

{
    $csvFile = "C:\Users\ri251\Documents\DCSchema\Site URL.csv"
    $credentials= Get-Credential
    $siteurls = Import-Csv $csvFile
    foreach($siteurl in $siteurls)
    {
        
        Connect-PnPOnline -Url $siteurl.SiteURL -Credentials $credentials

        Write-Host "Connected Site:" $siteurl.SiteURL

        Initiate
    }
  
  }
  
  function TestWrite()
  {
    Write-Host "Test Passed"
  } 

  

function Initiate()  
{  
     write-host -ForegroundColor Green "Initiating the script"   
 
 
     # Call the required functions
      
     #CreateLists
     
     #CreateSiteColumns

     #CreateContentTypes

     #AddSiteColumnsToContentType

     #AddContentTypetoLists

     AddSiteColumnsToLists

     #CreateListColumns

     #CreateListViews

     #CreatePages


     #CreateFolder

     #CreateSubFolder

     #GetAllRepositoryItems

     

     #SetCreatedByModifiedBy

     #SetConfidentiality

     #UserProfilePropertyUpdate
     
     #CreateGroupandAddPermissions

     #AddUserstoGroupbyXML

     #ImportDatafromExcel

     #getusersfromdl

     #deletelistitems

     #UploadFiletoSharepoint

     #CopyFolder

     #RenamecolumsExcel

     #CreateWebPartPages

     #DeleteContentTypes

     #DeleteLists

     #DeleteSiteColumns

     #RemovalofcolumsExcel

     #RemovalofcolumsExcelContentType

     #AdditionofcolumsExcel

     #RemovalofcolumsfromContentType

     #TestWrite

     
     Disconnect-PnPOnline # Disconnecting from the server
      
     write-host -ForegroundColor Green "Completed!!!!"   
}  
LoopSiteUrl  
Stop-Transcript