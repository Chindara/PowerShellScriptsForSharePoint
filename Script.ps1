#Add references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#Specify tenant admin and site URL
$User = Read-Host -Prompt "Please enter your username"
$Password = Read-Host -Prompt "Please enter your password" -AsSecureString
$SiteURL = "Your SharePoint Site Url"

#Get Credentials to connect
$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User,$Password)
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
$Context.Credentials = $Creds

Function Get-ListIdByName
{
    param
    (
        [Parameter(Mandatory=$true)] [string]$ListName = $(throw "Please Enter the List Name!")
    )

    Try {
        $List = $Context.Web.Lists.GetByTitle($ListName)
        return $List.Id;
    }
    Catch {
        write-host -f Red "Error Creating List!" $_.Exception.Message
    }
}

Function New-SPOList
{
    param
    (
        [Parameter(Mandatory=$true)] [string]$ListName = $(throw "Please Enter the List Name!"),
        [Parameter(Mandatory=$true)] [string]$ListTitle = $(throw "Please Enter the Title for the List!")
    )

	Try {
		#Retrieve lists
		$Lists = $Context.Web.Lists
		$Context.Load($Lists)
		$Context.ExecuteQuery()

		#Create list with "custom" list template
		$ListInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
		$ListInfo.Title = $ListName
		$ListInfo.TemplateType = "100"
		$List = $Context.Web.Lists.Add($ListInfo)
		$List.Description = $ListName
		$List.Update()
		$Context.ExecuteQuery()

		#Get the List
        $List=$Context.Web.Lists.GetByTitle($ListName) 
 
        $List.Title = $ListTitle
        $List.Update()
        $Context.ExecuteQuery()
             
        Write-Host "Success, List: $ListName" -f Green
	}
	Catch {
        write-host -f Red "Error Creating List!" $_.Exception.Message
    }
}

Function Add-SPOListColumn-Text
{
    param
    (
        [Parameter(Mandatory=$true)] [string]$ListName = $(throw "Please Enter the List Name!"),
        [Parameter(Mandatory=$true)] [string]$FieldName = $(throw "Please Enter the FieldName!"),
        [Parameter(Mandatory=$true)] [string]$DisplayName = $(throw "Please Enter the DisplayName!"),
        [Parameter(Mandatory=$true)] [string]$Description = $(throw "Please Enter the Description!"),
        [Parameter(Mandatory=$true)] [boolean]$IsRequired = $(throw "Please Enter the IsRequired!")
    )
    Try {
        #Get the List
        $List = $Context.Web.Lists.GetByTitle($ListName)
        $Fields = $List.Fields
        $Context.Load($List)
        $Context.Load($Fields)
        $Context.ExecuteQuery()

        $NewField = $Fields | where { ($_.Internalname -eq $Name) -or ($_.Title -eq $DisplayName) }
 
        if($NewField -ne $NULL)
        {
             Write-host "Field '$FieldName' already exists in the List!" -f Yellow
        }     
        else
        {
            $FieldSchema = "<Field Type='Text' DisplayName='$DisplayName' Name='$FieldName' Description='$Description' Required='$IsRequired' ></Field>"
            $NewField = $List.Fields.AddFieldAsXml($FieldSchema, $true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
            $List.Update()
            $Context.ExecuteQuery()
            Write-host "Success, List: $ListName, Field: $FieldName, Type: Text" -f Green 
        } 
    }
    Catch {
        write-host -f Red "Error Adding Column!" $_.Exception.Message
    }
}

Function Add-SPOListColumn-MultiText
{
    param
    (
        [Parameter(Mandatory=$true)] [string]$ListName = $(throw "Please Enter the List Name!"),
        [Parameter(Mandatory=$true)] [string]$FieldName = $(throw "Please Enter the FieldName!"),
        [Parameter(Mandatory=$true)] [string]$DisplayName = $(throw "Please Enter the DisplayName!"),
        [Parameter(Mandatory=$true)] [string]$Description = $(throw "Please Enter the Description!"),
        [Parameter(Mandatory=$true)] [boolean]$IsRequired = $(throw "Please Enter the IsRequired!")
    )
    Try {
        #Get the List
        $List = $Context.Web.Lists.GetByTitle($ListName)
        $Fields = $List.Fields
        $Context.Load($List)
        $Context.Load($Fields)
        $Context.ExecuteQuery()

        $NewField = $Fields | where { ($_.Internalname -eq $Name) -or ($_.Title -eq $DisplayName) }
 
        if($NewField -ne $NULL)
        {
             Write-host "Field '$FieldName' already exists in the List!" -f Yellow
        }     
        else
        {
            $FieldSchema = "<Field Type='Note' DisplayName='$DisplayName' Name='$FieldName' Description='$Description' NumLines='5' RichText='FALSE' Sortable='FALSE'></Field>"
            $NewField = $List.Fields.AddFieldAsXml($FieldSchema, $true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
            $List.Update()
            $Context.ExecuteQuery()
            Write-host "Success, List: $ListName, Field: $FieldName, Type: MultiText" -f Green 
        } 
    }
    Catch {
        write-host -f Red "Error Adding Column!" $_.Exception.Message
    }
}

Function Add-SPOListColumn-Boolean
{
    param
    (
        [Parameter(Mandatory=$true)] [string]$ListName = $(throw "Please Enter the List Name!"),
        [Parameter(Mandatory=$true)] [string]$FieldName = $(throw "Please Enter the FieldName!"),
        [Parameter(Mandatory=$true)] [string]$DisplayName = $(throw "Please Enter the DisplayName!"),
        [Parameter(Mandatory=$true)] [string]$Description = $(throw "Please Enter the Description!"),
        [Parameter(Mandatory=$true)] [boolean]$IsRequired = $(throw "Please Enter the IsRequired!"),
        [Parameter(Mandatory=$true)] [Int32]$DefaultValue = $(throw "Please Enter the DefaultValue!")
    )
    Try {
        #Get the List
        $List = $Context.Web.Lists.GetByTitle($ListName)
        $Fields = $List.Fields
        $Context.Load($List)
        $Context.Load($Fields)
        $Context.ExecuteQuery()

        $NewField = $Fields | where { ($_.Internalname -eq $Name) -or ($_.Title -eq $DisplayName) }
 
        if($NewField -ne $NULL)
        {
             Write-host "Field '$FieldName' already exists in the List!" -f Yellow
        }     
        else
        {
            $FieldSchema = "<Field Type='Boolean' DisplayName='$DisplayName' Name='$FieldName' Description='$Description'><Default>$DefaultValue</Default></Field>"
            $NewField = $List.Fields.AddFieldAsXml($FieldSchema, $true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
            $List.Update()
            $Context.ExecuteQuery()
            Write-host "Success, List: $ListName, Field: $FieldName, Type: Boolean" -f Green 
        } 
    }
    Catch {
        write-host -f Red "Error Adding Column!" $_.Exception.Message
    }
}

Function Add-SPOListColumn-DateTime
{
    param
    (
        [Parameter(Mandatory=$true)] [string]$ListName = $(throw "Please Enter the List Name!"),
        [Parameter(Mandatory=$true)] [string]$FieldName = $(throw "Please Enter the FieldName!"),
        [Parameter(Mandatory=$true)] [string]$DisplayName = $(throw "Please Enter the DisplayName!"),
        [Parameter(Mandatory=$true)] [string]$Description = $(throw "Please Enter the Description!"),
        [Parameter(Mandatory=$true)] [boolean]$IsRequired = $(throw "Please Enter the IsRequired!")
    )
    Try {
        #Get the List
        $List = $Context.Web.Lists.GetByTitle($ListName)
        $Fields = $List.Fields
        $Context.Load($List)
        $Context.Load($Fields)
        $Context.ExecuteQuery()

        $NewField = $Fields | where { ($_.Internalname -eq $Name) -or ($_.Title -eq $DisplayName) }
 
        if($NewField -ne $NULL)
        {
             Write-host "Field '$FieldName' already exists in the List!" -f Yellow
        }     
        else
        {
            $FieldSchema = "<Field Type='DateTime' DisplayName='$DisplayName' Name='$FieldName' Description='$Description'></Field>"
            $NewField = $List.Fields.AddFieldAsXml($FieldSchema, $true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
            $List.Update()
            $Context.ExecuteQuery()
            Write-host "Success, List: $ListName, Field: $FieldName, Type: DateTime" -f Green 
        } 
    }
    Catch {
        write-host -f Red "Error Adding Column!" $_.Exception.Message
    }
}

Function Add-SPOListColumn-Number
{
    param
    (
        [Parameter(Mandatory=$true)] [string]$ListName = $(throw "Please Enter the List Name!"),
        [Parameter(Mandatory=$true)] [string]$FieldName = $(throw "Please Enter the FieldName!"),
        [Parameter(Mandatory=$true)] [string]$DisplayName = $(throw "Please Enter the DisplayName!"),
        [Parameter(Mandatory=$true)] [string]$Description = $(throw "Please Enter the Description!"),
        [Parameter(Mandatory=$true)] [boolean]$IsRequired = $(throw "Please Enter the IsRequired!")
    )
    Try {
        #Get the List
        $List = $Context.Web.Lists.GetByTitle($ListName)
        $Fields = $List.Fields
        $Context.Load($List)
        $Context.Load($Fields)
        $Context.ExecuteQuery()

        $NewField = $Fields | where { ($_.Internalname -eq $Name) -or ($_.Title -eq $DisplayName) }
 
        if($NewField -ne $NULL)
        {
             Write-host "Field '$FieldName' already exists in the List!" -f Yellow
        }     
        else
        {
            $FieldSchema = "<Field Type='Number' DisplayName='$DisplayName' Name='$FieldName' Description='$Description'></Field>"
            $NewField = $List.Fields.AddFieldAsXml($FieldSchema, $true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
            $List.Update()
            $Context.ExecuteQuery()
            Write-host "Success, List: $ListName, Field: $FieldName, Type: Number" -f Green 
        } 
    }
    Catch {
        write-host -f Red "Error Adding Column!" $_.Exception.Message
    }
}

Function Add-SPOListColumn-Lookup
{
    param
    (
        [Parameter(Mandatory=$true)] [string]$ListName = $(throw "Please Enter the List Name!"),
        [Parameter(Mandatory=$true)] [string]$FieldName = $(throw "Please Enter the FieldName!"),
        [Parameter(Mandatory=$true)] [string]$DisplayName = $(throw "Please Enter the DisplayName!"),
        [Parameter(Mandatory=$true)] [string]$Description = $(throw "Please Enter the Description!"),
        [Parameter(Mandatory=$true)] [boolean]$IsRequired = $(throw "Please Enter the IsRequired!"),
        [Parameter(Mandatory=$true)] [string]$LookupList = $(throw "Please Enter the LookupList!"),
        [Parameter(Mandatory=$true)] [string]$ShowField = $(throw "Please Enter the ShowField!")
    )
    Try {
        #Get the List
        $List = $Context.Web.Lists.GetByTitle($ListName)
        $Fields = $List.Fields
        $Context.Load($List)
        $Context.Load($Fields)
        $Context.ExecuteQuery()

        $NewField = $Fields | where { ($_.Internalname -eq $Name) -or ($_.Title -eq $DisplayName) }
 
        if($NewField -ne $NULL)
        {
             Write-host "Field '$FieldName' already exists in the List!" -f Yellow
        }     
        else
        {
            $FieldSchema = "<Field Type='Lookup' DisplayName='$DisplayName' Name='$FieldName' Description='$Description' List='$LookupList' ShowField='$ShowField'></Field>"
            $NewField = $List.Fields.AddFieldAsXml($FieldSchema, $true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
            $List.Update()
            $Context.ExecuteQuery()
            Write-host "Success, List: $ListName, Field: $FieldName, Type: Lookup, List: $LookupList, ShowField: $ShowField" -f Green 
        } 
    }
    Catch {
        write-host -f Red "Error Adding Column!" $_.Exception.Message
    }
}

Function Add-SPOListColumn-Url
{
    param
    (
        [Parameter(Mandatory=$true)] [string]$ListName = $(throw "Please Enter the List Name!"),
        [Parameter(Mandatory=$true)] [string]$FieldName = $(throw "Please Enter the FieldName!"),
        [Parameter(Mandatory=$true)] [string]$DisplayName = $(throw "Please Enter the DisplayName!"),
        [Parameter(Mandatory=$true)] [string]$Description = $(throw "Please Enter the Description!"),
        [Parameter(Mandatory=$true)] [boolean]$IsRequired = $(throw "Please Enter the IsRequired!")
    )
    Try {
        #Get the List
        $List = $Context.Web.Lists.GetByTitle($ListName)
        $Fields = $List.Fields
        $Context.Load($List)
        $Context.Load($Fields)
        $Context.ExecuteQuery()

        $NewField = $Fields | where { ($_.Internalname -eq $Name) -or ($_.Title -eq $DisplayName) }
 
        if($NewField -ne $NULL)
        {
             Write-host "Field '$FieldName' already exists in the List!" -f Yellow
        }     
        else
        {
            $FieldSchema = "<Field Type='URL' DisplayName='$DisplayName' Name='$FieldName' Description='$Description' Format='Hyperlink'></Field>"
            $NewField = $List.Fields.AddFieldAsXml($FieldSchema, $true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
            $List.Update()
            $Context.ExecuteQuery()
            Write-host "Success, List: $ListName, Field: $FieldName, Type: URL" -f Green 
        } 
    }
    Catch {
        write-host -f Red "Error Adding Column!" $_.Exception.Message
    }
}

Function Add-SPOListColumn-User
{
    param
    (
        [Parameter(Mandatory=$true)] [string]$ListName = $(throw "Please Enter the List Name!"),
        [Parameter(Mandatory=$true)] [string]$FieldName = $(throw "Please Enter the FieldName!"),
        [Parameter(Mandatory=$true)] [string]$DisplayName = $(throw "Please Enter the DisplayName!"),
        [Parameter(Mandatory=$true)] [string]$Description = $(throw "Please Enter the Description!"),
        [Parameter(Mandatory=$true)] [boolean]$IsRequired = $(throw "Please Enter the IsRequired!")
    )
    Try {
        #Get the List
        $List = $Context.Web.Lists.GetByTitle($ListName)
        $Fields = $List.Fields
        $Context.Load($List)
        $Context.Load($Fields)
        $Context.ExecuteQuery()

        $NewField = $Fields | where { ($_.Internalname -eq $Name) -or ($_.Title -eq $DisplayName) }
 
        if($NewField -ne $NULL)
        {
             Write-host "Field '$FieldName' already exists in the List!" -f Yellow
        }     
        else
        {
            $FieldSchema = "<Field Type='User' DisplayName='$DisplayName' Name='$FieldName' Description='$Description' Format='Hyperlink'></Field>"
            $NewField = $List.Fields.AddFieldAsXml($FieldSchema, $true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
            $List.Update()
            $Context.ExecuteQuery()
            Write-host "Success, List: $ListName, Field: $FieldName, Type: User" -f Green 
        } 
    }
    Catch {
        write-host -f Red "Error Adding Column!" $_.Exception.Message
    }
}

Function Add-SPOListColumn-Choice
{
    param
    (
        [Parameter(Mandatory=$true)] [string]$ListName = $(throw "Please Enter the List Name!"),
        [Parameter(Mandatory=$true)] [string]$FieldName = $(throw "Please Enter the FieldName!"),
        [Parameter(Mandatory=$true)] [string]$DisplayName = $(throw "Please Enter the DisplayName!"),
        [Parameter(Mandatory=$true)] [string]$Description = $(throw "Please Enter the Description!"),
        [Parameter(Mandatory=$true)] [boolean]$IsRequired = $(throw "Please Enter the IsRequired!"),
        [Parameter(Mandatory=$true)] [string[]]$Choices = $(throw "Please Enter the Choices!"),
        [Parameter(Mandatory=$true)] [string]$DefaultChoice = $(throw "Please Enter the Default Choice!")
    )
    Try {
        #Get the List
        $List = $Context.Web.Lists.GetByTitle($ListName)
        $Fields = $List.Fields
        $Context.Load($List)
        $Context.Load($Fields)
        $Context.ExecuteQuery()

        $NewField = $Fields | where { ($_.Internalname -eq $Name) -or ($_.Title -eq $DisplayName) }
 
        if($NewField -ne $NULL)
        {
             Write-host "Field '$FieldName' already exists in the List!" -f Yellow
        }     
        else
        {
            $start1 = "<CHOICES>"
            $end1 = "</CHOICES>"
            $start2 = "<CHOICE>"
            $end2 = "</CHOICE>"
            $default1 = "<Default>"
            $default2 = "</Default>"

            $choiseString = $default1 + $DefaultChoice + $default2 + $start1

            Foreach($choice in $Choices)
            {
                $choiseString = $choiseString + $start2 + $choice+ $end2
            }
            $choiseString = $choiseString + $end1

            #Write-Host $choiseString

            $FieldSchema = "<Field Type='Choice' DisplayName='$DisplayName' Name='$FieldName' Description='$Description' Format='Dropdown'>'$choiseString'</Field>"
            $NewField = $List.Fields.AddFieldAsXml($FieldSchema, $true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
            $List.Update()
            $Context.ExecuteQuery()
            Write-host "Success, List: $ListName, Field: $FieldName, Type: Choice" -f Green 
        } 
    }
    Catch {
        write-host -f Red "Error Adding Column!" $_.Exception.Message
    }
}

Function New-SPOGroup
{
    param
    (
        [Parameter(Mandatory=$true)] [string]$GroupName = $(throw "Please Enter the Group Name!"),
        [Parameter(Mandatory=$true)] [string]$GroupDescription = $(throw "Please Enter the Group Description!"),
        [Parameter(Mandatory=$true)] [string]$PermissionLevel = $(throw "Please Enter the Permission Level!")
    )

    Try {
        #Get all existing groups of the site
        $Groups = $Context.Web.SiteGroups
        $Context.load($Groups)
        $Context.ExecuteQuery()

        #Get Group Names
        $GroupNames =  $Groups | Select -ExpandProperty Title

        $NewGroup = $Groups | where { ($_.Title -eq $GroupName) }

        if($NewGroup -ne $NULL)
        {
             Write-host "Group '$GroupName' already exists in the site!" -f Yellow
        }     
        else
        {
            #sharepoint online powershell create group
            $GroupInfo = New-Object Microsoft.SharePoint.Client.GroupCreationInformation
            $GroupInfo.Title = $GroupName
            $GroupInfo.Description = $GroupDescription
            $Group = $Context.web.SiteGroups.Add($GroupInfo)
            $Context.ExecuteQuery()

            #Get Group Title and ID
            $Groups | Select Title, ID

            #Assign permission to the group
            $RoleDef = $Context.web.RoleDefinitions.GetByName($PermissionLevel)
            $RoleDefBind = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($Context)
            $RoleDefBind.Add($RoleDef)
            $Context.Load($Context.Web.RoleAssignments.Add($Group,$RoleDefBind))
            $Context.ExecuteQuery()

            write-Host -f Green "Success, Group: $GroupName"
        }
    }
    Catch {
        write-host -f Red "Error Creating Group!" $_.Exception.Message
    }
}

Function Rename-SPOListColumn
{
    param
    (
        [Parameter(Mandatory=$true)] [string]$ListName = $(throw "Please Enter the Column Name!"),
        [Parameter(Mandatory=$true)] [string]$FieldName = $(throw "Please Enter the New Name!"),
        [Parameter(Mandatory=$true)] [string]$Title = $(throw "Please Enter the List Name!")
    )

    Try
    {
        $List = $Context.Web.Lists.GetByTitle($ListName) 
        $Field = $List.Fields.GetByInternalNameOrTitle($FieldName)
        $Field.Title = $Title
        $Field.Update()
        $Context.ExecuteQuery()
        Write-Host "Field Name $FieldName Changed To: $Title" -f Green
    }
    Catch
    {
      write-host -f Red "Something Went Wrong!" $_.Exception.Message
    }
}