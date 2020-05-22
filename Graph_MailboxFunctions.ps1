### Setting up Enum for $FolderSet parameter of New-GraphMailArchiveFolders
$ENUM_NewGraphMailArchiveFolders_FolderSet = @"
    namespace GraphMailFunctions {
        public enum FolderSet {
            Default=1,
            Fire=2,
            Planning=3,
            PublicWorks=4,
            Parks=5,
            CityAttorney=6,
            HumanResources=7
        }
    }
"@

Add-Type -TypeDefinition $ENUM_NewGraphMailArchiveFolders_FolderSet -Language CSharpVersion3


### FUNCTIONS
Function Convert-GraphValueToGuid {

    [CmdletBinding()]
    PARAM(

        [Parameter(Mandatory=$true)]
        [String]$Base64Value

    )

    $base64ToString = [System.Convert]::FromBase64String($Base64Value)
    $guidValue = [System.GUID]$base64ToString

    return $guidValue

}

Function Convert-GuidToGraphValue {

    [CmdletBinding()]
    PARAM(

        [Parameter(Mandatory=$true)]
        [String]$GuidString

    )

    $guidValue = [System.Guid]$GuidString
    $guidBytes = $guidValue.ToByteArray()
    $guidToBase64String = [System.Convert]::ToBase64String($guidBytes)

    return $guidToBase64String

}

function Get-GraphBearerToken {

    ### For first time use, you need to setup a credential object and store it as an XML file to properly 
    ### secure your Azure application secret. Use the Azure "Application Id" as the username and the 
    ### API secret as the password

    # Get-Credential -UserName <application_id> -Message "Enter Secret" | Export-Clixml <path_to_directory\MyApp.xml>

    [CmdletBinding()]
    PARAM(

        [Parameter(Mandatory=$true,ParameterSetName='File')]
        [ValidateNotNullOrEmpty()]
        [String]$Path,

        [Parameter(Mandatory=$true,ParameterSetName='Credential')]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.PSCredential]$ServiceCredential,

        [Parameter(ParameterSetName='File')]
        [Parameter(ParameterSetName='Credential')]
        [String]$TokenResource = "https://graph.microsoft.com/",

        ### TokenEndpoint should be the oauth token endpoint for your tenant...
        ### e.g. https://login.microsoftonline.com/<tenant_GUID>/oauth2/token
        [Parameter(ParameterSetName='File')]
        [Parameter(ParameterSetName='Credential')]
        [ValidateNotNullOrEmpty()]
        [String]$TokenEndpoint

    )

    if($PSCmdlet.ParameterSetName -eq "File") {

        $ServiceCredential = Import-Clixml -Path $Path

    }
 
    $Body = @{

        grant_type = 'client_credentials'
        client_id = $ServiceCredential.UserName
        client_secret = $ServiceCredential.GetNetworkCredential().Password
        resource = $TokenResource

    }

    $Headers = @{

        'Content-Type' = 'application/x-www-form-urlencoded'

    }

    $restParams = @{
        
        Uri = $TokenEndpoint
        Method = 'POST'
        Headers = $Headers
        Body = $Body

    }

    $restResponse = Invoke-RestMethod @restParams

    return $restResponse

}

Function Get-GraphMailboxFolders {

    [CmdletBinding()]
    PARAM(

        [Parameter(ParameterSetName='NextLink',Mandatory=$true)]
        [Parameter(ParameterSetName='BuildURI',Mandatory=$true)]
        $BearerToken,

        [Parameter(ParameterSetName='NextLink',Mandatory=$true)]
        [Alias("URI","Resource")]
        [String]$NextLink,

        [Parameter(ParameterSetName='BuildURI',Mandatory=$true)]
        [Alias("UPN","Username")]
        [String]$UserPrincipalName,

        [Parameter(ParameterSetName='BuildURI')]
        [UInt32]$SkipCount,

        [Parameter(ParameterSetName='BuildURI')]
        [String]$BaseFolderId,

        [Parameter(ParameterSetName='BuildURI')]
        [Switch]$GetChildFolders

    )

    switch($PSCmdlet.ParameterSetName) {

        'NextLink' {

            $restUri = $NextLink

        }

        'BuildURI' {

            $graphMailFoldersUri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/mailFolders"
            $restUri = $graphMailFoldersUri

            if($BaseFolderId) {

                $restUri = "$restUri/$baseFolderId"

            }

            if($GetChildFolders) {

                $restUri = "$restUri/childFolders"

            }

            if($SkipCount -gt 0) {

                $restUri = "$restUri`?`$skip=$SkipCount"
                
            }

        }
    
    }

    $restHeaders = @{

        Authorization = "Bearer $($BearerToken.access_token)"

    }

    $restParams = @{

        Uri = $restUri
        Method = 'GET'
        Headers = $restHeaders

    }

    Write-Verbose "Final URI string is: $restUri"
    $restResponse = Invoke-RestMethod @restParams

    return $restResponse

}

Function Get-GraphMailboxFolderHierarchy {

    [CmdletBinding()]
    PARAM(

        [Parameter(Mandatory=$true)]
        [Alias("UPN","User","Username")]
        [String]$UserPrincipalName,

        [Parameter(Mandatory=$true)]
        [Alias("Token")]
        $BearerToken,

        [Parameter()]
        [String]$BaseFolderId

    )

        #Setup script level variables
        $folderList = New-Object System.Collections.ArrayList
        $moreFolders = $true
        
        #If the function was called recursively with a folder id, make sure we target the specific folder
        if($PSBoundParameters.ContainsKey('BaseFolderId')) {

            Write-Verbose "BaseFolderId: $BaseFolderId"
            $mailFolders = Get-GraphMailboxFolders -BearerToken $BearerToken -UserPrincipalName $UserPrincipalName -BaseFolderId $BaseFolderId -GetChildFolders

        }

        else {

            $mailFolders = Get-GraphMailboxFolders -BearerToken $BearerToken -UserPrincipalName $UserPrincipalName

        }

        #Main process loop.  Recursively calls this function to enumerate childFolders
        do {

            foreach ($folder in $mailFolders.value) {

                $folder | Add-Member -MemberType NoteProperty -Name MailboxOwner -Value $UserPrincipalName
                
                #Query the Graph API for the retention tag Single Value Extended Property (SVEP)
                $retentionTagSVEPValues = Get-GraphMailboxFolderRetentionTag -UserPrincipalName $UserPrincipalName -BearerToken $BearerToken -FolderId $folder.id
                
                if($retentionTagSVEPValues.singleValueExtendedProperties) {

                    $retentionTagBase64Value = $retentionTagSVEPValues.singleValueExtendedProperties[0].value
                    $retentionTagGUID = Convert-GraphValueToGuid -Base64Value $retentionTagBase64Value
                    $folder | Add-Member -MemberType NoteProperty -Name RetentionTagGUID -Value $retentionTagGUID

                }                    

                $folderList.Add($folder) | Out-Null

                if($folder.childFolderCount -gt 0) {

                    $childFolders = Get-GraphMailboxFolderHierarchy -UserPrincipalName $UserPrincipalName -BearerToken $BearerToken -BaseFolderId $folder.id
                    $folder | Add-Member -MemberType NoteProperty -Name childFolders -Value $childFolders

                    

                }

            }

            if($mailFolders.'@odata.nextLink') {

                $mailFolders = Get-GraphMailboxFolders -BearerToken $token -NextLink $mailFolders.'@odata.nextLink'

            }
            
            else {
                
                $moreFolders = $false

            }

        } while ($moreFolders -eq $true)

        return $folderList

}

Function Get-GraphMailboxFolderRetentionTag {

    [CmdletBinding()]
    PARAM(

        [Parameter(Mandatory=$true)]
        $BearerToken,

        [Parameter(Mandatory=$true)]
        [Alias("UPN","Username")]
        [String]$UserPrincipalName,

        [Parameter(Mandatory=$true)]
        [String]$FolderId

    )

    $baseUri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/mailFolders"
    $singleValueExtendedPropertyParam = "`?`$expand=singleValueExtendedProperties(`$filter=id eq 'Binary 0x3019')"
    $restUri = "$baseUri/$FolderId$singleValueExtendedPropertyParam"

    $restHeaders = @{

        Authorization = "Bearer $($BearerToken.access_token)"

    }

    $restParams = @{

        Uri = $restUri
        Method = 'GET'
        Headers = $restHeaders

    }

    $restResponse = Invoke-RestMethod @restParams

    return $restResponse

}

Function New-GraphMailArchiveFolders {

    [CmdletBinding()]
    PARAM(

        [Parameter(Mandatory=$true)]
        $BearerToken,

        [Parameter(Mandatory=$true)]
        [Alias("UPN","Username")]
        [String]$UserPrincipalName,

        [Parameter(Mandatory=$true)]
        [GraphMailFunctions.FolderSet]$FolderSet,

        <#
            FolderSetDataPath needs to be the path to a CSV file that contains the following
            three data columns

                FolderName - This is the name of the folder that the retention tag applies to
                RetentionId - This is the 'RetentionId' of the retention tag (from Get-RetentionPolicyTag)
                FolderSet - The 'FolderSet' that the tag belongs to (a tag can be added to multiple FolderSets)
        #>
        [Parameter()]
        [String]$FolderSetDataPath

    )

    $folderSets = Import-Csv $FolderSetDataPath

    $archiveFolder = New-GraphMailFolder -BearerToken $BearerToken -UserPrincipalName $UserPrincipalName -FolderName "Archive Folders"
    
    foreach($folder in $folderSets) {

        if( ($folder.FolderSet -eq $FolderSet) -or ($folder.FolderSet -eq "Default") ) {
            $newFolder = New-GraphMailFolder -BearerToken $BearerToken -UserPrincipalName $UserPrincipalName -FolderName $folder.FolderName -BaseFolderId $archiveFolder.id
            Set-GraphMailFolderRetentionTag -BearerToken $BearerToken -UserPrincipalName $UserPrincipalName -FolderId $newFolder.id -RetentionTagGuid $folder.RetentionId | Out-Null
        }

    }

}

Function New-GraphMailFolder {

    [CmdletBinding()]
    PARAM(

        [Parameter(Mandatory=$true)]
        $BearerToken,

        [Parameter(Mandatory=$true)]
        [Alias("UPN","Username")]
        [String]$UserPrincipalName,

        [Parameter(Mandatory=$true)]
        [String]$FolderName,

        [Parameter()]
        [String]$BaseFolderId

    )

    $restUri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/mailFolders"

    if($BaseFolderId) {

        $restUri = "$restUri/$BaseFolderId/childFolders"

    }

    $restBody = @{

        displayName = $FolderName

    }
    
    $restHeaders = @{

        Authorization = "Bearer $($BearerToken.access_token)"
        'Content-Type' = 'application/json'

    }

    $restParams = @{

        Uri = $restUri
        Method = 'POST'
        Headers = $restHeaders
        Body = ($restBody | ConvertTo-Json)

    }

    $restResponse = Invoke-RestMethod @restParams

    return $restResponse

}

Function Set-GraphMailFolderRetentionTag {

    [CmdletBinding()]
    PARAM(

        [Parameter(Mandatory=$true)]
        $BearerToken,

        [Parameter(Mandatory=$true)]
        [Alias("UPN","Username")]
        [String]$UserPrincipalName,

        [Parameter(Mandatory=$true)]
        [String]$FolderId,

        [Parameter(Mandatory=$true)]
        [String]$RetentionTagGuid

    )

    $restUri = "https://graph.microsoft.com/beta/users/$UserPrincipalName/mailFolders(`'$FolderId`')"
    #$restUri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/mailFolders/$FolderId"

    $base64RetentionValue = Convert-GuidToGraphValue -GuidString $RetentionTagGuid

    $svepRetentionTag = [PSCustomObject]@{
        id = 'Binary 0x3019'
        value = $base64RetentionValue
    }

    $svepRetentionFlag = [PSCustomObject]@{
        id = 'Integer 0x301D'
        value = '89'        
    }

    $singleValueExtendedPropertyArray = New-Object System.Collections.ArrayList
    $singleValueExtendedPropertyArray.Add($svepRetentionTag) | Out-Null
    $singleValueExtendedPropertyArray.Add($svepRetentionFlag) | Out-Null

    $restBody = @{

        singleValueExtendedProperties = $singleValueExtendedPropertyArray

    }
    
    
    $restHeaders = @{

        Authorization = "Bearer $($BearerToken.access_token)"
        'Content-Type' = 'application/json'

    }

    $restParams = @{

        Uri = $restUri
        Method = 'PATCH'
        Headers = $restHeaders
        Body = ($restBody | ConvertTo-Json)

    }

    $restResponse = Invoke-RestMethod @restParams

    return $restResponse

}