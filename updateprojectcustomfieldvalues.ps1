<#
  Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
  See LICENSE in the project root for license information.
#>

# Updates a Project Custom field value using ReST API
param
(
    # SharepointOnline project site collection URL
    $SiteUrl = $(throw "SiteUrl parameter is required"),
    $projectId = $(throw "projectId parameter is required"),
    $customFieldId = $(throw "customFieldId parameter is required"),
    $lookupEntryId,     #use this parameter for custom fields associated with a lookup table
    $customFieldValue,   #use this parameter for non-lookup custom fields,
    [switch] $UseWebLogin
)
# Load ReST helper methods
. .\ReST.ps1

# Set up the request authentication
Set-SPOAuthenticationTicket $siteUrl $UseWebLogin
Set-DigestValue $siteUrl

# ReST request to check out the project
Post-ReSTRequest $SiteUrl "ProjectServer/Projects('$projectid')/checkOut" $null

# Set up the request parameters
# ReST request to update the project custom field
$customFieldInternalName = ([xml] (Get-ReSTRequest $siteUrl "ProjectServer/CustomFields('$customFieldId')")).entry.content.properties.InternalName
if (-not [String]::IsNullOrEmpty($lookupEntryId))
{
    $customFieldValue = ([xml] (Get-ReSTRequest $siteUrl "ProjectServer/CustomFields('$customFieldId')/LookupEntries('$lookupEntryId')")).entry.content.properties.InternalName
}
# ReST request to update the project custom field (see http://msdn.microsoft.com/en-us/library/hh642428(v=office.12).aspx for parameter information)
$body = "{'customFieldDictionary':[{'Key':'$customFieldInternalName','Value':'$customFieldValue', 'ValueType':'Edm.String'}]}"
Post-ReSTRequest $SiteUrl "ProjectServer/Projects('$projectid')/Draft/UpdateCustomFields" $body

# ReST request to publish and check-in the project
Post-ReSTRequest $SiteUrl "ProjectServer/Projects('$projectid')/Draft/publish(true)" $null

