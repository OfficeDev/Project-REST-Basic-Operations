<#
  Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
  See LICENSE in the project root for license information.
#>


# Get list of projects using OData ReST API
param
(
    # SharepointOnline project site collection URL
    $SiteUrl = $(throw "SiteUrl parameter is required")
)
# Load ReST helper methods
. .\ReST.ps1

# Set up the request authentication
Set-SPOAuthenticationTicket $siteUrl
Set-DigestValue $siteUrl

# Get list of projects using OData ReST API
Get-ReSTRequest $SiteUrl "ProjectData/Projects"