<#
  Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
  See LICENSE in the project root for license information.
#>

$ErrorActionPreference = "Stop"  # http://technet.microsoft.com/en-us/library/dd347731.aspx
Set-StrictMode -Version "Latest" # http://technet.microsoft.com/en-us/library/dd347614.aspx

# PS helper methods to call ReST API methods targeting Project Online tenants
$global:accessHeader = ''
$global:digestValue = ''

[Reflection.Assembly]::LoadFrom("$($PSScriptRoot)\Microsoft.Identity.Client.dll") | Out-Null

function Get-AuthToken([Uri] $Uri)
{
    # NOTE: Create an azure app and update $clientId,$tenantId and $redirectUri below
    $clientId = ""
    $tenantId = ""
    $redirectUri = ""
    # user's PJO login account
    $user = ""
    $scopes = New-Object System.Collections.Generic.List[string]
    # Project.Write Permission scope for app eg:"https://brismith.sharepoint.com/Project.Write"
    $writeScope = ""
    $scopes.Add($writeScope)
    $pcaConfig = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($clientId).WithTenantId($tenantId).WithRedirectUri($redirectUri);

    $authenticationResult = $pcaConfig.Build().AcquireTokenInteractive($scopes)
                .WithPrompt([Microsoft.Identity.Client.Prompt]::NoPrompt)
                .WithLoginHint($user).ExecuteAsync().Result;

    return $authenticationResult
}

# Gets Auth Token using MSAL library.
function Set-SPOAuthenticationTicket([string] $siteUrl)
{
    $siteUri = New-Object Uri -ArgumentList $siteUrl

    $authResult = Get-AuthToken -Uri $siteUri
    if ($authResult -ne $null)
    {
        $global:accessHeader = $authResult.AccessTokenType + " " + $authResult.AccessToken
    }
    
    if ([String]::IsNullOrEmpty($global:accessHeader))
    {
        throw 'Could not obtain authentication ticket based on provided credentials for specified site'
    }
}

function Build-ReSTRequest([string] $siteUrl, [string]$endpoint, [string]$method, [string]$body = $null)
{
    $url = ([string]$siteUrl).TrimEnd("/") + "/_api/" + $endpoint
    $req = [System.Net.WebRequest]::Create($url)
    $req.Timeout = 120000
    $req.Method = $method
    
    [bool]$isReadOnly = (('GET','HEAD') -contains $req.Method)
    [bool]$isDigestRequest = $endpoint -contains 'contextinfo'
    
    if ([String]::IsNullOrEmpty($body))
    {
        $req.ContentLength = 0;
    }
    else
    {
        $req.ContentLength = $body.Length
        $req.ContentType = "application/json"
    }

    # set Authorization header
    $req.Headers.Add("Authorization", $global:accessHeader)
    
    if (-not $isDigestRequest)
    {
        if (-not $isReadOnly)
        {
            $req.Headers.Add("X-RequestDigest", $global:digestValue)
        }
    }
    
    if (-not [String]::IsNullOrEmpty($body))
    {
        $writer = New-Object System.IO.StreamWriter $req.GetRequestStream()
        $writer.Write($body)
        $writer.Close()
        $writer.Dispose()
    }
    
    return $req
}

function Set-DigestValue([string]$siteUrl)
{
    $request = Build-ReSTRequest $siteUrl 'contextinfo' 'POST' $null
    if ($request -eq $null)
    {
        throw 'Could not obtain a request digest value based on provided credentials for specified site'
    }
    
    try
    {
        $resp = $request.GetResponse()
        $reader = [System.Xml.XmlReader]::Create($resp.GetResponseStream())
        if ($reader.ReadToDescendant("d:FormDigestValue"))
        {
            $global:digestValue = $reader.ReadElementContentAsString()
        }
        else
        {
            throw 'Could not obtain a request digest value based on provided credentials for specified site'
        }
    }
    finally
    {
        if ($reader -ne $null)
        {
            $reader.Close()
            $reader.Dispose()
        }
        if ($resp -ne $null)
        {
            $resp.Close()
            $resp.Dispose()
        }
    }
}

function Post-ReSTRequest([string]$siteUrl, [string]$endpoint, [string]$body = $null)
{
    $request = Build-ReSTRequest $siteUrl $endpoint 'POST' $body
    $resp = $request.GetResponse()
    if ($resp -ne $null)
    {    
        $reader = New-Object System.IO.StreamReader $resp.GetResponseStream()
        $reader.ReadToEnd()
        $reader.Dispose()
    }
}

function Patch-ReSTRequest([string]$siteUrl, [string]$endpoint, [string]$body)
{
    $request = Build-ReSTRequest $siteUrl $endpoint 'PATCH' $body
    $resp = $request.GetResponse()
    if ($resp -ne $null)
    {    
        $reader = New-Object System.IO.StreamReader $resp.GetResponseStream()
        $reader.ReadToEnd()
        $reader.Dispose()
    }
}

function Get-ReSTRequest([string]$siteUrl, [string]$endpoint)
{
    $request = Build-ReSTRequest $siteUrl $endpoint 'GET'
    $resp = $request.GetResponse()
    if ($resp -ne $null)
    {
        $reader = New-Object System.IO.StreamReader $resp.GetResponseStream()
        $reader.ReadToEnd()
        $reader.Dispose()
    }
}
