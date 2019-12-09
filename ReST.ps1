<#
  Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
  See LICENSE in the project root for license information.
#>

# PS helper methods to call ReST API methods targeting Project Online tenants
$global:fedAuthTicket = ''
$global:digestValue = ''

Add-Type -Path "$($PSScriptRoot)\WebLogin.cs" -ReferencedAssemblies System.Windows.Forms

function Get-AuthCookieWebLogin([Uri] $Uri)
{
    $cookieStr = [WebLogin.WebLogin]::GetWebLoginCookie($Uri)
    if ([System.String]::IsNullOrEmpty($cookieStr))
    {
        throw "Authentication failed"
    }

    return $cookieStr
}

function Set-SPOAuthenticationTicket([string] $siteUrl, $useWebLogin = $false)
{
    if ($useWebLogin)
    {
        $global:fedAuthTicket = Get-AuthCookieWebLogin($siteUrl)
    }
    else
    {
        $username = "ca1@mod607022.onmicrosoft.com"
        Write-Host 'Enter password for user' $username 'on site' $siteUrl
        $password = Read-Host -AsSecureString
        
        # load the SP client runtime code
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
        $onlineCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password)
        if ($onlineCredentials -ne $null)
        {
            $global:fedAuthTicket = $onlineCredentials.GetAuthenticationCookie($siteUrl, $true);
        }
        
        if ([String]::IsNullOrEmpty($global:fedAuthTicket))
        {
            throw 'Could not obtain authentication ticket based on provided credentials for specified site'
        }
    }
}

function Build-ReSTRequest([string] $siteUrl, [string]$endpoint, [string]$method, [string]$body = $null)
{
    $url = ([string]$siteUrl).TrimEnd("/") + "/_api/" + $endpoint
    $req = [System.Net.WebRequest]::Create($url)
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

    $cookies = New-Object System.Net.CookieContainer
    $cookies.SetCookies($siteUrl, $global:fedAuthTicket)
    
    $req.CookieContainer = $cookies
    
    if (-not $isDigestRequest)
    {
        if (-not $isReadOnly)
        {
            $req.Headers.Add("X-RequestDigest", $global:digestValue)
            
            if ($method -eq "PATCH")
            {
                $req.Headers.Add("If-Match", "*");
            }
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
