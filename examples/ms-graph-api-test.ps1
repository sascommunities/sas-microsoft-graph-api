# Use this script to authenticate with Microsoft Graph API and retrieve user information.
# It's useful to test that your Microsoft Graph API setup is working correctly.
# The script will guide you through the authentication process and save the access and refresh tokens to a file.
# It requires a configuration file named config.json with the following structure: 
# {
#   "tenant_id": "your-tenant-id",
#   "client_id": "your-client-id",
#   "redirect_uri": "your-redirect-uri",    
#   "resource": "https://graph.microsoft.com"
# } 

# Usage:
# 1. Create a config.json file with the details of your app and tenant.
# 2. Run this script in PowerShell, specifying the path to your config.json file if it's not in the same directory.
# 3. Follow the prompts to authenticate and retrieve your user information.

param(
    [string]$ConfigPath = "config.json"
)

if (-not (Test-Path -Path $ConfigPath)) {
    Write-Error "Configuration file '$ConfigPath' not found. Verify the path to your config.json file and specify with -ConfigPath."
    exit 1
}

$config = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
$tenantId = $config.tenant_id
$clientId = $config.client_id
$redirectUri = $config.redirect_uri
$resource = $config.resource

# Build the authorization URL (SAS macro equivalent)
$msloginBase = "https://login.microsoftonline.com"
$authUrl = "$msloginBase/$tenantId/oauth2/authorize" +
    "?client_id=$clientId" +
    "&response_type=code" +
    "&redirect_uri=$([uri]::EscapeDataString($redirectUri))" +
    "&resource=$([uri]::EscapeDataString($resource))"

Write-Host "Open the following URL in your browser to authenticate:"
Write-Host $authUrl

Start-Process $authUrl

Write-Host "After authenticating in the browser, paste the full redirected URL from the address bar below."
$redirectedUrl = Read-Host "Redirected URL"

try {
    $uri = [System.Uri]$redirectedUrl
    $queryParams = [System.Web.HttpUtility]::ParseQueryString($uri.Query)
    $authCode = $queryParams["code"]
    if (-not $authCode) {
        Write-Error "Authorization code not found in the URL. Please ensure you pasted the correct redirected URL."
        exit 1
    }
    Write-Host "Authorization code captured."
} catch {
    Write-Error "Invalid URL format. Please try again."
    exit 1
}


# Token endpoint
$tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/token"

# Request access and refresh tokens
$tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body @{
    client_id     = $clientId
    code          = $authCode
    redirect_uri  = $redirectUri
    grant_type    = "authorization_code"
    resource     = $resource
} -ContentType "application/x-www-form-urlencoded"

if (-not $tokenResponse.access_token) {
    Write-Error "Failed to obtain access token. Response: $($tokenResponse | ConvertTo-Json)"
    exit 1
}   

# save the token.json to an external file in the same directory as config.json
$tokenFilePath = Join-Path -Path (Split-Path -Parent $ConfigPath) -ChildPath "token.json"
$tokenResponse | ConvertTo-Json -Depth 5 | Set-Content -Path $tokenFilePath -Force  
Write-Host "Access and refresh tokens saved to $tokenFilePath"

# Use the access token to call the /me endpoint

$accessToken = $tokenResponse.access_token
$refreshToken = $tokenResponse.refresh_token

Write-Host "Access Token: $accessToken"
Write-Host "Refresh Token: $refreshToken"

# Call the /me endpoint
$meResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/me" -Headers @{
    Authorization = "Bearer $accessToken"
}
if (-not $meResponse) {
    Write-Error "Failed to retrieve user information from /me endpoint."
    exit 1
}

# Print the JSON response
$meResponse | ConvertTo-Json -Depth 5