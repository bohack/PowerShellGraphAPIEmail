# Jon Buhagiar
# 08/16/24
# Emails text to a specific email address via the Graph API

# Optional to force TLS1.2
#[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$clientId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
$tenantId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
$clientSecretValue = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
$fromEmail = "someone@fromemail.com"
$toEmail = "someone@toemail.com"

$tokenEndpointUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

$body = @{
    client_id     = $clientId
    client_secret = $clientSecretValue
    grant_type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
}

$accessToken = $(Invoke-RestMethod -Uri $tokenEndpointUri -Method Post -Body $body).access_token

$emailParams = @{
    "message" = @{
        "subject" = "Test Email $(Get-Date -Format 'MM-dd-yy')"
        "body" = @{
            "contentType" = "Text"
            "content" = "This is a test email $(Get-Date -Format 'MM-dd-yy')."
        }
        "toRecipients" = @(
            @{
                "emailAddress" = @{
                    "address" = $toEmail
                }
            }
        )
    }
}

$emailParamsJson = $emailParams | ConvertTo-Json -Depth 10

$sendMailUri = "https://graph.microsoft.com/v1.0/users/$fromEmail/sendMail"
$response = Invoke-RestMethod -Uri $sendMailUri -Method Post -Headers @{
    "Authorization" = "Bearer $accessToken"
    "Content-Type" = "application/json"
} -Body $emailParamsJson