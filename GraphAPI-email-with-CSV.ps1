# Jon Buhagiar
# 08/16/24
# Emails directory listing to specific email address via the Graph API

$clientId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
$tenantId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
$clientSecretValue = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
$fromEmail = "someone@fromemail.com"
$toEmail = "someone@toemail.com"
$DirPath = "C:\"

$tokenEndpointUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

$body = @{
    client_id     = $clientId
    client_secret = $clientSecretValue
    grant_type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
}

$accessToken = $(Invoke-RestMethod -Uri $tokenEndpointUri -Method Post -Body $body).access_token

$dirListing = dir $DirPath | select name,length,lastwritetime | ConvertTo-Csv -NoTypeInformation
$test = $dirListing -join "`n"
$fileBytes = [System.Text.Encoding]::UTF8.GetBytes($test)
$encodedString = [System.Convert]::ToBase64String($fileBytes)

$emailParams = @{
    "message" = @{
        "subject" = "Files in the directory of $DirPath on $(Get-Date -Format 'MM-dd-yy')"
        "body" = @{
            "contentType" = "Text"
            "content" = "The files in the CSV are from the directory of $DirPath on $(Get-Date -Format 'MM-dd-yy')."
        }
        "toRecipients" = @(
            @{
                "emailAddress" = @{
                    "address" = $toEmail
                }
            }
        )
		"attachments" = @(
            @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                "name" = "Files-$DirPath-$(Get-Date -Format 'MM-dd-yy').csv"
                "contentType" = "text/csv"
                "contentBytes" = $encodedString
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