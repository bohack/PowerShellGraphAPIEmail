# Jon Buhagiar
# 08/16/24
# Get the current BYOD list of users from Intune for the past 31 days and email the CSV.

Import-Module Microsoft.Graph.Intune

$clientId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
$tenantId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
$clientSecretValue = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
$fromEmail = "someone@fromemail.com"
$toEmail = "someone@toemail.com"

$tokenEndpointUri = "https://login.microsoftonline.com/$tenantId/oauth2/token"

$body = @{
    client_id     = $clientId
    client_secret = $clientSecret
    resource      = "https://graph.microsoft.com"
    grant_type    = "client_credentials"
}

$accessToken = $(Invoke-RestMethod -Uri $tokenEndpointUri -Method Post -Body $body).access_token

$headers = @{
    Authorization = "Bearer $accessToken"
}

$dataArray = @()
$intuneDevices = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices" -Headers $headers

$intuneDevices.value | ForEach-Object {
    If (($_.managedDeviceOwnerType -eq "personal") -and ($_.complianceState -eq "compliant")) {
        If ($(($(get-date) - $(Get-Date $_.lastSyncDateTime)).Days) -lt 31) {
            $intuneUser = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($_.EmailAddress)" -Headers $headers
            $dataArray += [PSCustomObject]@{
            DisplayName = $_.userDisplayName
            Title = $intuneUser.jobTitle
            EmailAddress = $_.EmailAddress
            LastSyncDateTime = $_.lastSyncDateTime
            PhoneType = "$($_.manufacturer) $($_.model)"
            }
        }
    }
}

$convertArray = $($dataArray | Sort-Object -Property DisplayName | ConvertTo-Csv -NoTypeInformation) -join "`n"
$fileBytes = [System.Text.Encoding]::UTF8.GetBytes($convertArray)
$encodedString = [System.Convert]::ToBase64String($fileBytes)

$emailParams = @{
    "message" = @{
        "subject" = "List of BYOD Devices $(Get-Date -Format 'MM-dd-yy')"
        "body" = @{
            "contentType" = "Text"
            "content" = "The enclosed file contains the current BYOD list. There are $($dataArray.count) users."
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
                "name" = "BYOD List $(Get-Date -Format 'MMddyy').csv"
                "contentType" = "text/csv"
                "contentBytes" = $encodedString
            }
		)
    }
}

$emailParamsJson = $emailParams | ConvertTo-Json -Depth 10

$sendMailUri = "https://graph.microsoft.com/v1.0/users/$userEmail/sendMail"
$response = Invoke-RestMethod -Uri $sendMailUri -Method Post -Headers @{
    "Authorization" = "Bearer $accessToken"
    "Content-Type" = "application/json"
} -Body $emailParamsJson