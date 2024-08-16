# oAuth2AzureEmail
These PowerShell scripts serve as an example of sending email through the Microsoft Graph API via oAuth2 (443/TCP) and not SMTP.

**GraphAPI-email.ps1** Simple example of sending an email via Graph API.<br />
**GraphAPI-email-with-CSV.ps1** Example of sending an email via Graph API with an attachment.<br />
**GraphAPI-email-with-CSV-report.ps1** More complex example of sending email via Graph API and converting an array into a CSV. This script requires additional permissions.<br />

## Azure Setup
To use any of these scripts you must setup the Azure components.

### Application Registration
First you need to register an application with Azure.

1. Log in to http://portal.azure.com
2. Select Microsoft Entra ID -> App Registrations or go directly to App Registrations
3. Click New Registration
4. Enter a unique name for your application (i.e. API-Email-Reports-PowerShell)
5. Select Accounts In This Organizational Directory Only (org only - Single tenant)
6. Click Register

### Setup the Secrets
Next you need a secret key pair for authentication.

1. Click Certificates & Secrets
2. Click New Client Secret
3. Enter a unique description (i.e. Intune report keys)
4. Select the expiration (The max is 24 months)
5. Click Add
6. Copy the Value from the secret and paste it into the script declarations

### Setup the Permissions
Last and most important is the permissions for the API. If you get a 40x you probably need to redo the permissions because you missed a step.

1. Click API Permissions
2. Click Add A Permission
3. Select Microsoft Graph
4. Select Application Permission
5. Type Mail in the search box
6. Expand Mail and select Mail.Send
7. Click Add Permissions
8. Click Grant Admin Consent for Organization
9. Confirm by selecting Yes
10. Click the three dots next to User.Read and select Revoke Admin Consent
11. Confirm by selecting Yes, Remove
12. Click the three dots next to User.Read and select Remove Permission
13. Confirm by selecting Yes, Remove
Note: The User.Read is not needed to send email.

### Setup Script
The scripts have generic declarations and will not work without setup.

1. Obtain the Application (Client) ID by clicking Overview
2. Copy the GUID to the script declarations
3. Obtain the Directory (Tenant) ID
4. Copy the GUID to the script declarations
5. Obtain the secret value from the prior step of setting the secret
6. Copy the GUID to the script declarations
7. Change your from Email and your to email
