Run e-mail merge with:

 python parse_recipient_info.py  --type excel --sheet 'MailMergeDataset' wafunif\ contact\ list.xlsx \
    --tokenfile tokens.json --subject "Request for Contact Information of Current Interns and Fellows" \
    --username 'l.bucur.cress@wafunif.org' --from_name 'WAFUNIF - World Association of Former United Nations Internes and Fellows' \
    --from_addr 'membership@wafunif.org' --email_template email_templates/requestforinterncontactinfo.html --no-test



Login to retrieve token with:

python m365_token_helper.py login --tenant '27fc3673-45b3-4ea0-8d45-04210f6434e9' --client-id '328053a5-2571-4315-bf19-7b1ed545a658'


Refresh token with:

python m365_token_helper.py refresh --tenant '27fc3673-45b3-4ea0-8d45-04210f6434e9' --client-id '328053a5-2571-4315-bf19-7b1ed545a658' --in tokens.json --out tokens.json

or use the helper script:
$> refresh_token.sh


Build with:
mkdir build
cd build
cmake ..
cmake --build .




Cmake Build Dependencies:

libcurl4-openssl-dev


Python Dependencies
openpyxl
requests
jinja2



export OAUTH2_TOKEN="eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9..."



Prerequisites

Admin consent: Your Azure AD tenant admin must allow SMTP AUTH for your mailbox.

In Microsoft 365 Admin Center
, under Users → Active Users → Mail, ensure SMTP AUTH is enabled for the account.

App registration in Azure AD:

Go to Azure Portal
 → Azure Active Directory → App registrations → New registration.

Give it a name, e.g., SMTP Client.

Choose "Accounts in this organizational directory only" (unless you need multi-tenant).

Register.

2. Configure the app

Under Certificates & secrets, create a client secret. Copy it (you’ll only see it once).

Under API permissions, add:

Microsoft Graph → Delegated permission → SMTP.Send

(You might also add offline_access so you can refresh tokens.)

Grant admin consent to these permissions.

3. Acquire tokens

You’ll need:

client_id (from the app registration),

client_secret,

tenant_id,

and the mailbox username (the account you want to send as).

The OAuth 2.0 flow is usually the Resource Owner Password (ROPC) or Authorization Code with PKCE. For service apps, ROPC is sometimes simpler, but it may be disabled by policy. Recommended way is Authorization Code Flow.

Example with curl (Authorization Code flow)

Direct the user (or yourself) to a consent URL:

https://login.microsoftonline.com/<TENANT_ID>/oauth2/v2.0/authorize?
    client_id=<CLIENT_ID>
    &response_type=code
    &redirect_uri=http://localhost
    &response_mode=query
    &scope=https%3A%2F%2Foutlook.office365.com%2FSMTP.Send%20offline_access


Sign in, consent, and copy the code returned to your redirect URI.

Exchange the code for tokens:

curl -X POST https://login.microsoftonline.com/<TENANT_ID>/oauth2/v2.0/token \
  -d client_id=<CLIENT_ID> \
  -d scope=https://outlook.office365.com/SMTP.Send offline_access \
  -d code=<AUTH_CODE_FROM_STEP_1> \
  -d redirect_uri=http://localhost \
  -d grant_type=authorization_code \
  -d client_secret=<CLIENT_SECRET>


The response JSON contains:

{
  "access_token": "eyJ0eXAiOiJKV1QiLCJhbGciOi...",
  "refresh_token": "0.AAA...",
  "expires_in": 3599,
  ...
}


The access_token is what you export as OAUTH2_TOKEN.

4. Using the token in your C program
export OAUTH2_TOKEN="eyJ0eXAiOiJKV1QiLCJhbGciOi..."
./m365_smtp


The program will then use XOAUTH2 with Microsoft’s SMTP server.

5. Refreshing the token

Access tokens last about 1 hour. Use the refresh_token to get a new one:

curl -X POST https://login.microsoftonline.com/<TENANT_ID>/oauth2/v2.0/token \
  -d client_id=<CLIENT_ID> \
  -d scope=https://outlook.office365.com/SMTP.Send offline_access \
  -d refresh_token=<REFRESH_TOKEN> \
  -d grant_type=refresh_token \
  -d client_secret=<CLIENT_SECRET>


That gives you a new access_token.

✅ Summary:

Register an app in Azure AD.

Add SMTP.Send + offline_access delegated permissions.

Get client_id, client_secret, tenant_id.

Use OAuth2 Authorization Code flow to obtain an access_token.

Export it as OAUTH2_TOKEN for your program.













How I registered the app:

Steps to Register an App in Azure AD (Microsoft Entra) for SMTP OAuth2

These steps assume you are an administrator or have rights to manage Azure AD / Microsoft Entra in your tenant.

Log into Azure Portal

Go to portal.azure.com

Sign in with an admin account in the tenant where you want to allow SMTP OAuth2.

Register a New App

Navigate to Azure Active Directory → App registrations → + New registration.

Enter a Name (e.g. SMTP OAuth2 Sender).

Choose supported account types (e.g. Accounts in this organizational directory only unless you need multi-tenant).

Set a Redirect URI if needed (for some flows like authorization code / interactive login). If you're using client credentials (app-only), you may not need an interactive redirect URI.

Save. Record the Application (client) ID and Directory (tenant) ID from the Overview page.

Add Permissions

In the App registration’s API Permissions pane, click Add a permission.

You want to add permissions for Exchange or Microsoft Graph, depending on Microsoft’s current policies. For SMTP OAuth2 (legacy protocols: IMAP/POP/SMTP), you may need the Office 365 Exchange Online permissions. 
Microsoft Learn
+1

Specific permissions to add:

SMTP.Send (application permission) if available.

Possibly POP.AccessAsApp / IMAP.AccessAsApp if you also need those. 
Microsoft Learn
+1

If SMTP.Send is not visible in the Azure UI, you may still use the scope https://outlook.office365.com/SMTP.Send in your token request. 
Stack Overflow
+1

Grant Admin Consent

After adding the required permissions, click Grant admin consent in the same pane.

This often requires a tenant admin. This ensures that your app can use those permissions across the tenant.

Create a Client Secret

In the App registration, go to Certificates & secrets → New client secret.

Give it a description and referential expiration (30 days, 90 days, etc.). When created, copy the secret value immediately — you will not be able to see it again later.

(If Needed) Register Service Principal / Exchange Permissions

For SMTP with OAuth2 (legacy protocol access), Exchange Online may require that you register the app / service principal in Exchange and grant it mailbox permissions. For example, allowing SendAs permissions if you're sending from a mailbox that is not the same as the authenticating identity. 
Email Architect
+1

In some tenants, you need to ensure SmtpClientAuthenticationDisabled setting is set to false for the mailbox. This ensures SMTP AUTH is enabled per mailbox. 
Email Architect

Use the Correct Token Request Parameters

When requesting your access token, you’ll use something like:

POST https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token
Content-Type: application/x-www-form-urlencoded

grant_type=client_credentials
client_id={your_client_id}
client_secret={your_client_secret}
scope=https://outlook.office365.com/.default


The scope .default means "use all application permissions you granted".

Ensure SSL/TLS, correct endpoints, etc.

The resulting access_token is what you use with AUTH XOAUTH2 in SMTP.

Testing

Once you have a token, test your SMTP connection with verbose logging to see the authentication exchange.

Make sure the username (UPN) is correct. Some examples require that the username in SMTP XOAUTH2 be the UPN for the mailbox you're sending from.

The From address must be valid and allowed. If sending as another mailbox, ensure you have SendAs or Send on behalf permissions.




****** Info on how to login/refresh the token ******

python m365_token_helper.py login \
      --tenant 27fc3673-45b3-4ea0-8d45-04210f6434e9 \
      --client-id 328053a5-2571-4315-bf19-7b1ed545a658 \
      --scopes "https://outlook.office365.com/SMTP.Send offline_access" \
      --out tokens.json



  python m365_token_helper.py refresh \
      --tenant 27fc3673-45b3-4ea0-8d45-04210f6434e9 \
      --client-id 328053a5-2571-4315-bf19-7b1ed545a658 \
      --in tokens.json \
      --out tokens.json




****** Azure AD application Registration Information ******


Display name
:
SMTP OAuth2 Sender


Application (client) ID
:
328053a5-2571-4315-bf19-7b1ed545a658


Object ID
:
d8c2b98e-6901-449e-ada0-87fb85d1539f


Directory (tenant) ID
:
27fc3673-45b3-4ea0-8d45-04210f6434e9


Supported account types
:
My organization only


Client credentials
:
Add a certificate or secret


Redirect URIs
:
Add a Redirect URI


Application ID URI
:
Add an Application ID URI


Managed application in local directory
:
SMTP OAuth2 Sender






