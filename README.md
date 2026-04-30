# Macro-for-sending-MSFT-Teams-Message
Send a message from a Cisco video device (RoomOS or MTR) to a person via Microsoft Teams. 

> NOTE: Because the clientID and refresh token are stored in a macro on the device, precautions need to be made.
> Any person who has read-only access to the video device or Control Hub, will be able to view the clientId and refresh token.
> As such, you can look to leverage a centralised service where these values are protected. Alternatively ensure the OAuth scopes are limited as not to cause any problems if compromised.

## Setup
The setup steps are splits into two halves: provisioning the service account (Entra ID + Teams license) and authenticating it (app registration + device-code flow). 

**Create a dedicated user in Microsoft Entra ID**
* In the Entra admin center (entra.microsoft.com) → Users → New user → Create new user:
* UPN: something obvious like teams-bot-svc@yourtenant.onmicrosoft.com
* Set a strong password manually (don't auto-generate — you need to keep it)

**Assign a Teams-capable license**

The account needs an actual Teams license — being a "user" isn't enough; Teams provisioning happens on first sign-in to a licensed account. Any of these SKUs works:
* Microsoft 365 Business Basic / Standard / Premium
* Microsoft 365 E3 / E5
* Office 365 E1 / E3 / E5
* Or the standalone Microsoft Teams Essentials / Teams Enterprise SKU

Assign the license via Users → [the account] → Licenses → Assignments.
From the overview tab for the new user, copy the "Object ID" value to notepad.

**Sign in to Teams once, interactively**

This is the step people skip and then wonder why messages fail. Open teams.microsoft.com in a private browser, sign in as the service account, accept any first-run prompts. This provisions the user's chat service backend. Until this happens, POST /chats will throw weird "user not found in Teams" errors.

**Lock the account down**

Since you're keeping a password (or a long-lived refresh token) for an account with real privileges, compensate:
* Disable interactive sign-in from anywhere except where your service runs (Conditional Access named locations)
* No admin roles
* Rotate the password on a schedule

**Register the app**

In Entra → App registrations → New registration:
* Name: e.g. teams-bot-graph-client
* Supported account types: single tenant
* Redirect URI: leave blank for now

After creation:
* Authentication blade → enable Allow public client flows: Yes. 
* API permissions blade → add the Graph delegated scopes from before: User.Read, User.ReadBasic.All, Chat.Create, ChatMessage.Send, plus offline_access for refresh tokens. Click Grant admin consent.

Copy the Application (client) ID and Directory (tenant) ID to Notepad.

**Acquire a token (device-code flow)**

A number of HTTP REST calls need to be made. Install [Postman](https://postman.com/). 

Request a device code via 
```rest
POST https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/devicecode
Content-Type: application/x-www-form-urlencoded

client_id={app-client-id}
&scope=ChatMessage.Send Chat.Create User.Read offline_access
```

The response gives you the following values: user_code, verification_uri, device_code, interval. Visit the verification_uri url provided. Enter the device code. Sign in using the above Microsoft Teams account you created (that you applied the Teams license to). 

Request a token via
```rest
POST https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/token
Content-Type: application/x-www-form-urlencoded

grant_type=urn:ietf:params:oauth:grant-type:device_code
&client_id={app-client-id}
&device_code={device_code-from-above}
```

Copy the refresh_token.

**Configure the video device**

Install and upload each of the 3 macros provided as part of this repo, but do not enable them yet.
Select the "MicrosoftManageTokens" macro then update the value for clientId, tenantId and TEAMS_BOT_ID.
Select the "MicrosoftSavedTokens" macro and update the refresh_token value (leave the other values as they are)

Enable the "MicrosoftSendMessage" macro.

From the video device, select the newly added "Teams Message" button. Enter the partial or full name of someone. Select Confirm. That person should now receive a Microsoft Teams message.
