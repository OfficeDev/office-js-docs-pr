---
title: Scenario - Implement single sign-on to your service
description: Learn about using the single-sign-on token and Exchange identity token provided by an Outlook add-in to implement SSO with your service.
ms.date: 08/14/2023
ms.topic: example-scenario
ms.localizationpriority: medium
---

# Scenario: Implement single sign-on to your service in an Outlook add-in

In this article we'll explore a recommended method of using the [single sign-on access token](authenticate-a-user-with-an-sso-token.md) and the [Exchange identity token](authenticate-a-user-with-an-identity-token.md) together to provide a single-sign on implementation to your own backend service. By using both tokens together, you can take advantage of the benefits of the SSO access token when it is available, while ensuring that your add-in will work when it is not, such as when the user switches to a client that does not support them, or if the user's mailbox is on an on-premises Exchange server.

> [!NOTE]
> The Single Sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](/javascript/api/requirement-sets/common/identity-api-requirement-sets).
> If you're working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy. For information about how to do this, see [Enable or disable modern authentication for Outlook in Exchange Online](/exchange/clients-and-mobile-in-exchange-online/enable-or-disable-modern-authentication-in-exchange-online).

## Why use the SSO access token?

The Exchange identity token is available in all requirement sets of the add-in APIs, so it may be tempting to just rely on this token and ignore the SSO token altogether. However, the SSO token offers some advantages over the Exchange identity token which make it the recommended method to use when it is available.

- The SSO token uses a standard OpenID format and is issued by Azure. This greatly simplifies the process of validating these tokens. In comparison, Exchange identity tokens use a custom format based on the JSON Web Token standard, requiring custom work to validate the token.
- The SSO token can be used by your backend to retrieve an access token for Microsoft Graph without the user having to do any additional sign in action.
- The SSO token provides richer identity information, such as the user's display name.

## Add-in scenario

For the purposes of this example, consider an add-in that consists of both the add-in UI and scripts (HTML + JavaScript) and a backend Web API that is called by the add-in. The backend Web API makes calls both to the [Microsoft Graph API](/graph/overview) and the Contoso Data API, a fictional API created by a third party. Like the Microsoft Graph API, the Contoso Data API requires OAuth authentication. The requirement is that the backend Web API should be able to call both APIs without having to prompt the user for their credentials every time an access token expires.

To do this, the backend API creates a secure database of users. Each user will get an entry in the database where the backend can store long-lived refresh tokens for both the Microsoft Graph API and the Contoso Data API. The following JSON markup represents a user's entry in the database.

```JSON
{
  "userDisplayName": "...",
  "ssoId": "...",
  "exchangeId": "...",
  "graphRefreshToken": "...",
  "contosoRefreshToken": "..."
}
```

The add-in includes either the SSO access token (if it is available) or the Exchange identity token (if the SSO token is not available) with every call it makes to the backend Web API.

### Add-in startup

1. When the add-in starts, it sends a request to the backend Web API to determine if the user is registered (i.e. has an associated record in the user database) and that the API has refresh tokens for both Graph and Contoso. In this call, the add-in includes both the SSO token (if available) and the identity token.

1. The Web API uses the methods in [Authenticate a user with an single-sign-on token in an Outlook add-in](authenticate-a-user-with-an-sso-token.md) and [Authenticate a user with an identity token for Exchange](authenticate-a-user-with-an-identity-token.md) to validate and generate a unique identifier from both tokens.

1. If an SSO token was provided, the Web API then queries the user database for an entry that has an `ssoId` value that matches the unique identifier generated from the SSO token.
   - If an entry did not exist, continue to the next step.
   - If an entry exists, proceed to step 5.

1. The Web API queries the database for an entry that has an `exchangeId` value that matches the unique identifier generated from the Exchange identity token.
   - If an entry exists and an SSO token was provided, update the user's record in the database to set the `ssoId` value to the unique identifier generated from the SSO token and proceed to step 5.
   - If an entry exists and no SSO token was provided, proceed to step 5.
   - If no entry exists, create a new entry. Set `ssoId` to the unique identifier generated from the SSO token (if available), and set `exchangeId` to the unique identifier generated from the Exchange identity token.

1. Check for a valid refresh token in the user's `graphRefreshToken` value.
   - If the value is missing or invalid and an SSO token was provided, use the [OAuth2 On-Behalf-Of flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of) to obtain an access token and refresh token for Graph. Save the refresh token in the `graphRefreshToken` value for the user.

1. Check for valid refresh tokens in both `graphRefreshToken` and `contosoRefreshToken`.
   - If both values are valid, respond to the add-in to indicate that the user is already registered and configured.
   - If either value is invalid, respond to the add-in to indicate that user setup is required, along with which services (Graph or Contoso) need to be configured.

1. The add-in checks the response.
   - If the user is already registered and configured, the add-in continues with normal operation.
   - If user setup is required, the add-in enters "setup" mode and prompts the user to authorize the add-in.

### Authorize the backend Web API

The procedure for authorizing the backend Web API to call the Microsoft Graph API and Contoso Data API should ideally only have to happen once, to minimize having to prompt the user for sign-in.

Based on the response from the backend Web API, the add-in may need to authorize the user for the Microsoft Graph API, the Contoso Data API, or both. Since both APIs use OAuth2 authentication, the method is similar for both.

1. The add-in notifies the user that it needs them to authorize their use of the API and asks them to click a link or button to start the process.

1. Once the flow completes, the add-in sends the refresh token to the backend Web API and includes the SSO token (if available) or the Exchange identity token.

1. The backend Web API locates the user in the database and updates the appropriate refresh token.

1. The add-in continues with normal operation.

### Normal operation

Whenever the add-in calls the backend Web API, it includes either the SSO token or the Exchange identity token. The backend Web API locates the user by this token, then uses the stored refresh tokens to obtain access tokens for the Microsoft Graph API and the Contoso Data API. As long as the refresh tokens are valid, the user will not have to sign in again.
