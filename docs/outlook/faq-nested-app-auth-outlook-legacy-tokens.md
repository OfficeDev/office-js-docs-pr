---
title: Nested app authentication FAQ
description: Nested app authentication FAQ
ms.service: microsoft-365
ms.subservice: add-ins
ms.topic: faq
ms.date: 11/03/2025
---

# Nested app authentication FAQ

## What is nested app authentication (NAA)?

Nested app authentication enables single sign-on (SSO) for applications nested inside of supported Microsoft applications such as Outlook. Compared with existing full-trust authentication models, and the on-behalf-of flow, NAA provides better security and greater flexibility in app architecture, enabling the creation of rich, client-driven applications. For more information, see [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md).

## Can I use Exchange Online tokens instead of NAA?

Legacy Exchange Online [user identity tokens](authentication.md#exchange-user-identity-token) and [callback tokens](authentication.md#callback-tokens) are no longer supported and turned off across all Microsoft 365 tenants. If an Outlook add-in requires delegated user access or user identity, we recommend using MSAL (Microsoft Authentication Library) and nested app authentication.

If your add-in calls an API that requests an Exchange token, such as `getUserIdentityTokenAsync`, it gets a generic error similar to the following with error codes such as 9017 or 9018.

- "GenericTokenError: An internal error has occurred."
- "InternalServerError: The Exchange server returned an error. Please look at the diagnostics object for more information."

Note that for Exchange on-premises, Exchange tokens still work and are supported.

## How do I report an issue with NAA?

Ask questions about NAA on the [OfficeDev/office-js GitHub issues site](https://github.com/OfficeDev/office-js/issues). Please put "NAA" in the title so we can group and prioritize issues.

If you submit an issue, please include the following information.

- Outlook client version.
- Outlook release channel audience (for client).
- Screen capture of the issue.
- The platform where the issue occurs (Windows, Outlook (new), Mac, iOS, Android).
- Session id where the issue is encountered.
- Type of account being used.
- Version of msal-browser.
- Logs from msal-browser.

## How do I get more debug information from MSAL and NAA?

Use the following code to enable debug information in the msalConfig when you initialize the nestable public client application. This will log additional details to the console.

```javascript
const msalConfig = {
  auth: {...},
  system: {
    loggerOptions: {
      logLevel: LogLevel.Verbose,
      loggerCallback: (level, message, containsPii) => {
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            return;
          case LogLevel.Info:
            console.info(message);
            return;
          case LogLevel.Verbose:
            console.debug(message);
            return;
          case LogLevel.Warning:
            console.warn(message);
            return;
        }
      },
    }
  }
};
```

## What tokens does MSAL return and are there minimum scopes to request?

When you request a token through MSAL, it always returns three tokens.

|Token          |Purpose  |Scopes  | `AuthencationResult` property |
|---------------|---------|---------|----------------------------|
|ID token | Provides information about the user to the client (task pane). | `profile` and `openid` | `authResult.idToken` |
|Refresh token  | Refreshes the ID and access tokens when they expire.     | `offline_access`       | Not available. |
|Access token   | Authenticates the user for specific scopes to a resource, such as Microsoft Graph. | Any resource scopes, such as `user.read`. | `authResult.accessToken` |

MSAL always returns these three tokens. It requests the `profile`, `openid`, and `offline_access` as default scopes even if your token request doesn't include them. This ensures the ID and refresh tokens are requested. However, you must include at least one resource scope, such as `user.read` so that you get an access token. If not, the request can fail.

## Should I validate the ID token from MSAL?

No. This is a legacy authentication pattern that was used with Exchange tokens to authorize access to your own resources. Passing the ID token over a network call to enable or authorize access to a service is a security anti-pattern. The ID token is intended only for the client (task pane) and there is no way for the service to reliably use the token to be sure the user has authorized access. For more information about ID token claims, see [ID token claims reference](/entra/identity-platform/id-token-claims-reference).

It's very important that you always request an access token to your own services. The access token also includes the same ID claims, so you don't need to pass the ID token. Instead create a custom scope for your service. For more information about app registration settings for your own services, see [Protected web API: App registration](/entra/identity-platform/scenario-protected-web-api-app-registration). When your service receives the access token, it can validate it, and use ID claims from inside the access token.

## Why am I getting errors from conditional access policies?

The **approved client app Conditional Access grant** is deprecated and will retire in March 2026. MSAL NAA does not support this policy and will return errors (even if you grant the add-in an exception to this policy.) To migrate off of this policy, see [Migrate approved client app to application protection policy in Conditional Access](/entra/identity/conditional-access/migrate-approved-client-app).

Some conditional access policies will cause issues for add-ins using MSAL NAA depending on what they require from the client. Often these are related to device management policies. For more information, see device management types in [How to create and assign app protection policies](/intune/intune-service/apps/app-protection-policies).

Sometimes you need to handle claims challenges based on policies. To learn more on how to handle a claims challenge in your add-in, see [Claims challenges, claims requests and client capabilities](/entra/identity-platform/claims-challenge).

## Why is the ID token not refreshed?

There is a known issue where MSAL sometimes doesn't refresh the ID token after it expires. This shouldn't cause any issues in your add-in since the ID token is only intended for use in your task pane to get basic user identity information, such as name and email. There's no reason to validate the ID token or check the expiration claim. If you need to authenticate the user to your own resources, use the access token which also contains user identity information. The ID token must never be passed outside of your client code that received it.

## In Outlook, how do I determine if the user is an online or on-premise account?

You can determine if the signed-in user has an Exchange Online account or on-premise Exchange account by using the [Office.UserProfile.accountType](/javascript/api/outlook/office.userprofile) property. If the account type property value is **enterprise**, then the mailbox is on an on-premises Exchange server. Note that volume-licensed perpetual Outlook 2016 doesn’t support the **accountType** property. To work around this, call the [ResolveNames](/exchange/client-developer/web-service-reference/resolvenames-operation) operation in Exchange Web Service (EWS) in the Exchange on-premise server to get the recipient types.

> [!NOTE]
> Outlook 2016 and Outlook 2019 are no longer supported. For more information, see [End of support for Office 2016 and Office 2019](https://support.microsoft.com/office/818c68bc-d5e5-47e5-b52f-ddf636cf8e16).

The **accountType** property requires mailbox requirement set 1.6. On older Outlook clients you’ll need to use the Autodiscover service as follows.

Call the Autodiscover endpoint for the outlook.office365.com domain. `https://outlook.office365.com/autodiscover/autodiscover.json/v1.0/{email}?Protocol=EWS&ServerLocation=true`

- For **online** accounts, the service will return a result with the `ServerLocation` key set to Exchange Online.
- For **on-premise** accounts, the service will not return a `ServerLocation` key.

> [!NOTE]
> For customers that use vanity URLs, you need to specifically configure your add-in to call the Autodiscover service on the vanity URL endpoint.

## How do I deploy my add-in to Microsoft Marketplace

If you're publishing a new add-in to Microsoft Marketplace, it will need to go through a certification process. For more information, see [Publish your Office Add-in to Microsoft Marketplace](../publish/publish-office-add-ins-to-appsource.md). If you're updating the manifest of an add-in that is already published to Microsoft Marketplace, you need to go through the certification process again. You can update the add-in's source code on your web server any time without a need to go through the certification process.

If you're add-in uses SSO through NAA, your add-in must be in compliance with the following publishing guidelines.

- [1000.3 Authentication options](/legal/marketplace/certification-policies#10003-authentication-options)
- [1120.3 Functionality](/legal/marketplace/certification-policies#11203-functionality)

Be sure to handle admin consent properly. See [Publish an add-in that requires admin consent for Microsoft Graph scopes](../publish/publish-nested-app-auth-add-in.md)

For additional deployment details, see [Make your solutions available in Microsoft Marketplace and within Office](/partner-center/marketplace-offers/submit-to-appsource-via-partner-center). If you update your add-in (change the manifest) you need to go through the [certification process again](../publish/publish-nested-app-auth-add-in.md). You can update your web server code any time without a need for review.

## Users get an unexplained error when signing in

When your add-in requests a token, users may see a sign-in popup dialog showing one of the following errors.

- **Something went wrong.** [*error code*]
- **You can't get there from here**

Check to see if the admin has any conditional access policies applied that enforce specific client restrictions, such as mobile location, or platform type. Also the **approved client app Conditional Access grant** is deprecated and will cause these errors with NAA token requests. An admin must completely remove this policy and switch over to the newer **application protection policy grant** for NAA to work. For more information, see [Migrate approved client app to application protection policy in Conditional Access](/entra/identity/conditional-access/migrate-approved-client-app).

## Related content

- [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md).
- [Outlook add-in with SSO using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA).
