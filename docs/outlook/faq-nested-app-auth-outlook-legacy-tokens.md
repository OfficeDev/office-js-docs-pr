---
title: Nested app authentication and Outlook legacy tokens deprecation FAQ
description: Nested app authentication and Outlook legacy tokens deprecation FAQ
ms.service: microsoft-365
ms.subservice: add-ins
ms.topic: faq
ms.date: 10/01/2025
---

# Nested app authentication and Outlook legacy tokens deprecation FAQ

Legacy Exchange Online [user identity tokens](authentication.md#exchange-user-identity-token) and [callback tokens](authentication.md#callback-tokens) are deprecated and are turned off across all Microsoft 365 tenants. If an Outlook add-in requires delegated user access or user identity, we recommend using MSAL (Microsoft Authentication Library) and nested app authentication.

## General FAQ

### What is nested app authentication (NAA)?

Nested app authentication enables single sign-on (SSO) for applications nested inside of supported Microsoft applications such as Outlook. Compared with existing full-trust authentication models, and the on-behalf-of flow, NAA provides better security and greater flexibility in app architecture, enabling the creation of rich, client-driven applications. For more information, see [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md).

### What is the timeline for shutting down legacy Exchange online tokens?

Legacy Exchange Online tokens are turned off. If you're an admin for a tenant and were granted an exemption from Microsoft for your tenant, most of this FAQ will still apply to you. All exemptions end on October 31st, 2025. **No more exemptions are allowed**.

### When is NAA generally available for my channel?

The general availability (GA) date for NAA depends on which channel you're using. The following table lists build and GA information for Outlook.

| Date     | NAA General Availability (GA) for Outlook |
| -------- | ------------------------------------------------------ |
| October 2024 | NAA is GA in Current Channel. |
| November 2024 | NAA is GA in Monthly Enterprise Channel. |
| January 2025 | NAA is GA in Semi-Annual Channel Version 2408 (Build 17928.20392). |
| June 2025 | NAA is GA in Semi-Annual Extended Channel Version 2408 (Build 17928.20604). |

### Are COM Add-ins affected by the deprecation of legacy Exchange Online tokens?

It's very unlikely any COM add-ins are affected by the deprecation of legacy Exchange Online tokens. Outlook web add-ins are primarily affected because they can use Office.js APIs that rely on Exchange tokens. For more information, see [How do I know if my Outlook add-in relies on legacy tokens?](#how-do-i-know-if-my-outlook-add-in-relies-on-legacy-tokens). The Exchange tokens are used to access Exchange Web Services (EWS) or Outlook REST APIs, both of which are also deprecated.

## Microsoft 365 administrator questions

### Which add-ins in my organization are impacted?

Use the `Get-AuthenticationPolicy` command to get a list of all Outlook add-ins that use legacy Exchange Online tokens on your tenant. For more information, see [Turn legacy Exchange Online tokens on or off](turn-exchange-tokens-on-off.md). Once you have the list of add-ins, you’ll need to reach out to the publishers to learn more about their plans to update. In some cases, the add-in may be developed by your own organization. You’ll need to reach out to the appropriate development team in your organization.

### What commands can I use to identify the publisher?

There are some Exchange Online PowerShell commands you can use to track down additional information about Outlook add-ins.

To find a list of add-ins installed on a user’s computer, the user can run the following command.

`Get-App | Select-Object -Property ProviderName, DisplayName, AppId`

The following screenshot shows an example of running the `Get-App` command.

:::image type="content" source="../images/get-app-cmdlet-providername.png" alt-text="Screenshot of running the Get-App command in PowerShell with results for Microsoft Polls and Microsoft Send to OneNote.":::

The **ProviderName** will help you identify who published the add-in so that you can contact them. The **AppId** can be used to get additional details about the add-in.

> [!NOTE]
> The `Get-App` command doesn’t show a complete list of all add-ins installed on the user’s computer. For example, sideloaded add-ins will not appear in this list. You may need to follow up with users in some cases to track down where the add-in came from.

To find information about an add-in by `AppId` use the following command.

`Get-App -Identity {identity} | Select-Object -Property ProviderName, DisplayName`

The following screenshot shows an example of using the ID of Bing Maps to get more information.

:::image type="content" source="../images/get-app-cmdlet-bing-maps.png" alt-text="Screenshot of running the Get-App command in PowerShell to get the ProviderName and DisplayName for Bing Maps.":::

You may also find additional information in the add-in's manifest file. The manifest contains URL endpoints which can also help you identify and contact the publisher. Use the following command to get the manifest.

`Get-App -Identity {identity} | Select-Object -Property ManifestXml`

The following screenshot shows an example of using the ID to get the XML manifest for Bing Maps.

:::image type="content" source="../images/get-app-cmdlet-bing-maps-manifestxml.png" alt-text="Screenshot of running the Get-App command in PowerShell to get the ManifestXml of Bing Maps":::

### How would ISVs know their add-in is using legacy tokens?

Add-ins may use the legacy tokens to get resources from Exchange through the EWS or Outlook REST APIs. Sometimes an add-in requires Exchange resources for some use cases and not others, making it difficult to figure out whether the add-in requires an update. We recommend reaching out to add-in developers and owners to ask them if their add-in code references the following APIs.

- `makeEwsRequestAsync`
- `getUserIdentityTokenAsync`
- `getCallbackTokenAsync`

If you rely on an ISV for your add-in, we recommend you contact them as soon as possible to confirm they have a plan and a timeline for moving off of legacy Exchange tokens. ISV developers should reach out directly to their Microsoft contacts with questions to ensure they're ready for the end of Exchange legacy tokens. If you rely on a developer within your organization, they should review this FAQ and the article [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md). Any questions should be raised on the [OfficeDev/office-js GitHub issues site](https://github.com/OfficeDev/office-js/issues).

### What do I do for add-ins I can't identify?

If you come across add-ins you can’t identify after running `Get-AuthenticationPolicy`, try performing a scream test to determine ownership.

> [!NOTE]
> You only need to perform the scream test if you turned legacy Exchange Online tokens on by using the `Set-AuthenticationPolicy` command. If you haven't run this command, then Exchange Online tokens should already be off by default.

Before performing the scream test you may want to let your users know in advance, such as through email, that there will be a test to turn off legacy tokens and that it may affect some Outlook add-ins. You should consider providing users the following information.

- The expected time period of the test.
- If there are known Outlook add-ins that will break, such as add-ins deployed from Microsoft Marketplace that you’ve already identified.
- That in general, Outlook add-ins shouldn’t break. However, if they do see issues, ask users to report the name, and description of the add-in, along with any error information observed.

Use the following steps to perform the test.

1. Run the following command to turn off legacy Exchange Online tokens on your tenant. For details on how to use this command, see [Turn legacy Exchange Online tokens on or off](turn-exchange-tokens-on-off.md).
    
    `Set-AuthenticationPolicy -BlockLegacyExchangeTokens -Identity "LegacyExchangeTokens"`
    
1. Wait a suitable amount of time for users to report any issues with add-ins. It takes approximately 24 hours for the command to turn off legacy Exchange Online tokens for all users. It may take another day or two for users to report any issues with Outlook add-ins.
1. Identify any affected Outlook add-ins. If users submit issues identifying breaking issues, be sure to get the name and description of the Outlook add-in affected. Also capture the error, or behavior so this information can be passed along to the publisher.
1. If any business-critical add-ins are broken, turn tokens back on using the following command. For details on how to use this command, see [Turn legacy Exchange Online tokens on or off](turn-exchange-tokens-on-off.md).
    
    `Set-AuthenticationPolicy -AllowLegacyExchangeTokens -Identity "LegacyExchangeTokens"`
    
    It takes approximately 24 hours for tokens to turn back on for all users on the tenant.
    
1. If there are no reports of breaking issues, we recommend you leave legacy Exchange Online tokens off as a security best practice.

### Can I turn Exchange Online legacy tokens back on?

You can only turn legacy tokens on if you were granted an exemption from Microsoft. For more information on how to turn legacy tokens on or off, see [Turn legacy Exchange Online tokens on or off](turn-exchange-tokens-on-off.md). All exemptions end on October 31st, 2025. **No more exemptions are allowed**.

### How does the admin consent flow work?

Independent software vendors (ISVs) are updating their add-ins to use Entra ID tokens and Microsoft Graph scopes. When the add-in requests an access token, it must have admin or user consent. If the administrator consents, all users on the tenant can use the add-in for the scopes the add-in requires. Otherwise, each end-user will be prompted for consent, if [user consent is enabled](/microsoft-365/admin/misc/user-consent). For a better experience because the users aren't prompted, complete admin consent.

One option for consent is that the ISV provides you with an admin consent URI.

1. The add-in developer provides an admin consent URI. If this is not in the documentation they provide, you need to contact them for more information.
1. The administrator browses to the admin consent URI.
1. The administrator is prompted to sign in and consent to a list of scopes that the add-in requires.
1. Once complete, the browser redirects to a web page from the ISV, which should show the consent was successful.

As an alternative, the ISV may provide an updated app manifest that will prompt for admin consent as part of central deployment. In this scenario, when you deploy the updated app manifest, you'll be prompted to consent before the deployment can complete. There is no need for an admin consent URI.

Finally, if the add-in is published in the Microsoft 365 store, the update will deploy automatically and the administrator will be prompted to consent to the scopes. If the administrator doesn't consent, users won't be able to use the updated add-in.

### What if the add-in doesn't work after admin consent?

Ensure you don't disable features, or revoke permissions that the add-in requires. For an example, see [modifying mailbox policy properties](/exchange/clients-and-mobile-in-exchange-online/outlook-on-the-web/configure-outlook-web-app-mailbox-policy-properties). The add-in uses delegated permissions and therefore has access to the same resources as the signed-in user. However, if a policy or setting blocks the user from a particular resource or action, the add-in will also be blocked.

### How do I deploy add-in updates from an ISV?

If you have an add-in that uses legacy Exchange tokens, you should reach out to your ISV for information about their timeline to migrate their add-in to use NAA. Once the ISV migrates their add-in, they will most likely provide an admin consent URL. For more information, see [How does the admin consent flow work?](#how-does-the-admin-consent-flow-work) .

The ISV may also provide you with an updated app manifest to deploy through centralized deployment. During centralized deployment, this may prompt you to consent to any Microsoft Graph scopes the add-in requires. In this scenario you won't need to use an admin consent URI.

If the add-in is deployed from Microsoft Marketplace, most likely you'll be prompted to consent to Microsoft Graph scopes when the ISV rolls out updates to the add-in. Until you consent, users on the tenant won't be able to use the new version of the add-in with NAA.

### Where do I find which add-ins have consent?

Once the admin or a user consents, it will be listed in the Microsoft Entra admin center. You can find app registrations using the following steps.

1. Go to [https://entra.microsoft.com/#home](https://entra.microsoft.com/#home) and sign in as admin on your tenant.
1. In the left navigation pane, select **Applications** > **Enterprise applications**.
1. On the **Enterprise applications** page, in the **Manage** section, select **All applications**.
1. Select the Add-in. This will open an overview page.
In the overview page, select **Permissions**.
There are two views for permissions; Admin consent, and User consent. Select User consent to see any individual consents.

### Is there a list of publishers that have updated their add-ins?

Some widely used Outlook add-in publishers have already updated their add-ins as listed below.

- [Appfluence Priority Matrix for Outlook](https://appsource.microsoft.com/product/office/wa104381735)
- [Atlassian Jira Cloud for Outlook](https://marketplace.atlassian.com/apps/1220666/jira-cloud-for-outlook-official?tab=overview&hosting=cloud)
- [Box for Outlook](https://appsource.microsoft.com/product/office/WA200000015)
- [Clickup for Outlook](https://appsource.microsoft.com/product/office/WA104382026)
- [iEnterprises® - Outlook Connector](https://ienterprises.com/connector/outlook-connector)
- [HubStar Connect](https://www.hubstar.com/solutions/connect/)
- [SalesForce for Outlook](https://appsource.microsoft.com/product/office/wa104379334)
- [LawToolBox](https://lawtoolbox.com/lawtoolbox-for-copilot/)
- [OnePlace Solutions](https://www.oneplacesolutions.com/oneplacemail-sharepoint-app-for-outlook.html)
- [Set-OutlookSignatures Benefactor Circle](https://set-outlooksignatures.com)
- [Wrike](https://appsource.microsoft.com/product/office/wa104381120)
- [Zoho CRM for Email](https://appsource.microsoft.com/product/office/WA104379468)
- [Zoho Recruit for Email](https://appsource.microsoft.com/product/office/WA200001485)
- [Zoho Projects for Email](https://appsource.microsoft.com/product/office/WA200006712)
- [Zoho Sign for Outlook](https://appsource.microsoft.com/product/office/WA200002326)
- [Zoho WorkDrive for Email](https://appsource.microsoft.com/product/office/WA200006673)
- [Invoice and Time Tracking - Zoho Invoice](https://appsource.microsoft.com/product/office/WA104381067)

If the publisher updated their manifest, and the add-in is deployed through the Microsoft store, you'll be prompted as an administrator to upgrade and deploy the updates. If the publisher updated their manifest, and the add-in is deployed through central deployment, you'll need to deploy the new manifest as an administrator. In some cases the publisher may have an admin consent URI you need to use to consent to new scopes for the add-in. Reach out to publishers if you need more information about updating an add-in.

## Outlook add-in migration FAQ

### Why is Microsoft making Outlook add-ins migrate?

Switching to Microsoft Graph using Entra ID tokens is a big improvement in security for Outlook and Exchange customers. Entra ID (formerly Azure Active Directory) is a leading cloud-based identity and access management service. Customers can take advantage of zero trust features such as conditional access, MFA requirements, continual token monitoring, real time safety heuristics, and more that aren't available with legacy Exchange tokens. Customers store important business data stored in Exchange, so it's vital that we ensure this data is protected. Migrating the whole Outlook ecosystem to use Entra ID tokens with Microsoft Graph greatly improves security for customer data.

### Does my Outlook add-in have to migrate to NAA?

No. Outlook add-ins don't have to use NAA, although NAA offers the best authentication experience for users and the best security posture for organizations. If add-ins aren't using legacy Exchange tokens, they won't be affected by the deprecation of Exchange tokens. Add-ins using MSAL.js or other SSO methods that rely on Entra ID will continue to work.

### How do I know if my Outlook add-in relies on legacy tokens?

To find out whether your add-in uses legacy Exchange user identity tokens and callback tokens, search your code for calls to the following APIs.

- `makeEwsRequestAsync`
- `getUserIdentityTokenAsync`
- `getCallbackTokenAsync`

If your add-in calls any of these APIs, you should adopt NAA and migrate to using Entra ID tokens to access Microsoft Graph instead.

### Which Outlook add-ins are in scope?

Many major add-ins are in scope. If your add-in is using EWS or Outlook REST to access Exchange Online resources, it almost certainly needs to migrate off of legacy Outlook tokens to NAA.
If your add-in is for Exchange on-premises only (for example, Exchange 2019), it's not affected by this change.

### What will happen to my Outlook add-ins if I don't migrate to NAA?

If you don't migrate your Outlook add-ins to NAA, they'll stop working as expected in Exchange Online. When Exchange tokens are turned off, Exchange Online will block legacy token issuance. Any add-in that uses legacy tokens won't be able to access Exchange online resources. When your add-in calls an API that requests an Exchange token, such as `getUserIdentityTokenAsync`, it gets a generic error similar to the following with error codes such as 9017 or 9018.

- "GenericTokenError: An internal error has occurred."
- "InternalServerError: The Exchange server returned an error. Please look at the diagnostics object for more information."

If your add-in only works on-premises or if your add-in is on a deprecation path, you may not need to update. However, most add-ins that access Exchange resources through EWS or Outlook REST must migrate to continue functioning as expected.

### How do I migrate my Outlook add-ins to NAA?

To support NAA in your Outlook add-in, please refer to the following documentation and sample.

- [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md).
- [Outlook add-in with SSO using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA).

### How do I keep up with the latest guidance?

We'll update this FAQ as any new information becomes available. We'll share additional guidance moving forward on the [Office Add-ins community call](../overview/office-add-ins-community-call.md) and the [M365 developer blog](https://devblogs.microsoft.com/microsoft365dev/).
Finally, you can ask questions about NAA and legacy Exchange Online token deprecation on the [OfficeDev/office-js GitHub issues site](https://github.com/OfficeDev/office-js/issues). Please put "NAA" in the title so we can group and prioritize issues.

If you submit an issue, please include the following information.

- Outlook client version.
- Outlook release channel audience (for client).
- Screen capture of the issue.
- The platform where the issue occurs (Windows, Outlook (new), Mac, iOS, Android).
- Session id where the issue is encountered.
- Type of account being used.
- Version of msal-browser.
- Logs from msal-browser.

## Developer questions

### How do I get more debug information from MSAL and NAA?

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

### Test your updated add-in

Once you've updated your add-in to use NAA, you should test it on all platforms you support, such as Mac, mobile, web, and Outlook on Windows.

#### Test when Exchange tokens turned off

To test that your add-in works correctly when Exchange tokens are turned off, deploy your add-in to a tenant with tokens turned off and test it. To turn tokens off, see [Turn legacy Exchange Online tokens on or off](turn-exchange-tokens-on-off.md).

If you've implemented a pattern where your code uses Exchange tokens but then falls over if they are unavailable, be sure you are checking for the correct errors. When a call to get an Exchange token fails, check the [asyncResult.diagnostics](/javascript/api/office/office.asyncresult). If either of the following errors is returned, switch to NAA.

- `GenericTokenError: An internal error has occurred.`
- `InternalServerError: The Exchange server returned an error. Please look at the diagnostics object for more information.`

#### Test fallback code for Trident+ webview

If your Outlook add-in supports Outlook 2016 or Outlook 2019 on Windows, test that it works correctly when the Trident+ (Internet Explorer 11) webview is used. When the Trident+ webview is used, your code must fall back to MSAL v2 to open a dialog and sign in the user. For more information on how to implement the fallback pattern, see [Outlook add-in with SSO using nested app authentication including Internet Explorer fallback](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA-IE).

> [!NOTE]
> Support for Outlook 2016 and Outlook 2019 ends October 2025. For more information, see [End of support for Office 2016 and Office 2019](https://support.microsoft.com/office/818c68bc-d5e5-47e5-b52f-ddf636cf8e16).

#### Testing in Trident+ and WebView2

Outlook 2016 and Outlook 2019 on Windows use the Trident+ or WebView2 based on various OS conditions.

- For more information on when Trident+ or Webview2 is used, see [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).
- For more information on how to determine which webview is running, see  [Support older Microsoft webviews and Office versions](../develop/support-ie-11.md#determine-the-webview-the-add-in-is-running-in-at-runtime)

> [!NOTE]
> Support for Outlook 2016 and Outlook 2019 ends October 2025. For more information, see [End of support for Office 2016 and Office 2019](https://support.microsoft.com/office/818c68bc-d5e5-47e5-b52f-ddf636cf8e16).

### What tokens does MSAL return and are there minimum scopes to request?

When you request a token through MSAL, it always returns three tokens.

|Token          |Purpose  |Scopes  | `AuthencationResult` property |
|---------------|---------|---------|----------------------------|
|ID token | Provides information about the user to the client (task pane). | `profile` and `openid` | `authResult.idToken` |
|Refresh token  | Refreshes the ID and access tokens when they expire.     | `offline_access`       | Not available. |
|Access token   | Authenticates the user for specific scopes to a resource, such as Microsoft Graph. | Any resource scopes, such as `user.read`. | `authResult.accessToken` |

MSAL always returns these three tokens. It requests the `profile`, `openid`, and `offline_access` as default scopes even if your token request doesn't include them. This ensures the ID and refresh tokens are requested. However, you must include at least one resource scope, such as `user.read` so that you get an access token. If not, the request can fail.

### Should I validate the ID token from MSAL?

No. This is a legacy authentication pattern that was used with Exchange tokens to authorize access to your own resources. Passing the ID token over a network call to enable or authorize access to a service is a security anti-pattern. The ID token is intended only for the client (task pane) and there is no way for the service to reliably use the token to be sure the user has authorized access. For more information about ID token claims, see [ID token claims reference](/entra/identity-platform/id-token-claims-reference).

It's very important that you always request an access token to your own services. The access token also includes the same ID claims, so you don't need to pass the ID token. Instead create a custom scope for your service. For more information about app registration settings for your own services, see [Protected web API: App registration](/entra/identity-platform/scenario-protected-web-api-app-registration). When your service receives the access token, it can validate it, and use ID claims from inside the access token.

### Why am I getting errors from conditional access policies?

The **approved client app Conditional Access grant** is deprecated and will retire in March 2026. MSAL NAA does not support this policy and will return errors (even if you grant the add-in an exception to this policy.) To migrate off of this policy, see [Migrate approved client app to application protection policy in Conditional Access](/entra/identity/conditional-access/migrate-approved-client-app).

Some conditional access policies will cause issues for add-ins using MSAL NAA depending on what they require from the client. Often these are related to device management policies. For more information, see device management types in [How to create and assign app protection policies](/intune/intune-service/apps/app-protection-policies).

Sometimes you need to handle claims challenges based on policies. To learn more on how to handle a claims challenge in your add-in, see [Claims challenges, claims requests and client capabilities](/entra/identity-platform/claims-challenge).

### Why is the ID token not refreshed?

There is a known issue where MSAL sometimes doesn't refresh the ID token after it expires. This shouldn't cause any issues in your add-in since the ID token is only intended for use in your task pane to get basic user identity information, such as name and email. There's no reason to validate the ID token or check the expiration claim. If you need to authenticate the user to your own resources, use the access token which also contains user identity information. The ID token must never be passed outside of your client code that received it.

### How do I determine if the user is an online or on-premise account?

You can determine if the signed-in user has an Exchange Online account or on-premise Exchange account by using the [Office.UserProfile.accountType](/javascript/api/outlook/office.userprofile) property. If the account type property value is **enterprise**, then the mailbox is on an on-premises Exchange server. Note that volume-licensed perpetual Outlook 2016 doesn’t support the **accountType** property. To work around this, call the [ResolveNames](/exchange/client-developer/web-service-reference/resolvenames-operation) operation in Exchange Web Service (EWS) in the Exchange on-premise server to get the recipient types.

> [!NOTE]
> Support for Outlook 2016 and Outlook 2019 ends October 2025. For more information, see [End of support for Office 2016 and Office 2019](https://support.microsoft.com/office/818c68bc-d5e5-47e5-b52f-ddf636cf8e16).

The **accountType** property requires mailbox requirement set 1.6. On older Outlook clients you’ll need to use the Autodiscover service as follows.

Call the Autodiscover endpoint for the outlook.office365.com domain. `https://outlook.office365.com/autodiscover/autodiscover.json/v1.0/{email}?Protocol=EWS&ServerLocation=true`

- For **online** accounts, the service will return a result with the `ServerLocation` key set to Exchange Online.
- For **on-premise** accounts, the service will not return a `ServerLocation` key.

> [!NOTE]
> For customers that use vanity URLs, you need to specifically configure your add-in to call the Autodiscover service on the vanity URL endpoint.

### How do I deploy my add-in to Microsoft Marketplace

If you're publishing a new add-in to Microsoft Marketplace, it will need to go through a certification process. For more information, see [Publish your Office Add-in to Microsoft Marketplace](../publish/publish-office-add-ins-to-appsource.md). If you're updating the manifest of an add-in that is already published to Microsoft Marketplace, you need to go through the certification process again. You can update the add-in's source code on your web server any time without a need to go through the certification process.

If you're add-in uses SSO through NAA, your add-in must be in compliance with the following publishing guidelines.

- [1000.3 Authentication options](/legal/marketplace/certification-policies#10003-authentication-options)
- [1120.3 Functionality](/legal/marketplace/certification-policies#11203-functionality)

Be sure to handle admin consent properly. See [Publish an add-in that requires admin consent for Microsoft Graph scopes](../publish/publish-nested-app-auth-add-in.md)

For additional deployment details, see [Make your solutions available in Microsoft Marketplace and within Office](/partner-center/marketplace-offers/submit-to-appsource-via-partner-center). If you update your add-in (change the manifest) you need to go through the [certification process again](../publish/publish-nested-app-auth-add-in.md). You can update your web server code any time without a need for review.

### Users get an unexplained error when signing in

When your add-in requests a token, users may see a sign-in popup dialog showing one of the following errors.

- **Something went wrong.** [*error code*]
- **You can't get there from here**

Check to see if the admin has any conditional access policies applied that enforce specific client restrictions, such as mobile location, or platform type. Also the **approved client app Conditional Access grant** is deprecated and will cause these errors with NAA token requests. An admin must completely remove this policy and switch over to the newer **application protection policy grant** for NAA to work. For more information, see [Migrate approved client app to application protection policy in Conditional Access](/entra/identity/conditional-access/migrate-approved-client-app).

## Related content

- [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md).
- [Outlook add-in with SSO using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA).
