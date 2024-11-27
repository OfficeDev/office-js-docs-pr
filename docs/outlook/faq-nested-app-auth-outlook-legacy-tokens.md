---
title: Nested app authentication and Outlook legacy tokens deprecation FAQ
description: Nested app authentication and Outlook legacy tokens deprecation FAQ
ms.service: microsoft-365
ms.subservice: add-ins
ms.topic: faq
ms.date: 11/26/2024
---

# Nested app authentication and Outlook legacy tokens deprecation FAQ

Exchange [user identity tokens](authentication.md#exchange-user-identity-token) and [callback tokens](authentication.md#callback-tokens) are deprecated and will be turned off starting in February 2025. We recommend moving Outlook add-ins that use legacy Exchange tokens to nested app authentication.

## General FAQ

### What is nested app authentication (NAA)?

Nested app authentication enables single sign-on (SSO) for applications nested inside of supported Microsoft applications such as Outlook. Compared with existing full-trust authentication models, and the on-behalf-of flow, NAA provides better security and greater flexibility in app architecture, enabling the creation of rich, client-driven applications. For more information, see [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md).

### What is the timeline for shutting down legacy Exchange online tokens?

Microsoft begins turning off legacy Exchange online tokens in February 2025. From now until February 2025, existing and new tenants will not be affected.  We'll provide tooling for administrators to reenable Exchange tokens for tenants and add-ins if those add-ins aren't yet migrated to NAA. See [Can I turn legacy tokens back on?](#can-i-turn-legacy-tokens-back-on) for more information.

| Date     | Legacy tokens status |
| -------- | ------------------------------------------------------ |
| Feb 2025 | Legacy tokens turned off for all tenants. Admins can reenable legacy tokens via PowerShell. |
| Jun 2025 | Legacy tokens turned off for all tenants. Admins can no longer reenable legacy tokens via PowerShell and must contact Microsoft for any exception. |
| Oct 2025 | Legacy tokens turned off for all tenants. Exceptions are no longer allowed. |

### When is NAA generally available for my channel?

The general availability (GA) date for NAA depends on which channel you are using.

| Date     | NAA General Availability (GA) |
| -------- | ------------------------------------------------------ |
| Oct 2024 | NAA is GA in Current Channel. |
| Nov 2024 | NAA is GA in Monthly Enterprise Channel. |
| Jan 2025 | NAA will GA in Semi-Annual Channel. |
| Jun 2025 | NAA will GA in Semi-Annual Extended Channel. |

## Microsoft 365 administrator questions

### Can I turn legacy tokens back on?

Yes, there are PowerShell commands you can use to turn legacy tokens on or off. Currently the commands are only intended for use in test tenants and testing Outlook add-ins. Don't use them in production tenants or you'll see side effects. We'll update the commands soon so that they can be used in production tenants as well. We'll update this FAQ with additional information once the commands can be used in production tenants.

For more information on how to turn legacy tokens on or off, see [Turn legacy Exchange Online tokens on or off](turn-exchange-tokens-on-off.md).

In June 2025, legacy tokens will be turned off and you won't be able to turn them back on without a specific exception granted by Microsoft. In October 2025, it won't be possible to turn on legacy tokens and they'll be disabled for all tenants. We'll update this FAQ with additional information once the tool is available 

### How does the admin consent flow work?

Independant software vendors (ISVs) are updating their add-ins to use Entra ID tokens and Microsoft Graph scopes. When the add-in requests an access token, it must have admin or user consent. If the administrator consents, all users on the tenant can use the add-in for the scopes the add-in requires. Otherwise, each end-user will be prompted for consent, if [user consent is enabled](/microsoft-365/admin/misc/user-consent). Completing admin consent provides a better experience because the users aren't prompted.

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

If the add-in is deployed from Microsoft AppSource, most likely you'll be prompted to consent to Microsoft Graph scopes when the ISV rolls out updates to the add-in. Until you consent, users on the tenant won't be able to use the new version of the add-in with NAA.

### Which add-ins in my organization are impacted?

We published a list of all Outlook add-ins published to the Microsoft store that use legacy tokens as of October 2024. For more information on how to use the list and build a report of Outlook add-ins that are potentially using legacy tokens, see [Find Outlook add-ins that use legacy Exchange Online tokens](https://aka.ms/naafaq). Also we're working on report tooling to make tracking add-ins using legacy tokens easier. We hope to have the report tooling available in early 2025.

Add-ins may use the legacy Exchange tokens to get resources from Exchange through the EWS or Outlook REST APIs. Sometimes an add-in requires Exchange resources for some use cases and not others, making it difficult to figure out whether the add-in requires an update. We recommend reaching out to add-in developers and owners to ask them if their add-in code references the following APIs.

- `makeEwsRequestAsync`
- `getUserIdentityTokenAsync`
- `getCallbackTokenAsync`

If you rely on an ISV for your add-in, we recommend you contact them as soon as possible to confirm they have a plan and a timeline for moving off of legacy Exchange tokens. ISV developers should reach out directly to their Microsoft contacts with questions to ensure they're ready for the end of Exchange legacy tokens. If you rely on a developer within your organization, they should review this FAQ and the article [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md). Any questions should be raised on the [OfficeDev/office-js GitHub issues site](https://github.com/OfficeDev/office-js/issues).

### Where do I find which add-ins have consent?

Once the admin or a user consents, it will be listed in the Microsoft Entra admin center. You can find app registrations using the following steps.

1. Go to [https://entra.microsoft.com/#home](https://entra.microsoft.com/#home).
1. In the left navigation pane, select **Applications** > **App registrations**.
1. On the **App registrations** page, select **All applications**.
1. Now you can search for any app registration by name or ID.

### Is there a list of publishers that have updated their add-ins?

Some widely used Outlook add-in publishers have already updated their add-ins as listed below.

| Outlook add-in  | More information  |
|---------|---------|---------|
| Salesforce for Outlook   | [SalesForce](https://appsource.microsoft.com/en-us/product/office/wa104379334)        |
| iEnterprises® - Outlook Connector  | [iEnterprises](https://ienterprises.com/connector/outlook-connector/) |
| HubStar Connect  | [HubStar](https://www.hubstar.com/solutions/connect/) |

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

If you don't migrate your Outlook add-ins to NAA, they'll stop working as expected in Exchange Online. When Exchange tokens are turned off, Exchange Online will block legacy token issuance. Any add-in that uses legacy tokens won't be able to access Exchange online resources.

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

## Developer troubleshooting questions

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

## Related content

- [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md).
- [Outlook add-in with SSO using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA).