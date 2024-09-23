---
title: Nested app authentication and Outlook legacy tokens deprecation FAQ
description: Nested app authentication and Outlook legacy tokens deprecation FAQ
ms.service: microsoft-365
ms.subservice: add-ins
ms.topic: faq
ms.date: 09/23/2024
---

# Nested app authentication and Outlook legacy tokens deprecation FAQ

Exchange [user identity tokens](https://learn.microsoft.com/office/dev/add-ins/outlook/authentication#exchange-user-identity-token) and [callback tokens](https://learn.microsoft.com/office/dev/add-ins/outlook/authentication#callback-tokens) are deprecated and will begin to be turned off in October 2024. We recommend moving Outlook add-ins that use legacy Exchange tokens to nested app authentication.

## General FAQ

### What is nested app authentication (NAA)?

Nested app authentication enables single sign-on (SSO) for applications nested inside of supported Microsoft applications such as Outlook. Compared with existing full-trust authentication models, and the on-behalf-of flow, NAA provides better security and greater flexibility in app architecture, enabling the creation of rich, client-driven applications. For more information, see [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md).

### What is the timeline for shutting down legacy Exchange online tokens?

The following tables list the key milestones based on which channel customers are using. Note that the general availability (GA) date for NAA will vary based on channel. We'll provide tooling for administrators to reenable Exchange tokens for tenants and add-ins if those add-ins aren't yet migrated to NAA.

| Date     | Release channel(s) | Legacy tokens status and NAA General Availability (GA) |
| -------- | ------------------ | ------------------------------------------------------ | 
| Oct 2024 | All channels | New PowerShell options for enabling/disabling legacy tokens for entire tenant or specific AppIDs. |
| Oct 2024 | Current Channel | Legacy tokens turned off for <b>tenants not using them</b>; NAA will GA in Current Channel. |
| Nov 2024 | Monthly Enterprise Channel | Legacy tokens turned off for <b>tenants not using them</b>; NAA will GA in Monthly Enterprise Channel. |
| Jan 2025 | Current and Semi-Annual Channels | Legacy tokens turned off for all tenants in Current and Semi-Annual Channels. Admins can reenable via PowerShell. NAA will GA in Semi-Annual Channels. |
| Feb 2025 | Monthly Enterprise Channel | Legacy tokens turned off for all tenants in Monthly Enterprise. Admins can reenable via PowerShell. |
| Jun 2025 | Semi-Annual Extended Channel | Legacy tokens off for all tenants in Semi-Annual Extended Channel. NAA will GA in Semi-Annual Extended Channel. |
| Jun 2025 | All channels | Admins can no longer re-enable legacy tokens via PowerShell; contact Microsoft. |
| Oct 2025 | All channels | Legacy tokens turned off for all tenants, there will be no re-enable option. |

> [!NOTE]
> If a single tenant uses multiple Microsoft 365 apps / Office release channels, Legacy Exchange Online tokens will be turned off based on the "slowest" release channel.

## Outlook add-in migration FAQ

### Why is Microsoft making Outlook add-ins migrate?

Switching to Microsoft Graph using Entra ID tokens is a big improvement in security for Outlook and Exchange customers. Entra ID (formerly Azure Active Directory) is a leader in the identity and access management space. Customers can take advantage of zero trust features such as conditional access, MFA requirements, continual token monitoring, real time safety heuristics, and more that aren't available with legacy Exchange tokens. Customers have much of their most important business data stored in Exchange, so it's vital that we ensure this data is protected. Migrating the whole Outlook ecosystem to using Entra ID tokens with Microsoft Graph greatly improves security for customer data.

### Can I opt out?

We'll provide tooling via PowerShell for Microsoft 365 administrators in October 2024 to turn legacy Exchange tokens on or off in your tenant. You can use it to ensure add-ins aren't broken if they haven't updated to use NAA yet. However, in June 2025, legacy Exchange Online tokens will be turned off and you won't be able to turn them back on without a specific exception granted by Microsoft. In October 2025 it won't be possible to turn on legacy Exchange Online tokens and they'll be disabled for all tenants. We'll update this FAQ when more information about this tooling is available.

### Does my Outlook add-in have to migrate to NAA?

No. Outlook add-ins don't have to use NAA, although NAA offers the best authentication experience for users and best security posture for organizations. If add-ins aren't using legacy Exchange tokens, they won't be affected by the deprecation of Exchange tokens. Add-ins using MSAL.js or other SSO methods that rely on Entra ID will continue to work.

### How do I know if my Outlook add-in relies on legacy tokens?

To find out whether your add-in uses legacy Exchange user identity tokens and callback tokens, you can search your code for calls to the following APIs.

- `makeEwsRequestAsync`
- `getUserIdentityTokenAsync`
- `getCallbackTokenAsync`

If your add-in calls any of these APIs, you should adopt NAA and migrate to using Entra ID tokens to access Microsoft Graph instead.

Also, We'll provide tooling via PowerShell for Microsoft 365 administrators in October 2024 to turn legacy Exchange tokens on or off in your tenant. This will allow you to test if any add-ins are using Exchange tokens. We'll provide more information when the tooling is ready in this FAQ.

### Which Outlook add-ins are in scope?

Many of our most important add-ins are in scope. If your add-in is using EWS or Outlook REST to access Exchange Online resources, it almost certainly needs to migrate off of legacy Outlook tokens to NAA.
If your add-in is for Exchange on-premises only (for example, Exchange 2019), it is not affected by this announcement.

### What will happen to my Outlook add-ins if I don't migrate to NAA?

If you don't migrate your Outlook add-ins to NAA, they'll stop working as expected in Exchange Online. When Exchange tokens are turned off (according to the previous tables), Exchange Online will block legacy token issuance. Any add-in that uses legacy tokens won't be able to access Exchange online resources.
If your add-in only works on-premises or if your add-in is on a deprecation path, you may not need to update. However, most add-ins that access Exchange resources through EWS or Outlook REST will have to migrate to continue functioning as expected.

### How do I migrate my Outlook add-ins to NAA?

To support NAA in your Outlook add-in, please refer to the following documentation and sample.

- [Enable SSO in an Office Add-in using nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md).
- [Outlook add-in with SSO using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA).

### As an admin, how do I know which add-ins in my org need to be updated?

Add-ins may use the legacy Exchange tokens to get resources from Exchange through the EWS or Outlook REST APIs. Sometimes an add-in requires Exchange resources for some use cases and not others, making it difficult to figure out whether the add-in requires an update. We recommend reaching out to add-in developers and owners to ask them if their add-in code references the following APIs:

- `makeEwsRequestAsync`
- `getUserIdentityTokenAsync`
- `getCallbackTokenAsync`

If you rely on an independent software vendor (ISV) for your add-in, we recommend you contact them as soon as possible to confirm they have a plan and a timeline for moving off legacy Exchange tokens. ISV developers should reach out directly to their Microsoft contacts with questions to ensure they're ready for the end of Exchange legacy tokens. If you rely on a developer within your organization, we recommend you ask them to review the [Updates on deprecating legacy Exchange Online tokens for Outlook add-ins blog](https://devblogs.microsoft.com/microsoft365dev/updates-on-deprecating-legacy-exchange-online-tokens-for-outlook-add-ins/?commentid=1131) and ask any questions to the Outlook extensibility PM team on the [OfficeDev/office-js GitHub issues site](https://github.com/OfficeDev/office-js/issues).

### As an admin, I don't own an add-in that needs an update. What should I do?

If you rely on an independent software vendor (ISV) for your add-in, we recommend reaching out as soon as possible to confirm they have a plan and a timeline for moving off of legacy Exchange tokens. ISV developers should reach out directly to their Microsoft contacts with questions to ensure they're ready for the end of Exchange legacy tokens.
If you rely on a developer within your organization, we recommend asking them to reach out to the Outlook extensibility PM team on GitHub: [https://github.com/officedev/office-js/issues](https://github.com/officedev/office-js/issues).

### How do I keep up with the latest guidance?

We'll share additional guidance moving forward on the [Office Add-ins community call](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins-community-call) and the [M365 developer blog](https://devblogs.microsoft.com/microsoft365dev/).
Finally, you can ask questions about NAA and legacy Exchange Online token deprecation on the OfficeDev/office-js GitHub issues site. Please put "NAA" in the title so we can group and prioritize issues.

## Developer troubleshooting questions

### NAA is not providing SSO and keeps prompting users to sign in

This can occur when NAA is not available in the Outlook client. If on Windows, check that you are using either the Beta Channel, or Current Channel (Preview). You need to join the [Microsoft 365 Insider Program](https://insider.microsoft365.com/join/windows) to switch to these channels.
A good way to check if NAA is available is to check the requirement set using the following code snippet.
`Office.context.requirements.isSetSupported("NestedAppAuth")`

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
- [NAA public preview blog](https://aka.ms/NAApreviewblog)
- Microsoft 365 developer blog, [Updates on deprecating legacy Exchange Online tokens for Outlook add-ins](https://devblogs.microsoft.com/microsoft365dev/updates-on-deprecating-legacy-exchange-online-tokens-for-outlook-add-ins/?commentid=1131)
