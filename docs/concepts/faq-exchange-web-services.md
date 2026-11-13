---
title: Exchange Web Services (EWS) deprecation FAQ for Office Add-ins
description: Learn how Exchange Online EWS retirement affects Office Add-ins, Office clients, and tenant administrators.
ms.service: microsoft-365
ms.subservice: add-ins
ms.topic: faq
ms.date: 11/13/2026
---

# Exchange Web Services (EWS) deprecation FAQ for Office Add-ins

## What is being deprecated?

Exchange Web Services (EWS) is being retired in Exchange Online. Phased disablement began in October 2026, and permanent shutdown starts April 1, 2027. This retirement doesn't apply to EWS in Exchange Server on-premises.

Add-ins and Office client features that depend on Exchange Online EWS must move to supported alternatives. For add-in code that calls EWS, Microsoft recommends migrating to Microsoft Graph.

This is separate from the retirement of legacy Exchange Online user identity and callback tokens, which are already turned off across all Microsoft 365 tenants.

For more information, see [Deprecation of Exchange Web Services in Exchange Online](/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online).

## Which Office client versions and applications are affected?

The Exchange Online EWS retirement can affect add-ins and Office client features that depend on EWS. For the Windows desktop versions of Word, Excel, and PowerPoint, we're backporting changes to LTSC 2024 that replace EWS calls used by the affected add-in-management features.

Earlier Office LTSC releases aren't planned to receive this backport because they'll be out of service before the permanent Exchange Online EWS shutdown in April 2027.

## Can my add-in be affected even if its code doesn't call EWS?

Yes. An add-in can be affected when the Office client uses Exchange Online EWS to discover, deploy, update, or manage the add-in, even if the add-in's own code and backend services don't call EWS.

The affected Office client scenarios include centralized deployment and private catalog operations. This is different from an add-in or backend service making its own EWS calls. Developers must still migrate any EWS calls made by their add-in or service to a supported alternative, such as Microsoft Graph.

## Do I need to update my add-in because of the EWS retirement?

If your add-in calls EWS in an Exchange Online environment, directly or through a backend service, migrate those calls to a supported alternative such as Microsoft Graph. This applies to Excel, Outlook, PowerPoint, and Word add-ins.

For Outlook add-ins, legacy Exchange Online user identity and callback tokens have already been retired. This is separate from the EWS retirement. For more information, see [Nested app authentication FAQ](../outlook/faq-nested-app-auth-outlook-legacy-tokens.md).

Using the Office JavaScript API doesn't, by itself, mean your add-in needs changes. Review its API calls and authentication flow to determine whether it's affected. If you use an add-in developed by someone else, contact its vendor or developer to confirm whether an update is needed.

These Exchange Online retirements don't apply to Exchange Server on-premises.

## What behavior might users see on an affected older Office client?

Depending on the client version and deployment method, users might find that deployed or private-catalog add-ins are missing, that the available add-in list appears empty, or that add-in information doesn't refresh as expected.

During the transition away from EWS, incompatible cached data can also prevent Office from displaying add-in information correctly. The LTSC 2024 mitigation includes recovery for invalid cached data and improved handling of unsuccessful service responses.

These symptoms don't, by themselves, confirm that EWS retirement is the cause. Record the Office application, version, build, deployment method, Exchange environment, and when the behavior began before troubleshooting or contacting support.

## My add-in stopped working in or after October 2026. Could the phased EWS retirement be the cause?

Possibly, but the date when the problem began isn't enough to confirm the cause. EWS retirement is a potential cause if:

- The affected mailbox is in Exchange Online.
- The add-in, its backend service, or an Office client feature that manages the add-in depends on EWS.
- The failure began after EWS was disabled for the tenant.

Ask your tenant administrator to review the organization's EWS Usage Report and follow the troubleshooting steps in [Exchange Online EWS application stopped working? Do this.](https://techcommunity.microsoft.com/blog/exchange/exchange-online-ews-application-stopped-working-do-this-/4566215).

## What actions do I need to take as a tenant administrator?

Review [Deprecation of Exchange Web Services in Exchange Online](/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online) for tools to identify EWS usage and migration guidance.

## See also

- [Deprecation of Exchange Web Services in Exchange Online](/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online)
- [Exchange Online EWS, Your Time is Almost Up](https://techcommunity.microsoft.com/blog/exchange/exchange-online-ews-your-time-is-almost-up/4492361)
- [Exchange Online EWS application stopped working? Do this.](https://techcommunity.microsoft.com/blog/exchange/exchange-online-ews-application-stopped-working-do-this-/4566215)
- [Use the Microsoft Graph REST API from an Outlook add-in](../outlook/microsoft-graph.md)
