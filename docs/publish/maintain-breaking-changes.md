---
title: Maintain your Office Add-in
description: Understand our commitments to compatibility and how to keep your add-in up to date.
ms.topic: best-practice
ms.date: 01/23/2023
ms.localizationpriority: medium
---

# Maintain your Office Add-in

After you publish your add-in, you should keep it up to date with any important changes from upstream libraries. Patching security issues is critical to building customer trust. Since these changes have no effect on the published manifest, your customers won't need to take any actions to get the latest versions of your add-in.

## Breaking changes in Office.js

The Microsoft 365 Developer Platform is committed to ensuring the compatibility of your add-in. We strive to avoid making breaking changes to the API surface and behavior. However, there are cases where we need to make breaking updates for the sake of security or reliability. In those rare cases, the following steps are taken to ensure users of your add-in are unaffected.

- Announcements that describe the impacted features and recommended changes are made on the [Microsoft 365 Developer Blog](https://devblogs.microsoft.com/microsoft365dev/).
- If your add-in is published in [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), you'll be contacted through the information you provided.
- Where possible, admins of impacted Microsoft 365 tenants (including [developer tenants](https://developer.microsoft.com/microsoft-365/dev-program)) will be contacted through [Message Center](/microsoft-365/admin/manage/message-center). It's the responsibility of the admin to contact providers of add-in solutions published outside of AppSource.

### Deprecation policy

APIs or tools with better alternatives may be deprecated. Microsoft undergoes a best effort to declare something as deprecated at least 24 months in advance of retiring it. Similarly, for individual APIs that are generally available (GA), Microsoft declares an API as deprecated at least 24 months in advance of removing it from the GA version.

Deprecation doesn't necessarily mean the feature or API will be removed and unusable by developers. It does show that after the 24-month time period, Microsoft will no longer support the API or feature.

When an API is marked as deprecated, we strongly recommend that you migrate to the latest version as soon as possible. In some cases, we'll announce that new applications must start to use the new APIs a short time after the original APIs are deprecated. In those cases, only active applications that currently use the deprecated APIs can continue to use them.

> [!IMPORTANT]
> The 24-month deprecation period will be accelerated if waiting that long poses a security risk for your add-in or Microsoft.

### App Assure

Microsoft’s [App Assure](https://www.microsoft.com/fasttrack/microsoft-365/app-assure) service fulfills Microsoft’s promise of application compatibility: your apps will work on Windows and Microsoft 365 Apps. App Assure engineers are available to help resolve any issues you might experience at no additional cost.

If you do encounter an app compatibility issue, App Assure engineers will work with you to help you resolve the issue. Our experts will:

- Help you troubleshoot and identify a root cause.
- Provide guidance to help you remediate the application compatibility issue.
- Engage with independent software vendors (ISVs) on your behalf to remediate some part of their app, so that it’s functional on the most modern version of our products.
- Work with Microsoft product engineering teams to fix product bugs.

To learn more about App Assure, watch [Bring your apps to Microsoft Edge with App Assure: tips and tricks](https://techcommunity.microsoft.com/t5/video-hub/bring-your-apps-to-microsoft-edge-with-app-assure-tips-and/ba-p/2167619). To submit your request for app compatibility with App Assure, complete the [Microsoft FastTrack Registration form](https://aka.ms/AppAssureRequest) or send an email to [achelp@microsoft.com](mailto:achelp@microsoft.com).

## Changes to Yeoman templates and web dependencies

The [Yeoman Generator for Office Add-ins](../develop/yeoman-generator-overview.md) relies on a number of libraries from Microsoft and others. These libraries are updated independently of any Microsoft 365 activity. Any projects created with the generator should be kept up to date as you develop, publish, and maintain your add-in. The following tools can help ensure your project is using secure versions of any dependent libraries.

- [npm audit](https://docs.npmjs.com/cli/v6/commands/npm-audit/)
- [Dependabot and other GitHub security features](https://github.com/features/security)

This guidance also applies to copies of samples taken from the [Office Add-in code samples](https://github.com/OfficeDev/Office-Add-in-samples) and other sources.

### office.js NPM package

The [office-js NPM package](https://www.npmjs.com/package/@microsoft/office-js) is a copy of what is hosted on the [Office.js content delivery network (CDN)](../develop/understanding-the-javascript-api-for-office.md#accessing-the-office-javascript-api-library). It's intended for scenarios where direct access to the CDN isn't possible. The NPM package isn't intended to provide versioned references to office.js. We strongly recommend always using the CDN to ensure you're using the latest version of the Office JavaScript APIs.

## Current best practices

While we strive to maintain backwards compatibility, the patterns and practices we recommend continually evolve. Our documentation strives to present the current best practices. To stay informed of new features that may improve your existing functionality, join our monthly [Office Add-ins Community Call](../overview/office-add-ins-community-call.md).

## Deploy updates

[!INCLUDE [deploy-updates-that-require-admin-consent](../includes/deploy-updates-that-require-admin-consent.md)]

## Community engagement

As updates are proposed for the Microsoft 365 Developer Platform, we will be listening for feedback. Please report concerns, potential consequences, or other questions to the channels listed in [Office Add-ins additional resources](../resources/resources-links-help.md).
