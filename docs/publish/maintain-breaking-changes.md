---
title: Update and maintain your Office Add-in
description: Understand Microsoft's commitments to compatibility, and learn how to update your add-in, handle breaking changes in Office.js, and version your manifest.
ms.topic: best-practice
ms.date: 05/07/2026
ms.localizationpriority: medium
---

# Update and maintain your Office Add-in

After you publish your add-in, you should keep it up to date with any important changes from upstream libraries. Patching security issues is critical to building customer trust. Since these changes have no effect on the published manifest, your customers won't need to take any actions to get the latest versions of your add-in.

## Choose your scenario

- **I need to release a new version.** See [Update your add-in](#update-your-add-in).
- **An API in the Office JavaScript Library was deprecated; will it break my code?** See [Deprecation policy](#deprecation-policy).
- **My add-in broke after a platform update.** See [App Assure](#app-assure).
- **I need to update npm/Yeoman dependencies.** See [Keep your add-in secure](#keep-your-add-in-secure).

## Breaking changes in Office.js

We strive to avoid making breaking changes to the API surface and behavior. However, there are cases where we need to make breaking updates for the sake of security or reliability. In those rare cases, the following steps are taken to ensure users of your add-in are unaffected.

- Announcements that describe the impacted features and recommended changes are made on the [Microsoft 365 Developer Blog](https://devblogs.microsoft.com/microsoft365dev/).
- If your add-in is published in [Microsoft Marketplace](/partner-center/marketplace-offers/submit-to-appsource-via-partner-center), you'll be contacted through the information you provided.
- Where possible, admins of impacted Microsoft 365 tenants (including [Microsoft 365 Developer Program tenants](https://aka.ms/m365devprogram)) will be contacted through [Message Center](/microsoft-365/admin/manage/message-center). It's the responsibility of the admin to contact providers of add-in solutions published outside of Microsoft Marketplace.

### Deprecation policy

Tools may be deprecated when something better is available, and APIs in the Office JavaScript Library (Office.js) that are generally available (GA) may be deprecated when their task is better done with newer APIs. Microsoft makes a best effort to declare deprecations at least 24 months in advance. Deprecation doesn't necessarily mean the feature or API will be removed and unusable by developers. But after 24 months, Microsoft will no longer support the tool or API, and the tool *may* be retired and the API *may* removed from the GA version of Office.js.

When an API is marked as deprecated, we strongly recommend that you migrate to the replacement APIs as soon as possible. In some cases, we'll announce that new applications must start to use the new APIs a short time after the original APIs are deprecated. In those cases, only active applications that currently use the deprecated APIs can continue to use them.

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

## Update your add-in

When you update your add-in, there are two pieces to consider: the web application and the add-in package, including the manifest. 

### Update the web application

Updates to your web application don't require any action from your users. Upload new HTML, CSS, code, and other files to your web application and CDNs and your users will automatically start seeing the new features. Some add-in artifacts, such as button icons, are cached when the add-in is first installed. Users may not see updated icons initially. If users are confused at seeing the older icons, closing and reopening the Office application should cause the newest version of the artifacts to be downloaded. If that doesn't happen, you can instruct them on how to close all Office applications and then [clear the Office cache](../testing/clear-cache.md). Doing so will remove the cached artifacts from all of their installed add-ins, but they are downloaded again the next time the Office application launches, so the only effect of this is a very slight delay in the icons appearing.  

### Update the app package files

Changes to the manifest or other files in the app package do require users to update. If you have published your add-in to Microsoft Marketplace, you will need to update your submission. More information about that process is found in the article [Update an existing offer](/partner-center/marketplace-offers/update-existing-offer).

Whenever you make a change to the manifest, or to any file in the app package, you must raise the version number of the manifest. (This includes changes to URLs in the manifest that point to supplementary configuration files in the app package, such as Copilot configuration files.)

- If the add-in uses the unified manifest, see [version property](/microsoft-365/extensibility/schema/root#version).
- If the add-in uses the add-in only manifest (in which case there is no app package), see [Version element](/javascript/api/manifest/version).

If your add-in is deployed by one or more admins to their organizations, some manifest changes require the admin to consent to the updates. Users are blocked from the add-in until consent is granted. The following manifest changes require the admin to consent again.

- Changes to requested permissions. See [Requesting permissions for API use in add-ins](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../outlook/understanding-outlook-add-in-permissions.md).
- Adding or changing the Office applications that are supported by the add-in.
- Adding or changing [events](../develop/event-based-activation.md) that are supported by the add-in.

## Keep your add-in secure

Add-in projects that are created with the [Yeoman Generator for Office Add-ins](../develop/yeoman-generator-overview.md) or [Microsoft 365 Agents Toolkit](../develop/teams-toolkit-overview.md) rely on a number of libraries from Microsoft and others. These libraries are updated independently of any Microsoft 365 activity. Your projects should be kept up to date as you develop, publish, and maintain your add-in. The following tools can help ensure your project is using secure versions of any dependent libraries.

- [npm audit](https://docs.npmjs.com/cli/v6/commands/npm-audit/)
- [Dependabot and other GitHub security features](https://github.com/features/security)

This guidance also applies to copies of samples taken from the [Office Add-in code samples](https://github.com/OfficeDev/Office-Add-in-samples) and other sources.

## Troubleshooting

- **My users can't access the updated add-in.** Check that manifest version was incremented and Marketplace submission updated.
- **Admin consent dialogs aren't appearing.** Ensure that permissions/scopes/events actually changed. Test in your development tenant.

## Community engagement

As updates are proposed for the Microsoft 365 Developer Platform, we will be listening for feedback. Please report concerns, potential consequences, or other questions to the channels listed in [Office Add-ins additional resources](../resources/resources-links-help.md).

Stay informed of new features and evolving best practices through our monthly [Office Add-ins Community Call](../overview/office-add-ins-community-call.md).

## See also

- [Manage both a unified manifest and an add-in only manifest version of your Office Add-in](../concepts/duplicate-legacy-metaos-add-ins.md)