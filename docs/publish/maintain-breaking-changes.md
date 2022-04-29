---
title: Maintain your Office Add-in
description: Understand our commitments to compatibility and how to keep your add-in up to date.
ms.date: 04/29/2022
ms.localizationpriority: medium
---

# Maintain your Office Add-in

After you publish your add-in, you should keep it up to date with any important changes from upstream libraries. Patching security issues is critical to building customer trust. Since these changes have no effect on the published manifest, your customers won't need to take any actions to get the latest versions of your add-in.

## Breaking changes in Office.js

The Microsoft 365 Developer Platform is committed to ensuring the compatibility of your add-in. We strive to avoid making breaking changes to the API surface and behavior. However, there are cases where we need to make breaking updates for the sake of security or reliability. In those rare cases, the following steps are taken to ensure users of your add-in are unaffected.

- Announcements that describe the impacted features and recommended changes are made on the [Microsoft 365 Developer Blog](https://devblogs.microsoft.com/microsoft365dev/).
- If your add-in is published in [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), you'll be contacted through your provided information.
- Where possible, admins of impacted Microsoft 365 tenants (including [developer tenants](https://developer.microsoft.com/microsoft-365/dev-program)) will be contacted through [Message Center](/microsoft-365/admin/manage/message-center). It's the responsibility of the admin to contact providers of add-in solutions published outside of AppSource.

> [!IMPORTANT]
> [App Assure](https://www.microsoft.com/fasttrack/microsoft-365/app-assure) guarantees our support to your add-in's compatibility as Microsoft products are updated.

## Changes to Yeoman templates and web dependencies

The [Yeoman Generator for Office Add-ins](../develop/yeoman-generator-overview.md) relies on a number of libraries from Microsoft and others. These libraries are updated independently of any Microsoft 365 activity. Any projects created with the generator should be kept up to date as you develop, publish, and maintain your add-in. The following tools can help ensure your project is using secure versions of any dependent libraries.

- [npm audit](https://docs.npmjs.com/cli/v6/commands/npm-audit/)
- [Dependabot and other GitHub security features](https://github.com/features/security)

This guidance also applies to copies of samples taken from the [Office Add-in code samples](https://github.com/OfficeDev/Office-Add-in-samples) and other sources.

### office.js NPM package

The [office-js NPM package](https://www.npmjs.com/package/@microsoft/office-js) is a copy of what is hosted on the [Office.js content delivery network (CDN)](../develop/understanding-the-javascript-api-for-office.md#accessing-the-office-javascript-api-library). It's intended for scenarios where direct access to the CDN isn't possible. The NPM package isn't intended to provide versioned references to office.js. We strongly recommended always using the latest version.

## Current best practices

While we strive to maintain backwards compatibility, the patterns and practices we recommend continually evolve. Our documentation strives to present the current best practices. To stay informed of new features that may improve your existing functionality, join our monthly [Office Add-ins Community Call](../overview/office-add-ins-community-call.md).
