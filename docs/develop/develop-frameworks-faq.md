---
title: Office Add-ins development frameworks FAQ
description: Frequently asked questions about tools and scaffolding for Office Add-in projects.
ms.service: microsoft-365
ms.subservice: add-ins
ms.topic: faq
ms.date: 10/26/2026
---

# Office Add-ins development frameworks FAQ

## Can I use the Visual Studio add-in templates to get started on a new Office Add-in?

The Visual Studio add-in templates create projects with the add-in only manifest. For add-ins targeting supported clients and features, we recommend the unified manifest for the best experience with Integrated Apps and compatibility with Copilot plugins. An add-in only manifest is still required for unsupported targets. See [Client and platform support](unified-manifest-overview.md#client-and-platform-support) and [Feature support](unified-manifest-overview.md#feature-support).

## I saw warnings after running `npm install`. How should I handle them?

In most cases, you can safely ignore deprecation warnings generated after running `npm install` for either Yo Office or the project. Security or audit warnings require separate assessment; address vulnerabilities in runtime dependencies, or dependencies whose origin you can't determine, before deploying to production. Sometimes, dependencies become deprecated and their replacements aren't supported by other packages that the project depends on. To review available dependency updates, use the `npm-check-updates` tool.

- In the command prompt while in the root project directory, run `npm i -g npm-check-updates`. This installs the tool globally.
- Run `ncu` to report available updates without modifying **package.json**.
- Review the proposed versions for compatibility and update only the appropriate dependency ranges in **package.json**, then run `npm install`. Note that `ncu -u` rewrites dependency ranges and can include breaking major-version upgrades.

For more information about warnings when running `npm install`, see [Warnings and dependencies in the Node.js and npm world](../overview/npm-warnings-advice.md).

## Where did the Office Add-ins Developer Kit go?

The developer kit is no longer available in the Visual Studio Code extensions store. Add-ins created with the kit use the add-in only manifest. We recommend add-ins have a unified manifest for the best experience with Integrated Apps and compatibility with Copilot plugins.
