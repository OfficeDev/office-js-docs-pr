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

We don't recommend using the Visual Studio add-in templates because they use the add-in only manifest. We recommend add-ins have a unified manifest for the best experience with Integrated Apps and compatibility with Copilot plugins.

## I saw warnings after running `npm install`. How should I handle them?

In most cases, you can safely ignore warnings generated after running `npm install` for either Yo Office or the project. Sometimes, dependencies become deprecated and their replacements aren't supported by other packages that the project depends on. If you'd like to resolve these warnings, use the `npm-check-updates` tool.

- In the command prompt while in the root project directory, run `npm i -g npm-check-updates`. This installs the tool globally.
- Run `ncu -u`. This provides a report of all packages and what versions they will be updated to.
- Run `npm install` to update all the packages.

For more information about warnings when running `npm install`, see [Warnings and dependencies in the Node.js and npm world](../overview/npm-warnings-advice.md).

## Where did the Office Add-ins Developer Kit go?

The developer kit is no longer available in the Visual Studio Code extensions store. Add-ins created with the kit use the add-in only manifest. We recommend add-ins have a unified manifest for the best experience with Integrated Apps and compatibility with Copilot plugins.
