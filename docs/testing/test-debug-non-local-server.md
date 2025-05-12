---
title: Test and debug Office Add-ins on a non-local server
description: Learn how to test and debug your Office Add-in on a non-local host.
ms.date: 05/19/2025
ms.localizationpriority: medium
---

# Test and debug Office Add-ins on a non-local server

When you've completed development and testing on a localhost and want to stage and test the add-in from a non-local server or cloud account, you can use the tool [office-addin-debugging](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-debugging) for any Node.js-based add-in project. (The tool isn't supported in projects created with Visual Studio.)

> [!NOTE]
> If you're working on a Windows computer, you may have another option for non-local testing. See [Sideload Office Add-ins for testing from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

## Projects created with Microsoft 365 Agents Toolkit or the Office Yeoman Generator (Yo Office)

If your project was created with [Agents Toolkit](../develop/agents-toolkit-overview.md) or [Office Yeoman Generator (Yo Office)](../develop/yeoman-generator-overview.md), then the office-addin-debugging tool is already installed and your package.json file has `start` and `stop` scripts that invoke the tool. To use it for non-local testing, update the domain part of the URLs in your manifest to point to your staging server (or CDN as needed). Then run `npm run start` at the command line (or Visual Studio Code TERMINAL) to sideload the add-in for testing and debugging.

> [!IMPORTANT]
> The office-addin-debugging tool registers the add-in in the Windows registry or a special folder on a Mac. For an Outlook add-in, it also registers the add-in in Exchange. To avoid subtle bugs when developing, always end a testing session by running `npm run stop` to ensure that these registrations are removed and that the server process is fully stopped. *Manually closing the server, the command line window (or TERMINAL), Visual Studio Code, or the Office application doesn't remove these registrations.*

## Other projects

If your project wasn't created with Agents Toolkit or Yo Office, run the tool with npx in the root of the project. Invoke it with its `start` command followed by the relative path to the manifest. The following is an example.

```command&nbsp;line
npx office-addin-debugging start manifest.json
```

This command sideloads the add-in for testing and debugging. The tool also works with an add-in only manifest.

There are many options for the `start` command. For details, see the README for the tool at [office-addin-debugging](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-debugging).

> [!IMPORTANT]
> The office-addin-debugging tool registers the add-in in the Windows registry or a special folder on a Mac. For an Outlook add-in, it also registers the add-in in Exchange. To avoid subtle bugs when developing, always end a testing session by running `npx office-addin-debugging stop` to ensure that these registrations are removed and that the server process is fully stopped. *Manually closing the server, the command line window (or TERMINAL), Visual Studio Code, or the Office application doesn't remove these registrations.* If you used the `--prod` option with the `start` command, use the same option with the `stop` command. 