---
title: Test and debug Office Add-ins on a non-local server
description: Learn how to sideload, test, and debug an Office Add-in hosted on a staging server or cloud account.
ms.date: 08/12/2026
ms.localizationpriority: medium
---

# Test and debug Office Add-ins on a non-local server

When you've completed testing on a localhost, and are ready to test an Office Add-in from a staging server or cloud account, use the [office-addin-debugging](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-debugging) tool with any Node.js-based add-in project. The tool isn't supported in projects created with Visual Studio.

> [!NOTE]
> If you're working on a Windows computer, you may have another option for non-local testing. See [Sideload Office Add-ins for testing from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

## Projects created with Microsoft 365 Agents Toolkit or the Office Yeoman Generator (Yo Office)

If your project was created with [Microsoft 365 Agents Toolkit](../develop/agents-toolkit-overview.md) or the [Yeoman Generator for Office Add-ins](../develop/yeoman-generator-overview.md), `office-addin-debugging` is already installed. The project's `package.json` file includes `start` and `stop` scripts that invoke the tool. For non-local testing, update the domain in the manifest URLs to point to your staging server or content delivery network (CDN). Then run `npm run start` from a terminal to sideload the add-in for testing and debugging.

> [!IMPORTANT]
> The `office-addin-debugging` tool registers the add-in in the Windows registry or a special folder on a Mac. For an Outlook add-in, it also registers the add-in in Exchange. To avoid subtle bugs during development, always end a testing session by running `npm run stop`. This command removes the registrations and fully stops the server process. *Manually closing the server, terminal, Visual Studio Code, or the Office application doesn't remove these registrations.*

## Other projects

If your project wasn't created with Agents Toolkit or Yo Office, run the tool with `npx` in the project root. Invoke it with its `start` command followed by the relative path to the manifest. The following example starts the tool with a manifest in the project root.

```command&nbsp;line
npx office-addin-debugging start manifest.json
```

This command sideloads the add-in for testing and debugging. The tool also works with an add-in only manifest.

For details about the available `start` options, see the [office-addin-debugging README](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-debugging).

> [!IMPORTANT]
> The `office-addin-debugging` tool registers the add-in in the Windows registry or a special folder on a Mac. For an Outlook add-in, it also registers the add-in in Exchange. To avoid subtle bugs during development, always end a testing session by running `npx office-addin-debugging stop`. This command removes the registrations and fully stops the server process. *Manually closing the server, terminal, Visual Studio Code, or the Office application doesn't remove these registrations.* If you used the `--prod` option with the `start` command, use the same option with the `stop` command.
