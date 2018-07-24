---
title: Sideload Office Add-ins Using Sideload Command
description: ''
ms.date: 07/24/2018
---

# Sideload Office Add-ins for testing using the **sideload command**
 (**NOTE**: this method only works for Excel, Word and PowerPoint add-ins).

1. Open a command prompt as an administrator, and change directories to the root of your add-in project folder.

2. Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"

3. Open a second command prompt as an administrator, and change directories to the root of your add-in project folder.

4. Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"

## See also

- [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md)
- [Publish your Office Add-in](../publish/publish.md)