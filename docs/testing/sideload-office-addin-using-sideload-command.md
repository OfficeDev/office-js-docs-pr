---
title: Sideload Office Add-ins using the sideload command
description: ''
ms.date: 07/24/2018
---

# Sideload Office Add-ins for testing using the **sideload command**
 >[!NOTE]
>The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins).

1. Open a command prompt as an administrator.

2. Change directories to the root of your add-in project folder.

3. Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"

4. Open a second command prompt as an administrator.

5. Change directories to the root of your add-in project folder.

6. Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"

## See also

- [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md)
- [Publish your Office Add-in](../publish/publish.md)