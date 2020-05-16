---
ms.date: 05/16/2020
description: 'Test your Office Add-in using Internet Explorer 11.'
title: Internet Explorer 11 testing
localization_priority: Normal
---

# Test your Office Add-in using Internet Explorer 11

Depending on the specifications of your add-in, you may plan to support older versions of Windows and Office, which require testing on Internet Explorer 11. This is often necessary as part of submitting your add-in to AppSource. You can use the following command line tooling to switch from more modern runtimes used by add-ins to the Internet Explorer 11 runtime for this testing.

## Pre-requisites

- [Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)
- A code editor. We recommend [Visual Studio Code](https://code.visualstudio.com/)
- [Be part of the Office Insider program](https://insider.office.com)

These instructions assume you have set up a Yo Office generator project before. If you haven't done this before, consider reading a quick start, such as [this one for Excel add-ins](../quickstarts/excel-quickstart-jquery.md).

## Using IE11 tooling

1. Create a Yo Office generator project. It doesn't matter what kind of project you select, this tooling will work with all project types.

> ![NOTE]
> If you have an existing project and want to add this tooling without creating a new project, skip this step and move to the next step. 

2. In the root folder of your new project, run the following in the command line:

```command&nbsp;line
office-add-dev-settings webview manifest.xml ie
```
You should see a note in the command line that the web view type is now set to IE.

> ![TIP]
> It isn't necessary to use this tooling, but it should help debug the majority of issues related to the Internet Explorer 11 runtime. For complete robustness, you should test using a computer with a copy of Windows 7 and Office 2013 installed.

## Command settings

Should you have a different manifest path, specify this in the command, as shown in the following:

`office-add-dev-settings webview [path to your manifest] ie`

The `office-addin-dev-settings webview` command can also take a number of runtimes as arguments:

- ie
- edge
- default

## See also
* [Test and debug Office Add-ins](test-debug-office-add-ins.md)
* [Sideload Office Add-ins for testing](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Debug add-ins using developer tools on Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [Attach a debugger from the task pane](attach-debugger-from-task-pane.md)