---
title: Start Here! A guide for beginners making Office Add-ins
description:  A recommended path for beginners through the learning resources for Office Add-ins.
ms.date: 02/19/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
---

# Start Here! A guide for beginners making Office Add-ins

Thank your for your interest in Office Add-ins, the Office extensions that run cross-platform, including Office on the web, Windows, and Mac. Like many people who are exploring a new framework, your first question might be **"Where do I start?"**. So, we created this recommended sequence of resources.

## 1 Begin with fundamentals

We know you're eager to start coding, but there are some things about Office Add-ins that you should read before you open your IDE or code editor.

- [Office Add-ins Platform Overview](office-add-ins.md): Find out what Office Web Add-ins are and how they differ from older ways of extending Office, such as VSTO add-ins. (Spoiler alert: they're web apps embedded in Office and they run in Office on Mac as well as Windows.)
- [Building Office Add-ins](office-add-ins-fundamentals.md): Get the 1000-meter view of Office add-in development and lifecycle including tools, creating an add-in UI, and using the JavaScript APIs to interact with the Office document.

There are a lot of links in those articles, but if you're a beginner with Office Add-ins, we recommend that you come back here when you've read them and continue with the next section.

## 2 Install tools and create your first add-in

You've got the big picture now, but you need tools before you can code. Use one of our quick starts. For purposes of learning the platform, we recommend the Excel quick start. For each we have a version that is based on Visual Studio and a version that is based in Node.js and Visual Studio Code:

- [Visual Studio](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Node.js and Visual Studio Code](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## 3 Code

You can't learn to drive by reading the owner's manual, so start coding with this [Excel tutorial](../tutorials/excel-tutorial.md). You'll be using the Office JavaScript library, and some XML in the add-in's manifest, but don't try to memorize anything. You'll be getting more background about both in a later steps.

## 4 Understand the JavaScript library

First, get the big picture of the Office JavaScript library with this tutorial from Microsoft Learn: [Understand the Office JavaScript APIs](/learn/modules/understand-office-javascript-apis/index).

Then explore the Office JavaScript APIs using [the Script Lab tool](explore-with-script-lab.md).

## 5 Understand the manifest

Get an understanding of the purposes of the add-in manifest and an introduction to its XML markup in [Office Add-ins XML manifest](../develop/add-in-manifests.md).

## Next Steps

You're not a beginner anymore. Here are some suggestions for further exploration of our documentation:

- Tutorials or quick starts for other Office applications:

  - [OneNote quick start](../quickstarts/onenote-quickstart.md)
  - [Outlook tutorial](/outlook/add-ins/addin-tutorial)
  - [PowerPoint tutorial](../tutorials/powerpoint-tutorial.md)
  - [Project quick start](../quickstarts/project-quickstart.md)
  - [Word tutorial](../tutorials/word-tutorial.md)

- Other important subjects:

  - [Develop Office Add-ins](../develop/develop-overview.md)
  - [Best practices for developing Office Add-ins](../concepts/add-in-development-best-practices.md)
  - [Design Office Add-ins](../design/add-in-design.md)
  - [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
  - [Deploy and publish Office Add-ins](../publish/publish.md)
  - [Resources](../resources/resources-links-help.md)
