---
title: Beginner's guide to Office Add-ins
description:  A recommended path for beginners through the learning resources for Office Add-ins.
ms.date: 11/10/2022
ms.topic: get-started
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Beginner's guide

Want to get started building your own cross-platform Office extensions? The following steps show you what to read first, what tools to install, and recommended tutorials to complete.

> [!NOTE]
> If you're experienced in creating VSTO add-ins for Office, we recommend that you immediately turn to [VSTO add-in developer's guide](learning-path-transition.md), which is a superset of the information in this article.

## Step 0: Prerequisites

- Office Add-ins are essentially web applications embedded in Office. So, you should first have a basic understanding of web applications and how they are hosted on the web. There is an enormous amount of information about this on the Internet, in books, and in online courses. A good way to start if you have no prior knowledge of web applications at all is to search for "What is a web app?" on Bing.
- The primary programming language you will use in creating Office Add-ins is JavaScript or TypeScript. You can think of TypeScript as a strongly-typed version of JavaScript. If you are not familiar with either of these languages, but you have experience with VBA, VB.Net, C#, you will probably find TypeScript easier to learn. Again, there is a wealth of information about these languages on the Internet, in books, and in online courses.

## Step 1: Begin with fundamentals

We know you're eager to start coding, but there are some things about Office Add-ins that you should read before you open your IDE or code editor.

- [Office Add-ins Platform Overview](office-add-ins.md): Find out what Office Web Add-ins are and how they differ from older ways of extending Office, such as VSTO add-ins.
- [Develop Office Add-ins](../develop/develop-overview.md): Get an overview of Office Add-in development and lifecycle including tooling, creating an add-in UI, and using the JavaScript APIs to interact with the Office document.
- ["Hello world" samples](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world): Learn how to build the simplest Office Add-in with only a manifest, HTML web page, and a logo. These samples will help you understand the fundamental parts of an Office Add-in.

There are a lot of links in those articles, but if you're a beginner with Office Add-ins, we recommend that you come back here when you've read them and continue with the next section.

## Step 2: Install tools and create your first add-in

You've got the big picture now, so dive in with one of our quick starts. For purposes of learning the platform, we recommend the Excel quick start. There is a version that is based on Visual Studio and a version that is based in Node.js and Visual Studio Code.

- [Visual Studio](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Node.js and Visual Studio Code](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## Step 3: Code

You can't learn to drive by reading the owner's manual, so start coding with this [Excel tutorial](../tutorials/excel-tutorial.md). You'll be using the Office JavaScript library and some XML in the add-in's manifest. There's no need to memorize anything, because you'll be getting more background about both in a later steps.

## Step 4: Understand the JavaScript library

First, get the big picture of the Office JavaScript library with this tutorial from Microsoft Learn training: [Understand the Office JavaScript APIs](/training/modules/understand-office-javascript-apis/index).

Then explore the Office JavaScript APIs with our [the Script Lab tool](explore-with-script-lab.md) -- a sandbox for running and exploring the APIs.

## Step 5: Understand the manifest

Get an understanding of the purposes of the add-in manifest and an introduction to its XML markup in [Office Add-ins XML manifest](../develop/add-in-manifests.md).

## Next Steps

Congratulations on finishing the beginner's learning path for Office Add-ins! Here are some suggestions for further exploration of our documentation:

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
  - [Learn about the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)
