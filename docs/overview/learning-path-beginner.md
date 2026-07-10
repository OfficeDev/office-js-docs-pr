---
title: Beginner's guide to Office Add-ins
description: A recommended learning path for beginners to build their first cross-platform Office Add-in.
ms.date: 07/09/2026
ms.topic: get-started
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Beginner's guide to Office Add-ins

Want to build your own cross-platform Office Add-in but not sure where to start? This guide gives you a step-by-step path: what to read first, which tools to install, and the tutorials to complete to go from zero to your first working add-in.

> [!NOTE]
> If you're experienced in creating VSTO add-ins for Office, we recommend that you immediately turn to [VSTO add-in developer's guide](learning-path-transition.md), which is a superset of the information in this article.

## Step 0: Prerequisites

- **Web development basics.** Office Add-ins are essentially web applications embedded in Office. You should first have a basic understanding of web applications and how they're hosted on the web. There's an enormous amount of information about this on the internet, in books, and in online courses. If you have no prior knowledge of web applications at all, a good way to start is to search for "What is a web app?".
- **JavaScript or TypeScript.** The primary programming language you use to create Office Add-ins is JavaScript or TypeScript. You can think of TypeScript as a strongly typed version of JavaScript. If you aren't familiar with either language but you have experience with VBA, VB.NET, or C#, you'll probably find TypeScript easier to learn. Again, there's a wealth of information about these languages on the internet, in books, and in online courses.

## Step 1: Begin with fundamentals

We know you're eager to start coding, but there are some things about Office Add-ins that you should read before you open your IDE or code editor.

- [Office Add-ins platform overview](office-add-ins.md): Find out what Office Web Add-ins are and how they differ from older ways of extending Office, such as VSTO add-ins.
- [Develop Office Add-ins](../develop/develop-overview.md): Get an overview of Office Add-in development and lifecycle including tooling, creating an add-in UI, and using the JavaScript APIs to interact with the Office document.
- ["Hello world" samples](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world): Learn how to build the simplest Office Add-in with only a manifest, HTML web page, and a logo. These samples will help you understand the fundamental parts of an Office Add-in.

There are a lot of links in those articles, but if you're a beginner with Office Add-ins, we recommend that you come back here when you've read them and continue with the next section.

## Step 2: Explore and try out existing samples

You've got the big picture now, so dive in by installing our [Script Lab add-in](explore-with-script-lab.md) to try out code samples in the various Office applications. The samples available in Script Lab show how to use many of the Office JavaScript APIs.

## Step 3: Install tools and create your first add-in

Next, create an add-in using one of our quick starts. For the purpose of learning the platform, we recommend the [Excel quick start](../quickstarts/excel-quickstart-jquery.md).

## Step 4: Code

You can't learn to drive by reading the owner's manual, so start coding with this [Excel tutorial](../tutorials/excel-tutorial.md). You'll use the Office JavaScript library and some JSON or XML in the add-in's manifest. There's no need to memorize anything, because you'll get more background about both in later steps.

## Step 5: Understand the JavaScript library

For an overview of the Office JavaScript library, see [Develop Office Add-ins](../develop/develop-overview.md).

Then return to Script Lab and use it like a playground: make your own code changes to the local copy of any samples you try and see how the results are affected.

## Step 6: Understand the manifest

Get an understanding of the purposes of the add-in manifest and an introduction to its XML markup or JSON in [Office Add-ins manifest](../develop/add-in-manifests.md).

## Step 7: Create a Partner Center account

If you plan to [publish your add-in to Microsoft Marketplace](../publish/publish.md), create a [Partner Center account](/partner-center/marketplace-offers/open-a-developer-account). This could take some time. Get this process going as soon as possible to avoid release delays.

## Next steps

Congratulations on finishing the beginner's learning path for Office Add-ins! Here are some suggestions for further exploration of our documentation:

- Tutorials or quick starts for other Office applications:

  - [OneNote quick start](../quickstarts/onenote-quickstart.md)
  - [Outlook tutorial](/outlook/add-ins/addin-tutorial)
  - [PowerPoint tutorial](../tutorials/powerpoint-tutorial-yo.md)
  - [Project quick start](../quickstarts/project-quickstart.md)
  - [Word tutorial](../tutorials/word-tutorial.md)

- Scenarios and other code samples:

  - [Excel: Create a spreadsheet from your web page and embed your add-in](../excel/pnp-open-in-excel.md)
  - [Outlook: Report spam or phishing emails](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-spam-reporting)
  - [Word: Import a document template](../word/import-template.md)
  - [Word: Manage citations](../word/citation-management.md)
  - [Office Add-in code samples](office-add-in-code-samples.md)

- Other important subjects:

  - [Develop Office Add-ins](../develop/develop-overview.md)
  - [Best practices for developing Office Add-ins](../concepts/add-in-development-best-practices.md)
  - [Design Office Add-ins](../design/add-in-design.md)
  - [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
  - [Deploy and publish Office Add-ins](../publish/publish.md)
  - [Resources](../resources/resources-links-help.md)
  - [Learn about the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)
