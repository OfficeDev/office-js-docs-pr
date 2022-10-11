---
title: Explore Office JavaScript API using Script Lab
description: Use Script Lab to explore the Office JS API and to prototype functionality.
ms.date: 09/27/2022
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Explore Office JavaScript API using Script Lab

The [Script Lab](https://appsource.microsoft.com/product/office/WA104380862) and [Script Lab for Outlook](https://appsource.microsoft.com/product/office/WA200001603) add-ins, available free from AppSource, enable you to explore the Office JavaScript API while you're working in an Office program such as Excel or Outlook. Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your own add-in.

## What is Script Lab?

Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Outlook, Word, and PowerPoint. It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code. Through Script Lab, you can access a library of samples to quickly try out features or you can use a sample as the starting point for your own code. You can even use Script Lab to try preview APIs.

Sounds good so far? Take a look at this one-minute video to see Script Lab in action.

[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)

## Key features

Script Lab offers a number of features to help you explore the Office JavaScript API and prototype add-in functionality.

### Explore samples

Get started quickly with a collection of built-in sample snippets that show how to complete tasks with the API. You can run the samples to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.

![Samples.](../images/script-lab-samples.jpg)

### Code and style

In addition to JavaScript or TypeScript code that calls the Office JS API, each snippet also contains HTML markup that defines content of the task pane and CSS that defines the appearance of the task pane. You can customize the HTML markup and CSS to experiment with element placement and styling as you prototype task pane design for your own add-in.

> [!TIP]
> To call preview APIs within a snippet, you'll need to update the snippet's libraries to use the beta content delivery network (CDN) (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) and the preview type definitions `@types/office-js-preview`. Additionally, some preview APIs are only accessible if you've signed up for the [Office Insider program](https://insider.office.com) and are running an Insider build of Office.

### Save and share snippets

By default, snippets that you open in Script Lab will be saved to your browser cache. To save a snippet permanently, you can export it to a [GitHub gist](https://gist.github.com). Create a secret gist to save a snippet exclusively for your own use, or create a public gist if you plan to share it with others.

![Sharing options.](../images/script-lab-share.jpg)

### Import snippets

You can import a snippet into Script Lab either by specifying the URL to the public [GitHub gist](https://gist.github.com) where the snippet YAML is stored or by pasting in the complete YAML for the snippet. This feature may be useful in scenarios where someone else has shared their snippet with you by either publishing it to a GitHub gist or providing their snippet's YAML.

![Import snippet option.](../images/script-lab-import-snippet.jpg)

## Supported clients

Script Lab is supported for Excel, Word, and PowerPoint on the following clients.

- Office on Windows\*
- Office 2016 or later on Mac
- Office on the web

Script Lab for Outlook is available on the following clients.

- Outlook on Windows\*
- Outlook 2016 or later on Mac
- Outlook on the web when using Chrome, Microsoft Edge, or Safari browsers

For more details on Script Lab for Outlook, see the related [blog post](https://devblogs.microsoft.com/microsoft365dev/script-lab-now-supports-outlook/).

> [!IMPORTANT]
> \* Script Lab no longer works with combinations of platform and Office version that use Internet Explorer to host add-ins. This includes perpetual versions of Office through Office 2019. For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

## Next steps

To use Script Lab in Excel, Word, or PowerPoint, install the [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862) from AppSource.

To use Script Lab for Outlook, install the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) from AppSource.

You're welcome to expand the sample library in Script Lab by contributing new snippets to the [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub repository.

When you're ready to create your first Office Add-in, try out the quick start for [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](../quickstarts/outlook-quickstart.md), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md), or [Project](../quickstarts/project-quickstart.md).

## See also

- [Get Script Lab for Excel, Word, or Powerpoint](https://appsource.microsoft.com/product/office/WA104380862)
- [Get Script Lab for Outlook](https://appsource.microsoft.com/product/office/wa200001603)
- [Learn more about Script Lab](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [Join the Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program)
- [Developing Office Add-ins](../develop/develop-overview.md)
- [Learn about the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)
