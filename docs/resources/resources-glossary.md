---
title: Office Add-ins glossary of terms
description: ''
ms.date: 01/03/2022
ms.localizationpriority: medium
---

# Office Add-ins glossary of terms

This is a glossary of terms commonly used throughout the Office Add-ins documentation.

## add-in

Office Add-ins are web applications embedded in Office applications. These web applications add new functionality to the Office application, such as bringing in external data, automating processes, or embedding interactive objects in Office documents.

Office Add-ins differ from VBA, COM, and VSTO add-ins because they offer cross-platform support (web, Windows, Mac, and iPad) and are based on standard web technologies (HTML, CSS, and JavaScript). The primary programming language of an Office Add-in is JavaScript or TypeScript.

## add-in commands

## application

In the Office Add-ins documentation, **application** refers to an Office application. The Office applications that support Office Add-ins are: Excel, OneNote, Outlook, PowerPoint, Project, Visio, and Word.

See also: platform, host, Office application, Office client.

## application-specific API

Application-specific APIs provide strongly-typed objects that can be used to interact with objects that are native to a specific Office application. For example, you can use the Excel JavaScript APIs to access worksheets, ranges, tables, charts, and more. application-specific APIs are currently available for Excel, OneNote, PowerPoint, and Word.

See also: Common API.

## Common API

Common APIs can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications. This API model uses [callbacks](https://developer.mozilla.org/docs/Glossary/Callback_function), which allow you to specify only one operation in each request sent to the Office application. Common APIs were introduced with Office 2013 and can be used to interact with Office 2013 or later. For details about the Common API object model, which includes APIs for interacting with Outlook, PowerPoint, and Project, see [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).

See also: application-specific API.

## Contoso

## content pane

See also: task pane.

## custom function

## host

In the Office Add-ins documentation, **host** typically refers to an Office application. The Office applications, or hosts, that support Office Add-ins are: Excel, OneNote, Outlook, PowerPoint, Project, Visio, and Word.

See also: application, host, Office application, Office client.

## Office

## Office application, Office client

In the Office Add-ins documentation, **Office client** refers to an Office application. The Office applications, or clients, that support Office Add-ins are: Excel, OneNote, Outlook, PowerPoint, Project, Visio, and Word.

See also: application, platform, host.

## platform

In the the Office Add-ins context, a **platform** usually refers to the operating system running an add-in. Platforms that support Office Add-ins are: Windows, Mac, iPad, and web browsers.

## requirement set

[!include[Requirement set note](../includes/office-js-requirement-sets.md)]

See also: application-specific API.

## ribbon, ribbon button

## runtime

## task pane

See also: content pane.

## tutorial

See also: quickstart.

## quickstart

See also: tutorial.

## UI-less custom function

## XLL

## Yeoman generator, yo office
