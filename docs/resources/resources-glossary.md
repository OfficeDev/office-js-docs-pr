---
title: Office Add-ins glossary of terms
description: 'A glossary of terms commonly used throughout the Office Add-ins documentation.'
ms.date: 01/27/2022
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

See also: client, host, Office application, Office client.

## application-specific API

Application-specific APIs provide strongly-typed objects that can be used to interact with objects that are native to a specific Office application. For example, you can use the Excel JavaScript APIs to access worksheets, ranges, tables, charts, and more. application-specific APIs are currently available for Excel, OneNote, PowerPoint, and Word.

See also: Common API.

## CDN

CDN is an acronym. It represents **content delivery network (CDN)** and refers to a distributed network of servers and data centers. A CDN typically provides higher resource availability and performance when compared to a single server or data center.

## client

In the Office Add-ins documentation, **client** typically refers to an Office application. The Office applications, or clients, that support Office Add-ins are: Excel, OneNote, Outlook, PowerPoint, Project, Visio, and Word.

See also: application, host, Office application, Office client.

## Common API

Common APIs can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications. This API model uses [callbacks](https://developer.mozilla.org/docs/Glossary/Callback_function), which allow you to specify only one operation in each request sent to the Office application. Common APIs were introduced with Office 2013 and can be used to interact with Office 2013 or later. For details about the Common API object model, which includes APIs for interacting with Outlook, PowerPoint, and Project, see [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).

See also: application-specific API.

## Contoso

## content add-in

**Content add-ins** are surfaces that can be embedded directly into Excel or PowerPoint documents. Content add-ins give users access to interface controls that run code to modify documents or display data from a data source. Use content add-ins when you want to embed functionality directly into the document.

## custom function

In the Office Add-ins documentation, a **custom function** is a user-defined function in Excel. Custom functions in Excel enable developers to add new functions, beyond the typical Excel features, by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel.

## host

In the Office Add-ins documentation, **host** typically refers to an Office application. The Office applications, or hosts, that support Office Add-ins are: Excel, OneNote, Outlook, PowerPoint, Project, Visio, and Word.

See also: application, client, Office application, Office client.

## Office application, Office client

In the Office Add-ins documentation, **Office client** refers to an Office application. The Office applications, or clients, that support Office Add-ins are: Excel, OneNote, Outlook, PowerPoint, Project, Visio, and Word.

See also: application, client, host.

## platform

In the Office Add-ins context, a **platform** usually refers to the operating system running an add-in. Platforms that support Office Add-ins are: Windows, Mac, iPad, and web browsers.

## requirement set

[!include[Requirement set note](../includes/office-js-requirement-sets.md)]

## ribbon, ribbon button

A **ribbon** is a command bar that organizes an application's features into a series of tabs or buttons at the top of a window. A **ribbon button** is one of the buttons within this series.

## runtime

In the Office Add-ins context, a **runtime** is a lifecycle, or the time during which an application is running.

## task pane

Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.

## tutorial

See also: quickstart.

## quickstart

See also: tutorial.

## UI-less custom function

A **UI-less custom function** is a custom functions add-in that doesn't have a task pane or other user-interface elements.

See also: custom function.

## Web add-ins

In the Office Add-ins documentation, **web add-in** is a legacy term for Office Add-ins.

See also: add-in.

## XLL

See also: custom function.

## Yeoman generator, yo office
