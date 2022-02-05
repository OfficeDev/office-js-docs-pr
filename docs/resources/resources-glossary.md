---
title: Office Add-ins glossary of terms
description: A glossary of terms commonly used throughout the Office Add-ins documentation.
ms.date: 02/04/2022
ms.localizationpriority: medium
---

# Office Add-ins glossary of terms

This is a glossary of terms commonly used throughout the Office Add-ins documentation.

## add-in

Office Add-ins are web applications embedded in Office applications. These web applications add new functionality to the Office application, such as bringing in external data, automating processes, or embedding interactive objects in Office documents.

Office Add-ins differ from VBA, COM, and VSTO add-ins because they offer cross-platform support (web, Windows, Mac, and iPad) and are based on standard web technologies (HTML, CSS, and JavaScript). The primary programming language of an Office Add-in is JavaScript or TypeScript.

## add-in commands

**Add-in commands** are UI elements, such as buttons and menus, that extend the Office UI for your add-in. When users select an add-in command element, they initiate actions such as running JavaScript code or displaying the add-in in a task pane. Add-in commands help users find and use your add-in, which can support your add-in's adoption and reuse. See [Add-in commands for Excel, PowerPoint, and Word](../design/add-in-commands.md) to learn more.

See also: [ribbon, ribbon button](#ribbon-ribbon-button).

## application

In the Office Add-ins documentation, **application** refers to an Office application. The Office applications that support Office Add-ins are: Excel, OneNote, Outlook, PowerPoint, Project, Visio, and Word.

See also: [client](#client), [host](#host), [Office application, Office client](#office-application-office-client).

## application-specific API

Application-specific APIs provide strongly-typed objects that can be used to interact with objects that are native to a specific Office application. For example, you can use the Excel JavaScript APIs to access worksheets, ranges, tables, charts, and more. Application-specific APIs are currently available for Excel, OneNote, PowerPoint, and Word. See [Application-specific API model](../develop/application-specific-api-model.md) to learn more.

See also: [Common API](#common-api).

## CDN

CDN is an acronym. It represents **content delivery network (CDN)** and refers to a distributed network of servers and data centers. A CDN typically provides higher resource availability and performance when compared to a single server or data center.

## client

In the Office Add-ins documentation, **client** typically refers to an Office application. The Office applications, or clients, that support Office Add-ins are: Excel, OneNote, Outlook, PowerPoint, Project, Visio, and Word.

See also: [application](#application), [host](#host), [Office application, Office client](#office-application-office-client).

## Common API

Common APIs can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications. This API model uses [callbacks](https://developer.mozilla.org/docs/Glossary/Callback_function), which allow you to specify only one operation in each request sent to the Office application. Common APIs were introduced with Office 2013 and can be used to interact with Office 2013 or later. For details about the Common API object model, which includes APIs for interacting with Outlook, PowerPoint, and Project, see [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).

See also: [application-specific API](#application-specific-api).

## Contoso

**Contoso** Ltd. (also known as Contoso and Contoso University) is a fictional company used by Microsoft as an example company and domain.

## content add-in

**Content add-ins** are surfaces that can be embedded directly into Excel or PowerPoint documents. Content add-ins give users access to interface controls that run code to modify documents or display data from a data source. Use content add-ins when you want to embed functionality directly into the document. See [Content Office Add-ins](../design/content-add-ins.md) to learn more.

## custom function

In the Office Add-ins documentation, a **custom function** is a user-defined function in Excel. Custom functions in Excel enable developers to add new functions, beyond the typical Excel features, by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel. See [Create custom functions in Excel](../excel/custom-functions-overview.md) to learn more.

## host

In the Office Add-ins documentation, **host** typically refers to an Office application. The Office applications, or hosts, that support Office Add-ins are: Excel, OneNote, Outlook, PowerPoint, Project, Visio, and Word.

See also: [application](#application), [client](#client), [Office application, Office client](#office-application-office-client).

## Office application, Office client

In the Office Add-ins documentation, **Office client** refers to an Office application. The Office applications, or clients, that support Office Add-ins are: Excel, OneNote, Outlook, PowerPoint, Project, Visio, and Word.

See also: [application](#application), [client](#client), [host](#host).

## platform

In the Office Add-ins context, a **platform** usually refers to the operating system running an add-in. Platforms that support Office Add-ins are: Windows, Mac, iPad, and web browsers.

## requirement set

[!include[Requirement set note](../includes/office-js-requirement-sets.md)]

## ribbon, ribbon button

A **ribbon** is a command bar that organizes an application's features into a series of tabs or buttons at the top of a window. A **ribbon button** is one of the buttons within this series.

## runtime

In the Office Add-ins context, a **runtime** is a lifecycle, or the time during which an application is running.

## task pane

Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document. See [Task panes in Office Add-ins](../design/task-pane-add-ins.md) to learn more.

## tutorial

A **tutorial** is a teaching aid designed to help people learn to use a product or procedure. In the Office Add-ins context, a tutorial guides an add-in developer through the complete add-in development process for a particular application, such as Excel. This involves following 20 or more steps and is a greater time investment than a [quickstart](#quickstart).

See also: [quickstart](#quickstart).

## quickstart

A **quickstart** is a a high level description of key skills and knowledge required for the basic operation of a particular program. In the Office Add-ins documentation, a quickstart is an introduction to developing an add-in for a particular application, such as Outlook. A quickstart contains series of steps that an add-in developer can complete in approximately 5 minutes, resulting in a functioning add-in.

See also: [tutorial](#tutorial).

## UI-less custom function

A **UI-less custom function** is a custom functions add-in that doesn't have a task pane or other user-interface elements.

See also: [custom function](#custom-function).

## web add-in

In the Office Add-ins documentation, **web add-in** is a legacy term for Office Add-ins.

See also: [add-in](#add-in).

## XLL

An **XLL** add-in is an Excel add-in file with the file extension **.xll**. An XLL file is a type of dynamic link library (DLL) file that can only be opened by Excel. XLL add-in files must be written in C or C++. See [Developing Excel XLLs](/office/client-developer/excel/developing-excel-xlls) to learn more.

See also: [custom function](#custom-function).

## Yeoman generator, yo office

The [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) uses the open source [Yeoman](https://github.com/yeoman/yo) tool to generate an Office Add-in via the command line. **Yo office** is the command line argument that runs the Yeoman generator for Office Add-ins. The Office Add-ins quickstarts and tutorials use the Yeoman generator.

## See also

- [Office Add-ins additional resources](resources-links-help.md)
