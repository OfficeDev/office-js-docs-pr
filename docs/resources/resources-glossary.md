---
title: Office Add-ins glossary of terms
description: A glossary of terms commonly used throughout the Office Add-ins documentation.
ms.date: 02/09/2022
ms.localizationpriority: medium
---

# Office Add-ins glossary

This is a glossary of terms commonly used throughout the Office Add-ins documentation.

## add-in

Office Add-ins are web applications that extend Office applications. These web applications add new functionality to the Office application, such as bring in external data, automate processes, or embed interactive objects in Office documents.

Office Add-ins differ from VBA, COM, and VSTO add-ins because they offer cross-platform support (web, Windows, Mac, and iPad) and are based on standard web technologies (HTML, CSS, and JavaScript). The primary programming language of an Office Add-in is JavaScript or TypeScript.

## add-in commands

**Add-in commands** are UI elements, such as buttons and menus, that extend the Office UI for your add-in. When users select an add-in command element, they initiate actions such as running JavaScript code or displaying the add-in in a task pane. Add-in commands let your add-in look and feel like a part of Office, which gives users more confidence in your add-in. See [Add-in commands for Excel, PowerPoint, and Word](../design/add-in-commands.md) to learn more.

See also: [ribbon, ribbon button](#ribbon-ribbon-button).

## application

**Application** refers to an Office application. The Office applications that support Office Add-ins are Excel, OneNote, Outlook, PowerPoint, Project, Visio, and Word.

See also: [client](#client), [host](#host), [Office application, Office client](#office-application-office-client).

## application-specific API

Application-specific APIs provide strongly-typed objects that interact with objects that are native to a specific Office application. For example, you call the Excel JavaScript APIs for access to worksheets, ranges, tables, charts, and more. Application-specific APIs are currently available for Excel, OneNote, PowerPoint, and Word. See [Application-specific API model](../develop/application-specific-api-model.md) to learn more.

See also: [Common API](#common-api).

## client

**Client** typically refers to an Office application. The Office applications, or clients, that support Office Add-ins are Excel, OneNote, Outlook, PowerPoint, Project, Visio, and Word.

See also: [application](#application), [host](#host), [Office application, Office client](#office-application-office-client).

## Common API

Common APIs can be used to access features such as UI, dialogs, and client settings that are common across multiple Office applications. This API model uses [callbacks](https://developer.mozilla.org/docs/Glossary/Callback_function), which allow you to specify only one operation in each request sent to the Office application. Common APIs were introduced with Office 2013 and can be used to interact with Office 2013 or later. For details about the Common API object model, which includes APIs for interacting with Outlook, PowerPoint, and Project, see [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).

See also: [application-specific API](#application-specific-api).

## content add-in

**Content add-ins** are surfaces that can be embedded directly into Excel, PowerPoint, or OneNote documents. Content add-ins give users access to interface controls that run code to modify documents or display data from a data source. Use content add-ins when you want to embed functionality directly into the document. See [Content Office Add-ins](../design/content-add-ins.md) to learn more.

## content delivery network (CDN)

A **content delivery network** or **CDN** is a distributed network of servers and data centers. It typically provides higher resource availability and performance when compared to a single server or data center.

## Contoso

**Contoso** Ltd. (also known as Contoso and Contoso University) is a fictional company used by Microsoft as an example company and domain.

## custom function

A **custom function** is a user-defined function in Excel. Custom functions in Excel enable developers to add new functions, beyond the typical Excel features, by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel. See [Create custom functions in Excel](../excel/custom-functions-overview.md) to learn more.

## host

**Host** typically refers to an Office application. The Office applications, or hosts, that support Office Add-ins are Excel, OneNote, Outlook, PowerPoint, Project, Visio, and Word.

See also: [application](#application), [client](#client), [Office application, Office client](#office-application-office-client).

## Office application, Office client

**Office client** refers to an Office application. The Office applications, or clients, that support Office Add-ins are Excel, OneNote, Outlook, PowerPoint, Project, Visio, and Word.

See also: [application](#application), [client](#client), [host](#host).

## platform

A **platform** usually refers to the operating system running the Office application. Platforms that support Office Add-ins are Windows, Mac, iPad, and web browsers.

## quick start

A **quick start** is a high-level description of key skills and knowledge required for the basic operation of a particular program. In the Office Add-ins documentation, a quick start is an introduction to developing an add-in for a particular application, such as Outlook. A quick start contains series of steps that an add-in developer can complete in approximately 5 minutes, resulting in a functioning add-in and functional development environment.

See also: [tutorial](#tutorial).

## requirement set

[!include[Requirement set note](../includes/office-js-requirement-sets.md)]

## ribbon, ribbon button

A **ribbon** is a command bar that organizes an application's features into a series of tabs or buttons at the top of a window. A **ribbon button** is one of the buttons within this series. See [Show or hide the ribbon in Office](https://support.microsoft.com/office/show-or-hide-the-ribbon-in-office-d946b26e-0c8c-402d-a0f7-c6efa296b527#ID0EBBD=Newer_Versions) for more information.

## runtime

A **runtime** is a lifecycle, or the time during which an application is running.

## task pane

Task panes are interface surfaces that typically appear on the right side of the window within Excel, Outlook, PowerPoint, and Word. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document. See [Task panes in Office Add-ins](../design/task-pane-add-ins.md) to learn more.

## tutorial

A **tutorial** is a teaching aid designed to help people learn to use a product or procedure. In the Office Add-ins context, a tutorial guides an add-in developer through the complete add-in development process for a particular application, such as Excel. This involves following 20 or more steps and is a greater time investment than a [quick start](#quick-start).

See also: [quick start](#quick-start).

## UI-less custom function

A **UI-less custom function** is a custom functions add-in that doesn't have a task pane or other user-interface elements.

See also: [custom function](#custom-function).

## web add-in

A **web add-in** is a legacy term for Office Add-ins. This term may be used when the Microsoft 365 documentation needs to distinguish modern Office Add-ins from other types of add-ins like VBA, COM, or VSTO.

See also: [add-in](#add-in).

## XLL

An **XLL** add-in is an Excel add-in file with the file extension **.xll**. An XLL file is a type of dynamic link library (DLL) file that can only be opened by Excel. XLL add-in files must be written in C or C++. See [Developing Excel XLLs](/office/client-developer/excel/developing-excel-xlls) to learn more.

See also: [custom function](#custom-function).

## Yeoman generator, yo office

The [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) uses the open source [Yeoman](https://github.com/yeoman/yo) tool to generate an Office Add-in via the command line. `yo office` is the command that runs the Yeoman generator for Office Add-ins. The Office Add-ins quick starts and tutorials use the Yeoman generator.

## See also

- [Office Add-ins additional resources](resources-links-help.md)
