---
title: Build Excel add-ins with Office Add-ins
description: Learn what Excel add-ins are, what you can build, and where to start with the Excel JavaScript API, workbook automation, and custom functions.
ms.date: 06/03/2026
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ai-usage: ai-assisted
---

# Build Excel add-ins for Excel

Use an Excel add-in when you want to automate workbook tasks, connect workbook data to external services, add custom calculations, or guide users with a web-based experience in Excel. Excel add-ins run in Excel on the web, Windows, Mac, and iPad, so you can build one solution for multiple platforms.

With the Office Add-ins platform and Office.js, you can:

- Read and write workbook data, including worksheets, ranges, tables, charts, and named items.
- Extend the **Ribbon** and **context menu** or add a task pane or content pane with web-based UI.
- Add custom functions that users call from worksheet cells.
- Open dialog boxes for sign-in, confirmation, and other focused tasks.

The platform also supports centralized deployment, standard web technologies such as HTML, CSS, and JavaScript, and publishing through the [Microsoft Marketplace](https://marketplace.microsoft.com/marketplace/apps?product=office).

> [!NOTE]
> Excel add-ins are different from COM and VSTO add-ins, which run only in Office on Windows. Excel add-ins don't require you to install code on a user's device or in Excel.

## Start with the most common Excel add-in tasks

If you're new to Excel add-ins, start with the [Excel quickstart](../quickstarts/excel-quickstart-jquery.md). Then use these articles to go deeper into the Excel object model and common workbook scenarios:

- [Learn the Excel JavaScript object model](excel-add-ins-core-concepts.md).
- [Work with worksheets](excel-add-ins-worksheets.md).
- [Work with tables](excel-add-ins-tables.md).
- [Work with charts](excel-add-ins-charts.md).
- [Create custom functions](custom-functions-overview.md).

## What you can build with an Excel add-in

In addition to working with workbook content, Excel add-ins can add commands, show task panes, define custom functions, open dialog boxes, and embed rich web content in a worksheet.

### Add-in commands

Add-in commands extend the Excel UI and start actions in your add-in. You can add a button to the **ribbon** or an item to a **context menu**. When users select a command, they can run JavaScript code or open a page from the add-in in a **task pane**.

:::image type="content" source="../images/excel-add-in-commands-script-lab.png" alt-text="Add-in commands in Excel.":::

For information about supported platforms and design guidance, see [Add-in commands for Excel, Word, and PowerPoint](../design/add-in-commands.md).

### Task panes

The **task pane** interface appears on the right side of the Excel window. It gives users controls that run code, update the workbook, or show data from another source.

:::image type="content" source="../images/excel-add-in-task-pane-insights.png" alt-text="Task pane add-in in Excel.":::

For more information, see [Task panes in Office Add-ins](../design/task-pane-add-ins.md). For a working sample, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends).

### Custom functions

Custom functions let you define new worksheet functions in JavaScript as part of an add-in. Users can call them the same way they call built-in functions such as `SUM()`.

:::image type="content" source="../images/SphereVolumeNew.gif" alt-text="Animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet.":::

For more information, see [Create custom functions in Excel](custom-functions-overview.md).

### Dialog boxes

Dialog boxes float above the active Excel window. Use them for sign-in flows that can't open directly in a **task pane**, to confirm an action, or to host content that needs more focused space. To open a dialog box in your add-in, use the [Dialog API](/javascript/api/office/office.ui).

:::image type="content" source="../images/excel-add-in-dialog-choose-number.png" alt-text="Add-in dialog box in Excel.":::

To learn more, see [Use the Dialog API in your Office Add-ins](../develop/dialog-api-in-office-add-ins.md).

### Content add-ins

A content add-in is embedded directly in a worksheet. Use one when you want users to interact with a chart, data visualization, media experience, or other web content in the document itself.

:::image type="content" source="../images/excel-add-in-content-map.png" alt-text="Content add-in in Excel.":::

To learn more, see [Content Office Add-ins](../design/content-add-ins.md). For a working sample, see [Excel content add-in: Humongous Insurance](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-content-add-in).

## What an Excel add-in includes

An Excel add-in has two basic parts: a web application and a manifest.

The web application uses the [Office JavaScript API](../reference/javascript-api-for-office.md) to work with objects in Excel and connect to online resources. For example, the web app can:

- Create, read, update, and delete workbook data.
- Authenticate users with an online service by using OAuth 2.0.
- Send requests to Microsoft Graph or another web API.

You can host the web app on any web server and build it with client-side frameworks such as Angular, React, or jQuery, or with server-side technologies such as ASP.NET, Node.js, or PHP.

The [manifest](../develop/add-in-manifests.md) is a configuration file that defines how the add-in integrates with Office. It specifies settings and capabilities such as:

- The URL of the add-in's web application.
- The add-in's display name, description, ID, version, and default locale.
- How the add-in integrates with Excel, including custom UI such as **Ribbon** buttons and **context menu** items.
- The permissions that the add-in requires, such as reading or writing document data.

To make an Excel add-in available to users, publish its manifest to Microsoft Marketplace or to an add-ins catalog. For details about Marketplace publishing, see [Make your solutions available in Microsoft Marketplace and within Office](/partner-center/marketplace-offers/submit-to-appsource-via-partner-center).

## JavaScript APIs for workbook content

[!include[The roles of the Common and application-specific APIs](../includes/excel-api-models.md)]

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Develop Office Add-ins](../develop/develop-overview.md)
- [Excel JavaScript API overview](../reference/overview/excel-add-ins-reference-overview.md)
- [Join the Microsoft 365 Developer Program](/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-)
