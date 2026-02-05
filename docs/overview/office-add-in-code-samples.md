---
title: Office Add-in code samples
description: A list of Office Add-in code samples to help you learn and build your own add-ins.
ms.date: 02/05/2026
ms.localizationpriority: high
---

# Office Add-in code samples

These code samples are written to help you learn how to use various features when developing Office Add-ins.

## Getting started

The following samples show how to build the simplest Office Add-in with only a manifest, HTML web page, and a logo. These components are the fundamental parts of an Office Add-in. For additional getting started information, see our [quick starts](../quickstarts/excel-quickstart-jquery.md) and [tutorials](/search/?terms=tutorial&scope=Office%20Add-ins).

### Task pane add-ins

- [Excel "Hello world" add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/excel-hello-world)
- [Outlook "Hello world" add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/outlook-hello-world)
- [PowerPoint "Hello world" add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/powerpoint-hello-world)
- [Word "Hello world" add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/word-hello-world)

### Content add-ins

- [Excel "Hello world" content add-in](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/hello-world/excel-content-hello-world)
- [PowerPoint "Hello world" content add-in](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/hello-world/powerpoint-content-hello-world)

## Blazor WebAssembly

If your development background is in building VSTO Add-ins, the following samples show how to build Office Web Add-ins using .NET Blazor WebAssembly. You can keep much of your code in C# and Visual Studio.

- [Create a Blazor WebAssembly Excel add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/blazor-add-in/excel-blazor-add-in)
- [Create a Blazor WebAssembly Outlook add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/blazor-add-in/outlook-blazor-add-in)
- [Create a Blazor WebAssembly Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/blazor-add-in/word-blazor-add-in)

## Excel

| Name                | Description         |
|:--------------------|:--------------------|
| [Data types explorer](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-data-types-explorer) (preview) | Builds an Excel add-in that allows you to create and explore data types in your workbooks. Data types enable add-in developers to organize complex data structures as objects, such as formatted number values, web images, and entity values. |
| [Open in Teams](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-open-in-teams) | Create a new Excel spreadsheet in Microsoft Teams containing data you define.|
| [Insert an external Excel file and populate it with JSON data](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-insert-file)  | Insert an existing template from an external Excel file into the currently open Excel workbook. Then, populate the template with data from a JSON web service. |
| [Create custom contextual tabs on the ribbon](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-contextual-tabs) | Create a custom contextual tab on the ribbon in the Office UI. The sample creates a table, and when the user moves the focus inside the table, the custom tab is displayed. When the user moves outside the table, the custom tab is hidden. |
| [Custom function sample using web worker](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/web-worker) | Use web workers in custom functions to prevent blocking the UI of your Office Add-in. |
| [Use storage techniques to access data from an Office Add-in when offline](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin) | Implement localStorage to enable limited functionality for your Office Add-in when a user experiences lost connection. |
| [Custom function batching pattern](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/Batching)| Batch multiple calls into a single call to reduce the number of network calls to a remote service.|

## Outlook

| Name                | Description         |
|:--------------------|:--------------------|
| [Report spam or phishing emails in Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-spam-reporting) | Build an integrated spam-reporting solution that's easily discoverable in the Outlook client ribbon. The solution provides the user with a dialog to report an email. It also saves a copy of the reported email to a file for further processing in your backend system. |
| [Encrypt attachments, process meeting request attendees, and react to appointment date/time changes using Outlook event-based activation](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-encrypt-attachments) | Use event-based activation to encrypt attachments when added by the user. Also use event handling for recipients changed in a meeting request, and changes to the start or end date or time in a meeting request. |
| [Identify and tag external recipients using Outlook event-based activation](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external) | Use event-based activation to run an Outlook add-in when the user changes recipients while composing a message. The add-in also uses the `appendOnSendAsync` API to add a disclaimer. |
| [Set your signature using Outlook event-based activation](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature) | Use event-based activation to run an Outlook add-in when the user creates a new message or appointment. The add-in can respond to events, even when the task pane isn't open. It also uses the `setSignatureAsync` API. |
| [Verify the color categories of a message or appointment before it's sent using Smart Alerts](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories) | Use Outlook Smart Alerts to verify that required color categories are applied to a new message or appointment before it's sent. |
| [Verify the sensitivity label of a message](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-verify-sensitivity-label) | Use the sensitivity label API in an event-based add-in to verify and apply the **Highly Confidential** sensitivity label to applicable outgoing messages. |

## Word

| Name                | Description         |
|:--------------------|:--------------------|
| [Get, edit, and set OOXML content in a Word document with a Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-get-set-edit-openxml) | This sample shows how to get, edit, and set OOXML content in a Word document. The sample add-in provides a scratch pad to get Office Open XML for your own content and test your own edited Office Open XML snippets.|
| [Import a Word document template with a Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-import-template) | Shows how to import templates in a Word document. |
| [Load and write Open XML in your Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml)  | This sample add-in shows you how to add a variety of rich content types to a Word document using the setSelectedDataAsync method with ooxml coercion type. The add-in also gives you the ability to show the Office Open XML markup for each sample content type right on the page. |
| [Manage citations with your Word add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-citation-management) | Shows how to manage citations in a Word document. |

## Authentication, authorization, and single sign-on (SSO)

| Name                | Description         |
|:--------------------|:--------------------|
| [Office Add-in with SSO using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-SSO-NAA) | Shows how to use MSAL.js nested app authentication (NAA) in an Office Add-in to access Microsoft Graph APIs for the signed-in user. The sample displays the signed-in user's name and email. It also inserts the names of files from the user's Microsoft OneDrive account into the document. |
| [Outlook add-in with SSO using nested app authentication](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA) | Shows how to use MSAL.js nested app authentication (NAA) in an Outlook Add-in to access Microsoft Graph APIs for the signed-in user. The sample displays the signed-in user's name and email. It also inserts the names of files from the user's Microsoft OneDrive account into a new message body.|
| [Use SSO with event-based activation in an Outlook add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-events) | Shows how to use SSO to access a user's Microsoft Graph data from an event fired in an Outlook add-in. |
| [Single Sign-on (SSO) Sample Outlook Add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO) | Use Office's SSO feature to give the add-in access to Microsoft Graph data.|
| [Get OneDrive data using Microsoft Graph and msal.js in an Office Add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React) | Build an Office Add-in, as a single-page application (SPA) with no backend, that connects to Microsoft Graph, and access workbooks stored in OneDrive for Business to update a spreadsheet.  |
| [Office Add-in auth to Microsoft Graph](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET) | Learn how to build a Microsoft Office Add-in that connects to Microsoft Graph, and access workbooks stored in OneDrive for Business to update a spreadsheet. |
| [Outlook Add-in auth to Microsoft Graph](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET). | Build an Outlook add-in that connects to Microsoft Graph, and access workbooks stored in OneDrive for Business to compose a new email message. |
| [Single Sign-on (SSO) Office Add-in with ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO) | Use the `getAccessToken` API in Office.js to give the add-in access to Microsoft Graph data. This sample is built on ASP.NET. |
| [Single Sign-on (SSO) Office Add-in with Node.js](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) | Use the `getAccessToken` API in Office.js to give the add-in access to Microsoft Graph data. This sample is built on Node.js.|

## Office

| Name                | Description         |
|:--------------------|:--------------------|
| [Save custom settings in your Office Add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-save-custom-settings) | Save custom settings inside an Office Add-in. The add-in stores data as key-value pairs, using the JavaScript API for Office property bag, browser cookies, web storage (localStorage and sessionStorage), or by storing the data in a hidden div in the document. |
| [Use keyboard shortcuts for Office Add-in actions](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-keyboard-shortcuts) | Create custom keyboard shortcuts to invoke certain actions for your Office Add-in. |

## Shared runtime

| Name                | Description         |
|:--------------------|:--------------------|
| [Share global data with a shared runtime](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-global-state) | Set up a basic project that uses the shared runtime to run code for ribbon buttons, task pane, and custom functions in a single browser runtime. |
| [Manage ribbon and task pane UI, and run code on doc open](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-scenario) | Create contextual ribbon buttons that are enabled based on the state of your add-in. |

## Additional samples

| Name                | Description         |
|:--------------------|:--------------------|
| [Use a shared library to migrate your Visual Studio Tools for Office add-in to an Office web add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/VSTO-shared-code-migration) | Provides a strategy for code reuse when migrating from VSTO Add-ins to Office Add-ins. |
| [Integrate an Azure function with your Excel custom function](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/AzureFunction) | Integrate Azure functions with custom functions to move to the cloud or integrate additional services. |
| [Dynamic DPI code samples](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/dynamic-dpi) | A collection of samples for handling DPI changes in COM, VSTO, and Office Add-ins. |

## Next steps

Join the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram) to get resources and information to help you build solutions for the Microsoft 365 platform, including recommendations tailored to your areas of interest.

You might also qualify for a free developer subscription that's renewable for 90 days and comes configured with sample data; for details, see the [FAQ](/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-).
