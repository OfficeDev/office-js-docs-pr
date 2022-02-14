---
title: Office Add-in code samples
description: 'A list of Office Add-in code samples to help you learn and build your own add-ins.'
ms.date: 11/18/2021
localization_priority: high
---

# Office Add-in code samples

These code samples are written to help you learn how to use various features when developing Office Add-ins.

## Getting started

The following samples show how to build the simplest Office Add-in with only a manifest, HTML web page, and a logo. These components are the fundamental parts of an Office Add-in. For additional getting started information, see our [quick starts](../quickstarts/excel-quickstart-jquery.md) and [tutorials](/search/?terms=tutorial&scope=Office%20Add-ins).

- [Excel "Hello world" add-in](/samples/officedev/pnp-officeaddins/excel-add-in-hello-world/)
- [Outlook "Hello world" add-in](/samples/officedev/pnp-officeaddins/outlook-add-in-hello-world/)
- [PowerPoint "Hello world" add-in](/samples/officedev/pnp-officeaddins/powerpoint-add-in-hello-world/)
- [Word "Hello world" add-in](/samples/officedev/pnp-officeaddins/word-add-in-hello-world/)

## Outlook

| Name                | Description         |
|:--------------------|:--------------------|
| [Encrypt attachments, process meeting request attendees, and react to appointment date/time changes](/samples/officedev/pnp-officeaddins/outlook-add-in-encrypt-attachments) | Use event-based activation to encrypt attachments when added by the user. Also use event handling for recipients changed in a meeting request, and changes to the start or end date or time in a meeting request. |
| [Use Outlook event-based activation to tag external recipients (preview)](/samples/officedev/pnp-officeaddins/outlook-add-in-tag-external-recipients/) | Use event-based activation to run an Outlook add-in when the user changes recipients while composing a message. The add-in also uses the `appendOnSendAsync` API to add a disclaimer. |
| [Use Outlook event-based activation to set the signature](/samples/officedev/pnp-officeaddins/outlook-add-in-set-signature/) | Use event-based activation to run an Outlook add-in when the user creates a new message or appointment. The add-in can respond to events, even when the task pane is not open. It also uses the `setSignatureAsync` API. |

## Excel

| Name                | Description         |
|:--------------------|:--------------------|
| [Open in Teams](/samples/officedev/pnp-officeaddins/office-excel-add-in-open-in-teams/) | Create a new Excel spreadsheet in Microsoft Teams containing data you define.|
| [Insert an external Excel file and populate it with JSON data](/samples/officedev/pnp-officeaddins/excel-add-in-insert-external-file/)  | Insert an existing template from an external Excel file into the currently open Excel workbook. Then, populate the template with data from a JSON web service. |
| [Create custom contextual tabs on the ribbon](/samples/officedev/pnp-officeaddins/office-add-in-contextual-tabs/) | Create a custom contextual tab on the ribbon in the Office UI. The sample creates a table, and when the user moves the focus inside the table, the custom tab is displayed. When the user moves outside the table, the custom tab is hidden. |
| [Use keyboard shortcuts for Office add-in actions](/samples/officedev/pnp-officeaddins/office-add-in-keyboard-shortcuts) | Set up a basic Excel add-in project that utilizes keyboard shortcuts. |
| [Custom function sample using web worker](/samples/officedev/pnp-officeaddins/excel-custom-function-web-worker-pattern/) | Use web workers in custom functions to prevent blocking the UI of your Office Add-in. |
| [Use storage techniques to access data from an Office Add-in when offline](/samples/officedev/pnp-officeaddins/use-storage-techniques-to-access-data-from-an-office-add-in-when-offline/) | Implement localStorage to enable limited functionality for your Office Add-in when a user experiences lost connection. |
| [Custom function batching pattern](/samples/officedev/pnp-officeaddins/excel-custom-function-batching-pattern/)| Batch multiple calls into a single call to reduce the number of network calls to a remote service.|

## Shared JavaScript runtime

| Name                | Description         |
|:--------------------|:--------------------|
[Share global data with a shared runtime](/samples/officedev/pnp-officeaddins/office-add-in-shared-runtime-global-data/) | Set up a basic project that uses the shared runtime to run code for ribbon buttons, task pane, and custom functions in a single browser runtime. |
| [Manage ribbon and task pane UI, and run code on doc open](/samples/officedev/pnp-officeaddins/office-add-in-ribbon-task-pane-ui/) | Create contextual ribbon buttons that are enabled based on the state of your add-in. |

## Authentication, authorization, and single sign-on (SSO)

| Name                | Description         |
|:--------------------|:--------------------|
| [Single Sign-on (SSO) Sample Outlook Add-in](/samples/officedev/pnp-officeaddins/outlook-add-in-sso-aspnet/) | Use Office's SSO feature to give the add-in access to Microsoft Graph data.|
| [Get OneDrive data using Microsoft Graph and msal.js in an Office Add-in](/samples/officedev/pnp-officeaddins/office-add-in-auth-graph-react/) | Build an Office Add-in, as a single-page application (SPA) with no backend, that connects to Microsoft Graph, and access workbooks stored in OneDrive for Business to update a spreadsheet.  |
| [Office Add-in auth to Microsoft Graph](/samples/officedev/pnp-officeaddins/office-add-in-auth-aspnet-graph/) | Learn how to build a Microsoft Office Add-in that connects to Microsoft Graph, and access workbooks stored in OneDrive for Business to update a spreadsheet. |
| [Outlook Add-in auth to Microsoft Graph](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET). | Build an Outlook add-in that connects to Microsoft Graph, and access workbooks stored in OneDrive for Business to compose a new email message. |
| [Single Sign-on (SSO) Office Add-in with ASP.NET](/samples/officedev/pnp-officeaddins/office-add-in-sso-aspnet/) | Use the `getAccessToken` API in Office.js to give the add-in access to Microsoft Graph data. This sample is built on ASP.NET. |
| [Single Sign-on (SSO) Office Add-in with Node.js](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) | Use the `getAccessToken` API in Office.js to give the add-in access to Microsoft Graph data. This sample is built on Node.js.|

## Additional samples

| Name                | Description         |
|:--------------------|:--------------------|
|[Use a shared library to migrate your Visual Studio Tools for Office add-in to an Office web add-in](/samples/officedev/pnp-officeaddins/vsto-shared-library-excel/) |Provides a strategy for code reuse when migrating from VSTO Add-ins to Office Add-ins. |
| [Integrate an Azure function with your Excel custom function](/samples/officedev/pnp-officeaddins/azure-function-with-excel-custom-function/) | Integrate Azure functions with custom functions to move to the cloud or integrate additional services. |
|[Dynamic DPI code samples](/samples/officedev/pnp-officeaddins/dynamic-dpi-code-samples/) |A collection of samples for handling DPI changes in COM, VSTO, and Office Add-ins. |

## Next steps

Join the Microsoft 365 Developer Program. Get a free sandbox, tools, and other resources you need to build solutions for the Microsoft 365 platform.

- [Free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Get a free, renewable 90-day Microsoft 365 E5 developer subscription.
- [Sample data packs](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Automatically configure your sandbox by installing user data and content to help you build your solutions.
- [Access to experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Access community events to learn from Microsoft 365 experts.
- [Personalized recommendations](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Find developer resources quickly from your personalized dashboard.
