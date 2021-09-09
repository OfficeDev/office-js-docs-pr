---
title: Office Add-in code samples
description: 'Office Add-in code samples'
ms.date: 08/02/2021
localization_priority: Normal
---

# Office Add-in code samples

The code samples listed in this article are written to help you learn how to use various features when developing Office Add-ins.

## Outlook

| Name                | Description         |
|:--------------------|:--------------------|
| [Use Outlook event-based activation to tag external recipients (preview)](/samples/officedev/pnp-officeaddins/outlook-add-in-tag-external-recipients) | This sample uses event-based activation to run an Outlook add-in when the user changes recipients while composing a message. The add-in also uses the appendOnSendAsync API to add a disclaimer. |
| [Use Outlook event-based activation to set the signature](/samples/officedev/pnp-officeaddins/outlook-add-in-set-signature/) | This sample uses event-based activation to run an Outlook add-in when the user creates a new message or appointment. The add-in can respond to events, even when the task pane is not open. It also uses the setSignatureAsync API. |

## Excel

| Name                | Description         |
|:--------------------|:--------------------|
| [Open Excel from your web page and embed your Office Add-in](../excel/pnp-open-in-excel.md)| Extend your SaaS web application so that your customers can open their data from a web page directly to Microsoft Excel. |
| [Integrate an Azure function with your Excel custom function](/samples/officedev/pnp-officeaddins/integrate-an-azure-function-with-your-excel-custom-function/) | Learn how to integrate Azure functions with custom functions to move to the cloud or integrate additional services. |
| [Custom function sample using web worker](/samples/officedev/pnp-officeaddins/custom-function-sample-using-web-worker/) | This sample shows how to use web workers in custom functions to prevent blocking the UI of your Office Add-in. |
| [Custom function batching pattern](/samples/officedev/pnp-officeaddins/custom-function-batching-pattern/) | Batch multiple calls into a single call to reduce the number of network calls to a remote service. |

## Shared JavaScript runtime

| Name                | Description         |
|:--------------------|:--------------------|
|[Custom contextual tabs on the ribbon](/samples/officedev/pnp-officeaddins/office-add-in-contextual-tabs/)  |This sample shows how to create a custom contextual tab on the ribbon in the Office UI. The sample creates a table, and when the user moves the focus inside the table, the custom tab is displayed. When the user moves outside the table, the custom tab is hidden. |
| [Use keyboard shortcuts for Office add-in actions](/samples/officedev/pnp-officeaddins/office-add-in-keyboard-shortcuts/) |Shows how to set up a basic Excel add-in project that utilizes keyboard shortcuts. Currently, the shortcuts are configured to show and hide the task pane as well as cycle through colors for a selected cell. |
| [Share global data with a shared runtime](/samples/officedev/pnp-officeaddins/office-add-in-shared-runtime-global-data/) |This sample shows how to set up a basic project that uses the shared runtime. The shared runtime runs all parts of the Excel add-in (ribbon buttons, task pane, custom functions) in a single browser runtime. This makes it easy to shared data through local storage, or through global variables. |
| [ Manage ribbon and task pane UI, and run code on doc open](/samples/officedev/pnp-officeaddins/office-add-in-ribbon-task-pane-ui/) |This sample shows how to create contextual ribbon buttons that are enabled based on the state of your add-in. It also shows how to use the Office.js API to show or hide the task pane. This sample also demonstrates how to run code when the task pane is closed, such as on document open. |

## Authentication, authorization, and single sign-on (SSO)

| Name                | Description         |
|:--------------------|:--------------------|
|[Single Sign-on (SSO) Sample Outlook Add-in](/samples/officedev/pnp-officeaddins/outlook-add-in-sso-aspnet/) |The sample implements an Outlook add-in that uses Office's SSO feature to give the add-in access to Microsoft Graph data. Specifically, it enables the user to save all attachments to their OneDrive. It also shows how to add custom buttons to the Outlook ribbon. |
|[Get OneDrive data using Microsoft Graph and MSAL.NET in an Office Add-in](/samples/officedev/pnp-officeaddins/office-add-in-auth-aspnet-graph/) |Learn how to build a Microsoft Office Add-in that connects to Microsoft Graph, finds the first three workbooks stored in OneDrive for Business, fetches their filenames, and inserts the names into an Office document using Office.js. |

## Additional samples

| Name                | Description         |
|:--------------------|:--------------------|
|[Use a shared library to migrate your Visual Studio Tools for Office add-in to an Office web add-in](/samples/officedev/pnp-officeaddins/vsto-shared-library-excel/) |Provides a strategy for code reuse when migrating from VSTO Add-ins to Office Add-ins. |
|[Dynamic DPI code samples](/samples/officedev/pnp-officeaddins/dynamic-dpi-code-samples/) |A collection of samples for handling DPI changes in COM, VSTO, and Office Add-ins. |

## Next steps

Join the Microsoft 365 Developer Program. Get a free sandbox, tools, and other resources you need to build solutions for the Microsoft 365 platform.

- [Free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Get a free, renewable 90-day Microsoft 365 E5 developer subscription.
- [Sample data packs](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Automatically configure your sandbox by installing user data and content to help you build your solutions.
- [Access to experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Access community events to learn from Microsoft 365 experts.
- [Personalized recommendations](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Find developer resources quickly from your personalized dashboard.