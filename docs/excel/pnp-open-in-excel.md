---
title: Create an Excel workbook from a web site with an auto-open task pane
description: Create an Excel workbook from your web page with data and configure a custom Office Add-in task pane that automatically opens.
ms.date: 01/14/2026
ms.topic: sample
ms.localizationpriority: medium
---

# Create an Excel workbook from a web site with an auto-open task pane

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Diagram illustrating how the Excel button on your web page opens a new Excel document and AutoOpens your add-in in the right pane.":::

Microsoft partners with SaaS web applications know that their customers often want to open their data from a web page in an Excel spreadsheet. They use Excel to do analysis on the data, or other types of number crunching. Then they upload the data back to the web site.

Instead of multiple steps to export the data from the web site to a .csv file, import the .csv file into Excel, work with the data, then export it from Excel, and upload it back to the web site, you can simplify this process to one button click.

This article shows how to add an Excel button to your web site. When a customer chooses the button, it automatically creates a new spreadsheet with the requested data, uploads it to the customer's OneDrive, and opens it in Excel on a new browser tab. With one click the requested data is opened in Excel and formatted correctly. Additionally the pattern embeds your own Office Add-in inside the spreadsheet so that customers can still access your services from the context of Excel.

Microsoft partners who implemented this pattern have seen increased customer satisfaction. They've also seen a significant increase in engagement with their add-ins by embedding them in the Excel spreadsheet. If you have a web site for customers to work with data, consider implementing this pattern in your own solution.

## Prerequisites

- [Node.js](https://nodejs.org/) version 16 or later.
- [Visual Studio Code](https://code.visualstudio.com/Download).
- A Microsoft 365 account. You can get one if you qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram); for details, see the [FAQ](/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g).
- At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.

## Run the sample code

The sample code for this article is named [Create an Excel workbook from a web site with an auto-open task pane](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-create-worksheet-from-web-site). To run the sample, follow the instructions in the [readme](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-create-worksheet-from-web-site).

## Solution architecture

:::image type="content" source="../images/open-in-excel-architecture.svg" alt-text="The sequence of steps to create a spreadsheet, populate it with data, and open it on a new browser tab for the user.":::

The solution described in this article adds an **Open in Microsoft Excel** button to the web site and interacts with a Node.js server API and the Microsoft Graph API. The following sequence of events occurs when the user wants to open their data in a new Excel spreadsheet.

1. The user selects the **Open in Microsoft Excel** button. The web page passes the data to an API endpoint on the Node.js server.
1. The server uses the ExcelJS library to create a new Excel spreadsheet in memory. It populates the spreadsheet with the data and embeds your Office Add-in.
1. The server returns the spreadsheet as a binary blob to the web page.
    > [!IMPORTANT]
    > The sample code is designed for development and demonstration purposes only. In a production environment, you **must** implement authentication and authorization for the `/api/create-spreadsheet` endpoint to ensure only authorized users can generate spreadsheets. Without proper security measures, bad actors could exploit this endpoint to generate excessive spreadsheets, consume server resources, or access data inappropriately.
1. The web page calls the Microsoft Graph API to upload the spreadsheet to the user's OneDrive.
1. Microsoft Graph returns the web url location of the new spreadsheet file.
1. The web page opens a new browser tab to open the spreadsheet at the web url. The spreadsheet contains the data and your embedded add-in.

## Key parts of the solution

The solution consists of a Node.js web application that serves both the client-side web pages and provides an API endpoint for creating spreadsheets. The following sections describe important concepts and implementation details for constructing the solution. A full reference implementation can be found in the [sample code](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-create-worksheet-from-web-site) for additional implementation details.

### Excel button and Fluent UI

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Fluent UI icons for Word, Excel, and PowerPoint.":::

You need a button on the web site that creates the Excel spreadsheet. A best practice is to use the Fluent UI to help your users transition between Microsoft products. Always use an Office icon to indicate which Office application the web page launches. For more information, see [Office Brand Icons](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) on the Fluent UI developer portal.  

### Sign in the user

The sample code uses authentication built from the Microsoft identity sample named [Vanilla JavaScript single-page application using MSAL.js to authenticate users to call Microsoft Graph](https://github.com/Azure-Samples/ms-identity-javascript-tutorial/blob/main/2-Authorization-I/1-call-graph/README.md). The authentication code is integrated directly into the web application. For more information about writing code for authentication and authorization, see this sample. For a full list of identity samples for a wide range of platforms, see [Microsoft identity platform code samples](/azure/active-directory/develop/sample-v2-code).

### Create the spreadsheet with ExcelJS

The sample code uses the [ExcelJS](https://github.com/exceljs/exceljs) library to create the spreadsheet. ExcelJS is a JavaScript library that you can use to read, manipulate, and write Excel files. The Node.js server provides an API endpoint at `/api/create-spreadsheet` that constructs the spreadsheet in memory and returns it as a binary blob.

To embed an add-in, the sample uses [JSZip](https://www.npmjs.com/package/jszip) and [xml2js](https://www.npmjs.com/package/xml2js) to manipulate the Office Open XML (OOXML) structure of the Excel file.

> [!CAUTION]
> The sample's `/api/create-spreadsheet` endpoint doesn't include authentication or authorization. Before deploying to production, add security measures to verify the user's identity and ensure they have permission to generate spreadsheets with the requested data. Consider implementing:
>
> - Token-based authentication, such as JWT tokens
> - Session validation
> - Rate limiting to prevent abuse
> - Input validation and sanitization

### Populate the spreadsheet with data

The `/api/create-spreadsheet` endpoint in the Node.js server accepts an HTTP POST request with a JSON body containing the row and column data. The server code iterates through all rows and columns and adds them to the worksheet using ExcelJS. The data format is a simple JSON structure with rows and columns, as defined in the sample's `tableData.js` file.

### Embed your Office Add-in inside the spreadsheet

The sample embeds a custom add-in into the spreadsheet. Embedding Office Add-ins requires manipulating the Office Open XML (OOXML) structure of the Excel file. The sample's `embedAddin` function in `server.js` performs the following operations:

- Adds `webextension1.xml` with the add-in reference.
- Adds `taskpanes.xml` to configure task pane behavior.
- Updates `[Content_Types].xml` to register the web extension parts.
- Updates `workbook.xml.rels` to link the taskpane configuration.

> [!IMPORTANT]
> The auto-open feature only works if the add-in is already installed (sideloaded or deployed) on the user's machine. If the user hasn't installed the add-in, the embedded reference is ignored. For security reasons, Office Add-ins can't force themselves to open without prior user interaction.

### Configure auto-open behavior

The sample demonstrates the auto-open feature, which allows the task pane to automatically open when the user opens the workbook. However, this feature requires a specific workflow.

1. **First open**: The user opens the downloaded file. The task pane doesn't auto-open yet.
1. **Manual activation**: The user selects the **Show Task Pane** button on the ribbon.
1. **Enable auto-open**: The user selects **Set auto-open ON** in the task pane.
1. **Save the file**: The user saves and closes the file.
1. **Subsequent opens**: The task pane now automatically opens when the user reopens the file.

Office.js controls the auto-open behavior by setting the document property `Office.AutoShowTaskpaneWithDocument`.

```javascript
function setAutoOpenOn() {
    Office.context.document.settings.set(
        'Office.AutoShowTaskpaneWithDocument',
        true
    );
    Office.context.document.settings.saveAsync();
}
```

This workflow is standard Office Add-in behavior. For more information, see [Automatically open a task pane with a document](../develop/automatically-open-a-task-pane-with-a-document.md).

### Upload the spreadsheet to OneDrive

When the spreadsheet is fully constructed, the server returns it as a binary blob to the web application. Then the web application uses the Microsoft Graph API to upload the spreadsheet to the user's OneDrive. The web application creates the file in the user's OneDrive root folder with a timestamped filename, but you can modify the code to use any folder and filename you prefer.

## Security considerations for production deployment

> [!IMPORTANT]
> The sample code is for educational purposes and isn't production-ready. Before deploying this pattern to a production environment, implement comprehensive security measures.

### Secure the API endpoint

The `/api/create-spreadsheet` endpoint acts as a critical security boundary. Without proper protection, malicious actors can:

- Generate unlimited spreadsheets, consuming server resources and storage.
- Access data they shouldn't have permission to view.
- Launch denial-of-service attacks against your server.

**Required security measures:**

1. **Authentication**: Verify the user's identity before processing any request. Use established authentication protocols such as:
   - OAuth 2.0 tokens
   - JSON Web Tokens (JWT)
   - Session-based authentication with secure cookies

1. **Authorization**: Confirm the authenticated user has permission to access the specific data they're requesting. Implement:
   - Role-based access control (RBAC)
   - Data-level permissions checking
   - Validation that the user owns or has access to the requested data

1. **Input validation**: Always validate and sanitize input data to prevent injection attacks and ensure data integrity.

1. **Rate limiting**: Implement rate limiting to prevent abuse and protect against denial-of-service attacks.

1. **HTTPS**: Always use HTTPS in production to encrypt data in transit.

1. **CORS configuration**: Configure Cross-Origin Resource Sharing (CORS) properly to allow only trusted domains.

### Additional security best practices

- Regularly update all npm packages to address security vulnerabilities by using `npm audit`.
- Implement logging and monitoring to detect suspicious activity.
- Use environment variables for sensitive configuration (never hard-code secrets).
- Consider implementing Content Security Policy (CSP) headers.
- Validate file sizes and content to prevent oversized or malicious spreadsheets.

## Additional considerations for your solution

Everyone’s solution is different in terms of technologies and approaches. The following considerations help you plan how to modify your solution to open documents and embed your Office Add-in.

### Store data in the spreadsheet when your add-in starts

If you implement add-in embedding, you can store metadata in the spreadsheet for your add-in to access. Excel provides several ways to persist data within a workbook:

- **Settings API**: Use [Excel.SettingCollection](/javascript/api/excel/excel.settingcollection) to store settings unique to your add-in and the workbook.
- **Custom properties**: Use [Excel.DocumentProperties.custom](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-custom-member) to store key-value pairs in the workbook's metadata.
- **Custom XML data**: Use [Excel.Workbook.customXmlParts](/javascript/api/excel/excel.workbook#excel-excel-workbook-customxmlparts-member) to store structured XML data in the workbook.

> [!CAUTION]
> Don't store sensitive information in settings or custom properties such as auth tokens or connection strings. Data stored in the spreadsheet isn't encrypted or protected.

For complete details on persisting data in Excel workbooks, see [Persist add-in state and settings](../develop/persisting-add-in-state-and-settings.md).

### Use single sign-on

To simplify authentication, implement single sign-on in your add-in. This approach ensures the user doesn't need to sign in a second time to access your add-in. For more information, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md).

## See also

- [Welcome to the Open XML SDK 2.5 for Office](/office/open-xml/open-xml-sdk)
- [Automatically open a task pane with a document](../develop/automatically-open-a-task-pane-with-a-document.md)
- [Persisting add-in state and settings](../develop/persisting-add-in-state-and-settings.md)
- [Create a spreadsheet document by providing a file name](/office/open-xml/spreadsheet/how-to-create-a-spreadsheet-document-by-providing-a-file-name)
