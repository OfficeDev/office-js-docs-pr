The Office JavaScript API includes two distinct models:

- **Host-specific** APIs provide strongly-typed objects that can be used to interact with objects that are native to a specific Office application. For example, you can use the Excel JavaScript APIs to access worksheets, ranges, tables, charts, and more. Host-specific APIs are currently available for the following hosts:
    - [Excel](../reference/overview/excel-add-ins-reference-overview.md)
    - [Word](../reference/overview/word-add-ins-reference-overview.md)
    - [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)

    This API model uses [promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) and allows you to specify multiple operations in each request you send to the Office host. Batching operations in this manner can significantly improve add-in performance in Office on the web applications. Host-specific APIs were introduced with Office 2016 and cannot be used to interact with Office 2013.

- **Common** APIs can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications. This API model uses [callbacks](https://developer.mozilla.org/docs/Glossary/Callback_function), where you can only specify one operation in each request sent to the Office host. Common APIs were introduced with Office 2013 and can be used to interact with Office 2013 or later. For details about the Common API object model, which includes APIs for interacting with Outlook and PowerPoint, see [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).

> [!NOTE]
> Excel Custom functions run within a unique runtime that prioritizes execution of calculations, and therefore uses a slightly different programming model. For details, see [Custom functions architecture](../excel/custom-functions-architecture.md).