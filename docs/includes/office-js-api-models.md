The Office JavaScript API includes two distinct models:

- **Application-specific** APIs provide strongly-typed objects that can be used to interact with objects that are native to a specific Office application. For example, you can use the Excel JavaScript APIs to access worksheets, ranges, tables, charts, and more. Application-specific APIs are currently available for the following Office applications.

    - [Excel](../reference/overview/excel-add-ins-reference-overview.md)
    - [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)
    - [PowerPoint](../reference/overview/powerpoint-add-ins-reference-overview.md)
    - [Word](../reference/overview/word-add-ins-reference-overview.md)

    This API model uses [promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) and allows you to specify multiple operations in each request you send to the Office application. Batching operations in this manner can significantly improve add-in performance in Office applications on the web. Application-specific APIs were introduced with Office 2016.

    > [!NOTE]
    > There's also an application-specific API for [Visio](../visio/visio-overview.md), but you can use it only in SharePoint Online pages to interact with Visio diagrams that have been embedded in the page. Office Web Add-ins are not supported in Visio.

    See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn more about this API model.

- **Common** APIs can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications. This API model uses [callbacks](https://developer.mozilla.org/docs/Glossary/Callback_function), which allow you to specify only one operation in each request sent to the Office application. Common APIs were introduced with Office 2013 and can be used to interact with any supported Office applications. For details about the Common API object model, which includes APIs for interacting with Outlook, PowerPoint, and Project, see [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).

> [!NOTE]
> Custom functions without a [shared runtime](../testing/runtimes.md#shared-runtime) run in a [JavaScript-only runtime](../testing/runtimes.md#javascript-only-runtime) that prioritizes execution of calculations. These functions use a slightly different programming model.
