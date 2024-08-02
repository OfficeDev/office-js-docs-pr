---
title: OneNote JavaScript API programming overview
description: Learn about the OneNote JavaScript API for OneNote add-ins on the web.
ms.date: 07/22/2024
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
---

# OneNote JavaScript API programming overview

OneNote introduces a JavaScript API for OneNote add-ins on the web. You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## Components of an Office Add-in

Add-ins consist of two basic components:

- A **web application** consisting of a webpage and any required JavaScript, CSS, or other files. These files are hosted on a web server or web hosting service, such as Microsoft Azure. In OneNote on the web, the web application displays in a webview control or iframe.

- A **manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the [add-in only manifest](../develop/add-in-manifests.md) format.

### Office Add-in = Manifest + Webpage

![Office Add-in consists of a manifest and webpage.](../images/onenote-add-in.png)

## Using the JavaScript API

Add-ins use the runtime context of the Office application to access the JavaScript API. The API has two layers:

- A **application-specific API** for OneNote-specific operations, accessed through the `Application` object.
- A **Common API** that's shared across Office applications, accessed through the `Document` object.

### Accessing the application-specific API through the *Application* object

Use the `Application` object to access OneNote objects such as **Notebook**, **Section**, and **Page**. With application-specific APIs, you run batch operations on proxy objects. The basic flow goes something like this:

1. Get the application instance from the context.

2. Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.

3. Call `load` on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.

   > [!NOTE]
   > Method calls to the API (such as `context.application.getActiveSection().pages;`) are also added to the queue.

4. Call `context.sync` to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.

For example:

```js
async function getPagesInSection() {
    await OneNote.run(async (context) => {

        // Get the pages in the current section.
        const pages = context.application.getActiveSection().pages;

        // Queue a command to load the id and title for each page.
        pages.load('id,title');

        // Run the queued commands, and return a promise to indicate task completion.
        await context.sync();
            
        // Read the id and title of each page.
        $.each(pages.items, function(index, page) {
            let pageId = page.id;
            let pageTitle = page.title;
            console.log(pageTitle + ': ' + pageId);
        });
    });
}
```

See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn more about the `load`/`sync` pattern and other common practices in the OneNote JavaScript APIs.

You can find supported OneNote objects and operations in the [API reference](../reference/overview/onenote-add-ins-javascript-reference.md).

#### OneNote JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For detailed information about OneNote JavaScript API requirement sets, see [OneNote JavaScript API requirement sets](/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets).

### Accessing the Common API through the *Document* object

Use the `Document` object to access the Common API, such as the [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1))
and [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) methods.

For example:  

```js
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            const error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```

OneNote add-ins support only the following Common APIs.

| API | Notes |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) | `Office.CoercionType.Text` and `Office.CoercionType.Matrix` only |
| [Office.context.document.setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) | `Office.CoercionType.Text`, `Office.CoercionType.Image`, and `Office.CoercionType.Html` only |
| [const mySetting = Office.context.document.settings.get(name);](/javascript/api/office/office.settings#office-office-settings-get-member(1)) | Settings are supported by content add-ins only |
| [Office.context.document.settings.set(name, value);](/javascript/api/office/office.settings#office-office-settings-set-member(1)) | Settings are supported by content add-ins only |
| [Office.EventType.DocumentSelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) |*None*|

In general, you use the Common API to do something that isn't supported in the application-specific API. To learn more about using the Common API, see [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).

<a name="om-diagram"></a>

## OneNote object model diagram

The following diagram represents what's currently available in the OneNote JavaScript API.

  ![OneNote object model diagram.](../images/onenote-om.png)

## See also

- [Developing Office Add-ins](../develop/develop-overview.md)
- [Build your first OneNote add-in](../quickstarts/onenote-quickstart.md)
- [OneNote JavaScript API reference](../reference/overview/onenote-add-ins-javascript-reference.md)
- [Sample: Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
