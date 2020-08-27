---
title: Fundamental programming concepts with the Word JavaScript API
description: 'Use the Word JavaScript API to build add-ins for Word.'
ms.date: 07/28/2020
localization_priority: Priority
---

# Fundamental programming concepts with the Word JavaScript API

This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins for Word 2016 or later.

## Referencing Office.js

You can reference Office.js from the following locations:

- `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - use this resource for production add-ins.

- `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - use this resource to try out preview features.

## Word JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For detailed information about Word JavaScript API requirement sets, see [Word JavaScript API requirement sets](../reference/requirement-sets/word-api-requirement-sets.md).

## Running Word add-ins

To run your add-in, use an `Office.initialize` event handler. For more information about add-in initialization, see [Understanding the API](../develop/understanding-the-javascript-api-for-office.md).

Add-ins that target Word 2016 or later can use the Word-specific APIs. They pass the Word-interaction logic as a function into the `Word.run()` method. See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about how to interact with the Word document in this programming model.

The following example shows how to initialize and run a Word add-in by using the `Word.run()` method.

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

## See also

- [Word JavaScript API overview](../reference/overview/word-add-ins-reference-overview.md)
- [Build your first Word add-in](../quickstarts/word-quickstart.md)
- [Word add-in tutorial](../tutorials/word-tutorial.md)
- [Word JavaScript API reference](/javascript/api/word)
