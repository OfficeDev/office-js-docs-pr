# Word JavaScript API reference

Word provides a rich set of APIs that you can use to create add-ins that interact with document content and metadata. Use these APIs to create compelling experiences that integrate with and extend Word. You can import and export content, assemble new documents from different data sources, and integrate with document workflows to create custom document solutions.

You can use two JavaScript APIs to interact with the objects and metadata in a Word document:

- Word JavaScript API - Introduced in Office 2016.
- [JavaScript API for Office](../javascript-api-for-office.md) (Office.js) - Introduced in Office 2013.

## Word JavaScript API

The Word JavaScript API is loaded by Office.js. The Word JavaScript API changes the way that you can interact with objects like documents and paragraphs. Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the Word JavaScript API provides “proxy” JavaScript objects that correspond to the real objects running in Word. You can interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them. These interactions with proxy objects aren’t immediately realized in the running script. The **context.sync** method synchronizes the state between your running JavaScript and the real objects in Office by executing queued instructions and retrieving properties of loaded Word objects for use in your script.

## JavaScript API for Office

You can reference Office.js from the following locations:

* https://appsforoffice.microsoft.com/lib/1/hosted/office.js - use this resource for production add-ins.
* https://appsforoffice.microsoft.com/lib/beta/hosted/office.js - use this resource when you're trying out preview features.

If you're using [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs), you can download the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) to get project templates that include Office.js.  You can also use [nuget to get Office.js](https://www.nuget.org/packages/Microsoft.Office.js/).

If you use TypeScript and have npm, you can get the the TypeScript definitions by typing this in your command line interface: ```typings install office-js --ambient```.

## Running Word add-ins

To run your add-in, use an Office.initialize event handler. For more information about add-in initialization, see [Understanding the API](../../docs/develop/understanding-the-javascript-api-for-office.md) .

Add-ins that target Word 2016 execute by passing a function into the **Word.run()** method. The function passed into the **run** method must have a context argument. This [context object](../../reference/word/requestcontext.md) is different than the context object you get from the Office object, but it is also used to interact with the Word runtime environment. The context object provides access to the Word JavaScript API object model. The following example shows how to initialize and execute a Word add-in by using the **Word.run()** method.

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

### Synchronizing Word documents with Word JavaScript API proxy objects

The Word JavaScript API object model is loosely coupled with the objects in Word. Word JavaScript API objects are proxies for objects in a Word document. Actions taken on proxy objects are not realized in Word until the document state has been synchronized. Conversely, the state of the Word document is not realized in the proxy objects until the document state has been synchronized. To synchronize the document state, you run the **context.sync()** method. The following example creates a proxy body object and a queued command to load the text property on the proxy body object, and uses the **context.sync()** method to synchronize the body of the Word document with the body proxy object.

```js
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        // The body object hasn't been set with any property values.
        var body = context.document.body;

        // Queue a command to load the text property for the proxy document body object.
        context.load(body, 'text');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });
    })
```

### Executing a batch of commands

The Word proxy objects have methods for accessing and updating the object model. These methods are executed sequentially in the order in which they were queued in the batch. All of the commands that are queued in the batch are executed when context.sync() is called.

The following example shows how the command queue works. When **context.sync()** is called, the [command to load](../../reference/word/loadoption.md) the body text is executed in Word. Then, the command to insert text into the body in Word occurs. The results are then returned to the body proxy object. The value of the **body.text** property in the Word JavaScript API is the value of the Word document body <u>before</u> the text was inserted into Word document.


```js
    // Run a batch operation against the Word JavaScript API.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to load the text property of the proxy body object.
        context.load(body, 'text');

        // Queue a command to insert text into the end of the Word document body.
        body.insertText('This is text inserted after loading the body.text property',
                        Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });
    })
```

## Open Word API specifications

As we design and develop new APIs for Word add-ins, we'll make them available for your feedback on our [Open API specifications](../../reference/openspec.md) page. Find out what new features are in the pipeline for the Word JavaScript APIs, and provide your input on our design specifications.

## Additional resources

* [Word add-ins overview](../../docs/word/word-add-ins-programming-overview.md )
* [Office Add-ins platform overview](../../docs/overview/office-add-ins.md)
* [Word add-in samples on GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Word)
