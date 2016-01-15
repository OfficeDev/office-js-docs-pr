# Word add-ins programming overview

_Applies to: Word 2016, Word for iPad, Word for Mac_

Word 2016 introduces a new object model for working with Word objects. This object model is an addition to the existing object model provided by Office.js to create add-ins for Word. This object model is accessed via JavaScript hosted by a web application.

## Manifest

The new Word add-in JavaScript API uses the same manifest format as is used for the Office 2013 add-in model. The manifest describes where the add-in is hosted, how it is displayed, permissions, and other information. Learn more about how you can customize the [add-in manifests](https://msdn.microsoft.com/en-us/library/office/fp161044.aspx). 

You have a few options for publishing Word add-in manifests. Read about how you can [publish your Office add-in](https://msdn.microsoft.com/EN-US/library/office/fp123515.aspx) to a network share, to an app catalog, or to the Office store.

## Understanding the JavaScript API for Word

The JavaScript API for Word is loaded by Office.js. It provides a set of JavaScript proxy objects that are used to queue a set of commands that interact with the contents of a Word document. These commands are run as a batch. The results of the batch are actions taken on the Word document, like inserting content, and synchronizing the Word objects with the JavaScript proxy objects. 

### Running your add-in

Let's take a look at what you'll need when you run your add-in. All add-ins should have an Office.initialize event handler.  Read [Understanding the API](https://msdn.microsoft.com/EN-US/library/fp160953.aspx) for more information about add-in initialization.  

Your Word add-in executes by passing a function into the Word.run() method. The function passed into the run method must have a context argument. This [context object](word-add-ins-javascript-reference/requestcontext.md) is different than the context object you get from the Office object, although it is used for the same purpose which is to interact with the Word runtime environment. The context object provides access to the Word JavaScript object model. Let's take a look at the comments and code of a basic Word add-in:

**Example 1. Initialization and execution of a Word add-in**

```javascript
    (function () {
        "use strict";

        // The initialize event handler is run each time the page is loaded.
        Office.initialize = function (reason) {
            
            // Checks for the DOM to load using the jQuery ready function.
            $(document).ready(function () {
                // Set your initialization code. You can use the reason 
                // argument to determine how the add-in was loaded.
                // You can also load saved settings from the Office object.
            });
        };

        // Run a batch operation against the Word object model.
        // Use the context argument to get access to the Word document.
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;
        })
    })();
```

Example 1. shows the basic code needed to create a Word add-in. It initializes Office.js and contains a run method for interacting with the Word document.

### Proxy objects

The Word JavaScript object model is loosely coupled with the objects in Word. The Word JavaScript objects are proxy objects for the real objects in a Word document. All actions taken on proxy objects are not realized in Word, and the state of the Word document is not realized in the proxy objects, until the document state has been synchronized. The document state is synchronized when context.sync() is run. The sync() method essentially runs the set of commands in queue for each proxy object.  Example 2 shows the creation of a proxy body object and a queued command to load the text property on the proxy body object, and then the synchronization of the body in the Word document with the body proxy object. 

**Example 2. Synchronization of the document body with the body proxy object.**

```javascript
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

### Command queue

The Word proxy objects have methods for accessing and updating the object model. These methods are executed sequentially in the order in which they were queued in the batch. A batch of commands is formed before a context.sync() call is made. All of the commands queued in all of the objects that use the context will be executed.  

In example 3, we demonstrate how the queue of commands works. When context.sync() is called, the first thing that happens is that the [command to load](Word%20Add-ins%20JavaScript%20Reference/loadoption.md) the body text is executed in Word. Then, the command to insert text into the body on Word occurs. The results are then returned to the body proxy object. The value of the body.text property in the Word JavaScript will be the value of the Word document body <u>before</u> the text was inserted into Word document. 

**Example 3. Executing a batch of commands.**

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to load the text in the proxy body object.
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

## Give us your feedback

Your feedback is important to us. 

* Check out the docs and let us know about any questions and issues you find in them by [submitting an issue](https://github.com/OfficeDev/office-js-docs/issues) directly in this repository.
* Let us know about your programming experience, what you would like to see in future versions, code samples, etc. Use [this site](http://officespdev.uservoice.com/) for entering your suggestions and ideas.


## Additional resources

* [Word add-ins](word-add-ins.md)
* [Word add-ins JavaScript reference](word-add-ins-javascript-reference.md)
* [Office Add-ins](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Get started with Office Add-ins](http://dev.office.com/getting-started/addins)
* [Word add-ins on GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Word)
* [Snippet Explorer for Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
