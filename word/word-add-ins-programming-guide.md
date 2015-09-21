# Word add-ins programming guide

*Applies to: Word 2016*

Microsoft Word 2016 (this list will be updated when we get new hosts) introduces a new object model for working with Word objects. This object model is an addition to the existing object model provided by OfficeJS to create add-ins for Word. This object model is accessed via JavaScript hosted by a web application.

<!--
What is the value proposition here? How is it better?
What can you do with it? 
Architecture - this is all covered in the existing content
What do people need to know?
Do I need to add content about how this relates to the rest of Office.js? 
Do I need to call out what you need to use Office.js for (gaps)?
-->

## Manifest

The new Word add-in Javascript API uses the same manifest format as is used for the old Word add-in model. The manifest describes where the add-in is hosted, how it is displayed, permissions, and other information. Learn more about how you can customize the [add-in manifests](https://msdn.microsoft.com/en-us/library/office/fp161044.aspx). 

You have a few options for publishing Word add-in manifests. Read about how you can [publish your Office add-in](https://msdn.microsoft.com/EN-US/library/office/fp123515.aspx) to a network share, and app catalog, or to the Office store.

## Understanding the Javascript API for Word

The Javascript API for Word is loaded by Office.js. It provides a set of proxy objects that are used to queue a set of commands that interact with the contents of a Word document. These commands are run as a batch. The results of the batch are actions taken on the Word document like inserting content, and synchronizing the Word objects with the Javascript proxy objects. 

### Running your add-in

Let's take a look at what you'll need when you run your add-in. All add-ins should have an Office.initialize event handler.  Read [Understanding the API](https://msdn.microsoft.com/EN-US/library/fp160953.aspx) for more information about add-in initialization.  Your Word add-in executes by passing a function into the Word.run() method. The function passed into the run method has a context argument. This context object is different than the context object you get from the Office object, although it is used for the same purpose which is to interact with the Word runtime enviroment. The context object provides access to the Word Javascript object model. Let's take a look at the comments and code of a basic Word add-in:

*Example 1. Initialization and execution of a Word add-in*

```javascript
    (function () {
        "use strict";

        // The initialize event handler is run each time a the page is loaded.
        Office.initialize = function (reason) {
            
            // Checks for the DOM to load using the jQuery ready function.
            $(document).ready(function () {
                // Set your initialization code. You can use the reason argument
                // to determine how the add-in was loaded.
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

### Proxy objects

The Word Javascript object model is loosely coupled with the objects in Word. The Word Javascript objects are proxy objects for the real objects in a Word document. All actions taken on proxy objects are not realized in Word, and the state of the Word document is not realized in the proxy objects, until the document state has been synchronized. The document state is synchronized when context.sync() is run. Example 2. shows the creation of a proxy object, a queued command to load the text property on the body proxy object, and then the synchronization of the body contents in the Word document with the body proxy object. 

*Example 2. Synchronization of the document body with the body proxy object.*

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        // The body object hasn't been set with any property values. 
        var body = context.document.body;

        // Queue a commmand to load the text property for the proxy document body object.
        context.load(body, 'text');

        // Synchronize the document state by executing the queued-up commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });  
    })
```

### Command queue

The Word proxy objects have methods for accessing and updating the object model. These methods are executed sequentially in the order in which they were queued in the batch. In example 3., we demonstrate how the queue of commands works. When context.sync() is run, the first thing that happens is that the commmand to load the body text is executed in Word. Then, the command to insert text into the body on Word occurs. The results are then returned to the body proxy object. The value of the body.text property in the Word Javascript will be the value of the Word document body <u>before</u> the text was inserted into Word document. 

*Example 3. Executing a batch of commands.*

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a commmand to load the text in document body.
        context.load(body, 'text');

        // Queue a command to insert text into the end of the body.
        body.insertText('This is text inserted after loading the body.text property',
                        Word.InsertLocation.end);

        // Synchronize the document state by executing the queued-up commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });  
    })
```


-- The methods, how they are used to send instruction to the document, how they, how there is a queue, when they executed

### Synchronize the document and proxy objects
- synchronize from the document to object model. sync from object model to the document.

### The Basics
This section introduces key concepts that you need to understand to work with the Word API. 

#### RequestContext()
All actions that target a Word document start by using the client request context returned by the Word.RequestContext method. The client request context serves two major roles:
* Contains the queue of commands that will be performed on the contents of a Word document.
* Provide the bridge between the Office add-in and the Word application since they run in two different processes. The JavaScript runs in the user's browser within the task pane. Word runs in a different process, and in the case of Word Online, on a remote server cluster.  

Here's how you get the request context:  

```javascript
    var ctx = new Word.RequestContext();
```

You can now create a queue of commands that will target the contents of a Word document.  For example, let's create a set of commands that will get the current selection and add some text to the selection. The selection will be contained in a [range](resources/range.md) object returned by document.getSelection(). We are going to add some text at
the end of the selection. We’ll use the context given in the previous line of code.

```javascript
    var range = ctx.document.getSelection();
    range.insertText("Hello World!", Word.InsertLocation.end);
```

At this point, no changes have occurred. You have specified a set of commands that will run in the future. Let's expand on this by looking at the load method.

#### executeAsync()
The Word JavaScript objects created in the add-ins are local proxy objects. Invoking methods and setting properties queues the set of commands in JavaScript, but does not submit them until executeAsync() is called. executeAsync submits the request queue to Word and returns a promise object, which can be used for chaining further commands. executeAsync runs each command in the order in which they were added to the queue. A best practice is to group related commands into a single executeAsync call.

##### executeAsync() example
This example shows how to insert text at the end of a selection. The queue is filled with two commands: getting the user's selection and inserting text at the end of the user's selection. These commands are ran when ctx.executeAsync() is called. executeAsync() returns a promise which can be used to chain it with other operations.

```javascript
    var ctx = new Word.RequestContext();

    // Queue: get the user's current selection and create a range object named range.
    // Queue: insert 'Hello World!' at the end of the selection.
    var range = ctx.document.getSelection();
    range.insertText("Hello World!", Word.InsertLocation.end);

    // Run the set of commands in the queue. In this case, we are inserting text
    // at the end of the range. 
    ctx.executeAsync()
        .then(function () {
            console.log("Done");
        })
        .catch(function(error){
            console.log("ERROR: " + JSON.stringify(error));
        });
```


#### load()
The load method specifies which collections, objects, and properties will be loaded into the object model.  You use the client request context to specify the load options and the object to load. There are two options for using the load method. We'll use the client request context we created above:

```javascript
    ctx.load(object, options); 
    // or
    object.load(options);
```    
        
`object` identifies the object that will be loaded into the object model.

`options` identifies which properties are loaded and the paging arguments. Properties to load can be specified as either a string, a string of comma-separated values, an array of strings, or in a [loadOption object](#loadOption-object). 

Note -- You can use multiple load statements that will be dispatched in a single executeAsync() call. Do this instead of creating complicated `select` and `expand` statements.

For example, we'll use the context given in the previous code to load the *text* content of all of the paragraphs contained in the current selection that was captured in the range object.

```javascript
    ctx.load(range.paragraphs, 'text');
```

Here is key information for using the load method:
+ You SHOULD specify the property set you want to load for the object in the options parameter. Not including the options parameter is the equivalent of using a "SELECT * from Table1", which will affect performance and SHOULD NOT be done for production applications.
+ If the loaded object is a collection, then the specified properties will be loaded for all objects in the collection.

##### loadOption object

The loadOption object specifies which properties to load and how to page through a collection. There are four loading options:

+ select
+ expand
+ top
+ skip

**select**

You use the select option to load properties that are primitive types. You can use either a string or an object literal to specify which properties to load.  For example, if you are going to make simple load statement, you don't need to create an object literal to specify the property. The following code will load the text string for a range object:

```javascript
    ctx.load(range, 'text');
```

Use commas to separate properties if you use the string form.
```javascript
    ctx.load(range, 'text, style, font');
```

You can specify the property set in the following object literal forms:
```javascript
    {select: 'propertyName'}
    {select: "propertyName1, propertyName2"}
    {select: ['propertyName1', 'propertyName2']}
```

Let's build on the last code snippet and load the *style* property on the range object.

```javascript
    ctx.load(range, 'style');
```

If you take a look at the [range](resources/range.md) object documentation, you can see that you can select the `style`, and `text` properties as they are all primitive types. You use methods to load HTML and OOXML properties. 

There's also a `select` path notation to access properties on objects specified by the `expand` statement.

**expand**

You use the expand option to load properties that are in nested Word API objects and collections. Using the range object from the previous examples, we can load the paragraphCollection and the font object for the range by specifying the objects in the expand option. We identify which properties are returned in the select statement.

```javascript
    ctx.load(range, {select: 'font/color, paragraphs/text', 
                     expand: 'font, paragraphs'});
```

Notice how we specify a path to the selected properties in the select statement. The select statement can be used not only to specify the properties on the loaded object, but also to specify the properties loaded on the child objects identified by the expand option. We would have gotten all the properties for the font object and paragraphs collection if we hadn't added the select statement. It is a best practice to always use the select statement with the expand statement.

Use multiple load method calls if you find that your loadOption objects are getting too complex. 

#### Pulling it all together

Let's put it all together by taking a look at a simple example that shows how you can use the client request context, load method, references, and the executeAsync() method.

**Example: How to load the font color and paragraph text for all fonts and paragraphs in a range** 

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (ctx) {

        // Create a proxy object for the document.
        var thisDocument = ctx.document;

        // Queue a command to get the user's current selection.
        // Create a proxy range object for the selection.
        var range = thisDocument.getSelection();

        // Queue a command to insert text into the selection.
        range.insertText("Hello World!", Word.InsertLocation.end);

        // Queue a command to load the range object's font color and the text 
        // for all paragraphs in the paragraph collection. 
        ctx.load(range, {select: 'font/color, paragraphs/text', 
                         expand: 'font, paragraphs'});    

        // Synchronize the document state by executing the queued-up commands, 
        // and return a promise to indicate task completion.
        return ctx.sync().then(function () {

            // The document has been updated with text inserted in to the selection.
            // The proxy range object that was created based on the selection 
            // has been loaded. You can access the font color and the text content 
            // in the paragraph collection on the proxy range object. 

            var contents = '';

            for (i=0; i < range.paragraphs.items.length; i++) {
                contents = contents + range.paragraphs.items[i].text;
            }

            // Show the contents of the paragraphs 
            console.log("OUTPUT: paragraph text in the range object: " + contents);

            // Queue a command to add a paragraph to the end of the range. 
            range.insertParagraph("This is a new paragraph.", Word.InsertLocation.after);

            // Synchronize the document state by executing the queued-up commands, 
            // and return a promise to indicate task completion. In this case,
            // we are inserting a paragraph into the selection.
            return ctx.sync();
        });  
    })
    .catch(function (error) {
        console.log("Error: " + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });

```
