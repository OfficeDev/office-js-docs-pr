# Word JavaScript APIs (Preview)

Welcome to the Word JavaScript API (Preview) documentation. Here you can find everything you need to create the next generation of Word add-ins in Office 2016 Preview for Windows. The new APIs provide a variety of Word-specific functionality related to documents, paragraphs, content controls, and other common Word objects. This API complements the functionality of our existing Office.js library. 

## Try it out

We've working on a Snippet Explorer (Preview) to let you browse through common code snippets and learn how the new APIs work. Give it a try. The code snippets referenced by the Snippet Explorer are available [here](https://github.com/robmhoward/robmhoward.github.io/tree/master/word/samples). 

You can also [get started with build 4429.1005](#getstarted).

## Give feedback

We've released this preview to give developers an early look at the direction we're heading with the Office developer platform.

* Check out the docs and let us know about any [issues](https://github.com/OfficeDev/office-js-docs/issues) you find in those docs by submitting issues directly in this repo.
* Share your code snippets by submitting a pull request to the Snippet Explorer repo.
* Let us know what you think about the APIs and the general programming experience. 

## API overview  

Below links provide the main-level Word objects and APIs.

* [Document](resources/document.md): The Document object is the top-level object. A Document object contains one or more 
[sections](resources/section.md), a body that contains the content of the document, and header/footer information.
* [Paragraph](resources/paragraph.md): A Paragraph object represents a single paragraph in a selection, range, or document. 
You can access a paragraph through the paragraphs collection in a selection, range, or document. 
* [ContentControl](resources/contentControl.md): A ContentControl object is a container for content. It is a bound and
 potentially labeled region in a document that serves as a container for specific types of content. For example, content 
 controls can contain contents such as paragraphs of formatted text and other content controls. You can access a 
 content control through the content control collection of the document, document body, paragraph, range, or on a content control.
* [Section](resources/section.md):  A Section object is commonly used to define different header and footers as well as 
different page layout configurations of a document. You can access sections from the Document object. 
* [Range](resources/range.md): A Range object represents a contiguous area in a document. You get a Range object when you
 get a selection, insert content into the body, insert content into a content control, insert content into a paragraph, 
 or get a search result. You can define and manipulate a range without changing the selection.
* [Selection](resources/document.md#getselection): Through the document object developers can gave access to the user's selection in the document, or the current insertion point, if nothing is selected.
* [Picture](resources/inlinePicture.md): A Picture object represents an inline image. You can access the inline picture
 collection of the body, content control, paragraph objects.
* [Font](resources/font.md): The Font object provides text formatting to a body, content control, paragraph, or range.

## Programming notes


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

#### references

So far, you specified the objects and properties that you want to load. While that will load the objects into the object model, you still need a handle to make changes to those objects. That is where the references property comes in. The references property gets an identifier for an object so that you can write back to the object. This happens because a reference to that object is persisted in memory. You **MUST** always remove that reference when you are done with the object. 

For example, if you want to use the range object after the executeAsync() call, you'll need to specify that you want a reference to it. Here's how you add a reference to the queue:

```javascript
    ctx.references.add(range);
```

After you've added a reference and acted upon the object, and after you have no more use for the object, you **MUST** remove the reference. You'll queue up the remove reference call before a code path that runs `ctx.executeAsync()`. You add the remove reference call to the queue in one of two ways: 

```javascript
    ctx.references.removeAll(); // removes all object references declared on this request context
    ctx.references.remove(object); // removes a single object reference where 'object' is the object passed into references.add()
```

#### Pulling it all together

Let's put it all together by taking a look at a simple example that shows how you can use the client request context, load method, references, and the executeAsync() method.

**Example: How to load the font color and paragraph text for all fonts and paragraphs in a range** 

```javascript
    // Create the client request context. You'll do this for all Word add-ins.
    var ctx = new Word.RequestContext();

    // Queue: get the user's current selection and create a range object named range.
    // Queue: insert 'Hello World!' at the end of the selection.
    var range = ctx.document.getSelection();
    range.insertText("Hello World!", Word.InsertLocation.end);

    // Queue: load the range object's font color and the text for all paragraphs in 
    // the paragraph collection. 
    ctx.load(range, {select: 'font/color, paragraphs/text', 
                     expand: 'font, paragraphs'});

    // Queue: adds a reference to the range object. You need this to act on the range 
    // object after executeAsync completes. You MUST use references.add() if you will
    // act on an object across executeAsync() calls. For example, we will act on this
    // range object after executeAsync() by inserting a paragraph in to it.
    ctx.references.add(range);

    // Run the set of commands in the queue. In this case, we are inserting text
    // at the end of range and loading font and paragraph collection properties. 
    ctx.executeAsync()
        .then(function () {

            // The range object has been loaded. You can access the font color and 
            // the text content in the paragraph collection on the range object. 

            var contents = '';

            for (i=0; i < range.paragraphs.items.length; i++) {
                contents = contents + range.paragraphs.items[i].text;
            }

            // Show the contents of the paragraphs 
            console.log("OUTPUT: paragraph text in the range object: " + contents);

            // Queue: add a paragraph to the end of the range. We need the reference for this.
            range.insertParagraph("This is a new paragraph.", Word.InsertLocation.after);

            // Queue: remove the reference to the range since we are done writing to it.
            ctx.references.remove(range);

            // Run the set of commands in the queue. In this case, we are adding a page break
            // and removing the reference to the range object. 
            ctx.executeAsync()
                .then(function(){
                    console.log("Success");
                };
        })
        .catch(function(error){
            console.log("ERROR: " + JSON.stringify(error));
        });
```

<a name="getstarted"></a>
### Get started with build 4429.1005

Use these steps to get you started with WordJS. Please open an issue if you encounter a problem using these steps.

1. [Download](https://products.office.com/en-us/office-2016-preview) the latest Office 16 preview (4229.1005 or greater).
2. Put [Word16SampleRegKey.reg](/resources/sampleFiles/Word16SampleRegKey.reg) and [WordAPIs.xml](/resources/sampleFiles/WordAPIs.xml) in the c:\temp directory. Modify the registry file if you place these files in a different directory. The registry key tells Word where it can find WordAPIs.xml. WordAPIs.xml is the manifest file that declares th functionality and the location of the add-in web application.
3. Close all Word, Excel, PowerPoint, and Outlook sessions.
4. Start Word.
5. Select the *Insert* tab, and then select the *My Add-ins* drop-down box. Select the *Word APIs (4229-1005)* add-in. This will load the add-in.
![Select the add-in from the Insert tab](/resources/images/insertAddIn.png)
6. Select the target build (in red), and the target object and sample (in green). ![Select the sample to run](/resources/images/chooseSample.png)
7. Select *Run!* to see the results of running the sample.

The code for this sample is found in this [sample library](https://github.com/trmini/robmhoward.github.io/tree/Word-APIs/word/samples/4229.1005). A great feature of this sample is that you can alter and run the code from within the sample. Contributions to this sample library are encouraged. Please provide feedback on this API, the experience of using it, blocking issues, and the documentation. Your input is appreciated! 


## Release Notes for build 4229.1005

Removed the 'html' and 'ooxml' properties. Added the getHtml() and getOoxml() methods.

Added search to the paragraph object.

Added insertBreak to the range object.
