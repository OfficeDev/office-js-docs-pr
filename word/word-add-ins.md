*Applies to:* Word 2016

# Word add-ins

Welcome to the Word add-in JavaScript API documentation. The Word JavaScript API is a part of the Office add-in programming model for extending Microsoft Office applications. The add-in programming model uses web applications to host your extension to Word. You can now extend Word with any web platform or language that you prefer. 

## Get started now

Are you the type that wants to read fewer words and just wants to see the code? Then let's go and [build your first Word add-in](build-your-first-word-add-in.md). 

## API Overview

Before we start, it is good to know that this new Word add-in model is different than what was made available with Word in Office 2013. The previous object model was not typed and provided a generic API for extending Office clients. While the previous model is still applicable to Word 2016, we strongly suggest that you start using the new Word object model. This new object model provides access to familiar Word objects like: Body, Sections, Paragraphs, Fonts, Content Controls, and Ranges.

The new JavaScript APIs for Word changes the way that you can interact with objects like documents and paragraphs. Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the new APIs provide “proxy” JavaScript objects that correspond to the real objects running in Word.  You can directly interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them.  These interactions with proxy objects aren’t immediately realized in the running script, though, so we provide a method on the context called **sync()** that synchronizes the state between your running JavaScript and the real objects in Office by executing instructions queued in your script and retrieving properties of loaded Office objects for use in your script.  

Let's take a look at this code and comments to get a better understanding of how the proxy objects are used to interact with the contents of a Word document.

```javascript

    // Get the context. This provides the bridge between the Word application and the add-in.
    var ctx = new Word.RequestContext();

    // Queue: get a handle on the document proxy object. Nothing has changed in the Word document.
    var thisDocument = ctx.document;

    // Queue: save the document proxy object. Nothing has changed in the Word document.
    thisDocument.save();
    
    // Queue: load the save state on the document proxy object. Nothing has changed in the Word document.
    // The current value for the saved property on the document proxy object is null.
    ctx.load(thisDocument, { select: 'saved'});
    
    // Run the batch of commands in the queue. The set of commands set on on the document proxy object
    // are sent to Word. If all of the commands are successful, the Word document will be saved and the
    // value of the *saved* property will be returned and set on the document proxy object. The document
    // proxy object's, and the actual Word document's, *saved* property will be in sync. 
    ctx.executeAsync();
```




## More information

Learn more about extending Word by reading the [Word add-ins programming guide](word-add-ins-programming-guide.md). Peruse the [Word add-ins JavaScript reference](word-add-ins-javascript-reference.md) to learn about the objects you can access. Check out our curated list of [Word add-in code samples](word-add-ins-code-samples.md) to get some ideas on how  you can create Word add-ins.

## Give feedback on the API

The documentation for this API is hosted on GitHub with the intention that we can improve the documentation and API by making it open for [issues](https://github.com/OfficeDev/office-js-docs/issues) against the documentation. Issues can include errors in the documentation, requests for clarification, or requests for improvements in the documentation. We also welcome general feedback about the API and the experience you have with it.

## Additional links

* [Office Add-ins](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Get started with Office Add-ins](http://dev.office.com/getting-started/addins)
* [Word add-ins on GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Word)