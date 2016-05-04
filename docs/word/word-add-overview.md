# Overview of Word JavaScript development

IMPORTANT INFORMATION BEFORE WE START:
CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js
Beta: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js


The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can create Web applications called add-ins that are hosted in a task pane within the Word UI. You can use these APIs to create add-ins that target:
* (Windows) Word 2013 and later
* (Web) Word Online
* (Mac) Word 2016 and later
* Word for iOS

Write once, and run your add-ins in all versions of Word on different platforms.

## The JavaScript API options

There are two sets of JavaScript APIs that you can use to interact with the objects and metadata in a Word document. The first set of APIs were introduced with Office 2013. These APIs are considered the shared OfficeJS APIs as many of the objects can be used in add-ins hosted by two or more Office clients. Use the [Word](http://dev.office.com/reference/add-ins/javascript-api-for-office?product=word) filter on [dev.office.com](dev.office.com) to get a view of the shared APIs that can be used by Word. This API uses callbacks extensively throughout it.

Starting with Word 2016 for both Mac and Windows, there is a new strongly-typed and Word specific JavaScript object model for creating Word add-ins. This new object model gives you access to familiar objects like [body](../../reference/word/body.md), [content controls](../../reference/word/contentcontrol.md), [inline pictures](../../reference/word/inlinepicture.md) and [paragraphs](../../reference/word/paragraph.md). This API uses promises through out it. This is the preferred API as it provides Word specific objects. We have TypeScript definitions and vsdoc files so that you can get code hints in your IDE.

As of the original publish date of this article, all of the Word clients support the shared API, and most of the clients support the Word specific JavaScript API. We are moving towards having all of the APIs available on all of the Word clients at the same time. Check the reference documentation to learn which clients are supported by the Word specific JavaScript API.

We recommend that you start with the Word specific APIs as the object model is easier to use. Use the Word specific API if you need to:
* Access the objects in the Word document.

Use the shared API when you need to:
* Target Word 2013
* Perform initial actions for the application.
* Check the supported requirement set.
* Access metadata, settings, and environmental information for the document.
* Bind to sections in a document and capture events.
* Use custom XML parts.
* Open a dialog box.

## How do add-ins work in Word?

The add-in acts as a client of the Word application. Sends batches of commands to the service.

## OfficeJS and Word

The OfficeJS that you download via Nuget, or that you reference from our CDN, is a module loader. OfficeJS loads modules from the CDN based on the Word client. Each Word client uses a slightly different library that targets the specific platform environment of the Word client.

Requirement sets - brief description and link to main topic.

## OpenSpec -
Identify openspec. How to find the branch and give feedback on features.



## Additional resources

Links to API reference and docs. Describe the organization of the content.

Links to other resources.





Before we start going into the specifics of the JavaScript API for Word, it is important to know that this new Word add-in object model is different than the model for Word in Office 2013. The Office 2013 add-in model was not typed and provided a generic API for extending Office clients. While the previous model is still applicable to Word 2016, we also suggest that you start using the new Word object model. We suggest that you read the [platform overview](https://msdn.microsoft.com/EN-US/library/office/jj220082.aspx) if you aren't familiar with the add-in platform.

The new JavaScript APIs for Word change the way that you can interact with objects like documents and paragraphs. Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the new APIs provide “proxy” JavaScript objects that correspond to the real objects running in Word.  You can directly interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them.  These interactions with proxy objects aren’t immediately realized in the running script, so we provide a method on the context named **sync()**. The context.sync method synchronizes the state between your running JavaScript and the real objects in Office by executing instructions queued in your script and by retrieving properties of loaded Word objects for use in your script.


## Give us your feedback

Your feedback is important to us.

* Check out the docs and let us know about any questions and issues you find in them by [submitting an issue](https://github.com/OfficeDev/office-js-docs/issues) directly in this repository.
* Let us know about your programming experience, what you would like to see in future versions, code samples, etc. Use [this site](http://officespdev.uservoice.com/) for entering your suggestions and ideas.

## Additional resources

* [Office Add-ins](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Get started with Office Add-ins](http://dev.office.com/getting-started/addins)
* [Word add-ins on GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Word)