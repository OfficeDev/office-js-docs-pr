
# Word add-in development overview

<!-- This should start with an introduction to add-ins for Word. Talk about extending the functionality of Word and the "shapes" available - content, add-in commands. Imagine that this topic begins right after the "Types of Office Add-ins" section of the Platform overview topic. (We might replace the content in "Word, Excel, and PowerPoint Add-ins that extend functionality" with links to these client-specific landing pages.) If someone wants to extend Word, what things do they want to know at a high level? Show examples, images, etc. -->

The Word JavaScript APIs let you wed the flexibility of Web development with extending the most popular document editing software to create compelling add-ins for Word. The Word JavaScript APIs let you create add-ins that reside in the Word UI. <!-- Replace this para, which is focused on the APIs, with the broader overview mentioned in previous comment. -->

You can use any server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework -- Ember, Backbone, Angular, React -- or stick with VanillaJS to develop your solution. Use services like Azure to [authenticate](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md) and host your application.

The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can use these APIs to create add-ins that target:

* Word 2013 for Windows
* Word 2016 for Windows
* Word Online
* Word 2016 for Mac
* Word for iOS

Write your add-in once, and it will run in all versions of Word across multiple platforms. For details, see [Office Add-in host and platform availability](https://dev.office.com/add-in-availability).

## JavaScript APIs for Word

You can use two sets of JavaScript APIs to interact with the objects and metadata in a Word document. The first is the **JavaScript API for Office**, which was introduced in Office 2013. This is a shared API -- many of the objects can be used in add-ins hosted by two or more Office clients. This API uses callbacks extensively. To learn more about the JavaScript API for Office, see the Shared API section of the [API Reference](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word) page. <!-- Unfortunately, the filtering doesn't work at the individual API topic level. --> 

The second is the **Word JavaScript API**. This is a strongly-typed object model that you can use to create Word add-ins that target Word 2016 for Mac and Windows. This object model uses promises, and provides access to Word-specific objects like [body](../../reference/word/body.md), [content controls](../../reference/word/contentcontrol.md), [inline pictures](../../reference/word/inlinepicture.md), and [paragraphs](../../reference/word/paragraph.md). The Word JavaScript API includes TypeScript definitions and vsdoc files so that you can get code hints in your IDE.

Currently, all Word clients support the shared JavaScript API for Office, and most clients support the Word JavaScript API. For details about supported clients, see the [API reference documentation](https://dev.office.com/reference/add-ins/word/paragraphcollection?product=word).

We recommend that you start with the Word JavaScript API because the object model is easier to use. Use the Word JavaScript API if you need to:

* Access the objects in a Word document.

Use the shared JavaScript API for Office when you need to:

* Target Word 2013.
* Perform initial actions for the application.
* Check the supported requirement set.
* Access metadata, settings, and environmental information for the document.
* Bind to sections in a document and capture events.
* Use custom XML parts.
* Open a dialog box.

### Word JavaScript API

The Word JavaScript API changes the way that you can interact with objects like documents and paragraphs. Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the Word JavaScript API provides “proxy” JavaScript objects that correspond to the real objects running in Word. You can interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them. These interactions with proxy objects aren’t immediately realized in the running script. The **context.sync** method synchronizes the state between your running JavaScript and the real objects in Office by executing queued instructions and retrieving properties of loaded Word objects for use in your script.

The Word JavaScript API is loaded by Office.js. <!-- I think this paragraph essentially says the same thing as the previous; combine? -->It provides a set of JavaScript proxy objects that are used to queue a set of commands that interact with the contents of a Word document. These commands are run as a batch. The results of the batch are actions taken on the Word document, like inserting content, and synchronizing the Word objects with the JavaScript proxy objects.

## Next steps

<!-- We should think about providing more clear next steps instead of lumping links together in an Additional resources section. -->

Ready to create your first Word add-in? See [Build your first Word add-in](word-add-ins.md). You can also try our interactive [Get started experience](http://dev.office.com/getting-started/addins?product=Word). Use the [add-in manifest](../overview/add-in-manifests.md) to describe where your add-in is hosted and how it is displayed, and define permissions and other information. 

<!-- We should add something here about design/best practices as another next step, like this... -->
To learn more about how to design a world class Word add-in that creates a compelling experience for your users, see [Design guidelines](../design/add-in-design.md) and [Best practices](../design/add-in-development-best-practices.md). 

After you develop your add-in, you can [publish](../publish/publish.md) it to a network share, to an app catalog, or to the Office Store.

## What's coming up for Word add-ins? <!-- Suggest a more clear/enticing heading here. -->

As we design and develop new APIs for Word add-ins, we'll make them available for your feedback on our [Open API specifications](../../reference/openspec.md) page. Find out what new features are in the pipeline for the Word JavaScript APIs, and provide your input on our design specifications.

You can also see what's new in the Word JavaScript API on the [change log](http://dev.office.com/changelog) page.


## Additional resources


* [Office Add-ins platform overview](../overview/office-add-ins.md)

