
# Word JavaScript add-in development overview

<!-- I added JavaScript to the H1 because we need to differentiate from the older add-in model. -->

Does your solution involve automated document assembly? Do you want to bind and access data in a Word document from other data sources? Do you want to create new tools for Word -- to make Word do things that it doesn't do out of the box? If so, the Word JavaScript add-in development model is the choice for building cross platform extensions to Word client applications.

Word add-ins are one of the many development options that you have on the [Office Add-ins platform](../overview/office-add-ins.md). You can extend the Word UI with [add-in commands](../design/add-in-commands.md) and task panes that can run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Word and Word add-ins have a host-client relationship. An add-in that interacts with content in a Word document takes on a client role by creating requests to act on Word objects and synchronize object state between the Word document and the add-in. Let's look at the following figure that shows a task pane loaded into Word.

**Figure 1. Word host and the task pane client**

![Word host and task pane](../../images/WordAddinShowHostClient.png)

The add-in loaded into a Word task pane (1), can send requests to the Word host (2). There is a paragraph object in the Word document that the add-in can access. That paragraph can be updated, deleted, or moved by running JavaScript in the Word task pane. For example, the following code shows how to append a new sentence to that paragraph.

```js
Word.run(function (context) {
    var paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load();
    return context.sync().then(function () {
        paragraphs.items[0].insertText(' New sentence in the paragraph.',
                                       Word.InsertLocation.end);
    }).then(context.sync);
});

```

You can use any Web server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework -- Ember, Backbone, Angular, React -- or stick with VanillaJS to develop your solution and you can use services like Azure to [authenticate](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md) and host your application.

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

Currently, all Word clients support the shared JavaScript API for Office, and most clients support the Word JavaScript API. For details about supported clients, see the [API reference documentation](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word).

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


## Next steps

<!-- We should think about providing more clear next steps instead of lumping links together in an Additional resources section. -->

Ready to create your first Word add-in? See [Build your first Word add-in](word-add-ins.md). You can also try our interactive [Get started experience](http://dev.office.com/getting-started/addins?product=Word). Use the [add-in manifest](../overview/add-in-manifests.md) to describe where your add-in is hosted and how it is displayed, and define permissions and other information.

<!-- We should add something here about design/best practices as another next step, like this... -->
To learn more about how to design a world class Word add-in that creates a compelling experience for your users, see [Design guidelines](../design/add-in-design.md) and [Best practices](../design/add-in-development-best-practices.md).

After you develop your add-in, you can [publish](../publish/publish.md) it to a network share, to an app catalog, or to the Office Store.

## What's coming up for Word add-ins?

As we design and develop new APIs for Word add-ins, we'll make them available for your feedback on our [Open API specifications](../../reference/openspec.md) page. Find out what new features are in the pipeline for the Word JavaScript APIs, and provide your input on our design specifications.

You can also see what's new in the Word JavaScript API on the [change log](http://dev.office.com/changelog) page.

## Additional resources

* [Office Add-ins platform overview](../overview/office-add-ins.md)
* [Word JavaScript API reference](../../reference/word/word-add-ins-reference-overview.md)

