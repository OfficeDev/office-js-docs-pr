---
title: Word add-ins overview
description: ''
ms.date: 11/05/2019
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Normal
---


# Word add-ins overview

Do you want to create a solution that extends the functionality of Word? For example, one that involves automated document assembly? Or a solution that binds to and accesses data in a Word document from other data sources? You can use the Office Add-ins platform, which includes the Word JavaScript API and the JavaScript API for Office, to extend Word clients running on a Windows desktop, on a Mac, or in the cloud.

Word add-ins are one of the many development options that you have on the [Office Add-ins platform](../overview/office-add-ins.md). You can use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state. 

> [!NOTE]
> When you build your add-in, if you plan to [publish](../publish/publish.md) your add-in to AppSource, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).

The following figure shows an example of a Word add-in that runs in a task pane.

*Figure 1. Add-in running in a task pane in Word*

![Add-in running in a task pane in Word](../images/word-add-in-show-host-client.png)

The Word add-in (1) can send requests to the Word document (2) and can use JavaScript to access the paragraph object and update, delete, or move the paragraph. For example, the following code shows how to append a new sentence to that paragraph.

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

You can use any web server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework -- Ember, Backbone, Angular, React -- or stick with VanillaJS to develop your solution, and you can use services like Azure to [authenticate](../develop/overview-authn-authz.md) and host your application.

The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can use these APIs to create add-ins that target:

* Word 2013 or later on Windows
* Word on the web
* Word 2016 or later on Mac
* Word on iPad

Write your add-in once, and it will run in all versions of Word across multiple platforms. For details, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).

## JavaScript APIs for Word

You can use two sets of JavaScript APIs to interact with the objects and metadata in a Word document. The first is the [Common API](/javascript/api/office), which was introduced in Office 2013. Many of the objects in the Common API can be used in add-ins hosted by two or more Office clients. This API uses callbacks extensively.

The second is the [Word JavaScript API](/javascript/api/word). This is a strongly-typed object model that you can use to create Word add-ins that target Word 2016 on Mac and Windows. This object model uses promises, and provides access to Word-specific objects like [body](/javascript/api/word/word.body), [content controls](/javascript/api/word/word.contentcontrol), [inline pictures](/javascript/api/word/word.inlinepicture), and [paragraphs](/javascript/api/word/word.paragraph). The Word JavaScript API includes TypeScript definitions and vsdoc files so that you can get code hints in your IDE.

Currently, all Word clients support the shared JavaScript API for Office, and most clients support the Word JavaScript API. For details about supported clients, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).

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

Ready to create your first Word add-in? See [Build your first Word add-in](word-add-ins.md). Use the [add-in manifest](../develop/add-in-manifests.md) to describe where your add-in is hosted, how it is displayed, and define permissions and other information.

To learn more about how to design a world class Word add-in that creates a compelling experience for your users, see [Design guidelines](../design/add-in-design.md) and [Best practices](../concepts/add-in-development-best-practices.md).

After you develop your add-in, you can [publish](../publish/publish.md) it to a network share, an app catalog, or AppSource.

## See also

* [Building Office Add-ins](../overview/office-add-ins-fundamentals.md)
* [Office Add-ins platform overview](../overview/office-add-ins.md)
* [Word JavaScript API reference](/office/dev/add-ins/reference/overview/word-add-ins-reference-overview)
