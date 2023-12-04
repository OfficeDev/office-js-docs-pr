---
title: Word add-ins overview
description: Learn the basics of Word add-ins.
ms.date: 02/24/2023
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Word add-ins overview

Do you want to create a solution that extends the functionality of Word? For example, one that involves automated document assembly? Or a solution that binds to and accesses data in a Word document from other data sources? You can use the Office Add-ins platform, which includes the Word JavaScript API and the Office JavaScript API, to extend Word clients running on a Windows desktop, on a Mac, or in the cloud.

Word add-ins are one of the many development options that you have on the [Office Add-ins platform](../overview/office-add-ins.md). You can use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

The following figure shows an example of a Word add-in that runs in a task pane.

*Figure 1. Add-in running in a task pane in Word*

![Add-in running in a task pane in Word.](../images/word-add-in-show-host-client.png)

The Word add-in can do the following:

  1. Send requests to the Word document.
  1. Use JavaScript to access the paragraph object and update, delete, or move the paragraph.

For example, the following code shows how to append a new sentence to that paragraph.

```js
await Word.run(async (context) => {
    const paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load();
    await context.sync();
    paragraphs.items[0].insertText(' New sentence in the paragraph.',
                                       Word.InsertLocation.end);
    await context.sync();
});

```

You can use any web server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework -- Ember, Backbone, Angular, React -- or stick with vanilla or plain JavaScript to develop your solution. You can also use services like Azure to [authenticate](../develop/overview-authn-authz.md) and host your application.

The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can use these APIs to create add-ins that target:

* Word 2013 or later on Windows
* Word on the web
* Word 2016 or later on Mac
* Word on iPad

Write your add-in once, and it will run in all versions of Word across multiple platforms. For details, see [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets).

## JavaScript APIs for Word

You can use two sets of JavaScript APIs to interact with the objects and metadata in a Word document. The first is the [Common API](/javascript/api/office), which was introduced in Office 2013. Many of the objects in the Common API can be used in add-ins hosted by two or more Office clients. This API uses callbacks extensively.

The second is the [Word JavaScript API](/javascript/api/word). This is an [application-specific API model](../develop/application-specific-api-model.md) that was introduced with Word 2016. It's a strongly-typed object model that you can use to create Word add-ins that target Word 2016 and later on Mac and on Windows. This object model uses promises and provides access to Word-specific objects like [body](/javascript/api/word/word.body), [content controls](/javascript/api/word/word.contentcontrol), [inline pictures](/javascript/api/word/word.inlinepicture), and [paragraphs](/javascript/api/word/word.paragraph). The Word JavaScript API includes TypeScript definitions and vsdoc files so that you can get code hints in your IDE.

Currently, all Word clients support the shared Office JavaScript API, and most clients support the Word JavaScript API. For details about supported clients, see [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets).

We recommend that you start with the Word JavaScript API because the object model is easier to use. Use the Word JavaScript API if you need to do the following:

* Access the objects in a Word document.

Use the shared Office JavaScript API when you need to do any of the following:

* Target Word 2013.
* Perform initial actions for the application.
* Check the supported requirement set.
* Access metadata, settings, and environmental information for the document.
* Bind to sections in a document and capture events.
* Open a dialog box.

## Next steps

Ready to enhance your productivity using Word add-ins? Discover how to [insert AI-generated content into a document](https://github.com/OfficeDev/Word-Scenario-based-Add-in-Samples/tree/main/Word-Add-in-AIGC).

When you're ready to explore creating your own solution, [build your first Word add-in](../quickstarts/word-quickstart.md). Then, learn more about [core concepts](word-add-ins-core-concepts.md) to further extend the functionality of your add-in.

## See also

* [Developing Office Add-ins](../develop/develop-overview.md)
* [Learn about the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)
* [Office Add-ins platform overview](../overview/office-add-ins.md)
* [Word JavaScript API reference](../reference/overview/word-add-ins-reference-overview.md)
