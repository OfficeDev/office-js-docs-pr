---
title: Office Add-ins development lifecycle
description: ''
ms.date: 03/19/2019
localization_priority: Priority
---

# Office Add-ins development lifecycle

> [!NOTE]
> If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)). 

The typical development lifecycle of an Office Add-in includes the following steps:


## 1. Decide on the purpose of the add-in

Ask the following questions:

- How is the add-in useful?

- How does it help your customers be more productive?

- What scenarios does your add-in's features support?

Decide the most important features and scenarios and focus your design around them.


## 2. Identify the data and data source for the add-in

- Is the data in a document, workbook, presentation, project, or an Access browser-based database?

- Is the data about an item or items in an Exchange Server or Exchange Online mailbox?

- Is the data from an external source such as a web service?


## 3. Identify the type of add-in and Office host applications that best support the purpose of the add-in

Consider the following to identify the scenarios:

- Will customers use the add-in to enrich the content of a document or Access browser-based database? If so, you may want to consider creating a **content add-in**.

- Will customers use the add-in while viewing or composing an email message or appointment? Is being able to expose the add-in according to the current context important? Is making the add-in available on not just the desktop, but also on tablets and phones a priority?

	If you answer yes to any of these questions, consider creating an **Outlook add-in**. Identify the context that will trigger your add-in (for example, the user being in a compose form, specific message types, the presence of an attachment, address, task suggestion, or meeting suggestion, or certain string patterns in the contents of an email or appointment). 

	To find out how you can contextually activate the Outlook add-in, see [Activation rules for Outlook add-ins](/outlook/add-ins/activation-rules).

- Will customers use the add-in to enhance the viewing or authoring experience of a document? If so, you may want to consider creating a **task pane add-in**.

Support for certain add-in APIs may differ between Office applications and the platform they are running on (Windows, Mac, Web, Mobile). To see the current API coverage by client and platform, see our [Office Add-in host and platform availability](../overview/office-add-in-availability.md) page.  


## 4. Design and implement the user experience and user interface for the add-in

Design a fast and fluid user experience that is consistent, easy to learn, with primary scenarios that require only a few steps to complete. Depending on the purpose of the add-in, make use of third-party APIs or web services.

You can choose from a variety of web development tools and use HTML and JavaScript to implement the user interface.


## 5. Create an XML manifest file based on the Office Add-ins manifest schema

Create an XML manifest to identify the add-in and its requirements, specify the locations of the HTML and any JavaScript and CSS files that the add-in uses, and depending on the type of the add-in, the default size and permissions.

For Outlook add-ins, you can specify the context, based on the current message or appointment, under which your add-in is relevant and you would like Outlook to make available in the UI. You can also decide which devices you want the add-in to support. In the manifest, specify the context as activation rules and the supported devices.


## 6. Install and test the add-in

Place the HTML files and any JavaScript and CSS files on the web servers that are specified in the add-in manifest file. The process to install an add-in depends on the type of the add-in. For details, see [Sideload Office Add-ins for testing](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

For Outlook add-ins, install it in an Exchange mailbox, and specify the location of the add-in manifest file in the Exchange Admin Center (EAC). For more information, see [Deploy and install Outlook add-ins for testing](/outlook/add-ins/testing-and-tips).


## 7. Publish the add-in

You can submit the add-in to AppSource, from which customers can install the add-in. In addition, you can publish task pane and content add-ins to an app catalog on SharePoint or to a shared network folder, and you can deploy an Outlook add-in directly on an Exchange server for your organization. For details, see [Publish your Office Add-in](../publish/publish.md).


## 8. Maintain the add-in

If your add-in calls a web service, and if you make updates to the web service after publishing the add-in, you do not have to republish the add-in. However, if you change any items or data you submitted for your add-in, such as the add-in manifest, screenshots, icons, HTML or JavaScript files, you will need to republish the add-in. 

In particular, if you have published the add-in to AppSource, you'll need to resubmit your add-in so that AppSource can implement those changes. You must resubmit your add-in with an updated add-in manifest that includes a new version number. You must also make sure to update the add-in version number in the submission form to match the new manifest's version number. For Outlook add-ins, you should make sure the [Id](/office/dev/add-ins/reference/manifest/id) element contains a different UUID in the add-in manifest.
