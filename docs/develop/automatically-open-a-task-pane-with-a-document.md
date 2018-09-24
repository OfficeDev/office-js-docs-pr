---
title: Automatically open a task pane with a document
description: ''
ms.date: 05/02/2018
---


# Automatically open a task pane with a document

You can use add-in commands in your Office Add-in to extend the Office UI by adding buttons to the Office ribbon. When users click your command button, an action occurs, such as opening a task pane. 

Some scenarios require that a task pane open automatically when a document opens, without explicit user interaction. You can use the autoopen taskpane feature, introduced in the AddInCommands 1.1 requirement set, to automatically open a task pane when your scenario requires it. 


## How is the autoopen feature different from inserting a task pane? 

When a user launches add-ins that don't use add-in commands - for example, add-ins that run in Office 2013 - they are inserted into the document, and persist in that document. As a result, when other users open the document, they are prompted to install the add-in, and the task pane opens. The challenge with this model is that in many cases, users don’t want the add-in to persist in the document. For example, a student who uses a dictionary add-in in a Word document might not want their classmates or teachers to be prompted to install that add-in when they open the document.  

With the autoopen feature, you can explicitly define or allow the user to define whether a specific task pane add-in persists in a specific document. 

## Support and availability
The autoopen feature is currently <!-- in **developer preview** and it is only --> supported in the following products and platforms.

|**Products**|**Platforms**|
|:-----------|:------------|
|<ul><li>Word</li><li>Excel</li><li>PowerPoint</li></ul>|Supported platforms for all products:<ul><li>Office for Windows Desktop. Build 16.0.8121.1000+</li><li>Office for Mac. Build 15.34.17051500+</li><li>Office Online</li></ul>|


## Best practices

Apply the following best practices when you use the autoopen feature:

- Use the autoopen feature when it will help make your add-in users more efficient, such as:
	- When the document needs the add-in in order to function properly. For example, a spreadsheet that includes stock values that are periodically refreshed by an add-in. The add-in should open automatically when the spreadsheet is opened to keep the values up to date. 
	- When the user will most likely always use the add-in with a particular document. For example, an add-in that helps users fill in or change data in a document by pulling information from a backend system. 
- Allow users to turn on or turn off the autoopen feature. Include an option in your UI for users to choose to no longer automatically open the add-in task pane.  
- Use requirement set detection to determine whether the autoopen feature is available, and provide a fallback behavior if it isn’t.
- Don't use the autoopen feature to artificially increase usage of your add-in. If it doesn’t make sense for your add-in to open automatically with certain documents, this feature can annoy users. 

    > [!NOTE]
    > If Microsoft detects abuse of the autoopen feature, your add-in might be rejected from AppSource. 

- Don't use this feature to pin multiple task panes. You can only set one pane of your add-in to open automatically with a document.  

## Implementation
To implement the autoopen feature:

- Specify the task pane to be opened automatically.
- Tag the document to automatically open the task pane.

> [!IMPORTANT]
> The pane that you designate to open automatically will only open if the add-in is already installed on the user's device. If the user does not have the add-in installed when they open a document, the autoopen feature will not work and the setting will be ignored. If you also require the add-in to be distributed with the document you need to set the visibility property to 1; this can only be done using OpenXML, an example is provided later in this article. 

### Step 1: Specify the task pane to open
To specify the task pane to open automatically, set the [TaskpaneId](https://docs.microsoft.com/javascript/office/manifest/action?view=office-js#taskpaneid) value to **Office.AutoShowTaskpaneWithDocument**. You can only set this value on one task pane. If you set this value on multiple task panes, the first occurrence of the value will be recognized and the others will be ignored. 

The following example shows the TaskPaneId value set to Office.AutoShowTaskpaneWithDocument.
          
```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```     

### Step 2: Tag the document to automatically open the task pane

You can tag the document to trigger the autoopen feature in one of two ways. Pick the alternative that works best for your scenario.  


#### Tag the document on the client side
Use the Office.js [settings.set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) method to set **Office.AutoShowTaskpaneWithDocument** to **true**, as shown in the following example.   

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

Use this method if you need to tag the document as part of your add-in interaction (for example, as soon as the user creates a binding, or chooses an option to indicate that they want the pane to open automatically).

#### Use Open XML to tag the document
You can use Open XML to create or modify a document and add the appropriate Open Office XML markup to trigger the autoopen feature. For a sample that shows you how to do this, see [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin). 

Add two Open XML parts to the document:

- A webextension part
- A task pane part

The following example shows how to add the webextension part.

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="[ADD-IN ID PER MANIFEST]">
  <we:reference id="[GUID or AppSource asset ID]" version="[your add-in version]" store="[Pointer to store or catalog]" storeType="[Store or catalog type]"/>
  <we:alternateReferences/>
  <we:properties>
 	<we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

The webextension part includes a property bag and a property named **Office.AutoShowTaskpaneWithDocument** that must be set to `true`.

The webextension part also includes a reference to the store or catalog with attributes for `id`, `storeType`, `store`, and `version`. Of the `storeType` values, only four are relevant to the autoopen feature. The values for the other three attributes depend on the value for `storeType`, as shown in the following table. 

| **`storeType` value** | **`id` value**	|**`store` value** | **`version` value**|
|:---------------|:---------------|:---------------|:---------------|
|OMEX (AppSource)|The AppSource asset ID of the add-in (see Note)|The locale of AppSource; for example, "en-us".|The version in the AppSource catalog (see Note)|
|FileSystem (a network share)|The GUID of the add-in in the add-in manifest.|The path of the network share; for example, "\\\\MyComputer\\MySharedFolder".|The version in the add-in manifest.|
|EXCatalog (deployment via the Exchange server) |The GUID of the add-in in the add-in manifest.|"EXCatalog". EXCatalog row is the row to use with add-ins that use Centralized Deployment in the Office 365 admin center.|The version in the add-in manifest.
|Registry (System registry)|The GUID of the add-in in the add-in manifest.|"developer"|The version in the add-in manifest.|

> [!NOTE]
> To find the asset ID and version of an add-in in AppSource, go to the AppSource landing page for the add-in. The asset ID appears in the address bar in the browser. The version is listed in the **Details** section of the page.

For more information about the webextension markup, see [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx).

The following example shows how to add the taskpane part.

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

Note that in this example, the `visibility` attribute is set to "0". This means that after the webextension and taskpane parts are added, the first time the document is opened, the user has to install the add-in from the **Add-in** button on the ribbon. Thereafter, the add-in task pane opens automatically when the file is opened. Also, when you set `visibility` to "0", you can use Office.js to enable users to turn on or turn off the autoopen feature. Specifically, your script sets the **Office.AutoShowTaskpaneWithDocument** document setting to `true` or `false`. (For details, see [Tag the document on the client side](#tag-the-document-on-the-client-side).) 

If `visibility` is set to "1", the task pane opens automatically the first time the document is opened. The user is prompted to trust the add-in, and when trust is granted, the add-in opens. Thereafter, the add-in task pane opens automatically when the file is opened. However, when `visibility` is set to "1", you can't use Office.js to enable users to turn on or turn off the autoopen feature. 

Setting `visibility` to "1" is a good choice when the add-in and the template or content of the document are so closely integrated that the user would not opt out of the autoopen feature. 

> [!NOTE]
> If you want to distribute your add-in with the document, so that users are prompted to install it, you must set the visibility property to 1. You can only do this via Open XML.

An easy way to write the XML is to first run your add-in and [tag the document on the client side](#tag-the-document-on-the-client-side) to write the value, and then save the document and inspect the XML that is generated. Office will detect and provide the appropriate attribute values. You can also use the [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) tool to generate C# code to programmatically add the markup based on the XML you generate.

## Test and verify opening taskpanes
You can deploy a test version of your add-in that will automatically open a taskpane using Centralized Deployment via the Office 365 admin center. The following example shows how add-ins are inserted from the Centralized Deployment catalog using the EXCatalog store version.

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```
To test the previous example, please consider joining the [Office 365 Developer Program](https://docs.microsoft.com/office/developer-program/office-365-developer-program) and signing up for an [Office 365 developer account](https://developer.microsoft.com/office/dev-program) if you don't already own an Office 365 subscription. You can actually test drive Centralized Deployment and verify that your add-in works as expected.


## See also

For a sample that shows you how to use the autoopen feature, see [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane). 
[Join the Office 365 developer program](https://docs.microsoft.com/office/developer-program/office-365-developer-program). 

