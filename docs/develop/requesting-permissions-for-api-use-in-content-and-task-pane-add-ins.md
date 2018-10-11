---
title: Requesting permissions for API use in content and task pane add-ins
description: ''
ms.date: 12/04/2017
---


# Requesting permissions for API use in content and task pane add-ins

This article describes the different permission levels that you can declare in your content or task pane add-in's manifest to specify the level of JavaScript API access your add-in requires for its features. 




## Permissions model


A five-level JavaScript API access-permissions model provides the basis for privacy and security for users of your content and task pane add-ins. Figure 1 shows the five levels of API permissions you can declare in your add-in's manifest.


*Figure 1. The five-level permission model for content and task pane add-ins*

![Levels of permissions for task pane apps](../images/office15-app-sdk-task-pane-app-permission.png)



These permissions specify the subset of the API that the add-in runtime will allow your content or task pane add-in to use when a user inserts, and then activates (trusts) your add-in. To declare the permission level your content or task pane add-in requires, specify one of the permission text values in the [Permissions](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/permissions?view=office-js) element of your add-in's manifest. The following example requests the **WriteDocument** permission, which will allow only methods that can write to (but not read) the document.




```XML
<Permissions>WriteDocument</Permissions>
```

As a best practice, you should request permissions based on the principle of  _least privilege_. That is, you should request permission to access only the minimum subset of the API that your add-in requires to function correctly. For example, if your add-in needs only to read data in a user's document for its features, you should request no more than the **ReadDocument** permission.

The following table describes the subset of the JavaScript API that is enabled by each permission level.



|**Permission**|**Enabled subset of the API**|
|:-----|:-----|
|**Restricted**|The methods of the [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) object, and the [Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getactiveviewasync-options--callback-) method.This is the minimum permission level that can be requested by a content or task pane add-in.|
|**ReadDocument**|In addition to the API allowed by the  **Restricted** permission, adds access to the API members necessary to read the document and manage bindings.This includes the use of:<br/><ul><li>The <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-" target="_blank">Document.getSelectedDataAsync</a> method to get the selected text, HTML (Word only), or tabular data, but not the underlying Open Office XML (OOXML) code that contains all of the data in the document.</p></li><li><p>The <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfileasync-filetype--options--callback-" target="_blank">Document.getFileAsync</a> method to get all of the text in the document, but not the underlying OOXML binary copy of the document.</p></li><li><p>The <a href="https://docs.microsoft.com/javascript/api/office/office.binding?view=office-js#getdataasync-options--callback-" target="_blank">Binding.getDataAsync</a> method for reading bound data in the document.</p></li><li><p>The <a href="https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromnameditemasync-itemname--bindingtype--options--callback-" target="_blank">addFromNamedItemAsync</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfrompromptasync-bindingtype--options--callback-" target="_blank">addFromPromptAsync</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-" target="_blank">addFromSelectionAsync</a> methods of the <span class="keyword">Bindings</span> object for creating bindings in the document.</p></li><li><p>The <a href="https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#getallasync-options--callback-" target="_blank">getAllAsync</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#getbyidasync-id--options--callback-" target="_blank">getByIdAsync</a>, and <a href="https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#releasebyidasync-id--options--callback-" target="_blank">releaseByIdAsync</a> methods of the <span class="keyword">Bindings</span> object for accessing and removing bindings in the document.</p></li><li><p>The <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-" target="_blank">Document.getFilePropertiesAsync</a> method to access document file properties, such as the URL of the document.</p></li><li><p>The <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-" target="_blank">Document.goToByIdAsync</a> method to navigate to named objects and locations in the document.</p></li><li><p>For task pane add-ins for Project, all of the "get" methods of the <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js" target="_blank">ProjectDocument</a> object. </p></li></ul>|
|**ReadAllDocument**|In addition to the API allowed by the  **Restricted** and **ReadDocument** permissions, allows the following additional access to document data:<br/><ul><li><p>The <span class="keyword">Document.getSelectedDataAsync</span> and <span class="keyword">Document.getFileAsync</span> methods can access the underlying OOXML code of the document (which in addition to the text may include formatting, links, embedded graphics, comments, revisions, and so forth).</p></li></ul>|
|**WriteDocument**|In addition to the API allowed by the  **Restricted** permission, adds access to the following API members:<br/><ul><li><p>The <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-" target="_blank">Document.setSelectedDataAsync</a> method to write to the user's selection in the document.</p></li></ul>|
|**ReadWriteDocument**|In addition to the API allowed by the  **Restricted**,  **ReadDocument**,  **ReadAllDocument**, and  **WriteDocument** permissions, includes access to all remaining API supported by content and task pane add-ins, including methods for subscribing to events.You must declare the  **ReadWriteDocument** permission to access these additional API members:<br/><ul><li><p>The <a href="https://docs.microsoft.com/javascript/api/office/office.binding?view=office-js#setdataasync-data--options--callback-" target="_blank">Binding.setDataAsync</a> method for writing to bound regions of the document.</p></li><li><p>The <a href="https://docs.microsoft.com/javascript/api/office/office.tablebinding?view=office-js#addrowsasync-rows--options--callback-" target="_blank">TableBinding.addRowsAsync</a> method for adding rows to bound tables.</p></li><li><p>The <a href="https://docs.microsoft.com/javascript/api/office/office.tablebinding?view=office-js#addcolumnsasync-tabledata--options--callback-" target="_blank">TableBinding.addColumnsAsync</a> method for adding columns to bound tables.</p></li><li><p>The <a href="https://docs.microsoft.com/javascript/api/office/office.tablebinding?view=office-js#deletealldatavaluesasync-options--callback-" target="_blank">TableBinding.deleteAllDataValuesAsync</a> method for deleting all data in a bound table.</p></li><li><p>The <a href="https://docs.microsoft.com/javascript/api/office/office.tablebinding?view=office-js#setformatsasync-cellformat--options--callback-" target="_blank">setFormatsAsync</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.tablebinding?view=office-js#clearformatsasync-options--callback-" target="_blank">clearFormatsAsync</a>, and <a href="https://docs.microsoft.com/javascript/api/office/office.tablebinding?view=office-js#settableoptionsasync-tableoptions--options--callback-" target="_blank">setTableOptionsAsync</a> methods of the <span class="keyword">TableBinding</span> object for setting formatting and options on bound tables.</p></li><li><p>All of the members of the <a href="https://docs.microsoft.com/javascript/api/office/office.customxmlnode?view=office-js" target="_blank">CustomXmlNode</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js" target="_blank">CustomXmlPart</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js" target="_blank">CustomXmlParts</a>, and <a href="https://docs.microsoft.com/javascript/api/office/office.customxmlprefixmappings?view=office-js" target="_blank">CustomXmlPrefixMappings</a> objects.</p></li><li><p>All of the methods for subscribing to the events supported by content and task pane add-ins, specifically the <span class="keyword">addHandlerAsync</span> and <span class="keyword">removeHandlerAsync</span> methods of the <a href="https://docs.microsoft.com/javascript/api/office/office.binding?view=office-js" target="_blank">Binding</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js" target="_blank">CustomXmlPart</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js" target="_blank">Document</a>, <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js" target="_blank">ProjectDocument</a>, and <a href="https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#settings" target="_blank">Settings</a> objects.</p></li></ul>|

## See also

- [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md)
    


