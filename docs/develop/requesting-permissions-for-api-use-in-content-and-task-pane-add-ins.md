
# Requesting permissions for API use in content and task pane add-ins
This article describes the different permission levels that you can declare in your content or task pane add-in's manifest to specify the level of JavaScript API access your add-in requires for its features. 




## Permissions model


A five-level JavaScript API access-permissions model provides the basis for privacy and security for users of your content and task pane add-ins. Figure 1 shows the five levels of API permissions you can declare in your add-in's manifest.


**Figure 1. The five-level permission model for content and task pane add-ins**

![Levels of permissions for task pane apps](../../images/off15appsdk_TaskPaneAppPermission.gif)



These permissions specify the subset of the API that the add-in runtime will allow your content or task pane add-in to use when a user inserts, and then activates (trusts) your add-in. To declare the permission level your content or task pane add-in requires, specify one of the permission text values in the [Permissions](http://msdn.microsoft.com/en-us/library/d4cfe645-353d-8240-8495-f76fb36602fe%28Office.15%29.aspx) element of your add-in's manifest. The following example requests the **WriteDocument** permission, which will allow only methods that can write to (but not read) the document.




```XML
<Permissions>WriteDocument</Permissions>
```

As a best practice, you should request permissions based on the principle of  _least privilege_. That is, you should request permission to access only the minimum subset of the API that your add-in requires to function correctly. For example, if your add-in needs only to read data in a user's document for its features, you should request no more than the **ReadDocument** permission.

The following table describes the subset of the JavaScript API that is enabled by each permission level.



|**Permission**|**Enabled subset of the API**|
|:-----|:-----|
|**Restricted**|The methods of the [Settings](../../reference/shared/settings.md) object, and the [Document.getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md) method.This is the minimum permission level that can be requested by a content or task pane add-in.|
|**ReadDocument**|In addition to the API allowed by the  **Restricted** permission, adds access to the API members necessary to read the document and manage bindings.This includes the use of:<br/><ul><li>The <a href="http://msdn.microsoft.com/en-us/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69(Office.15).aspx" target="_blank">Document.getSelectedDataAsync</a> method to get the selected text, HTML (Word only), or tabular data, but not the underlying Open Office XML (OOXML) code that contains all of the data in the document.</p></li><li><p>The <a href="http://msdn.microsoft.com/en-us/library/78047418-89c4-4c7d-9427-4735b8559518(Office.15).aspx" target="_blank">Document.getFileAsync</a> method to get all of the text in the document, but not the underlying OOXML binary copy of the document.</p></li><li><p>The <a href="http://msdn.microsoft.com/en-us/library/5372ffd8-579d-4fcb-9e5b-e9a2128f3201(Office.15).aspx" target="_blank">Binding.getDataAsync</a> method for reading bound data in the document.</p></li><li><p>The <a href="http://msdn.microsoft.com/en-us/library/afbadac7-60c7-47cb-9477-6e9466ded44c(Office.15).aspx" target="_blank">addFromNamedItemAsync</a>, <a href="http://msdn.microsoft.com/en-us/library/9dc03608-b08b-4700-8be1-3c86ae236799(Office.15).aspx" target="_blank">addFromPromptAsync</a>, <a href="http://msdn.microsoft.com/en-us/library/edc99214-e63e-43f2-9392-97ead42fc155(Office.15).aspx" target="_blank">addFromSelectionAsync</a> methods of the <span class="keyword">Bindings</span> object for creating bindings in the document.</p></li><li><p>The <a href="http://msdn.microsoft.com/en-us/library/ef902b73-cc4c-4551-95de-d8a51eeba82f(Office.15).aspx" target="_blank">getAllAsync</a>, <a href="http://msdn.microsoft.com/en-us/library/2727c891-bc05-465c-9324-113fbfeb3fbb(Office.15).aspx" target="_blank">getByIdAsync</a>, and <a href="http://msdn.microsoft.com/en-us/library/ad285984-8b44-435d-9b84-f0ade570c896(Office.15).aspx" target="_blank">releaseByIdAsync</a> methods of the <span class="keyword">Bindings</span> object for accessing and removing bindings in the document.</p></li><li><p>The <a href="http://msdn.microsoft.com/en-us/library/2533a563-95ae-4d52-b2d5-a6783e4ef5b4(Office.15).aspx" target="_blank">Document.getFilePropertiesAsync</a> method to access document file properties, such as the URL of the document.</p></li><li><p>The <a href="http://msdn.microsoft.com/en-us/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">Document.goToByIdAsync</a> method to navigate to named objects and locations in the document.</p></li><li><p>For task pane add-ins for Project, all of the "get" methods of the <a href="http://msdn.microsoft.com/en-us/library/1908af4f-93b9-4859-87e3-06942014fae1(Office.15).aspx" target="_blank">ProjectDocument</a> object. </p></li></ul>|
|**ReadAllDocument**|In addition to the API allowed by the  **Restricted** and **ReadDocument** permissions, allows the following additional access to document data:<br/><ul><li><p>The <span class="keyword">Document.getSelectedDataAsync</span> and <span class="keyword">Document.getFileAsync</span> methods can access the underlying OOXML code of the document (which in addition to the text may include formatting, links, embedded graphics, comments, revisions, and so forth).</p></li></ul>|
|**WriteDocument**|In addition to the API allowed by the  **Restricted** permission, adds access to the following API members:<br/><ul><li><p>The <a href="http://msdn.microsoft.com/en-us/library/998f38dc-83bd-4659-a759-4758c632a6ef(Office.15).aspx" target="_blank">Document.setSelectedDataAsync</a> method to write to the user's selection in the document.</p></li></ul>|
|**ReadWriteDocument**|In addition to the API allowed by the  **Restricted**,  **ReadDocument**,  **ReadAllDocument**, and  **WriteDocument** permissions, includes access to all remaining API supported by content and task pane add-ins, including methods for subscribing to events.You must declare the  **ReadWriteDocument** permission to access these additional API members:<br/><ul><li><p>The <a href="http://msdn.microsoft.com/en-us/library/6a59bb6d-40b6-4a95-9b98-d70d4616de09(Office.15).aspx" target="_blank">Binding.setDataAsync</a> method for writing to bound regions of the document.</p></li><li><p>The <a href="http://msdn.microsoft.com/en-us/library/1cd23454-8435-4e13-98b3-d0d29ed278a8(Office.15).aspx" target="_blank">TableBinding.addRowsAsync</a> method for adding rows to bound tables.</p></li><li><p>The <a href="http://msdn.microsoft.com/en-us/library/8f1bfa81-3850-4ea1-ba2e-c9bcf5847a44(Office.15).aspx" target="_blank">TableBinding.addColumnsAsync</a> method for adding columns to bound tables.</p></li><li><p>The <a href="http://msdn.microsoft.com/en-us/library/8f5cc783-384d-4520-a218-190dfed74dd2(Office.15).aspx" target="_blank">TableBinding.deleteAllDataValuesAsync</a> method for deleting all data in a bound table.</p></li><li><p>The <a href="http://msdn.microsoft.com/en-us/library/49712906-f582-4055-9ef8-6edde6e97679(Office.15).aspx" target="_blank">setFormatsAsync</a>, <a href="http://msdn.microsoft.com/en-us/library/cc56e9c0-b33c-4d9b-b676-a7e50f757c10(Office.15).aspx" target="_blank">clearFormatsAsync</a>, and <a href="http://msdn.microsoft.com/en-us/library/2885fc57-4527-4ca4-a43d-9ee447ec27d3(Office.15).aspx" target="_blank">setTableOptionsAsync</a> methods of the <span class="keyword">TableBinding</span> object for setting formatting and options on bound tables.</p></li><li><p>All of the members of the <a href="http://msdn.microsoft.com/en-us/library/dc1518de-47fa-4108-aab7-04a022724b04(Office.15).aspx" target="_blank">CustomXmlNode</a>, <a href="http://msdn.microsoft.com/en-us/library/83f0e668-8236-4f2f-a20f-b173a9e3f65f(Office.15).aspx" target="_blank">CustomXmlPart</a>, <a href="http://msdn.microsoft.com/en-us/library/ba40cd4c-29bb-4f31-875d-6f1382fd1ee8(Office.15).aspx" target="_blank">CustomXmlParts</a>, and <a href="http://msdn.microsoft.com/en-us/library/18b9aa8c-83e7-4c2f-8530-6a0ac8ce5535(Office.15).aspx" target="_blank">CustomXmlPrefixMappings</a> objects.</p></li><li><p>All of the methods for subscribing to the events supported by content and task pane add-ins, specifically the <span class="keyword">addHandlerAsync</span> and <span class="keyword">removeHandlerAsync</span> methods of the <a href="http://msdn.microsoft.com/en-us/library/42882642-d22b-47d2-a8d3-3aa8c6a4435e(Office.15).aspx" target="_blank">Binding</a>, <a href="http://msdn.microsoft.com/en-us/library/83f0e668-8236-4f2f-a20f-b173a9e3f65f(Office.15).aspx" target="_blank">CustomXmlPart</a>, <a href="http://msdn.microsoft.com/en-us/library/f8859516-cc1f-4b20-a8f3-cee37a983e70(Office.15).aspx" target="_blank">Document</a>, <a href="http://msdn.microsoft.com/en-us/library/1908af4f-93b9-4859-87e3-06942014fae1(Office.15).aspx" target="_blank">ProjectDocument</a>, and <a href="http://msdn.microsoft.com/en-us/library/ad733387-a58c-4514-8fc2-53e64fad468d(Office.15).aspx" target="_blank">Settings</a> objects.</p></li></ul>|

## Additional resources

    
- [Privacy and security for Office Add-ins](../../docs/develop/privacy-and-security.md)
    


