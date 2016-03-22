
# Outlook add-in APIs

To use APIs in your Outlook add-in, you must specify the location of the Office.js library, the requirement set, the schema, and the permissions.

## Office.js

To interact with the Outlook add-in API, developers must use the JavaScript APIs in Office.js. The CDN for the library is  _https://appsforoffice.microsoft.com/lib/1/hosted/Office.js_. Add-ins submitted to the store must reference Office.js by this CDN; they can't use a local reference. 

Declare the CDN in the **[head]** tag of the web page (.html, .aspx, or .php file) that implements the UI of your add-in, in the **[src]** attribute of the **[script]** tag:


```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

As we add new APIs, the URL to Office.js will stay the same. We will change the version in the URL only if we break an existing API behavior.

The library must load within 5 seconds of the add-in being launched, or else Outlook will determine the page is unresponsive, and will show an error dialog.


## Requirement sets

All Outlook APIs belong to the Mailbox requirement set. The Mailbox requirement set has versions, and each new set of APIs that we release belongs to a higher version of the set. Not all Outlook clients will support the newest set of APIs when we release them, but if an Outlook client declares support for a requirement set, it will support all the APIs in that requirement set. 

Specifying a minimum requirement set version in the manifest controls which Outlook clients the add-in will appear in. For example, if requirement set version 1.3 is specified, this means the add-in will not show up in any Outlook client that doesn't support at least v1.3. 

However, specifying a requirement set doesn't limit your add-in to the APIs in that version. If the add-in specifies requirement set v1.1 but is running in an Outlook client that supports v1.3, the add-in can still use these new APIs. The requirement set only controls which Outlook clients the add-in appears in.

To check availability of any APIs from a requirement set greater than the one specified in the manifest, you can use standard JavaScript:




```js
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction...  
}
```

> **Note:** No such checks are necessary for any APIs that are in the requirement set version specified in the manifest.

Developers should specify the minimum requirement set that supports the critical set of APIs for their scenario, without which the critical features of the add-in won't work. You specify the requirement set in the manifest in the  **Requirements**, **Sets**, and **Set** elements. For more information, see [Outlook add-in manifests](../outlook/manifests/manifests.md).

The  **Methods** element doesn't apply to mail add-ins, so you can't declare support for specific methods.


## Permissions

Your add-in requires the appropriate permissions to use the APIs that it needs. There are four levels of permissions, summarized below. For full details, see [Understanding Outlook add-in permissions](../outlook/understanding-outlook-add-in-permissions.md).


|**Permission level**|**Description**|
|:-----|:-----|
|Restricted|Allows use of entities but not regular expressions.|
|Read item|In addition to what is allowed in  _Restricted_, it allows:<ul><li>regular expressions</li><li>Outlook add-in API read access</li><li>getting the item properties and the callback token</li></ul>|
|Read/write|In addition to what is allowed in  _Read item_, it allows:<ul><li>full Outlook add-in API access except <b>makeEwsRequestAsync</b></li><li>setting the item properties</li></ul>|
|Read/write mailbox|In addition to what is allowed in  _Read/write_, it allows:<ul><li>creating, reading, writing items and folders</li><li>sending items</li><li>calling [makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md#makeewsrequestasyncdata-callback-usercontext)</li></ul>|
In general, you should specify the minimum permission needed for your add-in. Permissions are declared in the  **Permissions** element in the manifest. For more information, see [Outlook add-in manifests](../outlook/manifests/manifests.md). For information about security issues, see [Privacy, permissions, and security for Outlook add-ins](../outlook/../../docs/develop/privacy-and-security.md).


## Additional resources



- [Outlook Add-in API](../../reference/outlook/index.md)
    
- [Outlook add-in manifests](../outlook/manifests/manifests.md)
    
- [Privacy, permissions, and security for Outlook add-ins](../outlook/../../docs/develop/privacy-and-security.md)
    
