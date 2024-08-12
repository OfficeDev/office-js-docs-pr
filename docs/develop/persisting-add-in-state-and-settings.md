---
title: Persist add-in state and settings
description: Learn how to persist data in Office Web Add-in applications running in the stateless environment of a browser control.
ms.date: 08/07/2024
ms.localizationpriority: medium
---

# Persist add-in state and settings

Office Add-ins are essentially web applications running in the stateless environment of a browser iframe or a webview control. (For brevity hereafter, this article uses "browser control" to mean "browser or webview control".) When in use, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:

- [Use techniques provided by the underlying browser control](#browser-storage).
- [Use the application-specific Office JavaScript APIs for Excel, Word, and Outlook that store data](#application-specific-settings-and-persistence).

If you need to persist state across documents, such as tracking user preferences across any documents they open, you'll need to use a different approach. For example, you could use [SSO](use-sso-to-get-office-signed-in-user-token.md) to obtain the user identity, and then save the user ID and their settings to an online database.

## Browser storage

Persist data across add-in instances with tools from the underlying browser control, such as browser cookies or HTML5 web storage ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).

Some browsers or the user's browser settings may block browser-based storage techniques. You should test for availability as documented in [Using the Web Storage API](https://developer.mozilla.org/docs/Web/API/Web_Storage_API/Using_the_Web_Storage_API).

### Storage partitioning

As a best practice, any private data should be stored in partitioned `localStorage`. [Office.context.partitionKey](/javascript/api/office/office.context#office-office-context-partitionkey-member) provides a key for use with local storage. This ensures that data stored in local storage is only available in the same context. The following example shows how to use the partition key with `localStorage`. Note that the partition key is undefined in environments without partitioning, such as the browser controls for Windows applications.

```js
// Store the value "Hello" in local storage with the key "myKey1".
setInLocalStorage("myKey1", "Hello");

// ... 

// Retrieve the value stored in local storage under the key "myKey1".
const message = getFromLocalStorage("myKey1");
console.log(message);

// ...

function setInLocalStorage(key: string, value: string) {
  const myPartitionKey = Office.context.partitionKey;

  // Check if local storage is partitioned. 
  // If so, use the partition to ensure the data is only accessible by your add-in.
  if (myPartitionKey) {
    localStorage.setItem(myPartitionKey + key, value);
  } else {
    localStorage.setItem(key, value);
  }
}

function getFromLocalStorage(key: string) {
  const myPartitionKey = Office.context.partitionKey;

  // Check if local storage is partitioned.
  if (myPartitionKey) {
    return localStorage.getItem(myPartitionKey + key);
  } else {
    return localStorage.getItem(key);
  }
}
```

Starting in Version 115 of Chromium-based browsers, such as Chrome and Edge, [storage partitioning](https://developer.chrome.com/docs/privacy-sandbox/storage-partitioning/) is enabled to prevent specific side-channel cross-site tracking (see also [Microsoft Edge browser policies](/deployedge/microsoft-edge-policies#defaultthirdpartystoragepartitioningsetting)). Similar to the Office key-based partitioning, data stored by storage APIs, such as local storage, is only available to contexts with the same origin and the same top-level site.

## Application-specific settings and persistence

Excel, Word, and Outlook provide application-specific APIs to save settings and other data. Use these instead of the [Common APIs mentioned later in this article](#common-api-settings-and-persistence) so that your add-in follows consistent patterns and is optimized for the targeted application.

### Settings in Excel and Word

The application-specific JavaScript APIs for Excel and for Word also provide access to the custom settings. Settings are unique to a single Excel file and add-in pairing. For more information, see [Excel.SettingCollection](/javascript/api/excel/excel.settingcollection) and [Word.SettingCollection](/javascript/api/word/word.settingcollection).

The following example shows how to create and access a setting in Excel. The process is functionally equivalent in Word, which uses [Document.settings](/javascript/api/word/word.document#word-word-document-settings-member) instead of `Workbook.settings`.

```js
await Excel.run(async (context) => {
    const settings = context.workbook.settings;
    settings.add("NeedsReview", true);
    const needsReview = settings.getItem("NeedsReview");
    needsReview.load("value");

    await context.sync();
    console.log("Workbook needs review : " + needsReview.value);
});
```

#### Custom XML data in Excel and Word

The Open XML **.xlsx** and **.docx** file formats let your add-in embed custom XML data in the Excel workbook or Word document. This data persists with the file, independent of the add-in.

A [Word.Document](/javascript/api/word/word.document#word-word-document-customxmlparts-member) and [Excel.Workbook](/javascript/api/excel/excel.workbook#excel-excel-workbook-customxmlparts-member) contain a `CustomXmlPartCollection`, which is a list of `CustomXmlParts`. These give access to the XML strings and a corresponding unique ID. By storing these IDs as settings, your add-in can maintain the keys to its XML parts between sessions.

The following samples show how to use custom XML parts with an Excel workbook. The first code block demonstrates how to embed XML data. It stores a list of reviewers, then uses the workbook's settings to save the XML's `id` for future retrieval. The second block shows how to access that XML later. The "ContosoReviewXmlPartId" setting is loaded and passed to the workbook's `customXmlParts`. The XML data is then printed to the console. The process is functionally equivalent in Word, which uses [Document.customXmlParts](/javascript/api/word/word.document#word-word-document-customxmlparts-member) instead of `Workbook.customXmlParts`.

```js
await Excel.run(async (context) => {
    // Add reviewer data to the document as XML
    const originalXml = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    const customXmlPart = context.workbook.customXmlParts.add(originalXml);
    customXmlPart.load("id");
    await context.sync();

    // Store the XML part's ID in a setting
    const settings = context.workbook.settings;
    settings.add("ContosoReviewXmlPartId", customXmlPart.id);
});
```

> [!NOTE]
> `CustomXMLPart.namespaceUri` is only populated if the top-level custom XML element contains the `xmlns` attribute.

#### Custom properties in Excel and Word

The [Excel.DocumentProperties.custom](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-custom-member) and [Word.DocumentProperties.customProperties](/javascript/api/word/word.documentproperties#word-word-documentproperties-customproperties-member) properties represent collections of key-value pairs for user-defined properties. The following Excel example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.

```js
await Excel.run(async (context) => {
    const customDocProperties = context.workbook.properties.custom;
    customDocProperties.add("Introduction", "Hello");
    await context.sync();
});

// ...

await Excel.run(async (context) => {
    const customDocProperties = context.workbook.properties.custom;
    const customProperty = customDocProperties.getItem("Introduction");
    customProperty.load(["key", "value"]);
    await context.sync();

    console.log("Custom key  : " + customProperty.key); // "Introduction"
    console.log("Custom value : " + customProperty.value); // "Hello"
});
```

> [!TIP]
> In Excel, custom properties can also be set at the worksheet level with the [Worksheet.customProperties](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-customproperties-member) property. These are similar to document-level custom properties, except that the same key can be repeated across different worksheets.

### How to save settings in an Outlook add-in

For information about how to save settings in an Outlook add-in, see [Get and set add-in metadata for an Outlook add-in](../outlook/metadata-for-an-outlook-add-in.md) and [Get and set internet headers on a message in an Outlook add-in](../outlook/internet-headers.md).

## Common API settings and persistence

The [Common APIs](understanding-the-javascript-api-for-office.md#api-models) provide objects to save add-in state across sessions. The saved settings values are associated with the [Id](/javascript/api/manifest/id) of the add-in that created them. Internally, the data accessed with the `Settings`, `CustomProperties`, or `RoamingSettings` objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs. The name (key) for each value must be a `string`, and the stored value can be a JavaScript `string`, `number`, `date`, or `object`, but not a **function**.

This example of the property bag structure contains three defined **string** values named `firstName`,  `location`, and  `defaultView`.

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session. During the session, the settings are managed in entirely in memory using the `get`, `set`, and `remove` methods of the object that corresponds to the kind of settings you're creating (**Settings**, **CustomProperties**, or **RoamingSettings**).

> [!IMPORTANT]
> To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the `saveAsync` method of the corresponding object used to work with that kind of settings. The `get`, `set`, and `remove` methods operate only on the in-memory copy of the settings property bag. If your add-in is closed without calling `saveAsync`, any changes made to settings during that session will be lost.

### How to save add-in state and settings per document for content and task pane add-ins

To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, use the [Settings](/javascript/api/office/office.settings) object and its methods. The property bag created with the methods of the `Settings` object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.

The `Settings` object is automatically loaded as part of the [Document](/javascript/api/office/office.document) object, and is available when the task pane or content add-in is activated. After the `Document` object is instantiated, you can access the `Settings` object with the [settings](/javascript/api/office/office.document#office-office-document-settings-member) property of the `Document` object. During the lifetime of the session, you can use the `Settings.get`, `Settings.set`, and `Settings.remove` methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.

Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with, you must call the [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)) method.

#### Create or update a setting value

The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#office-office-settings-set-member(1)) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  *name* (Id) of the setting to set or create. The second parameter is the *value* of the setting.

```js
Office.context.document.settings.set('themeColor', 'green');
```

The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist. Use the `Settings.saveAsync` method to persist the new or updated settings to the document.

#### Get the value of a setting

The following example shows how use the [Settings.get](/javascript/api/office/office.settings#office-office-settings-get-member(1)) method to get the value of a setting called "themeColor". The only parameter of the `get` method is the case-sensitive *name* of the setting.

```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

The `get` method returns the value that was previously saved for the setting *name* that was passed in. If the setting doesn't exist, the method returns **null**.

#### Remove a setting

The following example shows how to use the [Settings.remove](/javascript/api/office/office.settings#office-office-settings-remove-member(1)) method to remove a setting with the name "themeColor". The only parameter of the `remove` method is the case-sensitive *name* of the setting.

```js
Office.context.document.settings.remove('themeColor');
```

Nothing will happen if the setting doesn't exist. Use the `Settings.saveAsync` method to persist removal of the setting from the document.

#### Save your settings

To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)) method to store them in the document. The only parameter of the `saveAsync` method is *callback*, which is a callback function with a single parameter.

```js
Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        write('Settings saved.');
    }
});
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

The anonymous function passed into the `saveAsync` method as the *callback* parameter is executed when the operation is completed. The *asyncResult* parameter of the callback provides access to an `AsyncResult` object that contains the status of the operation. In the example, the function checks the `AsyncResult.status` property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.

### How to save custom XML to the document

A custom XML part is an available storage option for when you want to store information that has a structured character or need the data to be accessible across instances of your add-in. Note that data stored this way can also be accessed by other add-ins. You can persist custom XML markup in a task pane add-in for Word (and for Excel and Word using application-specific API as mentioned in the previous paragraph). In Word, you can use the [CustomXmlPart](/javascript/api/office/office.customxmlpart) object and its methods. The following code creates a custom XML part and displays its ID and then its content in divs on the page. Note that there must be an `xmlns` attribute in the XML string.

```js
function createCustomXmlPart() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            $("#xml-id").text("Your new XML part's ID: " + asyncResult.value.id);
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

To retrieve a custom XML part, use the [getByIdAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbyidasync-member(1)) method, but the ID is a GUID that is generated when the XML part is created, so you can't know when coding what the ID is. For that reason, it's a good practice when creating an XML part to immediately store the ID of the XML part as a setting and give it a memorable key. The following method shows how to do this.

 ```js
function createCustomXmlPartAndStoreId() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            Office.context.document.settings.set('ReviewersID', asyncResult.value.id);
            Office.context.document.settings.saveAsync();
        }
    );
}
```

The following code shows how to retrieve the XML part by first getting its ID from a setting.

 ```js
function getReviewers() {
    const reviewersXmlId = Office.context.document.settings.get('ReviewersID');
    Office.context.document.customXmlParts.getByIdAsync(reviewersXmlId,
        (asyncResult) => {
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```


## See also

- [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Outlook add-ins](../outlook/outlook-add-ins-overview.md)
- [Get and set add-in metadata for an Outlook add-in](../outlook/metadata-for-an-outlook-add-in.md)
- [Get and set internet headers on a message in an Outlook add-in](../outlook/internet-headers.md)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
