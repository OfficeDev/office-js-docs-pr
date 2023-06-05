---
title: Persist add-in state and settings
description: Learn how to persist data in Office Web Add-in applications running in the stateless environment of a browser control.
ms.date: 06/05/2023
ms.localizationpriority: medium
---

# Persist add-in state and settings

[!include[information about the common API](../includes/alert-common-api-info.md)]

Office Add-ins are essentially web applications running in the stateless environment of a browser iframe or a webview control. (For brevity hereafter, this article uses "browser control" to mean "browser or webview control".) When in use, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:

- Use members of the Office JavaScript API that store data as either:
  - Name/value pairs in a property bag stored in a location that depends on add-in type.
  - Custom XML stored in the document.

- Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).
    > [!NOTE]
    > Some browsers or the user's browser settings may block browser-based storage techniques. You should test for availability as documented in [Using the Web Storage API](https://developer.mozilla.org/docs/Web/API/Web_Storage_API/Using_the_Web_Storage_API).

This article focuses on how to use the Office JavaScript API to persist add-in state to the current document. It's recommended that you use the application-specific object if it's available for your selected Office client instead of the Common Office JavaScript version. If you need to persist state across documents, such as tracking user preferences across any documents they open, you'll need to use a different approach. For example, you could use [SSO](use-sso-to-get-office-signed-in-user-token.md) to obtain the user identity, and then save the user ID and their settings to an online database.

## Persist add-in state and settings with the Office JavaScript API

The Office JavaScript API provides objects, such as [Settings](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [CustomProperties](/javascript/api/outlook/office.customproperties), to save add-in state across sessions as described in the following table. In all cases, the saved settings values are associated with the [Id](/javascript/api/manifest/id) of the add-in that created them.

|Object|Add-in type support|Storage location|Office application support|
|:-----|:-----|:-----|:-----|
|[Settings](/javascript/api/office/office.settings)|<ul><li>content</li><li>task pane</li></ul>|The document, spreadsheet, or presentation the add-in is working with. Content and task pane add-in settings are available only to the add-in that created them from the document where they're saved.<br><br>**Important**: Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it's stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.|<ul><li>Excel</li><li>PowerPoint</li><li>Word</li></ul><br>**Note**: Task pane add-ins for Project don't support the **Settings** API for storing add-in state or settings. However, for add-ins running in Project and other Office client applications, you can use techniques such as browser cookies or web storage. For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings). |
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|mail|The user's Exchange mailbox where the add-in is installed. Because these settings are stored in the user's mailbox, they can "roam" with the user and are available to the add-in when it's running in the context of any supported Office client application or browser accessing that user's mailbox.<br><br>Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.|Outlook|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|mail|The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they're saved.|Outlook<br><br>**Note**: A version of this object is available for [Excel](/javascript/api/excel/excel.custompropertycollection) and [Word](/javascript/api/word/word.custompropertycollection). Task pane and content add-ins for Excel support [Excel.CustomProperty](/javascript/api/excel/excel.customproperty). Task pane add-ins for Word support [Word.CustomProperty](/javascript/api/word/word.customproperty). Any add-in can access any custom properties saved in the document. The key and value of a custom property are each limited to 255 characters.|
|[InternetHeaders](/javascript/api/outlook/office.internetheaders)|mail|The message, appointment, or meeting request item the add-in is working with. Custom internet headers persist after the mail item leaves Exchange and are available to the item's recipients.|Outlook|
|[CustomXmlParts](/javascript/api/office/office.customxmlparts)|task pane|The document or spreadsheet the add-in is working with. Task pane add-in custom XML parts are available to any add-in in the document where they're saved.<br><br>**Important**: Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it's stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.|<ul><li>Word (using the application-specific Word JavaScript API [Word.CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection) (recommended) or using the Office JavaScript Common API)</li><li>Excel (using the application-specific Excel JavaScript API [Excel.CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection))</li></ul>|

## Settings data is managed in memory at runtime

The following two sections discuss settings in the context of the Office Common JavaScript API. The application-specific JavaScript APIs for Excel and for Word also provide access to the custom settings. The application-specific APIs and programming patterns are somewhat different from the Common version. For more information, see [Excel.SettingCollection](/javascript/api/excel/excel.settingcollection) and [Word.SettingCollection](/javascript/api/word/word.settingcollection).

Internally, the data in the property bag accessed with the `Settings`, `CustomProperties`, or `RoamingSettings` objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs. The name (key) for each value must be a `string`, and the stored value can be a JavaScript `string`, `number`, `date`, or `object`, but not a **function**.

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

## How to save add-in state and settings per document for content and task pane add-ins

To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, use the [Settings](/javascript/api/office/office.settings) object and its methods. The property bag created with the methods of the `Settings` object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.

The `Settings` object is automatically loaded as part of the [Document](/javascript/api/office/office.document) object, and is available when the task pane or content add-in is activated. After the `Document` object is instantiated, you can access the `Settings` object with the [settings](/javascript/api/office/office.document#office-office-document-settings-member) property of the `Document` object. During the lifetime of the session, you can use the `Settings.get`, `Settings.set`, and `Settings.remove` methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.

Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with, you must call the [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)) method.

### Create or update a setting value

The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#office-office-settings-set-member(1)) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.

```js
Office.context.document.settings.set('themeColor', 'green');
```

The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist. Use the `Settings.saveAsync` method to persist the new or updated settings to the document.

### Get the value of a setting

The following example shows how use the [Settings.get](/javascript/api/office/office.settings#office-office-settings-get-member(1)) method to get the value of a setting called "themeColor". The only parameter of the `get` method is the case-sensitive _name_ of the setting.

```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

The `get` method returns the value that was previously saved for the setting _name_ that was passed in. If the setting doesn't exist, the method returns **null**.

### Remove a setting

The following example shows how to use the [Settings.remove](/javascript/api/office/office.settings#office-office-settings-remove-member(1)) method to remove a setting with the name "themeColor". The only parameter of the `remove` method is the case-sensitive _name_ of the setting.

```js
Office.context.document.settings.remove('themeColor');
```

Nothing will happen if the setting doesn't exist. Use the `Settings.saveAsync` method to persist removal of the setting from the document.

### Save your settings

To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)) method to store them in the document. The only parameter of the `saveAsync` method is _callback_, which is a callback function with a single parameter.

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

The anonymous function passed into the `saveAsync` method as the _callback_ parameter is executed when the operation is completed. The _asyncResult_ parameter of the callback provides access to an `AsyncResult` object that contains the status of the operation. In the example, the function checks the `AsyncResult.status` property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.

## How to save custom XML to the document

This section discusses custom XML parts in the context of the Office Common JavaScript API which is supported in Word. The application-specific JavaScript APIs for Excel and for Word also provide access to the custom XML parts. The application-specific APIs and programming patterns are somewhat different from the Common version. For more information, see [Excel.CustomXmlPart](/javascript/api/excel/excel.customxmlpart) and [Word.CustomXmlPart](/javascript/api/word/word.customxmlpart).

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
            Office.context.document.settings.set('ReviewersID', asyncResult.id);
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

## How to save settings in an Outlook add-in

For information about how to save settings in an Outlook add-in, see [Get and set add-in metadata for an Outlook add-in](../outlook/metadata-for-an-outlook-add-in.md) and [Get and set internet headers on a message in an Outlook add-in](../outlook/internet-headers.md).

## See also

- [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Outlook add-ins](../outlook/outlook-add-ins-overview.md)
- [Get and set add-in metadata for an Outlook add-in](../outlook/metadata-for-an-outlook-add-in.md)
- [Get and set internet headers on a message in an Outlook add-in](../outlook/internet-headers.md)
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
