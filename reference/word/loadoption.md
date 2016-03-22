# LoadOption object (JavaScript API for Word)

An object that specifies paging information and the properties to load when context.sync() is called.

_Applies to: Word 2016, Word for iPad, Word for Mac_

## Properties
| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|select|object|Contains a comma delimited list or an array of parameter/relationship names. Optional.|
|expand|object|Contains a comma delimited list or an array of relationship names. Optional.|
|top|int| Specifies the maximum number of collection items that can be included in the result. Optional. You can only use this option when you use the object notation option.|
|skip|int|Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the result set will start after skipping the specified number of items. Optional. You can only use this option when you use the object notation option.|

## More information

The preferred method for specifying the properties and paging information is by using a string literal. The first two examples show the preferred way to request the text and font size properties for paragraphs in a paragraph collection:

<code>context.load(paragraphs, 'text, font/size');</code>

<code>paragraphs.load('text, font/size');</code>

Here is a similar example using object notation (includes paging):

<code>context.load(paragraphs, {select: 'text, font/size',
                                expand: 'font',
                                top: 50,
                                skip: 0});</code>

<code>paragraphs.load({select: 'text, font/size',
                       expand: 'font',
                       top: 50,
                       skip: 0});</code>

Note that if we don't specify the specific properties on the font object in the select statement, the expand statement by itself would indicate that all of the font properties are loaded.

## Examples

This example shows how to get the paragraphs in the Word document along with their text and font size properties.

```js
        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy object for the paragraphs collection.
            var paragraphs = context.document.body.paragraphs;

            // Queue a commmand to load the text and font properties.
            // It is best practice to always specify the property set. Otherwise, all properties are
            // returned in on the object.
            context.load(paragraphs, 'text, font/size');

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

            // Insert code that works with the paragraphs loaded by context.load().

        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });

```

## Support details

Use the [requirement set](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) in run time checks to make sure your application is supported by the host version of Word. For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx).