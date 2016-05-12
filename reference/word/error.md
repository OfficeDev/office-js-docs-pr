# OfficeExtension.Error object (JavaScript API for Word)

Represents errors that occur when you use the Word JavaScript API.

_Applies to: Word 2016, Word for iPad, Word for Mac_

## Properties
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|code|string|Gets a value that indicates the type of error. The value can be "AccessDenied", "GeneralException", "ActivityLimitReached", "InvalidArgument", "ItemNotFound", or "NotImplemented". <!-- Values come from OfficeExtension.Error and Word.ErrorCodes. -->|
|debugInfo|string|Gets a value that indicates what happened when the error occurred. This value is only intended for use during development / debugging.  |
|message |string| Gets a localized human readable string that corresponds to the error code.|
|name |string| Gets a value that is always "OfficeExtension.Error". |
|traceMessages |string[]| Gets an array of values that correspond to the instrumention messages set with context.trace(); |

_See property access [examples.](#property-access-examples)_

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[toString()](#tostring)|string|Returns the error code and message values in the following format: "{0}: {1}", code, message.|

## Method details

### toString()
Returns the error code and message values in the following format: "{0}: {1}", code, message.

#### Syntax
```js
error.toString()
```

#### Parameters
None.

#### Returns
string

#### Examples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

	// Create a proxy object for the document body.
	var body = context.document.body;

	// Queue a commmand to insert text in to the beginning of the body.
    // This will cause an OfficeExtension.Error.
	body.insertText(0);

	// Synchronize the document state by executing the queued-up commands,
	// and return a promise to indicate task completion.
	return context.sync();
})
.catch(function (error) {
	if (error instanceof OfficeExtension.Error) {
		console.log('Error code and message: ' + error.toString());
	}
});

```

## Property access examples

### Trace message instrumentation

The following example shows how you can instrument a batch of commands to determine where an error occurred. The first batch successfully inserts the first two paragraphs into the document and cause no errors. The second batch successfully inserts the third and fourth paragraphs but fails in the call to insert the fifth paragraph. All other commands after the failed command in the batch are not executed, including the command that adds the fifth trace message. In this case, the error occurred after the fourth paragraph was inserted, and before adding the fifth trace message.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

	// Create a proxy object for the document body.
	var body = context.document.body;

	// Queue a commmand to insert the paragraph at the end of the document body.
    // Start a batch of commands.
	body.insertParagraph('1st paragraph', Word.InsertLocation.end);
	// Queue a command for instrumenting this part of the batch.
	context.trace('1st paragraph successful');

	body.insertParagraph('2nd paragraph', Word.InsertLocation.end);
	context.trace('2nd paragraph successful');

	// Synchronize the document state by executing the queued-up commands,
	// and return a promise to indicate task completion.
	return context.sync().then(function () {
		// Queue a commmand to insert the paragraph at the end of the document body.
        // Start a new batch of commands.
		body.insertParagraph('3rd paragraph', Word.InsertLocation.end);
		context.trace('3rd paragraph successful');

		body.insertParagraph('4th paragraph', Word.InsertLocation.end);
		context.trace('4th paragraph successful');

		// This command will cause an error. The trace messages in the queue up to
        // this point will be available via Error.traceMessages.
		body.insertParagraph(0, '5th paragraph', Word.InsertLocation.end);
		// Queue a command for instrumenting this part of the batch.
        // This trace message will not be set on Error.traceMessages.
		context.trace('5th paragraph successful');
	}).then(context.sync);
})
.catch(function (error) {
	if (error instanceof OfficeExtension.Error) {
		console.log('Trace messages: ' + error.traceMessages);
	}
});

// Output: "Trace messages: 3rd paragraph successful,4th paragraph successful"

```
