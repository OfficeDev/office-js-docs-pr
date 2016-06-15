# PageContent Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a region on a page that contains top-level content types such as Outline or Image. A PageContent object can be assigned an XY position.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|:-------|
|id|string|Gets the ID of the PageContent object. Read-only.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-id)|
|left|double|Gets or sets the left (X-axis) position of the PageContent object.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-left)|
|top|double|Gets or sets the top (Y-axis) position of the PageContent object.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-top)|
|type|string|Gets the type of the PageContent object. Read-only. Possible values are: Outline, Image, Other.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-type)|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|:-------|
|image|[Image](image.md)|Gets the Image in the PageContent object. Returns null if PageContentType is not Image. Read-only.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-image)|
|outline|[Outline](outline.md)|Gets the Outline in the PageContent object. Returns null if PageContentType is not Outline. Read-only.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-outline)|
|parentPage|[Page](page.md)|Gets the page that contains the PageContent object. Read-only.||[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-parentPage)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[delete()](#delete)|void|Deletes the PageContent object.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-delete)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-load)|
|[select()](#select)|void|Selects the PageContent object.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContent-select)|

## Method Details


### delete()
Deletes the PageContent object.

#### Syntax
```js
pageContentObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples
```js
OneNote.run(function (context) {

	var page = context.application.getActivePage();
	var pageContents = page.contents;

	var firstPageContent = pageContents.getItemAt(0);
	firstPageContent.load('type');

	// Run the queued commands, and return a promise to indicate task completion.
	return context.sync()
		.then(function () {
			if(firstPageContent.isNull === false) {
				firstPageContent.delete();
				return context.sync();
			}
		});
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```
### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

### select()
Selects the PageContent object.

#### Syntax
```js
pageContentObject.select();
```

#### Parameters
None

#### Returns
void

#### Examples
```js
OneNote.run(function (context) {

	var page = context.application.getActivePage();
	var pageContents = page.contents;
	pageContents.load('type');

	// Run the queued commands, and return a promise to indicate task completion.
	return context.sync()
		.then(function () {
			var firstPageContent = pageContents.getItemAt(0);
			firstPageContent.select();
			return context.sync();
		});
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

