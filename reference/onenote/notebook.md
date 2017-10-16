# Notebook Object (JavaScript API for OneNote)

_Applies to: OneNote Online_   

Represents a OneNote notebook. Notebooks contain section groups and sections. 

To provide feedback on this API, you can [file an issue in GitHub](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook).

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|baseUrl|string|The base site URL for the notebook, if it is in a SharePoint site (it will be null for OneDrive notebooks). Use this property to interact with the OneNote REST API to fetch the **SiteCollectionId** and **SiteId** (using **FromUrl**). Read-only.<br/><br/>For details, see the [OneNote development blog](https://blogs.msdn.microsoft.com/onenotedev/2015/06/11/and-sharepoint-makes-three/).|
|clientUrl|string|The client URL of the notebook. Read-only.|
|id|string|Gets the ID of the notebook. Read-only.|
|name|string|Gets the name of the notebook. Read-only.|

_See [property access examples](#property-access-examples)_.

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|sectionGroups|[SectionGroupCollection](sectiongroupcollection.md)|The section groups in the notebook. Read-only.|
|sections|[SectionCollection](sectioncollection.md)|The sections of the notebook. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[addSection(name: String)](#addsectionname-string)|[Section](section.md)|Adds a new section to the end of the notebook.|
|[addSectionGroup(name: String)](#addsectiongroupname-string)|[SectionGroup](sectiongroup.md)|Adds a new section group to the end of the notebook.|
|[getRestApiId()](#getRestApiId)|string|Gets the ID that is compatible with the REST API.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.|

## Method details


### addSection(name: String)
Adds a new section to the end of the notebook.

#### Syntax

```js
notebookObject.addSection(name);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|The name of the new section.|

#### Returns

[Section](section.md)

#### Examples

```js          
OneNote.run(function (context) {

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();

    // Queue a command to add a new section. 
    var section = notebook.addSection("Sample section");
    
    // Queue a command to load the new section. This example reads the name property later.
    section.load("name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("New section name is " + section.name);
        });
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
}); 
```

<br/>

### addSectionGroup(name: String)
Adds a new section group to the end of the notebook.

#### Syntax

```js
notebookObject.addSectionGroup(name);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|The name of the new section group.|

#### Returns

[SectionGroup](sectiongroup.md)

#### Examples

```js          
OneNote.run(function (context) {

	// Gets the active notebook.
	var notebook = context.application.getActiveNotebook();

	// Queue a command to add a new section group.
	var sectionGroup = notebook.addSectionGroup("Sample section group");

	// Queue a command to load the new section group.
	sectionGroup.load();

	// Run the queued commands, and return a promise to indicate task completion.
	return context.sync()
		.then(function() {
			console.log("New section group name is " + sectionGroup.name);
		});
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
}); 
```

<br/>

### getRestApiId()

Gets the ID that is compatible with the REST API.

#### Syntax

```js
notebookObject.getRestApiId();
```

#### Parameters
None

#### Returns

String

#### Examples

```js

OneNote.run(function(ctx){
    // Get the current notebook.         
    var notebook = ctx.application.getActiveNotebook();
    var restApiId = notebook.getRestApiId();

    return ctx.sync().
        then(function(){
            console.log("The REST API ID is " + restApiId.value);
            // Note that the REST API ID isn't all you need to interact with the OneNote REST API. For SharePoint notebooks, the notebook baseUrl should be used to talk to the OneNote REST API according to [OneNote Development Blog](https://blogs.msdn.microsoft.com/onenotedev/2015/06/11/and-sharepoint-makes-three/)
            // (this is only required for SharePoint notebooks, baseUrl will be null for OneDrive notebooks)
        });
});
```

<br/>

### load(param: object)
Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.

#### Syntax

```js
object.load(param);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns

Void

<br/>

### Property access examples

**baseUrl**

```js
OneNote.run(function (context) {
        
    // Get the current notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('baseUrl');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Base url: " + notebook.baseUrl);
            // This baseUrl should be used to talk to OneNote REST APIs according to https://blogs.msdn.microsoft.com/onenotedev/2015/06/11/and-sharepoint-makes-three/ (only required for SharePoint notebooks, will be null for OneDrive notebooks)
        });
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

**id**

```js
OneNote.run(function (context) {
        
    // Get the current notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Notebook ID: " + notebook.id);
            
        });
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

**name**

```js
OneNote.run(function (context) {
        
    // Get the current notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Notebook name: " + notebook.name);
            
        });
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

**sectionGroups**

```js          
OneNote.run(function (context) {

    // Get the section groups in the notebook. 
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the sectionGroups. 
    sectionGroups.load("name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            $.each(sectionGroups.items, function(index, sectionGroup) {
                console.log("Section group name: " + sectionGroup.name);
            });
        });
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

<br/>

**sections**

```js
OneNote.run(function (context) {

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();
    
    // Queue a command to get immediate child sections of the notebook. 
    var childSections = notebook.sections;

    // Queue a command to load the childSections. 
    context.load(childSections);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            $.each(childSections.items, function(index, childSection) {
                console.log("Immediate child section name: " + childSection.name);
            });            
        });
})
.catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});   
```

