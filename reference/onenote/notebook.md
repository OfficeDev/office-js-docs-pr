# Notebook Object (JavaScript API for OneNote)

_Applies to: OneNote Online_   


Represents a OneNote notebook. Notebooks contain section groups and sections.

## Properties

| Property	   | Type	|Description|Feedback|
|:---------------|:--------|:----------|:-------|
|baseUrl|string|The base site url for the notebook, if it is in a SharePoint site (it will be null for OneDrive notebooks) - this is to be used for interacting with the OneNote REST API, according to our [OneNote API documentation](https://blogs.msdn.microsoft.com/onenotedev/2015/06/11/and-sharepoint-makes-three/) to fetch the SiteCollectionId and SiteId (using FromUrl). Read only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-baseUrl)|
|clientUrl|string|The client url of the notebook. Read only Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-clientUrl)|
|id|string|Gets the ID of the notebook. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-id)|
|name|string|Gets the name of the notebook. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-name)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|sectionGroups|[SectionGroupCollection](sectiongroupcollection.md)|The section groups in the notebook. Read only Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-sectionGroups)|
|sections|[SectionCollection](sectioncollection.md)|The the sections of the notebook. Read only Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-sections)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[addSection(name: String)](#addsectionname-string)|[Section](section.md)|Adds a new section to the end of the notebook.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-addSection)|
|[addSectionGroup(name: String)](#addsectiongroupname-string)|[SectionGroup](sectiongroup.md)|Adds a new section group to the end of the notebook.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-addSectionGroup)|
|[getRestApiId()](#getRestApiId)|string|Gets the id that is compatible with the REST API.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-getRestApiId)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-load)|

## Method Details


### addSection(name: String)
Adds a new section to the end of the notebook.

#### Syntax
```js
notebookObject.addSection(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|String|The name of the new section.|

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


### addSectionGroup(name: String)
Adds a new section group to the end of the notebook.

#### Syntax
```js
notebookObject.addSectionGroup(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|String|The name of the new section.|

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

### getRestApiId()
Gets the id that is compatible with the REST API.

#### Syntax
```js
notebookObject.getRestApiId();
```

#### Parameters
None

#### Returns
string

#### Examples
```js

OneNote.run(function(ctx){
    // Get the current notebook.         
    var notebook = ctx.application.getActiveNotebook();
    var restApiId = notebook.getRestApiId();

    return ctx.sync().
        then(function(){
            console.log("The rest api id is " + restApiId.value);
            // Note that the rest API id isn't all you need to interact with the OneNote Rest API. For SharePoint notebooks, the notebook baseUrl should be used to talk to OneNote REST APIs according to https://blogs.msdn.microsoft.com/onenotedev/2015/06/11/and-sharepoint-makes-three/
            // (this is only required for SharePoint notebooks, baseUrl will be null for OneDrive notebooks)
        });
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

