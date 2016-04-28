# Notebook Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a OneNote notebook. Notebooks contain section groups and sections.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID of the notebook. Read-only.|
|name|string|Gets the name of the notebook. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[addSection(name: String)](#addsectionname-string)|[Section](section.md)|Adds a new section to the end of the notebook.|
|[getSectionGroups()](#getsectiongroups)|[SectionGroupCollection](sectiongroupcollection.md)|Gets the section groups in the notebook.|
|[getSections(recursive: bool)](#getsectionsrecursive-bool)|[SectionCollection](sectioncollection.md)|Gets the sections in the notebook.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

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
    var notebook = context.application.activeNotebook;

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


### getSectionGroups()
Gets the section groups in the notebook.

#### Syntax
```js
notebookObject.getSectionGroups();
```

#### Parameters
None

#### Returns
[SectionGroupCollection](sectiongroupcollection.md)

#### Examples
```js          
OneNote.run(function (context) {

    // Get the section groups in the notebook. 
    var sectionGroups = context.application.activeNotebook.getSectionGroups();

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


### getSections(recursive: bool)
Gets the sections in the notebook.

#### Syntax
```js
notebookObject.getSections(recursive);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|recursive|bool|true to retrieve all child sections, or false to retrieve direct child sections only. Default is false.|

#### Returns
[SectionCollection](sectioncollection.md)

#### Examples
```js
OneNote.run(function (context) {

    // Gets the active notebook.
    var notebook = context.application.activeNotebook;
    
    // Queue a command to get immediate child sections of the notebook. 
    var childSections = notebook.getSections(false);

    // Queue a command to get all sections in the notebook, including sections in section groups.
    var allChildSections = notebook.getSections(true);

    // Queue a command to load the childSections. 
    context.load(childSections);

    // Queue a command to load the allChildSections. 
    context.load(allChildSections);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            $.each(childSections.items, function(index, childSection) {
                console.log("Immediate child section name: " + childSection.name);
            });

            $.each(allChildSections.items, function(index, childSection) {
                console.log("Child section name: " + childSection.name);
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
**id**
```js
OneNote.run(function (context) {
        
    // Get the current notebook.
    var notebook = context.application.activeNotebook;
            
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
    var notebook = context.application.activeNotebook;
            
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

