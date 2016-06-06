# SectionGroupCollection Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a collection of section groups.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|count|int|Returns the number of section groups in the collection. Read-only.|
|items|[SectionGroup[]](sectiongroup.md)|A collection of sectionGroup objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getByName(name: string)](#getbynamename-string)|[SectionGroupCollection](sectiongroupcollection.md)|Gets the collection of section groups with the specified name.|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[SectionGroup](sectiongroup.md)|Gets a section group by ID or by its index in the collection. Read-only.|
|[getItemAt(index: number)](#getitematindex-number)|[SectionGroup](sectiongroup.md)|Gets a section group on its position in the collection.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getByName(name: string)
Gets the collection of section groups with the specified name.

#### Syntax
```js
sectionGroupCollectionObject.getByName(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|The name of the section group.|

#### Returns
[SectionGroupCollection](sectiongroupcollection.md)

#### Examples
```js
OneNote.run(function (context) {

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the section groups. 
    // For best performance, request specific properties.
    sectionGroups.load("id"); 

    // Get the section groups with the specified name.
    var labsSectionGroups = sectionGroups.getByName("Labs");

    // Queue a command to load the section groups with the specified properties.
    labsSectionGroups.load("id,name"); 
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index.
            if (labsSectionGroups.items.length > 0) {
                console.log("Section group name: " + labsSectionGroups.items[0].name);
                console.log("Section group ID: " + labsSectionGroups.items[0].id);
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


### getItem(index: number or string)
Gets a section group by ID or by its index in the collection. Read-only.

#### Syntax
```js
sectionGroupCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|The ID of the section group, or the index location of the section group in the collection.|

#### Returns
[SectionGroup](sectiongroup.md)

### getItemAt(index: number)
Gets a section group on its position in the collection.

#### Syntax
```js
sectionGroupCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[SectionGroup](sectiongroup.md)

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

**items**
```js
OneNote.run(function (context) {

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the section groups. 
    // For best performance, request specific properties.
    sectionGroups.load("name"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Iterate through the collection or access items individually by index, for example: sectionGroups.items[0]
            $.each(sectionGroups.items, function(index, sectionGroup) {
                console.log("Section group name: " + sectionGroup.name);  
                console.log("Section group ID: " + sectionGroup.id);  
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

