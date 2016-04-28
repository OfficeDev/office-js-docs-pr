# Application Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents the top-level object that contains all globally addressable OneNote objects such as notebooks, the active notebook, and the active section.

## Properties

None

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|activeNotebook|[Notebook](notebook.md)|Gets the active notebook. Read-only.|
|activeOutline|[Outline](outline.md)|Gets the active outline. Read-only.|
|activePage|[Page](page.md)|Gets the active page. Read-only.|
|activeSection|[Section](section.md)|Gets the active section. Read-only.|
|notebooks|[NotebookCollection](notebookcollection.md)|Gets the collection of notebooks that are open in the OneNote application instance. In OneNote Online, only one notebook at a time is open in the application instance. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[navigateToPage(page: Page)](#navigatetopagepage-page)|void|Opens the specified page in the application instance.|
|[navigateToPageWithClientUrl(url: string)](#navigatetopagewithclienturlurl-string)|void|Opens the specified page in the application instance.|
|[navigateToPageWithPagePath(pagePath: string)](#navigatetopagewithpagepathpagepath-string)|void|Opens the specified page in the application instance.|

## Method Details


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

### navigateToPage(page: Page)
Opens the specified page in the application instance.

#### Syntax
```js
applicationObject.navigateToPage(page);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|page|Page|The page to open.|

#### Returns
void

#### Examples
```js        
OneNote.run(function (context) {
        
    // Get the pages in the current section.
    var pages = context.application.activeSection.getPages();
            
    // Queue a command to load the pages. 
    // For best performance, request specific properties.           
    pages.load('id');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // This example loads the first page in the section.
            var page = pages.items[0];
                        
            // Open the page in the application.                    
            context.application.navigateToPage(page);
                    
            // Run the queued command.
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
### navigateToPageWithClientUrl(url: string)
Opens the specified page in the application instance.

#### Syntax
```js
applicationObject.navigateToPageWithClientUrl(url);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|url|string|The client url of the page to open|

#### Returns
void

### navigateToPageWithPagePath(pagePath: string)
Opens the specified page in the application instance.

#### Syntax
```js
applicationObject.navigateToPageWithPagePath(pagePath);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|pagePath|string|The path of the page to open|

#### Returns
void
### Property access examples

**activeNotebook**
```js
OneNote.run(function (context) {
        
    // Get the current notebook.
    var notebook = context.application.activeNotebook;
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Notebook name: " + notebook.name);
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

**activePage**
```js
OneNote.run(function (context) {
        
    // Get the current notebook.
    var page = context.application.activePage;
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    page.load('id,title');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Page title: " + page.title);
            console.log("Page ID: " + page.id);
            
        });
    })
    .catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
```

**activeSection**
```js
OneNote.run(function (context) {
        
    // Get the current section.
    var section = context.application.activeSection;
            
    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Section name: " + section.name);
            console.log("Section ID: " + section.id);
            
        });
    })
    .catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
```

**notebooks**
```js
OneNote.run(function (context) {
        
    // Get the current notebook.
    var notebooks = context.application.notebooks;
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebooks.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                        
            $.each(notebooks.items, function(index, notebook) {
                
                // Show some properties.
                console.log("Notebook name: " + notebook.name);
                console.log("Notebook ID: " + notebook.id);
                
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

