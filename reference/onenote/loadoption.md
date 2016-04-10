# Object Load Options 

Represents an object that can be passed to the load method to specify the set of properties and relations to be loaded upon execution of sync() method that synchronizes the states between OneNote objects and corresponding JavaScript proxy objects in the add-in. This takes in options such as select and expand parameters to specify set of properties to be loaded on the object and also allows pagination on the collection.

It is also valid to supply a string containing the properties and relations to be loaded or to provide an array containing the list of properties and relations to be loaded. See the following example.

```js	
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

## Properties
| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|select|object|Provide comma-delimited list or an array of parameter/relationship names to be loaded upon an executeAsync call, for example, "property1, relationship1", [ "property1", "relationship1"]. Optional.|
|expand|object|Provide comma-delimited list or an array of relationship names to be loaded upon an executeAsync call, for example, "relationship1, relationship2", [ "relationship1", "relationship2"]. Optional.|
|top|int| Specify the number of items in the queried collection to be included in the result. Optional.|
|skip|int|Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the selection of result will start after skipping the specified number of items. Optional.|

#### Examples

In the example, select the top 100 rows of the table.

```js
OneNote.run(function (context) { 
    
    // Get the pages in the current section.
    var pages = context.application.activeSection.getPages();
            
    // Queue a command to load the pages.           
    pages.load({ "select":"title,pageLevel", "top":5, "skip":0 });
	return context.sync()
        .then(function() {
            
            // Iterate through the collection of pages.    
            $.each(pages.items, function(index, object) {
                var title = object.title;
                var pageLevel = object.pageLevel;
                
                // Show some properties.
                console.log("Page title: " + title);
                console.log("Indentation level: " + pageLevel);
                
            });
        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        })
    });
```