# Object Load Options (JavaScript API for Visio)

>**Note:** The Visio JavaScript APIs are not currently available for use in preview or production environments.

Represents an object that can be passed to the load method to specify the set of properties and relations to be loaded upon execution of **sync()** method that synchronizes the states between Visio objects and corresponding JavaScript proxy objects. This takes in options such as select and expand parameters to specify a set of properties to be loaded on the object and also allows pagination on the collection.

It is also valid to supply a string containing the properties and relations to be loaded or to provide an array containing the list of properties and relations to be loaded. See the following example.

```js
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

## Properties

| Property | Type  | Description |
|:---------|:------|:------------|
|select    |object |Provide comma-delimited list or an array of parameter/relationship names to be loaded upon an executeAsync call, for example, "property1, relationship1", [ "property1", "relationship1"]. Optional.|
|expand    |object |Provide comma-delimited list or an array of relationship names to be loaded upon an executeAsync call, for example, "relationship1, relationship2", [ "relationship1", "relationship2"]. Optional.|
|top       |int    |Specify the number of items in the queried collection to be included in the result. Optional.|
|skip      |int    |Specify the number of items in the collection that are to be skipped and not included in the result. If **top** is specified, the selection of result will start after skipping the specified number of items. Optional.|

