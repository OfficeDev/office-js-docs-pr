# Setting Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Represents a setting of the add-in.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|key|string|Gets the key of the setting. Read only. Read-only.|[1.4](../requirement-sets/word-api-requirement-sets.md)|
|value|object|Gets or sets the value of the setting.|[1.4](../requirement-sets/word-api-requirement-sets.md)|

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the setting.|[1.4](../requirement-sets/word-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/word-api-requirement-sets.md)|

## Method Details


### delete()
Deletes the setting.

#### Syntax
```js
settingObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples


```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue commands add a setting.
    var settings = context.document.settings;
    var startMonth = settings.add('startMonth', { month: 'March', year: 1998 });

    // Queue a command to get the count of settings.
    var count = settings.getCount();

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log(count.value);

        // Queue a command to delete the setting.
        startMonth.delete();

        // Queue a command to get the new count of settings.
        count = settings.getCount();
    })

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    .then(context.sync)
    .then(function () {
        console.log(count.value);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
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
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
