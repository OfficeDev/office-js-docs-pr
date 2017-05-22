# SettingCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Contains the collection of [setting](setting.md) objects.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Setting[]](setting.md)|A collection of setting objects. Read-only.|[1.4](../requirement-sets/word-api-requirement-sets.md)|

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[deleteAll()](#deleteall)|void|Deletes all settings in this add-in.|[1.4](../requirement-sets/word-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Gets the count of settings.|[1.4](../requirement-sets/word-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Setting](setting.md)|Gets a setting object by its key, which is case-sensitive. Throws if the setting does not exist.|[1.4](../requirement-sets/word-api-requirement-sets.md)|
|[getItemOrNullObject(key: string)](#getitemornullobjectkey-string)|[Setting](setting.md)|Gets a setting object by its key, which is case-sensitive. Returns a null object if the setting does not exist.|[1.4](../requirement-sets/word-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[set(key: string, value: object)](#setkey-string-value-object)|[Setting](setting.md)|Creates or sets a setting.|[1.4](../requirement-sets/word-api-requirement-sets.md)|

## Method Details


### deleteAll()
Deletes all settings in this add-in.

#### Syntax
```js
settingCollectionObject.deleteAll();
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
    settings.add('startMonth', { month: 'March', year: 1998 });

    // Queue a command to get the count of settings.
    var count = settings.getCount();

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log(count.value);

        // Queue a command to delete all settings.
        settings.deleteAll();

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


### getCount()
Gets the count of settings.

#### Syntax
```js
settingCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue commands add a setting.
    var settings = context.document.settings;
    settings.add('startMonth', { month: 'March', year: 1998 });

    // Queue a command to get the count of settings.
    var count = settings.getCount();

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log(count.value);

        // Queue a command to delete all settings.
        settings.deleteAll();

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



### getItem(key: string)
Gets a setting object by its key, which is case-sensitive. Throws if the setting does not exist.

#### Syntax
```js
settingCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|string|The key that identifies the setting object.|

#### Returns
[Setting](setting.md)

#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue commands add a setting.
    var settings = context.document.settings;
    settings.add('startMonth', { month: 'March', year: 1998 });

    // Queue a command to retrieve a setting.
    var startMonth = settings.getItem('startMonth');

    // Queue a command to load the setting.
    context.load(startMonth);

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log(JSON.stringify(startMonth.value));
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


#### Examples

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue commands add a setting.
    var settings = context.document.settings;
    settings.add('startMonth', { month: 'March', year: 1998 });
    
    // Queue commands to retrieve settings.
    var startMonth = settings.getItemOrNullObject('startMonth');
    var endMonth = settings.getItemOrNullObject('endMonth');

    // Queue commands to load settings.
    context.load(startMonth);
    context.load(endMonth);

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
       return context.sync().then(function () {
           if (startMonth.isNullObject) {
               console.log("No such setting.");
           }
           else {
               console.log(JSON.stringify(startMonth.value));
           }
            if (endMonth.isNullObject) {
               console.log("No such setting.");
           }
           else {
               console.log(JSON.stringify(endMonth.value));
           }
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```


### getItemOrNullObject(key: string)
Gets a setting object by its key, which is case-sensitive. Returns a null object if the setting does not exist.

#### Syntax
```js
settingCollectionObject.getItemOrNullObject(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|string|Required. The key that identifies the setting object.|

#### Returns
[Setting](setting.md)

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

### set(key: string, value: object)
Creates or sets a setting.

#### Syntax
```js
settingCollectionObject.set(key, value);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|string|Required. The setting's key, which is case-sensitive.|
|value|object|Required. The setting's value.|

#### Returns
[Setting](setting.md)
