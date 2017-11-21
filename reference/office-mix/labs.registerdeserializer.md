
# Labs.registerDeserializer

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Deserializes a specified JSON object into an object. Should be used by component authors only.

```
function registerDeserializer(type: string, deserialize: (json: Core.ILabObject): any): void
```


## Parameters


|**Name**|**Description**|
|:-----|:-----|
|json|The [Labs.Core.ILabObject](https://dev.office.com/reference/add-ins/office-mix/labs.core.ilabobject) to deserialize.|

## Return value

Returns an [Labs.Core.ILabObject](https://dev.office.com/reference/add-ins/office-mix/labs.core.ilabobject) instance.

