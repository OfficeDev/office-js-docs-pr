
# Labs.Command

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

General command used to pass messages between the client and host.

```
class Command
```


## Properties


|**Name**|**Description**|
|:-----|:-----|
| `public var type: string`|The type of the command.|
| `public var commandData: any`|Optional data associated with the command.|

## Methods




### constructor

 `function constructor(type: string, commandData?: any)`

Description

 **Parameters**


|||
|:-----|:-----|
| `type`|The type of the command.|
| `commandData`|Optional data associated with the command.|
