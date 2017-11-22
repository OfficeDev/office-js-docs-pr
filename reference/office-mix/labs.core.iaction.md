
# Labs.Core.IAction

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Represents a lab action, which is an interaction that a user has with a specified lab.

```
interface IAction
```


## Properties


|||
|:-----|:-----|
| `type: string`|The type of action taken by the user.|
| `options: Core.IActionOptions`|The [Labs.Core.IActionOptions](https://dev.office.com/reference/add-ins/office-mix/labs.core.iactionoptions) options sent with the action taken by the user.|
| `result: Core.IActionResult`|The [Labs.Core.IActionResult](https://dev.office.com/reference/add-ins/office-mix/labs.core.iactionresult) result of the action.|
| `time: number`|The time at which the action was completed, represented in milliseconds elapsed since 01 January 1970 00:00:00 UTC.|
