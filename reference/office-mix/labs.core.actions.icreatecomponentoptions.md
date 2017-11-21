
# Labs.Core.Actions.ICreateComponentOptions

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Creates a new component.

```
interface ICreateComponentOptions extends Core.IActionOptions
```


## Properties


|||
|:-----|:-----|
| `componentId: string`|The component invoking the create component action.|
| `component: Core.IComponent`|The [Labs.Core.IComponent](https://dev.office.com/reference/add-ins/office-mix/labs.core.icomponent) component to create|
| `correlationId?: string`|Optional field to correlate this component across all instances of a lab. Allows the host to identify different attempts at the same component.|
