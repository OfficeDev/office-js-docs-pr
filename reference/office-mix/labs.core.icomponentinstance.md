
# Labs.Core.IComponentInstance

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Base class for instances of lab components.

```
interface IComponentInstance extends Core.ILabObject, Core.IUserData
```


## Properties


|||
|:-----|:-----|
| `componentId: string`|The ID of the component this instance is associated with.|
| `name: string`|Name of the component.|
| `values: {[type:string]: Core.IValueInstance[]}`|The value property map associated with the component.|

## Remarks

A component instance is an instantiation of a component for a user. It contains a translated view of the component for a particular run of the lab. This view may exclude hidden information (answers, hints, and so forth) and also contains IDs to identify the various instances.

