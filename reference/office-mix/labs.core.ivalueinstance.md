
# Labs.Core.IValueInstance

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

An [Labs.Core.IValue](../../reference/office-mix/labs.core.ivalue.md) object instance that contains value data, if any.

```
interface IValueInstance
```


## Properties


|||
|:-----|:-----|
| `valueId: string`|ID of the value which this instance represents.|
| `isHint: boolean`|Boolean  **true** if this value is considered a hint.|
| `hasValue: boolean`|Boolean  **true** if the instance information contains the value.|
| `value?: any`|The value. This parameter may or may not be set depending whether it has been hidden.|
