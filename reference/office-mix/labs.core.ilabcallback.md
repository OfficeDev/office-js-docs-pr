
# Labs.Core.ILabCallback

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

The interface for handling Labs.js callback methods.

```
interface ILabCallback<T>
```


## Callback signature

 `(err: any, data: T): void`

 **Callback parameters**


|||
|:-----|:-----|
| _err_|**Null** if no errors occur. Non- **null** if an error has occurred.|
| _data_|The data returned with the callback.|
