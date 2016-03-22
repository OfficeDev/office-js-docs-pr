
# Labs.IMessageHandler

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Interface that allows you to define event handlers.

```
interface IMessageHandler(origin: Window, data: any, callback: Labs.Core.ILabCallback<any>): void
```


## 

 **Parameters**


|||
|:-----|:-----|
| `origin`|The lab window from which the message originated.|
| `data`|The contents of the message.|
| `callback`|Callback function that fires once the message is received.|
