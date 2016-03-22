
# Labs.connect (overload)

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Initializes a connection with the host.

```
function connect(labHost: Core.ILabHost, callback: Core.ILabCallback<Core.IConnectionResponse>)
```


## Parameters


|||
|:-----|:-----|
| _labHost_|Optional. The [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md)instance to which to connect. If the host is not specified, one will be constructed using [Labs.DefaultHostBuilder](../../reference/office-mix/labs.defaulthostbuilder.md).|
| _callback_|Callback that fires once the connection has been established.|

## Return value

Returns a connection to the host.

