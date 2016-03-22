
# Labs.takeLab

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Runs the specified lab and enables sending lab results to the server. Note that you cannot run a lab while it is being edited.

```
function takeLab(callback: Core.ILabCallback<LabInstance>): void
```


## Parameters


|**Name**|**Description**|
|:-----|:-----|
| _callback_|The callback method fired once the [Labs.LabInstance](../../reference/office-mix/labs.labinstance.md) object is created.|
