
# Labs.editLab

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Opens the specified lab for editing. You can specify the lab's configuration data while in edit mode. However, you cannot edit a lab while it is being taken (that is, the lab is running).

```
function editLab(callback: Core.ILabCallback<LabEditor>): void
```


## Parameters


|**Name**|**Description**|
|:-----|:-----|
| _callback_|The callback method that is fired once the [Labs.LabInstance](../../reference/office-mix/labs.labinstance.md) object is created.|
