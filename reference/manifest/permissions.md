
# Permissions element
Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.

 **Add-in type:** Content, Task pane, Mail


## Syntax:

For content and task pane add-ins:


```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

For mail add-ins:




```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```


## Contained in:

 _[OfficeApp](../../reference/manifest/officeapp.md)_


## Remarks

For more detail, see [Requesting permissions for API use in content and task pane add-ins](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../docs/outlook/understanding-outlook-add-in-permissions.md).

