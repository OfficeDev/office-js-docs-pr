# Permissions element

Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.

**Add-in type:** Content, Task pane, Mail

## Syntax

For content and task pane add-ins:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

For mail add-ins

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## Contained in

[OfficeApp](officeapp.md)

## Remarks

For more detail, see [Requesting permissions for API use in content and task pane add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).
