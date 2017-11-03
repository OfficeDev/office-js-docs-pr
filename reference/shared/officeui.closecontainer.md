# UI.closeContainer method

Closes the UI container where the JavaScript is executing. 

## Supported hosts

- Outlook. Minimum requirement set: Mailbox 1.5

The behavior of this method is specified by the following table.

| When called from | Behavior |
|:-----------------|:---------|
| A UI-less command button | No effect. Any dialog opened by [displayDialogAsync](officeui.displaydialogasync.md) will remain open. |
| A taskpane | The taskpane will close. Any dialog opened by `displayDialogAsync` will also close. If the taskpane supports pinning and was pinned by the user, it will be un-pinned. |
| A module extension | No effect. |

## Syntax

```js
Office.context.ui.closeContainer();
```

## Returns
void