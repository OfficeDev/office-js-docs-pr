<!-- Add to https://learn.microsoft.com/office/dev/add-ins/develop/dialog-api-in-office-add-ins when modal dialog feature is deployed.  -->

> [!TIP]
> If the child dialog is the [preview modal dialog](modal-dialog.md), then a call of `messageChild` can't be triggered by user interaction with the add-in's task pane or add-in commands, because user interaction is blocked while the modal dialog is open. So, if your dialog use case requires messaging from the parent to the dialog, you will nearly always need to use the non-modal dialog API. However, calling `Office.context.ui.messageParent` in the dialog triggers the `DialogMessageReceived` event in the parent, and code in the handler for that event can call `messageChild`.
