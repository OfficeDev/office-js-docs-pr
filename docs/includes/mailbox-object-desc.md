Outlook add-ins primarily use a subset of the API exposed through the [Mailbox](/javascript/api/outlook/office.mailbox) object. To access the objects and members for use in Outlook add-ins, such as the [Item](/javascript/api/outlook/office.item) object in compose or read mode, use the [mailbox](/javascript/api/office/office.context#office-office-context-mailbox-member) property of the **Context** object to access the **Mailbox** object. The following code is an example.

```js
// Access the Item object.
const item = Office.context.mailbox.item;
```

> [!IMPORTANT]
> When calling `Office.context.mailbox.item` on a message, note that the Reading Pane in the Outlook client must be turned on. For guidance on how to configure the Reading Pane, see [Use and configure the Reading Pane to preview messages](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0).

Additionally, Outlook add-ins can use the following objects.

- **Office** object: for initialization.

- **Context** object: for access to content and display language properties.

For information about using JavaScript in Outlook add-ins, see [Outlook add-ins](../outlook/outlook-add-ins-overview.md). To explore the Outlook JavaScript API, see the [Outlook API reference](/javascript/api/outlook) page.
