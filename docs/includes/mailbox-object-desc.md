Outlook add-ins primarily use a subset of the API exposed through the [Mailbox](/javascript/api/outlook/Office.mailbox) object. To access the objects and members specifically for use in Outlook add-ins, such as the [Item](/javascript/api/outlook/Office.mailbox) object, you use the [mailbox](/javascript/api/outlook/Office.mailbox) property of the **Context** object to access the **Mailbox** object, as shown in the following line of code.

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

Additionally, Outlook add-ins can use the following objects:

-  **Office** object: for initialization.

-  **Context** object: for access to content and display language properties.

-  **RoamingSettings** object: for saving Outlook add-in-specific custom settings to the user's mailbox where the add-in is installed.

For information about using JavaScript in Outlook add-ins, see [Outlook add-ins](../outlook/outlook-add-ins-overview.md).