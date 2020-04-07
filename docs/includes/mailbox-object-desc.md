Outlook add-ins primarily use the APIs exposed through the [Mailbox](/javascript/api/outlook/Office.mailbox) object. To access the objects and members specifically for use in Outlook add-ins, such as the [Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) object, you use the [mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) property of the **Context** object to access the **Mailbox** object, as shown in the following line of code.

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

Additionally, Outlook add-ins can use the following objects:

-  **Office** object: for initialization.

-  **Context** object: for access to content and display language properties.

-  **RoamingSettings** object: for saving Outlook add-in-specific custom settings to the user's mailbox where the add-in is installed.

For information about using the Outlook JavaScript API, see [Outlook add-ins](../outlook/outlook-add-ins-overview.md).