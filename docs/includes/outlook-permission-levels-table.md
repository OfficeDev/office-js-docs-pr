|**Permission level</br>canonical name**|**XML manifest name**|**Teams manifest name**|**Summary description**|
|:-----|:-----|:-----|:-----|
|**restricted**|Restricted|MailboxItem.Restricted.User|Allows use of entities, but not regular expressions. |
|**read item**|ReadItem|MailboxItem.Read.User|In addition to what is allowed in **restricted**, it allows:<ul><li>regular expressions</li><li>Outlook add-in API read access</li><li>getting the item properties and the callback token</li><li>writing custom properties</li></ul> |
|**read/write item**|ReadWriteItem|MailboxItem.ReadWrite.User|In addition to what is allowed in **read item**, it allows:<ul><li>full Outlook add-in API access except `makeEwsRequestAsync`</li><li>setting the item properties</li></ul> |
|**read/write mailbox**|ReadWriteMailbox|Mailbox.ReadWrite.User|In addition to what is allowed in **read/write item**, it allows:<ul><li>creating, reading, writing items and folders</li><li>sending items</li><li>calling [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)</li></ul> |

Permissions are declared in the manifest. The markup varies depending on the type of manifest.

- **XML manifest**:  Use the **\<Permissions\>** element.
- **Teams manifest (preview)**: Use the "name" property of an object in the "authorization.permissions.resourceSpecific" array.

> [!NOTE]
>
> - There is a supplementary permission needed for add-ins that use the append-on-send feature. With the XML manifest, you specify the permission in the [ExtendedPermissions](/javascript/api/manifest/extendedpermissions) element. For details, see [Implement append-on-send in your Outlook add-in](../outlook/append-on-send.md). With the Teams manifest (preview), you specify this permission with the name **Mailbox.AppendOnSend.User** in an additional object in the "authorization.permissions.resourceSpecific" array.
> - There is a supplementary permission needed for add-ins that use shared folders. With the XML manifest, you specify the permission by setting the [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) element to `true`. For details, see [Enable shared folders and shared mailbox scenarios in an Outlook add-in](../outlook/delegate-access.md). With the Teams manifest (preview), you specify this permission with the name **Mailbox.SharedFolder** in an additional object in the "authorization.permissions.resourceSpecific" array.
