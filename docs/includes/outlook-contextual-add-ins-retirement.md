> [!IMPORTANT]
> Entity-based contextual Outlook add-ins will be retired in Q2 of 2024. The work to retire this feature will start in May and continue until the end of June. After June, contextual add-ins will no longer be able to detect entities in mail items to perform tasks on them. The following APIs will also be retired.
>
> - [Office.context.mailbox.item.getEntities()](/javascript/api/requirement-sets/outlook/requirement-set-1.13/office.context.mailbox.item#methods)
> - [Office.context.mailbox.item.getEntitiesByType(entityType)](/javascript/api/requirement-sets/outlook/requirement-set-1.13/office.context.mailbox.item#methods)
> - [Office.context.mailbox.item.getFilteredEntitiesByName(name)](/javascript/api/requirement-sets/outlook/requirement-set-1.13/office.context.mailbox.item#methods)
> - [Office.context.mailbox.item.getSelectedEntities()](/javascript/api/requirement-sets/outlook/requirement-set-1.13/office.context.mailbox.item#methods)
>
> To help minimize potential disruptions, the following will still be supported after entity-based contextual add-ins are retired.
>
> - An alternative implementation of the **Join Meeting** button, which is activated by online meeting add-ins, is being developed. Once support for entity-based contextual add-ins ends, online meeting add-ins will automatically transition to the alternative implementation to activate the **Join Meeting** button.
> - Regular expression rules will continue to be supported after entity-based contextual add-ins are retired. We recommend updating your contextual add-in to use regular expression rules as an alternative solution. For guidance on how to implement these rules, see [Use regular expression activation rules to show an Outlook add-in](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md).
>
> For more information, see [Retirement of entity-based contextual Outlook add-ins](https://devblogs.microsoft.com/microsoft365dev/retirement-of-entity-based-contextual-outlook-add-ins/).
