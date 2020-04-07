---
title: Get and set categories
description: 'How to manage categories on mailbox and item'
ms.date: 01/14/2020
localization_priority: Normal
---

# Get and set categories

In Outlook, a user can apply categories to messages and appointments as a means of organizing their mailbox data. The user defines the master list of color-coded categories for their mailbox, and can then apply one or more of those categories to any message or appointment item. Each [category](/javascript/api/outlook/office.categorydetails) in the master list is represented by the name and [color](/javascript/api/outlook/office.mailboxenums.categorycolor) that the user specifies. You can use the Office JavaScript API to manage the categories master list on the mailbox and the categories applied to an item.

> [!NOTE]
> Support for this feature was introduced in requirement set 1.8. See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.

## Manage categories in the master list

Only categories in the master list on your mailbox are available for you to apply to a message or appointment. You can use the API to add, get, and remove master categories.

> [!IMPORTANT]
> For the add-in to manage the categories master list, you must set the `Permissions` node in the manifest to `ReadWriteMailbox`.

### Add master categories

The following example shows how to add a category named "Urgent!" to the master list by calling [addAsync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).

```js
var masterCategoriesToAdd = [
    {
        "displayName": "Urgent!",
        "color": Office.MailboxEnums.CategoryColor.Preset0
    }
];

Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories to master list");
    } else {
        console.log("masterCategories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### Get master categories

The following example shows how to get the list of categories by calling [getAsync](/javascript/api/outlook/office.mastercategories#getasync-options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        var masterCategories = asyncResult.value;
        console.log("Master categories:");
        masterCategories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### Remove master categories

The following example shows how to remove the category named "Urgent!" from the master list by calling [removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).

```js
var masterCategoriesToRemove = ["Urgent!"];

Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories from master list");
    } else {
        console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## Manage categories on a message or appointment

You can use the API to add, get, and remove categories for a message or appointment item.

> [!IMPORTANT]
> Only categories in the master list on your mailbox are available for you to apply to a message or appointment. See the earlier section [Manage categories in the master list](#manage-categories-in-the-master-list) for more information.
>
> In Outlook on the web, you can't use the API to manage categories on a message in Read mode.

### Add categories to an item

The following example shows how to apply the category named "Urgent!" to the current item by calling [addAsync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-) on `item.categories`.

```js
var categoriesToAdd = ["Urgent!"];

Office.context.mailbox.item.categories.addAsync(categoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories");
    } else {
        console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### Get an item's categories

The following example shows how to get the categories applied to the current item by calling [getAsync](/javascript/api/outlook/office.categories#getasync-options--callback-) on `item.categories`.

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        var categories = asyncResult.value;
        console.log("Categories:");
        categories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### Remove categories from an item

The following example shows how to remove the category named "Urgent!" from the current item by calling [removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-) on `item.categories`.

```js
var categoriesToRemove = ["Urgent!"];

Office.context.mailbox.item.categories.removeAsync(categoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories");
    } else {
        console.log("categories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## See also

- [Outlook permissions](understanding-outlook-add-in-permissions.md)
- [Permissions element in the manifest](../reference/manifest/permissions.md)
