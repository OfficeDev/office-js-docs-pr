---
title: Get and set categories
description: How to manage categories on mailbox and item.
ms.date: 04/12/2024
ms.topic: how-to
ms.localizationpriority: medium
---

# Get and set categories

In Outlook, a user can apply categories to messages and appointments as a means of organizing their mailbox data. The user defines the master list of color-coded categories for their mailbox, and can then apply one or more of those categories to any message or appointment item. Each [category](/javascript/api/outlook/office.categorydetails) in the master list is represented by the name and [color](/javascript/api/outlook/office.mailboxenums.categorycolor) that the user specifies. You can use the Office JavaScript API to manage the categories master list on the mailbox and the categories applied to an item.

> [!NOTE]
> Support for this feature was introduced in requirement set 1.8. See [clients and platforms](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.

## Manage categories in the master list

Only categories in the master list on your mailbox are available for you to apply to a message or appointment. You can use the API to add, get, and remove master categories.

> [!IMPORTANT]
> For the add-in to manage the categories master list, it must request the **read/write mailbox** permission in the manifest. The markup varies depending on the type of manifest.
>
> - **Add-in only manifest**: Set the **\<Permissions\>** element to **ReadWriteMailbox**.
> - **Unified manifest for Microsoft 365**: Set the `"name"` property of an object in the [`"authorization.permissions.resourceSpecific"`](/microsoft-365/extensibility/schema/root-authorization-permissions#resourcespecific) array to `"Mailbox.ReadWrite.User"`.

### Add master categories

The following example shows how to add a category named "Urgent!" to the master list by calling [addAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-addasync-member(1)) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

```js
const masterCategoriesToAdd = [
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

The following example shows how to get the list of categories by calling [getAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-getasync-member(1)) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        const masterCategories = asyncResult.value;
        console.log("Master categories:");
        masterCategories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### Remove master categories

The following example shows how to remove the category named "Urgent!" from the master list by calling [removeAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-removeasync-member(1)) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member).

```js
const masterCategoriesToRemove = ["Urgent!"];

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
> In Outlook on the web or [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627), you can't use the API to manage categories on a message in Compose mode.

### Add categories to an item

The following example shows how to apply the category named "Urgent!" to the current item by calling [addAsync](/javascript/api/outlook/office.categories#outlook-office-categories-addasync-member(1)) on `item.categories`.

```js
const categoriesToAdd = ["Urgent!"];

Office.context.mailbox.item.categories.addAsync(categoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories");
    } else {
        console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### Get an item's categories

The following example shows how to get the categories applied to the current item by calling [getAsync](/javascript/api/outlook/office.categories#outlook-office-categories-getasync-member(1)) on `item.categories`.

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        const categories = asyncResult.value;
        console.log("Categories:");
        categories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### Remove categories from an item

The following example shows how to remove the category named "Urgent!" from the current item by calling [removeAsync](/javascript/api/outlook/office.categories#outlook-office-categories-removeasync-member(1)) on `item.categories`.

```js
const categoriesToRemove = ["Urgent!"];

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
