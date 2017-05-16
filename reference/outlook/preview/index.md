# Outlook add-in API Preview requirement set

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

> **Note**: This documentation is for a **preview** [requirement set](tutorial-api-requirement-sets.html). This requirement set is not fully implemented yet, and clients will not accurately report support for it. You should not specify this requirement set in your add-in manifest. Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.

The Preview Requirement set includes all of the features of [Requirement set 1.5](../1.5/index.md). 

## Features in preview

The following features are in preview.

- [Event.completed](https://dev.office.com/reference/add-ins/outlook/preview/Event?product=outlook&version=preview#completedoptions) - A new optional parameter `options`, which is a dictionary with one valid value `allowEvent`. This value is used to cancel execution of an event.
- [Office.context.mailbox.item.getSelectedEntities](https://dev.office.com/reference/add-ins/outlook/preview/Office.context.mailbox.item?product=outlook&version=preview#getselectedentities--entities) - Added a new function that gets the entities found in a highlighted match a user has selected. Highlighted matches apply to contextual add-ins.
- [Office.context.mailbox.item.getSelectedRegExMatches](https://dev.office.com/reference/add-ins/outlook/preview/Office.context.mailbox.item?product=outlook&version=preview#getselectedregexmatches--object) - Added a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to contextual add-ins.

## Additional resources

- [Outlook add-ins](../../../docs/outlook/outlook-add-ins.md)
- [Outlook add-in code samples](https://developer.microsoft.com/en-us/outlook/code-samples)
- [Get started](https://docs.microsoft.com/en-us/outlook/add-ins/addin-tutorial)
