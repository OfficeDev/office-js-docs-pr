---
title: Dialog boxes in Office Add-ins
description: ''
ms.date: 2/28/2019
localization_priority: Priority
---


# Dialog boxes in Office Add-ins
 
Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.

*Figure 1. Typical layout for a dialog box*

![An example image displaying a typical layout for a dialog box](../images/overview-with-app-dialog.png)

## Best practices

|**Do**|**Don't**|
|:-----|:--------|
|<ul><li>Include a descriptive title that includes your add-in name along with the current task.</li></ul>|<ul><li>Don't append your company name to the title.</li></ul>|
||<ul><li>Don't open a dialog box unless the scenario requires it.</li></ul>|

## Implementation

For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.

## See also

- [Dialog object](https://docs.microsoft.com/javascript/api/office/office.dialog)
- [UX design patterns for Office Add-ins](../design/ux-design-pattern-templates.md)


