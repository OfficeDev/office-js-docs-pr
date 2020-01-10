---
title: Enabled element in the manifest file
description: ''
ms.date: 01/10/2020
localization_priority: Normal
---

# Enabled element

Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches. The **Enabled** element is a child element of [Control](control.md). If it is omitted, the default is `true`. 

The parent control can also be programmatically enabled and disabled. For more information, see [Enable and Disable Add-in Commands](/office/dev/add-ins/design/disable-add-in-commands).

## Example

```xml
<Enabled>false</Enabled>
```

