---
title: EquivalentAddin element in the manifest file
description: Specifies backwards compatibility for an equivalent COM add-in or XLL.
ms.date: 01/04/2022
ms.localizationpriority: medium
---

# EquivalentAddin element

Specifies backwards compatibility for an equivalent COM add-in or XLL.

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**Add-in type:** Task pane, Mail, Custom function

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## Syntax

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## Contained in

[EquivalentAddins](equivalentaddins.md)

## Must contain

[Type](type.md)

## Can contain

[ProgId](progid.md)
[FileName](filename.md)

## Remarks

To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements. To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.

## See also

- [Make your custom functions compatible with XLL user-defined functions](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Make your Office Add-in compatible with an existing COM add-in](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)