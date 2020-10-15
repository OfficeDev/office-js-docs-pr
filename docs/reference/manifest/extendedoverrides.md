---
title: ExtendedOverrides element in the manifest file
description: Specifies the URLs for a JSON-formatted extension of the manifest.
ms.date: 11/06/2020
localization_priority: Normal
---

# Description element

Specifies the full URL for a JSON-formatted file that extends the manifest.

**Add-in type:** Task pane

## Syntax

```XML
<ExtendedOverrides Url="string" [ResourceUrl="string"] />
```

## Contained in

[OfficeApp](officeapp.md)

## Can contain

_none_

## Attributes

|Attribute|Description|
|:-----|:-----|
|Url (required)| The full URL of the extended overrides JSON file.|
|ResourceUrl (optional) | ????????? |

## Example

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/extended-manifest-overrides.json" />
</OfficeApp>
```