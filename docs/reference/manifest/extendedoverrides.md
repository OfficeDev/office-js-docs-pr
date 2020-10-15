---
title: ExtendedOverrides element in the manifest file
description: Specifies the URLs for a JSON-formatted extension of the manifest.
ms.date: 11/06/2020
localization_priority: Normal
---

# Description element

Specifies the full URLs for a JSON-formatted files that extend the manifest. 

**Add-in type:** Task pane

## Syntax

```XML
<ExtendedOverrides Url="string" [ResourceUrl="string"] />
```

## Contained in

[OfficeApp](officeapp.md)

## Can contain

|Element|Content|Mail|TaskPane|
|:-----|:-----|:-----|:-----|
|[Tokens](tokens.md)|||x|

## Attributes

|Attribute|Description|
|:-----|:-----|
|Url (required)| The full URL of the extended overrides JSON file. The URL can use tokens defined by the [Tokens](tokens.md) element.|
|ResourceUrl (optional) | The full URL of a file that provides supplemental resources, such as localized strings, for the file specified in the `Url` attribute. The URL can use tokens defined by the [Tokens](tokens.md) element.|

## Example

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/extended-manifest-overrides.json">
    <Tokens>
    <Tokens>
  </ExtendedOverrides>
</OfficeApp>
```