---
title: ExtendedOverrides element in the manifest file
description: Specifies the URLs for a JSON-formatted extension of the manifest.
ms.date: 02/23/2021
ms.localizationpriority: medium
---

# ExtendedOverrides element

Specifies the full URLs for JSON-formatted files that extend the manifest. For detailed information about the use of this element and its descendent elements, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).

**Add-in type:** Task pane

## Syntax

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
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
|Url (required)| The full URL of the extended overrides JSON file. In the future, this value could be a URL template that uses tokens defined by the [Tokens](tokens.md) element. See [Examples](#examples).|
|ResourcesUrl (optional) | The full URL of a file that provides supplemental resources, such as localized strings, for the file specified in the `Url` attribute. This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.|

## Examples

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/extended-manifest-overrides.json"
                     ResourceUrl="https://contoso.com/addin/my-resources.json">
  </ExtendedOverrides>
</OfficeApp>
```

In the future, this value could be a URL template that uses tokens defined by the [Tokens](tokens.md) element. The following is an example.

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
</OfficeApp>
```
