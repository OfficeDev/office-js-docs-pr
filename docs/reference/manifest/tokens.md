---
title: Tokens element in the manifest file
description: Specifies tokens or wildcards that can be used with URL templates in the the manifest.
ms.date: 11/06/2020
ms.localizationpriority: medium
---

# Tokens element

Defines tokens that could be used in template URLs. For more information about the use of this element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).

**Add-in type:** Task pane

## Syntax

```XML
<Tokens></Tokens>
```

## Contained in

[ExtendedOverrides](extendedoverrides.md)

## Must contain

|Element|Content|Mail|TaskPane|
|:-----|:-----|:-----|:-----|
|[Token](token.md)|||x|

## Example

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