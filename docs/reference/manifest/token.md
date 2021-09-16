---
title: Token element in the manifest file
description: Specifies a token or wildcard that can be used with URL templates in the manifest.
ms.date: 11/06/2020
ms.localizationpriority: medium
---


# Token element

Defines an individual URL token. For more information about the use of this element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).

**Add-in type:** Task pane

## Syntax

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## Contained in

[Tokens](tokens.md)

## Can contain

|Element|Content|Mail|TaskPane|
|:-----|:-----|:-----|:-----|
|[Override](override.md)|||x|

## Attributes

|Attribute|Description|
|:-----|:-----|
|DefaultValue|Default value for this token if no condition in any child `<Override>` element matches.|
|Name|Token name. This name is user-defined. The type of the token is determined by the type attribute.|
|xsi:type|Defines the kind of Token. This attribute should be set to one of:  `"RequirementsToken"`,  or  `"LocaleToken"`.|

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