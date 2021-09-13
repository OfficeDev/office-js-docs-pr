---
title: SupportUrl element in the manifest file
description: The SupportUrl element specifies the URL of a page that provides support information for your add-in.
ms.date: 10/09/2018
ms.localizationpriority: medium
---

# SupportUrl element

Specifies the URL of a page that provides support information for your add-in.

## Syntax

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## Contained in

[OfficeApp](officeapp.md)

## Can contain

|  Element | Required | Description  |
|:-----|:-----|:-----|
|  [Override](override.md)   | No | Specifies the setting for additional locale urls |

## Attributes

|Attribute|Type|Required|Description|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|required|Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.|
