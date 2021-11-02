---
title: IconUrl element in the manifest file
description: The IconUrl element specifies the URL of the image that represents your Office Add-in in the insertion UX and Office Store.
ms.date: 03/30/2021
ms.localizationpriority: medium
---

# IconUrl element

Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.

**Add-in type:** Content, Task pane, Mail

## Syntax

```XML
<IconUrl DefaultValue="string" />
```

## Can contain

[Override](override.md)

## Attributes

|Attribute|Type|Required|Description|
|:-----|:-----|:-----|:-----|
|DefaultValue|string|required|Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.|

## Remarks

For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook on the web). For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI. For all add-in types, the icon is also used in [AppSource](https://appsource.microsoft.com), if you publish your add-in to AppSource.

The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF. For content and task pane apps, the image specified must be 32 x 32 pixels. For mail apps, the image resolution must be 64 x 64 pixels. You should also specify an icon for use with Office client applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element. For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).

Changing the value of the `IconUrl` element at runtime is not currently supported.
