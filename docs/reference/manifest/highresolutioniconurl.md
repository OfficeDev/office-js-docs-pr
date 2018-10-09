# HighResolutionIconUrl element

Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.

**Add-in type:** Content, Task pane, Mail

## Syntax

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## Can contain

[Override](override.md)

## Attributes

|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|string (URL)|required|Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.|

## Remarks

For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI . For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.

The image must be in one of the following file formats at a recommended resolution of 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP or TIFF. For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).
