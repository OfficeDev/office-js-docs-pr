---
title: Resources element in the manifest file
description: The Resources element contains icons, strings, and URLs for the VersionOverrides node.
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# Resources element

Contains icons, strings, and URLs for the [VersionOverrides](versionoverrides.md) node. A manifest element specifies a resource by using the **id** of the resource. This helps to keep the size of the manifest manageable, especially when resources have versions for different locales. An **id** must be unique within the manifest and can have a maximum of 32 characters.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

Each resource can have one or more **Override** child elements to define a different resource for a specific locale.

## Child elements

|  Element |  Type  |  Description  |
|:-----|:-----|:-----|
|  [Images](#images)            |  image   |  Provides the HTTPS URL to an image for an icon. |
|  **Urls**                |  url     |  Provides an HTTPS URL location. A URL can have a maximum of 2048 characters. |
|  **ShortStrings** |  string  |  The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters.|
|  **LongStrings**  |  string  | The text for **Description** attributes. Each **String** contains a maximum of 250 characters.|

> [!NOTE]
> You must use Secure Sockets Layer (SSL) for all URLs in the **Image** and **Url** elements.

### Images

Each icon must have three **Images** elements, one for each of the three mandatory sizes:

- 16x16
- 32x32
- 80x80

The following additional sizes are also supported, but not required.

- 20x20
- 24x24
- 40x40
- 48x48
- 64x64

> [!IMPORTANT]
>
> - If this image is your add-in's representative icon, see [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) for size and other requirements.
> - Office Add-ins require the ability to cache image resources for performance purposes. For this reason, the server hosting an image resource must not add any CACHE-CONTROL directives to the response header. These directives result in Office automatically substituting a generic or default image. To force the use of new icons on your development computer, [Clear the Office cache](../../testing/clear-cache.md). To force the use of new icons on your end-user's computers, you must give the new icons a different URL from the old ones.

## Resources examples

```XML
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp32-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp80-icon_default.png" />
        </bt:Image>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
        </bt:Url>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residLabel" DefaultValue="GetData">
          <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residToolTip" DefaultValue="Get data for your document.">
          <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
        </bt:String>
      </bt:LongStrings>
    </Resources>
```

```xml
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER//blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/blue-80.png"/>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
    <!-- other URLs -->
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="groupLabel" DefaultValue="Add-in Demo">
      <bt:Override Locale="ar-sa" Value="<Localized text>" />
    </bt:String>
    <!-- Other short strings -->
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment.">
      <bt:Override Locale="ar-sa" Value="<Localized text>." />
    </bt:String>
    <!-- Other long strings -->
  </bt:LongStrings>
</Resources>
```
