# Resources element

Contains icons, strings, and URLs for the [VersionOverrides](./versionoverrides.md) node. A manifest element specifies a resource by using the **Id** of the resource. This helps to keep the size of the manifest manageable, especially when resources have versions for different locales. An **Id** must be unique within the manifest and has a maximum of 32 characters.

The  **Resources** node defines the following resources. Each resource can have one or more **Override** child elements to define a resource for specific locales.

## Child elements

|  Element |  Type  |  Description  |
|:-----|:-----|:-----|
|  [Images](#images)            |  image   |  Provides the HTTPS URL to an image for an icon. |
|  [Urls](#urls)                |  url     |  Provides an HTTPS URL location. |
|  [ShotStrings](#shortstrings) |  string  |  The text for **Label** and **Title** elements. |
|  [LongStrings](#longstrings)  |  string  | The text for **Description** attributes. |

## Images
Provides the HTTPS URL to an image for an icon. Each icon must have three  **Image** elements, one for each of the three mandatory sizes:
- 16x16
- 32x32
- 80x80

The following additional sizes are also supported, but not required:
- 20x20
- 24x24
- 40x40
- 48x48
- 64x64

> **Important: ** Outlook requires the ability to cache image resources for performance purposes. For this reason, the server hosting an image resource must not add any CACHE-CONTROL directives to the response header. This will result in Outlook automatically substituting a generic or default image.    

## Urls
Provides an HTTPS URL location. A URL can be a maximum of 2048 characters. 

## ShortStrings
The text for  **Label** and **Title** elements. Each **String** contains a maximum of 125 characters.

## LongStrings
The text for  **Description** attributes. Each **String** contains a maximum of 250 characters.

## Resources example 
```xml
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/images/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER/images/blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/images/blue-80.png"/>
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