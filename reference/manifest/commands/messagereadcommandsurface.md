# MessageReadCommandSurface

This [extension point](./extensionpoint.md) puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.

## Child elements
|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](./customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

## OfficeTab example
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

##  CustomTab example
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```