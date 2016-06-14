# MessageComposeCommandSurface

This puts buttons on the ribbon for add-ins using mail compose form. 

## Child elements
|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](./customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

## OfficeTab example
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

##  CustomTab example
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```