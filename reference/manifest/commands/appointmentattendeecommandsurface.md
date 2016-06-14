# AppointmentAttendeeCommandSurface

This puts buttons on the ribbon for the form that's displayed to the attendee of the meeting. 

## Child elements
|  Element |  Description  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Adds the command(s) to the default ribbon tab.  |
|  [CustomTab](./customtab.md) |  Adds the command(s) to the custom ribbon tab.  |

## OfficeTab example
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

##  CustomTab example
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```