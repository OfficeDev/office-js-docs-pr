# ExtensionPoint element

 Defines where an add-in exposes functionality. The **ExtensionPoint** element is a child element of [FormFactor](./formfactor.md). 

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Yes  | The type of ExtentionPoint being defined.|

## xsi:type
For each form factor, you can define **ExtensionPoint** elements with the following **xsi:type** values, with the exception of the **Module** value which can only be used in the [DesktopFormFactor](./formfactor.md):

- [CustomPane](./custompane.md) 
- [MessageReadCommandSurface](./messagereadcommandsurface.md) 
- [MessageComposeCommandSurface](./messagecomposecommandsurface.md) 
- [AppointmentOrganizerCommandSurface](./appointmentorganizercommandsurface.md) 
- [AppointmentAttendeeCommandSurface](./appointmentattendeecommandsurface.md)
- [Module](./module.md)



