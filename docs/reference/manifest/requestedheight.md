# RequestedHeight element

Specifies the initial height (in pixels) of a content add-in or mail add-in. 

**Add-in type:** Content, Mail

## Syntax

```XML
<RequestedHeight>integer</RequestedHeight>
```

## Contained in

- [DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000
- [DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450
- [ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point