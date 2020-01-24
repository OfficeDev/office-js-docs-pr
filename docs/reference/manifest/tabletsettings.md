---
title: TabletSettings element in the manifest file
description: ''
ms.date: 01/13/2020
localization_priority: Normal
---

# TabletSettings element

Specifies control settings that apply when your mail add-in is used on a tablet.

> [!IMPORTANT]
> The `TabletSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows. To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](/outlook/add-ins/outlook-mobile-addins).

**Add-in type:** Mail

## Syntax

```XML
<Form xsi:type="ItemRead">
   <!--website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </DesktopSettings>
   <TabletSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://website.html" />
   </PhoneSettings>
</Form>
```

## Contained in

[Form](form.md)

