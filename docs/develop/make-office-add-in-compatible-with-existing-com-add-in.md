---
title: Make your Excel add-in compatible with an existing COM add-in
description: 'Enable compatibility with an equivalent COM add-in that has the same functionality as your Excel add-in'
ms.date: 05/06/2019
localization_priority: Normal
---

# Make your Office Add-in compatible with an existing COM add-in (preview)

If you have an existing COM add-in, you can build equivalent functionality in your Excel add-in to extend your solution features to other platforms such as online or macOS. However, Excel add-ins don't have all of the functionality available in COM add-ins. Your COM add-in may provide a better experience than the Excel add-in on Windows.

You can configure your Excel add-in so that when an equivalent COM add-in is already installed on the user's computer, Office runs the COM add-in instead of your Excel add-in. The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Excel add-in depending on which is installed on Windows.

[!include[COM add-in and XLL UDF compatibility requirements note](../includes/xll-compatibility-note.md)]

## Specify an equivalent COM add-in in the manifest

To enable compatibility with an existing COM add-in, identify the equivalent COM add-in in the manifest of your Excel add-in. Then Office will use the COM add-in instead of your Excel add-in when running on Windows.

Specify the `ProgID` of the equivalent COM add-in. Office will then use the COM add-in UI instead of your Excel add-in's UI when the COM add-in is installed.

The following example shows how to specify both a COM add-in and an XLL as equivalent. Often you will specify both so for completeness this example shows both in context. They are identified by their `ProgID` and `FileName` respectively. For more information on XLL compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).

```xml
<VersionOverrides>
...
<EquivalentAddins>
  <EquivalentAddin>
    <ProgID>ContosoCOMAddin</ProgID>
    <Type>COM</Type>
  </EquivalentAddin>

  <EquivalentAddin>
    <FileName>contosofunctions.xll</FileName>
    <Type>XLL</Type>
  </EquivalentAddin>
<EquivalentAddins>
...
</VersionOverrides>
```

## Equivalent behavior for users

When an equivalent COM add-in is specified in the Excel add-in manifest, Office suppresses your Excel add-in's UI on Windows when the equivalent COM add-in is installed. This does not affect your Excel add-in's UI on other platforms like online or macOS. Office only hides the ribbon buttons and does not prevent installation. Therefore your Excel add-in will still appear in the following UI locations:

- Under **My add-ins** because it is technically installed.
- As an entry in the ribbon manager.

The following scenarios describe what happens depending on how the user acquires the Excel add-in.

### AppSource acquisition of an Excel add-in

If a user downloads the Excel add-in from AppSource, and the equivalent COM add-in is already installed, then Office will:

1. Install the Excel add-in.
2. Hide the Excel add-in UI in the ribbon.
3. Display a call-out for the user that points out the COM add-in ribbon button.

### Centralized deployment of Excel add-in

If an admin deploys the Excel add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user needs to restart Office before they will see any changes. After Office restarts, it will:

1. Install the Excel add-in.
2. Hide the Excel add-in UI in the ribbon.
3. Display a call-out for the user that points out the COM add-in ribbon button.

### Document shared with embedded Excel add-in

If a user has the COM add-in installed, and then gets a shared document with the embedded Excel add-in, then when they open the document, Office will:

1. Prompt the user to trust the Excel add-in.
2. If trusted, the Excel add-in will install.
3. Hide the Excel add-in UI in the ribbon.

## Other COM add-in behavior

If a user uninstalls the COM add-in, then Office restores the Excel add-in UI on Windows for the equivalent installed Excel add-in.

Once you specify an equivalent COM add-in for your Excel add-in, Office stops processing updates for your Excel add-in. The user must uninstall the COM add-in order to get the latest updates for the Excel add-in.

## See also

- [Make your Custom Functions compatible with XLL User Defined Functions](../excel/make-custom-functions-compatible-with-xll-udf.md)
