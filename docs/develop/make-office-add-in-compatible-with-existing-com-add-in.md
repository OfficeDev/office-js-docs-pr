---
title: Make your Office Add-in compatible with an existing COM add-in
description: 'Enable compatibility with an equivalent COM add-in that has the same functionality as your Office Add-in'
ms.date: 06/13/2019
localization_priority: Normal
---

# Make your Office Add-in compatible with an existing COM add-in (preview)

If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Office on Mac. In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in. In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.

You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in. The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.

> [!NOTE]
> This feature is currently in preview and not supported for use in production environments. It's available in Excel, Word, and PowerPoint version 16.0.11629.20214 or later. To access this build, you must have an Office 365 subscription and join the [Office Insider](https://products.office.com/office-insider) program at the **Insider** level.

## Specify an equivalent COM add-in in the manifest

To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in. Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.

The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in. The value of the `ProgID` element identifies the COM add-in.

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgID>ContosoCOMAddin</ProgID>
      <Type>COM</Type>
    </EquivalentAddin>
  <EquivalentAddins>
  ...
</VersionOverrides>
```

> [!TIP]
> For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).

## Equivalent behavior for users

When an equivalent COM add-in is specified in the Office Add-in manifest, Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed. Office only hides the ribbon buttons of the Office Add-in and does not prevent installation. Therefore your Office Add-in will still appear in the following locations within the UI:

- Under **My add-ins**
- As an entry in the ribbon manager

> [!NOTE]
> Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or Office for Mac.

The following scenarios describe what happens depending on how the user acquires the Office Add-in.

### AppSource acquisition of an Office Add-in

If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:

1. Install the Office Add-in.
2. Hide the Office Add-in UI in the ribbon.
3. Display a call-out for the user that points out the COM add-in ribbon button.

### Centralized deployment of Office Add-in

If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes. After Office restarts, it will:

1. Install the Office Add-in.
2. Hide the Office Add-in UI in the ribbon.
3. Display a call-out for the user that points out the COM add-in ribbon button.

### Document shared with embedded Office Add-in

If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:

1. Prompt the user to trust the Office Add-in.
2. If trusted, the Office Add-in will install.
3. Hide the Office Add-in UI in the ribbon.

## Other COM add-in behavior

If a user uninstalls the COM add-in, then Office on Windows restores the Office Add-in UI for the equivalent installed Office Add-in.

After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in. To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.

## See also

- [Make your Custom Functions compatible with XLL User Defined Functions](../excel/make-custom-functions-compatible-with-xll-udf.md)
