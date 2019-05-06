---
title: 'Extend custom functions with XLL user-defined functions'
description: 'Enable compatibility with Excel XLL user-defined functions that have equivalent functionality to your custom functions' 
ms.date: 05/03/2019
localization_priority: Normal
---

# Extend custom functions with XLL user-defined functions (preview)

If you have existing Excel XLLs, you can build equivalent custom functions in an Office Add-in to extend your solution features to other platforms such as online or macOS. However, Office Add-ins don't have all of the functionality available in XLLs. Depending on the functionality your solution uses, the XLL may provide a better experience than the Office Add-in custom functions on Excel for Windows.

[!include[COM add-in and XLL UDF compatibility note](../includes/xll-compatibility-note.md)]

You can configure your Office Add-in so that when an equivalent XLL is already installed on the user's computer, Excel runs the XLL instead of your Office Add-in custom functions. The XLL is called equivalent because Excel will seamlessly transition between the XLL and the Office Add-in custom functions depending on which is installed on Windows.

## Specify equivalent XLL in the manifest

To enable compatibility with an existing XLL, identify the equivalent XLL in the manifest of your Office Add-in. Then Excel will use the XLL's functions instead of your Office Add-in custom functions when running on Windows.

To set the equivalent XLL for your custom functions, specify the `FileName` of the XLL. When the user opens a workbook with functions from the XLL, Excel converts the functions to compatible functions. The workbook then uses the XLL when opened in Excel on Windows, and it will use custom functions from your Office Add-in when opened online or on macOS.

The following example shows how to specify both a COM add-in and an XLL as equivalent. Often you will specify both so for completeness this example shows both in context. They are identified by their `ProgID` and `FileName` respectively. For more information on COM add-in compatibility, see [Make your Office Add-in compatible with an existing COM add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).

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

> [!NOTE]
> If an add-in declares its custom functions to be XLL compatible, changing the manifest at a later time could break a user’s workbook because it will change the file format.

## Office Add-in updates

Once you specify an equivalent XLL for your Office Add-in, Excel stops processing updates for your Office Add-in. The user must uninstall the XLL in order to get the latest updates for the Office Add-in.

## Custom function behavior for XLL compatible functions

When a spreadsheet is opened that contains XLL functions for which there is also an equivalent add-in, the XLL's functions are converted to XLL compatible custom functions. On the next save, they are written to the file in a compatible mode so that they work with both the XLL and Office Add-in custom functions (when on other platforms).

The following table compares features across XLL user-defined functions, XLL compatible custom functions, and Office Add-in custom functions.

|         |XLL user-defined function |XLL compatible custom functions |Office Add-in custom function |
|---------|---------|---------|---------|
| Supported platforms | Windows | Windows, macOS, Excel online | Windows, macOS, Excel online |
| Supported file formats | XLSX, XLSB, XLSM, XLS | XLSX, XLSB, XLSM | XLSX, XLSB, XLSM |
| Formula autocomplete | No | Yes | Yes |
| Streaming | Possible via xlfRTD and XLL callback. | No | Yes |
| Localization of functions | No | No. The Name and ID must match the existing XLL's functions. | Yes |
| Volatile functions | Yes | Yes | Yes |
| Multi-threaded recalculation support | Yes | Yes | Yes |
| Calculation behavior | No UI. Excel can be unresponsive during calculation. | Users will see #BUSY! until a result is returned. | Users will see #BUSY! until a result is returned. |
| Requirement sets | N/A | CustomFunctions 1.1 and later | CustomFunctions 1.1 and later |

## See also

- [Make your Office Add-in compatible with an existing COM add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Custom functions best practices](custom-functions-best-practices.md)
- [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
