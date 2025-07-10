---
title: Extend custom functions with XLL add-ins
description: Enable compatibility with Excel XLL add-ins that have equivalent functionality to your custom functions.
ms.date: 09/20/2024
ms.localizationpriority: medium
---

# Extend custom functions with XLL add-ins

> [!NOTE]
> An XLL add-in is an Excel add-in file with the file extension **.xll**. An XLL file is a type of dynamic link library (DLL) file that can only be opened by Excel. XLL add-in files must be written in C or C++. See [Developing Excel XLLs](/office/client-developer/excel/developing-excel-xlls) to learn more.

If you have existing Excel XLL add-ins, you can build equivalent custom function add-ins using the Excel JavaScript API to extend your solution features to other platforms, such as Excel on the web or on a Mac. However, Excel JavaScript API add-ins don't have all of the functionality available in XLL add-ins. Depending on the functionality your solution uses, the XLL add-in may provide a better experience in Excel on Windows than the Excel JavaScript API add-in.

[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

## Specify equivalent XLL in the manifest

To enable compatibility with an existing XLL add-in, identify the equivalent XLL add-in in the manifest of your Excel JavaScript API add-in. Excel then uses the XLL add-in functions when running on Windows, instead of your Excel JavaScript API add-in custom functions.

To set the equivalent XLL add-in for your custom functions, specify the `FileName` of the XLL file. When the user opens a workbook with functions from the XLL file, Excel converts the functions to compatible functions. The workbook then uses the XLL file when opened in Excel on Windows, but it continues to use custom functions from your Excel JavaScript API add-in when opened on the web or on Mac.

The manifest configuration depends on what type of manifest the add-in uses.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

The following example shows how to specify both a COM add-in and an XLL add-in as equivalents in a unified manifest. Often you specify both. For completeness, this example shows both equivalents in context. They're identified by their [`"alternates.prefer.comAddin.progId"`](/microsoft-365/extensibility/schema/extension-alternate-versions-array-prefer-com-addin#progid) and `"alternates.prefer.xllCustomFunctions.filename"` respectively. For more information on COM add-in compatibility, see [Make your Office Add-in compatible with an existing COM or VSTO add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).

```json
"extensions" [
  ...
  "alternates" [
    {
      "prefer": {
        "comAddin": {
          "progId": "ContosoCOMAddin"
        },
        "xllCustomFunctions": {
          "fileName": "contosofunctions.xll"
        }
      }
    }
  ]
]
```

# [Add-in only manifest](#tab/xmlmanifest)

The following example shows how to specify both a COM add-in and an XLL add-in as equivalents in an Excel JavaScript API add-in only manifest file. Often you specify both. For completeness, this example shows both equivalents in context. They're identified by their `ProgId` and `FileName` respectively. The `EquivalentAddins` element must be positioned immediately before the closing `VersionOverrides` tag. For more information on COM add-in compatibility, see [Make your Office Add-in compatible with an existing COM or VSTO add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>

    <EquivalentAddin>
      <FileName>contosofunctions.xll</FileName>
      <Type>XLL</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

---

> [!NOTE]
> If an Excel JavaScript API add-in declares its custom functions to be compatible with an XLL add-in, changing the manifest at a later time could break a user's workbook because it will change the file format.

## Custom function behavior for XLL compatible functions

An add-in's XLL functions are converted to XLL compatible custom functions when a spreadsheet is opened and there is an equivalent add-in available. On the next save, the XLL functions are written to the file in a compatible mode so that they work with both the XLL add-in and Excel JavaScript API add-in custom functions (when on platforms unsupported by XLL).

The following table compares features across XLL user-defined functions, XLL compatible custom functions, and Excel JavaScript API add-in custom functions.

|         |XLL user-defined function |XLL compatible custom functions |Excel JavaScript API add-in custom function |
|---------|---------|---------|---------|
| **Supported platforms** | Windows | Windows, macOS, web browser | Windows, macOS, web browser |
| **Supported file formats** | XLSX, XLSB, XLSM, XLS | XLSX, XLSB, XLSM | XLSX, XLSB, XLSM |
| **Formula autocomplete** | No | Yes | Yes |
| **Streaming** | Possible via xlfRTD and XLL callback. | Yes | Yes |
| **Localization of functions** | No | No. The Name and ID must match the existing XLL's functions. | Yes |
| **Volatile functions** | Yes | Yes | Yes |
| **Multi-threaded recalculation support** | Yes | Yes | Yes |
| **Calculation behavior** | No UI. Excel can be unresponsive during calculation. | Users see #BUSY! until a result is returned. | Users see #BUSY! until a result is returned. |
| **Requirement sets** | N/A | CustomFunctions 1.1 and later | CustomFunctions 1.1 and later |

## See also

- [Make your Office Add-in compatible with an existing COM or VSTO add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
