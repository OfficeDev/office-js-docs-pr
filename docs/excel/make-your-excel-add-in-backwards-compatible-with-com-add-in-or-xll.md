---
title: Make your Excel add-in backwards compatible with and existing COM add-in or Excel XLL
description: 'Enable backwards compatibility with COM add-ins or Excell XLLs that have equivalent functionality to your Office web add-in'
ms.date: 04/22/2019
localization_priority: Normal
---

# Make your Excel add-in backwards compatible with and existing COM add-in or Excel XLL

If you have existing COM add-ins or Excel XLL's, you can build equivalent Office web add-ins to extend your solution features to other platforms such as online or macOS. However Office web add-ins don't have all of the functionality available in COM add-ins and XLLs. Depending on the functionality your solution uses, users may have a better experience by using the original COM add-in or XLL when they use Excel on desktop.

You can configure your Office web add-in so that when an equivalent COM add-in or XLL is already installed on the user's computer, Excel runs the COM add-in or XLL instead of your web add-in.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## Specify equivalent add-ins in the manifest

To enable backwards compatibility, specify the equivalent add-ins in your manifest. You can specify a COM add-in, an XLL, or both. Then Excel will use the COM add-in or XLL instead of your web add-in when running on desktop.

To set the equivalent COM add-in, you specify its `ProgID`. Then Excel will create the mapping between your COM add-in UI and your web add-in UI.

To set the equivalent XLL for your custom functions, specify the `FileName` of the XLL. When the user opens a workbook with functions from the XLL, it will convert the functions to compatible functions. Then the workbook will use the XLL when opened in Excel on desktop, and it will use custom functions from your web add-in when opened online or on macOS.

The following example shows how to specify both a COM add-in and an XLL as equivalent. They are identified by their `ProgID` and `FileName` respectively.

```xml
<VersionOverrides>
...
<EquivalentAddins>  
<EquivalentAddin>  
       <ProgID>{progid}</ProgID>  
       <Type>COM</Type>  
  </EquivalentAddin>  
  
  <EquivalentAddin>  
       <FileName>{filename}.xll</FileName>  
       <Type>XLL</Type>  
  </EquivalentAddin>  
<EquivalentAddins>
...
</VersionOverrides>
```

> [!NOTE]
> If an add-in declares its custom functions to be XLL compatible, changing the manifest at a later time could break a user’s workbook because it will change the file format.

## How the compatibility behavior works for users

When an equivalent COM add-in is specified, Excel will suppress your web add-in UI on Windows when the specified COM add-in is installed. This does not affect your web add-in UI on other platforms like online or macOS. The following scenarios describe what happens depending on how the user acquires the web add-in.

### Office Store acquisition  

If a user downloads the web add-in from the Office store Excel will:

1. Install the web add-in.
2. Hide the web add-in UI in the ribbon.
3. Display a call-out for the user that points out the COM add-in ribbon button.

### Centralized deployment

If an admin deploys the web add-in to their tenant using centralized deployment, then the user needs to restart Excel before they will see any changes.

After they restart Excel, Excel will:

1. Install the web add-in.
2. Hide the web add-in UI in the ribbon.
3. Display a call-out for the user that points out the COM add-in ribbon button.

### Document share

If a user has the COM add-in installed, and then gets a shared document with the ebedded web add-in, then when they open the document:

1. The user will see a trust prompt for the web add-in.
2. If trusted, the web add-in will install.
3. Excel hides the web add-in UI in the ribbon.

### User uninstalls

If a user uninstalls the COM add-in, Excel restores the web add-in UI on Windows if the web add-in was previously installed.

Excel only hides the ribbon buttons and does not prevent installation. Therefore your web add-in will still appear in the following UI locations:

- Under **My Add-ins** because it is technically installed.
- As an entry in the ribbon manager.

### Web add-in updates

Once you specify an equivalent COM addin or XLL for your web add-in, Excel stops processing updates for your web add-in. In order to get the latest updates for the web add-in, the user must uninstall the COM add-in or XLL.

## Custom function behavior for XLL compatible function

When a spreadsheet is opened and contains XLL functions for which there is also an equivalent add-in, the XLL functions are converted to XLL compatible custom functions. They will be saved in the file in a compatible mode such that they will still work with the original XLL, or with custom functions (when on other platforms like online or macOS).

The following table compares features across XLL user defined functions, XLL compatible custom functions, and Office web add-in custom functions.

|         |XLL user defined function |XLL compatible custom functions |Office web add-in custom function |
|---------|---------|---------|---------|
| Supported platforms | Windows | Windows, macOS, Excel online | Windows, macOS, Excel online |
| Supported file formats | XLSX, XLSB, XLSM, XLS | XLSX, XLSB, XLSM | XLSX, XLSB, XLSM |
| Formula autocomplete | No | Yes | Yes |
| Streaming | Possible via xlfRTD and XLL callback. | Yes | Yes |
| Localization of functions | No | No. The Name and ID must match the existing XLL functions. | Yes |
| Volatile functions | Yes | Yes | Yes |
| Multi-threaded recalculation support | Yes | Yes | Yes |
| Calculation behavior | No UI. Excel can be unresponsive during calculation. | Users will see #BUSY! until a result is returned. | Users will see #BUSY! until a result is returned. |
| Requirement sets | N/A | CustomFunctions 1.1 only | CustomFunctions 1.1 and later |

## See also

- [Custom functions best practices](custom-functions-best-practices.md)
- [Custom functions changelog](custom-functions-changelog.md)
- [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
