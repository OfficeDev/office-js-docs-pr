---
title: Make your Office Add-in compatible with an existing COM add-in
description: Enable compatibility between your Office Add-in and equivalent COM add-in.
ms.date: 06/20/2024
ms.localizationpriority: medium
---

# Make your Office Add-in compatible with an existing COM add-in

If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Mac. In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in. In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.

[!INCLUDE [new-outlook-vsto-com-support](../includes/new-outlook-vsto-com-support.md)]

You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in. The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed on a user's computer.

[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

## Specify an equivalent COM add-in

### Obtain the ProgId of a COM add-in

Before you can specify an equivalent COM add-in, you must first identify its `ProgId`. To obtain the `ProgId` of a COM add-in:

1. Open Windows Registry Editor on the computer where the COM add-in is installed.
1. Go to **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\\*<application\>*\Addins** or **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\\*<application\>*\Addins**.
1. Copy the name of the registry key associated with the COM add-in you need. Note that the names are case-sensitive.

### Configure the manifest

> [!IMPORTANT]
> Applies to Excel, Outlook, PowerPoint, and Word.

To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in. Then, Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.

The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in. The value of the `ProgId` element identifies the COM add-in and the [EquivalentAddins](/javascript/api/manifest/equivalentaddins) element must be positioned immediately before the closing `VersionOverrides` tag.

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md). Not applicable for Outlook.

### Configure the Group Policy setting

> [!IMPORTANT]
> Applies to Outlook only.

To declare compatibility between your Outlook web add-in and COM add-in, identify the equivalent COM add-in in the **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed** Group Policy setting. This must be configured on the user's machine. Then, Outlook on Windows will use the COM add-in instead of the web add-in, if they're both installed.

1. Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030), paying attention to the tool's **Install Instructions**.
1. Open the Local Group Policy Editor (**gpedit.msc**).
1. Navigate to **User Configuration** > **Administrative Templates**  > **Microsoft Outlook 2016** > **Miscellaneous**.
1. Select the setting **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed**.
1. Open the link to edit the policy setting.
1. In the dialog **Outlook web add-ins to deactivate**:
    1. Set **Value name** to the `Id` found in the web add-in's manifest. **Important**: Do *not* add curly braces `{}` around the entry.
    1. Set **Value** to the `ProgId` of the equivalent COM add-in.
    1. Select **OK** to put the update into effect.

    ![The "Outlook web add-ins to deactivate" dialog.](../images/outlook-deactivate-gpo-dialog.png)

## Equivalent behavior for users

When an [equivalent COM add-in is specified](#specify-an-equivalent-com-add-in), Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed. Office only hides the ribbon buttons of the Office Add-in and doesn't prevent installation. Therefore, your Office Add-in will still appear in the following locations within the UI.

- Under **My add-ins**.
- As an entry on the ribbon manager (Excel, Word, and PowerPoint only).

> [!NOTE]
> Specifying an equivalent COM add-in in the manifest has no effect on other platforms, like Office on the web or on Mac.

The following scenarios describe what happens depending on how the user acquires the Office Add-in.

### AppSource acquisition of an Office Add-in

If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:

1. Install the Office Add-in.
2. Hide the Office Add-in UI on the ribbon.
3. Display a call-out for the user that points out the COM add-in ribbon button.

### Centralized deployment of Office Add-in

If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes. After Office restarts, it will:

1. Install the Office Add-in.
2. Hide the Office Add-in UI on the ribbon.
3. Display a call-out for the user that points out the COM add-in ribbon button.

### Document shared with embedded Office Add-in

If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:

1. Prompt the user to trust the Office Add-in.
2. If trusted, the Office Add-in will install.
3. Hide the Office Add-in UI on the ribbon.

## Other COM add-in behavior

### Excel, PowerPoint, Word

If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.

After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in. To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.

### Outlook

The COM add-in must be connected when Outlook is started in order for the corresponding web add-in to be disabled.

If the COM add-in is then disconnected during a subsequent Outlook session, the web add-in will likely remain disabled until Outlook is restarted.

## See also

- [Make your Custom Functions compatible with XLL User Defined Functions](../excel/make-custom-functions-compatible-with-xll-udf.md)
