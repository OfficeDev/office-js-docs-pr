---
title: Make your Office Add-in compatible with an existing COM or VSTO add-in
description: Enable compatibility between your Office Add-in and equivalent COM or VSTO add-in.
ms.date: 07/12/2025
ms.localizationpriority: medium
---

# Make your Office Add-in compatible with an existing COM or VSTO add-in

If you have an existing COM add-in or VSTO add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Mac. In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM or VSTO add-in. In these situations, your COM or VSTO add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.

> [!NOTE]
> Because COM and VSTO add-ins can only be installed on Windows, this article refers to COM and VSTO add-ins collectively as "Windows-only" add-ins.

[!INCLUDE [new-outlook-vsto-com-support](../includes/new-outlook-vsto-com-support.md)]

You can configure your Office Add-in so that when the equivalent Windows-only add-in is already installed on a user's computer, Office on Windows runs the Windows-only add-in instead of the Office Add-in. The Windows-only add-in is called "equivalent" because Office will seamlessly transition between the Windows-only add-in and the Office Add-in according to which one is installed on a user's computer.

[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

## Specify an equivalent Windows-only add-in

### Obtain the ID of the Windows-only add-in

Before you can specify an equivalent Windows-only add-in, you must first identify its name with these steps:

1. Open Windows Registry Editor on any computer where the Windows-only add-in is installed.
1. Go to **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\\*<Office application\>*\Addins** or **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\\*<Office application\>*\Addins**, where *\<Office application\>* is Excel, Outlook, PowerPoint, or Word. For example, **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\Excel\Addins**.
1. Copy the name of the registry key associated with the Windows-only add-in you need. Note that the names are case-sensitive.

### Configure the manifest

> [!IMPORTANT]
> Applies to Excel, Outlook, PowerPoint, and Word.

To enable compatibility between your Office Add-in and the Windows-only add-in, identify the equivalent Windows-only add-in in the [manifest](add-in-manifests.md) of your Office Add-in using the name that you obtained in the section [Obtain the ID of the Windows-only add-in](#obtain-the-id-of-the-windows-only-add-in). Then, Office on Windows will use the Windows-only add-in instead of the Office Add-in, if they're both installed. The configuration depends on the type of manifest that is being used.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

The following example shows the portion of the manifest that specifies a Windows-only add-in as an equivalent add-in. The value of the [`"alternates.prefer.comAddin.progId"`](/microsoft-365/extensibility/schema/extension-alternate-versions-array-prefer-com-addin#progid) property identifies the Windows-only add-in.

> [!NOTE]
> Although the property is called "comAddin", it actually refers to any Windows-only add-in, regardless of whether the add-in is COM or VSTO. Similarly, the term "progId" is usually associated with only COM add-ins, but in the manifest it refers to the name of the Windows-only add-in.

```json
"extensions" [
  ...
  "alternates" [
    {
      "prefer": {
        "comAddin": {
          "progId": "ContosoAddin"
        }
      }
    }
  ]
]
```

# [Add-in only manifest](#tab/xmlmanifest)

The following example shows the portion of the manifest that specifies a Windows-only add-in as an equivalent add-in. The value of the `ProgId` element identifies the Windows-only add-in and the [EquivalentAddins](/javascript/api/manifest/equivalentaddins) element must be positioned immediately before the closing `VersionOverrides` tag.

> [!NOTE]
> Although the term "ProgId" is usually associated with only COM add-ins, in the manifest it refers to the name of the Windows-only add-in, regardless of whether the add-in is COM or VSTO.

> [!IMPORTANT]
> Use "COM" as the value of the `<Type>` element for both COM and VSTO add-ins.

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

---

> [!TIP]
>
> - For information about XLL UDF compatibility with an Excel add-in that contains custom functions, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).
> - For Outlook only, an admin of a computer can specify an equivalent Windows-only add-in by using a Group Policy. For guidance, see [Configure the Group Policy setting for Outlook add-ins](#configure-the-group-policy-setting-for-outlook-add-ins).

## Configure the Group Policy setting for Outlook add-ins

A user with administrator powers on their computer can configure a Windows-only add-in to be the equivalent of an *Outlook* add-in by using the **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed** Group Policy setting on the computer. This can be achieved even when the manifest of the Outlook add-in doesn't have the equivalent add-in markup. Then, classic Outlook on Windows uses the Windows-only add-in instead of the Outlook add-in if they're both installed. The following are the steps to configure the policy.

1. Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030), paying attention to the tool's **Install Instructions**.
1. Open the Local Group Policy Editor (**gpedit.msc**).
1. Navigate to **User Configuration** > **Administrative Templates**  > **Microsoft Outlook 2016** > **Miscellaneous**.
1. Select the setting **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed**.
1. Open the link to edit the policy setting.
1. In the dialog **Outlook web add-ins to deactivate**:
    1. Set **Value name** to the ID of the Outlook add-in's manifest. In the add-in only manifest, use the value of the `<ID>` element. In the unified manifest, use the value of the `id` property in the root of the manifest. Do *not* add curly braces `{}` around the entry.
    1. Set **Value** to the name of the equivalent Windows-only add-in. See the section [Obtain the ID of the Windows-only add-in](#obtain-the-id-of-the-windows-only-add-in).
    1. Select **OK** to put the update into effect.

    ![The "Outlook web add-ins to deactivate" dialog.](../images/outlook-deactivate-gpo-dialog.png)

## Equivalent behavior for users

When an equivalent Windows-only add-in is specified, Office on Windows won't display your Office Add-in's user interface (UI) if the equivalent Windows-only add-in is installed. But Office only hides the ribbon buttons of the Office Add-in and doesn't prevent installation. Therefore, your Office Add-in will still appear in the following locations within the UI.

- Under **My add-ins**.
- As an entry on the ribbon manager (Excel, Word, and PowerPoint only).

> [!NOTE]
> Specifying an equivalent Windows-only add-in in the manifest has no effect on other platforms, like Office on the web or on Mac.

The following scenarios describe what happens depending on how the user acquires the Office Add-in.

### AppSource acquisition of an Office Add-in

If a user acquires the Office Add-in from AppSource and the equivalent Windows-only add-in is already installed, then Office will:

1. Install the Office Add-in.
2. Hide the Office Add-in UI on the ribbon.
3. Display a call-out for the user that points out the Windows-only add-in ribbon button.

### Centralized deployment of Office Add-in

If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent Windows-only add-in is already installed, the user must restart Office before they'll see any changes. After Office restarts, it will:

1. Install the Office Add-in.
2. Hide the Office Add-in UI on the ribbon.
3. Display a call-out for the user that points out the Windows-only add-in ribbon button.

### Document shared with embedded Office Add-in

If a user has the Windows-only add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:

1. Prompt the user to trust the Office Add-in.
2. If trusted, the Office Add-in will install.
3. Hide the Office Add-in UI on the ribbon.

## Other Windows-only add-in behavior

### Excel, PowerPoint, Word

If a user uninstalls the equivalent Windows-only add-in, then Office on Windows restores the Office Add-in UI.

After you specify an equivalent Windows-only add-in for your Office Add-in, Office stops processing updates for your Office Add-in. To acquire the latest updates for the Office Add-in, the user must first uninstall the Windows-only add-in.

### Outlook

The Windows-only add-in must be connected when Outlook is started in order for the corresponding Outlook add-in to be disabled.

If the Windows-only add-in is then disconnected during a subsequent Outlook session, the Outlook add-in will likely remain disabled until Outlook is restarted.

## See also

- [Make your Custom Functions compatible with XLL User Defined Functions](../excel/make-custom-functions-compatible-with-xll-udf.md)
