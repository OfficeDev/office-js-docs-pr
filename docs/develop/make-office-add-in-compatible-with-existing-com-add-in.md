---
title: Make your Office Add-in compatible with an existing COM add-in
description: Enable compatibility between your Office Add-in and similar COM or VSTO add-in.
ms.date: 11/04/2025
ms.localizationpriority: medium
---

# Make your Office Add-in compatible with an existing COM or VSTO add-in

If you have an existing COM add-in or VSTO add-in, you can build nearly equivalent functionality in your Office Web Add-in, thereby enabling your solution to run on other platforms such as in Office on the web or on Mac. In this scenario, the two add-ins are called "equivalent" and you can use the manifest of the Office Web Add-in to control which add-in is available to users on Windows computers (or give users the choice).

> [!NOTE]
> Because COM and VSTO add-ins can only be installed on Windows, this article refers to COM and VSTO add-ins collectively as "Windows-only" add-ins.

 As a general rule, add-ins should be treated as equivalent only when the following conditions are met, but there may be exceptional scenarios where add-ins are that don't meet these conditions can be treated as equivalent.

- They're both created by the same developer, typically the entity specified in the [`"developer"`](/microsoft-365/extensibility/schema/root-developer) property of the unified manifest or the [ProviderName](/javascript/api/manifest/providername) element in the add-in only manifest.
- They both are designed to be installed on the same Office application or applications, and they address the same workload with mainly overlapping functionality.
- They have identical, or highly similar, public names, and user interfaces including control names and icons.


[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

[!INCLUDE [new-outlook-vsto-com-support](../includes/new-outlook-vsto-com-support.md)]

Because the Windows-only add-in and the Office Add-in may not have exactly the same features, there are three scenarios you need to consider.

- Your Office Add-in does not provide all of the functionality that's available in the corresponding Windows-only add-in. In this scenario, you should [configure the Office add-in to be hidden](#hide-the-office-add-in) on any Windows computer on which the Windows-only add-in is installed. 
- Your Office Add-in provides more and better functionality than the corresponding Windows-only add-in. In this scenario, you should [configure the Office Add-in so that it detects the presence of the Windows-only add-in and hides it](#hide-the-windows-only-add-in).
- Neither add-in would be regarded as better than the other by all your users. In this scenario, you should configure the Office Add-in to [give each user on a Windows computer the choice of which add-in is hidden](#give-the-user-the-choice-of-which-add-in-to-use).

## Obtain the name of the Windows-only add-in

Before you can configure the Office Add-in's manifest, you must first identify the name of the Windows-only add-in in the Windows Registry with these steps:

1. Open Windows Registry Editor on any computer where the Windows-only add-in is installed.
1. Go to **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\\*<Office application\>*\Addins** or **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\\*<Office application\>*\Addins**, where *\<Office application\>* is Excel, Outlook, PowerPoint, or Word. For example, **HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\Excel\Addins**.
1. Copy the name of the registry key associated with the Windows-only add-in you need. Note that the names are case-sensitive.

## Hide the Office Add-in

You can configure your Office Add-in so that when the equivalent Windows-only add-in is already installed on a user's computer, Office on Windows runs the Windows-only add-in instead of the Office Add-in. If the Windows-only add-in is ever uninstalled, Office will automatically activate the Office Add-in the next time the host Office application is started. If the Windows-only add-in is installed *after* the Office Add-in, then when the Office application is restarted, the Office Add-in will be hidden. For more information, see [Behavior for users when the Office Add-in is hidden](#behavior-for-users-when-the-office-add-in-is-hidden). 

### Configure the manifest

> [!IMPORTANT]
> Applies to Excel, Outlook, PowerPoint, and Word.

To enable compatibility between your Office Add-in and the Windows-only add-in, identify the equivalent Windows-only add-in in the [manifest](add-in-manifests.md) of your Office Add-in using the name that you obtained in the section [Obtain the name of the Windows-only add-in](#obtain-the-name-of-the-windows-only-add-in). Then, Office on Windows will use the Windows-only add-in instead of the Office Add-in, if they're both installed. The configuration depends on the type of manifest that is being used.

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

> [!TIP]
> If you're unable to specify the `alternates.prefer` property in the manifest of your Outlook web add-in, you must configure Group Policy instead. This only applies to Outlook. For guidance, see [Configure the Group Policy setting for Outlook add-ins](#configure-the-group-policy-setting-for-outlook-add-ins).

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

### Configure the Group Policy setting for Outlook add-ins

A user with administrator powers on their computer can configure a Windows-only add-in to be the equivalent of an *Outlook* add-in by using the **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed** Group Policy setting on the computer. This can be achieved even when the manifest of the Outlook add-in doesn't have the equivalent add-in markup. Then, classic Outlook on Windows uses the Windows-only add-in instead of the Outlook add-in if they're both installed. The following are the steps to configure the policy.

1. Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030), paying attention to the tool's **Install Instructions**.
1. Open the Local Group Policy Editor (**gpedit.msc**).
1. Navigate to **User Configuration** > **Administrative Templates**  > **Microsoft Outlook 2016** > **Miscellaneous**.
1. Select the setting **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed**.
1. Open the link to edit the policy setting.
1. In the dialog **Outlook web add-ins to deactivate**:
    1. Set **Value name** to the ID of the Outlook add-in's manifest. In the add-in only manifest, use the value of the `<ID>` element. In the unified manifest, use the value of the `id` property in the root of the manifest. Do *not* add curly braces `{}` around the entry.
    1. Set **Value** to the name of the equivalent Windows-only add-in. See the section [Obtain the name of the Windows-only add-in](#obtain-the-name-of-the-windows-only-add-in).
    1. Select **OK** to put the update into effect.

    :::image type="content" source="../images/outlook-deactivate-gpo-dialog.png" alt-text="The 'Outlook web add-ins to deactivate' dialog.":::

### Behavior for users when the Office Add-in is hidden

When an equivalent Windows-only add-in is specified, Office on Windows won't display your Office Add-in's user interface (UI) if the equivalent Windows-only add-in is installed. But Office only hides the ribbon buttons of the Office Add-in and doesn't prevent installation. Therefore, your Office Add-in will still appear in the following locations within the UI.

- Under **My add-ins**.
- As an entry on the ribbon manager (Excel, Word, and PowerPoint only).

> [!NOTE]
> Specifying an equivalent Windows-only add-in in the manifest has no effect on other platforms, like Office on the web or on Mac.

The following scenarios describe what happens depending on how the user acquires the Office Add-in.

#### Microsoft Marketplace acquisition of the Office Add-in

If a user acquires the Office Add-in from Microsoft Marketplace and the equivalent Windows-only add-in is already installed, then Office will:

1. Install the Office Add-in.
2. Hide the Office Add-in UI on the ribbon.
3. Display a call-out for the user that points out the Windows-only add-in ribbon button.

#### Centralized deployment of the Office Add-in

If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent Windows-only add-in is already installed, the user must restart Office before they'll see any changes. After Office restarts, it will:

1. Install the Office Add-in.
2. Hide the Office Add-in UI on the ribbon.
3. Display a call-out for the user that points out the Windows-only add-in ribbon button.

#### Document shared with the Office Add-in embedded

If a user has the Windows-only add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:

1. Prompt the user to trust the Office Add-in.
2. If trusted, the Office Add-in will install.
3. Hide the Office Add-in UI on the ribbon.

### Other Windows-only add-in behavior

#### Excel, PowerPoint, Word

If a user uninstalls the equivalent Windows-only add-in, then Office on Windows restores the Office Add-in UI.

After you specify an equivalent Windows-only add-in for your Office Add-in, Office stops processing updates for your Office Add-in. To acquire the latest updates for the Office Add-in, the user must first uninstall the Windows-only add-in.

#### Outlook

The Windows-only add-in must be connected when Outlook is started in order for the corresponding Outlook add-in to be disabled.

If the Windows-only add-in is then disconnected during a subsequent Outlook session, the Outlook add-in will likely remain disabled until Outlook is restarted.

## Hide the Windows-only add-in

[!INCLUDE [Support limits on the equivalent add-ins feature](../includes/equivalent-add-ins-limitations.md)]

### Configure the manifest

The details to configure this feature depend on which type of manifest is being used by the Office Add-in.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

To disable the Windows-only add-in, take the following steps.

1. If it doesn't already exist, create an [`"extensions.alternates.hide.windowsExtensions"`](/microsoft-365/extensibility/schema/extension-alternate-versions-array-hide-windows-extensions) property.
1. Add a child `effect` property to the `windowsExtensions` property and set its value to  **disableWithNotification**.
1. Add a child [`"comAddin"`](/microsoft-365/extensibility/schema/extension-alternate-versions-array-hide-windows-extensions-com-addin) property to the `windowsExtensions` property.
1. Add a child `progIds` array property to the `comAddin` property.
1. Add the names of all the Windows-only add-ins that you want to hide to the `progIds` array. For information about how to obtain these names, see [Obtain the name of the Windows-only add-in](#obtain-the-name-of-the-windows-only-add-in).

> [!NOTE]
> The string "com" is in the name `comAddin` for historical reasons. In the manifest, the property refers to either COM or VSTO add-ins. Similarly, the term "progIds" is usually associated with only COM add-ins, but names of VSTO add-ins can also be included in the `progIds` array.

The following example shows how the JSON should look.

```json
"extensions" [
  {   
    ...
    "alternates": [
      {
        "hide": {
          "windowsExtensions": {
            "effect": "disableWithNotification",
            "comAddin": {
              "progIds": [
                "ContosoAddin"
              ]
            }
          }
        }
      }
    ]
  }
]
```

# [Add-in only manifest](#tab/xmlmanifest)

To disable the Windows-only add-in, take the following steps.

1. Create an [EquivalentAddins](/javascript/api/manifest/equivalentaddins) as the very last child of the **\<VersionOverrides/>** element.
1. Add an **\<EquivalentAddin\>** child for every Windows-only add-in that you want to hide.
1. For each **\<EquivalentAddin\>** element, add a **\<ProgId/>** and **\<Type/>** child, *in that order*.
1. Set the **\<ProgId/>** element to the name of a Windows-only add-in that you want to hide. For information about how to obtain these names, see [Obtain the name of the Windows-only add-in](#obtain-the-name-of-the-windows-only-add-in).
1. Set the **\<Type/>** element to "COM". 

   > [!NOTE]
   > The **\<Type\>** value is "COM" for historical reasons, but it refers to either a COM or a VSTO add-in. Similarly, the **\<ProgId\>** value can be either the progId of a COM add-in or the name of a VSTO add-in.

1. Add an **\<Effect\>** element *as the very last* child of the **\<EquivalentAddins\>** element, and set its value to **DisableWithNotification**.

The following is an example.

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
    <Effect>DisableWithNotification</Effect>
  </EquivalentAddins>
</VersionOverrides>
```

---

### Behavior for users when the Windows-only add-in is hidden

When a Windows-only add-in is hidden, Office on Windows won't display its user interface (UI) if the equivalent Office Add-in is installed. A popup dialog notifies the user that this is happening. But Office only hides the Windows-only add-in and doesn't prevent installation.

## Give the user the choice of which add-in to use

[!INCLUDE [Support limits on the equivalent add-ins feature](../includes/equivalent-add-ins-limitations.md)]

### Configure the manifest

The details to configure this feature depend on which type of manifest is being used by the Office Add-in.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

To give the user a choice of which add-in to hide, take the following steps.

1. If it doesn't already exist, create an [`"extensions.alternates.hide.windowsExtensions"`](/microsoft-365/extensibility/schema/extension-alternate-versions-array-hide-windows-extensions) property.
1. Add a child `effect` property to the `windowsExtensions` property and set its value to  **userOptionToDisable**.
1. Add a child [`"comAddin"`](/microsoft-365/extensibility/schema/extension-alternate-versions-array-hide-windows-extensions-com-addin) property to the `windowsExtensions` property.
1. Add a child `progIds` array property to the `comAddin` property.
1. Add the names of all the equivalent Windows-only add-ins to the `progIds` array. For information about how to obtain these names, see [Obtain the name of the Windows-only add-in](#obtain-the-name-of-the-windows-only-add-in).

> [!NOTE]
> The string "com" is in the name `comAddin` for historical reasons. In the manifest, the property refers to either COM or VSTO add-ins. Similarly, the term "progIds" is usually associated with only COM add-ins, but names of VSTO add-ins can also be included in the `progIds` array.

The following example shows how the JSON should look.

```json
"extensions" [
  {   
    ...
    "alternates": [
      {
        "hide": {
          "windowsExtensions": {
            "effect": "userOptionToDisable",
            "comAddin": {
              "progIds": [
                "ContosoAddin"
              ]
            }
          }
        }
      }
    ]
  }
]
```

# [Add-in only manifest](#tab/xmlmanifest)

To disable the Windows-only add-in, take the following steps.

1. Create an [EquivalentAddins](/javascript/api/manifest/equivalentaddins) as the very last child of the **\<VersionOverrides/>** element.
1. Add an **\<EquivalentAddin\>** child for every Windows-only add-in that you want to hide.
1. For each **\<EquivalentAddin\>** element, add a **\<ProgId/>** and **\<Type/>** child, *in that order*.
1. Set the **\<ProgId/>** element to the name of a Windows-only add-in whose hidden/visible status you want the user to decide. For information about how to obtain these names, see [Obtain the name of the Windows-only add-in](#obtain-the-name-of-the-windows-only-add-in).
1. Set the **\<Type/>** element to "COM". 

   > [!NOTE]
   > The **\<Type\>** value is "COM" for historical reasons, but it refers to either a COM or a VSTO add-in. Similarly, the **\<ProgId\>** value can be either the progId of a COM add-in or the name of a VSTO add-in.

1. Add an **\<Effect\>** element *as the very last* child of the **\<EquivalentAddins\>** element, and set its value to **UserOptionToDisable**.

The following is an example.

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
    <Effect>UserOptionToDisable</Effect>
  </EquivalentAddins>
</VersionOverrides>
```

---

### Behavior for users when they have the choice of which add-in is hidden

If you configure the manifest to give the user the choice, then the user sees a dialog similar to the following when they install the Office Add-in. The **Learn more** link in the dialog points to the following help page that provides information to help the user make the decision: [Resolve version conflicts for Office Add-ins](https://support.microsoft.com/office/1632ec51-82ed-4f8e-90b4-a246cbccccde).

:::image type="content" source="../images/com-preference-prompt.png" alt-text="A dialog titled 'Contoso JS Add-in Installer'. The first paragraph says, 'Contoso JS Add-in has identified the following older versions of the add-in that may cause conflicts:'. Below this paragraph is a single bulleted paragraph that says 'COM Add-in name: Contoso'. Below this is a paragraph that says, 'Do you want to disable the older add-in versions? (If you choose No, you will still get the new version if you open Office on Mac or Office on the web.)'. Below this is a link labelled 'Learn more'. Below this are two buttons labelled Yes and No.":::

If the user chooses **Yes** and Office successfully disables and hides the Windows-only add-in, then a dialog similar to the following opens that advises the user to restart the Office application.

:::image type="content" source="../images/com-preference-disable-success.png" alt-text="A dialog titled 'The Older Version Add-in Disabled'. The first paragraph says, 'The older version of the add-in has been successfully disabled.' Below this, a paragraph says, 'To ensure the changes take effect, we recommend restarting this Office application. You can continue without restarting, but some updates may not be applied until the next launch.' Below this is an OK button.":::

If the user chooses **Yes** but Office can't disable the Windows-only add-in for any reason, then a dialog similar to the following opens that advises the user to manually disable the Windows-only add-in.

:::image type="content" source="../images/com-preference-disable-failure.png" alt-text="A dialog titled 'Please Disable Add-in Manually'. The first paragraph says, 'The older version of the add-in could not be fully disabled. To complete the process, please manually disable the COM add-in in the Office application where it was installed.' Below this are instructions labelled 'Here's how'. The instructions read as follows: '1. Go to File, Options, Add-ins. 2. In the Manage dropdown at the bottom, choose the add-in type and click Go. 3. In the dialog, uncheck the add-in to disable it. 4. Restart the app to apply changes (optional but recommended).' Below this is an OK button.":::

## See also

- [Make your Custom Functions compatible with XLL User Defined Functions](../excel/make-custom-functions-compatible-with-xll-udf.md)
