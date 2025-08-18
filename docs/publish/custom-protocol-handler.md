---
title: Trust custom protocol handlers that launch add-ins
description: How to use group policies for protocol handler trust in the registry to launch add-ins.
ms.topic: how-to
ms.date: 08/15/2025
ms.localizationpriority: medium
---

# Trust custom protocol handlers to launch add-ins

Let your Office Add-in launch from a custom protocol handler (like `mailto:` or your own scheme) without prompting users for consent. This is useful when you want a seamless experience for users, or when your organization needs to centrally manage which add-ins launch from which protocols.

> [!IMPORTANT]
> This article applies to Windows only. Support for this feature starts with Office Version 2408 (Build 17928.20018).

## How protocol handler trust works

A protocol handler lets an app or add-in launch from a URI (for example, clicking a `mailto:` link opens your email client). Office add-ins also launch this way. By default, users are prompted to trust each add-in and protocol pair. As an admin, pre-approve or block these pairs using group policy and the Windows registry.

## Registry key format

To automatically trust a custom protocol handler for an add-in, create a registry key at one of the following locations (replace `<add-in ID>` with your add-in's [manifest ID](/javascript/api/manifest/id) or unified manifest [`id`](/microsoft-365/extensibility/schema/root#id).

- **Current user (64-bit Office)**: `HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\office\16.0\WEF\ProtocolHandlers\<add-in ID>`
- **Local machine (64-bit Office)**: `HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\office\16.0\WEF\ProtocolHandlers\<add-in ID>`

Add the following values to the key.

- **Name:** Protocol name (for example, `mailto`)
- **Type:** REG_SZ
- **Data:** `Allow` or `Block`

## Set group policies

Admins use group policy to manage protocol handler trust across the organization. The following sample files show how to set up these policies.

### Sample ADMX file

```xml
<?xml version="1.0" encoding="utf-16"?> 

<policyDefinitions xmlns="http://www.microsoft.com/GroupPolicy/PolicyDefinitions" revision="1.0" schemaVersion="1.0"> 
  <policyNamespaces> 
    <target prefix="osf16" namespace="osf16.Office.Microsoft.Policies.Windows" /> 
    <using prefix="windows" namespace="Microsoft.Policies.Windows" /> 
  </policyNamespaces> 
  <supersededAdm fileName="osf16" /> 
  <resources minRequiredRevision="1.0" /> 
  <categories> 
    <category name="L_MicrosoftOfficeAddins" displayName="$(string.L_MicrosoftOfficeAddins)" /> 
    <category name="L_ProtocolHandlers" displayName="$(string.L_ProtocolHandlers)"> 
      <parentCategory ref="L_MicrosoftOfficeAddins" /> 
    </category> 
  </categories> 
  <policies> 
    <!-- Protocol ListBox --> 
    <policy 
      name="L_Protocols" 
      class="Machine" 
      displayName="$(string.L_Protocols)" 
      explainText="$(string.L_ProtocolsExplain)" 
      key="Software\Policies\Microsoft\Office\16.0\WEF\ProtocolHandlers\[add-in id]" 
      presentation="$(presentation.L_CustomProtocolTaskpaneProtocols)"> 
      <parentCategory ref="L_ProtocolHandlers" /> 
      <supportedOn ref="windows:SUPPORTED_Windows7" /> 
      <elements> 
        <list id="L_ProtocolsListBox" explicitValue="true" additive="true"></list> 
      </elements>
    </policy> 
  </policies>
</policyDefinitions> 
```

### Sample ADML file

```xml
<?xml version="1.0" encoding="utf-16"?> 
<policyDefinitionResources xmlns="http://www.microsoft.com/GroupPolicy/PolicyDefinitions" revision="1.0" schemaVersion="1.0"> 
  <displayName>Microsoft Office Add-Ins</displayName> 
  <description>Microsoft Office Add-Ins</description> 
  <resources> 
    <stringTable> 
      <string id="L_MicrosoftOfficeAddins">Microsoft Office Add-ins</string> 
      <string id="L_ProtocolHandlers">Protocol Handlers</string> 
      <string id="L_Protocols">[add-in name]</string> 
      <string id="L_ProtocolsExplain">Defines URL protocol behavior. </string> 
    </stringTable> 
    <presentationTable> 
      <presentation id="L_Protocols"> 
        <listBox refId="L_ProtocolsListBox">Protocols</listBox> 
      </presentation> 
    </presentationTable> 
  </resources> 
</policyDefinitionResources> 
```

### Sample .REG file

The following example shows how to configure the registry directly using a .REG file.

```text
Windows Registry Editor Version 5.00
[HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\16.0\WEF\ProtocolHandlers]

[HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\16.0\WEF\ProtocolHandlers\[add-in id]]
"protocol1"="Allow"
"protocol2"="Block"
```

## See also

- [Office Add-ins manifest](../develop/add-in-manifests.md)
- [Use policy settings to manage privacy controls for Microsoft 365 Apps for enterprise](/microsoft-365-apps/privacy/manage-privacy-controls)
