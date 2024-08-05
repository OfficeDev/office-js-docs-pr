---
title: Trust custom protocol handlers that launch add-ins
description: How to use group policies for protocol handler trust in the registry to launch add-ins.
ms.topic: how-to
ms.date: 08/05/2024
ms.localizationpriority: medium
---

# Trust custom protocol handlers to launch add-ins

Protocol handlers are registered with the operating system to allow an app to be launched from a URI (for example, how `mailto:` launchs an email client). Add-ins can also be launched from protocol handlers. This article explains how to automatically trust these custom protocol handlers by using admin group policies.

Every add-in and protocol pair needs to be trusted. This either comes in the form end-user consent or admin group policies. Similarly, admins can block certain add-in and protocol pairs.

> [!IMPORTANT]
> This information about trusting custom protocol handlers that launch add-ins through the registry only pertains to Windows.

## Registry key format

To automatically trust a custom protocol handler that launches an add-in, create a registry key at one of the following locations. Note that `<add-in id>` refers to the [Id element](/javascript/api/manifest/id) specified in your add-in's manifest.

- Current user (64-bit Office): `HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\office\16.0\WEF\ProtocolHandlers\<add-in ID>`
- Local machine (64-bit Office): `HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\office\16.0\WEF\ProtocolHandlers\<add-in ID>`

Give the key the following values.

- **Name**: The protocol name based on the URI. For example, `mailto`.
- **Type**: REG_DWORD
- **Data**: ["Allow", "Block"]

## Set group policies

The following sample files show how admins define and trust these custom protocol handlers across their organization.

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

```text
Windows Registry Editor Version 5.00 

[HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\16.0\WEF\ProtocolHandlers] 
 
[HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\16.0\WEF\ProtocolHandlers\[add-in id]] 
"protocol1"="Allow" 
"protocol2"="Block" 
```
