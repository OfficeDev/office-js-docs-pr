# Define add-in commands in your manifest

Add-in commands provide an easy way to customize the default Office UI with UI elements that perform actions; for example, you can add custom buttons on the ribbon. To create commands, you add a **[VersionOverrides](../../reference/manifest/versionoverrides.md)** node to an existing manifest. 

When a manifest contains the **VersionOverrides** element, versions of Word, Excel, Outlook, and PowerPoint that support add-in commands will use the information within that element to load the add-in. Earlier versions of Office products that do not support add-in commands will ignore the element.

When client applications recognize the  **VersionOverrides** node, the add-in name appears in the ribbon, not in a task pane or a read/compose pane. The add-in won't appear in both places.
 
## VersionOverrides

The  [VersionOverrides](../../reference/manifest/versionoverrides.md) element is the root element that contains information for the add-in commands implemented by the add-in. It is supported in manifest schema v1.1 and later.

There are two versions of the **VersionOverrides** schema.

| Schema version | Description |
|----------------|-------------|
| 1.0 | Supports add-in commands for desktop versions of Office apps. | 
| 1.1 | Adds support for [pinnable taskpanes](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane) and mobile add-ins. **Note:** Currently only supported by Outlook 2016 for Windows and Outlook for iOS |

An add-in can support multiple versions of the **VersionOverrides** schema by nesting newer versions inside of the previous version. This allows clients to support the newer versions to take advantage of the new features, while allowing older clients to load the older version. For details, see [Implementing multiple versions](../../reference/manifest/versionoverrides.md#implementing-multiple-versions).

The **VersionOverrides** element includes the following child elements:

- [Description](../../reference/manifest/description.md)
- [Requirements](../../reference/manifest/requirements.md)
- [Hosts](../../reference/manifest/hosts.md)
- [Resources](../../reference/manifest/resources.md)
- [VersionOverrides](../../reference/manifest/versionoverrides.md)

The following diagram shows the hierarchy of elements used to define add-in commands. 

![Hierarchy of add-in commands elements in the manifest](../../images/080da303-51c4-4882-b74a-7ba11517c0ad.png)

## Sample manifests

For a sample manifest that implements add-in commands for Word, Excel, and PowerPoint, see [Simple add-in commands sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Simple).

For a sample manifest that implements add-in commands for Outlook, see [Sample manifest file for an Outlook add-in](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).

## Additional resources

- [Add-in commands for Outlook](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook)
    
- [Outlook add-in manifests](https://docs.microsoft.com/outlook/add-ins/manifests)
    
- [Outlook add-in command demo sample](https://github.com/OfficeDev/outlook-add-in-command-demo)
