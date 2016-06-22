# Define add-in commands in your Outlook add-in manifest

To support add-in commands, some additional elements have been added to the add-in manifest v1.1 within the [VersionOverrides](../../../reference/manifest/versionoverrides.md) element. When a manifest contains the **VersionOverrides** element, versions of Outlook that support add-in commands will use the information within that element to load the add-in. Earlier versions of Outlook that do not support add-in commands will ignore the element and continue to use the elements as described in [Outlook add-in manifests](../../outlook/manifests/manifests.md).

When the client application recognizes the  **VersionOverrides** node, the add-in name appears in the ribbon, not in the read/compose pane. The add-in won't appear in both places.
 

## VersionOverrides node

The  [VersionOverrides](../../reference/manifest/versionoverrides.md) element is the root element that contains information for the add-in commands implemented by the add-in. It is supported in manifest schema v1.1 or later but is defined in the VersionOverrides v1.0 schema. 

The VersionOverrides element includes the following child elements:

- [Description](../../reference/manifest/description.md)
- [Requirements](../../reference/manifest/requirements.md)
- [Hosts](../../reference/manifest/hosts.md)
- [Resources](../../reference/manifest/resources.md)

## Rule changes for add-in commands

The following changes affect the rules in the manifest:

- Activation rules are now inside each entry point.
    
- The **ItemIs** attribute of the [Rule](../../../reference/manifest/rule.md) element has been modified. **ItemType** can either be Message or AppointmentAttendee. The **FormType** attribute has been removed.
    
- The **ItemHasKnownEntity** attribute of the [Rule](../../../reference/manifest/rule.md) element has been udpated to accept a string for the EntityType.
    

## Sample manifest

For a sample manifest for an Outlook add-in that includes the VersionOverrides node, see [Sample manifest file for an Outlook add-in](https://gist.github.com/mlafleur/95b7ac030bb7a7ae742527e85a36b095) on GitHub.


## Additional resources


- [Add-in commands for Outlook](../../outlook/add-in-commands-for-outlook.md)
    
- [Outlook add-in manifests](../../outlook/manifests/manifests.md)
    
- [Outlook add-in command demo sample](https://github.com/jasonjoh/command-demo)
