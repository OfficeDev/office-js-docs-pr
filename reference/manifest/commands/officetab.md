# Office tab
On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in. 

The default tab is limited to one group per add-in. 

## Child elements
|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  OfficeTab  | Yes |  Always set to `TabDefault`.  |
|  Group      | Yes |  Defines a Group of commands.  |

## OfficeTab
Required. The pre-existing tab to use. Currently, the  **id** attribute can only be "TabDefault".

## Group
A group of user interface extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. It is a string with a maximum of 125 characters. See [Group element](./group.md).

## OfficeTab example
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```