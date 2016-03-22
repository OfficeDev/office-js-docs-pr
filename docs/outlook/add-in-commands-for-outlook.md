
# Add-in commands for Outlook


Outlook add-in commands provide ways to initiate specific add-in actions from the ribbon by adding buttons or drop-down menus. This lets users access add-ins in a simple, intuitive, and unobtrusive way. Because they offer increased functionality in a seamless manner, you can use add-in commands to create more engaging solutions.

Add-in commands are only available for add-ins that do not use [ItemHasAttachment](https://msdn.microsoft.com/en-us/library/fp123567.aspx%28Office.15%29.aspx), [ItemHasKnownEntity](https://msdn.microsoft.com/en-us/library/fp161166.aspx%28Office.15%29.aspx), or [ItemHasRegularExpressionMatch](https://msdn.microsoft.com/en-us/library/fp142215.aspx%28Office.15%29.aspx) rules to limit the types of items they activate on. However, add-ins can present different commands depending on whether the currently selected item is a message or appointment, and can choose to appear in read or compose scenarios. Using add-in commands if possible is a [best practice](../../docs/design/add-in-development-best-practices.md).


## Creating the add-in command

Add-in commands are declared in the add-in manifest in the  **VersionOverrides** element. This element is an addition to the manifest schema v1.1 that ensures backward compatibility. In a client that doesn't support **VersionOverrides**, existing add-ins will continue to function as they did without add-in commands.

The  **VersionOverrides** manifest entries specify many things for the add-in, such as the host, types of controls to add to the ribbon, the text, the icons, and any associated functions. For more information, see [Define add-in commands in your Outlook add-in manifest](../outlook/manifests/define-add-in-commands.md). 

When an add-in needs to provide status updates, such as progress indicators or error messages, it must do so through the [notification APIs](../../reference/outlook/NotificationMessages.md). The processing for the notifications must also be defined in a separate.md file that is specified in the  **FunctionFile** node of the manifest.

Developers should define icons for all needed sizes so that the add-in commands will adjust smoothly along with the ribbon. The icon sizes are 80 x 80 pixels, 32 x 32 pixels, and 16 x 16 pixels.


## How do add-in commands appear?

An add-in command appears on the ribbon as a button. When a user installs an add-in, its commands appear in the UI as a group of buttons labeled with the add-in name. This can either be on the ribbon's default tab or on a custom tab. For messages, the default is either the  **Home** or **Message** tab. For the calendar, the default is the **Meeting**,  **Meeting Occurrence**,  **Meeting Series**, or  **Appointment** tab. On the default tab, each add-in can have one ribbon group with up to 6 commands. On custom tabs, the add-in can have up to 10 groups, each with 6 commands. Add-ins are limited to only one custom tab.

As the ribbon gets more crowded, the add-in commands will adjust (collapse) in an orderly way. In all cases, the add-in commands for an add-in will be grouped together.


![Screenshots showing add-in command buttons in a normal and a collapsed state.](../../images/6fcb64d8-9598-41d1-8944-f6d1f6d2edb6.png)

When an add-in command is added to an add-in, the add-in name is removed from the app bar unless the add-in also includes a [custom pane Outlook add-in](../outlook/custom-pane-outlook-add-ins.md). Only the add-in command button on the ribbon remains.


## What UX shapes exist for add-in commands?

The UX shape for an add-in command consists of a ribbon tab in the host application that contains buttons that can perform various functions. Currently, three UI shapes are supported:


- A button that executes a JavaScript function
    
- A button that launches a task pane
    
- A button that shows a drop-down menu with one or more buttons of the other two types
    

### Executing a JavaScript function

Use an add-in command button that executes a JavaScript function for scenarios where the user doesn't need to make any additional selections to initiate the action. This can be for actions such as track, remind me, or print, or scenarios when the user wants more in-depth information from a service. 


![A button that executes a function on the Outlook ribbon.](../../images/23ab1de3-3ec4-41a5-ba5b-30b11d464e0c.png)


### Launching a task pane

Use an add-in command button to launch a task pane for scenarios where a user needs to interact with an add-in for a longer period of time. For example, the add-in requires changes to settings or the completion of many fields. 

The default width of the vertical task pane is 300 px. The vertical task pane can be resized in both the Outlook Explorer and inspector. The pane can be resized in the same way the to-do pane and list view resize.


![A button that opens a task pane on the Outlook ribbon.](../../images/c8e03da8-9f71-4f9b-813f-1cdea43d433c.png)

This screenshot shows an example of a vertical task pane. The pane opens with the name of the add-in command in the top left corner. Users can use the **X** button in the upper-right corner of the pane to close the add-in when they are finished using it. This pane will not persist across messages. All UI elements rendered in the task pane, aside from the add-in name and the close button, are provided by the add-in.

If a user chooses another add-in command that opens a task pane, the task pane is replaced with the recently used command. If a user chooses an add-in command button that executes a function, or drop-down menu while the task pane is open, the action will be completed and the task pane will remain open.


### Drop-down menu

A drop-down menu add-in command defines a static list of buttons. The buttons within the menu can be any mix of buttons that execute a function or buttons that open a task pane. Submenus are not supported.


![A button that drops down a menu on the Outlook ribbon.](../../images/3eff90d6-7822-4fdb-9153-68f754c0c746.png)


## Where do add-in commands appear in the UI?

Add-in commands are supported for four scenarios:


### Reading a message

When the user is reading a message, add-in commands added to the default tab appear on the  **Home** tab when viewing the message in the reading pane and in the **Message** tab for a pop-out read form.


### Composing a message

When the user is composing a message, add-in commands added to the default tab appear on the  **Message** tab.


### Creating or viewing an appointment or meeting as the organizer

When creating or viewing an appointment or meeting as the organizer, add-in commands added to the default tab appear on the  **Meeting**,  **Meeting Occurrence**,  **Meeting Series**, or  **Appointment** tabs on pop-out forms. However, if the user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon.


### Viewing a meeting as an attendee

When viewing a meeting as an attendee, add-in commands added to the default tab appear on the  **Meeting**,  **Meeting Occurrence**, or  **Meeting Series** tabs on pop-out forms. However, if a user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon


## Additional resources

- [Define add-in commands in your Outlook add-in manifest](../outlook/manifests/define-add-in-commands.md)
    
- [Add-in Command Demo Outlook Add-in](https://github.com/jasonjoh/command-demo.aspx)
    
