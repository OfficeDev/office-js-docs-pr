
# Create Outlook add-ins for compose forms

Starting with version 1.1 of the schema for Office Add-ins manifests and v1.1 of office.js, you can create compose add-ins, which are Outlook add-ins activated in compose forms. In contrast with read add-ins (Outlook add-ins that are activated in read mode when a user is viewing a message or appointment), compose add-ins are available in the following user scenarios:


- Composing a new message, meeting request, or appointment in a compose form.
    
- Viewing or editing an existing appointment, or meeting item in which the user is the organizer.
    
     >**Note**  If the user is on the RTM release of Outlook 2013 and Exchange 2013 and is viewing a meeting item organized by the user, the user can find read add-ins available. Starting in the Office 2013 SP1 release, there's a change such that in the same scenario, only compose add-ins can activate and be available.
- Composing an inline response message or replying to a message in a separate compose form.
    
- Editing a response ( **Accept**,  **Tentative**, or  **Decline**) to a meeting request or meeting item.
    
- Proposing a new time for a meeting item.
    
- Forwarding or replying to a meeting request or meeting item.
    
In each of these compose scenarios, any add-in command buttons defined by the add-in are shown. For older add-ins that do not implement add-in commands, users can choose  **Office Add-ins** in the ribbon to open the add-in selection pane, and then choose and start a compose add-in. The following figure shows add-in commands in a compose form.


![Shows an Outlook compose form with add-in commands.](../../images/583023e6-0534-4f17-9791-b91aa8bff07e.png)

The following figure shows the add-in selection pane consisting of two compose add-ins that do not implement add-in commands, activated when the user is composing an inline reply in Outlook.

![Templates mail app activated for composed item](../../images/mod_off15_MailApps_TemplatesAppSelectionPane.png)


## Types of add-ins available in compose mode


Compose add-ins are implemented as [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).


## API features available to compose add-ins



- For activating add-ins in compose forms, see Table 1 in [Specify activation rules in a manifest](../outlook/manifests/activation-rules.md#MailAppDefineRules_Manifest).
    
- [Add and remove attachments to an item in a compose form in Outlook](../outlook/add-and-remove-attachments-to-an-item-in-a-compose-form.md)
    
- [Get and set item data in a compose form in Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Get, set, or add recipients when composing an appointment or message in Outlook](../outlook/get-set-or-add-recipients.md)
    
- [Get or set the subject when composing an appointment or message in Outlook](../outlook/get-or-set-the-subject.md)
    
- [Insert data in the body when composing an appointment or message in Outlook](../outlook/insert-data-in-the-body.md)
    
- [Get or set the location when composing an appointment in Outlook](../outlook/get-or-set-the-location-of-an-appointment.md)
    
- [Get or set the time when composing an appointment in Outlook](../outlook/get-or-set-the-time-of-an-appointment.md)
    
- [Outlook-Power-Hour_Code-Samples](https://github.com/OfficeDev/Outlook-Power-Hour-Code-Samples):  `ComposeAppDemo`
    

## Additional resources



- [Get Started with Outlook add-ins for Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
    
- [Outlook add-ins](../outlook/outlook-add-ins.md)
    
