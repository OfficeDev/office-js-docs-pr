
# Outlook add-ins

Outlook add-ins are integrations built by third parties into Outlook using the new web technologies based platform. Outlook add-ins have three key aspects:


- The same add-in and business logic works across desktop Outlook for Windows and Mac, web (Office 365 and Outlook.com), and mobile.
    
-  Outlook add-ins consist of a manifest, which describes how the add-in integrates into Outlook (for example, a button or a task pane), and JavaScript/HTML code, which makes up the UI and business logic of the add-in.
    
- Outlook add-ins can be acquired from the Office store or side-loaded by end-users or administrators.
    
Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows. Unlike COM add-ins, Outlook add-ins don't have any code physically installed to the user's device or Outlook client. For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, then loads the JavaScript and HTML. This all executes in the context of a browser in a sandbox.

The Outlook items that support mail add-ins include email messages, meeting requests, responses and cancellations, and appointments. Each mail add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.


## Extension points


Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done:


- Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).
    
    **An add-in with command buttons on the ribbon**

    ![Add-in Command UI-less shape](../../images/41e46a9c-19ec-4ccc-98e6-a227283623d1.png)

- Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](../outlook/contextual-outlook-add-ins.md).
    
    **A contextual add-in for a highlighted entity (an address)**

    ![Shows a contextual app in a card](../../images/59bcabc2-7cb0-4b9b-bb9f-06089dca9c31.png)

- Add-ins can appear in a horizontal pane above the body of the message or appointment. This is based on complex rules, such as presence of attachments or Exchange item class of the message or appointment. For more information, see [Custom pane Outlook add-ins](../outlook/custom-pane-outlook-add-ins.md).
    
    **An add-in with a custom pane in read mode**

    ![Shows a custom pane in a message read form.](../../images/c585ab0a-6c33-42d0-a20f-5deb8b54f480.png)


## Mailbox items available to add-ins


Outlook add-ins are available on messages or appointments while composing or reading, but not other item types. Outlook does not activate add-ins if the current message item, in a compose or read form, is one of the following:


- Protected by Information Rights Management (IRM), in S/MIME format or encrypted in other ways for protection. A digitally signed message is an example since digital signing relies on one of these mechanisms.
    
- In the Junk Email folder.
    
- A delivery report or notification that has the message class IPM.Report.*, including delivery and Non-Delivery Report (NDR) reports, and read, non-read, and delay notifications.
    
- A .msg file which is an attachment to another message.
    
- A .msg file opened from the file system.
    
In general, Outlook can activate add-ins in read forms for items in the Sent Items folder, with the exception of add-ins that activate based on string matches of well-known entities. For more information about the reasons behind this, see [Support for well-known entities](../outlook/match-strings-in-an-item-as-well-known-entities.md#MailAppEntities_Supported).


## Supported hosts


Outlook add-ins are supported in Outlook 2013 and later versions, Outlook 2016 for Mac, Outlook Web App for Exchange 2013 on-premises, Outlook Web App in Office 365 and Outlook.com. Not all of the newest features are supported in all clients at the same time. Please refer to individual topics and API references, to see which hosts they are/are not supported in.


## Get started building Outlook add-ins


To get started building Outlook add-ins, see [Get Started with Outlook add-ins for Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted.aspx).


## Additional Resources


- [Overview of Outlook add-ins architecture and features](../outlook/overview.md)
- [Best practices for developing Office Add-ins](../../docs/design/add-in-development-best-practices.md)
- [Design guidelines for Office Add-ins](../../docs/design/add-in-design.md)
- [License your Office and SharePoint Add-ins](http://msdn.microsoft.com/library/3e0e8ff6-66d6-44ff-b0c2-59108ebd9181%28Office.15%29.aspx)
- [Publish your Office Add-in](../publish/publish.md)
- [Submit Office and SharePoint Add-ins and Office 365 web apps to the Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)

