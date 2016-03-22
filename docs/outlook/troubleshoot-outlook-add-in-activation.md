
# Troubleshoot Outlook add-in activation


Outlook add-in activation is contextual and is based on the activation rules in the add-in manifest. When conditions for the currently selected item satisfy the activation rules for the add-in, the host application activates and displays the add-in button in the Outlook UI (add-in selection pane for compose add-ins, add-in bar for read add-ins). However, if your add-in doesn't activate as you expect, you should look into the following areas for possible reasons.

## Is the user mailbox on a version of Exchange Server that is at least Exchange 2013?


First, ensure that the user's email account you're testing with is on a version of Exchange Server that is at least Exchange 2013. If you are using specific features that are released after Exchange 2013, make sure the user's account is on the appropriate version of Exchange.

You can verify the version of Exchange 2013 by using one of the following approaches:


- Check with your Exchange Server administrator.
    
- If you are testing the add-in on Outlook Web App or OWA for Devices, in a script debugger (for example, the JScript Debugger that comes with Internet Explorer), look for the  **src** attribute of the **script** tag that specifies the location from which scripts are loaded. The path should contain a substring **owa/15.0.516.x/owa2/...**, where  **15.0.516.x** represents the version of the Exchange Server, such as **15.0.516.2**.
    
- Alternatively, you can use the [Office.context.mailbox.diagnostics.hostVersion](../../reference/outlook/Office.context.mailbox.diagnostics.md) property to verify the version. On Outlook Web App and OWA for Devices, this property returns the version of the Exchange Server.
    
- If you can test the add-in on Outlook, you can use the following simple debugging technique that uses the Outlook object model and Visual Basic Editor:
    
      1. First, verify that macros are enabled for Outlook. Choose  **File**,  **Options**,  **Trust Center**,  **Trust Center Settings**,  **Macro Settings**. Ensure that  **Notifications for all macros** is selected in the Trust Center. You should have also selected **Enable Macros** during Outlook startup.
    
      2. On the  **Developer** tab of the ribbon, choose **Visual Basic**.
    
     >**Note**  Not seeing the  **Developer** tab? See [How to: Show the Developer Tab on the Ribbon](http://msdn.microsoft.com/en-us/library/ce7cb641-44f2-4a40-867e-a7d88f8e98a9%28Office.15%29.aspx) to turn it on.
      3. In the Visual Basic Editor, choose  **View**,  **Immediate Window**.
    
      4. Type the following in the Immediate window to display the version of the Exchange Server. The major version of the returned value must be equal to or greater than 15.
    
        - If there is only one Exchange account in the user's profile:
        
            ```vba
                  ?Session.ExchangeMailboxServerVersion
            ```
        
        - If there are multiple Exchange accounts in the same user profile:
        
                    ```vba
                      ?Session.Accounts.Item(emailAddress).ExchangeMailboxServerVersion
                    ```
        
        - _emailAddress_ represents a string that contains the user's primary STMP address. For example, if the user's primary SMTP address is randy@contoso.com, type the following:
        
                ```vba
                  ?Session.Accounts.Item("randy@contoso.com").ExchangeMailboxServerVersion
                ```


## Is the add-in disabled?


Any one of the Outlook rich clients can disable an add-in for performance reasons, including exceeding usage thresholds for CPU core or memory, tolerance for crashes, and length of time to process all the regular expressions for an add-in. When this happens, the Outlook rich client displays a notification that it is disabling the add-in. 


 >**Note**  Only Outlook rich clients monitor resource usage, but disabling an add-in in an Outlook rich client also disables the add-in in Outlook Web App and OWA for Devices.

Use one of the following approaches to verify whether an add-in is disabled: 


- In Outlook Web App, sign in directly to the email account, choose the Settings icon, and then choose  **Manage add-ins** to go to the Exchange Admin Center, where you can verify whether the add-in is enabled.
    
- In Outlook, go to the Backstage view and choose  **Manage add-ins**. Sign in to the Exchange Admin Center to verify whether the add-in is enabled.
    
- In Outlook for Mac, choose  **Manage add-ins** in the add-in bar. Sign in to the Exchange Admin Center to verify whether the add-in is enabled.
    

## Does the tested item support Outlook add-ins? Is the selected item delivered by a version of Exchange Server that is at least Exchange 2013?


If your Outlook add-in is a read add-in and is supposed to be activated when the user is viewing a message (including email messages, meeting requests, responses, and cancellations) or appointment, even though these items generally support add-ins, there are exceptions if the selected item is one of the following:


- Protected by Information Rights Management (IRM).
    
- In S/MIME format or encrypted in other ways for protection.
    
- A draft (does not have a sender assigned to it), or is in the Outlook Drafts folder.
    
- In the Junk Email folder.
    
- A delivery report or notification that has the message class IPM.Report.*, including delivery and Non-Delivery Report (NDR) reports, and read, non-read, and delay notifications.
    
- A .msg file that is attached to another message or was opened from the file system.
    
Also, because appointments are always saved in Rich Text Format, an [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) rule that specifies a **PropertyName** value of **BodyAsHTML** would not activate an add-in on an appointment or message that is saved in plain text or Rich Text Format.

Even if a mail item is not one of the above types, if the item was not delivered by a version of Exchange Server that is at least Exchange 2013, known entities and properties such as sender's SMTP address would not be identified on the item. Any activation rules that rely on these entities or properties would not be satisfied, and the add-in would not be activated.

If your add-in is a compose add-in and is supposed to be activated when the user is authoring a message or meeting request, make sure the item is not protected by IRM.


## Is the add-in manifest installed properly, and does Outlook have a cached copy?


This scenario applies to only Outlook for Windows. Normally, when you install an Outlook add-in for a mailbox, the Exchange Server copies the add-in manifest from the location you indicate to the mailbox on that Exchange Server. Every time the Outlook starts, it reads all the manifests installed for that mailbox into a temporary cache at the following location: 

%LocalAppData%\Microsoft\Office\15.0\WEF 

For example, for the user John, the cache might be at C:\Users\john\AppData\Local\Microsoft\Office\15.0\WEF.

If an add-in does not activate for any items, the manifest might not have been installed properly on the Exchange Server, or Outlook has not read the manifest properly on startup. Using the Exchange Admin Center, ensure that the add-in is installed and enabled for your mailbox, and reboot the Exchange Server, if necessary.

Figure 1 shows a summary of the steps to verify whether Outlook has a valid version of the manifest. 


**Figure 1. Flow chart of the steps to verify whether Outlook properly cached the manifest**

![Flow chart to check manifest](../../images/off15appsdk_TroubleshootManifest.png)The following procedure describes the details.



1. If you have modified the manifest while Outlook is open, and you are not using Napa, Visual Studio 2012, or a later version of Visual Studio to develop the add-in, you should uninstall the add-in and reinstall it using the Exchange Admin Center. 
    
2. Restart Outlook and test whether Outlook now activates the add-in.
    
3. If Outlook doesn't activate the add-in, check whether Outlook has a properly cached copy of the manifest for the add-in. Look under the following path:
    
    %LocalAppData%\Microsoft\Office\15.0\WEF
    
    You can find the manifest in the following subfolder:
    
    [GUID]\ [BASE 64 Hash]]\Manifests\[ManifestID]_[ManifestVersion]
    
     >**Note**  The following is an example of a path to a manifest installed for a mailbox for the user John:
    
    C:\Users\john\appdata\Local\Microsoft\Office\15.0\WEF\{8D8445A4-80E4-4D6B-B7AC-D4E6AF594E73}\GoRshCWa7vW8+jhKmyiDhA==\Manifests\b3d7d9d5-6f57-437d-9830-94e2aaccef16_1.2
    
    Verify whether the manifest of the add-in you're testing is among the cached manifests.
    
4. If the manifest is in the cache, skip the rest of this section and consider the other possible reasons following this section.
    
5. If the manifest is not in the cache, check whether Outlook indeed successfully read the manifest from the Exchange Server. To do that, use the Windows Event Viewer:
    
      1. Under  **Windows Logs**, choose  **Application**.
    
      2. Look for a reasonably recent event for which the Event ID equals 63, which represents Outlook downloading a manifest from an Exchange Server.
    
      3. If Outlook successfully read a manifest, the logged event should have the following description:
    
         **The Exchange web service request GetAppManifests succeeded.**
    
        Then skip the rest of this section and consider the other possible reasons following this section.
    

    For information about opening the Event Viewer in Windows 7, see [Open Event Viewer](http://windows.microsoft.com/en-US/windows7/Open-Event-Viewer).
    
6. If you don't see a successful event, close Outlook, and delete all the manifests in the following path:
    
    %LocalAppData%\Microsoft\Office\15.0\WEF\[GUID]\[BASE 64 Hash]]\Manifests\
    
    Start Outlook and test whether Outlook now activates the add-in.
    
7. If Outlook doesn't activate the add-in, go back to Step 3 to verify again whether Outlook has properly read the manifest.
    

## Are you using the appropriate activation rules?


Starting in version 1.1 of the Office Add-ins manifests schema, you can create add-ins that are activated when the user is in a compose form (compose add-ins) or in a read form (read add-ins). Make sure you specify the appropriate activation rules for each type of form that your add-in is supposed to activate in. For example, you can activate compose add-ins using only [ItemIs](http://msdn.microsoft.com/en-us/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx) rules with the **FormType** attribute set to **Edit** or **ReadOrEdit**, and you cannot use any of the other types of rules, such as [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) and [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) rules for compose add-ins. For more information, see [Activation rules for Outlook add-ins](../outlook/manifests/activation-rules.md).


## If you use a regular expression, is it properly specified?


Because regular expressions in activation rules are part of the XML manifest file for a read add-in, if a regular expression uses certain characters, be sure to follow the corresponding escape sequence that XML processors support. Table 1 lists these special characters. 


**Table 1. Escape sequences for regular expressions**


|**Character**|**Description**|**Escape sequence to use**|
|:-----|:-----|:-----|
|"|Double quotation mark|&amp;quot;|
|&amp;|Ampersand|&amp;amp;|
|'|Apostrophe|&amp;apos;|
|<|Less-than sign|&amp;lt;|
|>|Greater-than sign|&amp;gt;|

## If you use a regular expression, is the read add-in activating in Outlook Web App or OWA for Devices, but not in any of the Outlook rich clients?


Outlook rich clients use a regular expression engine that's different from the one used by Outlook Web App and OWA for Devices. Outlook rich clients use the C++ regular expression engine provided as part of the Visual Studio standard template library. This engine complies with ECMAScript 5 standards. Outlook Web App and OWA for Devices use regular expression evaluation that is part of JavaScript, is provided by the browser, and supports a superset of ECMAScript 5. 

While in most cases, these host applications find the same matches for the same regular expression in an activation rule, there are exceptions. For instance, if the regex includes a custom character class based on predefined character classes, an Outlook rich client may return results different from Outlook Web App and OWA for Devices. As an example, character classes that contain shorthand character classes  `[\d\w]` within them would return different results. In this case, to avoid different results on different hosts, use `(\d|\w)` instead.

Test your regular expression thoroughly. If it returns different results, rewrite the regular expression for compatibility with both engines. To verify evaluation results on an Outlook rich client, write a small C++ program that applies the regular expression against a sample of the text you are trying to match. Running on Visual Studio, the C++ test program would use the standard template library, simulating the behavior of the Outlook rich client when running the same regular expression. To verify evaluation results on Outlook Web App or OWA for Devices, use your favorite JavaScript regular expression tester.


## If you use an ItemIs, ItemHasAttachment, or ItemHasRegularExpressionMatch rule, have you verified the related item property?


If you use an  **ItemHasRegularExpressionMatch** activation rule, verify whether the value of the **PropertyName** attribute is what you expect for the selected item. The following are some tips to debug the corresponding properties:


- If the selected item is a message and you specify  **BodyAsHTML** in the **PropertyName** attribute, open the message, and then choose **View Source** to verify the message body in the HTML representation of that item.
    
- If the selected item is an appointment, or if the activation rule specifies  **BodyAsPlaintext** in the **PropertyName**, you can use the Outlook object model and the Visual Basic Editor in Outlook for Windows:
    
      1. Ensure that macros are enabled and the  **Developer** tab is displayed in the ribbon for Outlook. If you're not sure how to do this, see Steps 1 and 2 under [Is the user mailbox on a version of Exchange Server that is at least Exchange 2013?](#TroubleshootingMailApps_MailboxE15)
    
      2. In the Visual Basic Editor, choose  **View**,  **Immediate Window**.
    
      3. Type the following to display various properties depending on the scenario. 
    
      - The HTML body of the message or appointment item selected in the Outlook explorer:
    
            ```vba
              ?ActiveExplorer.Selection.Item(1).HTMLBody
            ```


     - The plain text body of the message or appointment item selected in the Outlook explorer:
    
                 ```vba
              ?ActiveExplorer.Selection.Item(1).Body
            ```


      - The HTML body of the message or appointment item opened in the current Outlook inspector:
    
            ```vba
              ?ActiveInspector.CurrentItem.HTMLBody
            ```
      - The plain text body of the message or appointment item opened in the current Outlook inspector:
    
            ```vba
              ?ActiveInspector.CurrentItem.Body
            ```

If the  **ItemHasRegularExpressionMatch** activation rule specifies **Subject** or **SenderSMTPAddress**, or if you use an  **ItemIs** or **ItemHasAttachment** rule, and you are familiar with or would like to use MAPI, you can use [MFCMAPI](http://mfcmapi.codeplex.com/) to verify the value in Table 2 that your rule relies on.


**Table 2. Activation rules and corresponding MAPI properties**


|**Type of rule**|**Verify this MAPI property**|
|:-----|:-----|
|**ItemHasRegularExpressionMatch** rule with **Subject**|[PidTagSubject](http://msdn.microsoft.com/en-us/library/aa7ba4d9-c5e0-4ce7-a34e-65f675223bc9%28Office.15%29.aspx)|
|**ItemHasRegularExpressionMatch** rule with **SenderSMTPAddress**|[PidTagSenderSmtpAddress](http://msdn.microsoft.com/en-us/library/321cde5a-05db-498b-a9b8-cb54c8a14e34%28Office.15%29.aspx) and [PidTagSentRepresentingSmtpAddress](http://msdn.microsoft.com/en-us/library/5ed122a2-0967-4de3-a2ee-69f81ae77b16%28Office.15%29.aspx)|
|**ItemIs**|[PidTagMessageClass](http://msdn.microsoft.com/en-us/library/1e704023-1992-4b43-857e-0a7da7bc8e87%28Office.15%29.aspx)|
|**ItemHasAttachment**|[PidTagHasAttachments](http://msdn.microsoft.com/en-us/library/fd236d74-2868-46a8-bb3d-17f8365931b6%28Office.15%29.aspx)|
After verifying the property value, you can then use a regular expression evaluation tool to test whether the regular expression finds a match in that value.


## Does the host application apply all the regular expressions to the portion of the item body as you expect?


This section applies to all activation rules that use regular expressions -- particularly those that are applied to the item body, which may be large in size and take longer to evaluate for matches. You should be aware that even if the item property that an activation rule depends on has the value you expect, the host application may not be able to evaluate all the regular expressions on the entire value of the item property. To provide reasonable performance and to control excessive resource usage by a read add-in, Outlook, Outlook Web App and OWA for Devices observe the following limits on processing regular expressions in activation rules at run time:


- The size of the item body evaluated -- There are limits to the portion of an item body on which the host application evaluates a regular expression. These limits depend on the host application, form factor, and format of the item body. See the details in Table 2 in [Limits for activation and JavaScript API for Outlook add-ins](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md).
    
- Number of regular expression matches -- The Outlook rich clients, Outlook Web App and OWA for Devices each returns a maximum of 50 regular expression matches. These matches are unique, and duplicate matches do not count against this limit. Do not assume any order to the returned matches, and do not assume the order in an Outlook rich client is the same as that in Outlook Web App and OWA for Devices. If you expect many matches to regular expressions in your activation rules, and you're missing a match, you may be exceeding this limit.
    
- Length of a regular expression match -- There are limits to the length of a regular expression match that the host application would return. The host application does not include any match above the limit and does not display any warning message. You can run your regular expression using other regex evaluation tools or a stand-alone C++ test program to verify whether you have a match that exceeds such limits. Table 3 summarizes the limits. For more information, see Table 3 in [Limits for activation and JavaScript API for Outlook add-ins](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md).
    
    **Table 3. Length limits for a regular expression match**


|**Limit on length of a regex match**|**Outlook rich clients**|**Outlook Web App or OWA for Devices**|
|:-----|:-----|:-----|
|Item body is plain text|1.5 KB|3 KB|
|Item body is HTML|3 KB|3 KB|
- Time spent on evaluating all regular expressions of a read add-in for an Outlook rich client: By default, for each read add-in, Outlook must finish evaluating all the regular expressions in its activation rules within 1 second. Otherwise Outlook retries up to three times and disables the add-in if Outlook cannot complete the evaluation. Outlook displays a message in the notification bar that the add-in has been disabled. The amount of time available for your regular expression can be modified by setting a group policy or a registry key. 
    
     >**Note**  Note that if the Outlook rich client disables a read add-in, the read add-in is not available for use for the same mailbox on the Outlook rich client, Outlook Web App and OWA for Devices.

## Additional resources



- [Deploy and install Outlook add-ins for testing](../outlook/testing-and-tips.md)
    
- [Activation rules for Outlook add-ins](../outlook/manifests/activation-rules.md)
    
- [Use regular expression activation rules to show an Outlook add-in](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [Limits for activation and JavaScript API for Outlook add-ins](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
- [Open Event Viewer](http://windows.microsoft.com/en-US/windows7/Open-Event-Viewer)
    
- [ItemHasAttachment complexType](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx)
    
- [ItemHasRegularExpressionMatch complexType](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)
    
- [ItemIs complexType](http://msdn.microsoft.com/en-us/library/926249ab-2d2f-39f5-1d73-fab1c989966f%28Office.15%29.aspx)
    
- [MailApp complexType](http://msdn.microsoft.com/en-us/library/696b9fcf-cd10-3f20-4d49-86d3690c887a%28Office.15%29.aspx)
    
