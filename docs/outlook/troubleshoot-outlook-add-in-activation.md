---
title: Troubleshoot Outlook contextual add-in activation
description: Possible reasons your contextual Outlook add-in doesn't activate as you expect.
ms.date: 07/14/2025
ms.topic: troubleshooting
ms.localizationpriority: medium
---

# Troubleshoot Outlook contextual add-in activation

Outlook contextual add-in activation is based on the activation rules in an add-in only manifest for the add-in. When conditions for the currently selected item satisfy the activation rules for the add-in, the application activates and displays the add-in button in the Outlook UI (add-in selection pane for compose add-ins, add-in bar for read add-ins). However, if your add-in doesn't activate as you expect, you should look into the following areas for possible reasons.

[!INCLUDE [outlook-contextual-add-ins-retirement](../includes/outlook-contextual-add-ins-retirement.md)]

## Is user mailbox on a version of Exchange Server that is at least Exchange 2016?

First, ensure that the user's email account you're testing with is on a version of Exchange Server that is at least Exchange 2016. If you're using specific features that are released after Exchange 2016, make sure the user's account is on the appropriate version of Exchange.

You can verify the version of Exchange by using one of the following approaches.

- Check with your Exchange Server administrator.

- If you're testing the add-in on Outlook on the web or mobile devices, in a script debugger (for example, the JScript Debugger that comes with Internet Explorer), look for the **src** attribute of the **script** tag that specifies the location from which scripts are loaded. The path should contain a substring **owa/15.0.516.x/owa2/...**, where **15.0.516.x** represents the version of the Exchange Server, such as **15.0.516.2**.

- Alternatively, you can use the [Office.context.mailbox.diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostversion-member) property to verify the version. In Outlook on the web, on mobile devices, and in [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627), this property returns the version of the Exchange Server.

## Is the add-in available?

Check your list of installed add-ins to verify whether an add-in is available. For instructions on how to view your add-ins in Outlook, see [Use add-ins in Outlook](https://support.microsoft.com/office/1ee261f9-49bf-4ba6-b3e2-2ba7bcab64c8).

If you also use the classic Outlook on Mac client, resource usage limits apply to add-ins that could affect their availability across supported platforms. For more information, see [Resource usage limits for add-ins](../concepts/resource-limits-and-performance-optimization.md#resource-usage-limits-for-add-ins).

## Does the tested item support Outlook add-ins? Is the selected item delivered by a version of Exchange Server that is at least Exchange 2016?

If your Outlook add-in is a read add-in and is supposed to be activated when the user is viewing a message (including email messages, meeting requests, responses, and cancellations) or appointment, even though these items generally support add-ins, there are exceptions. Check if the selected item is one of those [listed where Outlook add-ins don't activate](outlook-add-ins-overview.md#add-in-activation-limitations).

Also, because appointments are always saved in Rich Text Format, an [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule) rule that specifies a **PropertyName** value of **BodyAsHTML** wouldn't activate an add-in on an appointment or message that's saved in plain text or Rich Text Format.

## Is the add-in manifest installed properly, and does Outlook have a cached copy?

[!INCLUDE [Rule features not supported by the unified manifest for Microsoft 365](../includes/rules-not-supported-json-note.md)]

This scenario applies to only classic Outlook on Windows. Normally, when you install an Outlook add-in for a mailbox, the Exchange Server copies the add-in manifest from the location you indicate to the mailbox on that Exchange Server. Every time Outlook starts, it reads all the manifests installed for that mailbox into a temporary cache at the following location.

```text
%LocalAppData%\Microsoft\Office\16.0\WEF
```

For example, for the user John, the cache might be at C:\Users\john\AppData\Local\Microsoft\Office\16.0\WEF.

If an add-in doesn't activate for any items, the manifest might not have been installed properly on the Exchange Server, or Outlook hasn't read the manifest properly on startup. Using the Exchange Admin Center, ensure that the add-in is installed and enabled for your mailbox, and reboot the Exchange Server, if necessary.

The following figure shows a summary of the steps to verify whether Outlook has a valid version of the manifest.

![Flow chart to check manifest.](../images/troubleshoot-manifest-flow.png)

The following procedure describes the details.

1. If you have modified the manifest while Outlook is open, and you're not using Visual Studio 2015 or a later version of Visual Studio to develop the add-in, you should uninstall the add-in and reinstall it using the Exchange Admin Center.

1. Restart Outlook and test whether Outlook now activates the add-in.

1. If Outlook doesn't activate the add-in, check whether Outlook has a properly cached copy of the manifest for the add-in. Look under the following path.

    ```text
    %LocalAppData%\Microsoft\Office\16.0\WEF
    ```

    You can find the manifest in the following subfolder.

    ```text
    \<insert your guid>\<insert base 64 hash>\Manifests\<ManifestID>_<ManifestVersion>
    ```

    > [!NOTE]
    > The following is an example of a path to a manifest installed for a mailbox for the user John.
    >
    > ```text
    > C:\Users\john\appdata\Local\Microsoft\Office\16.0\WEF\{8D8445A4-80E4-4D6B-B7AC-D4E6AF594E73}\GoRshCWa7vW8+jhKmyiDhA==\Manifests\b3d7d9d5-6f57-437d-9830-94e2aaccef16_1.2
    > ```

    Verify whether the manifest of the add-in you're testing is among the cached manifests.

1. If the manifest is in the cache, skip the rest of this section and consider the other possible reasons following this section.

1. If the manifest is not in the cache, check whether Outlook indeed successfully read the manifest from the Exchange Server. To do that, use the Windows Event Viewer:

    1. Under **Windows Logs**, choose **Application**.

    1. Look for a reasonably recent event for which the Event ID equals 63, which represents Outlook downloading a manifest from an Exchange Server.

    1. If Outlook successfully read a manifest, the logged event should have the following description.

        ```text
        The Exchange web service request GetAppManifests succeeded.
        ```

        Then skip the rest of this section and consider the other possible reasons following this section.

1. If you don't see a successful event, close Outlook, and delete all the manifests in the following path.

    ```text
    %LocalAppData%\Microsoft\Office\16.0\WEF\<insert your guid>\<insert base 64 hash>\Manifests\
    ```

    Start Outlook and test whether Outlook now activates the add-in.

1. If Outlook doesn't activate the add-in, go back to Step 3 to verify again whether Outlook has properly read the manifest.

## Is the add-in manifest valid?

See [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md) to debug add-in manifest issues.

## Are you using the appropriate activation rules?

Starting in version 1.1 of the Office Add-ins manifests schema, you can create add-ins that are activated when the user is in a compose form (compose add-ins) or in a read form (read add-ins). Make sure you specify the appropriate activation rules for each type of form that your add-in is supposed to activate in. For example, you can activate compose add-ins using only [ItemIs](/javascript/api/manifest/rule#itemis-rule) rules with the **FormType** attribute set to **Edit** or **ReadOrEdit**, and you can't use any of the other types of rules, such as [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule) rules for compose add-ins.

## If you use a regular expression, is it properly specified?

Because regular expressions in activation rules are part of the XML-formatted add-in only manifest file for a read add-in, if a regular expression uses certain characters, be sure to follow the corresponding escape sequence that XML processors support. The following table lists these special characters.

|Character|Description|Escape sequence to use|
|:-----|:-----|:-----|
|`"`|Double quotation mark|&amp;quot;|
|`&`|Ampersand|&amp;amp;|
|`'`|Apostrophe|&amp;apos;|
|`<`|Less-than sign|&amp;lt;|
|`>`|Greater-than sign|&amp;gt;|

## If you use a regular expression, does the read add-in activate in Outlook on the web, on mobile devices, or in new Outlook on Windows, but not in Outlook on Windows (classic) or Outlook on Mac?

Outlook on Windows (classic) and Outlook on Mac use a regular expression engine that's different from the one used by Outlook on the web, on mobile devices, and on [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627). Classic Outlook on Windows and Outlook on Mac use the C++ regular expression engine provided as part of the Visual Studio standard template library. This engine complies with ECMAScript 5 standards. Outlook on the web, on mobile devices, and [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) use regular expression evaluation that's part of JavaScript, is provided by the browser, and supports a superset of ECMAScript 5.

While in most cases, these Outlook clients find the same matches for the same regular expression in an activation rule, there are exceptions. For instance, if the regex includes a custom character class based on predefined character classes, Outlook on Windows (classic) and Outlook on Mac may return results different from Outlook on the web, on mobile devices, and new Outlook on Windows. As an example, character classes that contain shorthand character classes `[\d\w]` within them would return different results. In this case, to avoid different results on different applications, use `(\d|\w)` instead.

Test your regular expression thoroughly. If it returns different results, rewrite the regular expression for compatibility with both engines. To verify evaluation results in Outlook on Windows (classic) and Outlook on Mac, write a small C++ program that applies the regular expression against a sample of the text you are trying to match. Running on Visual Studio, the C++ test program would use the standard template library, simulating the behavior of Outlook on Windows (classic) or Outlook on Mac when running the same regular expression. To verify evaluation results in Outlook on the web, on mobile devices, and in new Outlook on Windows, use your favorite JavaScript regular expression tester.

## If you use an ItemIs, ItemHasAttachment, or ItemHasRegularExpressionMatch rule, have you verified the related item property?

If you use an **ItemHasRegularExpressionMatch** activation rule, verify whether the value of the **PropertyName** attribute is what you expect for the selected item. The following are some tips to debug the corresponding properties.

- If the selected item is a message and you specify **BodyAsHTML** in the **PropertyName** attribute, open the message, and then choose **View Source** to verify the message body in the HTML representation of that item.

- If the selected item is an appointment, or if the activation rule specifies **BodyAsPlaintext** in the **PropertyName**, you can use the Outlook object model and the Visual Basic Editor in classic Outlook on Windows.

    1. Ensure that macros are enabled and the **Developer** tab is displayed on the ribbon for Outlook.

    1. In the Visual Basic Editor, choose **View**, **Immediate Window**.

    1. Type the following to display various properties depending on the scenario.

        - The HTML body of the message or appointment item selected in the Outlook explorer:

        ```vb
        ?ActiveExplorer.Selection.Item(1).HTMLBody
        ```

        - The plain text body of the message or appointment item selected in the Outlook explorer:

        ```vb
        ?ActiveExplorer.Selection.Item(1).Body
        ```

        - The HTML body of the message or appointment item opened in the current Outlook inspector:

        ```vb
        ?ActiveInspector.CurrentItem.HTMLBody
        ```

        - The plain text body of the message or appointment item opened in the current Outlook inspector:

        ```vb
        ?ActiveInspector.CurrentItem.Body
        ```

If the **ItemHasRegularExpressionMatch** activation rule specifies **Subject** or **SenderSMTPAddress**, or if you use an **ItemIs** or **ItemHasAttachment** rule, and you're familiar with or would like to use MAPI, use [MFCMAPI](https://github.com/stephenegriffin/mfcmapi). Verify the MAPI property that your rule relies on in the following table.

|Type of rule|Verify this MAPI property|
|:-----|:-----|
|**ItemHasRegularExpressionMatch** rule with **Subject**|[PidTagSubject](/office/client-developer/outlook/mapi/pidtagsubject-canonical-property)|
|**ItemHasRegularExpressionMatch** rule with **SenderSMTPAddress**|[PidTagSenderSmtpAddress](/office/client-developer/outlook/mapi/pidtagsendersmtpaddress-canonical-property) and [PidTagSentRepresentingSmtpAddress](/office/client-developer/outlook/mapi/pidtagsentrepresentingsmtpaddress-canonical-property)|
|**ItemIs**|[PidTagMessageClass](/office/client-developer/outlook/mapi/pidtagmessageclass-canonical-property)|
|**ItemHasAttachment**|[PidTagHasAttachments](/office/client-developer/outlook/mapi/pidtaghasattachments-canonical-property)|

After verifying the property value, you can then use a regular expression evaluation tool to test whether the regular expression finds a match in that value.

## Does Outlook apply all the regular expressions to the portion of the item body as you expect?

This section applies to all activation rules that use regular expressions, particularly, those that are applied to the item body, which may be large in size and take longer to evaluate for matches. You should be aware that even if the item property that an activation rule depends on has the value you expect, Outlook may not be able to evaluate all the regular expressions on the entire value of the item property. To provide reasonable performance and to control excessive resource usage by a read add-in, Outlook observes the following limits on processing regular expressions in activation rules at runtime.

- **The size of the item body evaluated**. There are limits to the portion of an item body on which Outlook evaluates a regular expression. These limits depend on the Outlook client, form factor, and format of the item body. For more information, see [Limits on the size of the item body evaluated](limits-for-activation-and-javascript-api-for-outlook-add-ins.md#limits-on-the-size-of-the-item-body-evaluated).

- **Number of regular expression matches**. Outlook on the web, on Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic), on Mac, and on mobile devices each returns a maximum of 50 regular expression matches. These matches are unique, and duplicate matches don't count against this limit. Don't assume any order to the returned matches, and don't assume the order in Outlook on Windows (classic) and Outlook on Mac is the same as that in Outlook on the web, on mobile devices, and in new Outlook on Windows. If you expect many matches to regular expressions in your activation rules, and you're missing a match, you may be exceeding this limit.

- **Length of a regular expression match**. There are limits to the length of a regular expression match that the Outlook application would return. Outlook doesn't include any match above the limit and doesn't display any warning message. You can run your regular expression using other regex evaluation tools or a stand-alone C++ test program to verify whether you have a match that exceeds such limits. The following table summarizes the limits. For more information, see [Limits of activation rules for contextual Outlook add-in](limits-for-activation-and-javascript-api-for-outlook-add-ins.md#limits-on-the-matches-returned).

    |Limit on length of a regex match|Outlook on the web, on new Windows client, and on mobile devices|Outlook on Windows (classic) and on Mac|
    |:-----|:-----|:-----|
    |Item body is plain text|3 KB|1.5 KB|
    |Item body is HTML|3 KB|3 KB|

- **Time spent on evaluating all regular expressions of a read add-in in Outlook on Windows (classic) and Outlook on Mac**. By default, for each read add-in, Outlook must finish evaluating all the regular expressions in its activation rules within one second. Otherwise, Outlook retries up to three times and makes the add-in unavailable if Outlook can't complete the evaluation. Outlook displays a message in the notification bar that the add-in is unavailable. The amount of time available for your regular expression can be modified by setting a group policy or a registry key.

   > [!NOTE]
   > If Outlook on Windows (classic) or Outlook on Mac makes a read add-in unavailable, the read add-in becomes unavailable on the same mailbox in Outlook on the web, on mobile devices, and in new Outlook on Windows.

## See also

- [Deploy and install Outlook add-ins for testing](testing-and-tips.md)
- [Contextual Outlook add-ins](contextual-outlook-add-ins.md)
- [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md)
