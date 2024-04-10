---
title: Troubleshoot event-based and spam-reporting add-ins
description: Learn how to troubleshoot development errors in Outlook add-ins that implement event-based activation or integrated spam reporting.
ms.date: 04/12/2024
ms.topic: troubleshooting
ms.localizationpriority: medium
---

# Troubleshoot event-based and spam-reporting add-ins

As you develop your [event-based](autolaunch.md) or [spam-reporting (preview)](spam-reporting.md) add-in, you may encounter issues, such as your add-in not loading or an event not occurring. The following sections provide guidance on how to troubleshoot your add-in.

## Review feature prerequisites

- Verify that the add-in is installed on a supported Outlook client. Some Outlook clients only support certain events or aspects of event-based activation or integrated spam reporting. For more information, see [Supported events](autolaunch.md#supported-events) and [Preview the integrated spam-reporting feature](spam-reporting.md#preview-the-integrated-spam-reporting-feature).
- Verify that your Outlook client supports the minimum requirement set needed.

  Event-based activation was introduced in [requirement set 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10), with additional events now supported in subsequent requirements sets. For more information, see [Supported events](autolaunch.md#supported-events) and [Requirement sets supported by Exchange servers and Outlook clients](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients). If you're developing an add-in that handles the `OnMessageSend` and `OnAppointmentSend` events, see the "Supported clients and platform section" of [Handle OnMessageSend and OnAppointmentSend events in your Outlook add-in with Smart Alerts](onmessagesend-onappointmentsend-events.md#supported-clients-and-platforms).

  The integrated spam reporting feature is currently in preview in Outlook on Windows. For more information, see [Preview the integrated spam-reporting feature](spam-reporting.md#preview-the-integrated-spam-reporting-feature)
- Review the expected behavior and limitations of the feature.

  - [Event-based activation behavior and limitations](autolaunch.md#event-based-activation-behavior-and-limitations)
  - [Smart Alerts behavior and scenarios](onmessagesend-onappointmentsend-events.md#smart-alerts-feature-behavior-and-scenarios)
  - [Integrated spam-reporting behavior and limitations](spam-reporting.md#review-feature-behavior-and-limitations)

## Check manifest and JavaScript requirements

- Ensure that the following conditions are met in your add-in's manifest.
  - Verify that your add-in's source file location URL is publicly available and isn't blocked by a firewall. This URL is specified in the [SourceLocation element](/javascript/api/manifest/sourcelocation) of the XML manifest or the "extensions.runtimes.code.page" property of the unified manifest for Microsoft 365.
  - Verify that the **\<Runtimes\>** element (XML manifest) or "extensions.runtimes.code" property (unified manifest) correctly references the HTML or JavaScript file containing the event handlers. Outlook on Windows uses the JavaScript file during runtime, while Outlook on the web and on new Mac UI use the HTML file. For an example of how this is configured in the manifest, see the "Configure the manifest" section of [Automatically set the subject of a new message or appointment](on-new-compose-events-walkthrough.md#configure-the-manifest).
  
    For Outlook on Windows, you must bundle all your event-handling JavaScript code into this JavaScript file referenced in the manifest. Note that a large JavaScript bundle may cause issues with the performance of your add-in. We recommend preprocessing heavy operations, so that they're not included in your event-handling code.
- Verify that your event-handling JavaScript file referenced by the Outlook client on Windows calls `Office.actions.associate`. This ensures that the event handler name specified in the manifest is mapped to its JavaScript counterpart.

  > [!TIP]
  > If your add-in has only one JavaScript file referenced by Outlook on the web, Windows, and Mac, it's recommended to check on which platform the add-in is running to determine when to call `Office.actions.associate`, as shown in the following code.
  >
  > ```js
  > if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  >   Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  >   Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
  > }
  > ```

- The JavaScript code of event-based and spam-reporting add-ins that run in Outlook on Windows only supports [ECMAScript 2016](https://262.ecma-international.org/7.0/) and earlier specifications. Some examples of programming syntax to avoid are as follows.
  - Avoid using `async` and `await` statements in your code. Including these in your JavaScript code will cause the add-in to time out.
  - Avoid using the [conditional (ternary) operator](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/Conditional_Operator) as it will prevent your add-in from loading.
  
  If your add-in has only one JavaScript file referenced by Outlook on the web, Windows, and Mac, you must limit your code to ECMAScript 2016 to ensure that your add-in runs in Outlook on Windows. However, if you have a separate JavaScript file referenced by Outlook on the web and Mac, you can implement a later ECMAScript specification in that file.

## Debug your add-in

- As you make changes to your add-in, be aware that:
  - If you update the manifest, [remove the add-in](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in), then sideload it again. If you're using Outlook on Windows, you must also close and reopen Outlook.
  - If you make changes to files other than the manifest, close and reopen the Outlook client on Windows or on Mac, or refresh the browser tab running Outlook on the web.
  - If you're still unable to see your changes after performing these steps, [clear your Office cache](../testing/clear-cache.md).
- As you test your add-in in Outlook on Windows:
  - For event-based add-ins, check [Event Viewer](/shows/inside/event-viewer) for any reported add-in errors.
    1. In Event Viewer, select **Windows Logs** > **Application**.
    1. From the **Actions** panel, select **Filter Current Log**.
    1. From the **Logged** dropdown, select your preferred log time frame.
    1. Select the **Error** checkbox.
    1. In the **Event IDs** field, enter **63**.
    1. Select **OK** to apply your filters.

    :::image type="content" source="../images/outlook-event-based-logs.png" alt-text="A sample of Event Viewer's Filter Current Log settings configured to only show Outlook errors with event ID 63 that occurred in the last hour.":::

  - Verify that the **bundle.js** file is downloaded to the following folder in File Explorer. Replace text enclosed in `[]` with your applicable information.
  
    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[Outlook mail account encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]
    ```

    [!INCLUDE [outlook-bundle-js](../includes/outlook-bundle-js.md)]

- As you test your add-in in Outlook on Windows or Mac, enable runtime logging to identify possible manifest and add-in installation issues. For guidance on how to use runtime logging, see [Debug your add-in with runtime logging](../testing/runtime-logging.md).
- Set breakpoints in your code to debug your add-in. For platform-specific instructions, see [Debug event-based and spam-reporting add-ins](debug-autolaunch.md).

## Seek additional help

If you still need help after performing the recommended troubleshooting steps, [open a GitHub issue](https://github.com/OfficeDev/office-js/issues/new?assignees=&labels=&template=bug_report.md&title=). Include screenshots, video recordings, or runtime logs to supplement your report.

## See also

- [Configure your Outlook add-in for event-based activation](autolaunch.md)
- [Implement an integrated spam-reporting add-in (preview)](spam-reporting.md)
