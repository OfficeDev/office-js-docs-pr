---
title: Debug add-ins using developer tools on Windows 10
description: 'Debug add-ins using Microsoft Edge developer tools on Windows 10'
ms.date: 12/16/2019
localization_priority: Priority
---

# Debug add-ins using developer tools on Windows 10

There are developer tools outside of IDEs available to help you debug your add-ins on Windows 10. These are useful when you need to investigate a problem while running your add-in outside the IDE.

The tool that you use depends on whether the add-in is running in Microsoft Edge or Internet Explorer. This is determined by the version of Windows 10 and the version of Office that are installed on the computer. To determine which browser is being used on your development computer, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

> [!NOTE]
> The instructions in this article cannot be used to debug an Outlook add-in that uses Execute Functions. To debug an Outlook add-in that uses Execute Functions, we recommend that you attach to Visual Studio in script mode or to some other script debugger.

## When the add-in is running in Microsoft Edge

When the add-in is running in Microsoft Edge, you can use the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).

### Enable debugging for add-in commands and UI-less code

[!include[Enable debugging on Microsoft Edge DevTools](../includes/enable-debugging-on-edge-devtools.md)]

### Debug using Microsoft Edge DevTools

1. Run the add-in.

2. Run the Microsoft Edge DevTools.

3. In the tools, open the **Local** tab. Your add-in will be listed by its name.

4. Click the add-in name to open it in the tools.

5. Open the **Debugger** tab. 

6. Choose the folder icon above the **script** (left) pane. From the list of available files shown in the dropdown list, select the JavaScript file that you want to debug.

7. To set a breakpoint, select the line. You will see a red dot to the left of the line and a corresponding line in the **Call stack** (bottom right) pane.

8. Execute functions in the add-in as needed to trigger the breakpoint.

## When the add-in is running in Internet Explorer

When the add-in is running in Internet Explorer, you can use the debugger from the F12 developer tools in Windows 10 to test your add-in. You can start the F12 developer tools after the add-in is running. The F12 tools are displayed in a separate window and do not use Visual Studio.

> [!NOTE]
> The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger. 

This example uses Word and a free add-in from AppSource.

1. Open Word and choose a blank document. 
    
2. On the **Insert** tab, in the Add-ins group, choose **Store** and select the **QR4Office** Add-in. (You can load any add-in from the Store or your add-in catalog.)
    
3. Launch the F12 development tools that corresponds to your version of Office:
    
   - For the 32-bit version of Office, use C:\Windows\System32\F12\IEChooser.exe
    
   - For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\IEChooser.exe
    
   When you launch IEChooser, a separate window named "Choose target to debug" displays the possible applications to debug. Select the application that you are interested in. If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL. 
    
   For example, select **home.html**. 
    
   ![IEChooser screen, pointing to bubbles add-in](../images/choose-target-to-debug.png)

4. In the F12 window, select the file you want to debug.
    
   To select the file in the F12 window, choose the folder icon above the **script** (left) pane. From the list of available files shown in the dropdown list, select **Home.js**.
    
5. Set the breakpoint.
    
   To set the breakpoint in **Home.js**, choose line 144, which is in the  `textChanged` function. You will see a red dot to the left of the line and a corresponding line in the **Call stack and Breakpoints** (bottom right) pane. For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)). 
    
   ![Debugger with breakpoint in home.js file](../images/debugger-home-js-02.png)

6. Run your add-in to trigger the breakpoint.
    
   In Word, choose the URL textbox in the upper part of the **QR4Office** pane and attempt to enter some text. In the Debugger, in the **Call stack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information. You might need to refresh the Debugger to see the results.
    
   ![Debugger with results from the triggered breakpoint](../images/debugger-home-js-01.png)


## See also

- [Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))
- [Using the F12 developer tools](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))
