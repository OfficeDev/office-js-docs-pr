---
title: Debug your add-in with runtime logging
description: Learn how to use runtime logging to debug your add-in.
ms.date: 11/06/2025
ms.localizationpriority: medium
---

# Debug your add-in with runtime logging

You can use runtime logging to debug your add-in's manifest as well as several installation errors. This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs. Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.

> [!IMPORTANT]
> Runtime Logging affects performance. Turn it on only when you need to debug issues with your add-in manifest.

## Use runtime logging from the command line

Enabling runtime logging from the command line is the fastest way to use this logging tool.

> [!IMPORTANT]
> The office-addin-dev-settings tool is not supported on Mac. See the section [Runtime logging on Mac](#runtime-logging-on-mac) for Mac-specific instructions.

- To enable runtime logging:

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable
    ```

- To enable runtime logging only for a specific file, use the same command with a filename:

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable [filename.txt]
    ```

- To disable runtime logging:

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --disable
    ```

- To display whether runtime logging is enabled:

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log
    ```

- To display help within the command line for runtime logging:

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --help
    ```

## Runtime logging on Mac

1. Open **Terminal** and set a runtime logging preference by using the `defaults` command:

    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    `<bundle id>` identifies which the host for which to enable runtime logging. `<file_name>` is the name of the text file to which the log will be written.

    Set `<bundle id>` to one of the following values to enable runtime logging for the corresponding application.

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

The following example enables runtime logging for Word and then opens the log file.

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE]
> You'll need to restart Office after running the `defaults` command to enable runtime logging.

To turn off runtime logging, use the `defaults delete` command:

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

The following example will turn off runtime logging for Word.

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

## Use runtime logging to troubleshoot issues with your manifest

To use runtime logging to troubleshoot issues loading an add-in:

1. [Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.

    > [!NOTE]
    > We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.

1. If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.

1. Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.

## Known issues with runtime logging

You might see messages in the log file that are confusing or that are classified incorrectly. For example:

- The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.

- If you see the message `Unexpected Add-in is missing required manifest fields    DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.

- Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.

## See also

- [Office Add-ins manifest](../develop/add-in-manifests.md)
- [Validate an Office Add-in's manifest](troubleshoot-manifest.md)
- [Clear the Office cache](clear-cache.md)
- [Sideload Office Add-ins for testing](sideload-office-add-ins-for-testing.md)
- [Debug add-ins using developer tools in Microsoft Edge](debug-add-ins-using-devtools-edge-chromium.md)
- [Runtimes in Office Add-ins](runtimes.md)