---
title: Convert an add-in to use the unified manifest for Microsoft 365
description: Learn the various methods for converting an add-in with an add-in only manifest to the unified manifest for Microsoft 365 and sideload the add-in.
ms.topic: how-to
ms.date: 04/18/2025
ms.localizationpriority: medium
---

# Convert an add-in to use the unified manifest for Microsoft 365

To add Teams capabilities to an add-in that uses the add-in only manifest, or to just future proof the add-in, you need to convert it to use the unified manifest for Microsoft 365.

> [!NOTE]
>
> - Projects created in Visual Studio, as distinct from Visual Studio Code, can't be converted at this time.
> - If you [created the project with Teams Toolkit](teams-toolkit-overview.md) or with the "unified manifest" option in the [Yeoman generator for Office Add-ins (Yo Office)](yeoman-generator-overview.md), it already uses the unified manifest.

   [!INCLUDE [Unified manifest support note for Office applications](../includes/unified-manifest-support-note.md)]

There are three basic tasks to converting an add-in project from the add-in only manifest to the unified manifest.

- Ensure that your add-in is ready to convert.
- Convert the XML-formatted add-in only manifest itself to the JSON format of the unified manifest.
- Package the new manifest and two icon image files (described later) into a zip file for sideloading or deployment. *Depending on how you sideload the converted add-in, this task may be done for you automatically.*

[!INCLUDE [non-unified manifest clients note](../includes/non-unified-manifest-clients.md)]

> [!NOTE]
>
> - Add-ins that use the unified manifest can be sideloaded only on Office Version 2304 (Build 16320.20000) or later.
> - Projects created in Visual Studio, as distinct from Visual Studio Code, can't be converted at this time.
> - If you [created the project with Teams Toolkit](teams-toolkit-overview.md) or with the "unified manifest" option in the [Office Yeoman Generator](yeoman-generator-overview.md), it already uses the unified manifest.

## Ensure that your add-in is ready to convert

The following sections describe conditions that must be met before you convert the manifest.

### Uninstall the existing version of the add-in

To avoid conflicts with UI control names and other problems, be sure the existing add-in isn't installed on the computer where you do the conversion. If you experience any difficulties uninstalling the add-in, see [Remove a ghost add-in](../testing/uninstall-add-in.md#remove-a-ghost-add-in).

### Ensure that you have two special image files

If your add-in only manifest doesn't already have both **\<IconUrl\>** and **\<HighResolutionIconUrl\>** (in that order) elements, then add them just below the **\<Description\>** element. The values of the **DefaultValue** attribute should be the full URLs of image files. The images must be a specified size as shown in the following table.

|Office application|\<IconUrl\>|\<HighResolutionIconUrl\>|
|:---------------|:---------------|:---------------|
|Outlook|64x64 pixels|128x128 pixels|
|All other Office</br>applications|32x32 pixels|64x64 pixels|

The following markup is an example.

```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="MailApp">
  <Id>01234567-89ab-cdef-0123-4567-89abcdef0123</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="Great Add-in"/>
  <Description DefaultValue="A great add-in."/>
  <IconUrl DefaultValue="https://localhost:3000/assets/icon-64.png" />
  <HighResolutionIconUrl DefaultValue="https://localhost:300/assets/icon-128.png" />

  <!-- Other markup omitted -->
```

### Reduce the number of add-in commands as needed

An add-in that uses the unified manifest may not have more than 20 [add-in commands](../design/add-in-commands.md). If the total number of  [**\<Action\>** elements](/javascript/api/manifest/action) in the add-in only manifest is greater than 20, you must redesign the add-in to have no more than 20.

### Update the add-in ID, version, domain, and function names in the manifest

1. Change the value of the `<ID>` element to a new random GUID.

1. Update the value of the `<Version>` element and ensure that it conforms to the [semver standard](https://semver.org/) (MAJOR.MINOR.PATCH). Each segment can have no more than five digits. For example, change the value `1.0.0.0` to `1.0.1`. The semver standard's prerelease and metadata version string extensions aren't supported.

1. Be sure that the domain segment of the add-in's URLs in the manifest are pointing to `https://localhost:3000`.

1. If your manifest has any **\<FunctionName\>** elements, make sure their values have fewer than 65 characters.

   > [!IMPORTANT]
   > The value of this element must exactly match the name of an action that's mapped to a function in a JavaScript or TypeScript file with the [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member(1)) function. If you change it in the manifest, be sure to change it in the `actionId` parameter passed to `associate()` too.

### Shorten the add-in's name and description as needed

Review and change, as needed, the `<DisplayName>` and `<Description>` values in light of the following effects of the conversion.

- The first 30 characters of `<DisplayName>` will become the value of "name.short" in the unified manifest.
- The first 100 characters of `<DisplayName>` will become the value of "name.long" in the unified manifest.
- The first 250 characters of `<Description>` will become the value of "description.short" in the unified manifest.
- The first 4000 characters of `<Description>` will become the value of "description.long" in the unified manifest.

### Verify that the modified add-in only manifest works

1. Validate the modified add-in only manifest. See [Validate an Office Add-in's manifest](../testing/troubleshoot-manifest.md).

1. Verify that the add-in can be sideloaded and run. See [Sideload an Office Add-in for testing](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

Resolve any problems before you attempt to convert the project.

## Conversion tools and options

There are several ways to carry out the remaining tasks, depending on the IDE and other tools you want to use for your project, and on the tool you used to create the project.

- [Convert the project with Teams Toolkit](#convert-the-project-with-teams-toolkit)
- [Convert projects created with the Yeoman generator for Office Add-ins (aka "Yo Office")](#convert-projects-created-with-the-yeoman-generator-for-office-add-ins-aka-yo-office)
- [Convert NodeJS and npm projects that weren't created with the Yeoman generator for Office Add-ins (Yo Office)](#convert-nodejs-and-npm-projects-that-werent-created-with-the-yeoman-generator-for-office-add-ins-yo-office)

### Convert the project with Teams Toolkit

The easiest way to convert is to use Teams Toolkit.

#### Prerequisites

- Install [Visual Studio Code](https://code.visualstudio.com/)
- Install [Teams Toolkit](/microsoftteams/platform/toolkit/install-teams-toolkit?tabs=vscode#install-teams-toolkit-for-visual-studio-code)

#### Import the add-in project to Teams Toolkit

1. Open Visual Studio Code and select the Teams Toolkit icon on the **Activity Bar**.

    :::image type="content" source="../images/teams-toolkit-icon.png" alt-text="Teams Toolkit icon.":::

1. Select **Create a New App**.
1. The **New Project** dropdown menu opens. The options listed will vary depending on your version of Teams Toolkit. Select **Office Add-in**.

    :::image type="content" source="../images/teams-toolkit-create-office-add-in.png" alt-text="The options in New Project dropdown menu. One option is called 'Office Add-in'.":::

1. The **App Features Using an Office Add-in** dropdown menu opens. The options listed will vary depending on your version of Teams Toolkit. Select **Import an Existing Office Add-in**.

    :::image type="content" source="../images/teams-toolkit-create-office-import-capability.png" alt-text="The options in the App Features Using an Office Add-in dropdown menu. The 'Import an Existing Office Add-in' option is selected.":::

1. In the **Existing add-in project folder** dropdown menu, browse to the root folder of the add-in project.
1. In the **Select import project manifest file** dropdown menu, browse to the add-in only manifest file, typically named **manifest.xml**.
1. In the **Workspace folder** dialog, select the folder where you want to put the converted project.
1. In the **Application name** dialog, give a name to the project (with no spaces). Teams Toolkit creates the project with your source files and scaffolding. It then opens the project *in a second Visual Studio Code window*. Close the original Visual Studio Code window.

#### Sideload the add-in in Visual Studio Code

You can sideload the add-in using the Teams Toolkit or in a command prompt, bash shell, or terminal. For more information, see:

- [Sideload with Teams toolkit](../testing/sideload-add-in-with-unified-manifest.md#sideload-with-the-teams-toolkit)
- [Sideload with a system prompt, bash shell, or terminal](../testing/sideload-add-in-with-unified-manifest.md#sideload-with-a-system-prompt-bash-shell-or-terminal)

> [!NOTE]
> Add-ins that use the unified manifest can be sideloaded only on Office Version 2304 (Build 16320.20000) or later.

### Convert projects created with the Yeoman generator for Office Add-ins (aka "Yo Office")

If the project was created with the Yeoman generator for Office Add-ins and you don't want to use the Teams Toolkit, convert it using the following steps.

1. In the root of the project, open a command prompt or bash shell and run the following command. This converts the manifest and updates the package.json to specify current tooling packages. The new unified manifest is in the root of the project and the old add-in only manifest is in a backup.zip file. For details about this command, see [Office-Addin-Project](https://www.npmjs.com/package/office-addin-project).

    ```command&nbsp;line
    npx office-addin-project convert -m <relative-path-to-XML-manifest>
    ```

1. Run `npm install`.
1. To sideload the add-in, see [Sideload add-ins created with the Yeoman generator for Office Add-ins (Yo Office)](../testing/sideload-add-in-with-unified-manifest.md#sideload-add-ins-created-with-the-yeoman-generator-for-office-add-ins-yo-office).

### Convert NodeJS and npm projects that weren't created with the Yeoman generator for Office Add-ins (Yo Office)

If you don't want to use the Teams Toolkit and your project wasn't created with Yo Office, use the office-addin-manifest-converter tool.

In the root of the project, open a command prompt or bash shell and run the following command. This command puts the unified manifest in a subfolder with the same name as the filename stem of the original add-in only manifest. For example, if the manifest is named **MyManifest.xml**, the unified manifest is created at **.\MyManifest\MyManifest.json**. For more details about this command, see [Office-Addin-Manifest-Converter](https://www.npmjs.com/package/office-addin-manifest-converter).

```command&nbsp;line
npx office-addin-manifest-converter convert <relative-path-to-XML-manifest>
```

Once you have the unified manifest created, there are two ways to create the zip file and sideload it. For more information, see [Sideload other NodeJS and npm projects](../testing/sideload-add-in-with-unified-manifest.md#sideload-other-nodejs-and-npm-projects).

## Next steps

Consider whether to maintain both the old and new versions of the add-in. See [Manage both a unified manifest and an add-in only manifest version of your Office Add-in](../concepts/duplicate-legacy-metaos-add-ins.md).
