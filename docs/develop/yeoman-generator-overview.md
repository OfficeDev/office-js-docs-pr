---
title: Create Office Add-in projects using the Yeoman Generator
description: Learn how to create Office Add-in projects using the Yeoman generator for Office Add-ins.
ms.topic: tutorial
ms.date: 07/07/2025
ms.localizationpriority: high
---

# Create Office Add-in projects using the Yeoman Generator

The [Yeoman Generator for Office Add-ins](https://github.com/OfficeDev/generator-office) (also called "Yo Office") is an interactive Node.js-based command line tool that creates Office Add-in development projects. These projects are Node.js-based. When you want the server-side code of the add-in to be in a .NET-based language (such as C# or VB.Net) or you want the add-in hosted in Internet Information Server (IIS), [use Visual Studio to create the add-in](develop-add-ins-visual-studio.md).

> [!NOTE]
> Office add-ins can also be created with the [Microsoft 365 Agents Toolkit](agents-toolkit-overview.md) or the [Office Add-in Development Kit](development-kit-overview.md).

The projects that the tool creates have the following characteristics.

- They have a standard [npm](https://www.npmjs.com/) configuration that includes a **package.json** file.
- They include several helpful scripts to build the project, start the server, sideload the add-in in Office, and other tasks.
- They use [webpack](https://webpack.js.org/) as a bundler and basic task runner.
- In development mode, they are hosted on localhost by webpack's Node.js-based webpack-dev-server, a development-oriented version of the [express](http://expressjs.com/) server that supports hot-reloading and recompile-on-change.
- By default, all dependencies are installed by the tool, but you can postpone the installation with a command line argument.
- They include a complete add-in manifest.
- They have a "Hello World"-level add-in that is ready run as soon as the tool has finished.
- They include a transpiler to transpile TypeScript to ES5 JavaScript, and a polyfill to enable ES5 JavaScript to use features from later versions of JavaScript. Together, they provide backward compatibility with legacy webviews such as Trident (Internet Explorer), although Microsoft doesn't support development of add-ins for versions of Office that use these old webviews.

> [!TIP]
> If you want to deviate from these choices significantly, such as using a different task runner or a different server, we recommend that when you run the tool you choose the [Manifest-only option](#manifest-only-option).

## Prerequisites

>[!NOTE]
> If you aren't familiar with Node.js or npm, you should start by [setting up your development environment](../overview/set-up-your-dev-environment.md).

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## Use the tool

Start the tool with the following command in a system prompt (not a bash window). This will create a new project in a new folder in the current directory.

```command&nbsp;line
yo office 
```

A lot needs to load, so it may take 40 seconds before the tool starts. The tool asks you a series of questions. For some, you just type an answer to the prompt. For others, you're given a list of possible answers. If given a list, use the up and down arrow keys to select one and then select <kbd>Enter</kbd>.

The first question asks you to choose between several types of projects. The options are:

- **Office Add-in Task Pane project**
- **Excel, PowerPoint, and/or Word Task Pane with unified manifest for Microsoft 365 (preview)**
- **Office Add-in Task Pane project using React framework**
- **Excel Custom Functions using a Shared Runtime**
- **Excel Custom Functions using a JavaScript-only Runtime**
- **Office Add-in Task Pane project supporting single sign-on**
- **Office Add-in Task Pane project supporting Nested App Auth single sign-on (preview)**
- **Office Add-in project containing the manifest only**

:::image type="content" source="../images/yo-office-project-type-prompt.png" alt-text="The prompt for project type, and the possible answers, in the Yeoman generator.":::

> [!NOTE]
> - The **Office Add-in project containing the manifest only** option produces a project that contains a basic add-in manifest and minimal scaffolding. For more information about the option, see [Manifest-only option](#manifest-only-option).
> - The **Excel, PowerPoint, and/or Word Task Pane with unified manifest for Microsoft 365 (preview)** option creates a project for Excel, PowerPoint, Word, or all three, that uses the unified manifest for Microsoft 365. For more information about the option, see [Word, PowerPoint, or Excel with unified manifest option](#word-powerpoint-or-excel-with-unified-manifest-option).

The next question asks you to choose between **TypeScript** and **JavaScript**. (This question is skipped if you chose the manifest-only option in the preceding question.)

:::image type="content" source="../images/yo-office-language-prompt.png" alt-text="The Yo Office interface after the user chose 'Office Add-in Task Pane project' to the preceding question. It shows the prompt for language, and the possible answers, TypeScript and JavaScript, in the Yeoman generator.":::

You're then prompted to give the add-in a name. The name you specify will be used in the add-in's manifest, but you can change it later. This is also the folder name for the project.

:::image type="content" source="../images/yo-office-name-prompt.png" alt-text="The Yo Office interface after the user chose TypeScript to the previous question. It shows the prompt for the add-in name in the Yeoman generator.":::

You're then prompted to choose which Office application the add-in should run in. There are six possible applications to choose from: **Excel**, **OneNote**, **Outlook**, **PowerPoint**, **Project**, and **Word**. You must choose just one, but you can change the manifest later to support the additional Office applications. The exception is Outlook. A manifest that supports Outlook cannot support any other Office application.

:::image type="content" source="../images/yo-office-host-prompt.png" alt-text="The Yo Office interface after the user named the project 'Contoso Add-in'. It shows the prompt for Office application, and the possible answers, in the Yeoman generator.":::

If you choose **Outlook** as the Office application, you get an additional question asking you which type of manifest you want to use. We recommend that you choose **unified manifest for Microsoft 365** unless your add-in will include an extensibility feature that isn't yet supported by the unified manifest.

[!INCLUDE [non-unified manifest clients note](../includes/non-unified-manifest-clients.md)]

After you've answered all questions, the generator creates the project and installs the dependencies. You may see **WARN** messages in the npm output on screen. You can ignore these. You may also see messages that vulnerabilities were found. You can ignore these for now, but you'll eventually need to fix them before your add-in is released to production. For more information, see [Warnings and dependencies in the Node.js and npm world](../overview/npm-warnings-advice.md).

If the creation is successful, you'll see a **Congratulations!** message in the command window, followed by some suggested next steps. (If you're using the generator as part of a quick start or tutorial, ignore the next steps in the command window and continue with the instructions in the article.)

> [!TIP]
> If you want to create the scaffolding of an Office Add-in project, but postpone the installation of the dependencies, add the `--skip-install` option to the `yo office` command. The following code is an example.
>
> ```command&nbsp;line
> yo office --skip-install
> ```
>
> When you're ready to install the dependencies, navigate to the root folder of the project in a command prompt and enter `npm install`.

> [!WARNING]
> If you choose **Office Add-in Task Pane project supporting single sign-on** and **TypeScript**, and you are using a Node.js version greater than 18.16.0, then a bug in Node.js may cause the project file **\<root\>\src\middle-tier\ssoauth-helper.ts** to be corrupted. To fix it, copy the contents of the file from the repo, [ssoauth-helper.ts](https://github.com/OfficeDev/Office-Addin-Taskpane-SSO/blob/master/src/middle-tier/ssoauth-helper.ts), over the contents of the file in the generated project.

## Manifest-only option

This option creates only a manifest for an add-in. The resulting project doesn't have a Hello World add-in, any of the scripts, or any of the dependencies. Use this option in the following scenarios.

- You want to use different tools from the ones a Yeoman generator project installs and configures by default. For example, you want to use a different bundler, transpiler, task runner, or development server.
- You want to use a web application development framework, other than React, such as Vue.

## Word, PowerPoint, or Excel with unified manifest option

The unified manifest for Microsoft 365 is in preview for Excel, PowerPoint, and Word add-ins. It should not be used for production add-ins, but you can select this option in Yo Office to create an add-in for one (or all three) of those Office applications. You'll be asked to choose which Office application. You can also choose **All** to create an add-in that is installable on all three Office applications. The project that is created uses TypeScript.

## Use command line parameters

You can also add parameters to the `yo office` command. The two most common are:

- `yo office --details`: This will output brief help about all of the other command line parameters.
- `yo office --skip-install`: This will prevent the generator from installing the dependencies.

For detailed reference about the command line parameters, see the readme for the generator at [Yeoman generator for Office Add-ins](https://github.com/officedev/generator-office).

## Troubleshooting

If you encounter problems using the tool, your first step should be to reinstall it to be sure that you have the latest version. (See [Prerequisites](#prerequisites) for details.) If doing so doesn't fix the problem, search the [issues of the GitHub repo for the tool](https://github.com/OfficeDev/generator-office/issues) to see if anyone else has encountered the same problem and found a solution. If no one has, [create a new issue](https://github.com/OfficeDev/generator-office/issues/new?assignees=&labels=needs+triage&template=bug_report.md&title=).
