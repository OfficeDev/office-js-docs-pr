---
title: Set up your development environment
description:  Set up your developer environment to build Office Add-ins.
ms.date: 01/15/2026
ms.topic: install-set-up-deploy
ms.localizationpriority: medium
---

# Set up your development environment

This guide helps you set up tools so you can create Office Add-ins by following our quick starts or tutorials. If you already have these installed, you're ready to begin a quick start, such as this [Excel React quick start](../quickstarts/excel-quickstart-react.md).

## Get Microsoft 365

You need a Microsoft 365 account. You might qualify for a Microsoft 365 E5 developer subscription, which includes all Office apps, through the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram); for details, see the [FAQ](/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-). Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/business/compare-all-microsoft-365-business-products-g).

## Install the environment

There are three kinds of development environments to choose from. The scaffolding of Office Add-in projects that is created in the three environments is different, so if multiple people will be working on an add-in project, they must all use the same environment.

- **Node.js environment**: In this environment, your tools are installed and run at a command line. The server-side of the web application part of the add-in is written in JavaScript or TypeScript and is hosted in a Node.js runtime. There are many helpful add-in development tools in this environment, such as an Office linter and a bundler/task-runner called webpack. The project creation and scaffolding tool is a command line tool called the Office Yeoman Generator (also called "Yo Office"), though you can still use the Visual Studio Code extensions mentioned in the next option.
- **Visual Studio Code**: Choose this environment if you use Visual Studio Code and would prefer to create projects from an extension rather than command line tools. The project creation and scaffolding tool is Microsoft 365 Agents Toolkit.
- **Visual Studio environment**: Choose this environment only if your development computer is Windows, and you want to develop the server-side of the add-in with a .NET based language and framework, such as ASP.NET. The add-in project templates in Visual Studio aren't updated as frequently as those in the Node.js environment. More information later on the **Visual Studio environment** tab.

> [!NOTE]
> Visual Studio for Mac doesn't include the project scaffolding templates for Office Add-ins, so if your development computer is a Mac, you should work with the Node.js or Visual Studio Code environments.

Select the tab for the environment you choose.

# [Node.js environment](#tab/yeomangenerator)

The main tools to be installed are:

- Node.js
- npm
- A code editor of your choice
- Office Yeoman Generator (Yo Office)
- The Office JavaScript linter

This guide assumes that you know how to use a command-line tool.

[!INCLUDE [Instructions to install Node.js.](../includes/install-node-js.md)]

### Install a code editor

You can use any code editor or IDE that supports client-side development to build your web part, such as:

- [Visual Studio Code](https://code.visualstudio.com/) (recommended)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

### Install the Yeoman generator &mdash; Yo Office

The project creation and scaffolding tool is [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md), affectionately known as **Yo Office**. You need to install the latest version of [Yeoman](https://github.com/yeoman/yo) and Yo Office. To install these tools globally, run the following command via the command prompt.

  ```command&nbsp;line
  npm install -g yo generator-office
  ```

### Install and use the Office JavaScript linter

Microsoft provides a JavaScript linter to help you catch common errors when using the Office JavaScript library. If you create an add-in project with either the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) or [Agents Toolkit](../develop/agents-toolkit-overview.md), then the linter is installed and configured for you. Skip to [Run the linter](#run-the-linter).

If you created your project manually, install and configure the linter with the following steps.

1. In the root of the project, run the following two commands (after you've [installed Node.js and npm](#install-nodejs-and-npm)).

   ```command&nbsp;line
   npm install office-addin-lint --save-dev
   npm install eslint-plugin-office-addins --save-dev
   ```

1. In the root of the project, create a text file named **eslint.config.js** (or **.mjs**), if there isn't one already there. Be sure to inherit the recommended configuration for `eslint-plugin-office-addins`. The `plugins` object should include a mapping to the `eslint-plugin-office-addins` plugin object. The following is a simple example that includes settings for TypeScript. Your **eslint.config.js** file may have additional properties and configurations.

   ```js
    const officeAddins = require("eslint-plugin-office-addins");
    const tsParser = require("@typescript-eslint/parser");
    const tsEsLint = require("typescript-eslint");
    
    export default [
      ...tsEsLint.configs.recommended,
      ...officeAddins.configs.recommended,
      {
        plugins: {
          "office-addins": officeAddins,
        },
        languageOptions: {
          parser: tsParser,
        },
      },
    ];
   ```

1. In the root of the project, open the **package.json** file and be sure that the `scripts` array has the following member.

   ```json
   "lint": "office-addin-lint check",
   ```

#### Run the linter

Run the linter with the following command either in the terminal of an editor, such as Visual Studio Code, or in a command prompt. Problems found by the linter appear in the terminal or prompt, and also appear directly in the code when you're using an editor that supports linter messages, such as Visual Studio Code.

```command&nbsp;line
npm run lint
```

# [Visual Studio Code](#tab/visualstudiocode)

The main tools to be installed are:

- Node.js
- npm
- Visual Studio Code
- Microsoft 365 Agents Toolkit

This guide assumes that you know how to use a command-line tool.

[!INCLUDE [Instructions to install Node.js.](../includes/install-node-js.md)]

### Install Visual Studio Code

Get the latest version of Visual Studio Code from [Visual Studio Code homepage](https://code.visualstudio.com). Select the appropriate option for your platform of choice.

### Install Agents Toolkit

When Visual Studio Code is installed, install the latest version of Agents Toolkit as described in [Install Microsoft 365 Agents Toolkit](/microsoftteams/platform/toolkit/install-teams-toolkit?tabs=vscode).

# [Visual Studio environment](#tab/visualstudio)

This option only installs Visual Studio.

[!INCLUDE [Deprecation warning for add-in projects in Visual Studio](../includes/vs-deprecation.md)]

### Install Visual Studio

Install the latest version of Visual Studio from [Visual Studio Downloads](https://visualstudio.microsoft.com/downloads/). Be sure to include the **Office/SharePoint development** workload when the installer asks you to specify workloads. Other workloads that you may need are **Web development tools for .NET**, **JavaScript and TypeScript language support** (for coding the client-side of the add-in), and ASP.NET-related workloads.

> [!NOTE]
> For information about debugging client-side code when you're using the Visual Studio environment, see [Debug Office Add-ins in Visual Studio](../develop/debug-office-add-ins-in-visual-studio.md). Debug the server-side code the same way you would any web application created in Visual Studio. See [Client-side or server-side](../testing/debug-add-ins-overview.md#server-side-or-client-side).

---

## Install Script Lab

Script Lab is a tool for quickly prototyping code that calls the Office JavaScript Library APIs. Script Lab is itself an Office Add-in and can be installed from Microsoft Marketplace at [Script Lab](https://marketplace.microsoft.com/marketplace/apps?search=script%20lab&page=1). There's a version for Excel, PowerPoint, and Word, and a separate version for Outlook. For information about how to use Script Lab, see [Explore Office JavaScript API using Script Lab](explore-with-script-lab.md).

## Next steps

Try creating your own add-in or use [Script Lab](explore-with-script-lab.md) to try built-in samples.

### Create an Office Add-in

You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](../index.yml). If you've previously completed a quick start and want to create a slightly more complex add-in, you should try a [tutorial](../index.yml).

### Explore the APIs with Script Lab

Explore the library of built-in samples in [Script Lab](explore-with-script-lab.md) to get a sense for the capabilities of the Office JavaScript APIs.

## See also

- [Core concepts for Office Add-ins](../overview/core-concepts-office-add-ins.md)
- [Developing Office Add-ins](../develop/develop-overview.md)
- [Design Office Add-ins](../design/add-in-design.md)
- [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
- [Publish Office Add-ins](../publish/publish.md)
- [Learn about the Microsoft 365 Developer Program](https://aka.ms/m365devprogram)
