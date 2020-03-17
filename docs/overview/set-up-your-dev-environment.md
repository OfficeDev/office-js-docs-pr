---
title: Set up your development environment
description:  'Set up your developer environment so you can build Office Add-ins' 
ms.date: 03/17/2020
localization_priority: Normal
---

# Set up your development environment

This guide will help you set up tools so you can create Office Add-ins following our quick starts or tutorials. You'll need to install tools in the list below. If you already have these installed, you should be ready to begin a quick start, such as this [Excel React quick start](../quickstarts/excel-quickstart-react.md).

- Node.js
- npm
- An Office 365 (the subscription version of Office) account
- A code editor of your choice

This guide assumes that you know how to use the command line and have one installed.

## Install Node.js

Node.js is a JavaScript runtime you will need in order to develop modern Office Add-ins.

To install Node.js, visit [their website to install the latest version](https://nodejs.org/about/releases) and follow the instructions.

## Install npm

Npm is an open source software registry, which allows you to download the packages which you will need to develop Office Add-ins.

To install npm, run the following in the command line:

```command&nbsp;line
    npm install npm -g
```

To check if you already have npm installed and see the installed version, run the following in the command line:

```command&nbsp;line
npm -v
```

You may wish to use a Node version manager to allow you to switch between multiple versions of Node.js and npm, but this is not strictly necessary. For details on how to do this, [see npm's instructions](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).

## Get Office 365

If you don't already have an Office 365 account, you can get a free, 90-day renewable Office 365 subscription by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).

## Install a code editor

You can use any code editor or IDE that supports client-side development to build your web part, such as:

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

## Next steps

Try creating your own add-in or use Script Lab to try built-in samples.

### Create an Office add-in

You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](../index.md). If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](../index.md).

### Explore the APIs with Script Lab

Explore the library of built-in samples in [Script Lab](explore-with-script-lab.md) to get a sense for the capabilities of the Office JavaScript APIs.

## See also

- [Building Office Add-ins](../overview/office-add-ins-fundamentals.md)
- [Core concepts for Office Add-ins](../overview/core-concepts-office-add-ins.md)
- [Develop Office Add-ins](../develop/develop-overview.md)
- [Design Office Add-ins](../design/add-in-design.md)
- [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md)
- [Publish Office Add-ins](../publish/publish.md)
