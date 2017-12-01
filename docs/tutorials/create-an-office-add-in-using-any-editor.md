---
title: Create an Office Add-in using any editor
description: 
ms.date: 11/20/2017 
---



# Create an Office Add-in using any editor

You can use the Yeoman generator to create your Office Add-in. The Yeoman generator provides the project scaffolding and build management. The  `manifest.xml` file tells the Office application where your add-in is located and how you want it to appear. The Office application takes care of hosting it within Office.

> [!NOTE]
> These instructions use Terminal on a Mac, but you can also use other shell environments. 

## Prerequisites for the Yeoman generator

To install the Yeoman Office generator, you must have [git](https://git-scm.com/downloads) and node.js installed on your computer. If you're on a Mac, we recommend that you use [Node Version Manager](https://github.com/creationix/nvm) to install node.js with the right permissions. If you're on Windows, you can install node.js from [nodejs.org](https://nodejs.org/en/).

> [!NOTE]
> If you're on Windows, use the default values when you install git, with the following exceptions:
> - Use git from the Windows command prompt.
> - Use the Windows default console window.

After you install node.js, open a Terminal and install the generator globally.

```bash
npm install -g yo generator-office
```


## Create the default files for your add-in

The Yeoman generator runs in the directory where you want to scaffold the project. Before you develop an Office Add-in, you should first create a folder for your project.

In Terminal, move to the parent folder where you want to create your project. Then use the following commands create a new folder named  _myHelloWorldaddin_ and shift the current directory to it:




```bash
mkdir myHelloWorldaddin
cd myHelloWorldaddin
```

Use the Yeoman generator to create the add-in of your choice. The steps in this article create a simple task pane add-in. To run the generator, enter the following command:




```bash
yo office
```

**Yeoman generator input for an add-in**

The generator will prompt you for the following: 


- New subfolder -- use _N_
- Add-in name -- use  _myHelloWorldaddin_ 
- The supported Office application - you can choose any application
- Create new add-in -- use _Yes, I want a new add-in._
- Add [TypeScript](https://www.typescriptlang.org/) -- use _N_
- Choose framework -- use _Jquery_

> [!NOTE]
> If you want to create an Office Add-in that uses Office UI Fabric React, enter the following:
> - Add [TypeScript](https://www.typescriptlang.org/) -- use _Y_
> - Choose framework -- use _React_

![Gif of yeoman generator prompting for project input](../images/gettingstarted-fast.gif)

This creates the structure and basic files for your add-in.


## Hosting your Office Add-in

Office Add-ins must be hosted, even in development, via HTTPS. Yo Office creates a bsconfig.json, which uses Browsersync to make it faster for you to tweak and test your add-in by synchronizing file changes across multiple devices. 

Launch the local HTTPS site on https://localhost:3000 by typing the following command in your console:


```bash
npm start
```

Browsersync will start a HTTPS server, and launch the index.html file in your project. You will see an error that states "There is a problem with this website's security certificate."


![Gif showing process to bypass error and see default index.html file](../images/ssl-chrome-bypass.png)

This error occurs because Browsersync includes a self-signed SSL certificate that your development environment must trust. For information about how to resolve this error, see [adding self-signed certificates](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

## Sideload the add-in into Office

You can use sideloading to install your add-in for testing within the Office clients:

- [Sideload Office Add-ins for testing](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Sideload Office Add-ins on iPad and Mac for testing](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)   
- [Sideload Outlook add-ins for testing](https://docs.microsoft.com/en-us/outlook/add-ins/testing-and-tips)

## Develop your Office Add-in

You can use any text editor to develop the files for your custom Office Add-in.

> [!IMPORTANT]
> The manifest-myHelloWorldaddin.xml file tells the Office client applications how to interact with your add-in. The value in the `<id>` tag is a GUID that Yo Office creates when it generates the project. Do not change the GUID for your add-in. If the host is Azure, the `SourceLocation` value will be a URL that is similar to _https:// [name-of-your-web-app].azurewebsites.net/[path-to-add-in]_. If you are using the self-hosted option, as in this example, it will be _https://localhost:3000/[path-to-add-in]_.


## Debug your Office Add-in


You can debug your add-in in several ways:

- Attach a debugger from the task pane (Office 2016 for Windows).
- Use your browser's developer tools.
- Use F12 developer tools in Windows 10.

### Attach debugger from the task pane

In Office 2016 for Windows, Build 77xx.xxxx or later, you can attach the debugger from the task pane. The attach debugger feature will directly attach the debugger to the correct Internet Explorer process for you. You can attach a debugger regardless of whether you are using Yeoman Generator, Visual Studio Code, node.js, Angular, or another tool. 

For more information, see [Attach debugger from the task pane](../testing/attach-debugger-from-task-pane.md).


### Browser developer tools 

You can use the Office web clients and open the browser's developer tools to debug your add-in the way you debug any other client-side JavaScript application. 

### F12 developer tools on Windows 10

If you're using the Office desktop client on Windows 10, you can [Debug add-ins using F12 developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).
    
## Next steps

- [Deploy and publish your Office Add-in](../publish/publish.md)
    
