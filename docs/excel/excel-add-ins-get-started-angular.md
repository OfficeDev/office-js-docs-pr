---
title: Build an Excel add-in using Angular
description: 
ms.date: 11/20/2017 
---


# Build an Excel add-in using Angular

In this article, you'll walk you through the process of building an Excel add-in using Angular and the Excel JavaScript API.

## Prerequisites

If you haven't done so previously, install the following tools:

1. Check whether you already have the [Angular CLI prerequisites](https://github.com/angular/angular-cli#prerequisites) and install any prerequistes that you are missing.

2. Install the [Angular CLI](https://github.com/angular/angular-cli) globally. 

    ```bash
    npm install -g @angular/cli
    ```

3. Install [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.

    ```bash
    npm install -g yo generator-office
    ```

## Generate a new Angular app

Use the Angular CLI to generate your Angular app. From the terminal, run the following command:

```bash
ng new my-addin
```

## Generate the manifest file and sideload the add-in

An add-in's manifest file defines its settings and capabilities.

1. Navigate to your app folder.

    ```bash
    cd my-addin
    ```

2. Use the Yeoman generator to generate the manifest file for your add-in. Run the following command and then answer the prompts as shown in the screenshot below.

    ```bash
    yo office
    ```
    ![Yeoman generator](../images/yo-office.png)
    
    > [!NOTE]
    > If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).

3. Open the manifest file (i.e., the file in the root directory of your app with a name ending in "manifest.xml"). Replace all occurrences of `https://localhost:3000` with `http://localhost:4200` and save the file.

    > [!NOTE]
    > Be sure to change the protocol to **http** in addition to changing the port number to **4200**.

4. Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.

    - Windows: [Sideload Office Add-ins for testing on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

## Update the app

1. Open **src/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

2. Open **src/main.ts**, replace `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` with the following code, and save the file. 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

3. Open **src/polyfills.ts**, add the following line of code above all other existing `import` statements, and save the file.

    ```typescript
    import 'core-js/client/shim';
    ```

4. In **src/polyfills.ts**, uncomment the following lines, and save the file.

    ```typescript
    import 'core-js/es6/symbol';
    import 'core-js/es6/object';
    import 'core-js/es6/function';
    import 'core-js/es6/parse-int';
    import 'core-js/es6/parse-float';
    import 'core-js/es6/number';
    import 'core-js/es6/math';
    import 'core-js/es6/string';
    import 'core-js/es6/date';
    import 'core-js/es6/array';
    import 'core-js/es6/regexp';
    import 'core-js/es6/map';
    import 'core-js/es6/weak-map';
    import 'core-js/es6/set';
    ```

5. Open **src/app/app.component.html**, replace file contents with the following HTML, and save the file. 

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <button (click)="onColorMe()">Color Me</button>
        </div>
    </div>
    ```

6. Open **src/app/app.component.css**, replace file contents with the following CSS code, and save the file.

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

7. Open **src/app/app.component.ts**, replace file contents with the following code, and save the file. 

    ```typescript
    import { Component } from '@angular/core';

    declare const Excel: any;

    @Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.css']
    })
    export class AppComponent {
    onColorMe() {
        Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = 'green';
        await context.sync();
        });
    }
    }
    ```

## Try it out

1. From the terminal, run the following command to start the dev server.

    ```bash
    npm start
    ```

2. In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.

    ![Excel Add-in button](../images/excel-quickstart-addin-2a.png)

3. Choose the **Color Me** button in the task pane to set the color of the selected range to green.

    ![Excel Add-in](../images/excel-quickstart-addin-2b.png)

## Next steps

Congratulations, you've successfully created an Excel add-in using Angular! Next, learn more about the [core concepts](excel-add-ins-core-concepts.md) of building Excel add-ins.

## Additional resources

* [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md)
* [Explore snippets with Script Lab](https://store.office.com/en-001/app.aspx?assetid=WA104380862&ui=en-US&rs=en-001&ad=US&appredirect=false)
* [Excel add-in code samples](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Excel JavaScript API reference](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
