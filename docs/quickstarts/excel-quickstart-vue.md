---
title: Build an Excel task pane add-in using Vue
description: Learn how to build a simple Excel task pane add-in by using the Office JS API and Vue.
ms.date: 01/16/2020
ms.prod: excel
localization_priority: Priority
---

# Build an Excel task pane add-in using Vue

In this article, you'll walk through the process of building an Excel task pane add-in using Vue and the Excel JavaScript API.

## Prerequisites

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Install the [Vue CLI](https://cli.vuejs.org/) globally.

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## Generate a new Vue app

Use the Vue CLI to generate a new Vue app. From the terminal, run the following command.

```command&nbsp;line
vue create my-add-in
```

Then select the `default` preset. If you are prompted to use either Yarn or NPM as a package you can choose either one.

## Generate the manifest file

Each add-in requires a manifest file to define its settings and capabilities.

1. Navigate to your app folder.

    ```command&nbsp;line
    cd my-add-in
    ```

2. Use the Yeoman generator to generate the manifest file for your add-in by running the following command:

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > When you run the `yo office` command, you may receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools. Use the information that's provided to respond to the prompts as you see fit. If you choose **Exit** in response to the second prompt, you'll need to run the `yo office` command again when you're ready to create your add-in project.

    When prompted, provide the following information to create your add-in project:

    - **Choose a project type:** `Office Add-in project containing the manifest only`
    - **What do you want to name your add-in?** `my-office-add-in`
    - **Which Office client application would you like to support?** `Excel`

    ![Yeoman generator](../images/yo-office-manifest-only-vue.png)

After you complete the wizard, it creates a `my-office-add-in` folder, which contains a `manifest.xml` file. You will use the manifest to sideload and test your add-in at the end of the quick start.

> [!TIP]
> You can ignore the *next steps* guidance that the Yeoman generator provides after the add-in project's been created. The step-by-step instructions within this article provide all of the guidance you'll need to complete this tutorial.

## Secure the app

[!include[HTTPS guidance](../includes/https-guidance.md)]

To enable HTTPS for your app, create a `vue.config.js` file in the root folder of the Vue project with the following contents:

```js
module.exports = {
  devServer: {
    port: 3000,
    https: true
  }
};
```

## Update the app

1. Open the `public/index.html` file and add the following `<script>` tag immediately before the `</head>` tag:

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

2. Open `src/main.js` and replace the contents with the following code:

   ```js
   import Vue from 'vue';
   import App from './App.vue';

   Vue.config.productionTip = false;

   window.Office.initialize = () => {
     new Vue({
       render: h => h(App)
     }).$mount('#app');
   };
   ```

3. Open `src/App.vue` and replace the file contents with the following code:

   ```html
   <template>
     <div id="app">
       <div class="content">
         <div class="content-header">
           <div class="padding">
             <h1>Welcome</h1>
           </div>
         </div>
         <div id="content-main">
           <div class="padding">
             <p>
               Choose the button below to set the color of the selected range to
               green.
             </p>
             <br />
             <h3>Try it out</h3>
             <button @click="onSetColor">Set color</button>
           </div>
         </div>
       </div>
     </div>
   </template>

   <script>
     export default {
       name: 'App',
       methods: {
         onSetColor() {
           window.Excel.run(async context => {
             const range = context.workbook.getSelectedRange();
             range.format.fill.color = 'green';
             await context.sync();
           });
         }
       }
     };
   </script>

   <style>
     .content-header {
       background: #2a8dd4;
       color: #fff;
       position: absolute;
       top: 0;
       left: 0;
       width: 100%;
       height: 80px;
       overflow: hidden;
     }

     .content-main {
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
   </style>
   ```

## Start the dev server

1. From the terminal, run the following command to start the dev server.

   ```command&nbsp;line
   npm run serve
   ```

2. In a web browser, navigate to `https://localhost:3000` (notice the `https`). If your browser indicates that the site's certificate is not trusted, you will need to [configure your computer to trust the certificate](https://github.com/OfficeDev/generator-office/blob/fd600bbe00747e64aa5efb9846295a3f66d428aa/src/docs/ssl.md#add-certification-file-through-ie).

3. When the page on `https://localhost:3000` is blank and without any certificate errors, it means that it is working. The Vue App is mounted after Office is initialized, so it only shows things inside of an Excel environment.

## Try it out

1. Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.

   - Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
   - Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)
   - iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.

   ![Excel add-in button](../images/excel-quickstart-addin-2a.png)

3. Select any range of cells in the worksheet.

4. In the task pane, choose the **Set color** button to set the color of the selected range to green.

   ![Excel add-in](../images/excel-quickstart-addin-2c.png)

## Next steps

Congratulations, you've successfully created an Excel task pane add-in using Vue! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.

> [!div class="nextstepaction"]
> [Excel add-in tutorial](../tutorials/excel-tutorial.md)

## See also

* [Office Add-ins platform overview](../overview/office-add-ins.md)
* [Building Office Add-ins](../overview/office-add-ins-fundamentals.md)
* [Develop Office Add-ins](../develop/develop-overview.md)
* [Fundamental programming concepts with the Excel JavaScript API](../excel/excel-add-ins-core-concepts.md)
* [Excel add-in code samples](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)
