---
title: Deploy an Office Add-in using SSO to a Microsoft Azure app service | Microsoft Docs
description: Learn how to deploy an Office Add-in using SSO to an Azure app service.
ms.date: 12/20/2020
ms.localizationpriority: medium
---

# Deploy an Office Add-in using SSO to a Microsoft Azure app service

To deploy an Office Add-in using SSO to Azure, you need to create an Azure app service. The steps in this article will deploy your Office Add-in to a Microsoft Azure app service for staging or deployment.

## Requirements

This article assumes that you created an Office Add-in using the [Yeoman Generator for Office Add-ins](https://github.com/OfficeDev/generator-office) using the `Office Add-in Task Pane project supporting single sign-on (localhost)` project type. Be sure you have configured the add-in project so that it runs on localhost successfully.

The steps in this article also require:
- [Azure Account extension](https://marketplace.visualstudio.com/items?itemName=ms-vscode.azure-account) for VS Code.
- [Azure App Service extension](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azureappservice) for VS Code.

## Create the Azure app service

These steps set up a basic deployment of the Office Add-in. There are multiple ways to configure deployment that are not covered in this documentation. For additional options on how you may want to configure your deployment, see [Deployment Best Practices](/azure/app-service/deploy-best-practices)

1. Open your Office Add-in project in VS Code.
1. Select the Azure icon in the Activity Bar. If the Activity Bar is hidden, open it by selecting **View** > **Appearance** > **Activity Bar**.
1. Select **Sign in to Azure** to sign in to your Azure account. If you don't already have an Azure account, create one by selecting **Create an Azure Account**. Follow the provided steps to set up your account.

    :::image type="content" source="../images/azure-extension-sign-in.png" alt-text="Sign in to Azure button selected in the Azure extension.":::

1. Once you are signed in to your Azure account, right-click **App Services** and select **Create New Web App...(Advanced)**.

    :::image type="content" source="../images/azure-extension-create-resource-button.png" alt-text="Create resource.":::

1. On step 1 of **Create new web app**, enter a globally unique name for your app service; for example, **username-sso-add-in**.
1. On step 2 of **Create new web app**, select the resource group you created for this deployment.
1. On step 3 of **Create new web app**, select **Node 16 LTS** for the runtime stack.
1. On step 4 of **Create new web app**, select **Windows** for the OS.
1. On step 5 of **Create new web app**, select a location in your region.
1. On step 6 of **Create new web app**, select the Windows App Service plan you created for this deployment.
1. On step 7 of **Create new web app**, select **Skip for now**.

    Azure will create the app service and it will appear under **App Services** on Azure in the Activity Bar. Don't deploy the add-in yet.

1. Right-click your app service and select **Browse Website**.
1. When the browser for your new web site opens, copy the URL and save it. You'll need it in later steps.

## Update manifest

It's useful to maintain multiple manifests for testing across localhost, staging, and deployment. We recommend you copy the existing file and create a new manifest named **manifest-deployment.xml**.

1. Open the **manifest-deployment.xml** file.
1. Find all instances of the text `localhost:3000` and replace it with the domain of the app service URL you saved previously.
1. In the `<AppDomains>` section, add an `<AppDomain>` entry for the app service from the URL you saved previously. For example `<AppDomain>https://contoso-sso.azurewebsites.net</AppDomain>`.
1. Save the file.

## Update package.json

1. In the Visual Studio Code terminal, run the command `npm pkg set 'scripts.start'='node middletier.js'`. This will configure the start script to run Node JS on deployment.

## Update webpack.config.js

1. Open the **webpack.config.js** file.
1. Find the first `CopyWebpackPlugin` section and update it to also copy the package.json file to the dist folder as shown in the following example.

   ```javascript
    new CopyWebpackPlugin({
          patterns: [
            {
              from: "assets/*",
              to: "assets/[name][ext][query]",
            },
            {
              from: "package.json",
              to: "package.json",
            },
            {
              from: "manifest*.xml",
              to: "[name]" + "[ext]",
              transform(content) {
                if (dev) {
                  return content;
                } else {
                  return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
                }
              },
            },
          ],
        }),
   ```

1. Save the file.

## Update fallbackauthdialog.js

The following steps also work for fallbackauthdialog.ts if you created a TypeScript project.

1. Open the **src/helpers/fallbackauthdialog.js** file.
1. Find the `redirectUri` on line 24 and change the value to use the app service URL you saved previously. For example, `redirectUri: "https://contoso-sso.azurewebsites.net/fallbackauthdialog.html",`
1. Save the file.

## Update .ENV

The **.ENV** file contains a client secret. For the purposes of learning in this article you can deploy the **.ENV** file to Azure. However for a production deployment, you should move the secret and any other confidential data into [Azure Key Vault](/azure/key-vault/general/basic-concepts).

1. Open the **.ENV** file.
1. Remove the entry for `PORT=3000`. Azure app service will provide a PORT variable to your project when deployed.
1. Change `NODE_ENV` to have the value `production`.
1. Save the file.

## Update app.js

The following steps also work for app.ts if you created a TypeScript project.

1. Open the **src/middle-tier/app.js** file.
1. Replace the entire file contents with the following code.

    ```javascript
    /*
     * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
     *
     * This file is the main Node.js server file that defines the express middleware.
     */
    
    require("dotenv").config();
    import * as createError from "http-errors";
    import * as path from "path";
    import * as cookieParser from "cookie-parser";
    import * as logger from "morgan";
    import express from "express";
    import { getUserData } from "./msgraph-helper";
    import { validateJwt } from "./ssoauth-helper";
    
    /* global console, process, require, __dirname */
    
    const app = express();
    const port = process.env.PORT;
    
    app.set("port", port);
    
    // view engine setup
    app.set("views", path.join(__dirname, "views"));
    app.set("view engine", "pug");
    
    app.use(logger("dev"));
    app.use(express.json());
    app.use(express.urlencoded({ extended: false }));
    app.use(cookieParser());
    
    /* Turn off caching when developing */
    if (process.env.NODE_ENV !== "production") {
      app.use(express.static(path.join(process.cwd(), "dist"), { etag: false }));
    
      app.use(function (req, res, next) {
        res.header("Cache-Control", "private, no-cache, no-store, must-revalidate");
        res.header("Expires", "-1");
        res.header("Pragma", "no-cache");
        next();
      });
    } else {
      // In production mode, let static files be cached.
      app.use(express.static(path.join(process.cwd())));
      console.log("static set up: " + path.join(process.cwd()));
    }
    
    const indexRouter = express.Router();
    indexRouter.get("/", function (req, res) {
      //   res.render("/taskpane.html");
      res.sendFile("/taskpane.html", { root: __dirname });
    });
    
    app.use("/", indexRouter);
    
    app.get("/getuserdata", validateJwt, getUserData);
    
    // Catch 404 and forward to error handler
    app.use(function (req, res, next) {
      next(createError(404));
    });
    
    // error handler
    app.use(function (err, req, temp, res) {
      // set locals, only providing error in development
    
      res.locals.message = err.message;
      res.locals.error = req.app.get("env") === "development" ? err : {};
    
      // render the error page
      res.status(err.status || 500).send({
        message: err.message,
      });
    });
    
    app.listen(process.env.PORT, () => console.log("Server listening on port: " + process.env.PORT));
    ```

1. Save the file.

## Update app registration

We recommend you create multiple app registrations for localhost, staging, and deployment testing. The following steps ensure that the app registration you use for deployment correctly uses the app service URL.

1. In the Azure portal, open your app registration. Note that the app registration may be in a different account than your app service. Be sure to sign in to the correct account.
1. In the left sidebar, select **Authentication**.
1. On the **Authentication** pane, find the `https://localhost:3000/fallbackauthdialog.html` and change it to use the app service URL you saved previously. For example, `https://contoso.sso.azurewebsites.net/fallbackauthdialog.html`.
1. Save the change.
1. In the left sidebar, select **Expose an API**.
1. Change the **Application ID URI** to use the app service URL you saved previously. For example, `api://contoso-sso.azurewebsites.net/628050c7-8d46-4f8f-a393-ac22eb688477`.
1. Save the changes.

## Build and deploy

Once the files and app registration are updated, you can deploy the add-in.

1. In VS Code open the terminal and run the command `npm run build`. This will build a folder named `dist` that you can deploy.
1. In the VS Code **Explorer** browse to the `dist` folder. Right-click the `dist` folder and select **Deploy to Web App..**.
1. When prompted to select a resource, select the app service you created previously.
1. When prompted if you are sure, select **Deploy**.
1. When prompted to always deploy the workspace, choose **Yes**.

## Guidelines for any project

Replace all localhost reference in your project with references to the app service URL.
Use Azure Key Vault for secrets.
Remove any port literal numbers, such as `3000`. Use the `process.env.PORT` environment variable that Azure app service provides to your add-in.
Replace any URL or domain references such as `https://localhost:3000` with the URL of your app service.
Use multiple app registrations and manifests for localhost, staging, and deployment testing.


For additional help see Azure help.

## Test the deployment

To check that the deployment is working as expected, open a browser and go to the URL for your app service. It should return the taskpane.html page (although it will not function without Office.)

You can also sideload the manifest-deployment.xml and test the functionality of the add-in in Office. For more information, see [Sideload an Office Add-in for testing](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

For additional support on Azure, see [Azure App Service FAQ](/troubleshoot/azure/app-service/create-delete-resources-faq#contact-us-for-help),
