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

Create the following resources (or be sure they exist) in your Azure account first:

- An Azure app plan that uses the Windows OS. You can use any pricing tier. We recommend the free tier for just learning the basics of deploying an Office Add-in using SSO.
- An Azure resource group.

Once complete, you'll see an Azure icon in the Activity Bar
If your activity bar is hidden, you won't be able to access the extension. Show the Activity Bar by clicking View > Appearance > Show Activity Bar

1. Open your Office Add-in project in VS Code.
1. Select the Azure icon in the activity bar.

    > [!NOTE]
    > If your activity bar is hidden, you won't be able to access the extension. Show the Activity Bar by selecting **View** > **Appearance** > **Show Activity Bar**

1. If you are not already signed in, select **Sign in to Azure...** to sign in to your Azure Account.
1. Once you are signed in to your Azure account and you have your app open in Visual Studio Code, select the **Create Resource** button - it's the plus icon - to create a Web App.
1. extension sidebar, then under your Azure subscription, select **App Services**.
1. Right-click **App Services** and select **Create New Web App... (Advanced)**.
1. On step 1 of **Create new web app** enter a globally unique name for your app service; for example, **username-sso-add-in**.
1. On step 2 of **Create new web app** select the resource group you want to use. If you don't have a resource group, see [Manage Azure resource groups by using the Azure portal](/azure/azure-resource-manager/management/manage-resource-groups-portal) for more information on creating a resource group.
1. On step 3 of **Create new web app** select **Node 16 LTS** for the runtime stack.
1. On step 4 of **Create new web app** select **Windows** for the OS.
1. On step 5 of **Create new web app** select a location that is ideally near you.
1. On step 6 of **Create new web app** select your App Service plan configured for **Windows**.
1. On step 7 of **Create new web app** if you are asked for Applicaoitn Insights, you can select **Skip for now**.

Azure will create the app service and it will appear under **App Services** in the sidebar.

1. Right-click your app service and select **Open in Portal**.
1. In the Azure portal, select **Configuration** in the sidebar.
1. In the **Application settings** pane, select **New application setting**. Create a new application setting named **SCM_DO_BUILD_DURING_DEPLOYMENT** with the value **true**. Then select **OK**.
1. In the **Application settings** pane, select **Save** and then **Continue** to save the new application setting.
1. In the sidebar, select **Overview** and then copy the **URL** value and save it. You'll need it in later steps.

## Update manifest

Often it's useful to maintain multiple manifests for localhost testing, staging, and deployment. We recommend you copy the existing file and create a new manifest named **manifest-deployment.xml**.

1. Open the **manifest-deployment.xml** file.
1. Search and replace all instances of `https://localhost:3000` with the URL you saved previously.
1. Find the `<WebApplicationInfo>` section near the bottom of the manifest. Replace the text `localhost:3000` in the `<Resource>` tag with the URL you saved previously. Don't include the `https://` portion of the URL.
1. Add an AppDomain entry for the app service domain. For example `<AppDomain>contoso-sso.azurewebsites.net</AppDomain>`.
1. Save.

## Update package.json

1. From the terminal, run the command `npm install ejs@3.1.8`.
1. Open the **package.json** file.
1. Modify the start to read `"start": "node middletier.js",`. Once deployed the project will use Node JS instead of a webpack server.
1. Save

## Update webpack.config.js

1. Open the **webpack.config.js** file.
1. Update the `CopyWebpackPlugin` section to include the package.json file as shown in the following example.

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

1. Save

## Update fallbackauthdialog.js

1. Open the **src/helpers/fallbackauthdialog.js** file.
1. Change the `redirectUri` on line 24 to reference the URL of the app service, not localhost:3000.
1. Save.

## Update app.js

1. Open the **app.js** file.
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
    app.use(function (err, req, res) {
      // set locals, only providing error in development
      res.locals.message = err.message;
      res.locals.error = req.app.get("env") === "development" ? err : {};
    
      // render the error page
      res.status(err.status || 500);
      //res.render("error");
    });
    
    app.listen(process.env.PORT, () => console.log("Server listening on port: " + process.env.PORT));
    ```

1. Save.

## Update app registration

You can create multiple app registrations which is useful to have one for testing, staging, deploying, and so on.

1. Change the redirect URI to use the new URL.
1. Go to expose API and update the Application ID URI.

## Build and deploy

run build
deploy
use streams for debug
