---
title: Deploy an Office Add-in that uses single sign-on (SSO) to Microsoft Azure App Service | Microsoft Docs
description: Learn how to deploy an Office Add-in that uses single sign-on (SSO) to Microsoft Azure App Service from Visual Studio Code.
ms.date: 01/12/2020
ms.localizationpriority: medium
---

# Deploy an Office Add-in that uses single sign-on (SSO) to Microsoft Azure App Service

Office Add-ins that use SSO require a web service that supports running the REST APIs and server-side code in the project. You can't deploy to a static website. Follow the steps in this article to deploy your Office Add-in to Azure App Service for staging or deployment.

## Prerequisites

The steps in this article work for an Office Add-in created by the [Yeoman Generator for Office Add-ins](https://github.com/OfficeDev/generator-office) using the `Office Add-in Task Pane project supporting single sign-on (localhost)` project type. Be sure you have configured the add-in project so that it runs on localhost successfully. For more information, see the [Single sign-on (SSO) quick start](../quickstarts/sso-quickstart.md).

The steps in this article also require:

- An Azure account. Get a trial subscription at [Microsoft Azure](https://azure.microsoft.com/free/).
- [Azure Account extension](https://marketplace.visualstudio.com/items?itemName=ms-vscode.azure-account) for VS Code.
- [Azure App Service extension](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azureappservice) for VS Code.

## Create the App Service app

The following steps set up a basic deployment of the Office Add-in. There are multiple ways to configure deployment that are not covered in this documentation. For additional options on how you may want to configure your deployment, see [Deployment Best Practices](/azure/app-service/deploy-best-practices)

### Sign in to Azure

1. Open your Office Add-in project in VS Code.
1. Select the Azure logo in the [Activity Bar](https://code.visualstudio.com/docs/getstarted/userinterface). If the Activity Bar is hidden, open it by selecting **View** > **Appearance** > **Activity Bar**.
1. In the **App Service** explorer, select **Sign in to Azure...** and follow the instructions.

    :::image type="content" source="../images/azure-extension-sign-in.png" alt-text="Sign in to Azure button selected in the Azure extension.":::

### Configure the App Service app

App Service supports various versions of Node.js on both Linux and Windows. Select the tab for the one you'd like to use and then follow the instructions to create your App Service app.

# [Deploy to Linux](#tab/linux)

1. Right-click on App Services and select **Create new Web App**. A Linux container is used by default.
1. Type a globally unique name for your web app and press **Enter**. The name must be unique across all of Azure and use only alphanumeric characters ('A-Z', 'a-z', and '0-9') and hyphens ('-').
1. In Select a runtime stack, select the **Node 16 LTS** runtime stack.
1. In Select a pricing tier, select **Free (F1)** and wait for the resources to be provisioned in Azure. When prompted to deploy, don't deploy the add-in yet. You'll do that in a later step.

# [Deploy to Windows](#tab/windows)

1. Right-click on App Services and select **Create new Web App... Advanced**.
1. Type a globally unique name for your web app and press **Enter**. The name must be unique across all of Azure and use only alphanumeric characters ('A-Z', 'a-z', and '0-9') and hyphens ('-').
1. Select the resource group you want to use. If you don't have a resource group, select **Create a new resource group**, then enter a name for the resource group, such as *AppServiceQS-rg*.
1. Select the **Node 16 LTS** runtime stack.
1. Select **Windows** for the operating system.
1. Select the location you want to serve your app from. For example, *West Europe*.
1. Select the App Service plan you want to use. If you don't have an App Service plan, select **Create new App Service plan**, then enter a name for the plan (such as *AppServiceQS-plan*), then select **F1 Free**.
1. For **Select an Application Insights resource for your app**, select **Skip for now** and wait the resources to be provisioned in Azure. When prompted to deploy, don't deploy the add-in yet. You'll do that in a later step.
1. In the **App Service** explorer in Visual Studio code, expand the node for the new app, right-click **Application Settings**, and select **Add New Setting**:

    :::image type="content" source="../images/azure-app-service-add-setting.png" alt-text="Add app setting command.":::

1. Enter `SCM_DO_BUILD_DURING_DEPLOYMENT` for the setting key.
1. Enter `true` for the setting value.

    This app setting enables build automation at deploy time, which automatically detects the start script and generates the *web.config* with it.

-----

1. Right-click your App Service app and select **Open in Portal**.
1. When the portal opens in the browser, copy the domain name of the **URL** (not the `https://` part) from the **Overview** pane and save it. You'll need it in later steps.

## Update package.json

1. Open the package.json file. Then replace the `start` command in the `"scripts"` section with the following entry.

    ```json
    "start": "node middletier.js",
    ```

1. Find the `"prestart"` entry in the '"scripts"' section and delete it. This section is not needed for this deployment.
1. Save the file.

## Update webpack.config.js

1. Open the **webpack.config.js** file.
1. Set the `urlDev` and `urlProd` constants to the following values (without the `https` protocol portion). This will cause webpack to replace `localhost:3000` with your web site domain name in the `/dest/manifest.xml` file.

    ```javascript
    const urlDev = "localhost:3000";
    const urlProd = "<your-web-site-domain-name>";   
    ```

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

## Update manifest

1. Open the **manifest-deployment.xml** file.
1. Replace `<SupportUrl DefaultValue="https://www.contoso.com/help"/>` with the URL of your web site help page.
1. Replace `<AppDomain>https://www.contoso.com</AppDomain>` with the URL of your web site.
1. In the `<Scopes>` section near the bottom of the file, add the `openid` scope as shown in the following XML.

    ```xml
    <Scopes>
        <Scope>User.Read</Scope>
        <Scope>profile</Scope>
        <Scope>openid</Scope>
    </Scopes>
    ```

1. Save the file.

## Update fallbackauthdialog.js (or fallbackauthdialog.ts)

1. Open the **src/helpers/fallbackauthdialog.js** file, or **src/helpers/fallbackauthdialog.ts** if your project uses TypeScript.
1. Find the `redirectUri` on line 24 and change the value to use your App Service app URL you saved previously. For example, `redirectUri: "https://contoso-sso.azurewebsites.net/fallbackauthdialog.html",`
1. Save the file.

## Update .ENV

The **.ENV** file contains a client secret. For the purposes of learning in this article you can deploy the **.ENV** file to Azure. However for a production deployment, you should move the secret and any other confidential data into [Azure Key Vault](/azure/key-vault/general/basic-concepts).

1. Open the **.ENV** file.
1. Set the `NODE_ENV` variable to the value `production`.
1. Save the file.

## Update app.js (or app.ts)

The app.js (or app.ts) requires several minor changes to run correctly in a deployment. It's easiest to just replace the file with an updated version for deployment.

1. Open the **src/middle-tier/app.js** file, or **src/middle-tier/app.ts** if your project uses TypeScript.
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
      res.sendFile("/taskpane.html", { root: __dirname });
    });
    
    // Route APIs
    indexRouter.get("/getuserdata", validateJwt, getUserData);
    
    app.use("/", indexRouter);
    
    // Catch 404 and forward to error handler
    app.use(function (req, res, next) {
      console.log("error 404");
      next(createError(404));
    });
    
    // error handler
    app.use(function (err, req, temp, res) {
      // set locals, only providing error in development
      console.log("error 500");
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

We recommend you create multiple app registrations for localhost, staging, and deployment testing. The following steps ensure that the app registration you use for deployment correctly uses the App Service app URL.

1. In the Azure portal, open your app registration. Note that the app registration may be in a different account than your App Service app. Be sure to sign in to the correct account.
1. In the left sidebar, select **Authentication**.

    :::image type="content" source="../images/azure-portal-authentication-page.png" alt-text="The authentication page in the Azure app registration.":::

1. On the **Authentication** pane, find the `https://localhost:3000/fallbackauthdialog.html` and change it to use the App Service app URL you saved previously. For example, `https://contoso.sso.azurewebsites.net/fallbackauthdialog.html`.
1. Save the change.
1. In the left sidebar, select **Expose an API**.
1. Edit the **Application ID URI** field and replace `localhost:3000` with the domain from the App Service app URL you saved previously.
1. Save the changes.

## Build and deploy

Once the files and app registration are updated, you can deploy the add-in.

1. In VS Code open the terminal and run the command `npm run build`. This will build a folder named `dist` that you can deploy.
1. In the VS Code **Explorer** browse to the `dist` folder. Right-click the `dist` folder and select **Deploy to Web App..**.
1. When prompted to select a resource, select the App Service app you created previously.
1. When prompted if you are sure, select **Deploy**.
1. When prompted to always deploy the workspace, choose **Yes**.

If you make additional code changes, you'll need to run `npm run build` again and redeploy the project.

## Test the deployment

Sideload the **manifest-deployment.xml** and test the functionality of the add-in in Office. For more information, see [Sideload an Office Add-in for testing](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

If you encounter any deployment issues, see the [Azure App Service troubleshooting documentation](/troubleshoot/azure/app-service/welcome-app-service).

## Next steps

- [Deploy to App Service using GitHub Actions](/azure/app-service/deploy-github-actions?tabs=applevel)
- [Deployment Best Practices](/azure/app-service/deploy-best-practices)
- [App Service documentation](/azure/app-service)
- [Azure community support](/answers/products/azure?product=all)
- [Create a Node.js web app in Azure](/azure/app-service/quickstart-nodejs?tabs=windows&pivots=development-environment-vscode)
