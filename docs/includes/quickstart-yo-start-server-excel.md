
Complete the following steps to start the local web server and sideload your add-in.

[!INCLUDE [alert use https](alert-use-https.md)]

> [!TIP]
> If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.
>
> ```command&nbsp;line
> npm run dev-server
> ```

- To test your add-in in Excel, run the following command in the root directory of your project. This starts the local web server and opens Excel with your add-in loaded.

    ```command&nbsp;line
    npm start
    ```

- To test your add-in in Excel on a browser, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of an Excel document on your OneDrive or a SharePoint library to which you have permissions.

    [!INCLUDE [npm start on web command syntax](../includes/start-web-sideload-instructions.md)]
