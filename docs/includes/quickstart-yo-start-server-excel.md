
Complete the following steps to start the local web server and sideload your add-in.

> [!NOTE]
> Office Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.

> [!TIP]
> If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server will start.
>
> ```command&nbsp;line
> npm run dev-server
> ```

- To test your add-in in Excel, run the following command in the root directory of your project. When you run this command, the local web server will start (if it's not already running) and Excel will open with your add-in loaded.

    ```command&nbsp;line
    npm start
    ```

- To test your add-in in Excel on a browser, run the following command in the root directory of your project. When you run this command, the local web server will start (if it's not already running).

    ```command&nbsp;line
    npm run start:web
    ```

    To use your add-in, open a new workbook in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).

