
Start the local web server, which runs in Node.js.

> [!NOTE]
> Office Web Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run one of the following commands to start the web server, accept the prompt to install the certificate that the Yeoman generator provides. 

- To test your add-in in Excel for Windows, run the following command to start the local web server, launch Excel, and sideload the add-in:

    ```
    npm start
    ```

    When you run this command, the local web server will start and Excel will open with your add-in loaded.

- To test your add-in in Excel Online, run the following command to start the local web server:

    ```
    npm run-script start:web
    ```

    When you run this command, the local web server will start. To use your add-in, open a new workbook in Excel Online and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online).

