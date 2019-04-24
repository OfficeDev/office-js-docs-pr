
Start the local web server and sideload your add-in.

> [!NOTE]
> Office Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides. 

- To test your add-in in Excel, run the following command. When you run this command, the local web server will start and Excel will open with your add-in loaded.

    ```
    npm start
    ```

- To test your add-in in Excel Online, run the following command. When you run this command, the local web server will start.

    ```
    npm run start:web
    ```

    To use your add-in, open a new workbook in Excel Online and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online).

