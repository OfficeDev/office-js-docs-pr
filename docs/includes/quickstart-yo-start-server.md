
Start the local web server, which runs in Node.js.

    - To test your add-in in Excel for Windows, run the following command to start the local web server, launch Excel, and sideload the add-in:

        ```
        npm start
        ```

        > [!NOTE]
        > Office Web Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run `npm start`, accept the prompt to install the certificate that Yo Office provides. 

        When you run this command, the local web server will start and Excel will start with your add-in loaded.

    - To test your add-in in Excel Online, run the following command to start the local web server:

        ```
        npm run-script start:web
        ```

        > [!NOTE]
        > Office Web Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run `npm start`, accept the prompt to install the certificate that Yo Office provides. 

        When you run this command, the local web server will start. To use your add-in, open a new workbook in Excel Online and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online).

