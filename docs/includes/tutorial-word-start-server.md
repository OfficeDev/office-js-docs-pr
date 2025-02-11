If the local web server is already running and your add-in is already loaded in Word, proceed to step 2. Otherwise, start the local web server and sideload your add-in.

- To test your add-in in Word, run the following command in the root directory of your project. This starts the local web server (if it isn't already running) and opens Word with your add-in loaded.

    ```command&nbsp;line
    npm start
    ```

- To test your add-in in Word on the web, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a Word document on your OneDrive or a SharePoint library to which you have permissions.

    [!INCLUDE [npm start on web command syntax](../includes/start-web-sideload-instructions.md)]

