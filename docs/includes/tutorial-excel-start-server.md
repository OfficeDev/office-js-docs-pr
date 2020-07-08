If the local web server is already running and your add-in is already loaded in Excel, proceed to step 2. Otherwise, start the local web server and sideload your add-in: 

- To test your add-in in Excel, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.

    ```command&nbsp;line
    npm start
    ```

- To test your add-in in Excel on the web, run the following command in the root directory of your project. When you run this command, the local web server will start (if it's not already running).

    ```command&nbsp;line
    npm run start:web
    ```

    To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office for the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-for-the-web).
