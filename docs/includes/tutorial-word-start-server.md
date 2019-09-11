Complete the following steps to start the local web server and sideload your add-in.

- To test your add-in in Word, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens Word with your add-in loaded.

    ```command&nbsp;line
    npm start
    ```

- To test your add-in in Word on a browser, run the following command in the root directory of your project. When you run this command, the local web server will start (if it's not already running).

    ```command&nbsp;line
    npm run start:web
    ```

    To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).
