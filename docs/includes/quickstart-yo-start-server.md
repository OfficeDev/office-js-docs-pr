1. Open a bash terminal in the root of the project (**[...]/My Office Add-in**) and run the following command to start the dev server.

    ```bash
    npm start
    ```

    This will start a web server at `https://localhost:3000` and open your default browser to that address.

2. Office Web Add-ins should use HTTPS, not HTTP, even when you are developing. If your browser indicates that the site's certificate is not trusted, you will need to add the certificate as a trusted certificate. See [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for details.

    > [!NOTE]
    > Chrome (web browser) may continue to indicate the the site's certificate is not trusted, even after you have completed the process described in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). You can disregard this warning in Chrome and can verify that the certificate is trusted by navigating to `https://localhost:3000` in either Internet Explorer or Microsoft Edge. 

3. After your browser loads the add-in page without any certificate errors, you're ready test your add-in. 
