1. Open a bash terminal in the root of the project (**[...]/My Office Add-in**) and run the following command to start the dev server.

    ```command&nbsp;line
    npm start
    ``` 

2. Open either Internet Explorer or Microsoft Edge and navigate to `https://localhost:3000`. If the page loads without any certificate errors, proceed to the next section in this article (**Try it out**). If your browser indicates that the site's certificate is not trusted, proceed to the following step.

3. Office Web Add-ins should use HTTPS, not HTTP, even when you are developing. If your browser indicates that the site's certificate is not trusted, you will need to add the certificate as a trusted certificate. See [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for details.

    > [!NOTE]
    > Chrome (web browser) may continue to indicate the the site's certificate is not trusted, even after you have completed the process described in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). Therefore, you should use either Internet Explorer or Microsoft Edge to verify that the certificate is trusted. 

4. After your browser loads the add-in page without any certificate errors, you're ready test your add-in.
