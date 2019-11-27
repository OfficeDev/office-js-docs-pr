An Office Add-in consists of a web application and a manifest file. The web application defines the add-in's user interface and functionality, while the manifest specifies the location of the web application and defines settings and capabilities of the add-in. 

While you're developing your add-in, you can run the add-in on your local web server (`localhost`), but when you're ready to publish it for other users to access, you'll need to deploy the web application to a web server or web hosting service (for example, Microsoft Azure) and update the manifest to specify the URL of the deployed application. 

When your add-in is working as desired and you're ready to publish it for other users to access, complete the following steps:

1. From the command line, in the root directory of your add-in project, run the following command to prepare all files for production deployment: 

    ```command&nbsp;line
    npm run build
    ```

    When the build completes, the **dist** folder in the root directory of your add-in project will contain the files that you'll deploy in subsequent steps.

2. Upload the contents of the **dist** folder to the web server that'll host your add-in. You can use any type of web server or web hosting service to host your add-in.

3. In VS Code, open the add-in's manifest file, located in the root directory of the project (`manifest.xml`). Replace all occurrences of `https://localhost:3000` with the URL of the web application that you deployed to a web server in the previous step.

    > [!TIP]
    > While you can update the existing `manifest.xml` file as described here, you might instead consider preserving the existing file in its original state, and creating a copy of the file where you'll replace all instances of `https://localhost:3000` with the deployed web application's URL. Doing things this way, you'd have two versions of the manifest file: one that could be used during ongoing development and testing of your add-in (referencing `localhost`) and another that could be used to publish your add-in for other users to access (referencing the deployed web application's URL). The two manifest files can be named however you choose (for example, you might choose a naming convention like `manifest_DEV.xml` and `manifest_PROD.xml`).

4. Choose the method you'd like to use to [deploy and publish your Office Add-in](../publish/publish.md) your add-in, and follow the instructions to publish the manifest file.
