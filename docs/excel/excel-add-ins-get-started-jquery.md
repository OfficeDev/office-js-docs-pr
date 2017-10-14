# Build an Excel add-in using jQuery

In this article, you'll walk through the process of building an Excel add-in by using jQuery and the Excel JavaScript API.

## Create the web app

1. Create a folder on your local drive and name it **my-addin**. This is where you'll create the files for your app.

2. Navigate to your app folder.

    ```bash
    cd my-addin
    ```

3. In your app folder, create a file named **Home.html** to specify the HTML that will be rendered in the add-in's task pane. Add the following code and save the file.

    ```html
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
            <title>My Excel Add-in</title>
            <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>

            <link href="Office.css" rel="stylesheet" type="text/css" />
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>

            <link href="Common.css" rel="stylesheet" type="text/css" />
            <script src="Notification.js" type="text/javascript"></script>
            <script src="Home.js" type="text/javascript"></script>

            <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
            <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">
        </head>
        <body class="ms-font-m">
            <div id="content-header">
                <div class="padding">
                    <h1>Welcome</h1>
                </div>
            </div>
            <div id="content-main">
                <div class="padding">
                    <p>Choose the button below to set the color of the selected range to green.</p>
                    <br />
                    <h3>Try it out</h3>
                    <button class="ms-Button" id="set-color">Set color</button>
                </div>
            </div>
        </body>
    </html>
    ```

4. In your app folder, create a file named **Home.js** to specify the jQuery script for the add-in. Add the following code and save the file.

    ```js
    (function () {
        "use strict";

        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#set-color').click(setColor);
            });
        };

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return ctx.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

5. In your app folder, create a file named **Common.css** to specify the custom styles for the add-in. Add the following code and save the file.

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

## Create the manifest file and sideload the add-in

1. In your app folder, create a file named **my-excel-add-in-manifest.xml** to define the add-in's settings and capabilities. Add the following XML to the file.

    ```xml
    <?xml version="1.0" encoding="UTF-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
        <Id>ab2991e7-fe64-465b-a2f1-c865247ef434</Id>
        <Version>1.0.0.0</Version>
        <ProviderName>Microsoft</ProviderName>
        <DefaultLocale>en-US</DefaultLocale>
        <DisplayName DefaultValue="My Office Add-in" />
        <Description DefaultValue="A task pane add-in for Excel built using jQuery"/>
        <Capabilities>
        <Capability Name="Workbook" />
        </Capabilities>
        <DefaultSettings>
        <SourceLocation DefaultValue="~remoteAppUrl/my-addin/Home.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

2. Generate a GUID using an online generator of your choice. Then, replace the value of the **Id** element shown in the previous step with that GUID.

3. Save the manifest file. 

## Deploy the web app and update the manifest

1. Deploy your web app (i.e., the contents of your app folder) to the web server of your choice.

2. In your local app folder, open the manifest file (**my-excel-add-in-manifest.xml**). Edit the attribute value within the **SourceLocation** element to specify the location of the **Home.html** file on the web server and save the file.

## Try it out

1. Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Excel.

    - Windows: [Sideload Office Add-ins for testing on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. In the right task pane, choose the **Set color** button to set the color of the selected range to green.

    ![Excel Add-in](../../images/excel_quickstart_addin_1.png)

## Next steps

Congratulations, you've successfully created an Excel add-in using jQuery! Next, learn more about the [core concepts](excel-add-ins-core-concepts.md) of building Excel add-ins.

## Additional resources

* [Excel JavaScript API core concepts](excel-add-ins-core-concepts.md)
* [Explore snippets with Script Lab](https://store.office.com/en-001/app.aspx?assetid=WA104380862&ui=en-US&rs=en-001&ad=US&appredirect=false)
* [Excel add-in code samples](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Excel JavaScript API reference](../../reference/excel/excel-add-ins-reference-overview.md)
