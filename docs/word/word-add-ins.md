---
title: Build your first Word add-in
description: 
ms.date: 11/20/2017 
---

# Build your first Word add-in

_Applies to: Word 2016, Word for iPad, Word for Mac_

A Word add-in runs inside Word and can interact with the contents of the document using the Word JavaScript API, which is part of the Office Add-ins programming model for extending Office applications. In this add-in programming model, you can use the platform and language of your choice to create the web application that hosts your extension to Word and then use the add-in's [manifest](../overview/add-in-manifests.md) to define its settings and capabilities.

In this article, you'll walk through the process of building a Word add-in by using jQuery and the Word JavaScript API. 

> [!NOTE]
> To develop an add-in for Word 2013, you'll need to use the shared [Office Javascript API](word-add-ins-programming-overview.md#javascript-apis-for-word). To learn more about the platforms and the different APIs that are available, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md). 

## Create the web app 

1. Create a folder on your local drive and name it **BoilerplateAddin**. This is where you'll create the files for your app.

2. In your app folder, create a file named **home.html** to specify the HTML that will be rendered in the add-in's task pane. This add-in will display three buttons and when any of the buttons are chosen, boilerplate text will be added to the document. Add the following code to **home.html** and save the file.

    ```html
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Boilerplate text app</title>
        <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
        <script src="home.js" type="text/javascript"></script>
        </head>
        <body>
            <div>
                <h1>Welcome</h1>
            </div>
            <div>
                <p>This sample shows how to add boilerplate text to a document by using the Word JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                <button id="checkhov">Add quote from Anton Chekhov</button>
                <button id="proverb">Add Chinese proverb</button>
            </div>
            <h3><div id="supportedVersion"/></h3>
        </body>
    </html>
    ```

3. In your app folder, create a file named **home.js** to specify the jQuery script for the add-in. This script contains initialization code as well as the code that makes changes to the Word document, by inserting text into the document when a button is chosen. Add the following code to **home.js** and save the file.

    ```javascript
    (function () {
        "use strict";

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

## Create the manifest file

1. In your app folder, create a file named **BoilerplateManifest.xml** to define the add-in's settings and capabilities. Add the following code to the file. 

    ```xml
    <?xml version="1.0" encoding="UTF-8"?>
        <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xsi:type="TaskPaneApp">
            <Id>2b88100c-656e-4bab-9f1e-f6731d86e464</Id>
            <Version>1.0.0.0</Version>
            <ProviderName>Microsoft</ProviderName>
            <DefaultLocale>en-US</DefaultLocale>
            <DisplayName DefaultValue="Boilerplate content" />
            <Description DefaultValue="Insert boilerplate content into a Word document." />
            <Hosts>
                <Host Name="Document"/>
            </Hosts>
            <DefaultSettings>
                <SourceLocation DefaultValue="\\MyShare\boilerplate\home.html" />
            </DefaultSettings>
            <Permissions>ReadWriteDocument</Permissions>
        </OfficeApp>
    ```

2. Generate a GUID using an online generator of your choice. Then, replace the value of the **Id** element shown in the previous step with that GUID.

3. Save the manifest file.

## Deploy the web app and update the manifest

1. Deploy your web app (i.e., the contents of your app folder) to the web server of your choice.

2. In your local app folder, open the manifest file (**BoilerplateManifest.xml**). Edit the attribute value within the **SourceLocation** element to specify the location of the **home.html** file on the web server and save the file.

## Try it out

1. To sideload the add-in within Word, follow the instructions for the platform you'll use to run your add-in.

    - Windows: [Sideload Office Add-ins for testing on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Word Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. In the right task pane, choose any of the buttons to add boilerplate text to the document.

![Picture of the Word application with the boilerplate add-in loaded.](../images/boilerplate-add-in.png)

## Next steps

Congratulations, you've successfully created a Word add-in using jQuery! Next, learn more about the [core concepts](word-add-ins-programming-overview.md) of building Word add-ins.

## Additional resources

* [Word add-ins overview](word-add-ins-programming-overview.md)
* [Word add-in code samples](http://dev.office.com/code-samples#?filters=word,office%20add-ins)
* [Word JavaScript API reference](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)
