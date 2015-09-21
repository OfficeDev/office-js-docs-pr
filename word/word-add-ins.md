# Word 2016 add-ins

*Applies to: Word 2016*

Welcome to the Word add-in JavaScript API documentation. The Word JavaScript API is a part of the Office add-in programming model for extending Microsoft Office applications. The add-in programming model uses web applications to host your extension to Word. You can now extend Word with any web platform or language that you prefer. 

## API Overview

Before we start going into the specifics of the Javascript API for Word, it is important to know that this new Word add-in object model is different than what was made available with Word in Office 2013. The previous object model was not typed and provided a generic API for extending Office clients. While the previous model is still applicable to Word 2016, we strongly suggest that you start using the new Word object model. We suggest that you read the [platform overview](https://msdn.microsoft.com/EN-US/library/office/jj220082.aspx) if you aren't familiar with the add-in platform. 

The new JavaScript APIs for Word changes the way that you can interact with objects like documents and paragraphs. Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the new APIs provide “proxy” JavaScript objects that correspond to the real objects running in Word.  You can directly interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them.  These interactions with proxy objects aren’t immediately realized in the running script, though, so we provide a method on the context called **sync()** that synchronizes the state between your running JavaScript and the real objects in Office by executing instructions queued in your script and retrieving properties of loaded Word objects for use in your script.  

<!--Let's take a look at this code and comments to get a better understanding of how the proxy objects are used to interact with the contents of a Word document.

```javascript
    
    // Run a batch operation against the Word object mode
    Word.run(function (ctx) {

        // Create a proxy object for the document. Nothing has changed in the Word document.
        var thisDocument = ctx.document;

        // Queue a command to load the save state on the document proxy object. Nothing has changed in the Word document.
        // The current value for the saved property on the document proxy object is null.
        ctx.load(thisDocument, { select: 'saved'});    

        // Synchronize the document state by executing the queued-up commands, 
        // and return a promise to indicate task completion.
        // If all of the commands are successful, the Word document will be saved and the
        // value of the *saved* property will be returned and set on the document proxy object. The document
        // proxy object's, and the actual Word document's, *saved* property will be in sync. 
        return ctx.sync().then(function () {

            if (thisDocument.saved === false) {
                // Queue a command to save this document.
                thisDocument.save();

                // Synchronize the document state by executing the queued-up commands, 
                // and return a promise to indicate task completion.
                return ctx.sync().then(function () {
                    console.log('Saved the document');
                });
            } else {
                console.log('The document has not changed since the last save.');
            }
        });  
    })
    .catch(function (error) {
        console.log(JSON.stringify(error));
    });    
    
```
-->

## Create your first Word add-in

A Word add-in runs inside Word and can interact with the contents of the document using the new Word JavaScript APIs available in Office 2016. Under the hood, there are two parts to create an add-in: 1) a web application that you can host anywhere, and 2) the add-in manifest that Word uses to discover where your web application is hosted (the manifest provides more than this, read more in the [programming guide](word-add-ins-programming-guide.md)).

>**Word add-in = manifest.xml + web app**

### Set it up
You will create a simple web app and the app manifiest in this section. The web app will allow you to add boilerplate text into the word document. 

1. Create a folder on your local drive called BoilerplateAddin (for example C:\BoilerplateAddin). Save all of the files created in the following steps into this folder.

2. Create a file named home.html and paste in the code below. <!-- TODO: Update CDN when it becomes official. -->

    ```html
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Boilerplate text app</title>    
        <script src="http://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
        <script src="https://eus-appsforoffice.edog.officeapps.live.com/afo/lib/1.1/hosted/office.js" type="text/javascript"></script>
        <link href="Styles.css" rel="stylesheet" type="text/css" />
        <script src="Home.js" type="text/javascript"></script>
        </head>
        <body>
            <div id="content-header">
                    <h1>Welcome</h1>
            </div>
            <div id="content-main">
                    <p>This sample shows how to add boilerplate text in to a document by using the Word JavaScript API.</p>
                    <br />
                    <h3>Try it out</h3>
                    <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                    <button id="checkhov">Add quote from Anton Chekhov</button>
                    <button id="proverb">Add Chinese proverb</button>
            </div>
        </body>
    </html>
    ```

3. Create a file named home.js and paste in the code below

    ```javascript
    (function () {
        "use strict";

        // The initialize function is run each time a the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#emerson').click(insertEmersonQuoteAtSelection);
                $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                $('#proverb').click(insertChineseProverbAtTheEnd);
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

                // Synchronize the document state by executing the queued-up commands, 
                // and return a promise to indicate task completion.
                return ctx.sync().then(function () {
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

                // Synchronize the document state by executing the queued-up commands, 
                // and return a promise to indicate task completion.
                return ctx.sync().then(function () {
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

                // Synchronize the document state by executing the queued-up commands, 
                // and return a promise to indicate task completion.
                return ctx.sync().then(function () {
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

4. Create an XML file named BoilerplateManifext.xml and paste in the code below.

    ```xml
    <?xml version="1.0" encoding="UTF-8"?>
        <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
            <Id>2b88100c-656e-4bab-9f1e-f6731d86e464</Id>
            <Version>1.0.0.0</Version>
            <ProviderName>Microsoft</ProviderName>
            <DefaultLocale>en-US</DefaultLocale>
            <DisplayName DefaultValue="Boilerplate content" />
            <Description DefaultValue="Insert boilerplate content into a Word document." />
            <Capabilities>
                <Capability Name="document" />
            </Capabilities>
            <DefaultSettings>
                <SourceLocation DefaultValue="\\tk2offfsm03\Users\mmainer\boilerplate\home.html" />
            </DefaultSettings>
            <Permissions>ReadWriteDocument</Permissions>
        </OfficeApp>
     ```

5. Generate a GUID and replace the value in the **OfficeApp/Id** element with your GUID.

6. Save all the files. You’ve now written your first Word add-in. 

7. Create a shared folder (for example, \\MyShare\boilerplate) and copy home.js, home.html, and BoilerplateManifext.xml to that location. 

8. Edit the <SourceLocation> element in BoilerplateManifext.xml so that it points to the shared folder location for home.html. 

At this point, you have your first add-in deployed. Now you need to let Word know where to find the add-in.

1. Launch Word and open a document.
2. Choose the File tab, and then choose Options.
3. Choose Trust Center, and then choose the Trust Center Settings button.
4. Choose Trusted Add-ins Catalogs.
5. In the Catalog Url box, enter the path to the folder share that contains BoilerplateManifext.xml and then choose Add Catalog.
6. Select the Show in Menu check box, and then choose OK.
7. A message is displayed to inform you that your settings will be applied the next time you start Office. Close and restart Word. 

### Try it out

Now you can run the add-in you created. Follow these steps to see it in action.

1. Open a Word document. 
2. On the Insert tab in Word 2016, choose My Add-ins. 
3. Choose *Boilerplate content*, and then select Insert.
4. The add-in will load in a task pane. 
5. Select the buttons to have boilerplate text entered into the Word document.

## Learn information

Learn more about extending Word by reading the [Word add-ins programming guide](word-add-ins-programming-guide.md). Peruse the [Word add-ins JavaScript reference](word-add-ins-javascript-reference.md) to learn about the objects you can access. Check out our curated list of [Word add-in code samples](word-add-ins-code-samples.md) to get some ideas on how  you can create Word add-ins.

## Give feedback on the API

The documentation for this API is hosted on GitHub with the intention that we can improve the documentation and API by making it open for [issues](https://github.com/OfficeDev/office-js-docs/issues) against the documentation. Issues can include errors in the documentation, requests for clarification, or requests for improvements in the documentation. We also welcome general feedback about the API and the experience you have with it.

## Additional links

* [Office Add-ins](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Get started with Office Add-ins](http://dev.office.com/getting-started/addins)
* [Word add-ins on GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Word)