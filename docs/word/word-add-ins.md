# Build your first Word add-in

_Applies to: Word 2016, Word for iPad, Word for Mac_

The Word JavaScript API is a part of the Office add-in programming model for extending Office applications. The add-in programming model uses web applications to host your extension to Word. You can now extend Word with any web platform or language that you prefer.

A Word add-in runs inside Word and can interact with the contents of the document using the Word JavaScript API that is available in Word 2016. Under the hood, there are two parts to create an add-in: 1) a web application that you can host anywhere, and 2) the [add-in manifest](../../docs/overview/add-in-manifests.md) that Word uses to discover where your web application is hosted (the manifest provides more than this, read more in the [programming overview](word-add-ins-programming-overview.md)).

>**Word add-in = manifest.xml + web app**

### Set it up
You will create a simple web app and the app manifest in this section. The web app will allow you to add boilerplate text into the Word document.

1- Create a folder on your local drive named BoilerplateAddin (for example C:\\BoilerplateAddin). Save all files created in the following steps to this folder.

2- Create a file named home.html for the add-in view. The add-in will have three buttons that, when they're selected, will add boilerplate text. Paste the following code into home.html.

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

3- Create a file named home.js and paste the following code into the file. This contains initialization code and all of our add-in code for making changes to the Word document. This code inserts text based on the cursor or the selection in the Word document.

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

4- Create an XML file named BoilerplateManifest.xml and paste the following code into the file. This is the manifest file that Word uses to discover information about an add-in such as its location or display name.
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

5- Generate a GUID and replace the value in the <code>OfficeApp/Id</code> element with your GUID.

6- Save all the files. You’ve now written your first Word add-in.

7- (Windows only) Create a network folder (for example, \\\MyShare\boilerplate) or [share a folder to the network](https://technet.microsoft.com/en-us/library/cc770880.aspx), and copy home.js, home.html, and BoilerplateManifest.xml to that location.

8- Edit the <code>SourceLocation</code> element in BoilerplateManifest.xml so that it points to the location of home.html.

At this point, you have your first add-in deployed. Now you need to let Word know where to find the add-in.

#### Try this out in Word 2016 for Windows

1. Launch Word and open a document.
2. Choose the **File** tab, and then choose **Options**.
3. Choose **Trust Center**, and then choose the **Trust Center Settings** button.
4. Choose **Trusted Add-ins Catalogs**.
5. In the **Catalog Url** box, enter the path to the folder share that contains BoilerplateManifest.xml and then choose **Add Catalog**.
6. Select the **Show in Menu** check box, and then choose **OK**.
7. A message is displayed to inform you that your settings will be applied the next time you start Office. Close and restart Word.

Now you can run the add-in you created. Follow these steps to see it in action:

1. Open a Word document.
2. On the **Insert** tab in Word 2016, choose **My Add-ins**.
3. Select the **Shared folder** tab.
4. Choose **Boilerplate content**, and then select **Insert**.
5. The add-in will load in a task pane. See figure 1 to see how it will look when it gets loaded.
6. Select the buttons to have boilerplate text entered into the Word document.


### Try it out in Word 2016 for Mac

Now you can run the add-in you created. Follow these steps to see it in action:

1. Create a folder called “wef” in Users/Library/Containers/com.microsoft.word/Data/Documents/
2. Put the manifest, BoilerplateManifest.xml, in the wef folder (Users/Library/Containers/com.microsoft.word/Data/Documents/wef)
3. Open Word 2016 on the Mac and click on the Insert tab > My Add-ins drop down. You should see the add-in listed in the drop down. Select it and it will load the add-in.

__Figure 1. The Boilerplate content add-in loaded in Word__
![Picture of the Word application with the boilerplate add-in loaded.](../../images/boilerplateAddin.png "A simple Word add-in for entering boilerplate text.")

## Learn more

Learn more about extending Word by reading the [Word add-ins programming guide](word-add-ins-programming-overview.md). Read the [Word add-ins JavaScript reference](../../reference/word/word-add-ins-javascript-reference.md) to learn about the objects that you can access.

## Give us your feedback

Your feedback is important to us.

* Check out the docs and let us know about any questions and issues you find in them by [submitting an issue](https://github.com/OfficeDev/office-js-docs/issues) directly in this repository.
* Let us know about your programming experience, what you would like to see in future versions, code samples, etc. Use [this site](http://officespdev.uservoice.com/) for entering your suggestions and ideas.

## Additional resources

* [Office Add-ins](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Get started with Office Add-ins](https://dev.office.com/getting-started/addins?product=word)
* [Word add-ins on GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Word)
