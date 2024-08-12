---
title: Create a dictionary task pane add-in
description: Learn how to create a dictionary task pane add-in.
ms.date: 06/13/2023
ms.topic: how-to
ms.localizationpriority: medium
---

# Create a dictionary task pane add-in

This article shows you an example of a task pane add-in with an accompanying web service that provides dictionary definitions or thesaurus synonyms for the user's current selection in a Word document.

A dictionary Office Add-in is based on the standard task pane add-in with additional features to support querying and displaying definitions from a dictionary XML web service in additional places in the Office application's UI.

In a typical dictionary task pane add-in, a user selects a word or phrase in their document, and the JavaScript logic behind the add-in passes this selection to the dictionary provider's XML web service. The dictionary provider's webpage then updates to show the definitions for the selection to the user.

The XML web service component returns up to three definitions in the format defined by the example OfficeDefinitions XML schema, which are then displayed to the user in other places in the hosting Office application's UI.

Figure 1 shows the selection and display experience for a Bing-branded dictionary add-in that's running in Word.

*Figure 1. Dictionary add-in displaying definitions for the selected word*

![Dictionary app displaying a definition.](../images/dictionary-add-in-01.png)

It's up to you to determine if selecting the **See More** link in the dictionary add-in's HTML UI displays more information within the task pane or opens a separate window to the full webpage for the selected word or phrase.

Figure 2 shows the **Define** command in the context menu that enables users to quickly launch installed dictionaries. Figures 3 through 5 show the places in the Office UI where the dictionary XML services are used to provide definitions in Word.

*Figure 2. Define command in the context menu*

:::image type="content" source="../images/dictionary-agave-02.jpg" alt-text="Define context menu.":::

*Figure 3. Definitions in the Spelling and Grammar panes*

![Definitions in the Spelling and Grammar panes.](../images/dictionary-agave-03.jpg)

*Figure 4. Definitions in the Thesaurus pane*

![Definitions in the Thesaurus pane.](../images/dictionary-agave-04.jpg)

*Figure 5. Definitions in Reading Mode*

:::image type="content" source="../images/dictionary-agave-05.jpg" alt-text="Definitions in Reading Mode.":::

To create a task pane add-in that provides a dictionary lookup, create two main components.

- An XML web service that looks up definitions from a dictionary service, and then returns those values in an XML format that can be consumed and displayed by the dictionary add-in.
- A task pane add-in that submits the user's current selection to the dictionary web service, displays definitions, and can optionally insert those values into the document.

The following sections provide examples of how to create these components.

## Prerequisites

[!include[Visual Studio project prerequisites](../includes/quickstart-vs-prerequisites.md)]

Next, create a Word add-in project in Visual Studio.

[!include[Visual Studio instructions to create Word solution](../includes/vs-word-instructions.md)]

To learn more about the projects in a Word add-in solution, see the [quick start](/office/dev/add-ins/quickstarts/word-quickstart?tabs=visualstudio#explore-the-visual-studio-solution).

## Create a dictionary XML web service

The XML web service must return queries to the web service as XML that conforms to the OfficeDefinitions XML schema. The following two sections describe the OfficeDefinitions XML schema, and provide an example of how to code an XML web service that returns queries in that XML format.

### OfficeDefinitions XML schema

The following code shows sample XSD for the OfficeDefinitions XML schema example.

```XML
<?xml version="1.0" encoding="utf-8"?>
<xs:schema
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xs="https://www.w3.org/2001/XMLSchema"
  targetNamespace="http://schemas.microsoft.com/contoso/OfficeDefinitions"
  xmlns="http://schemas.microsoft.com/contoso/OfficeDefinitions">
  <xs:element name="Result">
    <xs:complexType>
      <xs:sequence>
        <xs:element name="SeeMoreURL" type="xs:anyURI"/>
        <xs:element name="Definitions" type="DefinitionListType"/>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
  <xs:complexType name="DefinitionListType">
    <xs:sequence>
      <xs:element name="Definition" maxOccurs="3">
        <xs:simpleType>
          <xs:restriction base="xs:normalizedString">
            <xs:maxLength value="400"/>
          </xs:restriction>
        </xs:simpleType>
      </xs:element>
    </xs:sequence>
  </xs:complexType>
</xs:schema>
```

Returned XML consists of a root **\<Result\>** element that contains a **\<Definitions\>** element with zero to three **\<Definition\>** child elements. Each child element contains definitions that are at most 400 characters in length. Additionally, the URL to the full page on the dictionary site must be provided in the **\<SeeMoreURL\>** element. The following example shows the structure of returned XML that conforms to the OfficeDefinitions schema.

```XML
<?xml version="1.0" encoding="utf-8"?>
<Result xmlns="http://schemas.microsoft.com/contoso/OfficeDefinitions">
  <SeeMoreURL xmlns="">https://www.bing.com/search?q=example</SeeMoreURL>
  <Definitions xmlns="">
    <Definition>Definition1</Definition>
    <Definition>Definition2</Definition>
    <Definition>Definition3</Definition>
  </Definitions>
 </Result>

```

### Sample dictionary XML web service

The following C# code provides an example of how to write code for an XML web service that returns the result of a dictionary query in the OfficeDefinitions XML format.

```cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Xml;
using System.Text;
using System.IO;
using System.Net;
using System.Web.Script.Services;

/// <summary>
/// Summary description for _Default.
/// </summary>
[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this web service to be called from script, using ASP.NET AJAX, include the following line. 
[ScriptService]
public class WebService : System.Web.Services.WebService {

    public WebService () {

        // Uncomment the following line if using designed components.
        // InitializeComponent(); 
    }

    // You can replace this method entirely with your own method that gets definitions
    // from your data source and then formats it into the example OfficeDefinitions XML format. 
    // If you need a reference for constructing the returned XML, you can use this example as a basis.
    [WebMethod]
    public XmlDocument Define(string word)
    {

        StringBuilder sb = new StringBuilder();
        XmlWriter writer = XmlWriter.Create(sb);
        {
            writer.WriteStartDocument();
            
                writer.WriteStartElement("Result", "http://schemas.microsoft.com/contoso/OfficeDefinitions");

                    // See More URL should be changed to the dictionary publisher's page for that word on
                    // their website.
                    writer.WriteElementString("SeeMoreURL", "https://www.bing.com/search?q=" + word);

                    writer.WriteStartElement("Definitions");
            
                        writer.WriteElementString("Definition", "Definition 1 of " + word);
                        writer.WriteElementString("Definition", "Definition 2 of " + word);
                        writer.WriteElementString("Definition", "Definition 3 of " + word);
                   
                    writer.WriteEndElement(); // End of Definitions element.

                writer.WriteEndElement(); // End of Result element.
            
            writer.WriteEndDocument();
        }
        writer.Close();

        XmlDocument doc = new XmlDocument();
        doc.LoadXml(sb.ToString());

        return doc;
    }
}
```

To get started with development, you can do the following.

#### Create the web service

1. Add a **Web Service (ASMX)** to the add-in's web application project in Visual Studio and name it **DictionaryWebService**.
1. Replace the entire content of the associated .asmx.cs file with the preceding C# code sample.

#### Update the web service markup

1. In the **Solution Explorer**, select the **DictionaryWebService.asmx** file then open its context menu and choose **View Markup**.
1. Replace the contents of DictionaryWebService.asmx with the following code.

    ```XML
    <%@ WebService Language="C#" CodeBehind="DictionaryWebService.asmx.cs" Class="WebService" %>
    ```

#### Update the web.config

1. In the **Web.config** of the add-in's web application project, add the following to the **\<system.web\>** node.

    ```XML
    <webServices>
      <protocols>
        <add name="HttpGet" />
        <add name="HttpPost" />
      </protocols>
    </webServices>
    ```

1. Save your changes.

## Components of a dictionary add-in

A dictionary add-in consists of three main component files:

- An XML-formatted add-in only manifest file that describes the add-in.
- An HTML file that provides the add-in's UI.
- A JavaScript file that provides logic to get the user's selection from the document, sends the selection as a query to the web service, and then displays returned results in the add-in's UI.

### Example of a dictionary add-in's manifest file

The following is an example manifest file for a dictionary add-in.

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xsi:type="TaskPaneApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>15.0</Version>
  <ProviderName>Microsoft Office Demo Dictionary</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <!--DisplayName is the name that will appear in the user's list of applications.-->
  <DisplayName DefaultValue="Microsoft Office Demo Dictionary" />
  <!--Description is a 2-3 sentence description of this dictionary. -->
  <Description DefaultValue="The Microsoft Office Demo Dictionary is an example built to demonstrate how a
    publisher can create a dictionary that integrates with Office. It doesn't return real definitions." />
  <!--IconUrl is the URI for the icon that will appear in the user's list of applications.-->
  <IconUrl DefaultValue="http://contoso/_layouts/images/general/office_logo.jpg" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <!--Hosts specifies the kind of Office application your dictionary add-in will support.
      You shouldn't have to modify this area.-->
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <!--SourceLocation is the URL for your dictionary.-->
    <SourceLocation DefaultValue="http://contoso/ExampleDictionary/DictionaryHome.html" />
  </DefaultSettings>
  <!--Permissions is the set of permissions a user will have to give your dictionary.
      If you need write access, such as to allow a user to replace the highlighted word with a synonym,
      use ReadWriteDocument. -->
  <Permissions>ReadDocument</Permissions>
  <Dictionary>
    <!--TargetDialects is the set of regional languages your dictionary contains. For example, if your
        dictionary applies to Spanish (Mexico) and Spanish (Peru), but not Spanish (Spain), you can specify
        that here. Do not put more than one language (for example, Spanish and English) here. Publish
        separate languages as separate dictionaries. -->
    <TargetDialects>
      <TargetDialect>EN-AU</TargetDialect>
      <TargetDialect>EN-BZ</TargetDialect>
      <TargetDialect>EN-CA</TargetDialect>
      <TargetDialect>EN-029</TargetDialect>
      <TargetDialect>EN-HK</TargetDialect>
      <TargetDialect>EN-IN</TargetDialect>
      <TargetDialect>EN-ID</TargetDialect>
      <TargetDialect>EN-IE</TargetDialect>
      <TargetDialect>EN-JM</TargetDialect>
      <TargetDialect>EN-MY</TargetDialect>
      <TargetDialect>EN-NZ</TargetDialect>
      <TargetDialect>EN-PH</TargetDialect>
      <TargetDialect>EN-SG</TargetDialect>
      <TargetDialect>EN-ZA</TargetDialect>
      <TargetDialect>EN-TT</TargetDialect>
      <TargetDialect>EN-GB</TargetDialect>
      <TargetDialect>EN-US</TargetDialect>
      <TargetDialect>EN-ZW</TargetDialect>
    </TargetDialects>
    <!--QueryUri is the address of this dictionary's XML web service (which is used to put definitions in
        additional contexts, such as the spelling checker.)-->
    <QueryUri DefaultValue="http://contoso/ExampleDictionary/WebService.asmx/Define?word="/>
    <!--Citation Text, Dictionary Name, and Dictionary Home Page will be combined to form the citation line
        (for example, this would produce "Examples by: Contoso",
        where "Contoso" is a hyperlink to http://www.contoso.com).-->
    <CitationText DefaultValue="Examples by: " />
    <DictionaryName DefaultValue="Contoso" />
    <DictionaryHomePage DefaultValue="http://www.contoso.com" />
  </Dictionary>
</OfficeApp>
```

The **\<Dictionary\>** element and its child elements specific to creating a dictionary add-in's manifest file are described in the following sections. For information about the other elements in the manifest file, see [Office Add-ins with the add-in only manifest](../develop/xml-manifest-overview.md).

### Dictionary element

Specifies settings for dictionary add-ins.

**Parent element**

**\<OfficeApp\>**

**Child elements**

**\<TargetDialects\>**, **\<QueryUri\>**, **\<CitationText\>**, **\<Name\>**, **\<DictionaryHomePage\>**

**Remarks**

The **\<Dictionary\>** element and its child elements are added to the manifest of a task pane add-in when you create a dictionary add-in.

#### TargetDialects element

Specifies the regional languages that this dictionary supports. Required for dictionary add-ins.

**Parent element**

**\<Dictionary\>**

**Child element**

**\<TargetDialect\>**

**Remarks**

The **\<TargetDialects\>** element and its child elements specify the set of regional languages your dictionary contains. For example, if your dictionary applies to both Spanish (Mexico) and Spanish (Peru), but not Spanish (Spain), you can specify that in this element. Do not specify more than one language (e.g., Spanish and English) in this manifest. Publish separate languages as separate dictionaries.

**Example**

```XML
<TargetDialects>
  <TargetDialect>EN-AU</TargetDialect>
  <TargetDialect>EN-BZ</TargetDialect>
  <TargetDialect>EN-CA</TargetDialect>
  <TargetDialect>EN-029</TargetDialect>
  <TargetDialect>EN-HK</TargetDialect>
  <TargetDialect>EN-IN</TargetDialect>
  <TargetDialect>EN-ID</TargetDialect>
  <TargetDialect>EN-IE</TargetDialect>
  <TargetDialect>EN-JM</TargetDialect>
  <TargetDialect>EN-MY</TargetDialect>
  <TargetDialect>EN-NZ</TargetDialect>
  <TargetDialect>EN-PH</TargetDialect>
  <TargetDialect>EN-SG</TargetDialect>
  <TargetDialect>EN-ZA</TargetDialect>
  <TargetDialect>EN-TT</TargetDialect>
  <TargetDialect>EN-GB</TargetDialect>
  <TargetDialect>EN-US</TargetDialect>
  <TargetDialect>EN-ZW</TargetDialect>
</TargetDialects>
```

#### TargetDialect element

Specifies a regional language that this dictionary supports. Required for dictionary add-ins.

**Parent element**

**\<TargetDialects\>**

**Remarks**

Specify the value for a regional language in the RFC1766  `language` tag format, such as EN-US.

**Example**

```XML
<TargetDialect>EN-US</TargetDialect>
```

#### QueryUri element

Specifies the endpoint for the dictionary query service. Required for dictionary add-ins.

**Parent element**

**\<Dictionary\>**

**Remarks**

This is the URI of the XML web service for the dictionary provider. The properly escaped query will be appended to this URI.

**Example**

```XML
<QueryUri DefaultValue="http://msranlc-lingo1/proof.aspx?q="/>
```

#### CitationText element

Specifies the text to use in citations. Required for dictionary add-ins.

**Parent element**

**\<Dictionary\>**

**Remarks**

This element specifies the beginning of the citation text that will be displayed on a line below the content that is returned from the web service (for example, "Results by: " or "Powered by: ").

For this element, you can specify values for additional locales by using the **\<Override\>** element. For example, if a user is running the Spanish SKU of Office, but using an English dictionary, this allows the citation line to read "Resultados por: Bing" rather than "Results by: Bing". For more information about how to specify values for additional locales, see [Localization](../develop/xml-manifest-overview.md#localization).

**Example**

```XML
<CitationText DefaultValue="Results by: " />
```

#### DictionaryName element

Specifies the name of this dictionary. Required for dictionary add-ins.

**Parent element**

**\<Dictionary\>**

**Remarks**

This element specifies the link text in the citation text. Citation text is displayed on a line below the content that is returned from the web service.

For this element, you can specify values for additional locales.

**Example**

```XML
<DictionaryName DefaultValue="Bing Dictionary" />
```

#### DictionaryHomePage element

Specifies the URL of the home page for the dictionary. Required for dictionary add-ins.

**Parent element**

**\<Dictionary\>**

**Remarks**

This element specifies the link URL in the citation text. Citation text is displayed on a line below the content that is returned from the web service.

For this element, you can specify values for additional locales.

**Example**

```XML
<DictionaryHomePage DefaultValue="https://www.bing.com" />
```

### Update your dictionary add-in's manifest file

1. Open the manifest file in the add-in project.
1. Update the value of the **\<ProviderName\>** element with your name.
1. Replace the value of the **\<DisplayName\>** element's **\<DefaultValue\>** attribute with an appropriate name, for example, "Microsoft Office Demo Dictionary".
1. Replace the value of the **\<Description\>** element's **\<DefaultValue\>** attribute with an appropriate description, for example, "The Microsoft Office Demo Dictionary is an example built to demonstrate how a publisher could create a dictionary that integrates with Office. It doesn't return real definitions.".
1. Add the following code after the **\<Permissions\>** node, replacing "contoso" references with your own company name, then save your changes.

    ```XML
    <Dictionary>
      <!--TargetDialects is the set of regional languages your dictionary contains. For example, if your
          dictionary applies to Spanish (Mexico) and Spanish (Peru), but not Spanish (Spain), you can
          specify that here. Do not put more than one language (for example, Spanish and English) here.
          Publish separate languages as separate dictionaries. -->
      <TargetDialects>
        <TargetDialect>EN-AU</TargetDialect>
        <TargetDialect>EN-BZ</TargetDialect>
        <TargetDialect>EN-CA</TargetDialect>
        <TargetDialect>EN-029</TargetDialect>
        <TargetDialect>EN-HK</TargetDialect>
        <TargetDialect>EN-IN</TargetDialect>
        <TargetDialect>EN-ID</TargetDialect>
        <TargetDialect>EN-IE</TargetDialect>
        <TargetDialect>EN-JM</TargetDialect>
        <TargetDialect>EN-MY</TargetDialect>
        <TargetDialect>EN-NZ</TargetDialect>
        <TargetDialect>EN-PH</TargetDialect>
        <TargetDialect>EN-SG</TargetDialect>
        <TargetDialect>EN-ZA</TargetDialect>
        <TargetDialect>EN-TT</TargetDialect>
        <TargetDialect>EN-GB</TargetDialect>
        <TargetDialect>EN-US</TargetDialect>
        <TargetDialect>EN-ZW</TargetDialect>
      </TargetDialects>
      <!--QueryUri is the address of this dictionary's XML web service (which is used to put definitions in
          additional contexts, such as the spelling checker.)-->
      <QueryUri DefaultValue="~remoteAppUrl/DictionaryWebService.asmx"/>
      <!--Citation Text, Dictionary Name, and Dictionary Home Page will be combined to form the citation
          line (for example, this would produce "Examples by: Contoso", where "Contoso" is a hyperlink to
          http://www.contoso.com).-->
      <CitationText DefaultValue="Examples by: " />
      <DictionaryName DefaultValue="Contoso" />
      <DictionaryHomePage DefaultValue="http://www.contoso.com" />
    </Dictionary>
    ```

### Create a dictionary add-in's HTML user interface

The following two examples show the HTML and CSS files for the UI of the Demo Dictionary add-in. To view how the UI is displayed in the add-in's task pane, see Figure 6 following the code. To see how the implementation of the JavaScript provides programming logic for this HTML UI, see [Write the JavaScript implementation](#write-the-javascript-implementation) immediately following this section.

In the add-in's web application project in Visual Studio, you can replace the contents of the **./Home.html** file with the following sample HTML.

```HTML
<!DOCTYPE html>
<html>

<head>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />

    <!--The title will not be shown but is supplied to ensure valid HTML.-->
    <title>Example Dictionary</title>

    <!--Required library includes.-->
    <script type="text/javascript" src="https://ajax.microsoft.com/ajax/4.0/1/MicrosoftAjax.js"></script>
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>

    <!--Optional library includes.-->
    <script type="text/javascript" src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.5.1.js"></script>

    <!--App-specific CSS and JS.-->
    <link rel="Stylesheet" type="text/css" href="Home.css" />
    <script type="text/javascript" src="Home.js"></script>
</head>

<body>
    <div id="mainContainer">
        <div>INSTRUCTIONS</div>
        <ol>
            <li>Ensure there's text in the document.</li>
            <li>Select text.</li>
        </ol>
        <div id="header">
            <span id="headword"></span>
        </div>
        <div>DEFINITIONS</div>
        <ol id="definitions">
        </ol>
        <div id="SeeMore">
            <a id="SeeMoreLink" target="_blank">See More...</a>
        </div>
        <div id="message"></div>
    </div>
</body>

</html>
```

The following example shows the contents of the .css file.

In the add-in's web application project in Visual Studio, you can replace the contents of the **./Home.css** file with the following sample CSS.

```CSS
#mainContainer
{
    font-family: Segoe UI;
    font-size: 11pt;
}

#headword
{
    font-family: Segoe UI Semibold;
    color: #262626;
}

#definitions
{
    font-size: 8.5pt;
}
a
{
    font-size: 8pt;
    color: #336699;
    text-decoration: none;
}
a:visited
{
    color: #993366;
}
a:hover, a:active
{  
    text-decoration: underline;
}
```

*Figure 6. Demo dictionary UI*

:::image type="content" source="../images/dictionary-add-in-06.png" alt-text="Demo dictionary UI.":::

### Write the JavaScript implementation

The following example shows the JavaScript implementation in the .js file that's called from the add-in's HTML page to provide the programming logic for the Demo Dictionary add-in. This script uses the XML web service described previously. When placed in the same directory as the example web service, the script will get definitions from that service. It can be used with a public OfficeDefinitions-conforming XML web service by modifying the `xmlServiceURL` variable at the top of the file.

The primary members of the Office JavaScript API (Office.js) that are called from this implementation are shown in the following list.

- The [initialize](/javascript/api/office) event of the `Office` object, which is raised when the add-in context is initialized, and provides access to a [Document](/javascript/api/office/office.document) object instance that represents the document the add-in is interacting with.
- The [addHandlerAsync](/javascript/api/office/office.document#office-office-document-addhandlerasync-member(1)) method of the `Document` object, which is called in the `initialize` function to add an event handler for the [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event of the document to listen for user selection changes.
- The [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) method of the `Document` object, which is called in the `tryUpdatingSelectedWord()` function when the `SelectionChanged` event handler is raised to get the word or phrase the user selected, coerce it to plain text, and then execute the `selectedTextCallback` asynchronous callback function.
- When the  `selectTextCallback` asynchronous callback function that's passed as the *callback* argument of the `getSelectedDataAsync` method executes, it gets the value of the selected text when the callback returns. It gets that value from the callback's *selectedText* argument (which is of type [AsyncResult](/javascript/api/office/office.asyncresult)) by using the [value](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) property of the returned `AsyncResult` object.
- The rest of the code in the `selectedTextCallback` function queries the XML web service for definitions.
- The remaining code in the .js file displays the list of definitions in the add-in's HTML UI.

In the add-in's web application project in Visual Studio, you can replace the contents of the **./Home.js** file with the following sample JavaScript.

```js
// The document the dictionary add-in is interacting with.
let _doc;
// The last looked-up word, which is also the currently displayed word.
let lastLookup;

// The base URL for the OfficeDefinitions-conforming XML web service to query for definitions.
const xmlServiceUrl = "DictionaryWebService.asmx/Define";

// Initialize the add-in.
// Office.initialize or Office.onReady is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Store a reference to the current document.
        _doc = Office.context.document;
        // Check whether text is already selected.
        tryUpdatingSelectedWord();
        // Add a handler to refresh when the user changes selection.
        _doc.addHandlerAsync("documentSelectionChanged", tryUpdatingSelectedWord);
    });
}

// Executes when event is raised on the user's selection changes, and at initialization time.
// Gets the current selection and passes that to asynchronous callback function.
function tryUpdatingSelectedWord() {
    _doc.getSelectedDataAsync(Office.CoercionType.Text, selectedTextCallback);
}

// Async callback that executes when the add-in gets the user's selection. Determines whether anything should
// be done. If so, it makes requests that will be passed to various functions.
function selectedTextCallback(selectedText) {
    selectedText = $.trim(selectedText.value);
    // Be sure user has selected text. The SelectionChanged event is raised every time the user moves
    // the cursor, even if no selection.
    if (selectedText != "") {
        // Check whether the user selected the same word the pane is currently displaying to
        // avoid unnecessary web calls.
        if (selectedText != lastLookup) {
            // Update the lastLookup variable.
            lastLookup = selectedText;
            // Set the "headword" span to the word you looked up.
            $("#headword").text("Selected text: " + selectedText);
            // AJAX request to get definitions for the selected word; pass that to refreshDefinitions.
            $.ajax(xmlServiceUrl,
                {
                    data: { word: selectedText },
                    dataType: 'xml',
                    success: refreshDefinitions,
                    error: errorHandler
                });
    }
}

// This function is called when the add-in gets back the definitions target word.
// It removes the old definitions and replaces them with the definitions for the current word.
// It also sets the "See More" link.
function refreshDefinitions(data, textStatus, jqXHR) {
    $(".definition").remove();

    // Make a new list item for each returned definition that was returned, set the CSS class,
    // and append it to the definitions div.
    $(data).find("Definition").each(function () {
        $(document.createElement("li"))
            .text($(this).text())
            .addClass("definition")
            .appendTo($("#definitions"));
    });

    // Change the "See More" link to direct to the correct URL.
    $("#SeeMoreLink").attr("href", $(data).find("SeeMoreURL").text());
}

// Basic error handler that writes to a div with id='message'.
function errorHandler(jqXHR, textStatus, errorThrown) {
    document.getElementById('message').innerText
      += ("textStatus:- " + textStatus
          + "\nerrorThrown:- " + errorThrown
          + "\njqXHR:- " + JSON.stringify(jqXHR));
}
```

## Try it out

1. Using Visual Studio, test the newly created Word add-in by pressing **F5** or choosing **Debug** > **Start Debugging** to launch Word with the **Show Taskpane** add-in button displayed on the ribbon. The add-in will be hosted locally on IIS.

1. In Word, if the add-in task pane isn't already open, choose the **Home** tab, and then choose the **Show Taskpane** button to open the add-in task pane. (If you're using the volume-licensed perpetual version of Office, instead of the Microsoft 365 version or a retail perpetual version, then custom buttons aren't supported. Instead, the task pane will open immediately.)

    ![The Word application with the Show Taskpane button highlighted.](../images/word-quickstart-addin-0.png)

1. In Word, add text to the document then select any or all of that text.

    :::image type="content" source="../images/dictionary-add-in-06.png" alt-text="Dictionary task pane UI.":::
