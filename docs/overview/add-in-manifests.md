
# Office Add-ins XML manifest



The XML manifest file of an Office Add-in enables you to declaratively describe how your add-in should be activated when an end user installs and uses it with Office documents and applications. The [offappmanifest.xsd](http://msdn.microsoft.com/en-us/library/d5f72bff-3446-c64f-02ca-ab10b5648789%28Office.15%29.aspx) defines an XML schema that is common to all supported Office applications (both rich client applications and their corresponding web app web clients).

 >**Note**  Your Office Add-in must use manifest schema version 1.1. For information about how to update your add-in to use manifest 1.1, see [Update the version of your JavaScript API for Office and manifest schema files](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

An XML manifest file based on this schema enables an Office Add-in to do the following:

- Describe itself by providing an ID, version, description, display name, and default locale.
    
- Specify the location of the HTML file that provides the UI of the Office Add-in.
    
- Specify the requested default dimensions for content Office Add-ins, and requested height for Outlook Office Add-ins.
    
- Declare permissions that the Office Add-in requires, such as reading or writing to the document.
    
- Specify how they should be used and displayed, such as content (in the document), in a task pane, or contextually with a message, appointment, or meeting request item.
    
- For Outlook Office Add-ins, define the rule or rules that specify the context in which they will be activated and interact with a message, appointment, or meeting request item.
    
For examples of manifest v1.1 XML files, see [Manifest v1.1 XML file examples](#O15AgaveManifestOverview_Samplev1_1).

## Required elements


The following table specifies the elements that are required for the three types of Office Add-ins and the Dictionary task pane add-in.


 >**Important**  For add-ins submitted to the Office Store, all add-in locations, such as the source file locations specified in the  **SourceLocation** element, must be SSL-secured (HTTPS). For more information see, [What are some common submission errors to avoid?](http://msdn.microsoft.com/library/0ceb385c-a608-40cc-8314-78e39d6c75d0%28Office.15%29.aspx#bk_q2)


**Required elements by Office Add-in type**


|**Element**|**Content**|**Task pane**|**Outlook**|
|:-----|:-----|:-----|:-----|
|[OfficeApp](http://msdn.microsoft.com/en-us/library/68f1cada-66f8-4341-45f5-14e0634c24fb%28Office.15%29.aspx)|X|X|X|
|[Id](http://msdn.microsoft.com/en-us/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx)|X|X|X|
|[Version](http://msdn.microsoft.com/en-us/library/6a8bbaa5-ee8c-6824-4aba-cb1a804269f6%28Office.15%29.aspx)|X|X|X|
|[ProviderName](http://msdn.microsoft.com/en-us/library/0062693a-fafa-ea2d-051a-75dac0f6c323%28Office.15%29.aspx)|X|X|X|
|[DefaultLocale](http://msdn.microsoft.com/en-us/library/04796a3a-3afa-dc85-db66-4677560c185c%28Office.15%29.aspx)|X|X|X|
|[DisplayName](http://msdn.microsoft.com/en-us/library/529159ca-53bf-efcf-c245-e572dab0ef57%28Office.15%29.aspx)|X|X|X|
|[Description](http://msdn.microsoft.com/en-us/library/bcce6bad-23d0-7631-7d8c-1064b8453b5a%28Office.15%29.aspx)|X|X|X|
|[IconUrl](http://msdn.microsoft.com/library/c7dac2d4-4fda-6fc7-3774-49f02b2d3e1e%28Office.15%29.aspx)|X|X|X|
|[HighResolutionIconUrl](http://msdn.microsoft.com/library/ff7b2647-ec8e-70dc-4e4a-e1a1377ff3f2%28Office.15%29.aspx)|||X|
|[DefaultSettings (ContentApp)](http://msdn.microsoft.com/en-us/library/f7edc689-551f-1a17-ea81-ffd58f534557%28Office.15%29.aspx)<br/>[DefaultSettings (TaskPaneApp)](http://msdn.microsoft.com/en-us/library/36e3d139-56a4-fb3d-0a21-cbd14e606765%28Office.15%29.aspx)|X|X||
|[SourceLocation (ContentApp)](http://msdn.microsoft.com/en-us/library/00d95bb0-e8f5-647f-790a-0aa3aabc8141%28Office.15%29.aspx)<br/>[SourceLocation (TaskPaneApp)](http://msdn.microsoft.com/en-us/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx)|X|X||
|[DesktopSettings](http://msdn.microsoft.com/en-us/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c%28Office.15%29.aspx)|||X|
|[SourceLocation (MailApp)](http://msdn.microsoft.com/en-us/library/3792d389-bebd-d19a-9d90-35b7a0bfc623%28Office.15%29.aspx)|||X|
|[Permissions (ContentApp)](http://msdn.microsoft.com/en-us/library/9f3dcf9c-fced-c115-4f0d-38d60fb7c583%28Office.15%29.aspx)<br/>[Permissions (TaskPaneApp)](http://msdn.microsoft.com/en-us/library/d4cfe645-353d-8240-8495-f76fb36602fe%28Office.15%29.aspx)<br/>[Permissions (MailApp)](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx)|X|X|X|
|[Rule (RuleCollection)](http://msdn.microsoft.com/en-us/library/c6ce9d52-4b53-c6a6-de7e-c64106135c81%28Office.15%29.aspx)<br/>[Rule (MailApp)](http://msdn.microsoft.com/en-us/library/56dfc32e-2b8c-1724-05be-5595baf38aa3%28Office.15%29.aspx)|||X|
|[Dictionary](http://msdn.microsoft.com/en-us/library/f78898f4-059e-d5dc-5eab-1f6b92214068%28Office.15%29.aspx)||||
|[*Requirements (MailApp)](http://msdn.microsoft.com/en-us/library/9536ea30-34f7-76b5-7f30-1508626840e4%28Office.15%29.aspx)||X|
|[*Set](http://msdn.microsoft.com/en-us/library/1506daa1-332c-30e1-6402-3371bcd0b895%28Office.15%29.aspx)<br/>[**Sets (MailAppRequirements)](http://msdn.microsoft.com/en-us/library/2a6a2484-eeee-37e4-43bc-c185e8ae0d1d%28Office.15%29.aspx)|||X|
|[*Form](http://msdn.microsoft.com/en-us/library/77a8ac83-c22b-1225-4fc4-ba4038b68648%28Office.15%29.aspx)<br/>[**FormSettings](http://msdn.microsoft.com/en-us/library/0d1a311d-939d-78c1-e968-89ddf7ebc4b4%28Office.15%29.aspx)|||X|
|[*Sets (Requirements)](http://msdn.microsoft.com/en-us/library/509be287-b532-87c6-71ac-64f3a4bbd3af%28Office.15%29.aspx)||X|
|[*Hosts](http://msdn.microsoft.com/library/f9a739c1-3daf-c03a-2bd9-4a2a6b870101%28Office.15%29.aspx)||X|
*Added in the Office Add-in Manifest Schema version 1.1.


## Manifest v1.1 XML file examples


The following sections show examples of manifest v1.1 XML files for content, task pane, Outlook, and dictionary Office Add-ins.

If you're using Visual Studio to develop your Office Add-in, you can use the Visual Studio manifest designer to change Office Add-in manifest settings, rather than manually changing the underlying XML markup. By default, when you open an Office Add-in manifest file in Visual Studio, it opens in the manifest designer. The designer organizes the fields in the manifest, making them easier to find. Some fields have drop-down list boxes that contain valid field values, helping reduce data entry errors.


### Content Office Add-in manifest v1.1 example



```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type="ContentApp">
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/ENUSIcon.png" />
  <Hosts>
    <Host Name="Workbook" />
    <Host Name="Database" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" />
    </Sets>
  </Requirements>  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com/apps/content.html" />
    <RequestedWidth>400</RequestedWidth> 
    <RequestedHeight>400</RequestedHeight>
  </DefaultSettings>
  <Permissions>Restricted</Permissions>
  <AllowSnapshot>true</AllowSnapshot>
</OfficeApp>
```


### Task pane Office Add-in manifest v1.1 example



```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="TaskPaneApp">
  <Id>412ce350-4161-4ad0-a5f5-0ec9d2cd3570</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample task pane add-in">
    <Override Locale="en-US" Value="Project add-in"/> 
  </DisplayName>
  <Description DefaultValue="Describe the features of this app.">
    <Override Locale="en-US" Value="Adds project management information to documents"/>
  </Description>
  <IconUrl DefaultValue="https://contoso.com.sa/ProjectApp/Topgunas-SA.png">
    <Override Locale="en-US" Value="https://contoso.com/ProjectApp/Topgunen-US.png"/>
  </IconUrl>
  <AppDomains>
    <AppDomain>http://www.projectlogin.com</AppDomain>
    <AppDomain>http://m.projectlogin.com</AppDomain>
    <AppDomain>http://www.projectlogin.com.sa</AppDomain>
    <AppDomain>http://m.projectlogin.com.sa</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name = "Document"/>
    <Host Name = "Workbook"/>
    <Host Name = "Presentation"/>
    <Host Name = "Project"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com.sa/ProjectApp/ProjectAppar_SA.html">
      <Override Locale="en-US" Value="https://contoso.com/ProjectApp/ProjectAppen-US.html"/>
    </SourceLocation>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```


### Outlook Office Add-in manifest v1.1 example



```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type="MailApp">

  <Id>971E76EF-D73E-567F-ADAE-5A76B39052CF</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="YouTube"/>
  <Description DefaultValue=
    "Watch YouTube videos referenced in the e-mails you  
    receive without leaving your email client.">
    <Override Locale="fr-fr" Value="Visualisez les vid????os
      YouTube r????f????renc????es dans vos courriers ????lectronique
      directement depuis Outlook et Outlook Web App."/>
  </Description>
  <!-- Change the following line to specify    -->
  <!-- the web serverthat hosts the icon file. -->
  <IconUrl DefaultValue=
    "https://webserver/YouTube/YouTubeLogo.png"/>

  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_desktop.htm" />
        <RequestedHeight>216</RequestedHeight>
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_tablet.htm" />
        <RequestedHeight>216</RequestedHeight>
      </TabletSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_desktop.htm" />
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_tablet.htm" />
      </TabletSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or"> 
    <Rule xsi:type="RuleCollection" Mode="And">
      <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
      </Rule>
      <Rule xsi:type="ItemHasRegularExpressionMatch" 
        PropertyName="BodyAsPlaintext" RegExName="VideoURL" 
        RegExValue=
        "http://(((www\.)?youtube\.com/watch\?v=)|
        (youtu\.be/))[a-zA-Z0-9_-]{11}" />
    </Rule>
    <Rule xsi:type="RuleCollection" Mode="Or">
      <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
      <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
    </Rule> 
  </Rule>
</OfficeApp>

```


### Dictionary Office Add-in manifest v1.1 example



```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type="TaskPaneApp">
  <Id>6f28f837-8326-4231-a033-ea1d80ff9fb3</Id>
  <Version>15.0</Version>
  <ProviderName>Microsoft Office Demo Dictionary</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <!--DisplayName is the name that will appear in the user's list of applications.-->
  <DisplayName DefaultValue="Microsoft Office Demo Dictionary" />
  <!--Description is a 2-3 sentence description of this dictionary. -->
  <Description DefaultValue="The Microsoft Office Demo Dictionary is an example built to demonstrate how a publisher could create a dictionary that integrates with Office. It does not return real definitions." />
  <!--IconUrl is the URI for the icon that will appear in the user's list of applications.-->
  <IconUrl DefaultValue="https://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
  <Hosts>
    <Host Name = "Document"/>
    <Host Name = "Workbook"/>
    <Host Name = "Presentation"/>
    <Host Name = "Project"/>
  </Hosts>
  <DefaultSettings>
    <!--Replace the value specified for DefaultValue below with the URL for your dictionary-->
    <SourceLocation DefaultValue="https://contoso.com/ExampleDictionary/DictionaryHome.html" />
  </DefaultSettings>
  <!--Permissions is the set of permissions a user will have to give your dictionary. If you need write access, such as to allow a user to replace the highlighted word with a synonym, use ReadWriteDocument. -->
  <Permissions>ReadDocument</Permissions>
  <Dictionary>
    <!--TargetDialects is the set of dialects your dictionary contains. For example, if your dictionary applies to Spanish (Mexico) and Spanish (Peru), but not Spanish (Spain), you can specify that here. This is for different dialects of the same language. Please do NOT put more than one language (e.g. Spanish and English) here. Publish separate languages as separate dictionaries. -->
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
    <!--QueryUri is the address of this dictionary's XML webservice (which we use to put definitions in additional contexts, such as the spelling checker.)-->
    <QueryUri DefaultValue="https://contoso.com/ExampleDictionary/WebService.asmx/Define?word="/>
    <!--Citation Text, Dictionary Name, and Dictionary Home Page will be combined to form the citation line - e.g. this would produce "Examples by: Microsoft", where "Microsoft" is a hyperlink to http://www.microsoft.com-->
    <CitationText DefaultValue="Examples by: " />
    <DictionaryName DefaultValue="Microsoft" />
    <DictionaryHomePage DefaultValue="https://www.microsoft.com" />
  </Dictionary>
</OfficeApp>
```


## Validate the Office Add-ins manifest


To make sure that the manifest file that describes your Office Add-in is correct and complete, validate it against the XML Schema Definition (XSD) file. You can use an XML schema validation tool or Visual Studio to validate the manifest. You can also download the [Office App Compatibility Kit](https://www.microsoft.com/en-us/download/details.aspx?id=46831) and run it on your add-in.

For information about validating a manifest against a schema, see [XML Schema (XSD) validation tool](http://stackoverflow.com/questions/124865/xml-schema-xsd-validation-tool).


 >**Note**  Validating your manifest against the offappmanifest.xsd file is obsolete. For information about updating your manifest to schema version 1.1, see [Update the version of your JavaScript API for Office and manifest schema files](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).


## Specify domains you want to open in the add-in window


By default, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](http://msdn.microsoft.com/en-us/library/00d95bb0-e8f5-647f-790a-0aa3aabc8141%28Office.15%29.aspx) element of the manifest file), that URL will open in a new browser window outside add-in pane of the Office host application. This default behavior protects the user against unexpected page navigation within the add-in pane from embedded **iframe** elements.

To override this behavior, specify each domain you want to open in the add-in window in the list of domains specified in the [AppDomains](http://msdn.microsoft.com/en-us/library/13cf867d-9b24-786f-0687-6bcdc954628e%28Office.15%29.aspx) element of the manifest file. If the add-in tries to go to a URL in a domain that isn't in the list, that URL will open in a new browser window (outside of the add-in pane).

The following XML manifest example hosts its main add-in page in the  `https://www.contoso.com` domain as specified in the **SourceLocation** element. It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](http://msdn.microsoft.com/en-us/library/2a0353ec-5e09-6fbf-1636-4bb5dcebb9bf%28Office.15%29.aspx) element within the **AppDomains** element list. If the add-in navigates to a page in the www.northwindtraders.com domain, that page will open in the add-in pane.




```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```


## Additional resources



- [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md)
    
- [Create add-in commands in your manifest for Excel, Word, and PowerPoint (Preview)](../../docs/design/create-add-in-commands-in-your-manifest-preview.md)
    
- [Localization for Office Add-ins](../../docs/develop/localization.md)
    
- [Schema reference for Office Add-ins manifests (v1.1)](http://msdn.microsoft.com/en-us/library/7e0cadc3-f613-8eb9-57ef-9032cbb97f92%28Office.15%29.aspx)
    
