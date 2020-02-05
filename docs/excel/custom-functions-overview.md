---
ms.date: 12/28/2019
description: Create custom functions in Excel using JavaScript.
title: Create custom functions in Excel
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Normal
---

# Create custom functions in Excel 

Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`. This article describes how to create custom functions in Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

The following animated image shows your workbook calling a function you've created with JavaScript or Typescript. In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

The following code defines the custom function `=MYFUNCTION.SPHEREVOLUME`.

```js
/**
 * Returns the volume of a sphere. 
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
```

> [!NOTE]
> The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.

## How a custom function is defined in code

If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll find that it creates files which control your functions, your task pane, and your add-in overall. We'll concentrate on the files that are important to custom functions:

| File | File format | Description |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>or<br/>**./src/functions/functions.ts** | JavaScript<br/>or<br/>TypeScript | Contains the code that defines custom functions. |
| **./src/functions/functions.html** | HTML | Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions. |
| **./manifest.xml** | XML | Specifies the namespace for all custom functions within the add-in and the location of the JavaScript and HTML files that are listed previously in this table. It also lists the locations of other files your add-in might make use of, such as the task pane files and command files. |

### Script file

The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions and comments which define the function.

The following code defines the custom function `add`. The code comments are used to generate a JSON metadata file that describes the custom function to Excel. The required `@customfunction` comment is declared first, to indicate that this is a custom function. Additionally, you'll notice two parameters are declared, `first` and `second`, which are followed by their `description` properties. Finally, a `returns` description is given. For more information about what comments are required for your custom function, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number.
 * @param second Second number.
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}
```

Note that the **functions.html** file, which governs the loading of the custom functions runtime, must link to the current CDN for custom functions. Projects prepared with the current version of the Yo Office generator reference the correct CDN. If you are retrofitting a previous custom function project from March 2019 or earlier, you need to copy in the code below to the **functions.html** page.

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/custom-functions-runtime.js" type="text/javascript"></script>
```

### Manifest file

The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.

The following basic XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions. If using the Yo Office generator, your generated custom function files will contain a more complex manifest file, which you can compare on [this Github repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml).

> [!NOTE] 
> The URLs specified in the manifest file for the custom functions JavaScript, JSON, and HTML files must be publicly accessible and have the same subdomain.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> Functions in Excel are prepended by the namespace specified in your XML manifest file. A function's namespace comes before the function name and they are separated by a period. For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file. The namespace is intended to be used as an identifier for your company or the add-in. A namespace can only contain alphanumeric characters and periods.

[!include[manifest guidance](../includes/manifest-guidance.md)]

## Coauthoring

Excel on the web and Windows connected to an Office 365 subscription allow you to coauthor documents and this feature works with custom functions. If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in. Once you both have loaded the add-in, the custom function will share results through coauthoring.

For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).

## Known issues

See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## Next steps

Want to try out custom functions? Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.

Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel. You can try out creating your own custom function or play with the provided samples.

Ready to read more about the capabilities custom functions? Learn about an overview of [the custom functions architecture](custom-functions-architecture.md).

## See also 
* [Custom functions requirements](custom-functions-requirement-sets.md)
* [Naming guidelines](custom-functions-naming.md)
* [Make your custom functions compatible with XLL user-defined functions](make-custom-functions-compatible-with-xll-udf.md)
