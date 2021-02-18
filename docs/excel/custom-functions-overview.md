---
ms.date: 01/08/2020
description: 'Create an Excel custom function for your Office Add-in.'
title: Create custom functions in Excel
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
---
# Create custom functions in Excel

Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.

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

> [!TIP]
> If your custom function add-in will use a task pane or a ribbon button, in addition to running custom function code, you will need to set up a shared JavaScript runtime. See [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) to learn more.

## How a custom function is defined in code

If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, it creates files which control your functions and task pane. We'll concentrate on the files that are important to custom functions:

| File | File format | Description |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>or<br/>**./src/functions/functions.ts** | JavaScript<br/>or<br/>TypeScript | Contains the code that defines custom functions. |
| **./src/functions/functions.html** | HTML | Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions. |
| **./manifest.xml** | XML | Specifies the location of multiple files that your custom function use, such as the custom functions JavaScript, JSON, and HTML files. It also lists the locations of task pane files, command files, and specifies which runtime your custom functions should use. |

### Script file

The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions and comments which define the function.

The following code defines the custom function `add`. The code comments are used to generate a JSON metadata file that describes the custom function to Excel. The required `@customfunction` comment is declared first, to indicate that this is a custom function. Next, two parameters are declared, `first` and `second`, followed by their `description` properties. Finally, a `returns` description is given. For more information about what comments are required for your custom function, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).

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

### Manifest file

The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) does several things:

- Defines the namespace for your custom functions. A namespace prepends itself to your custom functions to help customers identify your functions as part of your add-in.
- Uses `<ExtensionPoint>` and `<Resources>` elements that are unique to a custom functions manifest. These elements contain the information about the locations of the JavaScript, JSON, and HTML files.
- Specifies which runtime to use for your custom function. We recommend always using a shared runtime unless you have a specific need for another runtime, because a shared runtime allows for the sharing of data between functions and the task pane. Note that using a shared runtime means your add-in will use Internet Explorer 11, not Microsoft Edge.

If you are using the Yo Office generator to create files, we recommend adjusting your manifest to use a shared runtime, as this is not the default for these files. To change your manifest, follow the instructions in [Configure your Excel add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

To see a full working manifest from a sample add-in, see [this Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).

[!include[manifest guidance](../includes/manifest-guidance.md)]

## Coauthoring

Excel on the web and on Windows connected to a Microsoft 365 subscription allow you to coauthor in Excel. If your workbook uses a custom function, your coauthoring colleague is prompted to load the custom function's add-in. Once you both have loaded the add-in, the custom function shares results through coauthoring.

For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).

## Next steps

Want to try out custom functions? Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.

Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel. You can try out creating your own custom function or play with the provided samples.

## See also 
* [Learn about the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program)
* [Custom functions requirement sets](custom-functions-requirement-sets.md)
* [Custom functions naming guidelines](custom-functions-naming.md)
* [Make your custom functions compatible with XLL user-defined functions](make-custom-functions-compatible-with-xll-udf.md)
* [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
