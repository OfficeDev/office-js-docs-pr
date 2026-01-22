---
description: Create an Excel custom function for your Office Add-in.
title: Create custom functions in Excel
ms.date: 01/22/2026
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---
# Create custom functions in Excel

Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.

[!include[Excel custom functions definition](../includes/excel-custom-functions-definition.md)]

The following animated image shows your workbook calling a function you've created with JavaScript or TypeScript. In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.

:::image type="content" source="../images/SphereVolumeNew.gif" alt-text="Animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet.":::

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

## How a custom function is defined in code

If you use the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) to create an Excel custom functions add-in project, it creates files which control your functions and task pane. We'll concentrate on the files that are important to custom functions.

| File | File format | Description |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>or<br/>**./src/functions/functions.ts** | JavaScript<br/>or<br/>TypeScript | Contains the code that defines custom functions. |
| **./src/functions/functions.html** | HTML | Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions. |
| **./manifest.xml** | XML | Specifies the location of multiple files that your custom function use, such as the custom functions JavaScript, JSON, and HTML files. It also lists the locations of task pane files, command files, and specifies which runtime your custom functions should use. |

> [!TIP]
> The Yeoman generator for Office Add-ins offers multiple **Excel Custom Functions** projects. We recommend selecting the project type **Excel Custom Functions using a Shared Runtime** and the script type **JavaScript**.

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

> [!TIP]
> In Excel on the web, custom function descriptions and parameter descriptions display inline. This gives users additional information when writing custom functions. Learn how to configure the inline descriptions by exploring any of the [custom functions Script Lab samples](https://github.com/OfficeDev/office-js-snippets/tree/prod/samples/excel/16-custom-functions) in Excel on the web. See the following screenshot for an example.
>
> :::image type="content" source="../images/custom-functions-inline-description.png" alt-text="A custom function with inline descriptions displayed in Excel on the web.":::

### Manifest file

The add-in only manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) creates) does several things.

- Defines the namespace for your custom functions. A namespace prepends itself to your custom functions to help customers identify your functions as part of your add-in.
- Uses `<ExtensionPoint>` and `<Resources>` elements that are unique to a custom functions manifest. These elements contain the information about the locations of the JavaScript, JSON, and HTML files.
- Specifies which runtime to use for your custom function. We recommend always using a shared runtime unless you have a specific need for another runtime, because a shared runtime allows for the sharing of data between functions and the task pane.

To see a full working manifest from a sample add-in, see the manifest in the [one of our Office Add-in samples Github repositories](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/excel-shared-runtime-global-state/manifest.xml).

[!include[manifest guidance](../includes/manifest-guidance.md)]

## Coauthoring

Excel on the web and on Windows connected to a Microsoft 365 subscription allow end users to coauthor in Excel. If an end user's workbook uses a custom function, that end user's coauthoring colleague is prompted to load the corresponding custom functions add-in. Once both users have loaded the add-in, the custom function shares results through coauthoring.

For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).

## Supported platforms

Excel custom functions are supported by most Office client applications. Excel custom functions aren't currently supported in **Office on iPad** or **volume-licensed perpetual versions of Office 2021 or earlier on Windows**. For more information, see [Custom functions requirement sets](/javascript/api/requirement-sets/excel/custom-functions-requirement-sets).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## Next steps

Want to try out custom functions? Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.

Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862), an add-in that allows you to experiment with custom functions right in Excel. You can try out creating your own custom function or play with the provided samples.

## See also

- [Learn about the Microsoft 365 Developer Program](https://aka.ms/m365devprogram)
- [Custom functions requirement sets](/javascript/api/requirement-sets/excel/custom-functions-requirement-sets)
- [Custom functions naming guidelines](custom-functions-naming.md)
- [Make your custom functions compatible with XLL user-defined functions](make-custom-functions-compatible-with-xll-udf.md)
- [Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Runtimes in Office Add-ins](../testing/runtimes.md)
