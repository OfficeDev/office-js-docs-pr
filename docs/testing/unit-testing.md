---
title: Unit testing in Office Add-ins
description: Learn how to unit test code that calls the Office JavaScript APIs.
ms.date: 11/25/2025
ms.localizationpriority: medium
---

# Unit testing in Office Add-ins

Unit tests check your add-in's functionality without requiring network or service connections, including connections to the Office application. Unit testing server-side code, and client-side code that does *not* call the [Office JavaScript APIs](../develop/understanding-the-javascript-api-for-office.md), is the same in Office Add-ins as it is in any web application, so it requires no special documentation. But client-side code that calls the Office JavaScript APIs is challenging to test. To solve these problems, we created a library to simplify the creation of mock Office objects in unit tests: [Office-Addin-Mock](https://www.npmjs.com/package/office-addin-mock). The library makes testing easier in the following ways:

- The Office JavaScript APIs must initialize in a webview control in the context of an Office application (such as Excel, PowerPoint, or Word), so they can't load in the process in which unit tests run on your development computer. You can import the Office-Addin-Mock library into your test files, which enables the mocking of Office JavaScript APIs inside the Node.js process in which the tests run.
- The [application-specific APIs](../develop/understanding-the-javascript-api-for-office.md#api-models) have [load](../develop/application-specific-api-model.md#load) and [sync](../develop/application-specific-api-model.md#sync) methods that you must call in a particular order relative to other functions and to each other. Moreover, you must call the `load` method with certain parameters depending on the properties of Office objects that will be read by the code *later* in the function being tested. But unit testing frameworks are inherently stateless, so they can't keep a record of whether `load` or `sync` was called or what parameters were passed to `load`. The mock objects that you create with the Office-Addin-Mock library have internal state that keeps track of these things. This internal state enables the mock objects to emulate the error behavior of actual Office objects. For example, if the function that is being tested tries to read a property that wasn't first passed to `load`, then the test returns an error similar to what Office would return.

The library doesn't depend on the Office JavaScript APIs and you can use it with any JavaScript unit testing framework, such as:

- [Jest](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [Jasmine](https://jasmine.github.io/)

The examples in this article use the Jest framework. For examples that use the Mocha framework, see [the Office-Addin-Mock home page](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#examples).

## Prerequisites

This article assumes that you're familiar with the basic concepts of unit testing and mocking, including how to create and run test files, and that you have some experience with a unit testing framework.

> [!TIP]
> If you're working with Microsoft Visual Studio (VS), we recommend that you read the article [Unit testing JavaScript and TypeScript in Visual Studio](/visualstudio/javascript/unit-testing-javascript-with-visual-studio) for some basic information about JavaScript unit testing in VS and then return to this article.

## Install the tool

To install the library, open a command prompt, navigate to the root of your add-in project, and enter the following command.

```command&nbsp;line
npm install office-addin-mock --save-dev
```

## Basic usage

1. Your project will have one or more test files. (See the instructions for your test framework and the example test files in [Examples](#examples) below.) Import the library with either the `require` or `import` keyword to any test file that has a test of a function that calls the Office JavaScript APIs, as shown in the following examples.

   ```javascript
   // CommonJS
   const OfficeAddinMock = require("office-addin-mock");
   ```

   ```javascript
   // ES6
   import OfficeAddinMock from "office-addin-mock";
   ```

1. Import the module that contains the add-in function that you want to test with either the `require` or `import` keyword. The following examples assume your test file is in a subfolder of the folder with your add-in's code files.

   ```javascript
   // CommonJS
   const myOfficeAddinFeature = require("../my-office-add-in");
   ```

   ```javascript
   // ES6
   import myOfficeAddinFeature from "../my-office-add-in";
   ```

1. Create a data object that has the properties and subproperties that you need to mock to test the function. The following example shows an object that mocks the Excel [Workbook.range.address](/javascript/api/excel/excel.range#excel-excel-range-address-member) property and the [Workbook.getSelectedRange](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedrange-member(1)) method. This object isn't the final mock object. Think of it as a seed object that `OfficeMockObject` uses to create the final mock object.

   ```javascript
   const mockData = {
     workbook: {
       range: {
         address: "C2:G3",
       },
       getSelectedRange: function () {
         return this.range;
       },
     },
   };
   ```

1. Pass the data object to the `OfficeMockObject` constructor. Note the following about the returned `OfficeMockObject` object.

   - It's a simplified mock of an [OfficeExtension.ClientRequestContext](/javascript/api/office/officeextension.clientrequestcontext) object.
   - The mock object has all the members of the data object and also has mock implementations of the `load` and `sync` methods.
   - The mock object mimics crucial error behavior of the `ClientRequestContext` object. For example, if the Office API you're testing tries to read a property without first loading the property and calling `sync`, then the test fails with an error similar to what would be thrown in production runtime: "Error, property not loaded".

   ```javascript
   const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);
   ```

    > [!NOTE]
    > Full reference documentation for the `OfficeMockObject` type is at [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

1. In the syntax of your test framework, add a test of the function. Use the `OfficeMockObject` object in place of the object that it mocks, in this case the `ClientRequestContext` object. The following continues the example in Jest. This example test assumes that the add-in function that is being tested is called `getSelectedRangeAddress`, that it takes a `ClientRequestContext` object as a parameter, and that it returns the address of the currently selected range. The full example is [later in this article](#mocking-a-clientrequestcontext-object).

   ```javascript
   test("getSelectedRangeAddress should return the address of the range", async function () {
     expect(await getSelectedRangeAddress(contextMock)).toBe("C2:G3");
   });
   ```

1. Run the test in accordance with documentation of the test framework and your development tools. Typically, there's a **package.json** file with a script that executes the test framework. For example, if Jest is the framework, **package.json** would contain the following:

   ```json
   "scripts": {
     "test": "jest",
     -- Other scripts omitted. --  
   }
   ```

   To run the test, enter the following in a command prompt in the root of the project.

   ```command&nbsp;line
   npm test
   ```

## Examples

The examples in this section use Jest with its default settings. These settings support CommonJS modules. For information about how to configure Jest and Node.js to use TypeScript and support ECMAScript modules, see the Jest documentation on [getting started](https://jestjs.io/docs/getting-started) and [ECMAScript Modules](https://jestjs.io/docs/ecmascript-modules).

To run any of these examples, take the following steps.

1. Create an Office Add-in project for the appropriate Office host application (for example, Excel or Word). One way to do this quickly is to use the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).
1. In the root of the project, [install Jest](https://jestjs.io/docs/getting-started).
1. [Install the office-addin-mock tool](#install-the-tool).
1. Create a file exactly like the first file in the example and add it to the folder that contains the project's other source files, often called `\src`.
1. Create a subfolder to the source file folder and give it an appropriate name, such as `\tests`.
1. Create a file exactly like the test file in the example and add it to the subfolder.
1. Add a `test` script to the **package.json** file, and then run the test, as described in [Basic usage](#basic-usage).

### Mocking the Office Common APIs

This example assumes an Office Add-ins for any host that supports the [Office Common APIs](../develop/office-javascript-api-object-model.md) (for example, Excel, PowerPoint, or Word). The add-in has one of its features in a file named `my-common-api-add-in-feature.js`. The following code shows the contents of the file. The `addHelloWorldText` function sets the text "Hello World!" to whatever is currently selected in the document; for example, a range in Word, or a cell in Excel, or a text box in PowerPoint.

```javascript
const myCommonAPIAddinFeature = {

    addHelloWorldText: async () => {
        const options = { coercionType: Office.CoercionType.Text };
        await Office.context.document.setSelectedDataAsync("Hello World!", options);
    }
}
  
module.exports = myCommonAPIAddinFeature;
```

The test file, named `my-common-api-add-in-feature.test.js`, is in a subfolder, relative to the location of the add-in code file. The following code shows the contents of the file. The top level property is `context`, an [Office.Context](/javascript/api/office/office.context) object, so the object that is being mocked is the parent of this property: an [Office](/javascript/api/office) object. Note the following about this code:

- The `OfficeMockObject` constructor doesn't add all of the Office enum classes to the mock `Office` object, so you must add the `CoercionType.Text` value that is referenced in the add-in method explicitly in the seed object.
- Because the Office JavaScript library isn't loaded in the node process, you must declare and initialize the `Office` object that is referenced in the add-in code.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myCommonAPIAddinFeature = require("../my-common-api-add-in-feature");

// Create the seed mock object.
const mockData = {
    context: {
      document: {
        setSelectedDataAsync: function (data, options) {
          this.data = data;
          this.options = options;
        },
      },
    },
    // Mock the Office.CoercionType enum.
    CoercionType: {
      Text: {},
    },
};
  
// Create the final mock object from the seed object.
const officeMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Create the Office object that is called in the addHelloWorldText function.
global.Office = officeMock;

/* Code that calls the test framework goes below this line. */

// Jest test
test("Text of selection in document should be set to 'Hello World'", async function () {
    await myCommonAPIAddinFeature.addHelloWorldText();
    expect(officeMock.context.document.data).toBe("Hello World!");
});
```

### Mocking the Outlook APIs

Although the Outlook APIs are part of the Common API model, they have a special architecture that's built around the [Mailbox](/javascript/api/outlook/office.mailbox) object. We provide a distinct example for Outlook. This example assumes an Outlook add-in that has one of its features in a file named `my-outlook-add-in-feature.js`. The following code shows the contents of the file. The `addHelloWorldText` function sets the text "Hello World!" to whatever is currently selected in the message compose window.

```javascript
const myOutlookAddinFeature = {

    addHelloWorldText: async () => {
        Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
      }
}

module.exports = myOutlookAddinFeature;
```

The test file, named `my-outlook-add-in-feature.test.js`, is in a subfolder relative to the location of the add-in code file. The following code shows the contents of the file. The top level property is `context`, an [Office.Context](/javascript/api/office/office.context) object. The object that the mock targets is the parent of this property: an [Office](/javascript/api/office) object. Note the following details about this code.

- The `host` property on the mock object is used internally by the mock library to identify the Office application. It's mandatory for Outlook. It currently serves no purpose for any other Office application.
- Because the Office JavaScript library isn't loaded in the node process, you must declare and initialize the `Office` object that is referenced in the add-in code.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myOutlookAddinFeature = require("../my-outlook-add-in-feature");

// Create the seed mock object.
const mockData = {
  // Identify the host to the mock library (required for Outlook).
  host: "outlook",
  context: {
    mailbox: {
      item: {
          setSelectedDataAsync: function (data) {
          this.data = data;
        },
      },
    },
  },
};
  
// Create the final mock object from the seed object.
const officeMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Create the Office object that is called in the addHelloWorldText function.
global.Office = officeMock;

/* Code that calls the test framework goes below this line. */

// Jest test
test("Text of selection in message should be set to 'Hello World'", async function () {
    await myOutlookAddinFeature.addHelloWorldText();
    expect(officeMock.context.mailbox.item.data).toBe("Hello World!");
});
```

### Mocking the Office application-specific APIs

When you test functions that use the application-specific APIs, mock the right type of object. You have two options.

- Mock a [OfficeExtension.ClientRequestObject](/javascript/api/office/officeextension.clientrequestcontext). Use this option when the function you're testing meets both of the following conditions.

  - It doesn't call a *Host*.`run` function, such as [Excel.run](/javascript/api/excel#Excel_run_batch_).
  - It doesn't reference any other direct property or method of a *Host* object.

- Mock a *Host* object, such as [Excel](/javascript/api/excel) or [Word](/javascript/api/word). Use this option when the preceding option isn't possible.

The following subsections show examples of both types of tests.

> [!NOTE]
> The Office-Addin-Mock library doesn't currently support mocking collection type objects. These objects are all the objects in the application-specific APIs that follow the *Collection* naming pattern, such as WorksheetCollection. We're working hard to add this support to the library.

#### Mocking a ClientRequestContext object

This example assumes an Excel add-in that has one of its features in a file named `my-excel-add-in-feature.js`. The following code shows the contents of the file. Note that the `getSelectedRangeAddress` function is a helper method called inside the callback that is passed to `Excel.run`.

```javascript
const myExcelAddinFeature = {
    
    getSelectedRangeAddress: async (context) => {
        const range = context.workbook.getSelectedRange();      
        range.load("address");

        await context.sync();
      
        return range.address;
    }
}

module.exports = myExcelAddinFeature;
```

The test file, named `my-excel-add-in-feature.test.js`, is in a subfolder relative to the location of the add-in code file. The following code shows the contents of the file. Note that the top level property is `workbook`, so the object that is being mocked is the parent of an `Excel.Workbook`: a `ClientRequestContext` object.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myExcelAddinFeature = require("../my-excel-add-in-feature");

// Create the seed mock object.
const mockData = {
    workbook: {
      range: {
        address: "C2:G3",
      },
      // Mock the Workbook.getSelectedRange method.
      getSelectedRange: function () {
        return this.range;
      },
    },
};

// Create the final mock object from the seed object.
const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);

/* Code that calls the test framework goes below this line. */

// Jest test
test("getSelectedRangeAddress should return address of selected range", async function () {
  expect(await myOfficeAddinFeature.getSelectedRangeAddress(contextMock)).toBe("C2:G3");
});
```

#### Mocking a host object

This example assumes a Word add-in that has one of its features in a file named `my-word-add-in-feature.js`. The following code shows the contents of the file.

```javascript
const myWordAddinFeature = {

  insertBlueParagraph: async () => {
    return Word.run(async (context) => {
      // Insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
  
      // Change the font color to blue.
      paragraph.font.color = "blue";
  
      await context.sync();
    });
  }
}

module.exports = myWordAddinFeature;
```

The test file, named `my-word-add-in-feature.test.js`, is in a subfolder relative to the location of the add-in code file. The following code shows the contents of the file. Note that the top level property is `context`, a `ClientRequestContext` object, so the object that is being mocked is the parent of this property: a `Word` object. Note the following details about this code.

- When the `OfficeMockObject` constructor creates the final mock object, it ensures that the child `ClientRequestContext` object has `sync` and `load` methods.
- The `OfficeMockObject` constructor doesn't add a `run` function to the mock `Word` object, so you must add it explicitly in the seed object.
- The `OfficeMockObject` constructor doesn't add all of the Word enum classes to the mock `Word` object, so you must add the `InsertLocation.end` value that is referenced in the add-in method explicitly in the seed object.
- Because the Office JavaScript library isn't loaded in the node process, you must declare and initialize the `Word` object that's referenced in the add-in code.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myWordAddinFeature = require("../my-word-add-in-feature");

// Create the seed mock object.
const mockData = {
  context: {
    document: {
      body: {
        paragraph: {
          font: {},
        },
        // Mock the Body.insertParagraph method.
        insertParagraph: function (paragraphText, insertLocation) {
          this.paragraph.text = paragraphText;
          this.paragraph.insertLocation = insertLocation;
          return this.paragraph;
        },
      },
    },
  },
  // Mock the Word.InsertLocation enum.
  InsertLocation: {
    end: "end",
  },
  // Mock the Word.run function.
  run: async function(callback) {
    await callback(this.context);
  },
};

// Create the final mock object from the seed object.
const wordMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Define and initialize the Word object that is called in the insertBlueParagraph function.
global.Word = wordMock;

/* Code that calls the test framework goes below this line. */

// Jest test set
describe("Insert blue paragraph at end tests", () => {

  test("color of paragraph", async function () {
    await myWordAddinFeature.insertBlueParagraph();  
    expect(wordMock.context.document.body.paragraph.font.color).toBe("blue");
  });

  test("text of paragraph", async function () {
    await myWordAddinFeature.insertBlueParagraph();
    expect(wordMock.context.document.body.paragraph.text).toBe("Hello World");
  });
})
```

> [!NOTE]
> Full reference documentation for the `OfficeMockObject` type is at [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

## See also

- [Office-Addin-Mock npm page](https://www.npmjs.com/package/office-addin-mock) installation point.
- The open source repo is [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock).
- [Jest](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [Jasmine](https://jasmine.github.io/)
