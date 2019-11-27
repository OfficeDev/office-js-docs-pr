---
ms.date: 12/02/2019
description: Guide to sharing code between VSTO Add-in and Office web add-in.
title: Tutorial: Migrate your VSTO Add-in to an Office web add-in with a shared code library
ms.prod: excel
localization_priority: Normal
---

# Tutorial: Migrate your VSTO Add-in to an Office web add-in with a shared code library

VSTO Add-ins are great for extending Office to provide solutions for your business or others. They've been around for a long time and there are thousands of solutions built with VSTO. However, they only run on Office on Windows. You can't run VSTO Add-ins on Mac, online, or mobile platforms.

Office web add-ins use HTML, JavaScript, and additional web technologies to build Office solutions on all platforms. Migrating your existing VSTO Add-in to an Office web add-in is a great way to make your solution available across all platforms.

However there are a number of reasons you may not want to migrate an entire VSTO Add-in codebase from .NET to JavaScript.

- It might be costly or just to big of a task to migrate all at once. You may prefer to build out the Office web add-in in phases while maintaining the current VSTO Add-in.
- Your existing customers might want to continue using your current VSTO Add-in, requiring you to continue supporting the existing codebase with bug fixes and new updates.
- There may be missing functionality in the Office.js library for Office web add-ins that prevents you from completely migrating to an Office web add-in.

If any of the previous factors applies, then the best strategy to be available across all platforms is to maintain both a VSTO Add-in, and an Office web add-in. However it can be costly to maintain two codebases. One way to simplify this is to create a shared code library.

## Shared code library

This guide will walk you through the steps of identifying and sharing common code between your VSTO Add-in and a modern Office web add-in. It uses a very simple VSTO Add-in example for the steps so that you can focus on the skills and techniques you will need for working with your own VSTO Add-ins.

The following diagram shows how the shared code library works for migration. Common code is refactored into a new shared code library. The code can remain written in its original language, such as C# or VB. This means you can continue using the code in the existing VSTO Add-in by creating a project reference. When you create the Office web add-in, it will also use the shared code library by calling into it through REST APIs.

![Diagram of VSTO Add-in and Office web add-in using a shared code library](../images/vsto-migration-shared-code-library.png)

Skills and techniques in this guide:

- Create a shared class library by refactoring code into a .NET class library.
- Create a REST API wrapper using ASP.NET Core for the shared class library.
- Call the REST API from the Office web add-in to access shared code.

## Prerequisites

To set up your development environment:

1. Install [Visual Studio 2019](https://visualstudio.microsoft.com/downloads/).
2. Install the following workloads:
    a. ASP.NET and web development
    b. .NET Core cross-platform development
    c. Office/SharePoint development
    d. Visual Studio Tools for Office (VSTO) Note that this is an Individual component.

You will also need the following:

- An Office 365 account. You can join the [Office 365 Developer Program](https://aka.ms/devprogramsignup) that includes a free 1 year subscription to Office 365.

## The Cell analyzer VSTO Add-in

This guide uses the [VSTO Add-in shared library for Office web add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/vstoshared/Samples/VSTO-shared-code-start) PnP solution. The **/start** folder contains the VSTO Add-in solution that you will migrate. Your goal is to migrate the VSTO Add-in to a modern Office web add-in by sharing code when possible.

> [!NOTE]
> The sample uses C# but you can apply the techniques in this guide to a VSTO Add-in written in any .NET language.

1. Download the [VSTO Add-in shared library for Office web add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/vstoshared/Samples/VSTO-shared-code-start) PnP solution to a working folder on your computer.
2. Start Visual Studio 2019 and open the **/start/Cell-Analyzer.sln** solution.
3. From the **Debug** menu choose **Start Debugging**.

The add-in is a custom task pane for Excel. You can select any cell with text, and then choose the **Show Unicode** button. The add-in will display a list of each character in the text along with its corresponding Unicode number.

![Screenshot of the Cell analyzer VSTO add-in running in Excel](../images/pnp-cell-analyzer-vsto-add-in.png)

## Analyze types of code in the VSTO Add-in

The first technique to apply is to analyze the add-in for which parts of code can be shared. In general there are three types of code to deal with when migrating.

### UI code

UI code interacts with the user. In VSTO UI code works through Windows Forms. Office web add-ins use HTML, CSS, and JavaScript for UI. Because of these differences you cannot share UI code to the Office web add-in. UI will need to be recreated in JavaScript.

### Document code

In VSTO code interacts with the document through .NET objects such as `Microsoft.Office.Interop.Excel.Range`. But Office web add-ins use the Office.js library. Although these are similar, they are not exactly the same. So again, you cannot share document interaction code to the Office web add-in.

### Logic code

Business logic, algorithms, helper functions, and similar code often make up the heart of a VSTO Add-in. This code works independently of the UI and document code to perform analysis, connect to backend services, run calculations, and more. This is the code that can be shared so that you don't have to rewrite it in JavaScript.

Let's examine the VSTO Add-in. In the following code, each section is identified as DOCUMENT, UI, or ALGORITHM code.

```csharp
// *** UI CODE ***
private void btnUnicode_Click(object sender, EventArgs e)
{
    // *** DOCUMENT CODE ***
    Microsoft.Office.Interop.Excel.Range rangeCell;
    rangeCell = Globals.ThisAddIn.Application.ActiveCell;

    string cellValue = "";

    if (null != rangeCell.Value)
    {
        cellValue = rangeCell.Value.ToString();
    }

    // *** ALGORITHM CODE ***
    //convert string to Unicode listing
    string result = "";
    foreach (char c in cellValue)
    {
        int unicode = c;

        result += $"{c}: {unicode}\r\n";
    }
    
    // *** UI CODE ***
    //Output the result
    txtResult.Text = result;
}
```

Using this approach you can see that one section of code can be shared to the Office web add-in. The following code will need to be refactored into a separate class library.

```csharp
// *** ALGORITHM CODE ***
//convert string to Unicode listing
string result = "";
foreach (char c in cellValue)
{
    int unicode = c;

    result += $"{c}: {unicode}\r\n";
}
```

## Create a shared class library

VSTO Add-ins are created in Visual Studio as .NET projects, so we'll reuse .NET as much as possible to keep things simple. Our next technique is to create a class library and refactor shared code into that class library.

1. If you haven't already, start Visual Studio 2019 and open the **\start\Cell-Analyzer.sln** solution.
2. Right-click the solution in **Solution Explorer** and choose **Add > New Project**.
3. In the **Add a new project dialog**, choose **Class Library (.NET Framework)** and choose **Next**.
    > [!NOTE]
    > Don't use the .NET Core class library because it will not work with your VSTO project.
5. In the **Configure your new project** dialog, set the following fields.
    - Set the  **Project name** to **CellAnalyzerSharedLibrary**.
    - Leave the **Location** at it's default value.
    - Set the **Framework** to **4.7.2**.
6. Choose **Create**.
7. After the project is created, rename the **Class1.cs** file to **CellOperations.cs**. You will be prompted to rename the class as well. Do that so the class name matches the file name.
8. Add the following code to the `CellOperations` class to create a method named `GetUnicodeFromText`.

```csharp
public class CellOperations
{
    static public string GetUnicodeFromText(string value)
    {
        string result = "";
        foreach (char c in value)
        {
            int unicode = c;

            result += $"{c}: {unicode}\r\n";
        }
        return result;
    }
}
```

### Use the shared class library in the VSTO Add-in

Now you need to update the VSTO Add-in to use the class library. This is important that both the VSTO Add-in and Office web add-in use the same shared class library so that future bug fixes or features are made in one location.

1. In **Solution Explorer** expand the **Cell-Analyzer** project, right-click the **CellAnalyzerPane.cs** file, and choose **View Code**.
2. In the `btnUnicode_Click` method, delete the following lines of code.

```csharp
//Convert to Unicode listing
    string result = "";
    foreach (char c in cellValue)
    {
        int unicode = c;
        result += $"{c}: {unicode}\r\n";
    }
```

3. Update the line of code under the `//Output the result` comment to read as follows:

```csharp
//Output the result
txtResult.Text = CellAnalyzerSharedLibrary.CellOperations.GetUnicodeFromText(cellValue);
```

6. From the **Debug** menu, choose **Start Debugging**. The custom task pane should still work as expected. Enter some text in a cell, and then test that you can convert it to a Unicode list with the add-in.

## Create a REST API wrapper

The VSTO Add-in can use the shared class library directly since they are both .NET projects. However the Office web add-in won't be able to use .NET since it uses JavaScript. Next you will need to create a REST API wrapper. This enables the Office web add-in to call a REST API, which then passes the call along to the shared class library.

1. If you haven't already, start Visual Studio 2019 and open the **\start\Cell-Analyzer.sln** solution.
2. Right-click the solution in **Solution Explorer** and choose **Add > New Project**.
3. In the **Add a new project dialog**, choose **ASP.NET Core Web Application** and choose **Next**.
4. In the **Configure your new project** dialog, set the following fields.
    - Set the  **Project name** to **CellAnalyzerRESTAPI**.
    - Leave the **Location** at it's default value.
5. Choose **Create**.
6. After the project is created, expand the **CellAnalyzerRESTAPI** project in **Solution Explorer**.
7. Right-click **Dependencies** and choose **Add Reference**.
8. Select **CellAnalyzerSharedLibrary** and choose **OK**.
9. Right-click the **Controllers** folder, and choose **Add > Controller**.
10. In the **Add New Scaffolded Item** dialog, choose **API Controller - Empty** and then **Add**.
11. In the **Add Empty API Controller** dialog, name the controller **AnalyzeUnicodeController** and then choose **Add**.
12. Open the **AnalyzeUnicodeController.cs** file and add the following code as a method to the `AnalyzeUnicodeController` class.

```csharp
[HttpGet]
public ActionResult<string> AnalyzeUnicode(string value)
{
    if (value == null)
    {
        return BadRequest();
    }
    return CellAnalyzerSharedLibrary.CellOperations.GetUnicodeFromText(value);
}
```

11. Right-click the CellAnalyzerRESTAPI project and choose **Set as Startup Project**.
12. From the **Debug** menu, choose **Start Debugging**.
13. A browser will launch. Enter the following URL to test that the REST API is working: **https://localhost:44323/api/analyzeunicode?value=test**. You should see a string returned with Unicode values for each character.

## Create the Office web add-in

When you create the Office web add-in, it will need to call the REST API. First you need to get the port number of the REST API server and save it.

### Save the SSL port number

1. If you haven't already, start Visual Studio 2019 and open the **\start\Cell-Analyzer.sln** solution.
2. In the **CellAnalyzerRESTAPI** project, expand **Properties**, and open the **launchSettings.json** file.
3. Find the line of code with the **sslPort** value, copy the port number and save it somewhere.

### Add the Office web add-in project

To keep things simple, you'll keep all the code in one solution. You'll add the Office web add-in project to the existing Visual Studio solution. However, if you are familiar with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) and Visual Studio Code you can also run `yo office` to build the project. The steps are very similar.

1. If you haven't already, start Visual Studio 2019 and open the **\start\Cell-Analyzer.sln** solution.
2. Right-click the solution in **Solution Explorer** and choose **Add > New Project**.
3. In the **Add a new project dialog**, choose **Excel Web Add-in** and choose **Next**.
5. In the **Configure your new project** dialog, set the following fields.
    - Set the  **Project name** to **CellAnalyzerWebAddin**.
    - Leave the **Location** at it's default value.
    - Set the **Framework** to **4.7.2** or later.
2. Choose **Create**.
3. In the **Choose the add-in type** dialog, select **Add new functionalities to Excel** and choose **Finish**.

### Add UI and functionality to the Office web add-in

1. Open the **Home.html** file and replace the `<body>` contents with the following HTML.

```html
<button id="btnShowUnicode" onclick="showUnicode()">Show Unicode</button>
<p>Result:</p>
<div id="txtResult"></div>
```

2. Open the **Home.js** file and replace the entire contents with the following code. Substitute the **sslPort** number you saved previously from the **launchSettings.json** file.

```js
(function () {
    "use strict";
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
        });
    };
})();

function showUnicode() {
    Excel.run(function (ctx) {
        const range = ctx.workbook.getSelectedRange();
        range.load("values");
        return ctx.sync(range).then(function (range) {
            const url = "https://localhost:<ssl port number>/api/analyzeunicode?value=" + range.values[0][0];
            $.ajax({
                type: "GET",
                url: url,
                success: function (data) {
                    let htmlData = data.replace(/\r\n/g, '<br>');
                    $("#txtResult").html(htmlData);
                }
            });
        });
    });
}
```

It's worth noting that in the previous code the returned string will be processed to replace carriage return line feeds with `<br>` HTML tags. You may occasionally run into situations where a return value that works perfectly fine for .NET in the VSTO Add-in will need to be adjusted on the Office web add-in side to work as expected. In this case the REST API and shared class library are only concerned with returning the string. The `showUnicode()` method is responsible for formatting return values correctly for presentation.

### Allow CORS from the Office web add-in

The Office.js library requires CORS on outgoing calls, such as the one made from the `ajax` call to the REST API server. Use the following steps to allow calls from the Office web add-in to the REST API.

1. In **Solution Explorer** select the **CellAnalyzerWebAddinWeb** project.
2. From the **View** menu, choose **Properties Window** (if the window is not already displayed).
3. In the properties window, copy the value of the **SSL URL** and save it somewhere. This is the URL that you need to allow through CORS.
4. In the **CellAnalyzerRESTAPI** project open the **Startup.cs** file.
5. Add the following code to the top of the `ConfigureServices` method. Be sure to substitute the URL SSL you copied previously for the `builder.WithOrigins` call.

```csharp
services.AddCors(options =>
    {
        options.AddPolicy(MyAllowSpecificOrigins,
        builder =>
        {
            builder.WithOrigins("<your URL SSL>")
            .AllowAnyMethod()
            .AllowAnyHeader()
            .AllowCredentials();
        });
    });
```

6. Add the following field to the `Startup` class:

```csharp
readonly string MyAllowSpecificOrigins = "_myAllowSpecificOrigins";
```

7. Add the following code to the `configure` method.

```csharp
app.UseCors(MyAllowSpecificOrigins);
```

When done, your `Startup` class should look similar to the following code:

```csharp
 public class Startup
{
    public Startup(IConfiguration configuration)
    {
        Configuration = configuration;
    }
    readonly string MyAllowSpecificOrigins = "_myAllowSpecificOrigins";
    public IConfiguration Configuration { get; }

    // This method gets called by the runtime. Use this method to add services to the container.
    public void ConfigureServices(IServiceCollection services)
    {
        services.AddCors(options =>
        {
            options.AddPolicy(MyAllowSpecificOrigins,
            builder =>
            {
                builder.WithOrigins("https://localhost:44397")
                .AllowAnyMethod()
                .AllowAnyHeader()
                .AllowCredentials();
            });
        });
        services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_2);
    }

    // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
    public void Configure(IApplicationBuilder app, IHostingEnvironment env)
    {
        if (env.IsDevelopment())
        {
            app.UseDeveloperExceptionPage();
        }
        else
        {
            // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
            app.UseHsts();
        }

        app.UseCors(MyAllowSpecificOrigins);
        app.UseHttpsRedirection();
        app.UseMvc();
    }
}
```
