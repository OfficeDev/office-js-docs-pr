---
title: Excel custom functions tutorial (preview)
description: In this tutorial, you’ll create an Excel add-in that contains a custom function that can perform calculations, request web data, or stream web data.
ms.date: 01/07/2019
ms.topic: tutorial
#Customer intent: As an add-in developer, I want to create custom functions in Excel to increase user productivity. 
---

# Tutorial: Create custom functions in Excel (preview)

Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`. You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.

In this tutorial, you will:
> [!div class="checklist"]
> * Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). 
> * Use a prebuilt custom function to perform a simple calculation.
> * Create a custom function that gets data from the web.
> * Create a custom function that streams real-time data from the web.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## Prerequisites

* [Node.js](https://nodejs.org/en/) (version 8.0.0 or later)

* [Git Bash](https://git-scm.com/downloads) (or another Git client)

* The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > Even if you have previously installed the Yeoman generator, we recommend updating your package to the latest version from npm.

* Excel for Windows (64-bit version 1810 or later) or Excel Online

* Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")

## Create a custom functions project

 To start, you'll create the code project to build your custom function add-in. The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some initial custom functions that you can try out.

1. Run the following command and then answer the prompts as follows.
    
    ```
    yo office
    ```
    
    * Choose a project type: `Excel Custom Functions Add-in project (...)`
    * Choose a script type: `JavaScript`
    * What do you want to name your add-in? `stock-ticker`
    
    ![Yeoman generator for Office Add-ins prompts for custom functions](../images/12-10-fork-cf-pic.jpg)
    
    The Yeoman generator creates the project files and installs supporting Node.js components.

2. Go to the project folder.
    
    ```
    cd stock-ticker
    ```

3. Trust the self-signed certificate that is needed to run this project. For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).  

4. Build the project.
    
    ```
    npm run build
    ```

5. Start the local web server, which runs in Node.js. You can try out the custom function add-in in Excel for Windows, or Excel online.

# [Excel for Windows](#tab/excel-windows)

Run the following command.

```
npm run start
```

This command starts the web server, and sideloads your custom function add-in into Excel for Windows.

> [!NOTE]
> If your add-in does not load, check that you have completed step 3 properly.

# [Excel online](#tab/excel-online)

Run the following command.

```
npm run start-web
```

This command starts the web server. Use the following steps to sideload your add-in.

<ol type="a">
   <li>In Excel Online, choose the <strong>Insert</strong> tab and then choose <strong>Add-ins</strong>.<br/>
   <img src="../images/excel-cf-online-register-add-in-1.png" alt="Insert ribbon in Excel Online with the My Add-ins icon highlighted"></li>
   <li>Choose <strong>Manage My Add-ins</strong> and select <strong>Upload My Add-in</strong>.</li> 
   <li>Choose <strong>Browse...</strong> and navigate to the root directory of the project that the Yeoman generator created.</li> 
   <li>Select the file <strong>manifest.xml</strong> and choose <strong>Open</strong>, then choose <strong>Upload</strong>.</li>
</ol>

> [!NOTE]
> If your add-in does not load, check that you have completed step 3 properly.

--- 
    
## Try out a prebuilt custom function

The custom functions project that you created alrady has two prebuilt custom functions named ADD and INCREMENT. The code for these prebuilt functions is in the  **src/customfunctions.js** file. The **./manifest.xml** file specifies that all custom functions belong to the `CONTOSO` namespace. You'll use the CONTOSO namespace to access the custom functions in Excel.

Next you'll try out the `ADD` custom function by completing the following steps:

1. In Excel, go to any cell and enter `=CONTOSO`. Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.

2. Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.

The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.

## Create a custom function that requests data from the web

Integrating data from the Web is a great way to extend Excel through custom functions. Next you’ll create a custom function named `stockPrice` that gets a stock quote from a Web API and returns the result to the cell of a worksheet. You’ll use the IEX Trading API, which is free and does not require authentication.

1. In the **stock-ticker** project, find the file **src/customfunctions.js** and open it in your code editor.

2. In **customfunctions.js**, locate the `increment` function and add the following code immediately after that function.

    ```js
    function stockPrice(ticker) {
        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        return fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                return parseFloat(text);
            });

        // Note: in case of an error, the returned rejected Promise
        //    will be bubbled up to Excel to indicate an error.
    }

3. In **customfunctions.js**, locate the line`CustomFunctions.associate("INCREMENT", increment);`. Add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctions.associate("stockprice", stockprice);
    ```

    The `CustomFunctions.associate` code associates the id of the function with the function address of `increment` in JavaScript so that Excel can call your function.

    Before Excel can use your custom function, you need to describe it using metadata. You need to define the `id` used in the `associate` method previously, along with some other metadata.


4. Open the **config/customfunctions.json** file. Add the following JSON object to the 'functions' array and save the file.

    ```JSON
    {
        "id": "STOCKPRICE",
        "name": "STOCKPRICE",
        "description": "Fetches current stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

    This JSON describes the `stockPrice` function, its parameters, and the type of result it returns.

5. Re-register the add-in in Excel so that the new function is available. 

# [Excel for Windows](#tab/excel-windows)

1. Close Excel and then reopen Excel.
2. In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.
    ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)
3. In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.
    ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)

# [Excel online](#tab/excel-online)

1. In Excel Online, choose the **Insert** tab and then choose **Add-ins**.
    ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)
2. Choose **Manage My Add-ins** and select **Upload My Add-in**. 
3. Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created. 
4. Select the file **manifest.xml** and choose **Open**, then choose **Upload**.

--- 

<ol start="6">
<li> Try out the new function. In cell <strong>B1</strong>, type the text <strong>=CONTOSO.STOCKPRICE("MSFT")</strong> and press enter. You should see that the result in cell <strong>B1</strong> is the current stock price for one share of Microsoft stock.</li>
</ol>

## Create a streaming asynchronous custom function

The `stockPrice` function returns the price of a stock at a specific moment in time, but stock prices are always changing. 
Next you’ll create a custom function named `stockPriceStream` that gets the price of a stock every 1000 milliseconds.

1. In the **stock-ticker** project, add the following code to **src/customfunctions.js** and save the file.

    ```js
    function stockPriceStream(ticker, handler) {
        var updateFrequency = 1000 /* milliseconds*/;
        var isPending = false;

        var timer = setInterval(function() {
            // If there is already a pending request, skip this iteration:
            if (isPending) {
                return;
            }

            var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
            isPending = true;

            fetch(url)
                .then(function(response) {
                    return response.text();
                })
                .then(function(text) {
                    handler.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    handler.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        handler.onCanceled = () => {
            clearInterval(timer);
        };
    }
    
    CustomFunctions.associate("stockpricestream", stockpricestream);
    ```
    
    Before Excel can use your custom function, you need to describe it using metadata .
    
2. In the **stock-ticker** project add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.
    
    ```json
    { 
        "id": "STOCKPRICESTREAM",
        "name": "STOCKPRICESTREAM",
        "description": "Streams real time stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ],
        "options": {
            "stream": true,
            "cancelable": true
        }
    }
    ```

    This JSON describes the `stockPriceStream` function. For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.

3. Re-register the add-in in Excel so that the new function is available.

# [Excel for Windows](#tab/excel-windows)

1. Close Excel and then reopen Excel.
2. In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.
    ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)
3. In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.
    ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)

# [Excel online](#tab/excel-online)

1. In Excel Online, choose the **Insert** tab and then choose **Add-ins**.
    ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)
2. Choose **Manage My Add-ins** and select **Upload My Add-in**.
3. Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.
4. Select the file **manifest.xml** and choose **Open**, then choose **Upload**.

--- 

<ol start="4">
<li>Try out the new function. In cell <strong>C1</strong>, type the text <strong>=CONTOSO.STOCKPRICESTREAM("MSFT")</strong> and press enter. Provided that the stock market is open, you should see that the result in cell <strong>C1</strong> is constantly updated to reflect the real-time price for one share of Microsoft stock.</li>
</ol>

## Next steps

Congratulations! You've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web. To learn more about custom functions in Excel, continue to the following article:

> [!div class="nextstepaction"]
> [Create custom functions in Excel](../excel/custom-functions-overview.md)

### Legal information

Data provided free by [IEX](https://iextrading.com/developer/). View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/). Microsoft's use of the IEX API in this tutorial is for educational purposes only.


