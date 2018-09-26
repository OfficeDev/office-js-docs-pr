---
title: Excel custom functions tutorial
description: In this tutorial, you’ll create an Excel add-in that contains a custom function that can perform calculations, request web data, or stream web data.
ms.date: 09/20/2018
ms.topic: tutorial
#Customer intent: As an add-in developer, I want to create a custom function in Excel to increase productivity. 
---

# Tutorial: Create custom functions in Excel

## Introduction

Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any other native function in Excel, such as `SUM()`. You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.

In this tutorial, you will:
> [!div class="checklist"]
> * Create a custom functions project by using the Yo Office generator
> * Use a prebuilt custom function to perform a simple calculation
> * Create a custom function that requests data from the web
> * Create a custom function that streams real-time data from the web

## Prerequisites

* [Node.js and npm](https://nodejs.org/en/)

* [Git Bash](https://git-scm.com/downloads) (or another Git client)

* The latest version of [Yeoman](http://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:

    ```bash
    npm install -g yo generator-office
    ```

* Excel for Windows (build number 10827 or later) or Excel Online

* [Join the Office Insider program](https://products.office.com/office-insider). 
    > [!NOTE]
    > Currently, you must join the Office Insider program in order to have access to custom functions. Custom functions are disabled across all Office builds unless you are a member of the Office Insider program.

## Create a custom functions project

You’ll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project.

1. Run the following command and then answer the prompts as follows.

    ```bash
    yo office
    ```

    * Choose a project type: `Excel Custom Functions Add-in project (...)`
    * Choose a script type: `JavaScript`
    * What do you want to name your add-in? `stock-ticker`

    ![Yo Office bash prompts for custom functions](../images/yo-office-cfs-stock-ticker-2.png)

    After you complete the wizard, the generator will create the project files and install supporting Node components.

2. Navigate to the project folder.

    ```bash
    cd stock-ticker
    ```

3. Start the local web server.

    * If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:

        ```bash
        npm start
        ```

    * If you'll be using Excel Online to test your custom functions, run the following command to start the local web server: 

        ```bash
        npm run start-web
        ```

4. Register your custom functions add-in in Excel by completing steps for the platform that you'll be using in this tutorial.

    * If you'll be using Excel for Windows to test your custom functions:

        1. In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.
            ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)

        2. In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.
            ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)

    * If you'll be using Excel Online to test your custom functions: 

        1. In Excel Online, choose the **Insert** tab and then choose **Add-ins**.
            ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)

        2. Choose **Manage My Add-ins** and select **Upload My Add-in**. 

        3. Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created. 

        4. Select the file **manifest.xml** and choose **Open**, then choose **Upload**.

## Try out a prebuilt custom function

The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/customfunction.js** file. The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.

At this point, the prebuilt custom functions in your project are loaded and available within Excel. Try out the `ADD` custom function by completing the following steps in Excel:

1. Within a cell, type **=CONTOSO**. Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.

2. Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by specifying the following value in the cell and pressing enter:

    ```
    =CONTOSO.ADD(10,200)
    ```

The `ADD` custom function computes the sum of the two numbers that you specify as input parameters. Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.

## Create a custom function that requests data from the web

What if you needed a function that could request the price of a stock from the web and display the result in the cell of a worksheet? Custom functions are designed so that you can easily request data from the web asynchronously.

Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock. This custom function uses the IEX Trading API, which is free and does not require authentication.

1. In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.

2. Add the following code to **customfunctions.js** and save the file.

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

    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

3. Before Excel can make this new function available to end-users, you must specify metadata that describes this function. In the **stock-ticker** project that the Yo Office generator created, find the file **config/customfunctions.json** and open it in your code editor. Add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.

    This JSON describes the `stockPrice` function.

    ```json
    {
        "id": "STOCKPRICE",
        "name": "STOCKPRICE",
        "description": "Fetches current stock price",
        "helpUrl": "http://yourhelpurl.com",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

4. You must reregister the add-in in Excel in order for the new function to be available to end-users. Reregister your custom functions add-in in Excel by completing the following steps:

    * If you're using Excel for Windows:

        1. In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.
            ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)

        2. In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.
            ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)

    * If you're using Excel Online: 

        1. In Excel Online, choose the **Insert** tab and then choose **Add-ins**.
            ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)

        2. Choose **Manage My Add-ins** and select **Upload My Add-in**. 

        3. Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created. 

        4. Select the file **manifest.xml** and choose **Open**, then choose **Upload**.

5. Now, let's try out the new function. In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter. You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.

## Create a streaming asynchronous custom function

The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing. Let's create a custom function that streams data from an API to get real-time updates on a stock price.

Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed). While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called. When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.

1. In the **stock-ticker** project that the Yo Office generator created, add the following code to **customfunctions.js** and save the file.

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
    CustomFunctionMappings.STOCKPRICESTREAM = stockPriceStream;
    ```

2. Before Excel can make this new function available to end-users, you must specify metadata that describes this function. In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.

    This JSON describes the `stockPriceStream` function. Notice that the `stream` property within the `options` object is set to `true`, to indicate that this is a streaming function.

    ```json
    { 
        "id": "STOCKPRICESTREAM",
        "name": "STOCKPRICESTREAM",
        "description": "Streams real time stock price",
        "helpUrl": "http://yourhelpurl.com",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock ticker name",
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

3. You must reregister the add-in in Excel in order for the new function to be available to end-users. Reregister your custom functions add-in in Excel by completing the following steps:

    * If you're using Excel for Windows:

        1. In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.
            ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)

        2. In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.
            ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)

    * If you're using Excel Online: 

        1. In Excel Online, choose the **Insert** tab and then choose **Add-ins**.
            ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)

        2. Choose **Manage My Add-ins** and select **Upload My Add-in**. 

        3. Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created. 

        4. Select the file **manifest.xml** and choose **Open**, then choose **Upload**.

4. Now, let's try out the new function. In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter. Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.

## Next steps

In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web. To learn more about custom functions in Excel, continue to the following article: 

> [!div class="nextstepaction"]
> [Create custom functions in Excel](../excel/custom-functions-overview.md)
