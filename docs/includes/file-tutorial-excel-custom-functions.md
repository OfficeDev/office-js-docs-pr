# Tutorial: Create custom functions in Excel

## Introduction

Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`. You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.

In this tutorial, you will:
> [!div class="checklist"]
> * Create a custom functions project by using the Yo Office generator
> * Use a prebuilt custom function to perform a simple calculation
> * Create a custom function that requests data from the web
> * Create a custom function that streams real-time data from the web

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## Prerequisites

* [Node.js](https://nodejs.org/en/) (version 8.0.0 or later)

* [Git Bash](https://git-scm.com/downloads) (or another Git client)

* The latest version of [Yeoman](https://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:

    ```bash
    npm install -g yo generator-office
    ```

* Excel for Windows (version 1810 or later) or Excel Online

* Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")

## Create a custom functions project

 You’ll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project. If you have previously installed yo office, make sure to update your package to pull the latest from npm. You can do this by running `npm install -g yo generator-office`.

1. Run the following command and then answer the prompts as follows.

    ```bash
    yo office
    ```

    * Choose a project type: `Excel Custom Functions Add-in project (...)`
    * Choose a script type: `JavaScript`
    * What do you want to name your add-in? `stock-ticker`

    ![Yo Office bash prompts for custom functions](../images/12-10-fork-cf-pic.jpg)

    After you complete the wizard, the generator will create the project files and install supporting Node components. The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.

2. Navigate to the project folder.

    ```
    cd stock-ticker
    ```

3. Trust the self-signed certificate that is needed to run this project. For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).  

4. Build the project.

    ```
    npm run build-dev
    ```
    
    After running this, you will see a readout in your command prompt about the build process. Note that you will be using this command later to re-build whenever you make changes to your functions files.

5. Start the local web server, which runs in Node.js. 

    * If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:

When running this command, you will see instructions to run `npm run start` however, run the below command instead: 
         ```
         npm run start-desktop
        ```
        After running this command, your command prompt will show details about what has been done, another npm window will open showing the details of the build, and Excel will start with your add-in loaded. If you add-in does not load, check that you have completed step 3 properly.  

    * If you'll be using Excel Online to test your custom functions, run the following command to start the local web server: 

        ```
        npm run start-web
        ```

         After running this command, another window will open showing you the details of the build. To use your functions, open a new workbook in Office Online. 

## Try out a prebuilt custom function

The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/functions/functions.js** file. The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.

In your Excel workbook, try out the `ADD` custom function by completing the following steps in Excel:

1. Within a cell, type **=CONTOSO**. Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.

2. Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.

The `ADD` custom function computes the sum of the two numbers that you specify as input parameters. Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.

## Create a custom function that requests data from the web

What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet? Custom functions are designed so that you can easily request data from the web asynchronously.

Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker symbol (e.g., **MSFT**) and returns the price of that stock. This custom function uses the IEX Trading API, which is free and does not require authentication.

1. In the **stock-ticker** project that the Yo Office generator created, find the file **src/functions/functions.js** and open it in your code editor.

2. In **functions.js**, locate the `increment` function and add the following code immediately after that function.

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

3. In **functions.js**, locate the line`CustomFunctionMappings.LOG = logMessage;`, add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```
    
4. Before Excel can make this new function available, you must specify metadata to describe the function to Excel. Open the **src/functions/functions.json** file. Add the following JSON object to the 'functions' array and save the file.


    This JSON describes the `stockPrice` function.

    ```json
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
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

5. You must re-register the add-in in Excel in order for the new function to be available to end-users. Complete the following steps for the platform that you're using in this tutorial.

    * If you're using Excel for Windows:

        1. Close Excel and then reopen Excel.

        2. In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.
            ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)

        1. In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.
            ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)

    * If you're using Excel Online: 

        1. In Excel Online, choose the **Insert** tab and then choose **Add-ins**.
            ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)

        2. Choose **Manage My Add-ins** and select **Upload My Add-in**. 

        3. Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created. 

        4. Select the file **manifest.xml** and choose **Open**, then choose **Upload**.
1. [ADD A STEP ABOUT RE_BUILDING]

5. Now, let's try out the new function. In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter. You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.

## Create a streaming asynchronous custom function

The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing. Let's create a custom function that streams data from an API to get real-time updates on a stock price.

Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed). While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called. When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.

1. In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/functions/functions.js** and save the file.

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

2. Before Excel can make this new function available to end-users, specify metadata that describes this function. In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **src/functions/functions.json** file and save the file.

    This JSON describes the `stockPriceStream` function. For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.

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

3. You must reregister the add-in in Excel in order for the new function to be available to end-users. Complete the following steps for the platform that you're using in this tutorial.

    * If you're using Excel for Windows:

        1. Close Excel and then reopen Excel.
        
        2. In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.
            ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)

        3. In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.
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

## Legal information

Data provided free by [IEX](https://iextrading.com/developer/). View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/). Microsoft's use of the IEX API in this tutorial is for educational purposes only.
