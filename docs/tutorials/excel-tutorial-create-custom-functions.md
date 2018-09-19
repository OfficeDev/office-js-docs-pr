---
title: Excel custom function tutorial
description: In this tutorial, you’ll create an add-in, a custom function in Excel, which can perform calculations, request web data, or stream web data.
ms.date: 09/05/2018
ms.topic: tutorial
#Customer intent: As an add-in developer, I want to create a custom function in Excel to increase productivity. 
---
# Create a streaming Excel custom function

## Introduction

Custom Functions give you the power to automate processes in Excel. You can use them for something as simple as creating your own custom calculation, similar to `=SUM()` or for more complex tasks, like streaming data from the web in real-time right into your worksheet.

In this tutorial, you will learn how to:
> [!div class="checklist"]
> * Create a custom function project using yo office
> * Build a custom function which performs a simple calculation
> * Use custom functions to request data from the web
> * Create a custom function which streams real-time data from the web

## Prerequisites

* [Node.js and npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads) (or another Git client)
* [Yeoman](http://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office)
* Office 2016 for Windows and [join the Office Insider program](https://products.office.com/office-insider). You must have Office build number 10827 or later.

## Create your add-in project

You’ll begin this tutorial by using the Yo Office Yeoman generator, which will automatically populate the files you need for your project.

    1. In your command line interface, create a scaffold of your project (by default, this should be in your C:\Users\YourUserName folder).

    ```bash
    yo office
    ```

    ![Yo Office bash prompts for custom functions](../images/yo-office-excel-cfs-stock-ticker.png)
    Answer the prompts as directed below:

    * Choose a project type: `Excel Custom Functions Add-in project (September 2018 Preview Refresh: Requires the Insider channel for Excel)`
    * What do you want to name your add-in? `stock-ticker`

    For this tutorial, choose Javascript as the language you would like to use to build your add-in.

    After you complete the wizard, the generator will create the project files and install supporting Node components.

    2. Next, start a localhost instance by running one of the below commands in your command line interface.

    If you are developing using the desktop version of Excel, use:

    ```bash
    npm start
    ```

    If you are using Excel Online, use:

    ```bash
    npm start-web
    ```

    3. You will also need to register your custom-functions add-in. In Excel, select **Insert > My Add-ins > Insert an Add-in**. This will bring up a list of available add-ins. Under "Developer Add-ins" you will see your add-in, under the name "Excel Custom Function". Select it to register it.

    Select **Insert > Add-ins**. Choose **Manage My Add-ins** and select **Upload My Add-in**. Click "Browse..." for your manifest file (`C:\Users\YourName\stock-ticker\manifest.xml`), then click Open, select **Upload**.

    4. Finally, change the script tag to point to the right custom functions source. Open up your add-in project in your favorite code editor. In **index.html** in the root folder, delete and replace the script tag immediately following the <title> tags with the code below:

    ```js
    <script src="https://unpkg.com/@microsoft/office-js@1.1.9-adhoc.22/dist/custom-functions-runtime.js" type="text/javascript"></script>
    ```

## Try out a basic computational custom function

Now the custom functions in your file will be loaded and ready to use. There are several pre-built functions for you in the Yo Office project. All are attached to a namespace called CONTOSO which is defined in the XML manifest file. Once you start typing =CONTOSO in a cell, the list of available functions will appear.

Let's call `=CONTOSO.ADD42()`. This function adds 42 to any two numbers you provide as arguments. In any cell, type `=CONTOSO.ADD42(1,2)`. It should deliver the answer 45.

## Create a custom function

What if you wanted a function which could fetch and display the price of Microsoft stock that day? Custom functions are designed so you can easily make requests for data from the web asynchronously.

You’ll be adding a new function, called `=CONTOSO.STOCKPRICE`, to the **customfunctions.j** file. The function will take in the name of a stock ticker, such as "MSFT", and return the price of that stock. You'll leverage the IEX Trading API, which is free and does not require authentication.

    1. Open your code editor of choice and navigate to the stock-ticker project folder. 
    2. Copy and paste the function below and add it to **customfunctions.js**.

    You'll notice in this code that your asynchronous function returns a JavaScript Promise with the data from the IEX Trading API. Asynchronous custom functions require you to either return a new Promise or use JavaScript's async/await syntax.

    ```js
    function STOCKPRICE(ticker) {
        return new Promise(
            function(resolve) {
                let xhr = new XMLHttpRequest();
                let url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price"
                //add handler for xhr
                xhr.onreadystatechange = function() {
                    if (xhr.readyState == XMLHttpRequest.DONE) {
                    //return result back to Excel
                    resolve(xhr.responseText);
                    }
                }
                //make request
                xhr.open('GET', url, true);
                xhr.send();
        });
    }
    ```

    3. In order for Excel to properly run this function, you must add some metadata to the **./config/customfunctions.json** file.

    You'll notice that this JSON file describes the function, listing the types and dimensionality of the results and parameters.

    ```json
    {
        "name": "STOCKPRICE",
        "description": "Multiplies number by 105",
        "helpUrl": "http://dev.office.com",
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
    }
    ```

    4. You will need to re-register this change once you have saved the file. In Excel, select **Insert > Add-ins > My Add-ins**. This will bring up a list of available add-ins. Under “Developer Add-ins" you will see your add-in, under the name “Excel Custom Function.” Select it to register it.

    5. In cell B1, run the function `=CONTOSO.STOCKPRICE("MSFT")`. It should show you the stock price for one share of Microsoft stock right now.

## Create a streaming asynchronous custom function

The previous function returned the stock price for Microsoft at a particular moment in time, but stock prices are always changing. With custom functions, it is possible to “stream” data from an API to get updates on stock prices in real time.

To do this, you’ll create a new function, `=CONTOSO.STOCKPRICESTREAM`. It makes a request for updated data every 1000 milliseconds. When a call is made, you may see `#GETTING_DATA` appear in a cell. Once a value is returned, this notification should disappear.

    1. Copy and paste the code below into **customfunctions.js**.

    ```js
        function STOCKPRICESTREAM(ticker, caller){
        let result = 0;

        //return every second
        setInterval(function(){
        let xhr = new XMLHttpRequest();
        let url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";

        //add handler for xhr
        xhr.onreadystatechange = function() {
            if (xhr.readyState == XMLHttpRequest.DONE) {
                //return result back to Excel
                caller.setResult(xhr.responseText);
            }
        }

        //make request
        xhr.open('GET', url, true);
        xhr.send();
            }, 1000);
        }
    ```

    3. Copy and paste the code below into to the **./config/customfunctions.json** file.

     You'll notice that this JSON file is very similar to the previous function's JSON file, but that a new section has been added for "options." Because this function is streaming, you must specify this as true in the JSON.

    ```json
    {
        "name": "STOCKPRICESTREAM",
        "description": "Streams real time stock price",
        "helpUrl": "http://dev.office.com",
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
            "stream": true
        }
    }
    ```

    4. Re-register this change once you have saved the file. In Excel, select **Insert > Add-ins > My Add-ins**. This will bring up a list of available add-ins. Under “Developer Add-ins" you will see your add-in, under the name “Excel Custom Function.” Select it to register it.

    5. In cell C1, run the function `=CONTOSO.STOCKPRICESTREAM("MSFT")`. You should see the price of Microsoft stock - which will update in real time right in your workbook.

## Next steps

You’ve completed the custom functions add-in tutorial. To learn more about custom functions, read [this overview article](../excel/custom-functions-overview.md).
> [!div class="nextstepaction"]
> [Overview of Custom Functions](../excel/custom-functions-overview.md)
