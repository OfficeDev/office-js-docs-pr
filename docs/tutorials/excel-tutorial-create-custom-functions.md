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
- Node.js[TODO]
- Latest version of [TODO]
- Gitbash[TODO]
- To use this tutorial, you must have Office 2016 for Windows and [join the Office Insider program](https://products.office.com/office-insider). You must have Office build number 8711 or later.

## Create your add-in project
You’ll begin this tutorial by using scaffolding tool Yeoman, which will automatically populate the files you need for your project.

1. In your command line interface, create a scaffold of your project.  
    
    ```bash
    yo office
    ```
    
    ![Yo Office bash prompts for custom functions](../images/yo-office-excel-cfs-stock-ticker.png)
    
    Answer the prompts as directed below:  
    - Choose a project type: `Excel Custom Funtions Add-in project (Preview: Requires the Insider channel for Excel)`
    - What do you want to name your add-in? `stock-ticker`
    
    After you complete the wizard, the generator will create the project files and install supporting Node components.  

2. In the root folder for your project, follow [these instructions](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to install the self-signed certificates.

3. Start a localhost instance by running the below in the command line:

    ```bash
    npm  start
    ```

4. In a separate git bash window, navigate to the root folder of your project. Next, sideload your project so it can be used in Excel by running the below in the command line:

    ```bash
    npm run sideload
    ```

5. Register your custom-functions add-in. In Excel, select **Insert > My Add-ins > Insert an Add-in**. This will bring up a list of available add-ins. Under "Developer Add-ins" you will see your add-in, under the name "Excel Custom Function". Select it to register it.

6. In `index.html` in the root folder, delete and replace the script tag immediately following the `<title>` tags with the code below:

    ```js
    <script src="https://unpkg.com/@microsoft/office-js@1.1.9-adhoc.22/dist/custom-functions-runtime.js" type="text/javascript"></script>
    ```

## Create a basic computational custom function
Now that your project is set up, you’ll learn how to create a function to calculate the value of multiple shares of Microsoft stock via a simple math problem.  

For the sake of this exercise, assume that Microsoft’s current stock price is $105/share. You’ll create a function which takes in the number of shares and multiples that number by 105. In the `src` folder, you will see there is a file called `customfunctions.js`. There are some pre-fabricated functions in this file like ADD42, which you will be adding to. Copy and paste the below code into `customfunctions.js`.

```js
function STOCKMULTIPLES(num1) {
    return num1 * 105;  
}
```

In order for Excel to properly run this function, you must add metadata describing the function to the `./config/customfunctions.json` file. Add the following JSON:  

```json
{
    "name": "STOCKMULTIPLES",
    "description": "Multiplies number by 105",
    "helpUrl": "http://dev.office.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
        },  
    "parameters": [
        {
            "name": "num1",
            "description": "variable to multiply by 105",
            "type": "number",
            "dimensionality": "scalar"
        }
    ]
}
```

Additionally, you will need to re-register this change once you have saved the file. In Excel, select **Insert > Add-ins > My Add-ins**. This will bring up a list of available add-ins. Under “Developer Add-ins" you will see your add-in, under the name “Excel Custom Function.” Select it to register it.  

In cell A1, run `=CONTOSO.STOCKMULTIPLES(5)`. This will tell us the value of 5 shares of Microsoft stock: $525.

## Create an asynchronous custom function
What if you wanted a function which could fetch and display the price of Microsoft stock that day? Custom functions are designed so you can easily make requests for data from the web asynchronously.
  
You’ll be adding a new function, called `=CONTOSO.STOCKPRICE`, to the `customfunctions.js` file.  The function will take in the name of a stock ticker, such as "MSFT", and return the price of that stock.  

Copy and paste the function below and add it to `customfunctions.js`.  

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

Again, in order for Excel to properly run this function, you must add some metadata to the `./config/customfunctions.json` file. Note that in the below code, the option for “sync” is changed from true to false, to accommodate the asynchronous nature of this function:  

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
}
```

Again, you will need to re-register this change once you have saved the file. In Excel, select **Insert > Add-ins > My Add-ins**. This will bring up a list of available add-ins. Under “Developer Add-ins" you will see your add-in, under the name “Excel Custom Function.” Select it to register it.  

In cell B1, run the function `=CONTOSO.STOCKPRICE("MSFT")`. It should show you the stock price for one share of Microsoft stock right now.

## Create a streaming asynchronous custom function
The previous function returned the stock price for Microsoft at a particular moment in time, but stock prices are always changing. With custom functions, it is possible to “stream” data from an API to get updates on stock prices in real time.  

To do this, you’ll create a new function, `=CONTOSO.STOCKPRICESTREAM`. It makes a request for updated data every 1000 milliseconds.  

When a call is made, you may see `#GETTING_DATA` appear in a cell. Once a value is returned, this notification should disappear.  

Copy and paste the code below into `customfunctions.js`.

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

Once again you’ll add to the `./config/customfunctions.json` file with the code below. Note that the “stream” option is marked as true.

Re-register this change once you have saved the file. In Excel, select **Insert > Add-ins > My Add-ins**. This will bring up a list of available add-ins. Under “Developer Add-ins" you will see your add-in, under the name “Excel Custom Function.” Select it to register it.  

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

In cell C1, run the function `=CONTOSO.STOCKPRICESTREAM("MSFT")`. You do not have to specify the caller because it only serves to hold the callback function, `setResult`, which passes data form the function to Excel to update the cell value.  

## Next steps
You’ve completed the custom functions add-in tutorial. For more information about custom functions, check out [this overview article](../excel/custom-functions-overview.md).[TODO format like Ops, make a button]

> [!div class="nextstepaction"]
> [Overview of Custom Functions](../excel/custom-functions-overview.md)