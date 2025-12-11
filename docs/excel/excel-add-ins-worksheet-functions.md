---
title: Calling built-in Excel worksheet functions using the Excel JavaScript API
description: Learn how to call built-in Excel worksheet functions such as `VLOOKUP` and `SUM` using the Excel JavaScript API.
ms.date: 02/17/2022
ms.localizationpriority: medium
---

# Call built-in Excel worksheet functions

This article explains how to call built-in Excel worksheet functions such as `VLOOKUP` and `SUM` using the Excel JavaScript API. It also provides the full list of built-in Excel worksheet functions that can be called using the Excel JavaScript API.

> [!NOTE]
> For information about how to create *custom functions* in Excel using the Excel JavaScript API, see [Create custom functions in Excel](custom-functions-overview.md).

## Calling a worksheet function

The following code snippet shows how to call a worksheet function, where `sampleFunction()` is a placeholder that should be replaced with the name of the function to call and the input parameters that the function requires. The `value` property of the `FunctionResult` object that's returned by a worksheet function contains the result of the specified function. As this example shows, you must `load` the `value` property of the `FunctionResult` object before you can read it. In this example, the result of the function is simply being written to the console.

```js
await Excel.run(async (context) => {
    let functionResult = context.workbook.functions.sampleFunction();
    functionResult.load('value');

    await context.sync();
    console.log('Result of the function: ' + functionResult.value);
});
```

> [!TIP]
> See the [Supported worksheet functions](#supported-worksheet-functions) section of this article for a list of functions that can be called using the Excel JavaScript API.

## Sample data

The following image shows a table in an Excel worksheet that contains sales data for various types of tools over a three month period. Each number in the table represents the number of units sold for a specific tool in a specific month. The examples that follow will show how to apply built-in worksheet functions to this data.

:::image type="content" source="../images/worksheet-functions-chaining-results.jpg" alt-text="Sales data in Excel for Hammer, Wrench, and Saw in months November, December, and January.":::

## Example 1: Single function

The following code sample applies the `VLOOKUP` function to the sample data described previously to identify the number of wrenches sold in November.

```js
await Excel.run(async (context) => {
    let range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    let unitSoldInNov = context.workbook.functions.vlookup("Wrench", range, 2, false);
    unitSoldInNov.load('value');

    await context.sync();
    console.log(' Number of wrenches sold in November = ' + unitSoldInNov.value);
});
```

## Example 2: Nested functions

The following code sample applies the `VLOOKUP` function to the sample data described previously to identify the number of wrenches sold in November and the number of wrenches sold in December, and then applies the `SUM` function to calculate the total number of wrenches sold during those two months.

As this example shows, when one or more function calls are nested within another function call, you only need to `load` the final result that you subsequently want to read (in this example, `sumOfTwoLookups`). Any intermediate results (in this example, the result of each `VLOOKUP` function) will be calculated and used to calculate the final result.

```js
await Excel.run(async (context) => {
    let range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    let sumOfTwoLookups = context.workbook.functions.sum(
        context.workbook.functions.vlookup("Wrench", range, 2, false),
        context.workbook.functions.vlookup("Wrench", range, 3, false)
    );
    sumOfTwoLookups.load('value');

    await context.sync();
    console.log(' Number of wrenches sold in November and December = ' + sumOfTwoLookups.value);
});
```

## Supported worksheet functions

The following built-in Excel worksheet functions can be called using the Excel JavaScript API.

| Function | Description |
|:---------------|:-----------|
| [ABS function](https://support.microsoft.com/office/3420200f-5628-4e8c-99da-c99d7c87713c) | Returns the absolute value of a number |
| [ACCRINT function](https://support.microsoft.com/office/fe45d089-6722-4fb3-9379-e1f911d8dc74) | Returns the accrued interest for a security that pays periodic interest |
| [ACCRINTM function](https://support.microsoft.com/office/f62f01f9-5754-4cc4-805b-0e70199328a7) | Returns the accrued interest for a security that pays interest at maturity |
| [ACOS function](https://support.microsoft.com/office/cb73173f-d089-4582-afa1-76e5524b5d5b) | Returns the arccosine of a number |
| [ACOSH function](https://support.microsoft.com/office/e3992cc1-103f-4e72-9f04-624b9ef5ebfe) | Returns the inverse hyperbolic cosine of a number |
| [ACOT function](https://support.microsoft.com/office/dc7e5008-fe6b-402e-bdd6-2eea8383d905) | Returns the arccotangent of a number |
| [ACOTH function](https://support.microsoft.com/office/cc49480f-f684-4171-9fc5-73e4e852300f) | Returns the hyperbolic arccotangent of a number |
| [AMORDEGRC function](https://support.microsoft.com/office/a14d0ca1-64a4-42eb-9b3d-b0dededf9e51) | Returns the depreciation for each accounting period by using a depreciation coefficient |
| [AMORLINC function](https://support.microsoft.com/office/7d417b45-f7f5-4dba-a0a5-3451a81079a8) | Returns the depreciation for each accounting period |
| [AND function](https://support.microsoft.com/office/5f19b2e8-e1df-4408-897a-ce285a19e9d9) | Returns `TRUE` if all of its arguments are true |
| [ARABIC function](https://support.microsoft.com/office/9a8da418-c17b-4ef9-a657-9370a30a674f) | Converts a Roman number to Arabic, as a number |
| [AREAS function](https://support.microsoft.com/office/8392ba32-7a41-43b3-96b0-3695d2ec6152) | Returns the number of areas in a reference |
| [ASC function](https://support.microsoft.com/office/0b6abf1c-c663-4004-a964-ebc00b723266) | Changes full-width (double-byte) English letters or katakana within a character string to half-width (single-byte) characters |
| [ASIN function](https://support.microsoft.com/office/81fb95e5-6d6f-48c4-bc45-58f955c6d347) | Returns the arcsine of a number |
| [ASINH function](https://support.microsoft.com/office/4e00475a-067a-43cf-926a-765b0249717c) | Returns the inverse hyperbolic sine of a number |
| [ATAN function](https://support.microsoft.com/office/50746fa8-630a-406b-81d0-4a2aed395543) | Returns the arctangent of a number |
| [ATAN2 function](https://support.microsoft.com/office/c04592ab-b9e3-4908-b428-c96b3a565033) | Returns the arctangent from x- and y-coordinates |
| [ATANH function](https://support.microsoft.com/office/3cd65768-0de7-4f1d-b312-d01c8c930d90) | Returns the inverse hyperbolic tangent of a number |
| [AVEDEV function](https://support.microsoft.com/office/58fe8d65-2a84-4dc7-8052-f3f87b5c6639) | Returns the average of the absolute deviations of data points from their mean |
| [AVERAGE function](https://support.microsoft.com/office/047bac88-d466-426c-a32b-8f33eb960cf6) | Returns the average of its arguments |
| [AVERAGEA function](https://support.microsoft.com/office/f5f84098-d453-4f4c-bbba-3d2c66356091) | Returns the average of its arguments, including numbers, text, and logical values |
| [AVERAGEIF function](https://support.microsoft.com/office/faec8e2e-0dec-4308-af69-f5576d8ac642) | Returns the average (arithmetic mean) of all the cells in a range that meet a given criteria |
| [AVERAGEIFS function](https://support.microsoft.com/office/48910c45-1fc0-4389-a028-f7c5c3001690) | Returns the average (arithmetic mean) of all cells that meet multiple criteria |
| [BAHTTEXT function](https://support.microsoft.com/office/5ba4d0b4-abd3-4325-8d22-7a92d59aab9c) | Converts a number to text, using the ß (baht) currency format |
| [BASE function](https://support.microsoft.com/office/2ef61411-aee9-4f29-a811-1c42456c6342) | Converts a number into a text representation with the given radix (base) |
| [BESSELI function](https://support.microsoft.com/office/8d33855c-9a8d-444b-98e0-852267b1c0df) | Returns the modified Bessel function In(x) |
| [BESSELJ function](https://support.microsoft.com/office/839cb181-48de-408b-9d80-bd02982d94f7) | Returns the Bessel function Jn(x) |
| [BESSELK function](https://support.microsoft.com/office/606d11bc-06d3-4d53-9ecb-2803e2b90b70) | Returns the modified Bessel function Kn(x) |
| [BESSELY function](https://support.microsoft.com/office/f3a356b3-da89-42c3-8974-2da54d6353a2) | Returns the Bessel function Yn(x) |
| [BETA.DIST function](https://support.microsoft.com/office/11188c9c-780a-42c7-ba43-9ecb5a878d31) | Returns the beta cumulative distribution function |
| [BETA.INV function](https://support.microsoft.com/office/e84cb8aa-8df0-4cf6-9892-83a341d252eb) | Returns the inverse of the cumulative distribution function for a specified beta distribution |
| [BIN2DEC function](https://support.microsoft.com/office/63905b57-b3a0-453d-99f4-647bb519cd6c) | Converts a binary number to decimal |
| [BIN2HEX function](https://support.microsoft.com/office/0375e507-f5e5-4077-9af8-28d84f9f41cc) | Converts a binary number to hexadecimal |
| [BIN2OCT function](https://support.microsoft.com/office/0a4e01ba-ac8d-4158-9b29-16c25c4c23fd) | Converts a binary number to octal |
| [BINOM.DIST function](https://support.microsoft.com/office/c5ae37b6-f39c-4be2-94c2-509a1480770c) | Returns the individual term binomial distribution probability |
| [BINOM.DIST.RANGE function](https://support.microsoft.com/office/17331329-74c7-4053-bb4c-6653a7421595) | Returns the probability of a trial result using a binomial distribution |
| [BINOM.INV function](https://support.microsoft.com/office/80a0370c-ada6-49b4-83e7-05a91ba77ac9) | Returns the smallest value for which the cumulative binomial distribution is less than or equal to a criterion value |
| [BITAND function](https://support.microsoft.com/office/8a2be3d7-91c3-4b48-9517-64548008563a) | Returns a 'Bitwise And' of two numbers |
| [BITLSHIFT function](https://support.microsoft.com/office/c55bb27e-cacd-4c7c-b258-d80861a03c9c) | Returns a value number shifted left by shift_amount bits |
| [BITOR function](https://support.microsoft.com/office/f6ead5c8-5b98-4c9e-9053-8ad5234919b2) | Returns a bitwise OR of 2 numbers |
| [BITRSHIFT function](https://support.microsoft.com/office/274d6996-f42c-4743-abdb-4ff95351222c) | Returns a value number shifted right by shift_amount bits |
| [BITXOR function](https://support.microsoft.com/office/c81306a1-03f9-4e89-85ac-b86c3cba10e4) | Returns a bitwise 'Exclusive Or' of two numbers |
| [CEILING.MATH, ECMA_CEILING functions](https://support.microsoft.com/office/80f95d2f-b499-4eee-9f16-f795a8e306c8) | Rounds a number up, to the nearest integer or to the nearest multiple of significance |
| [CEILING.PRECISE function](https://support.microsoft.com/office/f366a774-527a-4c92-ba49-af0a196e66cb) | Rounds a number the nearest integer or to the nearest multiple of significance. Regardless of the sign of the number, the number is rounded up. |
| [CHAR function](https://support.microsoft.com/office/bbd249c8-b36e-4a91-8017-1c133f9b837a) | Returns the character specified by the code number |
| [CHISQ.DIST function](https://support.microsoft.com/office/8486b05e-5c05-4942-a9ea-f6b341518732) | Returns the cumulative beta probability density function |
| [CHISQ.DIST.RT function](https://support.microsoft.com/office/dc4832e8-ed2b-49ae-8d7c-b28d5804c0f2) | Returns the one-tailed probability of the chi-squared distribution |
| [CHISQ.INV function](https://support.microsoft.com/office/400db556-62b3-472d-80b3-254723e7092f) | Returns the cumulative beta probability density function |
| [CHISQ.INV.RT function](https://support.microsoft.com/office/435b5ed8-98d5-4da6-823f-293e2cbc94fe) | Returns the inverse of the one-tailed probability of the chi-squared distribution |
| [CHOOSE function](https://support.microsoft.com/office/fc5c184f-cb62-4ec7-a46e-38653b98f5bc) | Chooses a value from a list of values |
| [CLEAN function](https://support.microsoft.com/office/26f3d7c5-475f-4a9c-90e5-4b8ba987ba41) | Removes all nonprintable characters from text |
| [CODE function](https://support.microsoft.com/office/c32b692b-2ed0-4a04-bdd9-75640144b928) | Returns a numeric code for the first character in a text string |
| [COLUMNS function](https://support.microsoft.com/office/4e8e7b4e-e603-43e8-b177-956088fa48ca) | Returns the number of columns in a reference |
| [COMBIN function](https://support.microsoft.com/office/12a3f276-0a21-423a-8de6-06990aaf638a) | Returns the number of combinations for a given number of objects |
| [COMBINA function](https://support.microsoft.com/office/efb49eaa-4f4c-4cd2-8179-0ddfcf9d035d) | Returns the number of combinations with repetitions for a given number of items |
| [COMPLEX function](https://support.microsoft.com/office/f0b8f3a9-51cc-4d6d-86fb-3a9362fa4128) | Converts real and imaginary coefficients into a complex number |
| [CONCATENATE function](https://support.microsoft.com/office/8f8ae884-2ca8-4f7a-b093-75d702bea31d) | Joins several text items into one text item |
| [CONFIDENCE.NORM function](https://support.microsoft.com/office/7cec58a6-85bb-488d-91c3-63828d4fbfd4) | Returns the confidence interval for a population mean |
| [CONFIDENCE.T function](https://support.microsoft.com/office/e8eca395-6c3a-4ba9-9003-79ccc61d3c53) | Returns the confidence interval for a population mean, using a Student's t distribution |
| [CONVERT function](https://support.microsoft.com/office/d785bef1-808e-4aac-bdcd-666c810f9af2) | Converts a number from one measurement system to another |
| [COS function](https://support.microsoft.com/office/0fb808a5-95d6-4553-8148-22aebdce5f05) | Returns the cosine of a number |
| [COSH function](https://support.microsoft.com/office/e460d426-c471-43e8-9540-a57ff3b70555) | Returns the hyperbolic cosine of a number |
| [COT function](https://support.microsoft.com/office/c446f34d-6fe4-40dc-84f8-cf59e5f5e31a) | Returns the cotangent of an angle |
| [COTH function](https://support.microsoft.com/office/2e0b4cb6-0ba0-403e-aed4-deaa71b49df5) | Returns the hyperbolic cotangent of a number |
| [COUNT function](https://support.microsoft.com/office/a59cd7fc-b623-4d93-87a4-d23bf411294c) | Counts how many numbers are in the list of arguments |
| [COUNTA function](https://support.microsoft.com/office/7dc98875-d5c1-46f1-9a82-53f3219e2509) | Counts how many values are in the list of arguments |
| [COUNTBLANK function](https://support.microsoft.com/office/6a92d772-675c-4bee-b346-24af6bd3ac22) | Counts the number of blank cells within a range |
| [COUNTIF function](https://support.microsoft.com/office/e0de10c6-f885-4e71-abb4-1f464816df34) | Counts the number of cells within a range that meet the given criteria |
| [COUNTIFS function](https://support.microsoft.com/office/dda3dc6e-f74e-4aee-88bc-aa8c2a866842) | Counts the number of cells within a range that meet multiple criteria |
| [COUPDAYBS function](https://support.microsoft.com/office/eb9a8dfb-2fb2-4c61-8e5d-690b320cf872) | Returns the number of days from the beginning of the coupon period to the settlement date |
| [COUPDAYS function](https://support.microsoft.com/office/cc64380b-315b-4e7b-950c-b30b0a76f671) | Returns the number of days in the coupon period that contains the settlement date |
| [COUPDAYSNC function](https://support.microsoft.com/office/5ab3f0b2-029f-4a8b-bb65-47d525eea547) | Returns the number of days from the settlement date to the next coupon date |
| [COUPNCD function](https://support.microsoft.com/office/fd962fef-506b-4d9d-8590-16df5393691f) | Returns the next coupon date after the settlement date |
| [COUPNUM function](https://support.microsoft.com/office/a90af57b-de53-4969-9c99-dd6139db2522) | Returns the number of coupons payable between the settlement date and maturity date |
| [COUPPCD function](https://support.microsoft.com/office/2eb50473-6ee9-4052-a206-77a9a385d5b3) | Returns the previous coupon date before the settlement date |
| [CSC function](https://support.microsoft.com/office/07379361-219a-4398-8675-07ddc4f135c1) | Returns the cosecant of an angle |
| [CSCH function](https://support.microsoft.com/office/f58f2c22-eb75-4dd6-84f4-a503527f8eeb) | Returns the hyperbolic cosecant of an angle |
| [CUMIPMT function](https://support.microsoft.com/office/61067bb0-9016-427d-b95b-1a752af0e606) | Returns the cumulative interest paid between two periods |
| [CUMPRINC function](https://support.microsoft.com/office/94a4516d-bd65-41a1-bc16-053a6af4c04d) | Returns the cumulative principal paid on a loan between two periods |
| [DATE function](https://support.microsoft.com/office/e36c0c8c-4104-49da-ab83-82328b832349) | Returns the serial number of a particular date |
| [DATEVALUE function](https://support.microsoft.com/office/df8b07d4-7761-4a93-bc33-b7471bbff252) | Converts a date in the form of text to a serial number |
| [DAVERAGE function](https://support.microsoft.com/office/a6a2d5ac-4b4b-48cd-a1d8-7b37834e5aee) | Returns the average of selected database entries |
| [DAY function](https://support.microsoft.com/office/8a7d1cbb-6c7d-4ba1-8aea-25c134d03101) | Converts a serial number to a day of the month |
| [DAYS function](https://support.microsoft.com/office/57740535-d549-4395-8728-0f07bff0b9df) | Returns the number of days between two dates |
| [DAYS360 function](https://support.microsoft.com/office/b9a509fd-49ef-407e-94df-0cbda5718c2a) | Calculates the number of days between two dates based on a 360-day year |
| [DB function](https://support.microsoft.com/office/354e7d28-5f93-4ff1-8a52-eb4ee549d9d7) | Returns the depreciation of an asset for a specified period by using the fixed-declining balance method |
| [DBCS function](https://support.microsoft.com/office/a4025e73-63d2-4958-9423-21a24794c9e5) | Changes half-width (single-byte) English letters or katakana within a character string to full-width (double-byte) characters |
| [DCOUNT function](https://support.microsoft.com/office/c1fc7b93-fb0d-4d8d-97db-8d5f076eaeb1) | Counts the cells that contain numbers in a database |
| [DCOUNTA function](https://support.microsoft.com/office/00232a6d-5a66-4a01-a25b-c1653fda1244) | Counts nonblank cells in a database |
| [DDB function](https://support.microsoft.com/office/519a7a37-8772-4c96-85c0-ed2c209717a5) | Returns the depreciation of an asset for a specified period by using the double-declining balance method or some other method that you specify |
| [DEC2BIN function](https://support.microsoft.com/office/0f63dd0e-5d1a-42d8-b511-5bf5c6d43838) | Converts a decimal number to binary |
| [DEC2HEX function](https://support.microsoft.com/office/6344ee8b-b6b5-4c6a-a672-f64666704619) | Converts a decimal number to hexadecimal |
| [DEC2OCT function](https://support.microsoft.com/office/c9d835ca-20b7-40c4-8a9e-d3be351ce00f) | Converts a decimal number to octal |
| [DECIMAL function](https://support.microsoft.com/office/ee554665-6176-46ef-82de-0a283658da2e) | Converts a text representation of a number in a given base into a decimal number |
| [DEGREES function](https://support.microsoft.com/office/4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1) | Converts radians to degrees |
| [DELTA function](https://support.microsoft.com/office/2f763672-c959-4e07-ac33-fe03220ba432) | Tests whether two values are equal |
| [DEVSQ function](https://support.microsoft.com/office/8b739616-8376-4df5-8bd0-cfe0a6caf444) | Returns the sum of squares of deviations |
| [DGET function](https://support.microsoft.com/office/455568bf-4eef-45f7-90f0-ec250d00892e) | Extracts from a database a single record that matches the specified criteria |
| [DISC function](https://support.microsoft.com/office/71fce9f3-3f05-4acf-a5a3-eac6ef4daa53) | Returns the discount rate for a security |
| [DMAX function](https://support.microsoft.com/office/f4e8209d-8958-4c3d-a1ee-6351665d41c2) | Returns the maximum value from selected database entries |
| [DMIN function](https://support.microsoft.com/office/4ae6f1d9-1f26-40f1-a783-6dc3680192a3) | Returns the minimum value from selected database entries |
| [DOLLAR, USDOLLAR functions](https://support.microsoft.com/office/a6cd05d9-9740-4ad3-a469-8109d18ff611) | Converts a number to text, using the $ (dollar) currency format |
| [DOLLARDE function](https://support.microsoft.com/office/db85aab0-1677-428a-9dfd-a38476693427) | Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number |
| [DOLLARFR function](https://support.microsoft.com/office/0835d163-3023-4a33-9824-3042c5d4f495) | Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction |
| [DPRODUCT function](https://support.microsoft.com/office/4f96b13e-d49c-47a7-b769-22f6d017cb31) | Multiplies the values in a particular field of records that match the criteria in a database |
| [DSTDEV function](https://support.microsoft.com/office/026b8c73-616d-4b5e-b072-241871c4ab96) | Estimates the standard deviation based on a sample of selected database entries |
| [DSTDEVP function](https://support.microsoft.com/office/04b78995-da03-4813-bbd9-d74fd0f5d94b) | Calculates the standard deviation based on the entire population of selected database entries |
| [DSUM function](https://support.microsoft.com/office/53181285-0c4b-4f5a-aaa3-529a322be41b) | Adds the numbers in the field column of records in the database that match the criteria |
| [DURATION function](https://support.microsoft.com/office/b254ea57-eadc-4602-a86a-c8e369334038) | Returns the annual duration of a security with periodic interest payments |
| [Dlet function](https://support.microsoft.com/office/d6747ca9-99c7-48bb-996e-9d7af00f3ed1) | Estimates variance based on a sample from selected database entries |
| [DVARP function](https://support.microsoft.com/office/eb0ba387-9cb7-45c8-81e9-0394912502fc) | Calculates variance based on the entire population of selected database entries |
| [EDATE function](https://support.microsoft.com/office/3c920eb2-6e66-44e7-a1f5-753ae47ee4f5) | Returns the serial number of the date that is the indicated number of months before or after the start date |
| [EFFECT function](https://support.microsoft.com/office/910d4e4c-79e2-4009-95e6-507e04f11bc4) | Returns the effective annual interest rate |
| [EOMONTH function](https://support.microsoft.com/office/7314ffa1-2bc9-4005-9d66-f49db127d628) | Returns the serial number of the last day of the month before or after a specified number of months |
| [ERF function](https://support.microsoft.com/office/c53c7e7b-5482-4b6c-883e-56df3c9af349) | Returns the error function |
| [ERF.PRECISE function](https://support.microsoft.com/office/9a349593-705c-4278-9a98-e4122831a8e0) | Returns the error function |
| [ERFC function](https://support.microsoft.com/office/736e0318-70ba-4e8b-8d08-461fe68b71b3) | Returns the complementary error function |
| [ERFC.PRECISE function](https://support.microsoft.com/office/e90e6bab-f45e-45df-b2ac-cd2eb4d4a273) | Returns the complementary ERF function integrated between x and infinity |
| [ERROR.TYPE function](https://support.microsoft.com/office/10958677-7c8d-44f7-ae77-b9a9ee6eefaa) | Returns a number corresponding to an error type |
| [EVEN function](https://support.microsoft.com/office/197b5f06-c795-4c1e-8696-3c3b8a646cf9) | Rounds a number up to the nearest even integer |
| [EXACT function](https://support.microsoft.com/office/d3087698-fc15-4a15-9631-12575cf29926) | Checks to see if two text values are identical |
| [EXP function](https://support.microsoft.com/office/c578f034-2c45-4c37-bc8c-329660a63abe) | Returns e raised to the power of a given number |
| [EXPON.DIST function](https://support.microsoft.com/office/4c12ae24-e563-4155-bf3e-8b78b6ae140e) | Returns the exponential distribution |
| [F.DIST function](https://support.microsoft.com/office/a887efdc-7c8e-46cb-a74a-f884cd29b25d) | Returns the F probability distribution |
| [F.DIST.RT function](https://support.microsoft.com/office/d74cbb00-6017-4ac9-b7d7-6049badc0520) | Returns the F probability distribution |
| [F.INV function](https://support.microsoft.com/office/0dda0cf9-4ea0-42fd-8c3c-417a1ff30dbe) | Returns the inverse of the F probability distribution |
| [F.INV.RT function](https://support.microsoft.com/office/d371aa8f-b0b1-40ef-9cc2-496f0693ac00) | Returns the inverse of the F probability distribution |
| [FACT function](https://support.microsoft.com/office/ca8588c2-15f2-41c0-8e8c-c11bd471a4f3) | Returns the factorial of a number |
| [FACTDOUBLE function](https://support.microsoft.com/office/e67697ac-d214-48eb-b7b7-cce2589ecac8) | Returns the double factorial of a number |
| [FALSE function](https://support.microsoft.com/office/2d58dfa5-9c03-4259-bf8f-f0ae14346904) | Returns the logical value `FALSE` |
| [FIND, FINDB functions](https://support.microsoft.com/office/c7912941-af2a-4bdf-a553-d0d89b0a0628) | Finds one text value within another (case-sensitive) |
| [FISHER function](https://support.microsoft.com/office/d656523c-5076-4f95-b87b-7741bf236c69) | Returns the Fisher transformation |
| [FISHERINV function](https://support.microsoft.com/office/62504b39-415a-4284-a285-19c8e82f86bb) | Returns the inverse of the Fisher transformation |
| [FIXED function](https://support.microsoft.com/office/ffd5723c-324c-45e9-8b96-e41be2a8274a) | Formats a number as text with a fixed number of decimals |
| [FLOOR.MATH function](https://support.microsoft.com/office/c302b599-fbdb-4177-ba19-2c2b1249a2f5) | Rounds a number down, to the nearest integer or to the nearest multiple of significance |
| [FLOOR.PRECISE function](https://support.microsoft.com/office/f769b468-1452-4617-8dc3-02f842a0702e) | Rounds a number down to the nearest integer or to the nearest multiple of significance. Regardless of the sign of the number, the number is rounded down. |
| [FV function](https://support.microsoft.com/office/2eef9f44-a084-4c61-bdd8-4fe4bb1b71b3) | Returns the future value of an investment |
| [FVSCHEDULE function](https://support.microsoft.com/office/bec29522-bd87-4082-bab9-a241f3fb251d) | Returns the future value of an initial principal after applying a series of compound interest rates |
| [GAMMA function](https://support.microsoft.com/office/ce1702b1-cf55-471d-8307-f83be0fc5297) | Returns the Gamma function value |
| [GAMMA.DIST function](https://support.microsoft.com/office/9b6f1538-d11c-4d5f-8966-21f6a2201def) | Returns the gamma distribution |
| [GAMMA.INV function](https://support.microsoft.com/office/74991443-c2b0-4be5-aaab-1aa4d71fbb18) | Returns the inverse of the gamma cumulative distribution |
| [GAMMALN function](https://support.microsoft.com/office/b838c48b-c65f-484f-9e1d-141c55470eb9) | Returns the natural logarithm of the gamma function, Γ(x) |
| [GAMMALN.PRECISE function](https://support.microsoft.com/office/5cdfe601-4e1e-4189-9d74-241ef1caa599) | Returns the natural logarithm of the gamma function, Γ(x) |
| [GAUSS function](https://support.microsoft.com/office/069f1b4e-7dee-4d6a-a71f-4b69044a6b33) | Returns 0.5 less than the standard normal cumulative distribution |
| [GCD function](https://support.microsoft.com/office/d5107a51-69e3-461f-8e4c-ddfc21b5073a) | Returns the greatest common divisor |
| [GEOMEAN function](https://support.microsoft.com/office/db1ac48d-25a5-40a0-ab83-0b38980e40d5) | Returns the geometric mean |
| [GESTEP function](https://support.microsoft.com/office/f37e7d2a-41da-4129-be95-640883fca9df) | Tests whether a number is greater than a threshold value |
| [HARMEAN function](https://support.microsoft.com/office/5efd9184-fab5-42f9-b1d3-57883a1d3bc6) | Returns the harmonic mean |
| [HEX2BIN function](https://support.microsoft.com/office/a13aafaa-5737-4920-8424-643e581828c1) | Converts a hexadecimal number to binary |
| [HEX2DEC function](https://support.microsoft.com/office/8c8c3155-9f37-45a5-a3ee-ee5379ef106e) | Converts a hexadecimal number to decimal |
| [HEX2OCT function](https://support.microsoft.com/office/54d52808-5d19-4bd0-8a63-1096a5d11912) | Converts a hexadecimal number to octal |
| [HLOOKUP function](https://support.microsoft.com/office/a3034eec-b719-4ba3-bb65-e1ad662ed95f) | Looks in the top row of an array and returns the value of the indicated cell |
| [HOUR function](https://support.microsoft.com/office/a3afa879-86cb-4339-b1b5-2dd2d7310ac7) | Converts a serial number to an hour |
| [HYPERLINK function](https://support.microsoft.com/office/333c7ce6-c5ae-4164-9c47-7de9b76f577f) | Creates a shortcut or jump that opens a document stored on a network server, an intranet, or the Internet |
| [HYPGEOM.DIST function](https://support.microsoft.com/office/6dbd547f-1d12-4b1f-8ae5-b0d9e3d22fbf) | Returns the hypergeometric distribution |
| [IF function](https://support.microsoft.com/office/69aed7c9-4e8a-4755-a9bc-aa8bbff73be2) | Specifies a logical test to perform |
| [IMABS function](https://support.microsoft.com/office/b31e73c6-d90c-4062-90bc-8eb351d765a1) | Returns the absolute value (modulus) of a complex number |
| [IMAGINARY function](https://support.microsoft.com/office/dd5952fd-473d-44d9-95a1-9a17b23e428a) | Returns the imaginary coefficient of a complex number |
| [IMARGUMENT function](https://support.microsoft.com/office/eed37ec1-23b3-4f59-b9f3-d340358a034a) | Returns the argument theta, an angle expressed in radians |
| [IMCONJUGATE function](https://support.microsoft.com/office/2e2fc1ea-f32b-4f9b-9de6-233853bafd42) | Returns the complex conjugate of a complex number |
| [IMCOS function](https://support.microsoft.com/office/dad75277-f592-4a6b-ad6c-be93a808a53c) | Returns the cosine of a complex number |
| [IMCOSH function](https://support.microsoft.com/office/053e4ddb-4122-458b-be9a-457c405e90ff) | Returns the hyperbolic cosine of a complex number |
| [IMCOT function](https://support.microsoft.com/office/dc6a3607-d26a-4d06-8b41-8931da36442c) | Returns the cotangent of a complex number |
| [IMCSC function](https://support.microsoft.com/office/9e158d8f-2ddf-46cd-9b1d-98e29904a323) | Returns the cosecant of a complex number |
| [IMCSCH function](https://support.microsoft.com/office/c0ae4f54-5f09-4fef-8da0-dc33ea2c5ca9) | Returns the hyperbolic cosecant of a complex number |
| [IMDIV function](https://support.microsoft.com/office/a505aff7-af8a-4451-8142-77ec3d74d83f) | Returns the quotient of two complex numbers |
| [IMEXP function](https://support.microsoft.com/office/c6f8da1f-e024-4c0c-b802-a60e7147a95f) | Returns the exponential of a complex number |
| [IMLN function](https://support.microsoft.com/office/32b98bcf-8b81-437c-a636-6fb3aad509d8) | Returns the natural logarithm of a complex number |
| [IMLOG10 function](https://support.microsoft.com/office/58200fca-e2a2-4271-8a98-ccd4360213a5) | Returns the base-10 logarithm of a complex number |
| [IMLOG2 function](https://support.microsoft.com/office/152e13b4-bc79-486c-a243-e6a676878c51) | Returns the base-2 logarithm of a complex number |
| [IMPOWER function](https://support.microsoft.com/office/210fd2f5-f8ff-4c6a-9d60-30e34fbdef39) | Returns a complex number raised to an integer power |
| [IMPRODUCT function](https://support.microsoft.com/office/2fb8651a-a4f2-444f-975e-8ba7aab3a5ba) | Returns the product of from 2 to 255 complex numbers |
| [IMREAL function](https://support.microsoft.com/office/d12bc4c0-25d0-4bb3-a25f-ece1938bf366) | Returns the real coefficient of a complex number |
| [IMSEC function](https://support.microsoft.com/office/6df11132-4411-4df4-a3dc-1f17372459e0) | Returns the secant of a complex number |
| [IMSECH function](https://support.microsoft.com/office/f250304f-788b-4505-954e-eb01fa50903b) | Returns the hyperbolic secant of a complex number |
| [IMSIN function](https://support.microsoft.com/office/1ab02a39-a721-48de-82ef-f52bf37859f6) | Returns the sine of a complex number |
| [IMSINH function](https://support.microsoft.com/office/dfb9ec9e-8783-4985-8c42-b028e9e8da3d) | Returns the hyperbolic sine of a complex number |
| [IMSQRT function](https://support.microsoft.com/office/e1753f80-ba11-4664-a10e-e17368396b70) | Returns the square root of a complex number |
| [IMSUB function](https://support.microsoft.com/office/2e404b4d-4935-4e85-9f52-cb08b9a45054) | Returns the difference between two complex numbers |
| [IMSUM function](https://support.microsoft.com/office/81542999-5f1c-4da6-9ffe-f1d7aaa9457f) | Returns the sum of complex numbers |
| [IMTAN function](https://support.microsoft.com/office/8478f45d-610a-43cf-8544-9fc0b553a132) | Returns the tangent of a complex number |
| [INT function](https://support.microsoft.com/office/a6c4af9e-356d-4369-ab6a-cb1fd9d343ef) | Rounds a number down to the nearest integer |
| [INTRATE function](https://support.microsoft.com/office/5cb34dde-a221-4cb6-b3eb-0b9e55e1316f) | Returns the interest rate for a fully invested security |
| [IPMT function](https://support.microsoft.com/office/5cce0ad6-8402-4a41-8d29-61a0b054cb6f) | Returns the interest payment for an investment for a given period |
| [IRR function](https://support.microsoft.com/office/64925eaa-9988-495b-b290-3ad0c163c1bc) | Returns the internal rate of return for a series of cash flows |
| [ISERR function](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Returns `TRUE` if the value is any error value except #N/A |
| [ISERROR function](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Returns `TRUE` if the value is any error value |
| [ISEVEN function](https://support.microsoft.com/office/aa15929a-d77b-4fbb-92f4-2f479af55356) | Returns `TRUE` if the number is even |
| [ISFORMULA function](https://support.microsoft.com/office/e4d1355f-7121-4ef2-801e-3839bfd6b1e5) | Returns `TRUE` if there is a reference to a cell that contains a formula |
| [ISLOGICAL function](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Returns `TRUE` if the value is a logical value |
| [ISNA function](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Returns `TRUE` if the value is the #N/A error value |
| [ISNONTEXT function](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Returns `TRUE` if the value is not text |
| [ISNUMBER function](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Returns `TRUE` if the value is a number |
| [ISO.CEILING function](https://support.microsoft.com/office/e587bb73-6cc2-4113-b664-ff5b09859a83) | Returns a number that is rounded up to the nearest integer or to the nearest multiple of significance |
| [ISODD function](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Returns `TRUE` if the number is odd |
| [ISOWEEKNUM function](https://support.microsoft.com/office/1c2d0afe-d25b-4ab1-8894-8d0520e90e0e) | Returns the number of the ISO week number of the year for a given date |
| [ISPMT function](https://support.microsoft.com/office/fa58adb6-9d39-4ce0-8f43-75399cea56cc) | Calculates the interest paid during a specific period of an investment |
| [ISREF function](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Returns `TRUE` if the value is a reference |
| [ISTEXT function](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Returns `TRUE` if the value is text |
| [KURT function](https://support.microsoft.com/office/bc3a265c-5da4-4dcb-b7fd-c237789095ab) | Returns the kurtosis of a data set |
| [LARGE function](https://support.microsoft.com/office/3af0af19-1190-42bb-bb8b-01672ec00a64) | Returns the k-th largest value in a data set |
| [LCM function](https://support.microsoft.com/office/7152b67a-8bb5-4075-ae5c-06ede5563c94) | Returns the least common multiple |
| [LEFT, LEFTB functions](https://support.microsoft.com/office/9203d2d2-7960-479b-84c6-1ea52b99640c) | Returns the leftmost characters from a text value |
| [LEN, LENB functions](https://support.microsoft.com/office/29236f94-cedc-429d-affd-b5e33d2c67cb) | Returns the number of characters in a text string |
| [LN function](https://support.microsoft.com/office/81fe1ed7-dac9-4acd-ba1d-07a142c6118f) | Returns the natural logarithm of a number |
| [LOG function](https://support.microsoft.com/office/4e82f196-1ca9-4747-8fb0-6c4a3abb3280) | Returns the logarithm of a number to a specified base |
| [LOG10 function](https://support.microsoft.com/office/c75b881b-49dd-44fb-b6f4-37e3486a0211) | Returns the base-10 logarithm of a number |
| [LOGNORM.DIST function](https://support.microsoft.com/office/eb60d00b-48a9-4217-be2b-6074aee6b070) | Returns the cumulative lognormal distribution |
| [LOGNORM.INV function](https://support.microsoft.com/office/fe79751a-f1f2-4af8-a0a1-e151b2d4f600) | Returns the inverse of the lognormal cumulative distribution |
| [LOOKUP function](https://support.microsoft.com/office/446d94af-663b-451d-8251-369d5e3864cb) | Looks up values in a vector or array |
| [LOWER function](https://support.microsoft.com/office/3f21df02-a80c-44b2-afaf-81358f9fdeb4) | Converts text to lowercase |
| [MATCH function](https://support.microsoft.com/office/e8dffd45-c762-47d6-bf89-533f4a37673a) | Looks up values in a reference or array |
| [MAX function](https://support.microsoft.com/office/e0012414-9ac8-4b34-9a47-73e662c08098) | Returns the maximum value in a list of arguments |
| [MAXA function](https://support.microsoft.com/office/814bda1e-3840-4bff-9365-2f59ac2ee62d) | Returns the maximum value in a list of arguments, including numbers, text, and logical values |
| [MDURATION function](https://support.microsoft.com/office/b3786a69-4f20-469a-94ad-33e5b90a763c) | Returns the Macauley modified duration for a security with an assumed par value of $100 |
| [MEDIAN function](https://support.microsoft.com/office/d0916313-4753-414c-8537-ce85bdd967d2) | Returns the median of the given numbers |
| [MID, MIDB functions](https://support.microsoft.com/office/d5f9e25c-d7d6-472e-b568-4ecb12433028) | Returns a specific number of characters from a text string starting at the position you specify |
| [MIN function](https://support.microsoft.com/office/61635d12-920f-4ce2-a70f-96f202dcc152) | Returns the minimum value in a list of arguments |
| [MINA function](https://support.microsoft.com/office/245a6f46-7ca5-4dc7-ab49-805341bc31d3) | Returns the smallest value in a list of arguments, including numbers, text, and logical values |
| [MINUTE function](https://support.microsoft.com/office/af728df0-05c4-4b07-9eed-a84801a60589) | Converts a serial number to a minute |
| [MIRR function](https://support.microsoft.com/office/b020f038-7492-4fb4-93c1-35c345b53524) | Returns the internal rate of return where positive and negative cash flows are financed at different rates |
| [MOD function](https://support.microsoft.com/office/9b6cd169-b6ee-406a-a97b-edf2a9dc24f3) | Returns the remainder from division |
| [MONTH function](https://support.microsoft.com/office/579a2881-199b-48b2-ab90-ddba0eba86e8) | Converts a serial number to a month |
| [MROUND function](https://support.microsoft.com/office/c299c3b0-15a5-426d-aa4b-d2d5b3baf427) | Returns a number rounded to the desired multiple |
| [MULTINOMIAL function](https://support.microsoft.com/office/6fa6373c-6533-41a2-a45e-a56db1db1bf6) | Returns the multinomial of a set of numbers |
| [N function](https://support.microsoft.com/office/a624cad1-3635-4208-b54a-29733d1278c9) | Returns a value converted to a number |
| [NA function](https://support.microsoft.com/office/5469c2d1-a90c-4fb5-9bbc-64bd9bb6b47c) | Returns the error value #N/A |
| [NEGBINOM.DIST function](https://support.microsoft.com/office/c8239f89-c2d0-45bd-b6af-172e570f8599) | Returns the negative binomial distribution |
| [NETWORKDAYS function](https://support.microsoft.com/office/48e717bf-a7a3-495f-969e-5005e3eb18e7) | Returns the number of whole workdays between two dates |
| [NETWORKDAYS.INTL function](https://support.microsoft.com/office/a9b26239-4f20-46a1-9ab8-4e925bfd5e28) | Returns the number of whole workdays between two dates using parameters to indicate which and how many days are weekend days |
| [NOMINAL function](https://support.microsoft.com/office/7f1ae29b-6b92-435e-b950-ad8b190ddd2b) | Returns the annual nominal interest rate |
| [NORM.DIST function](https://support.microsoft.com/office/edb1cc14-a21c-4e53-839d-8082074c9f8d) | Returns the normal cumulative distribution |
| [NORM.INV function](https://support.microsoft.com/office/54b30935-fee7-493c-bedb-2278a9db7e13) | Returns the inverse of the normal cumulative distribution |
| [NORM.S.DIST function](https://support.microsoft.com/office/1e787282-3832-4520-a9ae-bd2a8d99ba88) | Returns the standard normal cumulative distribution |
| [NORM.S.INV function](https://support.microsoft.com/office/d6d556b4-ab7f-49cd-b526-5a20918452b1) | Returns the inverse of the standard normal cumulative distribution |
| [NOT function](https://support.microsoft.com/office/9cfc6011-a054-40c7-a140-cd4ba2d87d77) | Reverses the logic of its argument |
| [NOW function](https://support.microsoft.com/office/3337fd29-145a-4347-b2e6-20c904739c46) | Returns the serial number of the current date and time |
| [NPER function](https://support.microsoft.com/office/240535b5-6653-4d2d-bfcf-b6a38151d815) | Returns the number of periods for an investment |
| [NPV function](https://support.microsoft.com/office/8672cb67-2576-4d07-b67b-ac28acf2a568) | Returns the net present value of an investment based on a series of periodic cash flows and a discount rate |
| [NUMBERVALUE function](https://support.microsoft.com/office/1b05c8cf-2bfa-4437-af70-596c7ea7d879) | Converts text to number in a locale-independent manner |
| [OCT2BIN function](https://support.microsoft.com/office/55383471-3c56-4d27-9522-1a8ec646c589) | Converts an octal number to binary |
| [OCT2DEC function](https://support.microsoft.com/office/87606014-cb98-44b2-8dbb-e48f8ced1554) | Converts an octal number to decimal |
| [OCT2HEX function](https://support.microsoft.com/office/912175b4-d497-41b4-a029-221f051b858f) | Converts an octal number to hexadecimal |
| [ODD function](https://support.microsoft.com/office/deae64eb-e08a-4c88-8b40-6d0b42575c98) | Rounds a number up to the nearest odd integer |
| [ODDFPRICE function](https://support.microsoft.com/office/d7d664a8-34df-4233-8d2b-922bcf6a69e1) | Returns the price per $100 face value of a security with an odd first period |
| [ODDFYIELD function](https://support.microsoft.com/office/66bc8b7b-6501-4c93-9ce3-2fd16220fe37) | Returns the yield of a security with an odd first period |
| [ODDLPRICE function](https://support.microsoft.com/office/fb657749-d200-4902-afaf-ed5445027fc4) | Returns the price per $100 face value of a security with an odd last period |
| [ODDLYIELD function](https://support.microsoft.com/office/c873d088-cf40-435f-8d41-c8232fee9238) | Returns the yield of a security with an odd last period |
| [OR function](https://support.microsoft.com/office/7d17ad14-8700-4281-b308-00b131e22af0) | Returns `TRUE` if any argument is true |
| [PDURATION function](https://support.microsoft.com/office/44f33460-5be5-4c90-b857-22308892adaf) | Returns the number of periods required by an investment to reach a specified value |
| [PERCENTILE.EXC function](https://support.microsoft.com/office/bbaa7204-e9e1-4010-85bf-c31dc5dce4ba) | Returns the k-th percentile of values in a range, where k is in the range 0..1, exclusive |
| [PERCENTILE.INC function](https://support.microsoft.com/office/680f9539-45eb-410b-9a5e-c1355e5fe2ed) | Returns the k-th percentile of values in a range |
| [PERCENTRANK.EXC function](https://support.microsoft.com/office/d8afee96-b7e2-4a2f-8c01-8fcdedaa6314) | Returns the rank of a value in a data set as a percentage (0..1, exclusive) of the data set |
| [PERCENTRANK.INC function](https://support.microsoft.com/office/149592c9-00c0-49ba-86c1-c1f45b80463a) | Returns the percentage rank of a value in a data set |
| [PERMUT function](https://support.microsoft.com/office/3bd1cb9a-2880-41ab-a197-f246a7a602d3) | Returns the number of permutations for a given number of objects |
| [PERMUTATIONA function](https://support.microsoft.com/office/6c7d7fdc-d657-44e6-aa19-2857b25cae4e) | Returns the number of permutations for a given number of objects (with repetitions) that can be selected from the total objects |
| [PHI function](https://support.microsoft.com/office/23e49bc6-a8e8-402d-98d3-9ded87f6295c) | Returns the value of the density function for a standard normal distribution |
| [PI function](https://support.microsoft.com/office/264199d0-a3ba-46b8-975a-c4a04608989b) | Returns the value of pi |
| [PMT function](https://support.microsoft.com/office/0214da64-9a63-4996-bc20-214433fa6441) | Returns the periodic payment for an annuity |
| [POISSON.DIST function](https://support.microsoft.com/office/8fe148ff-39a2-46cb-abf3-7772695d9636) | Returns the Poisson distribution |
| [POWER function](https://support.microsoft.com/office/d3f2908b-56f4-4c3f-895a-07fb519c362a) | Returns the result of a number raised to a power |
| [PPMT function](https://support.microsoft.com/office/c370d9e3-7749-4ca4-beea-b06c6ac95e1b) | Returns the payment on the principal for an investment for a given period |
| [PRICE function](https://support.microsoft.com/office/3ea9deac-8dfa-436f-a7c8-17ea02c21b0a) | Returns the price per $100 face value of a security that pays periodic interest |
| [PRICEDISC function](https://support.microsoft.com/office/d06ad7c1-380e-4be7-9fd9-75e3079acfd3) | Returns the price per $100 face value of a discounted security |
| [PRICEMAT function](https://support.microsoft.com/office/52c3b4da-bc7e-476a-989f-a95f675cae77) | Returns the price per $100 face value of a security that pays interest at maturity |
| [PRODUCT function](https://support.microsoft.com/office/8e6b5b24-90ee-4650-aeec-80982a0512ce) | Multiplies its arguments |
| [PROPER function](https://support.microsoft.com/office/52a5a283-e8b2-49be-8506-b2887b889f94) | Capitalizes the first letter in each word of a text value |
| [PV function](https://support.microsoft.com/office/23879d31-0e02-4321-be01-da16e8168cbd) | Returns the present value of an investment |
| [QUARTILE.EXC function](https://support.microsoft.com/office/5a355b7a-840b-4a01-b0f1-f538c2864cad) | Returns the quartile of the data set, based on percentile values from 0..1, exclusive |
| [QUARTILE.INC function](https://support.microsoft.com/office/1bbacc80-5075-42f1-aed6-47d735c4819d) | Returns the quartile of a data set |
| [QUOTIENT function](https://support.microsoft.com/office/9f7bf099-2a18-4282-8fa4-65290cc99dee) | Returns the integer portion of a division |
| [RADIANS function](https://support.microsoft.com/office/ac409508-3d48-45f5-ac02-1497c92de5bf) | Converts degrees to radians |
| [RAND function](https://support.microsoft.com/office/4cbfa695-8869-4788-8d90-021ea9f5be73) | Returns a random number between 0 and 1 |
| [RANDBETWEEN function](https://support.microsoft.com/office/4cc7f0d1-87dc-4eb7-987f-a469ab381685) | Returns a random number between the numbers you specify |
| [RANK.AVG function](https://support.microsoft.com/office/bd406a6f-eb38-4d73-aa8e-6d1c3c72e83a) | Returns the rank of a number in a list of numbers |
| [RANK.EQ function](https://support.microsoft.com/office/284858ce-8ef6-450e-b662-26245be04a40) | Returns the rank of a number in a list of numbers |
| [RATE function](https://support.microsoft.com/office/9f665657-4a7e-4bb7-a030-83fc59e748ce) | Returns the interest rate per period of an annuity |
| [RECEIVED function](https://support.microsoft.com/office/7a3f8b93-6611-4f81-8576-828312c9b5e5) | Returns the amount received at maturity for a fully invested security |
| [REPLACE, REPLACEB functions](https://support.microsoft.com/office/8d799074-2425-4a8a-84bc-82472868878a) | Replaces characters within text |
| [REPT function](https://support.microsoft.com/office/04c4d778-e712-43b4-9c15-d656582bb061) | Repeats text a given number of times |
| [RIGHT, RIGHTB functions](https://support.microsoft.com/office/240267ee-9afa-4639-a02b-f19e1786cf2f) | Returns the rightmost characters from a text value |
| [ROMAN function](https://support.microsoft.com/office/d6b0b99e-de46-4704-a518-b45a0f8b56f5) | Converts an Arabic numeral to Roman, as text |
| [ROUND function](https://support.microsoft.com/office/c018c5d8-40fb-4053-90b1-b3e7f61a213c) | Rounds a number to a specified number of digits |
| [ROUNDDOWN function](https://support.microsoft.com/office/2ec94c73-241f-4b01-8c6f-17e6d7968f53) | Rounds a number down, toward zero |
| [ROUNDUP function](https://support.microsoft.com/office/f8bc9b23-e795-47db-8703-db171d0c42a7) | Rounds a number up, away from zero |
| [ROWS function](https://support.microsoft.com/office/b592593e-3fc2-47f2-bec1-bda493811597) | Returns the number of rows in a reference |
| [RRI function](https://support.microsoft.com/office/6f5822d8-7ef1-4233-944c-79e8172930f4) | Returns an equivalent interest rate for the growth of an investment |
| [SEC function](https://support.microsoft.com/office/ff224717-9c87-4170-9b58-d069ced6d5f7) | Returns the secant of an angle |
| [SECH function](https://support.microsoft.com/office/e05a789f-5ff7-4d7f-984a-5edb9b09556f) | Returns the hyperbolic secant of an angle |
| [SECOND function](https://support.microsoft.com/office/740d1cfc-553c-4099-b668-80eaa24e8af1) | Converts a serial number to a second |
| [SERIESSUM function](https://support.microsoft.com/office/a3ab25b5-1093-4f5b-b084-96c49087f637) | Returns the sum of a power series based on the formula |
| [SHEET function](https://support.microsoft.com/office/44718b6f-8b87-47a1-a9d6-b701c06cff24) | Returns the sheet number of the referenced sheet |
| [SHEETS function](https://support.microsoft.com/office/770515eb-e1e8-45ce-8066-b557e5e4b80b) | Returns the number of sheets in a reference |
| [SIGN function](https://support.microsoft.com/office/109c932d-fcdc-4023-91f1-2dd0e916a1d8) | Returns the sign of a number |
| [SIN function](https://support.microsoft.com/office/cf0e3432-8b9e-483c-bc55-a76651c95602) | Returns the sine of the given angle |
| [SINH function](https://support.microsoft.com/office/1e4e8b9f-2b65-43fc-ab8a-0a37f4081fa7) | Returns the hyperbolic sine of a number |
| [SKEW function](https://support.microsoft.com/office/bdf49d86-b1ef-4804-a046-28eaea69c9fa) | Returns the skewness of a distribution |
| [SKEW.P function](https://support.microsoft.com/office/76530a5c-99b9-48a1-8392-26632d542fcb) | Returns the skewness of a distribution based on a population: a characterization of the degree of asymmetry of a distribution around its mean |
| [SLN function](https://support.microsoft.com/office/cdb666e5-c1c6-40a7-806a-e695edc2f1c8) | Returns the straight-line depreciation of an asset for one period |
| [SMALL function](https://support.microsoft.com/office/17da8222-7c82-42b2-961b-14c45384df07) | Returns the k-th smallest value in a data set |
| [SQRT function](https://support.microsoft.com/office/654975c2-05c4-4831-9a24-2c65e4040fdf) | Returns a positive square root |
| [SQRTPI function](https://support.microsoft.com/office/1fb4e63f-9b51-46d6-ad68-b3e7a8b519b4) | Returns the square root of (number * pi) |
| [STANDARDIZE function](https://support.microsoft.com/office/81d66554-2d54-40ec-ba83-6437108ee775) | Returns a normalized value |
| [STDEV.P function](https://support.microsoft.com/office/6e917c05-31a0-496f-ade7-4f4e7462f285) | Calculates standard deviation based on the entire population |
| [STDEV.S function](https://support.microsoft.com/office/7d69cf97-0c1f-4acf-be27-f3e83904cc23) | Estimates standard deviation based on a sample |
| [STDEVA function](https://support.microsoft.com/office/5ff38888-7ea5-48de-9a6d-11ed73b29e9d) | Estimates standard deviation based on a sample, including numbers, text, and logical values |
| [STDEVPA function](https://support.microsoft.com/office/5578d4d6-455a-4308-9991-d405afe2c28c) | Calculates standard deviation based on the entire population, including numbers, text, and logical values |
| [SUBSTITUTE function](https://support.microsoft.com/office/6434944e-a904-4336-a9b0-1e58df3bc332) | Substitutes new text for old text in a text string |
| [SUBTOTAL function](https://support.microsoft.com/office/7b027003-f060-4ade-9040-e478765b9939) | Returns a subtotal in a list or database |
| [SUM function](https://support.microsoft.com/office/043e1c7d-7726-4e80-8f32-07b23e057f89) | Adds its arguments |
| [SUMIF function](https://support.microsoft.com/office/169b8c99-c05c-4483-a712-1697a653039b) | Adds the cells specified by a given criteria |
| [SUMIFS function](https://support.microsoft.com/office/c9e748f5-7ea7-455d-9406-611cebce642b) | Adds the cells in a range that meet multiple criteria |
| [SUMSQ function](https://support.microsoft.com/office/e3313c02-51cc-4963-aae6-31442d9ec307) | Returns the sum of the squares of the arguments |
| [SYD function](https://support.microsoft.com/office/069f8106-b60b-4ca2-98e0-2a0f206bdb27) | Returns the sum-of-years' digits depreciation of an asset for a specified period |
| [T function](https://support.microsoft.com/office/fb83aeec-45e7-4924-af95-53e073541228) | Converts its arguments to text |
| [T.DIST function](https://support.microsoft.com/office/4329459f-ae91-48c2-bba8-1ead1c6c21b2) | Returns the Percentage Points (probability) for the Student t-distribution |
| [T.DIST.2T function](https://support.microsoft.com/office/198e9340-e360-4230-bd21-f52f22ff5c28) | Returns the Percentage Points (probability) for the Student t-distribution |
| [T.DIST.RT function](https://support.microsoft.com/office/20a30020-86f9-4b35-af1f-7ef6ae683eda) | Returns the Student's t-distribution |
| [T.INV function](https://support.microsoft.com/office/2908272b-4e61-4942-9df9-a25fec9b0e2e) | Returns the t-value of the Student's t-distribution as a function of the probability and the degrees of freedom |
| [T.INV.2T function](https://support.microsoft.com/office/ce72ea19-ec6c-4be7-bed2-b9baf2264f17) | Returns the inverse of the Student's t-distribution |
| [TAN function](https://support.microsoft.com/office/08851a40-179f-4052-b789-d7f699447401) | Returns the tangent of a number |
| [TANH function](https://support.microsoft.com/office/017222f0-a0c3-4f69-9787-b3202295dc6c) | Returns the hyperbolic tangent of a number |
| [TBILLEQ function](https://support.microsoft.com/office/2ab72d90-9b4d-4efe-9fc2-0f81f2c19c8c) | Returns the bond-equivalent yield for a Treasury bill |
| [TBILLPRICE function](https://support.microsoft.com/office/eacca992-c29d-425a-9eb8-0513fe6035a2) | Returns the price per $100 face value for a Treasury bill |
| [TBILLYIELD function](https://support.microsoft.com/office/6d381232-f4b0-4cd5-8e97-45b9c03468ba) | Returns the yield for a Treasury bill |
| [TEXT function](https://support.microsoft.com/office/20d5ac4d-7b94-49fd-bb38-93d29371225c) | Formats a number and converts it to text |
| [TIME function](https://support.microsoft.com/office/9a5aff99-8f7d-4611-845e-747d0b8d5457) | Returns the serial number of a particular time |
| [TIMEVALUE function](https://support.microsoft.com/office/0b615c12-33d8-4431-bf3d-f3eb6d186645) | Converts a time in the form of text to a serial number |
| [TODAY function](https://support.microsoft.com/office/5eb3078d-a82c-4736-8930-2f51a028fdd9) | Returns the serial number of today's date |
| [TRIM function](https://support.microsoft.com/office/410388fa-c5df-49c6-b16c-9e5630b479f9) | Removes spaces from text |
| [TRIMMEAN function](https://support.microsoft.com/office/d90c9878-a119-4746-88fa-63d988f511d3) | Returns the mean of the interior of a data set |
| [TRUE function](https://support.microsoft.com/office/7652c6e3-8987-48d0-97cd-ef223246b3fb) | Returns the logical value `TRUE` |
| [TRUNC function](https://support.microsoft.com/office/8b86a64c-3127-43db-ba14-aa5ceb292721) | Truncates a number to an integer |
| [TYPE function](https://support.microsoft.com/office/45b4e688-4bc3-48b3-a105-ffa892995899) | Returns a number indicating the data type of a value |
| [UNICHAR function](https://support.microsoft.com/office/ffeb64f5-f131-44c6-b332-5cd72f0659b8) | Returns the Unicode character that is references by the given numeric value |
| [UNICODE function](https://support.microsoft.com/office/adb74aaa-a2a5-4dde-aff6-966e4e81f16f) | Returns the number (code point) that corresponds to the first character of the text |
| [UPPER function](https://support.microsoft.com/office/c11f29b3-d1a3-4537-8df6-04d0049963d6) | Converts text to uppercase |
| [VALUE function](https://support.microsoft.com/office/257d0108-07dc-437d-ae1c-bc2d3953d8c2) | Converts a text argument to a number |
| [VAR.P function](https://support.microsoft.com/office/73d1285c-108c-4843-ba5d-a51f90656f3a) | Calculates variance based on the entire population |
| [VAR.S function](https://support.microsoft.com/office/913633de-136b-449d-813e-65a00b2b990b) | Estimates variance based on a sample |
| [VARA function](https://support.microsoft.com/office/3de77469-fa3a-47b4-85fd-81758a1e1d07) | Estimates variance based on a sample, including numbers, text, and logical values |
| [VARPA function](https://support.microsoft.com/office/59a62635-4e89-4fad-88ac-ce4dc0513b96) | Calculates variance based on the entire population, including numbers, text, and logical values |
| [VDB function](https://support.microsoft.com/office/dde4e207-f3fa-488d-91d2-66d55e861d73) | Returns the depreciation of an asset for a specified or partial period by using a declining balance method |
| [VLOOKUP function](https://support.microsoft.com/office/0bbc8083-26fe-4963-8ab8-93a18ad188a1) | Looks in the first column of an array and moves across the row to return the value of a cell |
| [WEEKDAY function](https://support.microsoft.com/office/60e44483-2ed1-439f-8bd0-e404c190949a) | Converts a serial number to a day of the week |
| [WEEKNUM function](https://support.microsoft.com/office/e5c43a03-b4ab-426c-b411-b18c13c75340) | Converts a serial number to a number representing where the week falls numerically with a year |
| [WEIBULL.DIST function](https://support.microsoft.com/office/4e783c39-9325-49be-bbc9-a83ef82b45db) | Returns the Weibull distribution |
| [WORKDAY function](https://support.microsoft.com/office/f764a5b7-05fc-4494-9486-60d494efbf33) | Returns the serial number of the date before or after a specified number of workdays |
| [WORKDAY.INTL function](https://support.microsoft.com/office/a378391c-9ba7-4678-8a39-39611a9bf81d) | Returns the serial number of the date before or after a specified number of workdays using parameters to indicate which and how many days are weekend days |
| [XIRR function](https://support.microsoft.com/office/de1242ec-6477-445b-b11b-a303ad9adc9d) | Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic |
| [XNPV function](https://support.microsoft.com/office/1b42bbf6-370f-4532-a0eb-d67c16b664b7) | Returns the net present value for a schedule of cash flows that is not necessarily periodic |
| [XOR function](https://support.microsoft.com/office/1548d4c2-5e47-4f77-9a92-0533bba14f37) | Returns a logical exclusive OR of all arguments |
| [YEAR function](https://support.microsoft.com/office/c64f017a-1354-490d-981f-578e8ec8d3b9) | Converts a serial number to a year |
| [YEARFRAC function](https://support.microsoft.com/office/3844141e-c76d-4143-82b6-208454ddc6a8) | Returns the year fraction representing the number of whole days between start_date and end_date |
| [YIELD function](https://support.microsoft.com/office/f5f5ca43-c4bd-434f-8bd2-ed3c9727a4fe) | Returns the yield on a security that pays periodic interest |
| [YIELDDISC function](https://support.microsoft.com/office/a9dbdbae-7dae-46de-b995-615faffaaed7) | Returns the annual yield for a discounted security; for example, a Treasury bill |
| [YIELDMAT function](https://support.microsoft.com/office/ba7d1809-0d33-4bcb-96c7-6c56ec62ef6f) | Returns the annual yield of a security that pays interest at maturity |
| [Z.TEST function](https://support.microsoft.com/office/d633d5a3-2031-4614-a016-92180ad82bee) | Returns the one-tailed probability-value of a z-test |

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Functions Class (JavaScript API for Excel)](/javascript/api/excel/excel.functions)
- [Workbook Functions Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook#excel-excel-workbook-functions-member)
