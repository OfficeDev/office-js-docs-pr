---
title: Create a Copilot skill for Excel that uses custom functions (preview)
description: Learn how to create an Excel Copilot skill that inserts formulas that call custom functions.
ms.date: 10/09/2026
ms.topic: tutorial
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Create a Copilot skill for Excel that uses custom functions (preview)

In this tutorial, you create a Copilot skill for Excel that uses Office.js to compare quotes in different currencies. The skill finds a table of vendor quotes, converts each quote to the workbook's reporting currency, and classifies each quote against its budget. The calculations use custom functions that get the latest exchange rate and determine the budget status.

> [!NOTE]
> - This tutorial assumes that you're familiar with [Overview of Copilot skills for Excel (preview)](excel-skills.md), [Create custom functions in Excel](custom-functions-overview.md), and [Build plugins for Copilot Cowork](/microsoft-365/copilot/cowork/cowork-plugin-development). Although the latter article is in the context of Cowork, the general packaging, manifest, icon, installation, and publishing guidance in that article also applies to Excel skills. This tutorial focuses on the Excel-specific pieces and uses the same plugin package model described in that article.
> - Custom skills for Excel are in preview and require version 2608 (Build 20305.20002) or later on Windows, or version 16.112.26070718 or later on Mac. They're available in only the **Beta** and **Current Channel (Preview)** channels, through the [Microsoft 365 Insider Program](https://techcommunity.microsoft.com/kb/microsoft-365-insider-kb/microsoft-365-insider-handbook/4401152). Don't use them in a production Copilot extension.

## What you'll build

You'll build a Microsoft 365 app package with one skill and two custom functions. The Office.js script validates the workbook and inserts formulas, formatting, and conditional formatting. The custom functions provide the latest exchange rate and budget classification used by those formulas. A **SKILL.md** file tells Copilot when to use the script and how to explain the result to the user.

The custom functions are:

```excel
=PURCHASEPLANNER.FXRATE(sourceCurrency, reportingCurrency)
=PURCHASEPLANNER.BUDGETSTATUS(convertedQuote, budgetLimit, warningThreshold)
```

## Prerequisites

Before you start, make sure you have the following.

- [Node.js](https://nodejs.org/) version 22 or later.
- The Microsoft 365 Agents Toolkit CLI as described in "Step 7: Test" of [Build plugins for Copilot Cowork](/microsoft-365/copilot/cowork/cowork-plugin-development#step-7-test).
- A Microsoft 365 work or developer account with access to Copilot in Excel.

## Task 1: Create the project folders

Create the following folders.

```text
international-purchase-planner/
|-- appPackage/
|   |-- assets/
|   |-- skills/
|       |-- international-purchase-planner/
|           |-- resources/
|           |-- scripts/
|-- src/
    |-- functions/
    |-- taskpane/
```

## Task 2: Create SKILL.md

1. In the **appPackage/skills/international-purchase-planner/** folder, create a file called **SKILL.md**.
1. Add the following YAML frontmatter (including the two `---` lines) to the very top of the file. Note that the `metadata.tags` value includes `excel`, which helps identify the skill as Excel-oriented for purposes of skill discovery.

    ```yaml
    ---
    name: international-purchase-planner
    description: |
      Use this skill in Excel when the user asks to compare complete vendor quotes
      in one reporting currency or identify quotes that are near or over budget.
    metadata:
      category: Excel analysis
      version: 1.0.0
      tags: excel, office-js, custom-functions, currency, purchasing, budgeting
    ---
    ```

    > [!IMPORTANT]
    > The value of the `name` property must exactly match the name of the child folder directly under the **skills** folder.

1. Below the frontmatter, add the following Markdown to focus Copilot on the purpose and resources of the skill. You create the two resource files in later steps.

    ```md
    # International Purchase Planner

    Convert complete vendor quotes with the latest exchange rates and classify each converted quote against its budget.

    ## Reference resources

    Before running the script, consult:

    - `resources/workbook-data-guardrails.md`
    - `resources/excel-vs-agent-execution.md`
    ```

1. Below the resources, add the following workflow section. You create the JavaScript file that calls Office.js in a later step.

    ```md
    ## Workflow

    1. Confirm that the current context is Excel.
    2. Confirm that the workbook contains a valid table named `Settings`.
    3. Run `scripts/prepare-purchase-comparison.js`.
    4. Report the quote table name, reporting currency, warning threshold, and number of processed rows.
    ```

1. Below the workflow, add the following sections that give instructions about the output.

    ```md
    ## Workbook output

    The script must use the first qualifying quote table in workbook order. It must add or update these calculated columns:

    - `Exchange rate`
    - `Converted quote`
    - `Budget status`

    The script must insert formulas, not static results. After calculation finishes, it must format numeric output, apply status-based conditional formatting, and autofit each output column to its header and calculated values. It must not create a summary or modify source columns.

    ## Copilot chat output

    Report:

    - the name of the processed table;
    - the reporting currency;
    - the warning threshold; and
    - the number of processed quote rows.

    If no table qualifies or `Settings` is invalid, report the exact error returned by the script.
    ```

1. Below the output instructions, add the following guidance to Copilot.

    ```md
    ## Common pitfalls to avoid

    - Do not process any table other than the first qualifying table in workbook order.
    - Do not infer or prompt for settings.
    - Do not process incomplete or incorrectly typed quote data.
    - Do not replace formulas with static values.
    - Do not run the Office.js script outside Excel.
    ```

## Task 3: Add workbook data guardrails

1. In the **appPackage/skills/international-purchase-planner/resources** folder, create a Markdown file named **workbook-data-guardrails.md**. The resource file keeps scenario-specific workbook assumptions out of **SKILL.md**, while still giving Copilot guardrails for when it invokes the script.
1. Give the file the following content.

    ```md
    # Workbook data guardrails

    ## Settings table

    The workbook must contain an Excel table named `Settings` with columns named `Setting` and `Value`. It must contain exactly one nonblank value for each required setting:

    | Setting | Required value |
    | --- | --- |
    | `Reporting currency` | A three-letter currency code |
    | `Warning threshold` | A number from 0 through 1 |

    Do not create, repair, infer, or prompt for these settings. Stop if the table or either value is invalid.

    ## Qualifying quote table

    Search worksheets and their tables in workbook order. Ignore the `Settings` table and use the first table that qualifies. Never prioritize the current selection.

    A qualifying table:

    - has at least one data row;
    - has exactly one column for each required header: `Item`, `Vendor`, `Quote amount`, `Quote currency`, and `Budget limit`;
    - has a nonblank string in every `Item` and `Vendor` cell;
    - has a finite nonnegative number in every `Quote amount` cell;
    - has a three-letter currency code in every `Quote currency` cell; and
    - has a finite positive number in every `Budget limit` cell.

    Header comparison is case-insensitive after trimming. A table with an incomplete or incorrectly typed required cell does not qualify. Continue searching for the next table; do not repair it or report individual rows.

    ## Change boundaries

    - Add or update only `Exchange rate`, `Converted quote`, and `Budget status`.
    - Do not rename or overwrite source columns.
    - Stop if an output column name conflicts with a source column or duplicated header.
    - Use only the latest exchange rate.
    - Handle only one quote table per invocation.
    - Do not create a summary.
    ```

## Task 4: Add Excel execution guidance

1. In the **appPackage/skills/international-purchase-planner/resources** folder, create a Markdown file named **excel-vs-agent-execution.md**. The rules in this file are needed because the skill may be accessible in Copilot outside the context of Excel, in which case the Office.js APIs can't run. If the skill runs outside Excel, Copilot shouldn't pretend that it changed the workbook.
1. Give it the following content.

    ```md
    # Excel vs non-Excel execution guidance

    ## Inside Excel

    - Use the workbook as the source of truth.
    - Run `scripts/prepare-purchase-comparison.js` only for requests to convert complete vendor quotes or identify quotes near or over budget.
    - Let the script perform all workbook changes.

    ## Outside Excel

    - Do not run the script or claim that a workbook was changed.
    - Explain that the skill must be invoked in Copilot in Excel with the workbook open.

    ## Quality rules

    - Do not run the script for general purchasing, currency, or budgeting advice.
    - Do not ask for missing workbook values.
    - Report script errors without claiming partial success.
    - Never claim that a workbook update succeeded unless the script completed without error.
    ```

## Task 5: Add the custom functions

1. In the **src/functions** folder, create a TypeScript file named **functions.ts**.
1. Add the following code. The `FXRATE` function calls the Frankfurter currency API for the latest exchange rate. If the source and reporting currencies are the same, it returns `1` without calling the service.

    ```typescript
    /* global CustomFunctions, fetch, Response */

    const FUNCTION_CURRENCY_CODE_PATTERN = /^[A-Z]{3}$/;
    const FRANKFURTER_API = "https://api.frankfurter.dev/v2";

    interface FrankfurterRate {
      rate?: unknown;
    }

    /**
     * Returns the latest exchange rate between two currencies.
     * @customfunction FXRATE
     * @param sourceCurrency Three-letter source currency code.
     * @param reportingCurrency Three-letter reporting currency code.
     * @returns The latest exchange rate.
     */
    export function fxrate(sourceCurrency: string, reportingCurrency: string): Promise<number> {
      const source = normalizeFunctionCurrency(sourceCurrency, "source");
      const reporting = normalizeFunctionCurrency(reportingCurrency, "reporting");

      if (source === reporting) {
        return Promise.resolve(1);
      }

      return fetch(`${FRANKFURTER_API}/rate/${source.toLowerCase()}/${reporting.toLowerCase()}`).then(
        (response: Response) => {
          if (!response.ok) {
            const errorCode =
              response.status === 400 || response.status === 404 || response.status === 422
                ? CustomFunctions.ErrorCode.invalidValue
                : CustomFunctions.ErrorCode.notAvailable;
            throw new CustomFunctions.Error(
              errorCode,
              "No exchange rate is available for this currency pair."
            );
          }

          return response.json().then(
            (payload: FrankfurterRate) => {
              if (
                typeof payload.rate !== "number" ||
                !Number.isFinite(payload.rate) ||
                payload.rate <= 0
              ) {
                throw new CustomFunctions.Error(
                  CustomFunctions.ErrorCode.notAvailable,
                  "The exchange-rate service did not return a valid rate."
                );
              }
              return payload.rate;
            },
            () => {
              throw new CustomFunctions.Error(
                CustomFunctions.ErrorCode.notAvailable,
                "The exchange-rate service returned an invalid response."
              );
            }
          );
        },
        () => {
          throw new CustomFunctions.Error(
            CustomFunctions.ErrorCode.notAvailable,
            "The exchange-rate service could not be reached."
          );
        }
      );
    }
    ```

1. Add the following function to **functions.ts**. `BUDGETSTATUS` returns one of three values based on the converted quote and the warning threshold.

    ```typescript
    /**
     * Classifies a converted quote relative to its budget.
     * @customfunction BUDGETSTATUS
     * @param convertedQuote Converted quote amount in the reporting currency.
     * @param budgetLimit Positive budget limit in the reporting currency.
     * @param warningThreshold Warning threshold from zero through one.
     * @returns Within budget, Near limit, or Over budget.
     */
    export function budgetStatus(
      convertedQuote: number,
      budgetLimit: number,
      warningThreshold: number
    ): string {
      if (!Number.isFinite(convertedQuote) || convertedQuote < 0) {
        throw new CustomFunctions.Error(
          CustomFunctions.ErrorCode.invalidValue,
          "The converted quote must be a nonnegative number."
        );
      }
      if (!Number.isFinite(budgetLimit) || budgetLimit <= 0) {
        throw new CustomFunctions.Error(
          CustomFunctions.ErrorCode.invalidValue,
          "The budget limit must be a positive number."
        );
      }
      if (!Number.isFinite(warningThreshold) || warningThreshold < 0 || warningThreshold > 1) {
        throw new CustomFunctions.Error(
          CustomFunctions.ErrorCode.invalidValue,
          "The warning threshold must be from zero through one."
        );
      }

      if (convertedQuote > budgetLimit) {
        return "Over budget";
      }
      if (convertedQuote >= budgetLimit * (1 - warningThreshold)) {
        return "Near limit";
      }
      return "Within budget";
    }
    ```

1. Add the following helper function to **functions.ts**.

    ```typescript
    function normalizeFunctionCurrency(value: string, label: string): string {
      if (typeof value !== "string") {
        throw new CustomFunctions.Error(
          CustomFunctions.ErrorCode.invalidValue,
          `The ${label} currency must be a three-letter code.`
        );
      }

      const currency = value.trim().toUpperCase();
      if (!FUNCTION_CURRENCY_CODE_PATTERN.test(currency)) {
        throw new CustomFunctions.Error(
          CustomFunctions.ErrorCode.invalidValue,
          `The ${label} currency must be a three-letter code.`
        );
      }
      return currency;
    }
    ```

1. In the **src/functions** folder, create a file named **functions.html**, and give it the following content. The custom functions runtime uses this page.

    ```html
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8" />
      <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
      <meta http-equiv="Expires" content="0" />
      <title></title>
      <script
        src="https://appsforoffice.microsoft.com/lib/1.1/hosted/custom-functions-runtime.js"
        type="text/javascript">
      </script>
    </head>
    <body>
    </body>
    </html>
    ```

## Task 6: Add the task pane activation surface

A task pane runtime is required because custom functions in a JavaScript-only runtime aren't registered unless the add-in also has a browser-based runtime. The task pane in this tutorial is only an activation surface. It doesn't implement purchase-planning features or add a ribbon button.

1. In the **src/taskpane** folder, create a file named **taskpane.ts**, and give it the following content.

    ```typescript
    /* global Office */

    Office.onReady();
    ```

1. In the same folder, create a file named **taskpane.html**, and give it the following content.

    ```html
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8" />
      <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
      <meta name="viewport" content="width=device-width, initial-scale=1" />
      <title>Purchase Planner</title>
      <script src=" https://officeapis.public.onecdn.static.microsoft/1/office.js"></script>
    </head>
    <body>
      <main>
        <h1>Purchase Planner</h1>
        <p>The Purchase Planner custom functions are available in this workbook.</p>
      </main>
    </body>
    </html>
    ```

## Task 7: Add the script that calls Office.js

1. In the **appPackage/skills/international-purchase-planner/scripts** folder, create a JavaScript file named **prepare-purchase-comparison.js**.
1. Add the following constants and main code. Note the following about this code.

    - The script consists of a call to `Excel.run`. Copilot creates the runtime, initializes Office.js, and executes the skill script.
    - The script calls helper methods that you create in later steps.
    - The script inserts formulas that call the two custom functions. It doesn't calculate exchange rates or budget status itself.
    - The script waits until every formula returns a final value before it autofits the output columns and reports success.
    - The function is parameterless because the workbook itself supplies all input.

    > [!NOTE]
    > The preview of Office.js-based skills doesn't support passing parameters to the functions that call Office.js. We're working hard to provide this support in the future.

    ```javascript
    const REQUIRED_QUOTE_COLUMNS = [
      "Item",
      "Vendor",
      "Quote amount",
      "Quote currency",
      "Budget limit",
    ];
    const OUTPUT_COLUMNS = ["Exchange rate", "Converted quote", "Budget status"];
    const CURRENCY_CODE_PATTERN = /^[A-Z]{3}$/;
    const CALCULATION_TIMEOUT_MS = 120000;
    const CALCULATION_POLL_INTERVAL_MS = 250;

    await Excel.run(async (context) => {
      const settings = await readSettings(context);
      const quoteTable = await findFirstQualifyingTable(context);

      if (!quoteTable) {
        throw new Error(
          "No qualifying quote table was found. Every required source cell must contain a value of the appropriate data type."
        );
      }

      await addMissingOutputColumns(context, quoteTable.table, quoteTable.columnNames);

      const exchangeRateColumn = quoteTable.table.columns.getItem(
        quoteTable.columnNames.get(normalizeHeader("Exchange rate"))
      );
      const convertedQuoteColumn = quoteTable.table.columns.getItem(
        quoteTable.columnNames.get(normalizeHeader("Converted quote"))
      );
      const budgetStatusColumn = quoteTable.table.columns.getItem(
        quoteTable.columnNames.get(normalizeHeader("Budget status"))
      );

      const exchangeRateRange = exchangeRateColumn.getDataBodyRange();
      const convertedQuoteRange = convertedQuoteColumn.getDataBodyRange();
      const budgetStatusRange = budgetStatusColumn.getDataBodyRange();

      exchangeRateRange.formulas = repeatFormula(
        quoteTable.rowCount,
        `=PURCHASEPLANNER.FXRATE([@[Quote currency]],"${settings.reportingCurrency}")`
      );
      convertedQuoteRange.formulas = repeatFormula(
        quoteTable.rowCount,
        "=[@[Quote amount]]*[@[Exchange rate]]"
      );
      budgetStatusRange.formulas = repeatFormula(
        quoteTable.rowCount,
        `=PURCHASEPLANNER.BUDGETSTATUS([@[Converted quote]],[@[Budget limit]],${settings.warningThreshold})`
      );

      exchangeRateRange.numberFormat = repeatFormat(quoteTable.rowCount, "0.0000");
      convertedQuoteRange.numberFormat = repeatFormat(
        quoteTable.rowCount,
        `[$${settings.reportingCurrency}]#,##0.00`
      );

      budgetStatusRange.conditionalFormats.clearAll();
      addStatusFormat(budgetStatusRange, "Within budget", "#E2F0D9", "#375623");
      addStatusFormat(budgetStatusRange, "Near limit", "#FFF2CC", "#7F6000");
      addStatusFormat(budgetStatusRange, "Over budget", "#F4CCCC", "#9C0006");

      await context.sync();
      await waitForFinalValues(
        context,
        exchangeRateRange,
        convertedQuoteRange,
        budgetStatusRange
      );

      exchangeRateColumn.getRange().format.autofitColumns();
      convertedQuoteColumn.getRange().format.autofitColumns();
      budgetStatusColumn.getRange().format.autofitColumns();
      await context.sync();

      return JSON.stringify({
        table: quoteTable.table.name,
        reportingCurrency: settings.reportingCurrency,
        warningThreshold: settings.warningThreshold,
        processedRows: quoteTable.rowCount,
      });
    ```

1. Add the following helper method inside the callback to `Excel.run`. It validates the **Settings** table and returns the reporting currency and warning threshold.

    ```javascript
      async function readSettings(excelContext) {
        const table = excelContext.workbook.tables.getItemOrNullObject("Settings");
        table.load(["isNullObject", "name"]);
        await excelContext.sync();

        if (table.isNullObject) {
          throw new Error('The workbook must contain a table named "Settings".');
        }

        const columns = table.columns;
        const rows = table.rows;
        columns.load("items/name");
        rows.load("count");
        await excelContext.sync();

        const columnMap = createColumnMap(columns.items);
        if (
          !columnMap.has(normalizeHeader("Setting")) ||
          !columnMap.has(normalizeHeader("Value")) ||
          rows.count === 0
        ) {
          throw new Error(
            'The "Settings" table must have nonempty "Setting" and "Value" columns.'
          );
        }

        const settingRange = table.columns
          .getItem(columnMap.get(normalizeHeader("Setting")))
          .getDataBodyRange();
        const valueRange = table.columns
          .getItem(columnMap.get(normalizeHeader("Value")))
          .getDataBodyRange();
        settingRange.load("values");
        valueRange.load("values");
        await excelContext.sync();

        const valuesBySetting = new Map();
        for (let rowIndex = 0; rowIndex < rows.count; rowIndex += 1) {
          const setting = settingRange.values[rowIndex][0];
          const value = valueRange.values[rowIndex][0];
          if (typeof setting !== "string" || setting.trim() === "" || isBlank(value)) {
            throw new Error(
              'Every row in the "Settings" table must contain a setting and value.'
            );
          }

          const key = setting.trim().toLowerCase();
          if (valuesBySetting.has(key)) {
            throw new Error(
              `The "Settings" table contains duplicate "${setting.trim()}" rows.`
            );
          }
          valuesBySetting.set(key, value);
        }

        const reportingValue = valuesBySetting.get("reporting currency");
        const reportingCurrency =
          typeof reportingValue === "string"
            ? reportingValue.trim().toUpperCase()
            : "";
        if (!CURRENCY_CODE_PATTERN.test(reportingCurrency)) {
          throw new Error(
            'The "Reporting currency" setting must be a three-letter currency code.'
          );
        }

        const warningThreshold = valuesBySetting.get("warning threshold");
        if (
          typeof warningThreshold !== "number" ||
          !Number.isFinite(warningThreshold) ||
          warningThreshold < 0 ||
          warningThreshold > 1
        ) {
          throw new Error(
            'The "Warning threshold" setting must be a number from zero through one.'
          );
        }

        return { reportingCurrency, warningThreshold };
      }
    ```

1. Add the following helper method inside the callback to `Excel.run`. It searches worksheets and tables in workbook order and returns the first complete quote table with correctly typed source data.

    ```javascript
      async function findFirstQualifyingTable(excelContext) {
        const worksheets = excelContext.workbook.worksheets;
        worksheets.load("items/name");
        await excelContext.sync();

        for (const worksheet of worksheets.items) {
          const tables = worksheet.tables;
          tables.load("items/name");
          await excelContext.sync();

          for (const table of tables.items) {
            if (table.name.toLowerCase() === "settings") {
              continue;
            }

            const columns = table.columns;
            const rows = table.rows;
            columns.load("items/name");
            rows.load("count");
            await excelContext.sync();

            if (rows.count === 0) {
              continue;
            }

            const columnMap = createColumnMap(columns.items);
            if (
              !hasRequiredColumns(columnMap) ||
              hasDuplicateRequiredHeaders(columns.items)
            ) {
              continue;
            }

            const ranges = {};
            for (const requiredName of REQUIRED_QUOTE_COLUMNS) {
              const actualName = columnMap.get(normalizeHeader(requiredName));
              ranges[requiredName] = table.columns
                .getItem(actualName)
                .getDataBodyRange();
              ranges[requiredName].load("values");
            }
            await excelContext.sync();

            if (hasValidQuoteData(ranges, rows.count)) {
              table.load("name");
              await excelContext.sync();
              return { table, columnNames: columnMap, rowCount: rows.count };
            }
          }
        }

        return null;
      }
    ```

1. Add the following helper methods inside the callback to `Excel.run`. They add missing output columns and validate the source values.

    ```javascript
      async function addMissingOutputColumns(excelContext, table, existingColumns) {
        for (const outputName of OUTPUT_COLUMNS) {
          const normalizedName = normalizeHeader(outputName);
          if (!existingColumns.has(normalizedName)) {
            table.columns.add(null, null, outputName);
            existingColumns.set(normalizedName, outputName);
          }
        }
        await excelContext.sync();
      }

      function hasValidQuoteData(ranges, rowCount) {
        for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
          const item = ranges["Item"].values[rowIndex][0];
          const vendor = ranges["Vendor"].values[rowIndex][0];
          const quoteAmount = ranges["Quote amount"].values[rowIndex][0];
          const quoteCurrency = ranges["Quote currency"].values[rowIndex][0];
          const budgetLimit = ranges["Budget limit"].values[rowIndex][0];

          if (
            typeof item !== "string" ||
            item.trim() === "" ||
            typeof vendor !== "string" ||
            vendor.trim() === "" ||
            typeof quoteAmount !== "number" ||
            !Number.isFinite(quoteAmount) ||
            quoteAmount < 0 ||
            typeof quoteCurrency !== "string" ||
            !CURRENCY_CODE_PATTERN.test(quoteCurrency.trim().toUpperCase()) ||
            typeof budgetLimit !== "number" ||
            !Number.isFinite(budgetLimit) ||
            budgetLimit <= 0
          ) {
            return false;
          }
        }
        return true;
      }
    ```

1. Add the following header and array helper methods inside the callback to `Excel.run`.

    ```javascript
      function createColumnMap(columns) {
        const result = new Map();
        for (const column of columns) {
          const normalizedName = normalizeHeader(column.name);
          if (!result.has(normalizedName)) {
            result.set(normalizedName, column.name);
          }
        }
        return result;
      }

      function hasRequiredColumns(columnMap) {
        return REQUIRED_QUOTE_COLUMNS.every((name) =>
          columnMap.has(normalizeHeader(name))
        );
      }

      function hasDuplicateRequiredHeaders(columns) {
        const counts = new Map();
        for (const column of columns) {
          const normalizedName = normalizeHeader(column.name);
          counts.set(normalizedName, (counts.get(normalizedName) || 0) + 1);
        }
        return [...REQUIRED_QUOTE_COLUMNS, ...OUTPUT_COLUMNS].some(
          (name) => (counts.get(normalizeHeader(name)) || 0) > 1
        );
      }

      function normalizeHeader(value) {
        return value.trim().toLowerCase();
      }

      function isBlank(value) {
        return value === null ||
          value === undefined ||
          (typeof value === "string" && value.trim() === "");
      }

      function repeatFormula(rowCount, formula) {
        return Array.from({ length: rowCount }, () => [formula]);
      }

      function repeatFormat(rowCount, format) {
        return Array.from({ length: rowCount }, () => [format]);
      }
    ```

1. Add the following helper methods inside the callback to `Excel.run`. They wait for the custom functions to finish and apply conditional formatting. The final `});` closes the call to `Excel.run`.

    ```javascript
      async function waitForFinalValues(
        excelContext,
        exchangeRates,
        convertedQuotes,
        budgetStatuses
      ) {
        const application = excelContext.workbook.application;
        application.calculate(Excel.CalculationType.full);
        await excelContext.sync();

        const deadline = Date.now() + CALCULATION_TIMEOUT_MS;
        while (Date.now() < deadline) {
          exchangeRates.load("values");
          convertedQuotes.load("values");
          budgetStatuses.load("values");
          await excelContext.sync();

          const ratesAreFinal = exchangeRates.values.every(
            ([value]) =>
              typeof value === "number" && Number.isFinite(value) && value > 0
          );
          const quotesAreFinal = convertedQuotes.values.every(
            ([value]) =>
              typeof value === "number" && Number.isFinite(value) && value >= 0
          );
          const statusesAreFinal = budgetStatuses.values.every(([value]) =>
            ["Within budget", "Near limit", "Over budget"].includes(value)
          );

          if (ratesAreFinal && quotesAreFinal && statusesAreFinal) {
            return;
          }
          await delay(CALCULATION_POLL_INTERVAL_MS);
        }

        throw new Error(
          "The purchase comparison formulas did not return final values within 120 seconds."
        );
      }

      function delay(milliseconds) {
        return new Promise((resolve) => setTimeout(resolve, milliseconds));
      }

      function addStatusFormat(range, status, fillColor, fontColor) {
        const format = range.conditionalFormats.add(
          Excel.ConditionalFormatType.cellValue
        );
        format.cellValue.rule = {
          formula1: `"${status}"`,
          operator: Excel.ConditionalCellValueOperator.equalTo,
        };
        format.cellValue.format.fill.color = fillColor;
        format.cellValue.format.font.color = fontColor;
      }
    });
    ```

## Task 8: Configure the build

1. From the **international-purchase-planner** root folder, run the following command to create a **package.json** file.

    ```command line
    npm init -y
    ```

1. Install the runtime dependencies.

    ```command line
    npm install core-js regenerator-runtime
    ```

1. Install the development dependencies.

    ```command line
    npm install --save-dev @babel/core @babel/preset-env @babel/preset-typescript @types/custom-functions-runtime @types/office-js @types/office-runtime babel-loader copy-webpack-plugin custom-functions-metadata-plugin html-loader html-webpack-plugin office-addin-dev-certs office-addin-lint office-addin-manifest office-addin-prettier-config source-map-loader typescript webpack webpack-cli webpack-dev-server
    ```

1. In **package.json**, replace the "scripts" object with the following.

    ```json
    "scripts": {
      "build": "webpack --mode production",
      "build:dev": "webpack --mode development",
      "dev-server": "npm run build:dev && webpack serve --mode development",
      "lint": "office-addin-lint check",
      "lint:fix": "office-addin-lint fix",
      "prettier": "office-addin-lint prettier",
      "validate": "office-addin-manifest validate appPackage/manifest.json"
    }
    ```

1. Add the following properties at the root of **package.json**.

    ```json
    "engines": {
      "node": ">=22"
    },
    "prettier": "office-addin-prettier-config",
    "browserslist": [
      "last 2 versions",
      "ie 11"
    ]
    ```

1. In the project root, create **babel.config.json** with the following content.

    ```json
    {
      "presets": [
        [
          "@babel/preset-env",
          {
            "targets": {
              "ie": "11"
            }
          }
        ],
        "@babel/preset-typescript"
      ]
    }
    ```

1. In the project root, create **tsconfig.json** with the following content.

    ```json
    {
      "compilerOptions": {
        "allowJs": true,
        "baseUrl": ".",
        "esModuleInterop": true,
        "experimentalDecorators": true,
        "noEmitOnError": true,
        "outDir": "lib",
        "sourceMap": true,
        "target": "es5",
        "lib": [
          "es2015",
          "dom"
        ]
      },
      "exclude": [
        "node_modules",
        "dist",
        "lib",
        "lib-amd"
      ]
    }
    ```

1. In the project root, create **webpack.config.js**. Give it the following content.

    ```javascript
    /* eslint-disable no-undef */

    const devCerts = require("office-addin-dev-certs");
    const CopyWebpackPlugin = require("copy-webpack-plugin");
    const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
    const HtmlWebpackPlugin = require("html-webpack-plugin");
    const path = require("path");

    const urlDev = "https://localhost:3000/";
    const urlProd = "https://www.contoso.com/";

    async function getHttpsOptions() {
      const httpsOptions = await devCerts.getHttpsServerOptions();
      return {
        ca: httpsOptions.ca,
        key: httpsOptions.key,
        cert: httpsOptions.cert,
      };
    }

    module.exports = async (env, options) => {
      const dev = options.mode === "development";
      return {
        devtool: "source-map",
        entry: {
          functions: "./src/functions/functions.ts",
          taskpane: [
            "./src/taskpane/taskpane.ts",
            "./src/taskpane/taskpane.html",
          ],
        },
        output: {
          clean: true,
        },
        resolve: {
          extensions: [".ts", ".html", ".js"],
        },
        module: {
          rules: [
            {
              test: /\.ts$/,
              exclude: /node_modules/,
              use: {
                loader: "babel-loader",
                options: {
                  presets: ["@babel/preset-typescript"],
                },
              },
            },
            {
              test: /\.html$/,
              exclude: /node_modules/,
              use: "html-loader",
            },
          ],
        },
        plugins: [
          new CustomFunctionsMetadataPlugin({
            output: "functions.json",
            input: "./src/functions/functions.ts",
          }),
          new HtmlWebpackPlugin({
            filename: "functions.html",
            template: "./src/functions/functions.html",
            chunks: ["functions"],
          }),
          new HtmlWebpackPlugin({
            filename: "taskpane.html",
            template: "./src/taskpane/taskpane.html",
            chunks: ["taskpane"],
          }),
          new CopyWebpackPlugin({
            patterns: [
              {
                from: "appPackage/assets/*",
                to: "assets/[name][ext][query]",
              },
              {
                from: "appPackage/manifest*.json",
                to: "[name][ext]",
                transform(content) {
                  return dev
                    ? content
                    : content.toString().replace(new RegExp(urlDev, "g"), urlProd);
                },
              },
              {
                from: "appPackage/skills",
                to: "skills",
              },
            ],
          }),
        ],
        devServer: {
          static: {
            directory: path.join(__dirname, "dist"),
            publicPath: "/public",
          },
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Cache-Control": "no-store",
          },
          server: {
            type: "https",
            options:
              env.WEBPACK_BUILD || options.https !== undefined
                ? options.https
                : await getHttpsOptions(),
          },
          port: 3000,
        },
      };
    };
    ```

## Task 9: Create the manifest

1. Create a **manifest.json** file in the **appPackage** folder.
1. Give the file the following content. You add the two icon files in a later step. Note that the `"validDomains"` array includes `"api.frankfurter.dev"` which is the source that the custom functions use for currency exchange rates.

    ```json
    {
      "$schema": "https://developer.microsoft.com/json-schemas/teams/v1.30/MicrosoftTeams.schema.json",
      "manifestVersion": "1.30",
      "id": "00000000-0000-0000-0000-000000000000",
      "version": "1.0.0",
      "name": {
        "short": "Purchase Planner",
        "full": "International Purchase Planner for Excel"
      },
      "description": {
        "short": "Compare vendor quotes in a common reporting currency.",
        "full": "An educational Excel custom skill that converts complete vendor quotes with current exchange rates and identifies quotes near or over budget."
      },
      "developer": {
        "name": "Contoso",
        "websiteUrl": "https://www.contoso.com",
        "privacyUrl": "https://www.contoso.com/privacy",
        "termsOfUseUrl": "https://www.contoso.com/servicesagreement"
      },
      "icons": {
        "outline": "assets/outline.png",
        "color": "assets/color.png"
      },
      "accentColor": "#217346",
      "localizationInfo": {
        "defaultLanguageTag": "en-us",
        "additionalLanguages": []
      },
      "authorization": {
        "permissions": {
          "resourceSpecific": [
            {
              "name": "Document.ReadWrite.User",
              "type": "Delegated"
            }
          ]
        }
      },
      "validDomains": [
        "localhost",
        "api.frankfurter.dev"
      ]
    }
    ```

1. Replace the string `1.30` in the `"$schema"` and `"manifestVersion"` properties at the top with the number of the latest version of the Microsoft 365 unified manifest schema.
1. Replace the placeholder `"id"` value with a randomly generated GUID.
1. Add the following `"extensions"` array to the root object. The first runtime provides the browser-based activation surface. The second runtime registers the custom functions, their namespace, and the URLs of their script and generated metadata.

    ```json
    "extensions": [
      {
        "requirements": {
          "scopes": [
            "workbook"
          ],
          "capabilities": [
            {
              "name": "CustomFunctionsRuntime",
              "minVersion": "1.1"
            }
          ]
        },
        "runtimes": [
          {
            "id": "TaskPaneRuntime",
            "type": "general",
            "code": {
              "page": "https://localhost:3000/taskpane.html"
            },
            "lifetime": "short",
            "actions": [
              {
                "id": "TaskPaneRuntimeShow",
                "type": "openPage",
                "pinnable": false,
                "view": "dashboard"
              }
            ]
          },
          {
            "id": "FunctionsRuntime",
            "type": "general",
            "code": {
              "page": "https://localhost:3000/functions.html",
              "script": "https://localhost:3000/public/functions.js"
            },
            "lifetime": "short",
            "customFunctions": {
              "namespace": {
                "id": "PURCHASEPLANNER",
                "name": "PURCHASEPLANNER"
              },
              "allowCustomDataForDataTypeAny": false,
              "metadataUrl": "https://localhost:3000/public/functions.json"
            }
          }
        ]
      }
    ]
    ```

1. Add the following `"agentSkills"` array to the root object.

    ```json
    "agentSkills": [
      {
        "folder": "./skills/international-purchase-planner"
      }
    ]
    ```

## Task 10: Add icons

Add the required package icons, **color.png** and **outline.png**, in the **appPackage/assets** folder. For details about the size requirements of the icons, see ["icons"](/microsoft-365/extensibility/schema/root-icons).

> [!TIP]
> To obtain the required files quickly, use Microsoft 365 Agents Toolkit to create any kind of App for Microsoft 365. The project that is created has the required **color.png** and **outline.png** files, usually in a folder named **assets**.

## Task 11: Build and package the skill

1. From the **international-purchase-planner** root, build the project.

    ```command line
    npm run build:dev
    ```

    The build generates the custom-function metadata and bundles the localhost files.

1. Validate the manifest.

    ```command line
    npm run validate
    ```

1. Start the localhost server.

    ```command line
    npm run dev-server
    ```

    If this is the first time you've run a development add-in on your computer, or the first time in a month, you may be prompted to delete expired certificates and install new ones. Respond **Yes** to both prompts.

1. In a separate command prompt, run the following command from the project root to create the package.

    ```command line
    atk package --manifest-file ./appPackage/manifest.json --output-package-file ./international-purchase-planner.zip --output-folder .
    ```

    The ZIP file should have the following package structure.

    ```text
    |-- manifest.json
    |-- assets/
    |   |-- color.png
    |   |-- outline.png
    |-- skills/
        |-- international-purchase-planner/
            |-- SKILL.md
            |-- resources/
            |   |-- workbook-data-guardrails.md
            |   |-- excel-vs-agent-execution.md
            |-- scripts/
                |-- prepare-purchase-comparison.js
    ```

    The custom-function JavaScript, metadata, and HTML files aren't included in the package. The add-in's server hosts them.

## Task 12: Test in Excel

1. Install the package by following the testing guidance in "Step 7: Test" of [Build plugins for Copilot Cowork](/microsoft-365/copilot/cowork/cowork-plugin-development#step-7-test). For example, run the following command.

    ```command line
    atk install --file-path ./international-purchase-planner.zip --scope Personal
    ```

1. Create or open a workbook that contains a table named **Settings** with the following data.

    > [!TIP]
    > To get a workbook that is already configured, download the **Text.xlsx** from the sample at **https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/CopilotExtensions/custom-functions-skill**.


    | Setting | Value |
    | --- | --- |
    | Reporting currency | USD |
    | Warning threshold | 0.1 |

    Don't enter the warning threshold as text such as "10%". Enter it as a number, such as `0.1`. Excel may then render the value as "10%".

1. Add a second table with at least the following columns and data.

    | Item | Vendor | Quote amount | Quote currency | Budget limit |
    | --- | --- | ---: | --- | ---: |
    | Laptops | Fabrikam | 18500 | EUR | 23000 |
    | Monitors | Contoso | 10200 | USD | 11000 |
    | Docking stations | Northwind | 14800 | CAD | 10000 |

    Every **Item** and **Vendor** value must be nonblank text. Every quote amount must be a nonnegative number, every quote currency must be a three-letter code, and every budget limit must be a positive number.

1. Verify that the custom functions are installed by entering `=PUR` in the formula bar. You should see **PURCHASEPLANNER.BUDGETSTATUS** and **PURCHASEPLANNER.FXRATE** in the autocomplete dropdown.

    > [!NOTE]
    > If the custom functions don't appear, select **Home** > **Add-ins**, and then select **Purchase Planner** to activate the add-in. If it doesn't appear, close and reopen Excel, wait two minutes, and try again.

1. Open Copilot in Excel.
1. Verify that your skill is installed with the following steps.
    1. Select the **+** icon in the chat area.
    1. Select **Choose plugins and skills**.
    1. Search for **Purchase Planner**, and make sure that it's enabled.
1. In the chat, ask the skill to compare the vendor quotes. Somewhere in your prompt, mention the full name of the skill in the form `@skill-name`. For example:

    ```text
    @international-purchase-planner, compare these vendor quotes.
    ```

    Or:

    ```text
    @international-purchase-planner, show me which quotes are near or over budget.
    ```

1. Approve the request to run the workbook script if Excel asks for confirmation.
1. Verify that the skill does the following.

    - Processes the first qualifying quote table in workbook order.
    - Adds or updates the **Exchange rate**, **Converted quote**, and **Budget status** columns.
    - Fills these columns in every quote row with formulas rather than static values.
    - Autofits each output column after the formulas finish calculating.
    - Applies green, yellow, and red conditional formatting to budget statuses.
    
    > [!NOTE]
    > The cells are typically populated in this order: first, `#NAME?` appears, then the custom function formulas appear, and finally the resolved values appear. The first time you use the skill, it may take several minutes for the workbook to recalculate and resolve the formulas to their final values and then autofit the columns. 

1. Verify that Copilot reports the processed table name, reporting currency, warning threshold, and number of processed rows. Copilot responses in the chat are nondeterministic, so the wording may differ between runs.
1. Remove a required quote value or replace a numeric source value with text.
1. Invoke the skill again. Verify that the skill doesn't modify the invalid table. If no other table qualifies, Copilot should report the exact error from the script.
1. Restore the quote table, and then in the **Settings** table, enter an invalid currency code or a nonnumber as the warning threshold.
1. Invoke the skill again. Verify that it stops without modifying the quote table and reports the settings error.
1. After each test session, uninstall the skill with the following steps.

    1. Close Excel.
    1. Open Teams and sign in with the same account that you used to install the skill.
    1. On the Teams app bar, select the **+** icon.
    1. In the upper-right corner of the **Store** page, select the gear icon.
    1. Find **Purchase Planner**.
    1. Expand the app entry, select the trash can icon, and confirm **Remove**.
    1. Stop the localhost server by pressing <kbd>Ctrl+C</kbd>.

## Troubleshooting

| Problem | Likely cause | Fix |
| --- | --- | --- |
| The skill doesn't appear in Copilot. | The package isn't installed for the account signed in to Excel, or Excel hasn't refreshed the package. | Confirm that Excel and the Agents Toolkit CLI use the same account, completely uninstall any older package, reinstall, and reopen Excel. |
| The custom functions don't appear in formula autocomplete. | The add-in isn't activated or the localhost server isn't running. | Confirm that the server is running, then select **Home** > **Add-ins** > **Purchase Planner**. |
| A function returns `#NAME?` and doesn't resolve to a formula or value even after several minutes. | Excel hasn't registered the custom-function metadata. | Confirm that the server provides `https://localhost:3000/public/functions.json`, and then close and reopen Excel. |
| A function remains at `#BUSY!` and doesn't resolve to a formula or value even after several minutes. | The function bundle or exchange-rate service can't be reached. | Confirm that the server provides `https://localhost:3000/public/functions.js` and that `api.frankfurter.dev` is reachable. |
| A formula is displayed as text. | The table column uses the **Text** number format. | Change the column format to **General** or **Number**, and then invoke the skill again. |
| Copilot says it can't run the script. | The skill is running outside Excel. | Open the workbook in Excel and invoke the skill there. |
