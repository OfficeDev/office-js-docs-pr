---
title: Create Copilot skills for Excel (preview)
description: Learn how to create agent skills that call the Office JavaScript Library (Office.js) that can be plugged into Copilot agents in Excel.
ms.date: 06/25/2026
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Create Copilot skills for Excel (preview)

You can create custom skills to accomplish complex tasks for AI agents and Copilot for Excel. Optionally, these skills can use the Office JavaScript Library (Office.js).

> [!NOTE]
> Custom skills for Excel are in preview. Do not use them in a production skill. 

The main guidance for creating skills is at [Build plugins for Copilot Cowork](/microsoft-365/copilot/cowork/cowork-plugin-development). The context of the article refers to Microsoft Cowork, but the guidance applies generally to all custom skills that can be plugged into an agent running in Copilot.

Start with the section [What you'll build](/microsoft-365/copilot/cowork/cowork-plugin-development#what-youll-build) and then skip to [Build a plugin from scratch](/microsoft-365/copilot/cowork/cowork-plugin-development#build-a-plugin-from-scratch). 

*This article provides only supplementary guidance for creating skills that use Office.js to interact with an Excel workbook.*

## Supplement to Step 1

1. In the YAML frontmatter of a SKILL.md file, include a `metadata.tags` property with an `excel` tag, and other tags as needed. The following is an example.

   ```yaml
   name: my-excel-skill
   description: |
     Use this skill in Excel to ... 
   metadata:
     version: 1.0.0
     tags: excel
   ```

1. For most Excel skills, the output should be in the workbook, not a chat or CLI interface. So, in place of the "Output Format" section of a SKILL.md file, have a "Workbook Output" section. The following is an example.

   ```md
   ## Workbook output

   Produce a dashboard sheet with the following characteristics:

   - It should be named "My Skill Output".
   - ... other characteristics here.
   ```

1. Include a "Common pitfalls to avoid" section in the SKILL.md. The following is an example.

   ```md
   ## Common pitfalls to avoid

   - Do not search broadly for unrelated sheets or columns.
   - Do not invent missing data.
   - Do not claim to be finished when the current environment can't update the workbook.
   ```

## Supplement to Step 2

1. Create JavaScript scripts that call APIs in Office.js and put them in the **\scripts** folder of the skill. Each script must run to completion without further interaction from the user or the agent. Any Office.js APIs can be called, but typically, the script will consist of a single asynchronous call of `Excel.run`, as in the following edited example.

   ```javascript
   await Excel.run(async (context) => {
      ...
      await context.sync();
   });
   ```

   > [!IMPORTANT]
   > The script should *not* call `Office.onReady` or define an `Office.initialize`. Copilot in Excel automatically creates a runtime and initializes Office.js. 

1. Include as resources, two markdown files that will impose some boundaries on how the skill interacts with the workbook. 

   1. The first resource should give instructions on how to the existing workbook should be cleaned up or normalized before the skill adds any output to it. These instructions are specific to the scenario the skill addresses, not generic Excel best practices. The following is an example from a skill that creates a dashboard with data about an athletic rivalry between two American colleges.

      ```md
      # Workbook data guardrails

      Use this reference to improve data quality before the scenario skill generates an output.

      ## Core principles

      - Start from workbook labels already present in the workbook.
      - Search only for the exact missing sheet or column names the scenario needs.
      - Ask precise questions instead of inventing scores, quotes, colors, or team metadata.

      ## Search rules

      - Prefer exact names such as `Game Results`, `Player Stats`, `Team Stats`, `Quotes Colors`, `Team Profiles`, `Roster Recruiting`, or `Opponent PPG`.
      - If a required field is missing, ask for the exact replacement label instead of broad keyword guesses.
      - Stop searching once the needed data for this scenario is located.

      ## Organization playbook

      - If the workbook is messy, ask the user to rename sheets or headers before analysis.
      - Use one sheet for game results, one for player or team stats, and one for branding or quotes when possible.
      - Don't assume one table can serve multiple scenarios.

      ## Quality checks

      - Confirm the selected school, season range, metric choice, and other custom hooks before generating output.
      - Keep caveats visible for limited games, freshman samples, or incomplete rows.
      - Label synthetic rows clearly when they are used.

      ```
   
   1. The second resource should do the following.

      - Provide some general quality rules that Copilot should always following when using the skill.
      - Explain when the skill's scripts should be called and when they shouldn't be.
      - Explain what to do, and not do, if a user has invoked the skill in Copilot outside of Excel.

      The following is an example.

      ```md
      # Excel vs non-Excel execution guidance

      ## When the skill runs inside Excel

      - Prefer workbook-native ranges, named tables, and Office.js helpers already available in the current session.
      - Keep the output focused on workbook changes, sheet creation, chart updates, and branded formatting.
      - Use the workbook as the primary source of truth.

      ## When the skill runs outside Excel

      - Treat the workbook as a data file and provide the exact next steps, schema assumptions, and missing inputs.
      - Avoid claiming that charts or sheets were completed unless the current environment can perform the action.
      - Ask the user to open the workbook in Excel or provide the exact labels needed for the scenario.

      ## Quality checks for either environment

      - Confirm the execution mode before generating output.
      - State clearly what was completed and what still needs user or workbook action.
      - Never invent missing data or pretend a workbook update succeeded.

      ## When to use the scripts folder

      - Run `scripts/my-first-script.js` when, and only when, the user asks for ... 
      - Run `scripts/another-script.js` when, and only when, the user asks for ...
      - Use these scripts after the workbook context is ready for shape and sheet generation, not during an analysis-only pass.
      - If the current environment is outside Excel, explain the workbook steps and do not attempt to execute the Office.js scripts.
      ```

## Supplement to Step 4

1. Although your skill calls APIs from Office.js, you don't need an `"extensions"` section in the manifest. Copilot in Excel automatically creates a runtime, loads Office.js and executes the scripts in your skill. 

   > [!NOTE]
   > Because the runtime isn't explicitly configured in an "extensions" element, there is no `"extensions.requirements.capabilities"` key in the manifest, so there is no way to limit the runtime to a version of Excel that supports a specific Requirement Set. It is unlikely, but theoretically possible, that a user's channel may have a version of Excel that doesn't support an API that is called in a script in the skill. The effect of such a call is undefined. 

1. Because a skill plugin is an App for Microsoft 365 and is implemented with the unified manifest and package for Microsoft 365, you can include your custom skill in the same manifest and package as any other type of app for Microsoft 365 such as a Teams Tab or an Office Add-in. However, you can't share your Office.js code between the skill and one of these other apps because the skills scripts are in the app package, not hosted online. You must duplicate code in two or more files if it needs to be run by both the skill and the other app.

## Supplement to Step 7

It is important to cleanly uninstall the skill after each test session. Follow these steps.

1. Open Teams and be sure you're signed in with the same credentials you used to install the skill. 
1. Select the apps button: the plus sign in a box.
1. On the **Apps** pane, select **Manage your apps**.
1. Find your add-in in the list of apps. It has the name specified in the **name** property of the YAML frontmatter in the SKILL.md file.
1. Select the add-in from the list of apps to expand its row.
1. Select the trash can icon and then select **Remove** in the prompt.

