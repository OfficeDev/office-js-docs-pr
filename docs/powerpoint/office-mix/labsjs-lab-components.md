---
title: LabsJS lab components
description: ''
ms.date: 12/04/2017
---


# LabsJS lab components

Labs.js provides you with four component types that you can use to assemble your lab. Each component type supports a specific type of lab interaction, including, for example, multiple choice problems, free response problems, or activities like viewing web pages in the lesson's HTML iFrame.

## Components

Office Mix supports the following four lab component types: 


-  **Activity component** ( **IActivityComponent**). Presents the user with an activity that must be completed; for example, read a piece of text, watch a video, or interact with a simulation. For more information, see [Labs.Components.ActivityComponentInstance](https://dev.office.com/reference/add-ins/office-mix/labs.components.activitycomponentinstance).
    
-  **Choice component** ( **IChoiceComponent**). Presents the user with a list of choices from which the user must select. Supports single or multiple responses (or no answer at all). Use this component type for true/false, multiple choice, multiple response, or polls. For more information, see [Labs.Components.ChoiceComponentInstance](https://dev.office.com/reference/add-ins/office-mix/labs.components.choicecomponentinstance).
    
-  **Input component** ( **IInputComponent**). Enables free form user input. Use this component type when you want to get responses to questions or math problems from the user, for example, or for other problem types that require text inputs from the user. For more information, see [Labs.Components.InputComponentInstance](https://dev.office.com/reference/add-ins/office-mix/labs.components.inputcomponentinstance).
    
-  **Dynamic component** ( **IDynamicComponent**). Generates other component types at runtime. Use this component type when you have branching questions, for example, where follow-up component types vary depending on a previous user input. This type also enables creating quiz banks or generating problems at runtime. For more information, see [Labs.Components.DynamicComponentInstance](https://dev.office.com/reference/add-ins/office-mix/labs.components.dynamiccomponentinstance).
    

## See also

- [Office Mix add-ins](office-mix-add-ins.md)
- [Configuring and editing LabsJS labs for Office Mix](configuring-and-editing-labsjs-labs-for-office-mix.md)
- [Walkthrough: Creating your first lab for Office Mix](creating-your-first-lab-for-office-mix.md)
    
