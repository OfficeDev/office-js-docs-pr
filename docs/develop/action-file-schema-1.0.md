---
title: Action file schema 1.0 for Office
description: Learn about the 1.0 schema for an action file referenced by the manifest's actionFile property.
ms.localizationpriority: medium
ms.date: 10/19/2026
ms.topic: reference
---

# Action file schema 1.0 for Office

An action file describes the data format for a single Office.js action that's referenced by an Office Add-in's `actionFile` property. The action file provides metadata about the action, such as its ID and description, and it describes the action's parameters, its return value, and per-state instructions for the Copilot orchestrator.

The following article describes the 1.0 schema used by action files. For more information about Office.js actions and how they're referenced from an add-in, see QUESTION4.

## JSON schema

The schema described in this document is available in [JSON Schema format](https://developer.microsoft.com/json-schemas/office-js/1.0/action.schema.json).

## Action file object

The root of the action file document is a JSON object that contains properties that describe the action.

The action file object contains the following properties.

| Property | Type | Description |
| -------- | ---- | ----------- |
| `$schema` | String | Optional. The URI/URL of this schema. In a production app this should always be set to `https://developer.microsoft.com/json-schemas/office-js/1.0/action.schema.json`. |
| `id` | String | Required. The ID of the action, which must match an action ID in the manifest. Max chars: 64. |
| `actionDescription` | String | Required. Describes what the action does and when to use it. Max chars: 1024. |
| `parameters` | Array of [Parameter object](#parameter-object) | Optional. The parameters that can be passed to the action. |
| `returnedValue` | [Returned value object](#returned-value-object) | Optional. The value returned by the action. |
| `states` | [States object](#states-object) | Optional. Per-state instructions for the Copilot orchestrator. |

Unrecognized properties in the action file object aren't allowed.

### Action file object example

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/office-js/1.0/action.schema.json",
  "id": "addTodo",
  "actionDescription": "Adds a new to-do item to the current list.",
  "parameters": [
    {
      "name": "description",
      "parameterDataType": "string",
      "parameterDescription": "The text of the to-do item.",
      "isRequired": true
    }
  ],
  "returnedValue": {
    "dataType": "string",
    "returnedValueDescription": "The ID of the newly created to-do item."
  }
}
```

### Parameter object

Describes a single action parameter.

The parameter object contains the following properties.

| Property | Type | Description |
| -------- | ---- | ----------- |
| `name` | String | Required. The name of the parameter. Value must match the `^[A-Za-z0-9_]+$` regular expression. Max chars: 64. |
| `parameterDataType` | String | Required. Specifies the parameter's type. Possible values are: `string`, `boolean`, `number`. |
| `parameterDescription` | String | Optional. The description of the parameter. Max chars: 512. |
| `isRequired` | Boolean | Optional. Whether the parameter is required. Defaults to `false`. |
| `defaultValue` | String, Boolean, Number | Optional. The default value for an optional parameter. |

Unrecognized properties in the parameter object aren't allowed.

#### Parameter object example

```json
{
  "name": "description",
  "parameterDataType": "string",
  "parameterDescription": "The text of the to-do item.",
  "isRequired": true
}
```

### Returned value object

Describes the value returned by the action.

The returned value object contains the following properties.

| Property | Type | Description |
| -------- | ---- | ----------- |
| `dataType` | String | Required. The type of the returned value. Possible values are: `string`. |
| `returnedValueDescription` | String | Optional. The description of the returned value. Max chars: 512. |

Unrecognized properties in the returned value object aren't allowed.

#### Returned value object example

```json
{
  "dataType": "string",
  "returnedValueDescription": "The ID of the newly created to-do item."
}
```

### States object

Contains per-state instructions for the Copilot orchestrator.

The states object contains the following properties. At least one property MUST be present.

| Property | Type | Description |
| -------- | ---- | ----------- |
| `reasoning` | [State configuration object](#state-configuration-object) | Optional. The state where the model can call actions. |
| `responding` | [State configuration object](#state-configuration-object) | Optional. The state where the model responds to the user. |

Unrecognized properties in the states object aren't allowed.

#### States object example

```json
{
  "states": {
    "reasoning": {
      "stateDescription": "Use this action to add a to-do item when the user asks to track a new task."
    },
    "responding": {
      "stateDescription": "Confirm to the user that the to-do item was added."
    }
  }
}
```

### State configuration object

Contains instructions for the action in a specific orchestrator state.

The state configuration object contains the following property.

| Property | Type | Description |
| -------- | ---- | ----------- |
| `stateDescription` | String | Required. The purpose of the action in this state. Max chars: 512. |

Unrecognized properties in the state configuration object aren't allowed.

#### State configuration object example

```json
{
  "stateDescription": "Use this action to add a to-do item when the user asks to track a new task."
}
```

## Action file example

The following example shows an action file that uses all of the properties described in this article.

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/office-js/1.0/action.schema.json",
  "id": "addTodo",
  "actionDescription": "Adds a new to-do item to the current list.",
  "parameters": [
    {
      "name": "description",
      "parameterDataType": "string",
      "parameterDescription": "The text of the to-do item.",
      "isRequired": true
    },
    {
      "name": "dueDate",
      "parameterDataType": "string",
      "parameterDescription": "The due date of the to-do item, in ISO 8601 format.",
      "isRequired": false
    }
  ],
  "returnedValue": {
    "dataType": "string",
    "returnedValueDescription": "The ID of the newly created to-do item."
  },
  "states": {
    "reasoning": {
      "stateDescription": "Use this action to add a to-do item when the user asks to track a new task."
    },
    "responding": {
      "stateDescription": "Confirm to the user that the to-do item was added."
    }
  }
}
```

## Related content

- QUESTION7

