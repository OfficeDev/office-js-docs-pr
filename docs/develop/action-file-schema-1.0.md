---
title: Action file schema 1.0 for Office
description: Learn about the 1.0 schema for an action file referenced by the manifest's actionFile property.
ms.localizationpriority: medium
ms.date: 10/19/2026
ms.topic: reference
---

# Action file schema 1.0 for Office

An action file describes the data format for a single Office.js action that's referenced by an Office Add-in's `actionFile` property. The action file provides metadata about the action, such as its ID and description, and it describes the action's parameters and its return value.

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
| `description` | String | Required. Describes what the action does and when to use it. Max chars: 1024. |
| `parameters` | [Parameters object](#parameters-object) | Optional. The parameters that can be passed to the action. |
| `returns` | [Returns object](#returns-object) | Optional. The value returned by the action. |

Unrecognized properties in the action file object aren't allowed.

### Action file object example

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/office-js/1.0/action.schema.json",
  "id": "addTodo",
  "description": "Adds a new to-do item to the current list.",
  "parameters": {
    "type": "object",
    "properties": {
      "description": {
        "type": "string",
        "description": "The text of the to-do item."
      }
    },
    "required": ["description"]
  },
  "returns": {
    "type": "string",
    "description": "The ID of the newly created to-do item."
  }
}
```

### Parameters object

Describes the parameters that can be passed to the action. This object mirrors the shape of an MCP tool's `inputSchema`, but supports only the subset of JSON Schema keywords described here.

The parameters object contains the following properties.

| Property | Type | Description |
| -------- | ---- | ----------- |
| `type` | String | Required. Must be `object`. |
| `properties` | Object | Required. Maps parameter names to [Parameter objects](#parameter-object). Each name must match the `^[A-Za-z0-9_]+$` regular expression. |
| `required` | Array of String | Optional. The names of the required parameters. Each name must match a key in `properties`. Parameters not listed here are optional. |
| `additionalProperties` | Boolean | Optional. Whether parameters not listed in `properties` are allowed. |

Unrecognized properties in the parameters object aren't allowed. The `additionalProperties` property controls extra parameters in the action's input, not extra fields in the parameters object itself.

#### Parameters object example

```json
{
  "type": "object",
  "properties": {
    "description": {
      "type": "string",
      "description": "The text of the to-do item."
    }
  },
  "required": ["description"],
  "additionalProperties": false
}
```

### Parameter object

Describes a single action parameter.

The parameter object contains the following properties.

| Property | Type | Description |
| -------- | ---- | ----------- |
| `type` | String | Required. Specifies the parameter's type. Possible values are: `string`, `number`, `integer`, `boolean`. |
| `description` | String | Optional. The description of the parameter. Max chars: 512. |
| `enum` | Array of String | Optional. The valid values for this parameter. Only valid when `type` is `string`. |
| `default` | String, Number, Integer, Boolean | Optional. The default value for an optional parameter. Must match the type specified by `type`. |

Unrecognized properties in the parameter object aren't allowed. The parameter's name is its key in `parameters.properties`, and whether it's required is specified in `parameters.required`.

#### Parameter object example

```json
{
  "type": "string",
  "description": "The text of the to-do item."
}
```

### Returns object

Describes the value returned by the action.

The returns object contains the following properties.

| Property | Type | Description |
| -------- | ---- | ----------- |
| `type` | String | Required. The type of the returned value. Possible values are: `string`. |
| `description` | String | Optional. The description of the returned value. Max chars: 512. |

Unrecognized properties in the returns object aren't allowed.

#### Returns object example

```json
{
  "type": "string",
  "description": "The ID of the newly created to-do item."
}
```

## Action file example

The following example shows an action file that uses all of the properties described in this article.

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/office-js/1.0/action.schema.json",
  "id": "addTodo",
  "description": "Adds a new to-do item to the current list.",
  "parameters": {
    "type": "object",
    "properties": {
      "description": {
        "type": "string",
        "description": "The text of the to-do item."
      },
      "dueDate": {
        "type": "string",
        "description": "The due date of the to-do item, in ISO 8601 format."
      },
      "priority": {
        "type": "string",
        "description": "The priority of the to-do item.",
        "enum": ["low", "normal", "high"],
        "default": "normal"
      }
    },
    "required": ["description"],
    "additionalProperties": false
  },
  "returns": {
    "type": "string",
    "description": "The ID of the newly created to-do item."
  }
}
```

## Related content

- QUESTION7
