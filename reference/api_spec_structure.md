# Structure of the Office JavaScript APIs

## Rules and conventions
 
* File name: lowercase, one word, inlcude only object name (not the full namespace); for example, binding.md
* Each object gets its own unique file. An object is anything that is not scalar (simple and complex types, full objects with properties and methods, non-string, int, and so on.)
* Don't inlcude information that repeats in each page. Provide links. 
* Don't include information that is not relevant or useful (support history, when introduced).
* Include all properties, relationships, and methods that part of an object in the same file. 
* Members of the object belong to one of the following categories: properties/relationships or methods. 
* Include examples for each method to show how the method is used. You can also provide broader examples at the object level. 
* Use simple code examples -- do not cover broader scenarios or use-case examples. If you have a larger example to add, do it as part of a separate content page and provide links. The idea is to keep the spec short and concise. 
* Enclose examples in a ```js {code} ``` block.
* Do not include HTML tags other than **br/** tags inside Markdown. HTML tags are treated differently and might cause issues with the HTML conversion. 

## Structure 

The following example shows the structure to follow. The `<html>` tags are used to show components/parts of the spec. Not all component are applicable for every object. For example, complex types won't have any methods. 

The variables are shown within `%percent%` symbols. 

Repeated rows are prefixed with `>r`.

Comments are provided inside <!-- {comment} --> blocks.

```md
<header>
# %name% resource type
%description%
</header>

<extendedremarks>
%remarks%
</extendedremarks> 

<requirements>
## Requirement set and supported hosts

| Requirement set | Application     |
|:---------------|:--------|
|%req%|%apps%|

</requirements>

<properties>
### Properties

| Property                | Type | Description| Req. Set Ver#| 
|:-------------|:-------|:-----------|:---|
>r|%name%      | %type% | %description% | %req% |

%propertygetset%
%propertynotes%
</properties>

<relationships>
### Relationships
| Relationship | Type         | Description| Requirement Set|
|:-------------|:-------|:-----------|:---|
>r|%name%      | [%type%](%link%) | %description% | %req% |

%relationshipnotes%
</relationships>

<methods>

## Methods

| Method                 | Return Type    | Description | Requirement Set|
|:-------------|:---------------|:------------|:----|
>r| [%name%](%link%)     | %dtype% | %description% | %req%|

%methodnotes%

## Method Details

<api>
<!-- Repeat <api/> block for each method--> 
### %apisignature%
%apidescription%
%syntax%
<parameter>
#### Parameters
%noparam%
| Method                 | Type    | Description | 
|:-------------|:---------------|:------------|
>r| %name%     | %dtype% | %description% | 

</parameter>
#### Returns
%returntype% 

<example>
#### Example
%examplelines%
</example>

</api>

</methods>
```
