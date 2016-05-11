# Structure of the Office JavaScript APIs

## Rules and conventions
 
* File namename: lowecase, one word, inlcude only object name (not the full namespace). e.g., binding.md
* Each "object" gets it's own unique file. Object is anything that is not scalar (simple and complex types, full objects with properties, methods, etc. non string, int, so on..)
* Don't inlcude information that repeats in each page. Provide links. 
* Don't include information that is not relevant or useful (e.g., support history, when introduced, etc.)
* All properties, relationships and methods part of an object are in the same file. 
* Members of the object belong to one of the following categories: properties/relationships or methods. 
* Include examples for each method to showcase the usage. Additional broader examples could be provided at the object level. 
* Think of example code as simple code-snippers -- not broader scenario or use-case examples. If you have a larger example to add, do it as part of a separate content page and provide links. The idea is to keep the spec short and sweet. 
* Examples should be enclosed in ```js {code} ``` block.
* Do not include any HTML tags other thab <br/> inside markdown. HTML tags are treated differently and may cause issues with HTML conversion of markdown page. 

## Structure 

Below block provides the structure to follow. The `<html>` looking tags are used to show components/parts of the spec. Not all component may be applicable for an object file. Example: complex types won't have any methods. 

The variables are shown within `%percent%` symbol. 

Repeated rows are prefixed with `>r`.

Comments are provided inside <!-- {comment} --> block.

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
