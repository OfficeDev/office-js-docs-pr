
# Labs.Components.IChoiceComponent

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Enables interactions with a choice component.

```
interface IChoiceComponent extends Labs.Core.IComponent
```


## Properties


|Name|Description|
|:-----|:-----|
| `choices: Components.IChoice[]`|An array representing the list of choices associated with the problem.|
| `timeLimit: number`|Time limit for completing the problem.|
| `maxAttempts: number`|Maximum number of attempts allowed for the problem.|
| `maxScore: number`|The maximum score for the problem.|
| `hasAnswer: boolean`|**True** if the problem has an answer.|
| `answer: any`|The answer for the problem. Either an array if multiple answers are supported or a single ID if only one answer is supported.|
| `secure: boolean`|Whether or not the quiz is secure, meaning that secure fields are withheld from the user.|
