
# Labs.Core.Actions.ISubmitAnswerResult

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

The result of submitting an answer for an attempt.

```
interface ISubmitAnswerResult extends Core.IActionResult
```


## Properties


|||
|:-----|:-----|
| `submissionId: string`|An ID associated with the submission. Provided by the server.|
| `complete: boolean`|Returns  **true** if the attempt is completed due to the current submission.|
| `score: any`|Score information associated with the submission.|
