
# Labs.ValueHolder

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

A container object that holds and tracks values for a specified lab. The value may be stored either locally or on the server.

```
class ValueHolder<T>
```


## Variables


|||
|:-----|:-----|
| `public var isHint: boolean`|**True** if the value is a hint.|
| `public var hasBeenRequested: boolean`|**True** if the value has been requested by the lab.|
| `public var hasValue: boolean`|**True** if the value container currently has the desired value.|
| `public var value: T`|The value that is held in the container.|
| `public var id: string`|The ID of the value.|

## Methods




### getValue

 `public function getValue(callback: Labs.Core.ILabCallback<T>): void`

Retrieves the specified value.

 **Parameters**


|||
|:-----|:-----|
| _callback_|Callback function that returns the specified value.|

### provideValue

 `public function provideValue(value: T): void`

Internal method that provides the value to the value container.

 **Parameters**


|||
|:-----|:-----|
| _value_|The value to provide to the value container.|
