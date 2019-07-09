---
title: Use Office UI Fabric React in Office Add-ins
description: Learn how to use Office UI Fabric React in Office Add-ins.
ms.date: 07/10/2019
localization_priority: Priority
---

# Use Office UI Fabric React in Office Add-ins

Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.

To get started using Fabric React components in your add-in, perform the following steps.

> [!NOTE]
> If you follow the steps in this article, Fabric Core is also available in your add-in.

## Step 1 - Create your project with the Yeoman generator for Office

To create an add-in that uses Fabric React, we recommend that you use the Yeoman generator for Office. The Yeoman generator for Office provides the project scaffolding and build management needed to develop an Office Add-in.

To create your project, perform the following steps using **Windows PowerShell** (not the command prompt):

1. Install the prerequisites.
2. Run `yo office` to create the project files for your add-in.
3. When prompted to select an Office client application, choose **Word**.
4. Ensure you are in the directory with the project files, and then run `npm start`. A browser window showing a spinner opens automatically.
5. [Sideload your manifest](../testing/test-debug-office-add-ins.md) to view the full UI of the add-in.

## Step 2 - Add a Fabric React component

Next, add Fabric React components to your add-in. Complete the following steps to create a new React component, named `ButtonPrimaryExample`, that consists of a `Label` and `PrimaryButton` from Fabric React:

1. Open the project folder created by the Yeoman generator, and go to **src\taskpane\components**.
2. In that folder, create a new file named **Button.tsx**.
3. Open **Button.tsx** and add the following code to define the `ButtonPrimaryExample` component.

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';

export class ButtonPrimaryExample extends React.Component<IButtonProps, {}> {
  public constructor(props) {
    super(props);
  }

  insertText = async () => {
    // In the click event, write text to the document.
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph('Hello Office UI Fabric React!', Word.InsertLocation.end);
      await context.sync();
    });
  }

  public render() {
    let { disabled } = this.props;
    return (
      <div className='ms-BasicButtonsExample'>
        <Label>Click the button to insert text.</Label>
        <PrimaryButton
          data-automation-id='test'
          disabled={ disabled }
          text='Insert text...'
          onClick={ this.insertText } />
      </div>
    );
  }
}
```

This code does the following:

- References the React library using `import * as React from 'react';`.
- References the Fabric components (`PrimaryButton`, `IButtonProps`, `Label`) that are used to create `ButtonPrimaryExample`.
- Declares the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.
- Declares the `insertText` function that will handle the button's `onClick` event.
- Defines the HTML markup of the React component in the `render` function. The markup specifies that when the `onClick` event fires, the `insertText` function will run.

## Step 3 - Add the React component to your add-in

Add `ButtonPrimaryExample` to your add-in by opening **src\components\App.tsx** and completing the following steps:

1. Add the following import statement to reference `ButtonPrimaryExample` from **Button.tsx**.

  ```typescript
  import {ButtonPrimaryExample} from './Button';
  ```

2. Remove the following two import statements.

  ```typescript
  import { Button, ButtonType } from 'office-ui-fabric-react';
  ...
  import Progress from './Progress';
  ```

3. Replace the default `render()` function with the following code that uses `ButtonPrimaryExample`.

  ```typescript
  render() {
      return (
          <div className="ms-welcome">
          <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
          <HeroList message="Discover what this add-in can do for you today!" items={this.state.listItems} >
              <ButtonPrimaryExample />
          </HeroList>
          </div>
      );
  }
  ```

Save your changes. In Word, notice that the default text and button at the bottom of the task pane automatically updates to show the UI that's defined by the `ButtonPrimaryExample` React component.

![Screenshot of the Word application with the Insert text... button and preceeding text highlighted](../images/word-task-pane-with-react-component.png)

## See also

- [Office UI Fabric React](https://developer.microsoft.com/fabric)
- [UX design patterns for Office Add-ins](../design/ux-design-pattern-templates.md)
- [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [Office Add-in Fabric UI sample (uses Fabric 1.0)](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [Yeoman generator for Office](https://github.com/OfficeDev/generator-office)
