---
title: Use Office UI Fabric React in Office Add-ins
description: 'Learn how to use Office UI Fabric React in Office Add-ins.'
ms.date: 01/16/2020
localization_priority: Normal
---

# Use Office UI Fabric React in Office Add-ins

Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.

This article describes how to create an add-in that's built with React and uses Fabric React components. 

> [!NOTE]
> [Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors) is included with Fabric React, which means your add-in will also have access to Fabric Core after you've completed the steps in this article.

## Create an add-in project

You'll use the Yeoman generator for Office Add-ins to create an add-in project that uses React.

### Install the prerequisites

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### Create the project

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type:** `Office Add-in Task Pane project using React framework`
- **Choose a script type:** `TypeScript`
- **What do you want to name your add-in?** `My Office Add-in`
- **Which Office client application would you like to support?** `Word`

![Yeoman generator](../images/yo-office-word-react.png)

After you complete the wizard, the generator creates the project and installs supporting Node components.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### Try it out

1. Navigate to the root folder of the project.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. Complete the following steps to start the local web server and sideload your add-in.

    > [!NOTE]
    > Office Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.

    > [!TIP]
    > If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - To test your add-in in Word, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens Word with your add-in loaded.

        ```command&nbsp;line
        npm start
        ```

    - To test your add-in in Word on a browser, run the following command in the root directory of your project. When you run this command, the local web server will start (if it's not already running).

        ```command&nbsp;line
        npm run start:web
        ```

        To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).

3. In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane. Notice the default text and the **Run** button at the bottom of the task pane. In the remainder of this walkthrough, you'll redefine this text and button by creating a React component that uses UX components from Fabric React.

    ![Screenshot of the Word application with the Show Taskpane ribbon button highlighted and the Run... button and preceeding text highlighted in the task pane](../images/word-task-pane-yo-default.png)


## Create a React component that uses Fabric React

At this point, you've created a very basic task pane add-in that's built using React. Next, complete the following steps to create a new React component (`ButtonPrimaryExample`) within the add-in project. The component uses the `Label` and `PrimaryButton` components from Fabric React.

1. Open the project folder created by the Yeoman generator, and go to **src\taskpane\components**.
2. In that folder, create a new file named **Button.tsx**.
3. In **Button.tsx**, add the following code to define the `ButtonPrimaryExample` component.

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
- Defines the UI of the React component in the `render` function. The HTML markup uses the `Label` and `PrimaryButton` components from Fabric React and specifies that when the `onClick` event fires, the `insertText` function will run.

## Add the React component to your add-in

Add the `ButtonPrimaryExample` component to your add-in by opening **src\components\App.tsx** and completing the following steps:

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

  4. Save the changes you've made to **App.tsx**.

## See the result

In Word, the add-in task pane automatically updates when you save changes to **App.tsx**. The default text and button at the bottom of the task pane now shows the UI that's defined by the `ButtonPrimaryExample` component. Choose the **Insert text...** button to insert text into the document.

![Screenshot of the Word application with the Insert text... button and preceeding text highlighted](../images/word-task-pane-with-react-component.png)

Congratulations, you've successfully created a task pane add-in using React and Office UI Fabric React! 

## See also

- [Office UI Fabric in Office Add-ins](office-ui-fabric.md)
- [Office UI Fabric React](https://developer.microsoft.com/fabric)
- [UX design patterns for Office Add-ins](ux-design-pattern-templates.md)
- [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
