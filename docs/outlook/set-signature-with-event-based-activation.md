---
title: Set the email signature using event-based activation
description: Set the email signature using event-based activation
ms.topic: article
ms.date: 07/12/2021
localization_priority: Normal
---

# Set the email signature using event-based activation

This sample uses event-based activation to run an Outlook add-in when the user creates a new message or appointment. The add-in can respond to events, even when the task pane is not open. It also uses the setSignatureAsync API. If no signature is set, the add-in prompts the user to set a signature, and can then open the task pane for the user.
scenario

image

Scenario: Event-based activation
In this scenario, the add-in helps the user manage their email signature, even when the task pane is not open. When the user sends a new email message, or creates a new appointment, the add-in displays an information bar prompting the user to create a signature. If the user chooses to set a signature, the add-in opens the task pane for the user to continue setting their signature.

Prereqs

## Create a new Outlook Office Add-in

1. Use yo office to create a new Office Add-in

## Configure event-based activation in the manifest

1. Add the following XML to the manifest.

```xml
<Runtimes>
  <Runtime resid="Autorun">
    <Override type="javascript" resid="runtimeJs"/>
  </Runtime>
</Runtimes>
```

## Create the loading page for Outlook on the web

1. Add the following html to a new file named autorunweb.html

```html
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title></title>
    <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.9.1.min.js" type="text/javascript"></script>
    <script src="https://appsforoffice.microsoft.com/lib/beta/hosted/office.js" type="text/javascript"></script>
    <script type="text/javascript" src="../Js/autorunshared.js"></script>
    <script type="text/javascript" src="../Js/autorunweb.js"></script>
</head>
<body>
    <!-- NOTE: The body is empty on purpose. Since this is invoked via a button, there is no UI to render. -->
</body>
</html>
```

## Create shared code file

1. Add the following code to autorunshared.js

```javascript
/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
 */
function checkSignature(eventObj) {
  let user_info_str = Office.context.roamingSettings.get("user_info");
  if (!user_info_str) {
    display_insight_infobar();
  } else {
    let user_info = JSON.parse(user_info_str);

    if (Office.context.mailbox.item.getComposeTypeAsync) {
      //Find out if the compose type is "newEmail", "reply", or "forward" so that we can apply the correct template.
      Office.context.mailbox.item.getComposeTypeAsync(
        {
          asyncContext: {
            user_info: user_info,
            eventObj: eventObj,
          },
        },
        function (asyncResult) {
          if (asyncResult.status === "succeeded") {
            insert_auto_signature(
              asyncResult.value.composeType,
              asyncResult.asyncContext.user_info,
              asyncResult.asyncContext.eventObj
            );
          }
        }
      );
    } else {
      // Appointment item. Just use newMail pattern
      let user_info = JSON.parse(user_info_str);
      insert_auto_signature("newMail", user_info, eventObj);
    }
  }
}
```

## Try it out

1. Run the code and try it out.
