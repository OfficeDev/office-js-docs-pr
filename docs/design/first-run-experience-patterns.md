# First-run experience

A first-run experience is the experience a user has when they open your add-in for the first time. This experience shapes the user's impression of your add-in and can strongly influence their likelihood to continue use. Follow these best practices when crafting your first-run experience. 


## Best practices

|Do|Don't|
|:------|:------|
|Provide a simple and brief introduction to the main actions in  the add-in. | Don't include information and call-outs that are not relevant to getting started.
|Give users the opportunity to complete an action that will positively impact their use of the add-in.| Don't expect users to learn everything at once. Focus on the action that provides the most value.
|Create an engaging experience that users will want to complete. | Don't force the users to click through the first-run experience. Give users an option to bypass the first-run experience. |


Consider whether showing users the first-run experience once or many times is important to your scenario. For example, if users use your add-in periodically, they might forget how to use it, and it might be helpful to see the first-run experience more than once. 

Apply the following patterns as applicable to create or enhance the first-run experience for your add-in. 

## Carousel

The carousel takes users through a series of features or information pages before they start using the add-in.

> Formerly named **Paging Panel**.

Recommended flow for when using the carousel. 

![First Run - Carousel - Flowchart](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/carousel_flow.png)

1. Allow users to advance or skip the beginning pages of the carousel flow. 
![First Run - Carousel - Specifications for desktop task pane](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/carousel_taskPaneCallouts.png)

2. Provide a clear call to action to exit the first-run-experience.
![First Run - Carousel - Specifications for desktop task pane](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/carousel_taskPaneCallouts2.png)



### Code sample
* [Carousel code sample](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/carousel)

> [!NOTE]
> Some UX patterns do not match the source code. We're working hard to bring all assets into alignment.

<br/>

## Value Placemat

The value placement communicates your add-in's value proposition through logo placement, a clear value proposition, feature summary, and a call-to-action.


*Specification for desktop task pane*
![First Run - Value Placemat - Specifications for desktop task pane](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/valuePlacemat_taskPaneCallouts.png)


### Code sample
* [Value placemat code sample](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/value-placemat)

> [!NOTE]
> Some UX patterns do not match the source code. We're working hard to bring all assets into alignment.

<br/>

## Video Placemat

The video placemat shows users a video before they start using your add-in. 

Recommended screen flow when using the video placemat in your add-in. 
![Video Placemat - Flowchart](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/videoPlacemat_flow.png)

1. First Run Placemat - The screen contains a clear call to action button.
![Video Placemat - Specifications for desktop task pane](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/videoPlacemat_taskPaneCallouts.png)

2. Video Player - End users are presented with a video within a dialog window. 
![Video Placemat - Specifications for desktop task pane](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-UX-Design-Patterns/master/assets/images/videoPlacemat_taskPaneCallouts2.png)


### Code sample
* [Video placemat code sample](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat)

> [!NOTE]
> Some UX patterns do not match the source code. We're working hard to bring all assets into alignment. 

