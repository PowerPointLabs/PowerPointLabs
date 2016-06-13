##Software Design

![Architecture](https://raw.githubusercontent.com/PowerPointLabs/PowerPointLabs/master/doc/DesignAndConventions.png)  
PowerPointLabs is an add-in for PowerPoint. Given above is an overview of the main components.

* **Add-in Ribbon**: The UI seen by users in the PowerPoint Ribbon tabs or context menu. It consists in [`ThisAddIn.cs`](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/ThisAddIn.cs), [`Ribbon1.cs`](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/Ribbon1.cs), and [`Ribbon1.xml`](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/Ribbon1.xml). `ThisAddIn.cs` is in charge of add-in's lifecycle and other important events. `Ribbon1.xml` defines the styles of add-in ribbon and context menu, and `Ribbon1.cs` is the entry point that routes the user requests to the UI and Logic through the Action Framework. Any changes made in `Ribbon1.cs` or `ThisAddIn.cs` should be generic enough to be used by every feature.
* **UI**: The UI seen by users as a sidebar (AKA task pane) or window. [`WPF`](https://msdn.microsoft.com/en-us/library/mt149842(v=vs.110).aspx) and <span style="color:gray">`Winform (deprecated)`</span> techniques are used to build the UI, and [`MVVM pattern`](https://msdn.microsoft.com/en-us/library/hh848246.aspx) is preferred to implement UI-related features. Not all features require a UI.
* **Logic**: The main part of the add-in that implements features logic.
* **Test Driver**: PowerPointLabs relies on the test automation to prevent regression. [`Functional Test`](https://github.com/PowerPointLabs/PowerPointLabs/tree/master/PowerPointLabs/Test) and [`Unit Test`](https://github.com/PowerPointLabs/PowerPointLabs/tree/master/PowerPointLabs/Test) is used to automate testing against UI and Logic. [`Test data`](https://github.com/PowerPointLabs/PowerPointLabs/tree/master/doc/test) is accessed during testing. 
* **Storage**: PowerPointLabs generally uses `Temp` folder to store temporary data and `Documents` folder to save user data and settings.
* **Model**: This includes the [`PowerPoint Object Model`](https://msdn.microsoft.com/en-us/library/microsoft.office.interop.powerpoint(v=office.14).aspx) and some wrapper classes of the `PowerPoint Object Model`.
* Many [`Windows APIs`](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/NativeMethods.cs) are used as a supplement to Office APIs.

##Add-in Ribbon & UI

The diagram below shows the structure of Ribbon & UI with [Action Framework](https://github.com/PowerPointLabs/PowerPointLabs/tree/master/PowerPointLabs/PowerPointLabs/ActionFramework). 

![ActionFramework](https://raw.githubusercontent.com/PowerPointLabs/PowerPointLabs/master/doc/ActionFramework.png)

When a request (e.g. click a button) comes to the `Ribbon`, `HandlerFactory` will create a `Handler` to handle the request. In the `Handler`, it can use `ActionFrameworkExtensions` to access the current context (e.g. current selected shape, current slide), use some `Feature Logic` (e.g. fit to width) to handle the request, or display some `Feature UI` (e.g. a sidebar).

###To create a new feature

- Set up [ribbon.xml](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/Ribbon1.xml#L394). Provide a unique id for the ribbon control.
```xml
<button id="fitToWidthShape"
        getLabel="GetLabel"
        getImage="GetImage"
        onAction="OnAction"/>
```
- Set up handlers. In this example, we will need handlers for [Label](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/ActionFramework/Label/FitToWidthLabelHandler.cs), [Image](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/ActionFramework/Image/FitToWidthImageHandler.cs), and [Action](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/ActionFramework/Action/FitToWidthActionHandler.cs).
```cs
// Provide the ribbon control id to link this handler to the ribbon control
[ExportActionRibbonId("fitToWidthShape")]
class FitToWidthActionHandler : ActionHandler
```
- Access PowerPoint context if required.
To access the PowerPoint context, type `this.` in a ActionHandler or WPF UI control, and then you'll be able to access a list of context getters provided by [ActionFrameworkExtensions](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/ActionFramework/Common/Extension/ActionFrameworkExtensions.cs).
- Set up WPF UI if required. To set up a sidebar UI, you'll need to wrap the WPF UI in a Winform UI [[example]](https://github.com/PowerPointLabs/PowerPointLabs/tree/master/PowerPointLabs/PowerPointLabs/DrawingsLab), and then call [`this.RegisterTaskPane`](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/ActionFramework/Common/Extension/ActionFrameworkExtensions.cs#L77) to register the sidebar. For the UI style, we're using Metro style UI provided by [Mahapps](http://mahapps.com).
- Call the feature logic from the ActionHandler or UI to complete the request.

##Logic & Testing

The diagram below shows the structure of backend. 

![Backend](https://raw.githubusercontent.com/PowerPointLabs/PowerPointLabs/master/doc/Backend.png)

UI and ActionHandler can call feature logic to process the request. Test component (unit test and functional test) can call feature logic to do test-automation. Feature logic is built upon `PowerPoint Object Model`, [`other model classes`](https://github.com/PowerPointLabs/PowerPointLabs/tree/master/PowerPointLabs/PowerPointLabs/Models), and some other components. 

###Notes

- The feature logic should be [`SOLID`](http://www.codeproject.com/Articles/703634/SOLID-architecture-principles-using-simple-Csharp) and [`testable`](http://www.toptal.com/qa/how-to-write-testable-code-and-why-it-matters), and be organized into its own package/folder.
- For testable logic, it can be tested by `Unit Test (UT)`. For untestable/legacy/UI logic, it can be tested by `Functional Test (FT)`. Instructions of testing can be found [here](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/Test/README.md).
- It's highly recommended to use [Logger](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/ActionFramework/Common/Log/Logger.cs) to capture significant events in features.
- It's highly recommended to model slide-level or presentation-level behaviours by extending [`PowerPointSlide.cs`](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/Models/PowerPointSlide.cs) and [`PowerPointPresentation.cs`](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/Models/PowerPointPresentation.cs).

## Conventions

* Ensure the codes are [`SOLID`](http://www.codeproject.com/Articles/703634/SOLID-architecture-principles-using-simple-Csharp) and [`testable`](http://www.toptal.com/qa/how-to-write-testable-code-and-why-it-matters).
