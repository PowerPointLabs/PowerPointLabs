##Software Design

(graph here)  
PowerPointLabs is an add-in for PowerPoint. Given above is an overview of the main components.

* **Add-in Ribbon**: The UI seen by users in the PowerPoint Ribbon tabs or context menu. It consists in [`ThisAddIn.cs`](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/ThisAddIn.cs), [`Ribbon1.cs`](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/Ribbon1.cs), and [`Ribbon1.xml`](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/Ribbon1.xml). `ThisAddIn.cs` is in charge of add-in's lifecycle and other important events. `Ribbon1.xml` defines the styles of add-in ribbon and context menu, and `Ribbon1.cs` routes the user requests to the UI and Logic.
* **Test Driver**: PowerPointLabs relies on the test automation to prevent regression. [`Functional Test`](https://github.com/PowerPointLabs/PowerPointLabs/tree/master/PowerPointLabs/FunctionalTest) and MSTest is used to automate testing against UI and Logic. [`Test data`](https://github.com/PowerPointLabs/PowerPointLabs/tree/master/doc/test) is accessed during testing.
* **UI**: The UI seen by users as a sidebar (AKA task pane) or window. [`WPF`](https://msdn.microsoft.com/en-us/library/mt149842(v=vs.110).aspx) and <span style="color:gray">`Winform (deprecated)`</span> techniques are used to build the UI, and [`MVVM pattern`](https://msdn.microsoft.com/en-us/library/hh848246.aspx) is preferred to implement UI-related features. Not all features require a UI.
* **Logic**: The main part of the add-in that implements features logic.
* **Storage**: PowerPointLabs generally uses `Temp` folder to store temporary data and `Documents` folder to save user data and settings.
* **Model**: This includes the [`PowerPoint Object Model`](https://msdn.microsoft.com/en-us/library/microsoft.office.interop.powerpoint(v=office.14).aspx) and some wrapper classes of the `PowerPoint Object Model`.
* Many [`Windows APIs`](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/PowerPointLabs/NativeMethods.cs) are used as a supplement to Office APIs.

## Conventions

* Ensure the codes are [`SOLID`](http://www.codeproject.com/Articles/703634/SOLID-architecture-principles-using-simple-Csharp) and [testable](http://www.toptal.com/qa/how-to-write-testable-code-and-why-it-matters).
