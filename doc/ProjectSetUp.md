# Project Set-Up/Testing

This section will go through: the basic prerequisites needed to develop for PowerPointLabs, how to set up the development environment and how to run tests within the development environment.

## Dev Prerequisites
1. Install Office 2016 or 2013 with PowerPoint
1. Install SourceTree (recommended), GitHub for Windows, or at least, Git
1. Install Visual Studio 2017
1. Install `VSTO for IDE` (Visual Studio Tools for Office, AKA the Office SDK):<br>
   for VS2015 - http://aka.ms/OfficeDevToolsForVS2015<br>
1. Install Microsoft .NET Framework 4.6

## Setup
1. [Fork](http://help.github.com/fork-a-repo/) and clone the source codes from this repo
1. Turn off *Office Version Upgrading*: Open the solution `PowerPointLabs\PowerPointLabs.sln` >> open Tools (menu) >> Options >> Office Tools >> Project Migration >> uncheck ‘Always upgrade to installed version of Office’
1. Set up *External Office Program*: Open ‘PowerPointLabs’ Properties >> Debug >> select ‘Start external program’ and choose `POWERPNT.exe` in the Office folder. Ensure that both **Debug** and **Release** configurations have set up this
1. Delete the .vs folder in ./PowerPointLabs/PowerPointLabs for Visual Studio to recreate the configuration files.
1. Run the solution by pressing F5 and then PowerPointLabs tab will appear in the PowerPoint ribbon. If you have installed PowerPointLabs add-in, you may have to uninstall it first and rebuild the solution
1. If failed to build PowerPointLabs solution, try to install `VSTO for PowerPoint` from [this link](http://powerpointlabs.info/vsto-redirect.html)

## Testing
1. Click Build (menu) >> Rebuild Solution
1. Click Test (menu) >> Windows >> Test Explorer
1. In the open Test Explorer window, click `Group by Traits` >> right click `FT` >> click `Run Selected Tests`. During the test, *DO NOT move the mouse & ensure the Windows UI is in English*
1. In the open Test Explorer window, click `Group by Traits` >> right click `UT` >> click `Run Selected Tests`
