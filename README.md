<img src="https://raw.githubusercontent.com/PowerPointLabs/PowerPointLabs-Installer/master/PowerPointLabsInstaller/PowerPointLabsInstallerUi/Resources/logo.png" width='300'>

The typical PowerPoint presentation isn't very interesting. Walls of text or bullet points, with few visuals - it's no wonder audiences find it hard to pay attention.
Your slides don't have to be this way, and it doesn't take a whole lot of effort to make them better.
PowerPointLabs makes creating engaging PowerPoint presentations easy. Check out what it can do for you here: http://powerpointlabs.info

[![Build status](https://img.shields.io/appveyor/ci/kai33/powerpointlabs/master.svg)](https://ci.appveyor.com/project/kai33/powerpointlabs)

### Contributing To PowerPointLabs
Interested to contribute? Please take a moment to review the [guidelines for contributing](https://github.com/PowerPointLabs/powerpointlabs/blob/master/doc/CONTRIBUTING.md) and [the design](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/doc/SoftwareDesign.md).

### Dev Prerequisites
0. Install Office 2016 or 2013 with PowerPoint.
1. Install SourceTree (recommended), GitHub for Windows, or at least, Git.
2. Install Visual Studio 2015.
3. Install `VSTO for IDE` (Visual Studio Tools for Office, AKA the Office SDK):<br>
   for VS2015 - http://aka.ms/OfficeDevToolsForVS2015<br>
4. Install Microsoft .NET Framework 4.6.

### Setup
0. [Fork](http://help.github.com/fork-a-repo/) and clone the source codes from this repo.
1. Turn off *Office Version Upgrading*: Open the solution `PowerPointLabs\PowerPointLabs.sln` >> open Tools (menu) >> Options >> Office Tools >> Project Migration >> uncheck ‘Always upgrade to installed version of Office’.
2. Set up *External Office Program*: Open ‘PowerPointLabs’ Properties >> Debug >> select ‘Start external program’ and choose `POWERPNT.exe` in the Office folder. Ensure that both **Debug** and **Release** configurations have set up this.
3. Run the solution by pressing F5 and then PowerPointLabs tab will appear in the PowerPoint ribbon. If you have installed PowerPointLabs add-in, you may have to uninstall it first and rebuild the solution.
4. If failed to build PowerPointLabs solution, try to install `VSTO for PowerPoint` from the link inside [this file](https://github.com/PowerPointLabs/PowerPointLabs-Website/blob/master/vsto-redirect.html).

### Testing
0. Click Build (menu) >> Rebuild Solution.
1. Click Test (menu) >> Windows >> Test Explorer. 
2. In the open Test Explorer window, click `Group by Traits` >> right click `FT` >> click `Run Selected Tests`. During the test, *DO NOT move the mouse & ensure the Windows UI is in English*.
3. In the open Test Explorer window, click `Group by Traits` >> right click `UT` >> click `Run Selected Tests`.

### Acknowledgements
PowerPointLabs is developed at the School of Computing, National University of Singapore, with funding from an NUS Learning Innovation Fund, Technology (LIFT) grant.

### License
PowerPointLabs is released under GPLv2.
