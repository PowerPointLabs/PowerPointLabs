<img src="https://raw.githubusercontent.com/PowerPointLabs/powerpointlabs/master/PowerPointLabsInstallerUi/PowerPointLabsInstallerUi/Resources/logo.png" width='300'>

The typical PowerPoint presentation isn't very interesting. Walls of text or bullet points, with few visuals - it's no wonder audiences find it hard to pay attention.
Your slides don't have to be this way, and it doesn't take a whole lot of effort to make them better.
PowerPointLabs makes creating engaging PowerPoint presentations easy. Check out what it can do for you here: http://powerpointlabs.info

[![Build status](https://img.shields.io/appveyor/ci/kai33/powerpointlabs/master.svg)](https://ci.appveyor.com/project/kai33/powerpointlabs)

### Dev Prerequisites
0. Install Office 2016, 2013 or 2010 with PowerPoint.
1. Install GitHub for Windows (recommended), or at least, Git.
2. Install Visual Studio 2015 (recommended), 2013, or 2012.
3. Install VSTO (Visual Studio Tools for Office, AKA the Office SDK):<br>
   for VS2012 - http://aka.ms/OfficeDevToolsForVS2012<br>
   for VS2013 - http://aka.ms/OfficeDevToolsForVS2013<br>
   for VS2015 - http://aka.ms/OfficeDevToolsForVS2015<br>

### Setting Up Environment
0. Fork and clone the source codes from this repo.
1. Turn off *Office Version Upgrading*: Open the solution `PowerPointLabs\PowerPointLabs.sln` >> open Tools (menu) >> Options >> Office Tools >> Project Migration >> uncheck ‘Always upgrade to installed version of Office’.
2. Set up *External Office Program*: Open ‘PowerPointLabs’ Properties >> Debug >> select ‘Start external program’ and choose `POWERPNT.exe` in the Office folder. Ensure that both **Debug** and **Release** configurations have set up this.
3. Run the solution by pressing F5 and then PowerPointLabs tab will appear in the PowerPoint ribbon. If you have installed PowerPointLabs add-in, you may have to uninstall it first and rebuild the solution.
