# PowerPointLabs
The typical PowerPoint presentation isn't very interesting. Walls of text or bullet points, with few visuals - it's no wonder audiences find it hard to pay attention.
Your slides don't have to be this way, and it doesn't take a whole lot of effort to make them better.
PowerPointLabs makes creating engaging PowerPoint presentations easy. Check out what it can do for you here: http://powerpointlabs.info

## Dev Prerequisites
0. Install GitHub for Windows (recommended), or at least, Git.
1. Install Visual Studio 2012 or 2013 (at least Professional edition; Premium or Ultimate is better)(You can get it from DreamSpark).
2. Install Office 2010 or 2013.
3. For VS2012, install VSTO (Visual Studio Tools for Office):
download: http://aka.ms/OfficeDevToolsForVS2012

## Setting Up Environment
0. Clone the source code from https://github.com/PowerPointLabs/repo
1. Setup the project by following this article http://ayulin.net/blog/2014/version-control.html
In short, open the project >> open Tools (menu) >> Options >> Office Tools >> Project Migration >> uncheck ‘Always upgrade to installed version of Office’; 
Then open Project (menu) >> ‘Project Name’ Properties >> Debug >> select ‘Start external program’ and choose POWERPNT.exe in your office folder. Ensure that both Debug and Release configurations have set up this.
2. Run the project. If you have installed PowerPointLabs, you may have to uninstall it first.
