<p>
    <img src="vnc.png">
</p>
## VNCOffice

VSTO Addins for the common MS Office Applications

## Table of Contents

### GrrReferences\
  Dlls used by the addins
  
### SupportTools_Excel\
  VSTO Addin for Excel
  
### SupportTools_PowerPoint\
  VSTO Addin for PowerPoint
  
### SupportTools_Word\
  VSTO Addin for Word
  
### SupportTools_Visio\
  VSTO Addin for Visio
	
### Visio Templates\
  Custom Visio Templates.  Some are used in Videos (infra)
  
### VNCAddinHelper\
  Common code used by the Addins.

## How to use

The addins use WPF UI components from DevExpress.  These will need to be replaced with something free.
All the Addins use VNC.Logging that uses the EnterpriseLibrry logging framework.  If you don't want logging, just comment out the code.

VNC.Logging is available in https://github.com/chrhodes/VNCDevelopment

## Links

## Contributors

Christopher Rhodes

I am retired now and don't work on this code base much.
  
## History
I started automating Office applications decades ago using scripts.  That evolved to VB VSTO addins that evolved to C# VSTO addins.  Most of the early work was done in Excel.  Most of the recent work has been in Visio.  Over the years a lot of stuff was added to the SupportTools_Excel addin that related to companies I worked at.  There were things to manage SharePoint and AzureDevOps.  Then tons of stuff was added to SupportTools_Visio to support my love of Visio. 

### SupportTools_Excel

I am in the process of cleaning up the repo and reposting a thined down version of SupportTools_Excel that reflects my current use.  I am extracting the functionality that is not generally used, e.g. SharePoint, AzureDevOps, and will leave what I still use on a regular basis in SupportTools_Excel.  I will move the TaskPane functionality into WPF windows like in SupportTools_Visio.

### SupportTools_PowerPoint

Very little is here.  Not planning on future work.

### SupportTools_Word

Very little is here.  Not planing on future work.

### SupportTools_Visio

SupportTools_Visio is reflective of the latest thinking.  Visio did not support TaskPanes and I moved to WPF windows to host commands along with the Ribbon.  If you like Visio there is some cool stuff here.

### VisioTemplates

As I learned more about Visio I started creating my own templates that took advantage of the VSTO code behand.  The NapkinMaking folder has a lot of stuff I use all the time.

### VNC.AddInHelper

VNC.AddInHelper has code that is common across the SupportTools_X Addins.  Haven't touched this in years.

## Support

You may contact me at chrhodesvnc@gmail.com

## Documentation

This is it

## Videos &amp; Training

Visio Training Trailer
https://www.youtube.com/watch?v=8fK3XM5jwiE

Visio Training Lesson 1 Setup
https://www.youtube.com/watch?v=HfwaG99psak&t=4s

Visio Training Lesson 2 Building a Template from an Existing Drawing
https://www.youtube.com/watch?v=WnfK4jIzmRg&t=6s

Visio Training Lesson 3 Creating Pages and Page Navigation
https://www.youtube.com/watch?v=RrHnzKxcPoc

Visio Training Lesson 4 Creating State Model Page
https://www.youtube.com/watch?v=MdoMyjWuC8k

Visio Training Lesson 5 Extending Visio with Code
https://www.youtube.com/watch?v=7gTaZF65Io4

Visio Training Lesson 6 Creating a Custom Shape
https://www.youtube.com/watch?v=cB6B8GyzZr0

## NuGet Packages

## Visual Studio Templates

