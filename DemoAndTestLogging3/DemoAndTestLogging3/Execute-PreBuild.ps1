<# 

.SYNOPSIS 
A brief description of the function or script. 
This keyword can be used only once in each topic.

.DESCRIPTION
A detailed description of the function or script.
This keyword can be used only once in each topic
		
.PARAMETER firstNamedArgument
The description of a parameter. You can include a Parameter keyword for
each parameter in the function or script syntax.

The Parameter keywords can appear in any order in the comment block, but
the function or script syntax determines the order in which the parameters
(and their descriptions) appear in Help topic. To change the order,
change the syntax.
 
You can also specify a parameter description by placing a comment in the
function or script syntax immediately before the parameter variable name.
If you use both a syntax comment and a Parameter keyword, the description
associated with the Parameter keyword is used, and the syntax comment is
ignored.

.PARAMETER secondNamedArgument
blah blah about secondNamedArgument

.EXAMPLE
A sample command that uses the function or script, optionally followed
by sample output and a description. Repeat this keyword for each example.
.EXAMPLE
Example2

.INPUTS
Inputs to this cmdlet (if any)

.OUTPUTS
Output from this cmdlet (if any)

.NOTES
Additional information about the function or script.

.LINK
The name of a related topic. Repeat this keyword for each related topic.

This content appears in the Related Links section of the Help topic.

The Link keyword content can also include a Uniform Resource Identifier
(URI) to an online version of the same Help topic. The online version 
opens when you use the Online parameter of Get-Help. The URI must begin
with "http" or "https".

.COMPONENT
The technology or feature that the function or script uses, or to which
it is related. This content appears when the Get-Help command includes
the Component parameter of Get-Help.

.ROLE
The user role for the Help topic. This content appears when the Get-Help
command includes the Role parameter of Get-Help.

.FUNCTIONALITY
The intended use of the function. This content appears when the Get-Help
command includes the Functionality parameter of Get-Help.


<ScriptName - Consider Verb-Noun>.ps1

SCC:
	This script is under source code control.  Modifications should be 
	checked into the TFS repository located at 
		<Team Project Collection>
	under a project 
		$<TeamProject>/<Path>

Last Update:

v1.0.0 <Author>, <Date>, <Company>

Be sure to leave two blank lines after end of block comment.
#>

##############################################    
# Script Level Parameters.
##############################################

param
(
# <TODO: Add script level parameters>
    # [switch] $SwitchArg1,
    # [switch] $SwitchArg2,
    [string] $ProjectFileName,
    [string] $Configuration, 
    [string] $Platform,
	[string] $TargetName,
    [switch] $Contents,
    [switch] $Verbose
)

##############################################    
# Script Level Variables
##############################################

<#
$ScriptVar1 = "Foo"
$ScriptVar2 = 42
$ScriptVar3 = @(
    ("Apple")
    ,("Pear")
    ,("Yoghurt")
)
#>

$SCRIPTNAME = $MyInvocation.MyCommand.Name
$SCRIPTPATH = & { $myInvocation.ScriptName }
$CURRENTDIRECTORY = $PSScriptRoot

##############################################
# Main function
##############################################

function Main
{
    if ($SCRIPT:Verbose)
    {
        "SCRIPTNAME         = $SCRIPTNAME"
		"SCRIPTPATH         = $SCRIPTPATH"
		"CURRENTDIRECTORY   = $CURRENTDIRECTORY"
		
        "ProjectFileName    = $ProjectFileName"
        "Configuration      = $Configuration"
        "Platform           = $Platform"
        "TargetName         = $TargetName"
<#
        "ScriptVar2         = $ScriptVar2"
        "ScriptVar3         = $ScriptVar3"
        ""
#>
		"`$Verbose           = $Verbose"
    }
    
    # if ( ! (VerifyPrerequisites))
    # {
    #     LogMessage "Error Verifying Prerequisites" "Main" "Error"
    #     exit
    # }
    # else
    # {
    #     LogMessage "Prerequisites OK" "Main" "Info"
    # }

    # $message = "Beginning " + $SCRIPTNAME + ": " + (Get-Date)
    # LogMessage $message "Main" "Info"
    
# <TODO: Add code, functional calls here to do something cool>

    Set-Location $CURRENTDIRECTORY

    UpdateFileVersion
   
    # $message = "Ending   " + $SCRIPTNAME + ": " + (Get-Date)
    # LogMessage $message "Main" "Info"
}

##############################
# Internal Functions
##############################

function UpdateFileVersion()
{
    $message = "Updating FileVersion"
    LogMessage $message "Info"

    # Read the existing project file

    # This does not preserve the whitespace
    #[xml]$xmlDoc = Get-Content $ProjectFileName

    # Opening like this does
    
    $xmlDoc = [xml]::new()
    $xmlDoc.PreserveWhitespace = $true
    $xmlDoc.Load($ProjectFileName)

    $currentFileVersion = $xmlDoc.Project.PropertyGroup.FileVersion
    $existingFileVersion = $xmlDoc.Project.PropertyGroup[0].FileVersion

    $propertyGroups = $xmlDoc.SelectNodes("//Project/PropertyGroup")

    if ($propertyGroups.Count -ne 1)
    {
        Write-Output "Project file has multiple PropertyGroups"
        # exit
    }

    # Is there an existing property?
    $currentFileVersion2 = $propertyGroups[0].SelectSingleNode("FileVersion")
    # if ($node) {
    # Write-Output "Project already has TreatWarningsAsErrors set to $($node.InnerText)"
    # exit
    # }

    # $currentFileVersion2 = (Select-Xml //ns:FileVersion $ProjectFileName -Namespace @{ ns='http://schemas.microsoft.com/developer/msbuild/2003' })[0].Node

    LogMessage "existingFileVerison >$($existingFileVersion)<  currentFileVersion >$($currentFileVersion)<  currentFileVersion2 >$($currentFileVersion2.InnerText)<" "Info"

    [string]$currentDate = Get-Date -Format "yyyy.MM.dd"

    LogMessage "CurrentDate >$($currentDate)<" "Info"

    # This doesn't work because there are multiple PropertyGroups

    # $xmlDoc.Project.PropertyGroup.FileVersion = $currentDate

    # This works
    $propertyGroups[0].SelectSingleNode("FileVersion").InnerText = $currentDate

    # This does not work
    # $xmlDoc.Project.PropertyGroup[0].FileVersion.InnerText = $currentDate

    # If it was one specific element you can just do like so:
    #$xmlDoc.config.button.command = "C:\Prog32\folder\test.jar"
    # however this wont work since there are multiple elements

    # Since there are multiple elements that need to be 
    # changed use a foreach loop
    #foreach ($element in $xmlDoc.config.button)
    #{
    #    $element.command = "C:\Prog32\folder\test.jar"
    #}
    
    # Then you can save that back to the xml file

    $xmlDoc.Save($ProjectFileName)
}

if ($SCRIPT:Contents)
{
	$myInvocation.MyCommand.ScriptBlock
	exit
}
	
# Call the main function.  Use Dot Sourcing to ensure executed in Script scope.

. Main

#
# End Execute-PreBuild.ps1
#