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

Execute-PostBuild.ps1
#>

##############################################
# Script Level Parameters.
##############################################

[CmdletBinding()]
param
(
    [string] $ProjectFileName
    , [string] $TargetName
    , [string] $PackageVersion
    # , [string] $Configuration
    # , [string] $Platform
    # , [switch] $Contents
    # , [switch] $Verbose
)

##############################################
# Script Level Variables
##############################################

$SCRIPTNAME = $myInvocation.InvocationName
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
        "TargetName         = $TargetName"
        "PackageVersion     = $PackageVersion"
    }

    Set-Location $CURRENTDIRECTORY

    Write-Verbose "Update-VNCNugetPackage $ProjectFileName $TargetName $PackageVersion"
    Update-VNCNugetPackage $ProjectFileName $TargetName $PackageVersion
}

##############################
# Internal Functions
##############################

# Call the main function.  Use Dot Sourcing to ensure executed in Script scope.

. Main

#
# End Execute-PostBuild.ps1
#