#* FileName: build.ps1
#*==========================================================================================
#* Script Name: build
#* Created: 09/02/2012
#* Author: Lloyd Holman
#* Company: Infigo Software Ltd

#* Requirements:
#* 1. Install PowerShell 2.0+ on local machine
#* 2. Execute from build.bat

#* Parameters: -task* (The build task type to run).
#*	(*) denotes required parameter, all others are optional.

#* Example use to run the default compile task:  
#* .\default.ps1 -task "compile"

#*==========================================================================================
#* Purpose: Wraps the core default.ps1 script and does the following
#* - starts by importing the psake PowerShell module (we have this in a relative path in source control) .
#* - it then invokes the default psake build script in the current working folder (i.e. default.ps1),
#* passing the first parameter passed to the batch file in as the psake task.  Default.ps1 obviously does
#* all the build work for us.
#* - the if block then simply checks the Powershell $error array and returns a count of the errors, this
#* gives TeamCity its non zero return code.
#* - finally the psake PowerShell module is removed.

#*==========================================================================================
#*==========================================================================================
#* SCRIPT BODY
#*==========================================================================================
param([string]$task = "compile", [string]$tc_project_name = "Infigo_Catfish")
Import-Module '.\lib\psake\psake.psm1'; 
#$psake
$psake.use_exit_on_error = $true
Invoke-psake .\Build-CatfishDefault.ps1 -t $task -framework '4.0' -parameters @{"p1"=$tc_project_name};
if ($Error -ne '') 
{ 
	Write-Host "ERROR: $error" -fore RED; 
	Remove-Module psake -ea 'SilentlyContinue';
	exit $error.Count
} 
#$psake
Remove-Module psake -ea 'SilentlyContinue';
