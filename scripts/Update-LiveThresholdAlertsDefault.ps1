#* FileName: Update-LiveThresholdAlertsDefault.ps1
#*==========================================================================================
#* Script Name: build
#* Created: 03/09/2012
#* Author: Lloyd Holman
#* Company: Infigo Software Ltd

#* Requirements:
#* 1. Install PowerShell 2.0+ on local machine
#* 2. Execute from build.bat

#*==========================================================================================
#* Purpose: Reads specifically formatted UKFast ThresholdAlert emails from a defined Exchange 
#* mailbox, Processes the contents and Writes resulting stats to the ISL StatsD/Graphite 
#* instance, for close to real-time monitoring. 
#*==========================================================================================
#*==========================================================================================
#* SCRIPT BODY
#*==========================================================================================
Properties { 
	#Build properties

}

Task default -depends Compile

Task Compile -depends Clean {
	CompileSln("Release")
}

Task Compiled -depends Clean {
	CompileSln("Debug")
}


function DeletePluginsFolder()
{
	if (test-path "$base_dir\Catfish\Presentation\Nop.Web\Plugins")
	{
		Remove-Item "$base_dir\Catfish\Presentation\Nop.Web\Plugins\" -recurse -ErrorAction SilentlyContinue
	}
}
