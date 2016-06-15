#========================================================================
#
# Tool Name	: MDT Quick Import with right-click
# Author 	: Damien VAN ROBAEYS
#
#========================================================================

[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  				| out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.ComponentModel') 				| out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Data')           				| out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')        				| out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 				| out-null
[System.Reflection.Assembly]::LoadWithPartialName('PresentationCore')      				| out-null
[System.Reflection.Assembly]::LoadWithPartialName('MahApps.Metro.Controls.Dialogs')     | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\MahApps.Metro.dll')       				| out-null
[System.Reflection.Assembly]::LoadFrom('assembly\System.Windows.Interactivity.dll') 	| out-null

Add-Type -AssemblyName "System.Windows.Forms"
Add-Type -AssemblyName "System.Drawing"

function LoadXml ($global:filename)
{
    $XamlLoader=(New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($filename)
    return $XamlLoader
}

# Load MainWindow
$XamlMainWindow=LoadXml("PS1ToEXE_Generator.xaml")
$Reader=(New-Object System.Xml.XmlNodeReader $XamlMainWindow)
$Form=[Windows.Markup.XamlReader]::Load($Reader)

[System.Windows.Forms.Application]::EnableVisualStyles()

$browse_exe = $Form.findname("browse_exe") 
$exe_sources_textbox = $Form.findname("exe_sources_textbox") 
$exe_name = $Form.findname("exe_name") 
$icon_sources_textbox = $Form.findname("icon_sources_textbox") 
$browse_icon = $Form.findname("browse_icon") 
$Build = $Form.findname("Build") 
$Choose_ps1 = $Form.findname("Choose_ps1") 

$object = New-Object -comObject Shell.Application  

$openfiledialog1 = New-Object 'System.Windows.Forms.OpenFileDialog'
$openfiledialog1.DefaultExt = "ico"
$openfiledialog1.Filter = "Applications (*.ico) |*.ico"
$openfiledialog1.ShowHelp = $True
$openfiledialog1.filename = "Search for ICO files"
$openfiledialog1.title = "Select an icon"

$Choose_ps1.IsEnabled = $false
$Build.IsEnabled = $false
$exe_name.IsEnabled = $false
$exe_sources_textbox.IsEnabled = $false

$Global:Current_Folder =(get-location).path 
$Global:Conf_File = "$Current_Folder\PS1ToEXE_Generator.conf"
$Global:Temp_Conf_file = "$Current_Folder\PS1ToEXE_Generator_Temp.conf"

$User_Profile = $env:userprofile
$Global:User_Desktop = "$User_Profile\Desktop"

$browse_exe.Add_Click({		
	$folder = $object.BrowseForFolder(0, $message, 0, 0) 
	If ($folder -ne $null) 
		{ 		
			$global:EXE_folder = $folder.self.Path 
			$exe_sources_textbox.Text = $EXE_folder	

			$Global:Folder_name = Split-Path -leaf -path $EXE_folder			
					
			$Choose_ps1.IsEnabled = $true
			$Build.IsEnabled = $true
			$exe_name.IsEnabled = $true
					
			$Dir_EXE_Folder = get-childitem $EXE_folder -recurse
			$List_All_PS1 = $Dir_EXE_Folder | where {$_.extension -eq ".ps1"}				
			foreach ($ps1 in $List_All_PS1)
				{
					$Choose_ps1.Items.Add($ps1)	
					$Global:EXE_PS1_To_Run = $Choose_ps1.SelectedItem						
				}			
			$Choose_ps1.add_SelectionChanged({
			$Global:EXE_PS1_To_Run = $Choose_ps1.SelectedItem				
			write-host $EXE_PS1_To_Run	
			})	
		}
})	


$browse_icon.Add_Click({	

	If($openfiledialog1.ShowDialog() -eq 'OK')
		{	
			$icon_sources_textbox.Text = $openfiledialog1.FileName
			$Global:EXE_Icon_To_Set = $openfiledialog1.FileName 	
		}		
})	


$Build.Add_Click({		

	$EXE_File_Name = $exe_name.Text.ToString()	


	If ($exe_name.Text -eq "") 
		{
			$exe_name.BorderBrush = "Red"
		}
	Else
		{
			copy-item $Conf_File $Temp_Conf_file
			Add-Content $Temp_Conf_file "Setup=C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -sta -WindowStyle Hidden -noprofile -executionpolicy bypass -file %temp%\$EXE_PS1_To_Run"

			$command = "$Current_Folder\WinRAR\WinRAR.exe"	
			& $command a -ep1 -r -o+ -dh -ibck -sfx -iadm "-iicon$EXE_Icon_To_Set" "-z$Temp_Conf_file"  "$User_Desktop\$EXE_File_Name.exe" "$EXE_folder\*"
			
			# -iadm: request administrative access for SFX archive
			# -iiconC: Specify the icon
			# -sfx: Create an SFX self-extracting archive
			# -IIMG: Specify a logo
			# -zC: Read the conf file
			# -r: Repair an archive
			# -ep1: Exlude bas directory from names
			# -inul: Disable error messages
			# -ibck: Run Winrar in Background
			# -y: Assume Yes on all queries

			sleep 5
			remove-item $Temp_Conf_file -force
			
			[System.Windows.Forms.MessageBox]::Show("The EXE $EXE_File_Name has been created on your Desktop") 			
		}
})


# Show FORM
$Form.ShowDialog() | Out-Null