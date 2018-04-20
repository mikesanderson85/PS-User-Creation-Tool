<###-Account Creation Tool-###
Author - Michael Sanderson

---Quick Guide---
1. XML should be placed in Section 1 between the @" "@. XML can be generated from a GUI interface using VisualStudio and a WPF form project
2. Section 2 converts the form elements into interactable powershell varibales prefixed as $WPF
3. All workable PowerShell should be completed in Section 3
4. Add domain sites to the function fnSelectDomain using the switch format used for previous domains - also add to the combo box at the bottom of the page
---Revision History---
0.20 - Final Version 03/12/15
0.2.1 - Github Test Edit


#>

#============================rt
#========SECTION 1===========
#============================
#ERASE ALL THIS AND PUT XAML BELOW between the @" "@ 
$inputXML = @"
<Window x:Name="Account_Creation_Tool" x:Class="Account_Creation_Tool.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Account Creation Tool" Height="650.667" Width="506" WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <Grid>
        <TextBlock x:Name="label" HorizontalAlignment="Left" Margin="10,10,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Height="22" Width="211" FontWeight="Bold" FontSize="14"><Run Language="en-gb" Text="User Creation Tool"/></TextBlock>
        <Button x:Name="btnSubmit" Content="Submit" HorizontalAlignment="Left" Margin="329,580,0,0" VerticalAlignment="Top" Width="75" TabIndex="21"/>
        <TextBox x:Name="txtBox_fName" HorizontalAlignment="Left" Height="22" Margin="138,92,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="120" TabIndex="1"  />
        <TextBox x:Name="txtBox_sName" HorizontalAlignment="Left" Height="22" Margin="138,120,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="120" TabIndex="3" />
        <Label x:Name="lbl_fName" Content="* First Name: " HorizontalAlignment="Left" Margin="10,90,0,0" VerticalAlignment="Top" Width="123"/>
        <Label x:Name="lbl_sName" Content="* Surname:" HorizontalAlignment="Left" Margin="10,118,0,0" VerticalAlignment="Top" Width="123"/>
        <Button x:Name="btnCancel" Content="Close" HorizontalAlignment="Left" Margin="409,580,0,0" VerticalAlignment="Top" Width="75"  />
        <Label x:Name="lbl_initial" Content="Initial:" HorizontalAlignment="Left" Margin="268,90,0,0" VerticalAlignment="Top" Width="42"/>
        <TextBox x:Name="txtBox_initial" HorizontalAlignment="Left" Height="22" Margin="315,92,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="40" TabIndex="2"  />
        <Label x:Name="lbl_crNumber" Content="CR Number:" HorizontalAlignment="Left" Margin="10,146,0,0" VerticalAlignment="Top" Width="123"/>
        <Label x:Name="lbl_telNum" Content="Tel. Number:" HorizontalAlignment="Left" Margin="10,174,0,0" VerticalAlignment="Top" Width="123"/>
        <Label x:Name="lbl_mobNum" Content="Mob. Number:" HorizontalAlignment="Left" Margin="10,204,0,0" VerticalAlignment="Top" Width="123"/>
        <Label x:Name="lbl_costCenter" Content="Cost Center:" HorizontalAlignment="Left" Margin="10,232,0,0" VerticalAlignment="Top" Width="123"/>
        <Label x:Name="lbl_logonName" Content="* Login ID:" HorizontalAlignment="Left" Margin="10,258,0,0" VerticalAlignment="Top" Width="123"/>
        <DatePicker x:Name="date_Expire" HorizontalAlignment="Left" Margin="328,345,0,0" VerticalAlignment="Top" Width="155" FirstDayOfWeek="Monday" Text="Account Expiry Date" SelectedDateFormat="Short" Visibility="Hidden" Height="24"  IsTodayHighlighted="True" />
        <Label x:Name="lbl_accountExp" Content="Account Expiry:" HorizontalAlignment="Left" Margin="232,344,0,0" VerticalAlignment="Top" Width="123" Visibility="Hidden"/>
        <Label x:Name="lbl_newPassword" Content="* New Password:" HorizontalAlignment="Left" Margin="10,286,0,0" VerticalAlignment="Top" Width="123"/>
        <TextBox x:Name="txtBox_newPassword" HorizontalAlignment="Left" Height="22" Margin="138,288,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="120" Text="Password_001" TabIndex="9" />
        <TextBox x:Name="txtBox_crNumber" HorizontalAlignment="Left" Height="22" Margin="138,148,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="120" TabIndex="4" Text="CR"/>
        <TextBox x:Name="txtBox_telNum" HorizontalAlignment="Left" Height="22" Margin="138,176,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="120" TabIndex="5"  />
        <TextBox x:Name="txtBox_mobNum" HorizontalAlignment="Left" Height="23" Margin="138,205,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="120" TabIndex="6"  />
        <TextBox x:Name="txtBox_costCenter" HorizontalAlignment="Left" Height="22" Margin="138,234,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="120" TabIndex="7" />
        <TextBox x:Name="txtBox_loginName" HorizontalAlignment="Left" Height="22" Margin="138,260,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="120" TabIndex="8"  />
        <Separator HorizontalAlignment="Left" Height="19" Margin="10,396,0,0" VerticalAlignment="Top" Width="473"/>
        <Label x:Name="lbl_company" Content="Company:" HorizontalAlignment="Left" Margin="10,410,0,0" VerticalAlignment="Top" Width="123"/>
        <ComboBox x:Name="cmbbx_domain" HorizontalAlignment="Left" Margin="138,46,0,0" VerticalAlignment="Top" Width="120" SelectedIndex="0" TabIndex="0"  />
        <Label x:Name="lbl_domain" Content="Domain:" HorizontalAlignment="Left" Margin="10,44,0,0" VerticalAlignment="Top" Width="123"/>
        <Label x:Name="lbl_memorableWord" Content="Memorable Word:" HorizontalAlignment="Left" Margin="10,317,0,0" VerticalAlignment="Top" Width="123"/>
        <TextBox x:Name="txtBox_memorableWord" HorizontalAlignment="Left" Height="24" Margin="138,318,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="120" TabIndex="10" />
        <Label x:Name="lbl_location" Content="Location:" HorizontalAlignment="Left" Margin="10,370,0,0" VerticalAlignment="Top" Width="123"/>
        <ComboBox x:Name="cmbbx_location" HorizontalAlignment="Left" Margin="138,372,0,0" VerticalAlignment="Top" Width="120" SelectedIndex="0" TabIndex="12"  />
        <ComboBox x:Name="cmbbx_company" HorizontalAlignment="Left" Margin="138,412,0,0" VerticalAlignment="Top" Width="120" SelectedIndex="0" TabIndex="13"  />
        <Label x:Name="lbl_defenceAccount" Content="Defence Account:" HorizontalAlignment="Left" Margin="10,441,0,0" VerticalAlignment="Top" Width="123"/>
        <ComboBox x:Name="cmbbx_defenceAccount" HorizontalAlignment="Left" Margin="138,443,0,0" VerticalAlignment="Top" Width="120" SelectedIndex="0" TabIndex="14"  />
        <Label x:Name="lbl_hardware" Content="Hardware:" HorizontalAlignment="Left" Margin="10,468,0,0" VerticalAlignment="Top" Width="123"/>
        <ComboBox x:Name="cmbbx_hardware" HorizontalAlignment="Left" Margin="138,470,0,0" VerticalAlignment="Top" Width="120" SelectedIndex="0" TabIndex="15"  />
        <Separator HorizontalAlignment="Left" Height="19" Margin="10,490,0,0" VerticalAlignment="Top" Width="473"/>
        <CheckBox x:Name="chck_magicmikeemail" Content="magicmike Email Address" HorizontalAlignment="Left" Margin="10,507,0,0" VerticalAlignment="Top" IsChecked="True" TabIndex="16" />
        <CheckBox x:Name="chck_GSEemail" Content="GSE Email Address" HorizontalAlignment="Left" Margin="10,528,0,0" VerticalAlignment="Top" TabIndex="18" IsEnabled="True"/>
        <CheckBox x:Name="chck_hideAddress" Content="Hide Address From GAL" HorizontalAlignment="Left" Margin="191,507,0,0" VerticalAlignment="Top" TabIndex="17" IsEnabled="True"/>
        <Separator HorizontalAlignment="Left" Height="19" Margin="10,74,0,0" VerticalAlignment="Top" Width="473"/>
        <Button x:Name="btnresetForm" Content="Reset Form" HorizontalAlignment="Left" Margin="10,580,0,0" VerticalAlignment="Top" Width="75" />
        <CheckBox x:Name="chck_contractor" Content="Contractor?" HorizontalAlignment="Left" Margin="138,349,0,0" VerticalAlignment="Top" TabIndex="11"/>
        <Button x:Name="btn_checkLogin" Content="Check Login" HorizontalAlignment="Left" Margin="268,260,0,0" VerticalAlignment="Top" Width="75" />
        <Button x:Name="btn_randomise" Content="Generate" HorizontalAlignment="Left" Margin="268,319,0,0" VerticalAlignment="Top" Width="75"/>

    </Grid>
</Window>

"@

#Function Show PopUp
Function fnShowPopUp($message, $title, $type) {
	$wshell = New-Object -ComObject Wscript.Shell
	$wshell.Popup($message, 0, $title, $type)
}

<#
check powershell is running as admin
if (!([security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
fnShowPopUp "'User Creation Tool' must be run as an Administrator. Right-click on the icon and select 'Run as administrator'" "Error" 48
Stop-Process -Id $PID
return
}
#>

#check powershell version
if ($PSVersionTable.PSVersion.Major -lt "3") {
	fnShowPopUp "PowerShell 3.0 not detected. Cannot open User Account Creation Tool" "Error" 48
	Stop-Process -Id $PID
	return
}
#check .net version
if ($PSVersionTable.CLRVersion.Major -lt "4") {
	fnShowPopUp ".Net 4.0 not detected. Cannot open User Creation Tool" "Error" 48
	Stop-Process -Id $PID
	return
}
#check AD module is installed
if (!(Get-Module -ListAvailable -Name ActiveDirectory)) {
	fnShowPopUp "Active Directory PowerShell module cannot be found. Cannot open User Creation Tool" "Error" 48
	Stop-Process -Id $PID
	return
}

$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'


[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
	$Form = [Windows.Markup.XamlReader]::Load($reader)
} catch {
	Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
}


#============================
#========SECTION 2===========
#============================ 
#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================

$xaml.SelectNodes("//*[@Name]") | %{
	Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)
}
Function Get-FormVariables {
	if ($global:ReadmeDisplay -ne $true) {
		Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow; $global:ReadmeDisplay = $true
	}
	write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
	get-variable WPF*
}

#Get-FormVariables

#============================
#========SECTION 3===========
#============================  
#===========================================================================
# Actually make the objects work
#===========================================================================

#Remove any log files that older than 30 days
Function Set-Log {
	
	$logfolderName = "UCT Logs"
	$logFolderPath = "$PSScriptRoot"
	
	if (!(Test-Path $logFolderPath\$logfolderName)) {
		New-Item -ItemType Directory -Name $logfolderName -Path $logFolderPath
	}
	
	Get-ChildItem $logFolderPath\$logFolderName | Where-Object {
		!$_.PSIsContainer -and $_.CreationTime -lt $maxDays
	} | Remove-Item -Force
	
	$logFileDate = Get-Date -Format dd_MMMM_yyyy
	$currentLoggedinUser = [Environment]::UserName
	$script:LogFile = $("$logFolderPath\$logfolderName\") + $logFileDate + "_" + $currentLoggedinUser + ".log"
}


#function for writing to log file
function Write-Log {
	Param ([string]$logString)
	
	$logDateTime = Get-Date -Format F
	Add-Content $script:LogFile -Value (($logDateTime + ": ") + ($logString))
	$error.Clear()
}

#Import the ActiveDirectory Module
Import-Module ActiveDirectory

#Routed Event Handlers
$WPFchck_magicmikeemail.Add_Checked({
		$WPFchck_GSEemail.IsEnabled, $WPFchck_magicmikeemail.IsEnabled, $WPFchck_hideAddress.IsEnabled = 'True', 'True', 'True'
	})

$WPFchck_magicmikeemail.Add_UnChecked({
		$WPFchck_GSEemail.IsEnabled = ! 'True'; $WPFchck_hideAddress.IsEnabled = ! 'True'; $WPFchck_GSEemail.IsChecked = ! 'True'; $WPFchck_magicmikeemail.IsChecked = ! 'True'; $WPFchck_hideAddress.IsChecked = ! 'True'
	})

$WPFchck_contractor.Add_Checked({
		$WPFlbl_accountExp.Visibility, $WPFdate_Expire.Visibility = 'Visible', 'Visible'
	})
$WPFchck_contractor.Add_UnChecked({
		$WPFlbl_accountExp.Visibility, $WPFdate_Expire.Visibility = 'Hidden', 'Hidden'
	})

$WPFtxtBox_newPassword.add_GotFocus({
		$WPFtxtBox_newPassword.SelectAll()
	})
$WPFtxtBox_crNumber.add_GotFocus({
		$WPFtxtBox_crNumber.SelectAll()
	})

$WPFcmbbx_domain.Add_SelectionChanged({
		if ($WPFcmbbx_domain.SelectedValue -ne "magicmike.com") {
			$WPFchck_magicmikeemail.IsChecked = ! 'True'
			$WPFchck_magicmikeemail.IsEnabled = ! 'True'
		} else {
			$WPFchck_magicmikeemail.IsEnabled = 'True'
			$WPFchck_magicmikeemail.IsChecked = 'True'
		}
	})

#Select the domain to create the user account in
Function fnSelectDomain {
	switch ($WPFcmbbx_domain.SelectedValue) {
		
		magicmike.com {
			$script:fqdnDomain = "magicmike.com"
			$script:FileServer = "magicmike01"
			$script:exchServer = "magicmike02"
			$script:ADserver = "magicmikead"
			$script:lyncServer = "lspool.magicmike.local"
			$adDiskName = $script:ADserver + "_AD_DISK"
			
			#Set AD Session as PowerShell Drive
			if (!(Get-PSDrive $adDiskName -ErrorAction SilentlyContinue)) {
				if (!(New-PSDrive -Name $adDiskName -PSProvider ActiveDirectory -Server $script:ADserver -Root "//RootDSE/" -scope Global -ErrorAction SilentlyContinue)) {
					Write-Log $Error
					return $false
				}
			}
			cd $("$adDiskName" + ":")
			return $true
		}
		
		domain.com {
			$script:fqdnDomain = "domain.com"
			$script:DCName = "group1ad"
			$script:FileServer = "computer2"
			$script:ADserver = "group1ad"
			$adDiskName = $script:ADserver + "_AD_DISK"
			
			#Set AD Session as PowerShell Drive
			if (!(Get-PSDrive $adDiskName -ErrorAction SilentlyContinue)) {
				if (!(New-PSDrive -Name $adDiskName -PSProvider ActiveDirectory -Server $script:ADserver -Root "//RootDSE/" -scope Global -Credential $script:credential -ErrorAction SilentlyContinue)) {
					Write-Log $Error
					return $false
				}
			}
			cd $("$adDiskName" + ":")
			return $true
		}
	}
}
#end Function fnSelectDomain

#reset the form values
Function fnResetValues {
	$WPFtxtBox_fName.Clear()
	$WPFtxtBox_initial.Clear()
	$WPFtxtBox_sName.Clear()
	$WPFtxtBox_crNumber.text = "CR"
	$WPFtxtBox_telNum.Clear()
	$WPFtxtBox_mobNum.Clear()
	$WPFtxtBox_costCenter.Clear()
	$WPFtxtBox_loginName.Clear()
	$WPFtxtBox_memorableWord.Clear()
	$WPFtxtBox_newPassword.text = "Password_001"
	$WPFcmbbx_domain.SelectedIndex = 0
	$WPFcmbbx_company.SelectedIndex = 0
	$WPFcmbbx_defenceAccount.SelectedIndex = 0
	$WPFcmbbx_hardware.SelectedIndex = 0
	$WPFcmbbx_location.SelectedIndex = 0
	$WPFchck_contractor.IsChecked = ! "True"
	$WPFchck_UKPSemail.IsChecked = "True"
	$WPFdate_Expire.SelectedDate = Get-Date
	$WPFchck_GSEemail.IsChecked = ! "True"
	$WPFchck_hideAddress.IsChecked = ! "True"
	$WPFchck_magicmikeemail.IsChecked = ! "True"
}
#End fnResetValues

#Check login name exists
Function fnCheckLogin($loginName, $answer) {
	$loginName = $WPFtxtBox_loginName.Text.Trim()
	if (!$loginName) {
		fnShowPopUp "The field 'Login Name' is blank" "Error" 48
		return $false
	} else {
		$ADuser = Get-ADUser -Filter {
			SamAccountName -eq $loginName
		}
		
		if ($ADuser) {
			fnShowPopUp "The specified Login ID, '$loginName' already exists in the domain '$script:fqdnDomain'" "Error" 48
			return $true
		} else {
			if ($answer -eq $skip) {
				fnShowPopUp "The specified Login ID, '$loginName' is available for use in the domain '$script:fqdnDomain'" "Success" 64
				return $false
			}
		}
	}
}
#End function fnCheckLogin

#Check First name and Surname don't already exist
Function fnCheckName ($firstName, $surname) {
	$checkName = Get-ADUser -Filter {
		(GivenName -eq $firstName) -and (sn -eq $surname)
	}
	if ($checkName) {
		fnShowPopUp "The user $firstname $surname already exists on the domain" "Error" 48
		return $true
	} else {
		return $false
	}
}
#End function fnCheckName

#Check email groups don't already exist
Function fncheckEmailGroups ($firstName, $surname) {
	$samAccountNameArray = (($firstName + "." + $surname + "@maigcmike.com"), ($firstName + "." + $surname + "@maigcmike.com"))
	
	forEach ($samAccount in $samAccountNameArray) {
		$checkEmailGroup = Get-ADGroup -Filter {
			SamAccountName -eq $samAccount
		}
		$groupsExist += $checkEmailGroup.Name
		$checkEmailGroup = $null
	}
	if ($groupsExist) {
		fnShowPopUp "The following email groups are already applied to the user: $groupsExist Account will not be created!" "Error" 48
		$groupsExist = $null
		return $True
	} else {
		return $False
		
	}
	
}
#End function  fnCheckEmailGroups

#function - create a random memorable word using the CVC formwat 
Function fnRandomise {
	for ($i = 1; $i -le 3; $i++) {
		$random1 = "b", "c", "d", "f", "g", "h", "j", "k", "l", "m", "n", "p", "q", "r", "s", "t", "v", "x", "z" | Get-Random
		$random2 = "a", "e", "i", "o", "u" | GET-RANDOM
		$random3 = "b", "c", "d", "f", "g", "h", "j", "k", "l", "m", "n", "p", "q", "r", "s", "t", "v", "x", "z" | Get-Random
		$randomWord += $random1 + $random2 + $random3
	}
	$WPFtxtBox_memorableWord.Text = $randomWord
} #End function fnRandomise

#function - create the users home folder
Function fnCreateHomeFolder($loginName) {
	Write-Log -logString "Starting home folder creation for user '$loginName' in domain '$script:fqdnDomain'"
	#Set home directory. If Laptop there is no profile directory
	if ($WPFcmbbx_hardware.SelectedValue -eq "Laptop") {
		$homeDirectory = "\\" + $script:FileServer + "." + $script:fqdnDomain + "\Users$\" + $loginName
		$ADhomeDirectory = "\\" + $script:FileServer + "\Users$\" + $loginName
		$roamingProfile = $null
	} else {
		$homeDirectory = "\\" + $script:FileServer + "." + $script:fqdnDomain + "\Users$\" + $loginName
		$ADhomeDirectory = "\\" + $script:FileServer + "\Users$\" + $loginName
		$roamingProfile = "\\" + $script:FileServer + "\Profiles$\" + $loginName
	}
	
	$HomeDrive = 'H:'
	
	#Set the user home drive and roaming profile
	Set-ADUser $loginName -HomeDrive $homeDrive -HomeDirectory $ADhomeDirectory -ProfilePath $roamingProfile
	
	#Create folder on the root of the common Users Share
	#Connect to File Server
	
	
	$invokeCreateFolder = Invoke-Command -ComputerName ($script:FileServer + "." + $script:fqdnDomain) -ScriptBlock {
		try {
			New-Item -Path $args[0] -ItemType directory -Force
			"Created home directory"
		} catch {
			"Error creating home folder: $Error"
		}
	} -ArgumentList $homeDirectory
	
	if ($invokeCreateFolder) {
		Write-Log -logString $invokeCreateFolder
	} else {
		Write-Log -logString "There was a problem connecting to the File Server '$script:FileServer.$script:fqdnDomain'. Home folder not created"
		fnShowPopUp "There was a problem connecting to the File Server '$script:FileServer.$script:fqdnDomain'. Home folder not created" "Error" 48
	}
	
	$invokeCreateFolder = $null
	
	$domain = $script:fqdnDomain
	$IdentityReference = $domain + '\' + $loginName
	
	$invokeACL = Invoke-Command -ComputerName ($script:FileServer + "." + $script:fqdnDomain) -ScriptBlock {
		
		# Set parameters for Access rule
		$fileSystemAccessRights = [System.Security.AccessControl.FileSystemRights]"FullControl"
		$inheritanceFlags = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit, ObjectInherit"
		$propogationFlags = [System.Security.AccessControl.PropagationFlags]"None"
		$accessControl = [System.Security.AccessControl.AccessControlType]"Allow"
		
		$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentList `
		($args[0], $fileSystemAccessRights, $inheritanceFlags, $propogationFlags, $accessControl)
		
		#Get current access rule from home folder for user
		$homefolderACL = Get-Acl $args[1]
		$homefolderACL.AddAccessRule($accessRule)
		
		try {
			Set-Acl -Path $args[1] -AclObject $homefolderACL
			"Access rules applied to home folder"
		} catch {
			"There was a problem setting access rules on the home folder: $Error"
		}
	} -ArgumentList $IdentityReference, $homeDirectory
	
	if ($invokeACL) {
		Write-Log -logString $invokeACL
	} else {
		Write-Log -logString "There was a problem connecting to the File Server '$script:FileServer.$script:fqdnDomain'. Access rules not applied to the Home Folder"
		fnShowPopUp "There was a problem connecting to the File Server '$script:FileServer.$script:fqdnDomain'. Access rules not applied to the Home Folder" "Error" 48
	}
	
	$invokeACL = $null
	
} #End fnCreateHomeFolder


Function fnAddNameToAD {
	#set the input variables from the form
	$currentDate = Get-Date
	if ($WPFtxtBox_newPassword.Text.Trim()) {
		$password = ConvertTo-SecureString $WPFtxtBox_newPassword.Text.Trim() -AsPlainText -Force
	} else {
		$password = $null
	} #convert password to secure string
	if ($WPFtxtBox_fName.Text.Trim()) {
		$firstName = $WPFtxtBox_fName.Text.Trim()
	} else {
		$firstName = $null
	}
	if ($WPFtxtBox_initial.Text.Trim()) {
		$initial = $WPFtxtBox_initial.Text.Trim()
	} else {
		$initial = $null
	}
	if ($WPFtxtBox_sName.Text.Trim()) {
		$surname = $WPFtxtBox_sName.Text.Trim()
	} else {
		$surname = $null
	}
	if ($WPFtxtBox_crNumber.Text.Trim()) {
		$crNumber = $WPFtxtBox_crNumber.Text.Trim()
	} else {
		$crNumber = $null
	}
	if ($crNumber) {
		$crNumberTrimmed = "CR" + $crNumber.TrimStart("CR")
	} else {
		$crNumberTrimmed = $null
	}
	if ($WPFtxtBox_telNum.Text.Trim()) {
		$telNum = $WPFtxtBox_telNum.Text.Trim()
	} else {
		$telNum = $null
	}
	if ($WPFtxtBox_mobNum.Text.Trim()) {
		$mobNum = $WPFtxtBox_mobNum.Text.Trim()
	} else {
		$mobNum = $null
	}
	if ($WPFtxtBox_costCenter.Text.Trim()) {
		$costCenter = $WPFtxtBox_costCenter.Text.Trim()
	} else {
		$costCenter = $null
	}
	if ($WPFtxtBox_loginName.Text.Trim()) {
		$loginName = $WPFtxtBox_loginName.text.Trim()
	} else {
		$loginName = $null
	}
	if ($WPFchck_contractor.IsChecked) {
		$expireDate = $WPFdate_Expire.text
	} else {
		$expiredate = $null
	}
	if ($WPFcmbbx_company.Text.Trim()) {
		$company = $WPFcmbbx_company.Text.Trim()
	} else {
		$company = $null
	}
	if ($WPFtxtBox_memorableWord.Text.Trim()) {
		$memorableWord = $WPFtxtBox_memorableWord.Text.Trim()
	} else {
		$memorableWord = $null
	}
	$RDprofileDirectory = "\\" + $script:FileServer + "\TSProfiles$\" + $loginName
	
	#set the office location details
	switch ($WPFcmbbx_location.selectedValue) {
		Erskine {
			if (!$telNum) {
				$telNum = "0141 814 0000"
			}
			$officeStreet = @"
Erskine Office,
Ferry Road, Bishopton
"@
			$officeCity = "Glasgow"
		}
		Hook {
			if (!$telNum) {
				$telNum = "01256 742 222"
			}
			$officeStreet = @"
Hook Office,
Bartley Way, Hook
"@
			$officeCity = "Hook"
		}
		Telford {
			if (!$telNum) {
				$telNum = $null
			}
			$officeStreet = @"
Telford Office,
Stafford Park 7, Telford
"@
			$officeCity = "Telford"
		}
		Tewksbury{
			if (!$telNum) {
				$telNum = $null
			}
			$officeStreet = @"
Tewksbury Office,
Alexandra Way, Tewksbury
"@
			$officeCity = "Tewksbury"
		}
		Newcastle{
			if (!$telNum) {
				$telNum = $null
			}
			$officeStreet = @"
Newcastle Office,
The Silverlink N, Newcastle
"@
			$officeCity = "Newcastle"
		}
		Bracknell{
			if (!$telNum) {
				$telNum = $null
			}
			$officeStreet = @"
Bracknell Office,
Cain Rd, Amen Corner, Bracknell
"@
			$officeCity = "Bracknell"
		}
		Home/Other{
			if (!$telNum) {
				$telNum = $null
			}
			$officeStreet = @"
Other
"@
			$officeCity = "Other"
		}
	}
	
	#Choose the correct OU for the user depending on chosen attributes in the form
	if ($WPFcmbbx_domain.selectedValue -eq "magicmike.com") {
		switch ($WPFcmbbx_defenceAccount.SelectedValue) {
			"Apollo" {
				if (($WPFchck_contractor.IsChecked) -and ($WPFcmbbx_company.SelectedValue -eq "magicmike")) {
					$ouPath = "CN=Users,DC=magicmike,DC=com"
				} elseIf ((!$WPFchck_contractor.IsChecked) -and ($WPFcmbbx_company.SelectedValue -eq "magicmike")) {
					$ouPath = "CN=Users,DC=magicmike,DC=com"
				} else {
					$ouPath = "CN=Users,DC=magicmike,DC=com"
				}
				
			}
			"Helo" {
				if (($WPFchck_contractor.IsChecked) -and ($WPFcmbbx_company.SelectedValue -eq "magicmike")) {
					$ouPath = "CN=Users,DC=magicmike,DC=com"
				} elseIf ((!$WPFchck_contractor.IsChecked) -and ($WPFcmbbx_company.SelectedValue -eq "magicmike")) {
					$ouPath = "CN=Users,DC=magicmike,DC=com"
				} else {
					$ouPath = "CN=Users,DC=magicmike,DC=com"
				}
				
			}
			"Starbuck" {
				if (($WPFchck_contractor.IsChecked) -and ($WPFcmbbx_company.SelectedValue -eq "magicmike")) {
					$ouPath = "CN=Users,DC=magicmike,DC=com"
				} elseIf ((!$WPFchck_contractor.IsChecked) -and ($WPFcmbbx_company.SelectedValue -eq "magicmike")) {
					$ouPath = "CN=Users,DC=magicmike,DC=com"
				} else {
					$ouPath = "CN=Users,DC=magicmike,DC=com"
				}
				
			}
			"Showboat" {
				if (($WPFchck_contractor.IsChecked) -and ($WPFcmbbx_company.SelectedValue -eq "magicmike")) {
					$ouPath = "CN=Users,DC=magicmike,DC=com"
				} elseIf ((!$WPFchck_contractor.IsChecked) -and ($WPFcmbbx_company.SelectedValue -eq "magicmike")) {
					$ouPath = "CN=Users,DC=magicmike,DC=com"
				} else {
					$ouPath = "CN=Users,DC=magicmike,DC=com"
				}
				
			}
			default {
				if (($WPFchck_contractor.IsChecked) -and ($WPFcmbbx_company.SelectedValue -eq "magicmike")) {
					$ouPath = "CN=Users,DC=magicmike,DC=com"
				} elseIf ((!$WPFchck_contractor.IsChecked) -and ($WPFcmbbx_company.SelectedValue -eq "magicmike")) {
					$ouPath = "CN=Users,DC=magicmike,DC=com"
				} else {
					$ouPath = "CN=Users,DC=magicmike,DC=com"
				}
				
			}
			
		}
	} elseIf ($WPFcmbbx_domain.selectedValue -eq "domain.com") {
		$ouPath = "CN=Users,DC=magicmike,DC=com"
	}
	
	#Error check for null values and if user already exists in domwain
	if ($loginName -eq $null -or $firstName -eq $null -or $surname -eq $null -or $password -eq $null) {
		fnShowPopUp "All mandatory fields (*) must be completed" "Error" 48
	} elseIf (fnCheckLogin $loginName "yes") {
	} elseIf (fnCheckName $firstName $surname) {
	} elseIf (fncheckEmailGroups $firstName $surname) {
	} elseIf (($WPFchck_contractor.IsChecked) -and ($WPFdate_Expire.SelectedDate -le $currentDate)) {
		fnShowPopUp "Please select an expiry date in the future" "Error" 48
	} else {
		Write-Log -logString "==================================================="
		Write-Log -logString "Starting user creation for user: '$loginName' in domain '$script:fqdnDomain'"
		try {
			#create the user
			New-ADUser -Name ($firstName + " " + $surname) -SamAccountName $loginName -UserPrincipalName ($loginName + "@" + $script:fqdnDomain) -givenName $firstName -Initials $initial `
					   -Surname $surname -DisplayName ($surname + ", " + $firstName) -OfficePhone $telNum -MobilePhone $mobNum -StreetAddress $officeStreet -City $officeCity -Description $crNumberTrimmed `                        `
					   -Office $memorableWord -AccountPassword $password -enabled $true -ChangePasswordAtLogon $true -Company $costCenter -AccountExpirationDate $expireDate -path $ouPath
			
		} catch {
			fnShowPopUp "There was an error creating the account. If AD has created the user it will now be deleted. Please close the form and try again. View the log '$script:LogFile' for more details" "Error" 48
			Write-Log -logString $error
			try {
				Remove-ADUser $loginName -Confirm:$false -ErrorAction SilentlyContinue
			} catch {
				return
				Write-Log -logString $error
			}
			
			return
		}
		#Add user to groups
		#Removable Media Permissions
		#if the domain is magicmike.local only include these groups
		if ($WPFcmbbx_domain.selectedValue -eq "magicmike.com") {
			try {
				$Groups = @("AG_EDAM_CD_Read_Auth", "AG_EDAM_RemovableMedia_Deny")
				forEach ($group in $groups) {
					Add-ADGroupMember $group $loginName
					Write-Log -logString "Added user to group '$group'"
				}
			} catch {
				Write-Log -logString "Groups error: $error"
			}
		}
		
		#Internet Access Permissions
		if (($WPFchck_contractor.IsChecked) -or ($WPFcmbbx_company.SelectedValue -eq "None-magicmike")) {
			try {
				Add-ADGroupMember "Users Denied Internet" $loginName
				Write-Log -logString "Added user to group 'Users Denied Internet'"
			} catch {
				Write-Log -logString "Groups error: $error"
			}
		}
		
		#Company Groups
		if ($WPFcmbbx_company.SelectedValue -eq "None-magicmike") {
			try {
				Add-ADGroupMember "Contractors Denied Intranet" $loginName
				Write-Log -logString "Added user to group 'Contractors Denied Intranet'"
			} catch {
				Write-Log -logString "Groups error: $error"
			}
		}
		
		#magicmike.com User Creation Only
		if ($WPFcmbbx_domain.selectedValue -eq "magicmike.com") {
			#Mailbox enable user
			if ($WPFchck_UKPSemail.IsChecked) {
				
				#first connect to to PS Exchange session
				try {
					$script:exchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$script:exchServer.$script:fqdnDomain/PowerShell" -Authentication Kerberos
					Import-PSSession $script:exchSession
					Write-Log -logString "Connected to Exchange Server '$script:exchServer.$script:fqdnDomain' as '$script:exchSession'"
				} catch {
					fnShowPopUp "There was an error connecting to the Exchange Server $script:exchServer. The users Mailbox will not be created but account creation will continue..." "Warning" 64
					Write-Log -logString $error
				}
				
				#assign to correct Mailbox depending on Surname
				if ($surname.substring(0, 1) -le "K") {
					try {
						Enable-Mailbox $loginName -database "(A-K)"
						Write-Log -logString "Enabled Mailbox for '$loginName' in '(A-K)'"
					} catch {
						fnShowPopUp "There was an error creating the users Mailbox. The users Mailbox will not be created but account creation will continue..." "Warning" 64
						Write-Log -logString $error
					}
				} else {
					try {
						Enable-Mailbox $loginName -database "(L-Z)"
						Write-Log -logString "Enabled Mailbox for '$loginName' in '(L-Z)'"
					} catch {
						fnShowPopUp "There was an error creating the users Mailbox. The users Mailbox will not be created but account creation will continue..." "Warning" 64
						Write-Log -logString $error
					}
				}
				
				
				#Create Distribution Groups
				#Get the UsersGUID
				$userPN = Get-ADUser -Identity $loginName
				$userPN = $userPN.UserPrincipalName
				
				#Add maigcmike.com email
				if ($WPFchck_UKPSemail.IsChecked -and $WPFchck_GSEemail.IsChecked) {
					l
					try {
						$GSESamAccountName = ($firstName + "." + $surname + "@maigcmike.com")
						
						New-DistributionGroup ($firstName + "." + $surname + "@maigcmike.com") -members $userPN -SamAccountName $GSESamAccountName -displayname ($surname + ", " + $firstName + " (maigcmike.com)") -Alias ($firstName + "." + $surname + "maigcmike.com") -PrimarySmtpAddress $GSESamAccountName
						Add-AdPermission -Identity $GSESamAccountName -user $userPN -AccessRights GenericRead, GenericWrite -ExtendedRights "Send-As", "Send-To" -InheritanceType None
						Do {
							#Start-Sleep -Seconds 3
							$GSEgroup = Get-ADGroup -Filter {
								SamAccountName -eq $GSESamAccountName
							}
							#write-host "GSE group value is: $GSEgroup"
						} While ($GSEgroup -eq $null)
						
						Set-ADGroup $GSEgroup -Description $GSESamAccountName -GroupScope Global
						Move-ADObject $GSEgroup -TargetPath "CN=Users,DC=magicmike,DC=com"
						Write-Log -logString "Created the distribution group '$GSEgroup'"
					} catch {
						Write-Log -logString $error
					}
				}
				
				#Add magicmike.r.mil.ul e-mail
				if ($WPFchck_UKPSemail.IsChecked -and $WPFchck_magicmikeemail.IsChecked) {
					try {
						$magicmikeSamAccountName = ($firstName + "." + $surname + "@magicmike.com")
						
						New-DistributionGroup ($firstName + "." + $surname + "@magicmike.com") -members $userPN -SamAccountName $magicmikeSamAccountName -displayname ($surname + ", " + $firstName + " (magicmike.com)") -Alias ($firstName + "." + $surname + "magicmike.com") -PrimarySmtpAddress $magicmikeSamAccountName
						Add-AdPermission -Identity $magicmikeSamAccountName -user $userPN -AccessRights GenericRead, GenericWrite -ExtendedRights "Send-As", "Send-To" -InheritanceType None
						Do {
							#Start-Sleep -Seconds 3           
							$magicmikegroup = Get-ADGroup -Filter {
								SamAccountName -eq $magicmikeSamAccountName
							}
							#write-host "magicmike group value is: $magicmikegroup"
						} While ($magicmikegroup -eq $null)
						
						Set-ADGroup $magicmikegroup -Description $magicmikeSamAccountName -GroupScope Global
						Move-ADObject $magicmikegroup -TargetPath "CN=Users,DC=magicmike,DC=com"
						Write-Log -logString "Created the distribution group '$magicmikegroup'"
					} catch {
						Write-Log -logString $error
					}
				}
				
				#if hide from GAL is selected
				if ($WPFchck_hideAddress.IsChecked) {
					try {
						Set-ADUser $loginName -add @{
							msExchHideFromAddressLists	   = "TRUE"
						}
						Write-Log -logString "User $loginName is hidden from GAL"
					} catch {
						Write-Log -logString $error
					}
				}
				
				if ($script:exchSession.state -eq "Opened") {
					Remove-PSSession $script:exchSession
					Write-Log -logString "Closing Exchange Session: '$script:exchSession'"
				}
				
				if ($script:lyncSession.state -eq "Opened") {
					Remove-PSSession $script:lyncSession
					Write-Log -logString "Closing Lync Session: '$script:lyncSession'"
				}
			}
		}
		#set RD Services Profile path
		try {
			$TSuser = Get-ADUser -Filter {
				SamAccountName -eq $loginName
			}
			
			Invoke-Command -ComputerName ($script:ADserver + "." + $script:fqdnDomain) -ScriptBlock {
				$TSuser = $args[0].distinguishedName
				$TSuser = [ADSI]"LDAP://$TSuser"
				$TSuser.psbase.InvokeSet('terminalServicesProfilePath', $args[1])
				$TSuser.SetInfo()
			} -ArgumentList $TSuser, $RDprofileDirectory
			$TSuser = $null
		} catch {
			Write-Log -logString "RD Services profile path error: $error"
		}
		
		#Create the home folders and profile paths
		fnCreateHomeFolder($loginName)
		
		#Show success message  
		Write-Log -logString "The user '$firstName $surname' ($loginName) has been created in the domain '$script:fqdnDomain'"
		fnShowPopUp "The user '$firstName $surname' ($loginName) has been created in the domain '$script:fqdnDomain'" "Success" 64
	}
	
} #End fnAddNameToAD

$WPFbtnSubmit.Add_Click({
		Set-Log
		$Result = fnSelectDomain
		if ($Result) {
			fnAddNameToAD
		} else {
			fnShowPopUp "Cannot connect to AD. Unable to connect to the AD Server. Close the form and try again" "Error" 48
			return
		}
	})

$WPFbtnresetForm.Add_Click({
		fnResetValues
		
	})

$WPFbtn_checkLogin.Add_Click({
		Set-Log
		$Result = fnSelectDomain
		if ($Result) {
			fnCheckLogin
		} else {
			fnShowPopUp "Cannot connect to AD. Unable to connect to the AD Server. Close the form and try again" "Error" 48
			return
		}
		
	})

$WPFbtn_randomise.Add_Click({
		fnRandomise
	})

$WPFbtnCancel.Add_Click({
		#Close any open sessions
		
		Stop-Process -Id $PID
		$Form.Close() | out-null
		
	})

#Populate combo boxes
"magicmike.com", "domain.com" | ForEach-Object {
	$WPFcmbbx_domain.AddChild($_)
}
"Erskine", "Hook", "Telford", "Tewkesbury", "Newcastle", "Bracknell", "Home/Other" | ForEach-Object {
	$WPFcmbbx_location.AddChild($_)
}
"magicmike", "None-magicmike" | ForEach-Object{
	$WPFcmbbx_company.AddChild($_)
}
"Apollo", "Helo", "Starbuck", "Showboat" | ForEach-Object {
	$WPFcmbbx_defenceAccount.AddChild($_)
}
"Laptop", "Desktop" | ForEach-Object {
	$WPFcmbbx_hardware.AddChild($_)
}

$WPFdate_Expire.SelectedDate = Get-Date

#==========================================================================
# Shows the form
#===========================================================================

$Form.ShowDialog() | out-null

