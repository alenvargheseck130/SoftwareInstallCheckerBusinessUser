Write-Host "Script Executed"
#Destiantion folder to download installer for all the software
$username = [Environment]::UserName
$destination = "C:\Users\$username\Desktop\"
Install-Module PSWindowsUpdate
$driverName = Read-host "Please enter the disk name of the USB(in a single alphabet): "
$userInput =  $driverName.ToUpper()
$businessInstallerDestination = $userInput+":\AlenScript\SoftwareInstallCheckerBusinessUser";

#FOR TEAMS
$exeName = "teams.exe"
$software = "Teams Machine-Wide Installer";
$installed = $null -ne (Get-WmiObject -Class Win32_product | Where-Object { $_.Name -eq $software }) 
[Net.ServicePointManager]::SecurityProtocol +='tls12'
#Condition to check whether or not the machine has the software installed or not
If( -Not $installed) {
	$source = "https://go.microsoft.com/fwlink/p/?LinkID=2187327&clcid=0x1009&culture=en-ca&country=CA"
	Write-Host "'$software' is NOT installed.";
	#Download software
	Invoke-WebRequest -URI $source -OutFile $destination$exeName
	#install software
	Start-Process $destination$exeName
	Start-Sleep -Seconds 10
	
} 
else {
	Write-Host "'$software' is already installed"
	
}

#FOR GOOGLE CHROME
#sort to find chrome:- Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName | Format-Table -AutoSize
$exeName = "googleChrome.exe"
$software = "Google Chrome";
$installed =  $null -ne (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName | Where-Object { $_.DisplayName -eq "Google Chrome" })  

#Condition to check whether or not the machine has the software installed or not
If(-Not $installed) {
	$source = "https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D%26iid%3D%7B7E39B4A3-F9C3-D714-548A-A3D317C937C5%7D%26lang%3Den%26browser%3D4%26usagestats%3D1%26appname%3DGoogle%2520Chrome%26needsadmin%3Dprefers%26ap%3Dx64-stable-statsdef_1%26installdataindex%3Dempty/update2/installers/ChromeSetup.exe"
	Write-Host "'$software' is NOT installed.";
	#Download software
	Invoke-WebRequest -URI $source -OutFile $destination$exeName
	#install software
	Start-Process $destination$exeName
	Start-Sleep -Seconds 10
} 
else {
	Write-Host "'$software' is already installed"
	
}


#FOR GOOGLE DRIVE
#sort to find google Drive:- Get-PSDrive -PSProvider FileSystem | Select-Object Root
 #get-wmiobject -class win32_logicaldisk | Select-Object DeviceID 
 $installerExeName = "\GoogleDriveSetup.exe"
 $installerPath = $businessInstallerDestination+$installerExeName
 $software = "Google Drive"
 #method to verify if the google driver exists in the local machine disk
 $installed = $null -ne (Get-WmiObject -Class win32_logicaldisk  | Select-Object DeviceID | Where-Object { $_.DeviceID -eq "G:" }) 
 
 #Condition to check whether or not the machine has the software installed or not
 If(-Not $installed) {
	Write-Host "'$software' is NOT installed.";
	#install software
	Start-Process $installerPath
	Start-Sleep -Seconds 10
} 
 else {
	 Write-Host "'$software' is already installed"
}


#FOR OFFICE APPS
#to sort = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$exeName = "microsoftOffice.exe"
$software = "Microsoft Office"
$uninstallKeys = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$O365Exists = $null -ne ($uninstallKeys | Where-Object { $_.GetValue("DisplayName") -eq "Office 16 Click-to-Run Extensibility Component" })
if ($O365Exists) {
Write-Host "'$software' is already installed"

}
else {
	$source = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365ProPlusRetail&platform=Def&language=en-us&TaxRegion=pr&correlationId=224dcf7a-8131-486f-9832-9e4dc2336cdf&token=049fdd00-51a9-42b8-9328-f2159c510d3a&version=O16GA&source=O15OLSO365&Br=2"
	Write-Host "'$software' is NOT installed.";
	#Download software
	Invoke-WebRequest -URI $source -OutFile $destination$exeName
	#install software
	Start-Process $destination$exeName
	Start-Sleep -Seconds 10
}

#FOR WINDOWS ACTIVATION

#method to check if windows is installed - 
#(Get-CimInstance SoftwareLicensingProduct -Filter "Name like 'Windows%'" | 
#Where-Object { $_.PartialProductKey } | select Description, LicenseStatus | select #LicenseStatus | Where-Object {$_.LicenseStatus -eq "1"}) -ne $null
$software = "Windows"
$windowsActivated = $null -ne (Get-CimInstance SoftwareLicensingProduct -Filter "Name like 'Windows%'" | Where-Object { $_.PartialProductKey } | Select-Object Description, LicenseStatus | Select-Object LicenseStatus | Where-Object {$_.LicenseStatus -eq "1"})
if($windowsActivated){
	Write-Host "'$software' is already activated"
}
else{
	<# alternative #1
	$computer = gc env:computername
	$key = "HKGWV-79N29-F89Y8-VQT7H-XD72F"
	$service = get-wmiObject -query "select * from SoftwareLicensingService" -computername $computer
	$service.InstallProductKey($key)
	$service.RefreshLicenseStatus()
	Write-Host "Windows is now activated on your system...."
	
	#>
	slmgr.vbs /ipk HKGWV-79N29-F89Y8-VQT7H-XD72F
}

#FOR DOWNLOADING AND INSTALLING WINDOWS UPDATES
$flag = "true"
If($flag -eq "true"){
	Get-WindowsUpdate -AcceptAll -Install -AutoReboot
	Write-Host "Installed windows updates"
}
else{
	Write-Host "Updates already installed"
}