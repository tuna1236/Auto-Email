#IP Stats Email with updates
#Version 1.0 
#Last Updated 2/17/2022

#Adds Log Folder to system
mkdir "C:\Windows_Logs\Email"
attrib +h "C:\Windows_Logs"

#Adds Folder for script
New-Item "C:\Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup" -ItemType Directory -ea 0
mkdir "C:\Scripts\Email"
attrib +h "C:\Scripts"

#Removes Old Folder
Remove-Item -LiteralPath "C:\Email" -Recurse -Force
Remove-Item "C:\Scripts\Email\IP Stats Send Email.ps1" -Recurse -Force


#Waits for network conection
Start-Sleep 100

$ssid = (get-netconnectionProfile).Name
if($ssid -like '*name*' -or $ssid -like '*open*') {
	
	write-host "Conected to Network"
	
} else {
	
	write-host "Not Conected to Network"
	Start-Sleep 20
	}




#Adds Log
$username = $env:UserName
$date = get-date
add-content "C:\Windows_Logs\Email\Email_Started.txt" "Email was started on {$date} the logged in user is {$username} Version 1.0"
add-content "C:\Windows_Logs\Email\Email_Started.txt" " "


#infinite loop for calling connect function  
while(1){

	start-sleep -seconds 40

	# Send the Ip address
	$IP = (Invoke-WebRequest -UseBasicParsing ifconfig.me/ip).Content.Trim()
	
	#Location of Labtop
	$Location = Invoke-RestMethod -Uri ('http://ipinfo.io/'+(Invoke-WebRequest -uri "http://ifconfig.me/ip").Content)


	Add-Type -AssemblyName System.Device #Required to access System.Device.Location namespace
	[cultureinfo]::currentculture = 'en-US';


	#For Debugging
	write-host "Loading Files"


	#Open Ports
	$port = netstat -an | select-string -pattern "listening"
	$port | Out-File -FilePath "C:\Windows_Logs\Email\host.txt"
	$file = "C:\Windows_Logs\Email\host.txt"
	$attachment = new-object System.Net.Mail.Attachment $file


	#Hard Drive Stats
	$drive = Get-PSDrive
	$drive | Out-File -FilePath "C:\Windows_Logs\Email\psdrive.txt"
	$file1 = "C:\Windows_Logs\Email\psdrive.txt"
	$attachment1 = new-object System.Net.Mail.Attachment $file1



	#Network Conections
	$network1 = Get-NetIPConfiguration
	$network1 | Out-File -FilePath "C:\Windows_Logs\Email\network.txt"
	$file2 = "C:\Windows_Logs\Email\network.txt"
	$attachment2 = new-object System.Net.Mail.Attachment $file2


	#task List
	$tasklist = tasklist
	$tasklist | Out-File -FilePath "C:\Windows_Logs\Email\tasklist1.txt"
	$file8 = "C:\Windows_Logs\Email\tasklist1.txt"
	$attachment7 = new-object System.Net.Mail.Attachment $file8

			
	#Log File to prove script is running
	$username = $env:UserName
	$date = get-date
	add-content "C:\Windows_Logs\Email\Email_Stats.txt" "Email Powershell was ran at this time ($date)"
	add-content "C:\Windows_Logs\Email\Email_Stats.txt" "User logged on ($username)"
	add-content "C:\Windows_Logs\Email\Email_Stats.txt" "Labtops Plubic IP Address $IP"
	add-content "C:\Windows_Logs\Email\Email_Stats.txt" "__"
	
	#For Debugging
	write-host "Files Loaded"
	
	
	#Registry Key's
	#$prop1 = Get-ItemProperty Registry::"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Defender\Exclusions\Paths"
	#$prop1 | Out-File -FilePath C:\Email\network.txt
	#$attachment4 = new-object System.Net.Mail.Attachment $file4
	
	#$prop2 = Get-ItemProperty Registry::"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Defender\Exclusions\Extensions"
	#$prop2 | Out-File -FilePath C:\Email\network.txt
	#$attachment5 = new-object System.Net.Mail.Attachment $file5
	
	#$prop3 = Get-ItemProperty Registry::"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Defender\Exclusions\Processes"
	#$prop3 | Out-File -FilePath C:\Email\network.txt
	#$attachment6 = new-object System.Net.Mail.Attachment $file6
	
	
	#For Debugging
	write-host "Quick Settings"
		
	#Host Name
	$host1 = $env:computername
	#Last Loged in user
	$lastlog = $env:UserName

	#For Debugging
	write-host "Preparing Location"

	#Gets Lat and Long
	Add-Type -AssemblyName System.Device #Required to access System.Device.Location namespace
	$GeoWatcher = New-Object System.Device.Location.GeoCoordinateWatcher #Create the required object
	$GeoWatcher.Start() #Begin resolving current locaton

	while (($GeoWatcher.Status -ne 'Ready') -and ($GeoWatcher.Permission -ne 'Denied')) {
		Start-Sleep -Milliseconds 100 #Wait for discovery.
	}  

	if ($GeoWatcher.Permission -eq 'Denied'){
		Write-Error 'Access Denied for Location Information'
	} else {
		$GeoWatcher.Position.Location | Select Latitude,Longitude #Select the relevent results.
		$lat = $GeoWatcher.Position.Location.Latitude
		$long =  $GeoWatcher.Position.Location.Longitude
	}


	#For Debugging
	write-host "Location Loaded"
	write-host "Preparing Email"
	
	
	#Battery Levels
	$batt = (Get-WmiObject Win32_Battery)
	$battlevel = $batt.EstimatedChargeRemaining


	#Sending the Email
	$emailSmtpServer = "smtp.gmail.com"
	$emailSmtpServerPort = "587"
	$emailSmtpUser = "*Username*"
	$emailSmtpPass = "*Password*"
	 
	#Sender
	$emailFrom = "*Username*"
	#Reciver
	$emailTo = "*Username*"


	#Email Body
	$emailMessage = New-Object System.Net.Mail.MailMessage( $emailFrom , $emailTo )
	$emailMessage.Subject = "Labtop Email" 
	$emailMessage.Attachments.Add($attachment)
	$emailMessage.Attachments.Add($attachment1)
	$emailMessage.Attachments.Add($attachment2)
	#$emailMessage.Attachments.Add($attachment4)
	#$emailMessage.Attachments.Add($attachment5)
	#$emailMessage.Attachments.Add($attachment6)
	$emailMessage.Attachments.Add($attachment7)
	$emailMessage.IsBodyHtml = $true #true or false depends
	$emailMessage.Body = "
	<html>
	<head></head>
	<body>
	<p>
	($host1) Status Update
	<br>
	<br>
	Latitutude: $lat<br>
	Longatutude: $long<br>
	Battery Level: %$battlevel<br>
	Network IP Address: $IP<br>
	<br>
	$lastlog
	</p>
	</body>
	</html>
	"
	$SMTPClient = New-Object System.Net.Mail.SmtpClient( $emailSmtpServer , $emailSmtpServerPort )
	$SMTPClient.EnableSsl = $true
	$SMTPClient.Credentials = New-Object System.Net.NetworkCredential( $emailSmtpUser , $emailSmtpPass );
	$SMTPClient.Send( $emailMessage )


	
	#Log File to prove script is running
	$username = $env:UserName
	$date = get-date
	add-content "C:\Windows_Logs\Email\Email_Stats.txt" "Email has been sent at ($date)"
	add-content "C:\Windows_Logs\Email\Email_Stats.txt" "User logged on ($username)"
	add-content "C:\Windows_Logs\Email\Email_Stats.txt" "__"
	
	#For Debugging
	write-host "Email Sent"
	
	#Starts the Loop
	start-sleep -seconds 3600

}
