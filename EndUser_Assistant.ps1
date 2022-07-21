# Function create Log folder
    Function CreateCCMTuneFolder
{
    If(!(Test-Path "C:\Users\Default\AppData\Local\Tools_CCMTune\"))
    {
    New-Item -Force -Path "C:\Users\Default\AppData\Local\Tools_CCMTune\" -ItemType Directory
		}
		else 
		{ 
    Write-Host "The folder Tools_CCMTune already exists !"
    }
}
#Install CCMTune Favicon
New-Item "C:\Users\Default\AppData\Local\ToolsPS" -itemType Directory
$ico = new-object System.Net.WebClient
$ico.DownloadFile("https://raw.githubusercontent.com/ChrisMogis/O365-ManageCalendarPermissions/main/favicon-image.ico","C:\Users\Default\AppData\Local\ToolsPS\favicon-image.ico")

# Create Log Folder
    CreateCCMTuneFolder

#Variables
$ServiceName = "IntuneManagementExtension"
$PingTest = "172.1.1.1"
$Favico = "C:\Users\Default\AppData\Local\ToolsPS\favicon-image.ico"

#Listbox
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'User Assistant'
$form.Size = New-Object System.Drawing.Size(500,300)
$form.StartPosition = 'CenterScreen'
$form.Icon = $Favico

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(200,190)
$okButton.Size = New-Object System.Drawing.Size(80,25)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please select an action:'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,50)
$listBox.Size = New-Object System.Drawing.Size(455,40)
$listBox.FontSize = 12
$listBox.Height = 130
$listBox.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)

[void] $listBox.Items.Add('Test Internet Connectivity')
[void] $listBox.Items.Add('Clear DNS Cache')
[void] $listBox.Items.Add('Reset Network Configuration')
[void] $listBox.Items.Add('Restart Microsoft Intune service')
[void] $listBox.Items.Add('Launch Remote Control')
[void] $listBox.Items.Add('Launch Antivirus Quick Scan')
[void] $listBox.Items.Add('Launch Antivirus Full Scan')

$form.Controls.Add($listBox)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $Action = $listBox.SelectedItem
    $Action

#List of Actions
if ($Action -eq 'Test Internet Connectivity')
    {
        if($Ping = Test-Connection $PingTest -Count 1 -Quiet) 
            {
                [void] [System.Windows.MessageBox]::Show("Connectivity OK", "Internet connectivity test", "OK", "Information")
            }
        else 
            {
                [void] [System.Windows.MessageBox]::Show("Your internet connection have an issue, please contact your Administrator", "Internet connectivity test", "OK", "Error")
            }
    }

if ($Action -eq 'Clear DNS Cache')
    {
        Start-Process powershell "Clear-DnsClientCache"
            Start-Sleep -s 5
        [void] [System.Windows.MessageBox]::Show( "Your DNS cache has been reinitialized", "DNS cache reseted", "OK", "Information" )
    }

if ($Action -eq 'Reset Network Configuration')
    {
        Start-Process powershell "ipconfig /release"
        Start-Process powershell "ipconfig /renew"
            Start-Sleep -s 5
        [void] [System.Windows.MessageBox]::Show( "Your network configuration has been reseted", "Network configuration reseted", "OK", "Information" )
    }

if ($Action -eq 'Restart Microsoft Intune service')
    {
        Start-Process powershell "Net stop IntuneManagementExtension"
            Start-Sleep -s 5
        Start-Process powershell "Net start IntuneManagementExtension"
            Start-Sleep -s 5
        #Check service status
        $Service = Get-Service | Where-Object { $_.Name -eq $serviceName }
        if ($service.status -eq "Stopped")
            {
            [void] [System.Windows.MessageBox]::Show( "$($ServiceName) is not restarted, please contact your Administrator","Windows Intune Service", "OK", "Error" )
            }
        else 
            {       
            [void] [System.Windows.MessageBox]::Show( "$($ServiceName) is started","Windows Intune Service", "OK", "Information" )
            }
    }

if ($Action -eq 'Launch Remote Control')
    {
        [system.Diagnostics.Process]::start("msra.exe")
    }

if ($Action -eq 'Launch Antivirus Quick Scan')
    {
        Start-Process powershell "Start-Mpscan -ScanType QuickScan"
        [void] [System.Windows.MessageBox]::Show( "Windows Defender Quick Scan is finished", "Windows Defender Quick Scan", "OK", "Information" )
    }

if ($Action -eq 'Launch Antivirus Full Scan')
    {
        Start-Process powershell "Start-Mpscan -ScanType FullScan"
        [void] [System.Windows.MessageBox]::Show( "Windows Defender Full Scan is finished", "Windows Defender Full Scan", "OK", "Information" )
    }
}

[void] [System.Windows.MessageBox]::Show("Thank you for using this tool", "Thx", "OK", "Information")