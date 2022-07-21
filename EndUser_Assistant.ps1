# Function create Log folder
    Function CreateLogsFolder
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
#Favicon
New-Item "C:\Users\Default\AppData\Local\ToolsPS" -itemType Directory
$ico = new-object System.Net.WebClient
$ico.DownloadFile("https://raw.githubusercontent.com/ChrisMogis/O365-ManageCalendarPermissions/main/favicon-image.ico","C:\Users\Default\AppData\Local\ToolsPS\favicon-image.ico")

# Create Log Folder
    CreateLogsFolder

#Listbox
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'User Assistant'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'
$form.Icon = "C:\Users\Default\AppData\Local\ToolsPS\favicon-image.ico"

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(100,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
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
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

[void] $listBox.Items.Add('Test Internet Connectivity')
[void] $listBox.Items.Add('Clear DNS Cache')
[void] $listBox.Items.Add('Reset Network Configuration')
[void] $listBox.Items.Add('Remote Control')
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
if ($Action -eq 'Test connectivity')
    {
        PING.EXE 8.8.4.4 | Out-GridView -Title 'Ping Test'       
    }

if ($Action -eq 'Clear DNS Cache')
    {
       powershell.exe Clear-DnsClientCache
    }

if ($Action -eq 'Reset Network Configuration')
    {
        ipconfig /release
        ipconfig /renew
    }

if ($Action -eq 'Remote Control')
    {
        [system.Diagnostics.Process]::start("msra.exe")
    }

if ($Action -eq 'Launch Antivirus Quick Scan')
    {
        Start-Process powershell "Start-Mpscan -ScanType QuickScan"
    }

if ($Action -eq 'Launch Antivirus Full Scan')
    {
        Start-Process powershell "Start-Mpscan -ScanType FullScan"
    }
}