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

#Listbox
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'User Assistant'
$form.Size = New-Object System.Drawing.Size(500,300)
$form.StartPosition = 'CenterScreen'
$form.Icon = "C:\Users\Default\AppData\Local\ToolsPS\favicon-image.ico"

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
$listBox.Font = New-Object System.Drawing.Font("lucida Console",12,[System.Drawing.FontStyle]::Regular)

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
if ($Action -eq 'Test Internet Connectivity')
    {
        PING.EXE 8.8.4.4 | Out-GridView -Title 'Ping Test'       
    }

if ($Action -eq 'Clear DNS Cache')
    {
       powershell.exe Clear-DnsClientCache
       Start-Sleep -s "5"
       [void] [System.Windows.MessageBox]::Show( "Your DNS cache has been reinitialized", "DNS cache reseted", "OK", "Information" )
    }

if ($Action -eq 'Reset Network Configuration')
    {
        ipconfig /release
        ipconfig /renew
        Start-Sleep -s "5"
        [void] [System.Windows.MessageBox]::Show( "Your network configuration has been reseted", "DNS cache reseted", "OK", "Information" )
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

[void] [System.Windows.MessageBox]::Show("Thank you for using this tool", "Thx", "OK", "Information")