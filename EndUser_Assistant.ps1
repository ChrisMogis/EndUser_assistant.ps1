#Create Tools Directory
New-Item -Force -Path "C:\Users\Default\AppData\Local\Tools_CCMTune\" -ItemType Directory

#Install CCMTune Favicon
Invoke-WebRequest "https://raw.githubusercontent.com/ChrisMogis/O365-ManageCalendarPermissions/main/favicon-image.ico" -Outfile "C:\Users\Default\AppData\Local\Tools_CCMTune\favicon-image.ico"

#Variables
$ServiceName = "IntuneManagementExtension"
$Ico = "C:\Users\Default\AppData\Local\Tools_CCMTune\favicon-image.ico"

#Listbox
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

$form = New-Object System.Windows.Forms.Form
$form.Text = 'User Assistant'
$form.Size = New-Object System.Drawing.Size(500,300)
$form.StartPosition = 'CenterScreen'
$form.Icon = $Ico

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
#$listBox.FontSize = 10
$listBox.Height = 130

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
$Ping = Test-Connection "8.8.4.4" -Count 1 -Quiet
if ($Ping -eq 'True')
        { 
        [void][System.Windows.MessageBox]::Show("OK","Thx","OK","Information")
        }
    else 
        {
        [void][System.Windows.MessageBox]::Show("NOK","Thx","OK","Information")
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
    }

if ($Action -eq 'Launch Antivirus Full Scan')
    {
        Start-Process powershell "Start-Mpscan -ScanType FullScan"
    }
}
