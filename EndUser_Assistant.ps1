#Create Tools Directory
New-Item -Force -Path "C:\Tools_CCMTune\Tools" -ItemType Directory
New-Item -Force -Path "C:\Tools_CCMTune\Logs" -ItemType Directory

#Install CCMTune Favicon
Invoke-WebRequest "https://raw.githubusercontent.com/ChrisMogis/O365-ManageCalendarPermissions/main/favicon-image.ico" -Outfile "C:\Tools_CCMTune\Tools\favicon-image.ico"

#Variables
$Date = Get-Date
$Logs = "C:\Tools_CCMTune\Logs\EndUserAssistant.log"
$ServiceName = "IntuneManagementExtension"
$Ico = "C:\Tools_CCMTune\Tools\favicon-image.ico"

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

$LinkLabel = New-Object System.Windows.Forms.LinkLabel
#$linkLabel.AutoSize = "True"
$LinkLabel.Location = New-Object System.Drawing.Point(205,240)
$LinkLabel.Size = New-Object System.Drawing.Size(80,20)
$LinkLabel.Font = $font
$LinkLabel.LinkColor = "BLUE"
$LinkLabel.ActiveLinkColor = "RED"
$LinkLabel.Text = "CCMTune.fr"
$LinkLabel.add_Click({[system.Diagnostics.Process]::start("https://ccmtune.fr")})
$Form.Controls.Add($LinkLabel)

$LinkLabel2 = New-Object System.Windows.Forms.LinkLabel
#$linkLabel.AutoSize = "True"
$LinkLabel2.Location = New-Object System.Drawing.Point(10,240)
$LinkLabel2.Size = New-Object System.Drawing.Size(280,20)
$LinkLabel2.Font = $font
$LinkLabel2.LinkColor = "BLUE"
$LinkLabel2.ActiveLinkColor = "RED"
$LinkLabel2.Text = "Logs"
$LinkLabel2.add_Click({[system.Diagnostics.Process]::start($($Logs))})
$Form.Controls.Add($LinkLabel2)

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

[void] $listBox.Items.Add('Computer Information')
[void] $listBox.Items.Add('Test Internet Connectivity')
[void] $listBox.Items.Add('Clear DNS Cache')
[void] $listBox.Items.Add('Reset Network Configuration')
[void] $listBox.Items.Add('Restart Microsoft Intune service')
[void] $listBox.Items.Add('Launch Remote Control')
[void] $listBox.Items.Add('Backup your data')
[void] $listBox.Items.Add('Launch Antivirus Quick Scan')
[void] $listBox.Items.Add('Launch Antivirus Full Scan')

$form.Controls.Add($listBox)

$form.Topmost = $true

$result = $form.ShowDialog() 

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $Action = $listBox.SelectedItem
    $Action
    Write-Output "$($Date) : $($Action)" | Tee-Object -FilePath $Logs -Append

#List of Actions
if ($Action -eq 'Computer Information')
    {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        
        $font = New-Object System.Drawing.Font("Time New Roman", 9)
        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'Computer-Info'
        $form.Size = New-Object System.Drawing.Size(300,205)
        $form.StartPosition = 'CenterScreen'
        $form.Icon = $Ico
         
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(105,120)
        $OKButton.Size = New-Object System.Drawing.Size(75,23)
        $OKButton.Text = 'OK'
        $OKButton.TextAlign = 'MiddleCenter'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.AcceptButton = $OKButton
        $form.Controls.Add($OKButton)
                 
        $hostn = New-Object System.Windows.Forms.Label
        #$hostn.AutoSize = "True"
        $hostn.Location = New-Object System.Drawing.Point(10,10)
        $hostn.Size = New-Object System.Drawing.Size(280,20)
        $hostn.Font = $font
        $hostn.Text = "Computer Name : $(hostname)"
        $form.Controls.Add($hostn)
         
        $os = New-Object System.Windows.Forms.Label
        #$os.AutoSize = "True"
        $os.Location = New-Object System.Drawing.Point(10,90)
        $os.Size = New-Object System.Drawing.Size(280,20)
        $os.Font = $font
        $os.Text = "OS: $((Get-CimInstance win32_operatingsystem).Caption.Trimstart('Microsoft '))"
        $form.Controls.Add($os)
         
        $user = New-Object System.Windows.Forms.Label
        #$user.AutoSize = "True"
        $user.Location = New-Object System.Drawing.Point(10,30)
        $user.Size = New-Object System.Drawing.Size(280,20)
        $user.Font = $font
        $user.Text = "User : $env:username"
        $form.Controls.Add($user)
         
        $ip = New-Object System.Windows.Forms.Label
        #$ip.AutoSize = "True"
        $ip.Location = New-Object System.Drawing.Point(10,70)
        $ip.Size = New-Object System.Drawing.Size(280,20)
        $ip.Font = $font
        $ip.Text = "IP : $((Get-NetAdapter | Where-Object Status -EQ 'up' | Get-NetIPAddress -AddressFamily IPv4).IPAddress)"
        $form.Controls.Add($ip)
        
        $mac = New-Object System.Windows.Forms.Label
        #$mac.AutoSize = "True"
        $mac.Location = New-Object System.Drawing.Point(10,50)
        $mac.Size = New-Object System.Drawing.Size(280,20)
        $mac.Font = $font
        $mac.Text = "MAC : $((Get-NetAdapter | Where-Object Status -EQ 'UP').MacAddress)"
        $form.Controls.Add($mac)
         
        $form.Topmost = $true
         
        $form.ShowDialog()
    }

if ($Action -eq 'Test Internet Connectivity') 
{
$Ping = Test-Connection "8.8.4.4" -Count 1 -Quiet
Write-Output "Ping test to Google" | Tee-Object -FilePath $Logs -Append
Write-Output "Result : $($Ping)" | Tee-Object -FilePath $Logs -Append
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
        Write-Output "Stopping IntuneManagementExtension service" | Tee-Object -FilePath $Logs -Append
            Start-Sleep -s 5
        Start-Process powershell "Net start IntuneManagementExtension"
        Write-Output "Starting IntuneManagementExtension service" | Tee-Object -FilePath $Logs -Append
            Start-Sleep -s 5
        #Check service status
        $Service = Get-Service | Where-Object { $_.Name -eq $serviceName }
        Write-Output "$($ServiceName) is $($Service.status)" | Tee-Object -FilePath $Logs -Append
        if ($service.status -eq "Stopped")
            {
            [void] [System.Windows.MessageBox]::Show( "$($ServiceName) is not restarted, please contact your Administrator","Windows Intune Service", "OK", "Error" )
            }
        else 
            {       
            [void] [System.Windows.MessageBox]::Show( "$($ServiceName) is started","Windows Intune Service", "OK", "Information" )
            }
    }

if ($Action -eq 'Backup your data')
    {
    #Variables
    $Source = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($Source.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $SourcedirectoryName = $Source.SelectedPath
    Write-Host "Directory selected is $SourcedirectoryName"
    }

    $Target = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($Target.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $TargetdirectoryName = $Target.SelectedPath
    Write-Host "Directory selected is $TargetdirectoryName"
    }

    #Data Backup
    Write-Output "Ping test to Google" | Tee-Object -FilePath $Logs -Append
    $Backup = Robocopy.exe $SourcedirectoryName $TargetdirectoryName /E
    Write-Output "$($Backup)" | Tee-Object -FilePath $Logs -Append
    }

if ($Action -eq 'Launch Remote Control')
    {
        Write-Output "Launch $($Action)" | Tee-Object -FilePath $Logs -Append
        [system.Diagnostics.Process]::start("msra.exe")
    }

if ($Action -eq 'Launch Antivirus Quick Scan')
    {
        Write-Output "Launch $($Action)" | Tee-Object -FilePath $Logs -Append
        Start-Process powershell "Start-Mpscan -ScanType QuickScan"
    }

if ($Action -eq 'Launch Antivirus Full Scan')
    {
        Write-Output "Launch $($Action)" | Tee-Object -FilePath $Logs -Append
        Start-Process powershell "Start-Mpscan -ScanType FullScan"
    }
}
