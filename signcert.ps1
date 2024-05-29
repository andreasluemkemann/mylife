#requires -version 5
##requires -RunAsAdministrator
Using namespace Terminal.Gui

function install-Prerequisites {
    param (
        [string]$OptionalParameters = $null
    )
    #Register-PackageSource -Name NuGet.org -Location https://www.nuget.org/api/v2 -ProviderName NuGet -Force -Trusted -ErrorAction SilentlyContinue
    #Install-Package -Name Terminal.Gui -Force -Source nuget.org -Scope AllUsers -SkipDependencies
    #Install-Package -Name NStack.Core -Force -Source nuget.org -Scope AllUsers -SkipDependencies
    #Get-Package Terminal.GUI -OutVariable p
    #get the assembly for your project
    #Expand-Archive $p.source -DestinationPath .\terminalgui
    #$src = (Get-Package NStack.core).source
    #Expand-Archive $src -DestinationPath .\NStack.Core -Force

    Get-ChildItem .\assemblies\ |
    Select-Object Name, @{Name = "Version"; Expression = { $_.VersionInfo.FileVersion } }
}

#region helper functions
. $PSScriptRoot/ConvertTo-DataTable.ps1
#endregion

#region generalFunctions

#region CertificateFunctions
function install-Cert {
    param (
        [string]$Subject
    )

    $tempFile = [System.IO.Path]::GetTempFileName()

    $InfraCert =
    @"
-----BEGIN CERTIFICATE-----
MIIG6DCCBNCgAwIBAgIQTEKyIf4bR09Zbqo3VUfcdDANBgkqhkiG9w0BAQsFADBW
MQswCQYDVQQGEwJQTDEhMB8GA1UEChMYQXNzZWNvIERhdGEgU3lzdGVtcyBTLkEu
MSQwIgYDVQQDExtDZXJ0dW0gQ29kZSBTaWduaW5nIDIwMjEgQ0EwHhcNMjQwNDE2
MDU1MTE0WhcNMjUwNDE2MDU1MTEzWjCBjTELMAkGA1UEBhMCREUxDzANBgNVBAgM
BkJlcmxpbjEPMA0GA1UEBwwGQmVybGluMS0wKwYDVQQKDCRJbmZyYXNwcmVhZCBV
RyAoaGFmdHVuZ3NiZXNjaHLDpG5rdCkxLTArBgNVBAMMJEluZnJhc3ByZWFkIFVH
IChoYWZ0dW5nc2Jlc2NocsOkbmt0KTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
AgoCggIBAKQuc6Ph2WjBFl39LeHwnNMBLntUMGRFxMbF0jPi1up79yVSbDjGsjKG
wsuvQDVuNcw1kXbJQWBtl8+fp3fskFMt4t+aDwtTY65lmkZ5kTJLweZMXFXWLyVO
71QbeEEZwuLm35ZBh4d9eJbOFRXJwi0ItnSIZgv/D3R3LTS0jy0ow5flWrVrKdCW
zEx7E9jjXjlzQsEca+Z8kljGA6jysk9arfYuyM9WyjL4ZUSl9nSPzRRw8zggg6j+
TeeNrrxcOnsY1YGc8AXP9SdesEkMj/P58VbtAnhabaXO38hjzl42pKeO/RPSp2qx
aH+8by4oE/LFBn1C0iX2DLPR7JBX0GvbMfhbMbnPk299lcgtiEkoyZBXY4nu2EGk
k9qIgtyozc85wfZP3nRt/vfuZj/7cSAiLU2QLGrE+/6eX4yjEj8yN3al4NeMzO04
kZTTyoXrNy6YGpOXcuqqAtbXrOfbPicWJDGx7yitdmplTPtXJpnBrD4D8R7gj0ly
bNx0X8oYBw24drKnslcki/uUsjQSTs8W1wvNRgAnIAnCAOi6rOryLk2lgeQ4VJJC
Ep+GkqBqpH+4p3ElVC26YniSsYWajDiCx9k3CyBWt0qMnHa9cKAQa/iX/fATfb6W
lqCtEiydPw74zh+eeqnA/27ncwGWdywCYRwa6VUhrkwd1CbSzFzHAgMBAAGjggF4
MIIBdDAMBgNVHRMBAf8EAjAAMD0GA1UdHwQ2MDQwMqAwoC6GLGh0dHA6Ly9jY3Nj
YTIwMjEuY3JsLmNlcnR1bS5wbC9jY3NjYTIwMjEuY3JsMHMGCCsGAQUFBwEBBGcw
ZTAsBggrBgEFBQcwAYYgaHR0cDovL2Njc2NhMjAyMS5vY3NwLWNlcnR1bS5jb20w
NQYIKwYBBQUHMAKGKWh0dHA6Ly9yZXBvc2l0b3J5LmNlcnR1bS5wbC9jY3NjYTIw
MjEuY2VyMB8GA1UdIwQYMBaAFN10XUwA23ufoHTKsW73PMAywHDNMB0GA1UdDgQW
BBS6UQ8kk3dEWZwNXI1kkyYXQFyj3DBLBgNVHSAERDBCMAgGBmeBDAEEATA2Bgsq
hGgBhvZ3AgUBBDAnMCUGCCsGAQUFBwIBFhlodHRwczovL3d3dy5jZXJ0dW0ucGwv
Q1BTMBMGA1UdJQQMMAoGCCsGAQUFBwMDMA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG
9w0BAQsFAAOCAgEAZ//W6YbrtmrCWhvViD6owRWKNh4BUfiJeNwH8wPNCs2iY+XH
HLT1gfJgxPiyjkolvsd+CUBLnPcGzPJgVzvc1m4UIt+azpVhH5BhA7XL3HV9ZcLJ
k4rueI79Jg+r4zfYb8rp3JokbBJ601UiJNgSbdqqCqkJAsu/Q3cgXgri8V9npkd+
USX187xUsy87gDvo3pGTIUF5T0fYPOoALsgHvPSdIsCzmK+kqom8p7/GeIRGHp0F
6lQfAjV84hMcCLjJtjBt5XO6x5U7HwLmtyRjr94rE81i568cvU4Bh8hBrx1cuJ6J
RuA2iMNH9OthQKzzQ3vHlD2y2jKkWYFFxGSMp5Nep9Et+LsMQmbt/EC0Mm9gm9rS
2t/q1JHX6EQK4hfx0Vrin/riEtn6mHnibeqkZkULhTP17covZJ8SMfJhRUpen4Cn
C8DJnidzMKxt0VhSN6stxvE+x5fyGbLZF7TkpZbnCbKDj2VyMD2TU7EFnRmWqSLB
MeLJUTC9bsU7w2Cky1W3KNS9UmmOJcn5DkUTqYPoirPvP43sbYwFToTzUCDZflEt
vWZJRDOqR+tWrRZDh495UQ3DfDijK1IUvmys50q2IcJ6E6cn2svzhgbPf1/fdfRS
emNUzEBZmyRyvedv43wGyb5BpEDxisvAvgCpY+STIjOfUPeZE2oJSRjH6jA=
-----END CERTIFICATE-----
"@

    try {
        Write-Host "Installing code signing certificate"
        $InfraCert | Out-File -FilePath $tempFile
        $import = Import-Certificate -FilePath $tempFile -CertStoreLocation Cert:\CurrentUser\TrustedPublisher
        Remove-Item -Path $tempFile
        $result = "Certificate installed successfully"
        Write-Host $result
        return $result
    }
    catch {
        $result = "Failed to install certificate"
        Write-Error "Failed to install certificate"
        return $result
    }
}

function Get-CodeSigningCert {
    param (
        [string]$Subject
    )
    Write-Information "Searching for certificate with subject: $Subject"
    $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object { $_.Subject -like "*$Subject*" }
    if ($cert) {
        Write-Information "Certificate found"
        Write-Information "Subject: " $cert.Subject
        return $cert
    }
    else {
        $warningText = "no Certificate with subject $Subject found"
        Write-Warning $warningText
    }
}

function checkSignature {
    param (
        [string]$FilePath
    )

    $signature = Get-AuthenticodeSignature -FilePath $FilePath
    return $signature.Status.ToString()

    if ($signature.Status -eq "Valid") {
        Write-Host "Signature is valid"
        Write-Host "Signer: " $signature.SignerCertificate.Subject
        return $signature
    }
    else {
        if ($signature.Status -eq "UnknownError") {
            Write-Host "Signature is unknown"
            return $signature.Status.ToString()
        }
        if ($signature.Status -eq "NotSigned") {
            Write-Host "File is not signed"
            return $signature.Status.ToString()
        }
        if ($signature.Status -eq "NotTrusted") {
            Write-Host "Signature is not trusted"
            return $signature.Status.ToString()
        }
        if ($signature.Status -eq "HashMismatch") {
            Write-Host "Signature hash mismatch"
            return $signature.Status.ToString()
        }
        if ($signature.Status -eq "Invalid") {
            Write-Host "Signature is invalid"
            return $signature.Status.ToString()
        }

    }
}
# endregion

Function resetForm {
    $txtComputer.Text = ''
    $TableView.Table = $null
    $txtUser.Text = ''
    $txtPass.Text = ''
    $radioGroup.SelectedItem = 0
    $txtComputer.SetFocus()
    $StatusBar.Items[0].Title = Get-Date -Format g
    $StatusBar.Items[3].Title = 'Ready'
    [Application]::Refresh()
}

function GetPrinterInfo {
    Param(
        $computerName = $txtComputer.Text.ToString().ToUpper()
    )

    $splat = @{
        Class       = 'Win32_Printer'
        ErrorAction = 'Stop'
    }

    $user = $txtUser.Text.ToString()
    $pass = $txtPass.Text.ToString() | ConvertTo-SecureString -AsPlainText -Force
    $cred = [PSCredential]::New($user, $pass)
    $splat['Credential'] = $cred
    $splat['Computername'] = $Computername

    $script:printers = Get-WmiObject @splat | Group-Object -Property Name -AsHashTable -AsString

    $TableView.Table = $script:printers.GetEnumerator() |
    ForEach-Object { $_.value |
        Select-Object Name, DriverName, PortName, Shared, Sharename, Location, Comment
    } | Sort-Object -Property Name | ConvertTo-DataTable

    $TableView.SetFocus()
    $txtComputer.SetFocus()
    [Application]::Refresh()
    [Application]::MainLoop.AddIdle({
            $script:time_start = Get-Date
            UpdateTimer
            return $true
        })
}

function GetServiceInfo {
    Param(
        $computerName = $txtComputer.Text.ToString().ToUpper()
    )

    $splat = @{
        ClassName   = 'win32_service'
        ErrorAction = 'Stop'
    }

    #check for alternate credentials
    If ((-Not [string]::IsNullOrEmpty($txtUser.text.toString()) -AND (-Not [string]::IsNullOrEmpty($txtPass.text.toString())))) {
        $user = $txtUser.Text.ToString()
        $pass = $txtPass.Text.ToString() | ConvertTo-SecureString -AsPlainText -Force
        $cred = [PSCredential]::New($user, $pass)
        Try {
            $cs = New-CimSession -ComputerName $Computername -Credential $cred -ErrorAction Stop
            $splat['CimSession'] = $cs
        }
        Catch {
            [MessageBox]::ErrorQuery('Error!', "Failed to create a session to $Computername. $($_.Exception.Message)", 0, @('OK'))
            $StatusBar.Items[0].Title = Get-Date -Format g
            $StatusBar.Items[3].Title = 'Ready'
            $txtComputer.SetFocus()
            [Application]::Refresh()
            return
        }
    }
    elseif ((-Not [string]::IsNullOrEmpty($txtUser.text.toString()) -AND ([string]::IsNullOrEmpty($txtPass.text.toString())))) {
        [MessageBox]::Query('Alert!', 'Did you forget to enter a password for your alternate credential?')
    }
    else {
        $splat['Computername'] = $Computername
    }

    Switch ($radioGroup.SelectedItem) {
        1 {
            $splat['filter'] = "State='Running'"
        }
        2 {
            $splat['filter'] = "State='Stopped'"
        }
    }
    # Query services using Get-CimInstance
    Try {
        $script:services = Get-CimInstance @splat |
        Group-Object -Property Name -AsHashTable -AsString
        $TableView.Table = $script:services.GetEnumerator() |
        ForEach-Object { $_.value |
            Select-Object Name, State, StartMode, DelayedAutoStart, StartName
        } | Sort-Object -Property Name | ConvertTo-DataTable

        $StatusBar.Items[0].Title = "Updated: $(Get-Date -Format g)"
        $StatusBar.Items[3].Title = $script:services[$TableView.Table.Rows[$TableView.SelectedRow].Name].DisplayName
        $TableView.SetFocus()
    }
    Catch {
        [MessageBox]::ErrorQuery('Error!', "Failed to query services on $($txtComputer.text.ToString()). $($_.Exception.Message)", 0, @('OK'))
        $StatusBar.Items[0].Title = Get-Date -Format g
        $StatusBar.Items[3].Title = 'Ready'
    }
    Finally {
        $txtComputer.SetFocus()
        [Application]::Refresh()
    }

}

Function ExportJson {
    if ($script:services) {
        $ReportDate = Get-Date
        $SaveDialog = [SaveDialog]::New()
        [Application]::Run($SaveDialog)
        if ((-Not $SaveDialog.Canceled) -AND ($SaveDialog.FilePath.ToString() -match 'json$')) {

            $StatusBar.Items[3].Title = "Exported to $($saveDialog.FilePath.ToString())"

            $script:services.GetEnumerator() |
            ForEach-Object { $_.value |
                Select-Object Name, State, StartMode, DelayedAutoStart, StartName,
                @{Name = 'Computername'; Expression = { $txtComputer.Text.toString() } },
                @{Name = 'ReportDate'; Expression = { $ReportDate } }
            } | ConvertTo-Json | Out-File -FilePath $SaveDialog.FilePath.ToString()
            [MessageBox]::Query('Export', "Service data exported to $($saveDialog.FilePath.ToString())", 0, @('OK'))`

        }
    } #if service data is found
    Else {
        [MessageBox]::ErrorQuery('Alert!', "No services to export from $($txtComputer.Text.toString())", 0, @('OK'))
    }
    $txtComputer.SetFocus()
}

Function ExportCsv {
    if ($script:services) {
        $ReportDate = Get-Date
        $SaveDialog = [SaveDialog]::New()
        [Application]::Run($SaveDialog)
        if ((-Not $SaveDialog.Canceled) -AND ($SaveDialog.FilePath.ToString() -match 'csv$')) {

            $StatusBar.Items[3].Title = "Exported to $($saveDialog.FilePath.ToString())"

            $script:services.GetEnumerator() |
            ForEach-Object { $_.value |
                Select-Object -Property Name, State, StartMode, DelayedAutoStart, StartName,
                @{Name = 'Computername'; Expression = { $txtComputer.Text.toString() } },
                @{Name = 'ReportDate'; Expression = { $ReportDate } }
            } | Export-Csv -Path $SaveDialog.FilePath.ToString() -NoTypeInformation

            [MessageBox]::Query('Export', "Service data exported to $($saveDialog.FilePath.ToString())", 0, @('OK'))`

        }
    } #if service data is found
    Else {
        [MessageBox]::ErrorQuery('Alert!', "No services to export from $($txtComputer.Text.toString())", 0, @('OK'))
    }
    $txtComputer.SetFocus()
}

Function ExportCliXML {
    if ($script:services) {
        $ReportDate = Get-Date
        $SaveDialog = [SaveDialog]::New()
        [Application]::Run($SaveDialog)
        if ((-Not $SaveDialog.Canceled) -AND ($SaveDialog.FilePath.ToString() -match 'xml$')) {

            $StatusBar.Items[3].Title = "Exported to $($saveDialog.FilePath.ToString())"

            $script:services.GetEnumerator() |
            ForEach-Object { $_.value |
                Select-Object -Property Name, State, StartMode, DelayedAutoStart, StartName,
                @{Name = 'Computername'; Expression = { $txtComputer.Text.toString() } },
                @{Name = 'ReportDate'; Expression = { $ReportDate } }
            } | Export-Clixml -Path $SaveDialog.FilePath.ToString()

            [MessageBox]::Query('Export', "Service data exported to $($saveDialog.FilePath.ToString())", 0, @('OK'))`

        }
    } #if service data is found
    Else {
        [MessageBox]::ErrorQuery('Alert!', "No services to export from $($txtComputer.Text.toString())", 0, @('OK'))
    }
    $txtComputer.SetFocus()
}
#endregion

#region setup

If ($host.name -ne 'ConsoleHost') {
    Write-Warning 'This must be run in a console host.'
    #Return
}

$dlls = "$PSScriptRoot/assemblies/NStack.dll", "$PSScriptRoot/assemblies/Terminal.Gui.dll"
#$dlls = "$PSScriptRoot\assemblies\Terminal.Gui.dll"
#$dlls = Get-ChildItem -Path "$PSScriptRoot\assemblies" -Filter '*.dll' -File | Select-Object -ExpandProperty FullName
ForEach ($item in $dlls) {
    Try {
        Add-Type -Path $item -ErrorAction Stop
    }
    Catch [System.IO.FileLoadException] {
        Write-Host "already loaded" -ForegroundColor yellow
    }
    Catch {
        Throw $_
    }
}
$scriptVer = '1.2.0'
$TerminalGuiVersion = [System.Reflection.Assembly]::GetAssembly([terminal.gui.application]).GetName().version
#$NStackVersion = [System.Reflection.Assembly]::GetAssembly([nstack.ustring]).GetName().version

[Application]::Init()
[Application]::QuitKey = 27

$Welcome = [Label]@{
    X    = 1
    Y    = 1
    Text = 0
}
#endregion
#region ColorScheme
<#
Disabled
The default foreground and background color for text, when the view is disabled.
Focus
The foreground and background color for text when the view has the focus.
HotFocus
The foreground and background color for text when the view is highlighted (hot) and has focus.
HotNormal
The foreground and background color for text when the view is highlighted (hot).
Normal
The foreground and background color for text when the view is not focused, hot, or disabled.

Black = 0
Blue = 1
BrightBlue = 9
BrightCyan = 11
BrightGreen = 10
BrightMagenta = 13
BrightRed = 12
BrightYellow = 14
Brown = 6
Cyan = 3
DarkGray = 8
Gray = 7
Green = 2
Magenta = 5
Red = 4
White = 15
#>

$n = [Terminal.Gui.Attribute]::new("BrightGreen", "Black") #Normal
$hn = [Terminal.Gui.Attribute]::new("Red", "Black") #HotNormal
$f = [Terminal.Gui.Attribute]::new("White", "Green") #Focus
$hf = [Terminal.Gui.Attribute]::new("Red", "Green") #HotFocus
$d = [Terminal.Gui.Attribute]::new("Red", "DarkGray") #Disabled

$cs = [Terminal.Gui.ColorScheme]::new()
$cs.Normal = $n
$cs.HotNormal = $hn
$cs.Focus = $f
$cs.HotFocus = $hf
$cs.Disabled = $d

$WindowColorScheme = [Terminal.Gui.ColorScheme]::new()
$WindowColorScheme.Normal = [Terminal.Gui.Attribute]::new( "BrightGreen", "Black" )
$WindowColorScheme.HotNormal = [Terminal.Gui.Attribute]::new( "Red", "Black" )
$WindowColorScheme.Focus = [Terminal.Gui.Attribute]::new( "White", "Green" )
$WindowColorScheme.HotFocus = [Terminal.Gui.Attribute]::new( "Red", "Green" )
$WindowColorScheme.Disabled = [Terminal.Gui.Attribute]::new( "Red", "DarkGray" )
#endregion
#endregion

#region tableview
$TableView = [TableView]@{
    X             = 1
    Y             = 6
    Width         = [Dim]::Fill()
    Height        = [Dim]::Fill()
    MultiSelect   = $true
    FullRowSelect = $true
    #    AllowsMarking = $true
    #AutoSize = $True
}
#Keep table headers always in view
$TableView.Style.AlwaysShowHeaders = $True

$TableView.Add_SelectedCellChanged({
        $selectedColumn = $tableView.Table.Columns[$tableView.SelectedColumn]
        $selectedRow = $script:services[$TableView.Table.Rows[$TableView.SelectedRow].Name].Name
        #$StatusBar.Items[3].Title = $script:services[$TableView.Table.Rows[$TableView.SelectedRow].Name].DisplayName
        $StatusBar.Items[3].Title = "Row: $selectedRow Col: $selectedColumn"
    })

#endregion

#region main window and status bar
$Border = @{
    Title         = 'Infraspread Software Signing Tool'
    Effect3DBrush = [Terminal.Gui.Attribute]::new("Gray", "DarkGray")
    BorderBrush   = "Blue"
    Effect3D      = $true
    #BorderThickness = @{'Left' = 1; 'Top' = 1; 'Right' = 1; 'Bottom' = 1 }
    #DrawMarginFrame = $true
    BorderStyle   = [Terminal.Gui.BorderStyle]::Single
}

$window = [Terminal.Gui.Window]@{
    Title       = 'Infraspread Software Signing Tool'
    ColorScheme = $windowColorScheme
    Border      = $Border
}

$StatusBar = [StatusBar]::New(
    @(
        [StatusItem]::New('Unknown', $(Get-Date -Format g), {}),
        [StatusItem]::New('Unknown', 'ESC to quit', {}),
        [StatusItem]::New('Unknown', "v$scriptVer", {}),
        [StatusItem]::New('Unknown', 'Ready', {}),
        [StatusItem]::New('Unknown', 'Activity', {})
    )
)

[Application]::Top.add($StatusBar)
#endregion

#region TUI functions

function UpdateTimer {
    $script:elapsed = (New-TimeSpan -Start $script:time_start -End $(Get-Date)).Seconds
    $StatusBar.Items[4].Title = $script:elapsed
    return $true
}

#region ColorSchemeViewer
function ColorSchemeViewer {
    $ColorSchemeViewerColorScheme = [Terminal.Gui.ColorScheme]::new()
    $ColorSchemeViewerColorScheme.Normal = [Terminal.Gui.Attribute]::new( "BrightGreen", "Black" )
    $ColorSchemeViewerColorScheme.HotNormal = [Terminal.Gui.Attribute]::new( "BrightGreen", "Black" )
    $ColorSchemeViewerColorScheme.Focus = [Terminal.Gui.Attribute]::new( "White", "Green" )
    $ColorSchemeViewerColorScheme.HotFocus = [Terminal.Gui.Attribute]::new( "Red", "Green" )
    $ColorSchemeViewerColorScheme.Disabled = [Terminal.Gui.Attribute]::new( "Red", "DarkGray" )

    $DialogColorSchemeViewer = [Dialog]@{
        Title       = "ColorSchemeViewer"
        #        Width       = [Dim]::Fill()
        #        Height      = [Dim]::Fill()
        ColorScheme = $ColorSchemeViewerColorScheme
    }

    $lblNormal = [Label]@{
        X    = 1
        Y    = 2
        Text = 'cs.Normal'
    }
    $DialogColorSchemeViewer.Add($lblNormal)

    $txtHotNormal = [TextField]@{
        X                       = [POS]::Right($lblNormal) + 1
        Y                       = 2
        Text                    = ' cs.HotNormal'
        TabIndex                = 0
        DesiredCursorVisibility = [CursorVisibility]::Underline
    }
    $DialogColorSchemeViewer.Add($txtHotNormal)
    $textWidth = $txtHotNormal.Text.Length + 1
    $txtHotNormal.Width = $textWidth

    #region Button Close
    $buttonDialogColorSchemeViewerClose = [Terminal.Gui.Button]@{
        Text  = "Close"
        Width = $buttonDialogColorSchemeViewerClose.Text.Length + 2
        X     = 1
        Y     = 4
    }
    $buttonDialogColorSchemeViewerClose.Add_Clicked({
            [Application]::Refresh()
            $DialogColorSchemeViewer.Running = $false
            $window.SetFocus()
        })
    $DialogColorSchemeViewer.Add($buttonDialogColorSchemeViewerClose)
    #endregion

    #region Button Test
    $buttonDialogColorSchemeViewerTest = [Terminal.Gui.Button]@{
        Text  = "Test"
        Width = $buttonDialogColorSchemeViewerTest.Text.Length + 2
        X     = [POS]::Right($buttonDialogColorSchemeViewerClose) + 1
        Y     = 4
    }
    $buttonDialogColorSchemeViewerTest.Add_Clicked({
            $DialogColorSchemeViewer.Running = $false
            [Application]::Refresh()
        })
    $DialogColorSchemeViewer.Add($buttonDialogColorSchemeViewerTest)
    #endregion

    [Application]::Run($DialogColorSchemeViewer)
    #$window.Add($DialogColorSchemeViewer)
    #$DialogColorSchemeViewer.SetFocus()
    $buttonDialogColorSchemeViewerClose.SetFocus()
    [Application]::Refresh()
}
#endregion

#region Backup
function backup-File {
    param (
        [string]$FilePath
    )
    $backupFileBaseName = (Get-Item $FilePath).BaseName.ToString()
    $FileExtension = (Get-Item $FilePath).Extension.ToString()
    $backupTime = Get-Date -Format FileDateTime
    $backupFileName = $backupFileBaseName + "_" + $backupTime + $FileExtension + ".bak"

    try {
        Copy-Item -Path $FilePath -Destination $backupFileName
        return $backupFileName
    }
    catch {
        Write-Error "Failed to create backup file"
        return
    }
}
#endregion

function selectFileDialog {
    #region open a file dialog
    $Dialog = [OpenDialog]::new("Sign Powershell Script", "")
    $Dialog.CanChooseDirectories = $false
    $Dialog.CanChooseFiles = $true
    $Dialog.AllowsMultipleSelection = $false
    $Dialog.DirectoryPath = "$PSScriptRoot"
    $Dialog.AllowedFileTypes = @(".ps1;.psm1;.psd1;.bak")
    $StatusBar.Items[3].Title = $Dialog.FilePath.ToString()
    $StatusBar.SetNeedsDisplay()
    [Application]::Run($Dialog)
    If (-Not $Dialog.Canceled -AND $dialog.FilePath.ToString()) {
        [string]$SelectedFile = $dialog.FilePath.ToString()
        Write-Host $SelectedFile
        [Application]::Refresh()
        return $SelectedFile
    }
}
#endregion

function TextEditor {
    param (
        [string]$FilePath
    )

    $TextEditorColorScheme = [Terminal.Gui.ColorScheme]::new()
    $TextEditorColorScheme.Normal = [Terminal.Gui.Attribute]::new( "BrightGreen", "Black" )
    $TextEditorColorScheme.HotNormal = [Terminal.Gui.Attribute]::new( "Red", "Black" )
    $TextEditorColorScheme.Focus = [Terminal.Gui.Attribute]::new( "White", "Blue" )
    $TextEditorColorScheme.HotFocus = [Terminal.Gui.Attribute]::new( "Red", "Black" )
    $TextEditorColorScheme.Disabled = [Terminal.Gui.Attribute]::new( "White", "Red" )

    $Dialog = [Dialog]@{
        Title       = "Viewing: $FilePath"
        Y           = [Pos]::Top($window) + 2
        Width       = [Dim]::Fill()
        Height      = [Dim]::Fill()
        ColorScheme = $WindowColorScheme
        MenuBar     = $MenuBar
        StatusBar   = $StatusBar
    }

    $Editor = [TextView]@{
        X                       = [Pos]::Left($Dialog) + 1
        Y                       = [Pos]::Top($Dialog) + 2
        ReadOnly                = $true
        Width                   = [Dim]::Fill()
        Height                  = [Dim]::Fill() - 3
        AutoSize                = $true
        AllowsTab               = $false
        CanFocus                = $true
        DesiredCursorVisibility = [CursorVisibility]::Underline
        ColorScheme             = $TextEditorColorScheme
    }
    $Editor.LoadFile($FilePath)
    #$Editor.Text = Get-Content -Path $FilePath -Raw
    #$Dialog.Add($MenuBar)
    #$Dialog.Add($StatusBar)
    $Dialog.Add($Editor)

    $EditorButtonClose = [Terminal.Gui.Button]@{
        Text  = "Close"
        Width = $EditorButtonClose.Text.Length + 2
        X     = 1
        Y     = [Pos]::Bottom($Editor) + 1
    }
    $Dialog.Add($EditorButtonClose)
    $EditorButtonClose.Add_Clicked({
            [Application]::Refresh()
            $Dialog.Running = $false
        })
    $EditorButtonSign = [Terminal.Gui.Button]@{
        Text  = "Sign"
        Width = $EditorButtonSign.Text.Length + 2
        X     = [Pos]::Right($EditorButtonClose) + 1
        Y     = [Pos]::Bottom($Editor) + 1
    }
    $Dialog.Add($EditorButtonSign)

    $EditorButtonSign.Add_Clicked({
            [Application]::Refresh()
            $Dialog.Running = $false
        })
    $EditorButtonCheckSignature = [Terminal.Gui.Button]@{
        Text  = "Check Signature"
        Width = $EditorButtonCheckSignature.Text.Length + 2
        X     = [Pos]::Right($EditorButtonSign) + 1
        Y     = [Pos]::Bottom($Editor) + 1
    }
    $Dialog.Add($EditorButtonCheckSignature)

    $EditorButtonCheckSignature.Add_Clicked({
            [string]$signature = (Get-AuthenticodeSignature -FilePath $FilePath).StatusMessage
            #            $signature = checkSignature -FilePath $FilePath
            [string]$msgText = "File: $FilePath`nSignature: $signature"
            [MessageBox]::Query('Signature', "Result: $msgText", 0, @('OK'))
            [Application]::Refresh()
            $Dialog.Running = $false
        })
    $EditorButtonBackup = [Terminal.Gui.Button]@{
        Text  = "Backup"
        Width = $EditorButtonBackup.Text.Length + 2
        X     = [Pos]::Right($EditorButtonCheckSignature) + 1
        Y     = [Pos]::Bottom($Editor) + 1
    }
    $Dialog.Add($EditorButtonBackup)
    $EditorButtonBackup.Add_Clicked({
            $backuppedFile = backup-File -FilePath $FilePath
            [Application]::Refresh()
            #$Dialog.Running = $false
            $StatusBar.Items[0].Title = Get-Date -Format g
            $StatusBar.Items[3].Title = 'File Backup complete: ' + $backuppedFile
            $StatusBar.SetNeedsDisplay()
            [MessageBox]::Query('Backup', "Backup created as $backuppedFile", 0, @('OK'))
        })
    $EditorButtonClose.SetFocus()
    [Terminal.Gui.Application]::Run($Dialog)
    [Application]::Refresh()
}

function PrinterFunction {
    $TableView.Table = $null
    #region Add a label and text box for the computer name
    $lblComputer = [Label]@{
        X    = 1
        Y    = 2
        Text = 'Computer Name:'
    }
    $window.Add($lblComputer)

    $txtComputer = [TextField]@{
        X        = 10
        Y        = 2
        Width    = 35
        Text     = 'print.infraspread.net'
        TabIndex = 0
    }

    #make the computername always upper case
    $txtComputer.Add_TextChanged({
            $txtComputer.Text = $txtComputer.Text.ToString().ToUpper()
        })

    $window.Add($txtComputer)
    #endregion

    #region alternate credentials
    $CredentialFrame = [FrameView]::New('Credentials')
    $CredentialFrame.x = 50
    $CredentialFrame.y = 1
    $CredentialFrame.width = 40
    $CredentialFrame.Height = 5

    $lblUser = [Label]@{
        Text = 'Username:'
        X    = 1
    }
    $CredentialFrame.Add($lblUser)

    $txtUser = [TextField]@{
        X        = $lblUser.Frame.Width + 2
        Width    = 25
        TabIndex = 1
        Text     = 'infraspread\administrator'
    }
    $CredentialFrame.Add($txtUser)

    $lblPass = [Label]@{
        Text = 'Password:'
        X    = 1
        Y    = 2
    }
    $CredentialFrame.Add($lblPass)
    $txtPass = [TextField]@{
        X        = $lblUser.Frame.Width + 2
        Y        = 2
        Width    = 25
        Secret   = $True
        TabIndex = 2
        Text     = 'P@ssw0rd!'
    }
    $CredentialFrame.Add($txtPass)

    $Window.Add($CredentialFrame)
    #endregion

    #region Add a button to query services
    $btnQuery = [Button]@{
        X        = 1
        Y        = 4
        Text     = '_Get Info'
        TabIndex = 4
    }
    $btnQuery.Add_Clicked({
            Switch ($radioGroup.SelectedItem) {
                0 {
                    $select = 'All'
                }
                1 {
                    $select = 'Running'
                }
                2 {
                    $select = 'Stopped'
                }
            }
            $StatusBar.Items[3].Title = "Getting $select printers from $($txtComputer.Text.ToString().toUpper())"
            $StatusBar.SetNeedsDisplay()
            $tableView.RemoveAll()
            $tableView.Clear()
            $tableView.SetNeedsDisplay()
            [Application]::Refresh()

            GetPrinterInfo
            [Application]::MainLoop.RemoveIdle({
                    UpdateTimer
                    return $false
                })
        })
    $window.Add($btnQuery)
    #endregion

    #region add radio group
    $RadioGroup = [RadioGroup]::New(15, 3, @('_All', '_Running', '_Stopped'), 0)
    $RadioGroup.DisplayMode = 'Horizontal'
    $RadioGroup.TabIndex = 3
    #put the radio group next to the Get Info button
    $RadioGroup.y = $btnQuery.y
    $Window.Add($RadioGroup)
    #endregion
    $lblFilter = [Label]@{
        Text  = 'Filter:'
        Width = 9
        X     = 1
        Y     = $btnQuery.Y + 1
    }
    $window.Add($lblFilter)

    $txtFilter = [TextField]@{
        X        = [POS]::Right($lblFilter) + 1
        Y        = $lblFilter.Y
        Width    = 25
        TabIndex = 5
        Text     = 'ECO'
    }

    $window.Add($txtFilter)

    $btnFilter = [Button]@{
        X        = [POS]::Right($txtFilter) + 1
        Y        = $lblFilter.Y
        Text     = 'Go'
        TabIndex = 6
    }
    $btnFilter.Add_Clicked({
            [string]$filter = $txtFilter.Text.ToString()
            #$script:printers.getEnumerator() | Where-Object { $_.Name -like "*$($filter)*" } | Out-GridView
            $script:printersFiltered = $script:printers.getEnumerator() | Where-Object { $_.Name -like "*$($filter)*" }

            $TableView.Table = $script:printersFiltered |
            ForEach-Object { $_.Value |
                Select-Object Name, DriverName, PortName, Shared, Sharename, Location, Comment
            } | Sort-Object -Property Name | ConvertTo-DataTable
            $TableView.SetFocus()
            $txtComputer.SetFocus()
            [Application]::Refresh()
        })
    <#
    $TableView.Table = $script:printers.GetEnumerator() |
    ForEach-Object { $_.value |
        Select-Object Name, DriverName, PortName, Shared, Sharename, Location, Comment
    } | Sort-Object -Property Name | ConvertTo-DataTable
    #>
    $window.Add($btnFilter)
    $window.Add($TableView)
    $btnFilter.SetFocus()
    [Application]::Run()
}

function ServiceInfo {
    $TableView.Table = $null
    #region Add a label and text box for the computer name
    $lblComputer = [Label]@{
        X    = 1
        Y    = 2
        Text = 'Computer Name:'
    }
    $window.Add($lblComputer)

    $txtComputer = [TextField]@{
        X        = 10
        Y        = 2
        Width    = 35
        Text     = $env:COMPUTERNAME
        TabIndex = 0
    }

    #make the computername always upper case
    $txtComputer.Add_TextChanged({
            $txtComputer.Text = $txtComputer.Text.ToString().ToUpper()
        })

    $window.Add($txtComputer)
    #endregion

    #region alternate credentials
    $CredentialFrame = [FrameView]::New('Credentials')
    $CredentialFrame.x = 50
    $CredentialFrame.y = 1
    $CredentialFrame.width = 40
    $CredentialFrame.Height = 5

    $lblUser = [Label]@{
        Text = 'Username:'
        X    = 1
    }
    $CredentialFrame.Add($lblUser)

    $txtUser = [TextField]@{
        X        = $lblUser.Frame.Width + 2
        Width    = 25
        TabIndex = 1
    }
    $CredentialFrame.Add($txtUser)

    $lblPass = [Label]@{
        Text = 'Password:'
        X    = 1
        Y    = 2
    }
    $CredentialFrame.Add($lblPass)
    $txtPass = [TextField]@{
        X        = $lblUser.Frame.Width + 2
        Y        = 2
        Width    = 25
        Secret   = $True
        TabIndex = 2
    }
    $CredentialFrame.Add($txtPass)

    $Window.Add($CredentialFrame)
    #endregion

    #region Add a button to query services
    $btnQuery = [Button]@{
        X        = 1
        Y        = 4
        Text     = '_Get Info'
        TabIndex = 4
    }
    $btnQuery.Add_Clicked({
            Switch ($radioGroup.SelectedItem) {
                0 {
                    $select = 'All'
                }
                1 {
                    $select = 'Running'
                }
                2 {
                    $select = 'Stopped'
                }
            }
            $StatusBar.Items[3].Title = "Getting $select services from $($txtComputer.Text.ToString().toUpper())"
            $StatusBar.SetNeedsDisplay()
            $tableView.RemoveAll()
            $tableView.Clear()
            $tableView.SetNeedsDisplay()
            [Application]::Refresh()
            GetServiceInfo
        })
    $window.Add($btnQuery)
    #endregion

    #region add radio group
    $RadioGroup = [RadioGroup]::New(15, 3, @('_All', '_Running', '_Stopped'), 0)
    $RadioGroup.DisplayMode = 'Horizontal'
    $RadioGroup.TabIndex = 3
    #put the radio group next to the Get Info button
    $RadioGroup.y = $btnQuery.y
    $Window.Add($RadioGroup)
    #endregion

    $window.Add($TableView)
    $txtComputer.SetFocus()
    [Application]::Run()
}
#endregion
#endregion

#region add menus
$MenuItemFileInstallCert = [MenuItem]::New('Install Certificate', '', {
        $findInfraCert = Get-CodeSigningCert -Subject 'Infraspread'
        If ($findInfraCert) {
            [MessageBox]::Query('Certificate Found', 'Certificate found', 0, @('OK'))
        }
        Else {
            [MessageBox]::Query('Certificate Not Found', 'Certificate not found', 0, @('OK'))

            [string]$InstallCertResult = install-Cert
            [MessageBox]::Query('Install Certificate', $InstallCertResult, 0, @('OK'))
        }
    }
)


$MenuItemFileOpen = [MenuItem]::New('_Open', '', {
        $myFile = selectFileDialog
        If ($myFile) {
            #[MessageBox]::Query('Open', 'You selected ' + $myFile, 0, @('OK', 'Cancel'))
            $StatusBar.Items[0].Title = Get-Date -Format g
            $StatusBar.Items[3].Title = 'Selected ' + $myFile
            $StatusBar.SetNeedsDisplay()
            TextEditor -FilePath $myFile
            [Application]::Refresh()
        }
    })
$MenuItemFileView = [MenuItem]::New('_View', '', {
        $myFile2 = selectFileDialogFrame2
        If ($myFile2) {
            #[MessageBox]::Query('Open', 'You selected ' + $myFile, 0, @('OK', 'Cancel'))
            $StatusBar.Items[0].Title = Get-Date -Format g
            $StatusBar.Items[3].Title = 'Selected ' + $myFile2
            $StatusBar.SetNeedsDisplay()
            TextEditorFrame2 -FilePath $myFile2
            [Application]::Refresh()
        }
        $window.Remove($Editor)
        [Application]::Refresh()
    })
$MenuItemFileQuit = [MenuItem]::New('_Quit', '', { [Application]::RequestStop() })
$MenuBarItemFile = [MenuBarItem]::New('_Software Signing', @($MenuItemFileInstallCert, $MenuItemFileOpen, $MenuItemFileView, $MenuItemFileQuit))

$MenuItemPrinterMgmt = [MenuItem]::New('Printer Management', '', { PrinterFunction })
$MenuBarItemPrinterMgmt = [MenuBarItem]::New('_Printer Management', @($MenuItemPrinterMgmt))

$MenuItemOpenConfig = [MenuItem]::New('Open Config', '', { ReadConfiguration })
$MenuBarItemConfiguration = [MenuBarItem]::New('_Configuration', @($MenuItemOpenConfig))

$MenuItemServices = [MenuItem]::New('_Get Services', '', { ServiceInfo })
$MenuItemServicesClear = [MenuItem]::New('_Clear form', '', { resetForm })
$MenuItemServicesExportCsv = [MenuItem]::New('as Csv', '', { ExportCsv })
$MenuItemServicesExportJson = [MenuItem]::New('as JSON', '', { ExportJson })
$MenuItemServicesExportXml = [MenuItem]::New('as XML', '', { ExportCliXML })
$MenuBarItemServices = [MenuBarItem]::New('_Services', @($MenuItemServices, $MenuItemServicesClear, $MenuItemServicesExportCsv, $MenuItemServicesExportJson, $MenuItemServicesExportXml))

#this is what will be on the menu bar


$about = @"
$($MyInvocation.MyCommand) v$scriptVer
PSVersion $($PSVersionTable.PSVersion)
Terminal.Gui $TerminalGuiVersion
NStack 0000
"@
$MenuItem3 = [MenuItem]::New('A_bout', '', { [MessageBox]::Query('About', $About) })
$MenuItem4 = [MenuItem]::New('_Documentation', '', { [MessageBox]::Query('Help', 'To be completed') })
$MenuItemColorScheme = [MenuItem]::New('Color _Scheme', '', { ColorSchemeViewer })
$MenuBarItem2 = [MenuBarItem]::New('_Help', @($MenuItem3, $MenuItem4, $MenuItemColorScheme))

$MenuBar = [MenuBar]::New(@($MenuBarItemFile, $MenuBarItemPrinterMgmt, $MenuBarItemServices, $MenuBarItem2, $MenuBarItemConfiguration))
$Window.Add($MenuBar)
#endregion







$Frame1 = [Terminal.Gui.FrameView]::new()
$Frame1.Y = [Terminal.Gui.Pos]::Bottom($MenuBar) + 1
$Frame1.Width = [Terminal.Gui.Dim]::Percent(50)
$Frame1.Height = [Terminal.Gui.Dim]::Fill()
$Frame1.Title = "Frame 1"
#$Window.Add($Frame1)

$Frame2 = [Terminal.Gui.FrameView]::new()
$Frame2.Y = $Frame1.Y
$Frame2.Width = [Terminal.Gui.Dim]::Percent(50)
$Frame2.Height = [Terminal.Gui.Dim]::Fill()
$Frame2.X = [Terminal.Gui.Pos]::Right($Frame1)
$Frame2.Title = "Frame 2"
#$Window.Add($Frame2)

function selectFileDialogFrame2 {
    #region open a file dialog
    $Dialog = [OpenDialog]::new("Sign Powershell Script", "")
    $Dialog.CanChooseDirectories = $false
    $Dialog.CanChooseFiles = $true
    $Dialog.AllowsMultipleSelection = $false
    $Dialog.DirectoryPath = "$PSScriptRoot"
    $Dialog.AllowedFileTypes = @(".ps1;.psm1;.psd1;.bak")
    $StatusBar.Items[3].Title = $Dialog.FilePath.ToString()
    $StatusBar.SetNeedsDisplay()
    [Application]::Run($Dialog)
    If (-Not $Dialog.Canceled -AND $dialog.FilePath.ToString()) {
        [string]$SelectedFile = $dialog.FilePath.ToString()
        Write-Host $SelectedFile
        [Application]::Refresh()
        return $SelectedFile
    }
}
#endregion

function TextEditorFrame2 {
    param (
        [string]$FilePath
    )
    #$window.Add($Frame2)

    $TextEditorColorScheme = [Terminal.Gui.ColorScheme]::new()
    $TextEditorColorScheme.Normal = [Terminal.Gui.Attribute]::new( "BrightGreen", "Black" )
    $TextEditorColorScheme.HotNormal = [Terminal.Gui.Attribute]::new( "Red", "Black" )
    $TextEditorColorScheme.Focus = [Terminal.Gui.Attribute]::new( "White", "Blue" )
    $TextEditorColorScheme.HotFocus = [Terminal.Gui.Attribute]::new( "Red", "Black" )
    $TextEditorColorScheme.Disabled = [Terminal.Gui.Attribute]::new( "White", "Red" )

    $Dialog2 = [Dialog]@{
        Title       = "Dialog2: $FilePath"
        Y           = [Pos]::Top($Frame2) + 1
        Width       = [Dim]::Fill()
        Height      = [Dim]::Fill()
        ColorScheme = $WindowColorScheme
    }

    $Editor2 = [TextView]@{
        X                       = [Pos]::Left($Dialog2)
        Y                       = [Pos]::Top($Dialog2) + 1
        ReadOnly                = $true
        Width                   = [Dim]::Fill()
        Height                  = [Dim]::Fill() - 3
        AutoSize                = $true
        AllowsTab               = $false
        CanFocus                = $true
        DesiredCursorVisibility = [CursorVisibility]::Underline
        ColorScheme             = $TextEditorColorScheme
    }
    $Editor2.LoadFile($FilePath)
    $Dialog2.Add($Editor2)
    #$Editor.Text = Get-Content -Path $FilePath -Raw

    $EditorButtonClose = [Terminal.Gui.Button]@{
        Text  = "Close"
        Width = 9
        X     = 2
        Y     = 25
    }
    $Frame2.Add($EditorButtonClose)

    $EditorButtonClose.Add_Clicked({
            [Application]::Refresh()
            $Frame2.RemoveAll()
        })

    $EditorButtonSign = [Terminal.Gui.Button]@{
        Text  = "Sign"
        Width = $EditorButtonSign.Text.Length + 2
        X     = [Pos]::Right($EditorButtonClose) + 1
        Y     = $Editor2.Y + 1
    }
    $Dialog2.Add($EditorButtonSign)

    $EditorButtonSign.Add_Clicked({
            [Application]::Refresh()
            $Dialog2.Running = $false
        })
    $EditorButtonBackup = [Terminal.Gui.Button]@{
        Text  = "Backup"
        Width = $EditorButtonBackup.Text.Length + 2
        X     = [Pos]::Right($EditorButtonSign) + 1
        Y     = $Editor2.Y + 1
    }
    $Dialog2.Add($EditorButtonBackup)
    $EditorButtonBackup.Add_Clicked({
            $backuppedFile = backup-File -FilePath $FilePath
            [Application]::Refresh()
            #$Dialog.Running = $false
            $StatusBar.Items[0].Title = Get-Date -Format g
            $StatusBar.Items[3].Title = 'File Backup complete: ' + $backuppedFile
            $StatusBar.SetNeedsDisplay()
            [MessageBox]::Query('Backup', "Backup created as $backuppedFile", 0, @('OK'))
        })
    $EditorButtonClose.SetFocus()
    #[Terminal.Gui.Application]::Run($Dialog2)
    $Frame2.Add($Editor2)
    [Application]::Refresh()
}

$Label1 = [Terminal.Gui.Label]::new()
$Label1.Text = "Frame 1 Content"
$Label1.Height = 1
$Label1.Width = 20
$Frame1.Add($Label1)
$Frame1.Add($Welcome)

$Label2 = [Terminal.Gui.Label]::new()
$Label2.Text = "Frame 2 Content"
$Label2.Height = 1
$Label2.Width = 20
$Frame2.Add($Label2)



#$window.Add($Welcome)

[Application]::Top.Add($window)
[Application]::Run()
$window.SetFocus()
[Application]::ShutDown()

#end of file

#region functions

function protect-File {
    param (
        [string]$FilePath,
        [string]$Subject
    )
    $backupFileBaseName = (Get-Item $FilePath).BaseName.ToString()
    $backupTime = Get-Date -Format FileDateTime
    $backupFileName = $backupFileBaseName + "_" + $backupTime + ".bak"

    try {
        $cert = Get-CodeSigningCert -Subject $Subject
    }
    catch {
        Write-Warning "Failed to get certificate"
        return
    }

    if ($cert) {
        try {
            try {
                Copy-Item -Path $FilePath -Destination $backupFileName
            }
            catch {
                Write-Error "Failed to create backup file"
                return
            }
            Set-AuthenticodeSignature -FilePath $FilePath -Certificate $cert
            verify-signature -FilePath $FilePath

        }
        catch {
            Write-Error "Failed to sign file"
            return
        }
    }
}
#endregion

