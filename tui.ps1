#! "netcoreapp2.0"
#using System.Formats.Asn1
using namespace Terminal.Gui


# Load the Terminal.Gui.dll assembly.
if ($PSVersionTable.PSVersion -ge '7.2') {
    # Load the Terminal.Gui assembly via the 'Microsoft.PowerShell.ConsoleGuiTools'
    # module, by installing that module on demand.
    if (-not (Get-Module -ListAvailable Microsoft.PowerShell.ConsoleGuiTools)) {
        #Write-Verbose -Verbose "Installing module Microsoft.PowerShell.ConsoleGuiTools on demand, in the current user's scope."
        #Install-Module -Scope CurrentUser -ErrorAction Stop Microsoft.PowerShell.ConsoleGuiTools
    }
    # Terminal.Gui.dll is inside the module's folder.
    try {
        #Add-Type -LiteralPath (Join-Path (Get-Module -ListAvailable Microsoft.PowerShell.ConsoleGuiTools).ModuleBase Terminal.Gui.dll)
    }
    catch {
        throw
    }
}
else {
    # Windows PowerShell (or earlier PS Core versions)
    # Unfortunately, there's no easy way to gain access to Terminal.Gui.dll, and the
    # best option is to use an aux. NET SDK project as shown in https://stackoverflow.com/a/50004706/45375
    # The next command assumes that the steps there have been followed.
    try {
        #Add-Type -Path C:\Users\jdoe\.nuget-pwsh\packages-winps\terminal.gui\*\Terminal.Gui.dll
    }
    catch {
        throw
    }
}
#Install-Module TerminalGuiDesigner
#Import-Module TerminalGuiDesigner
#Import-Module -Name .\oconsysFnCollection.psm1
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
#$guiTools = (Get-Module TerminalGuiDesigner -List).ModuleBase
Set-Location -Path $scriptDir
#Add-Type -Path (Join-path $guiTools Terminal.Gui.dll)
[PSCustomObject]$global:ObjRegCnt = @{}
$global:CSVRegistrationContent = "Key", "Value", "Text" | Join-String -Separator ',' -DoubleQuote -OutputSuffix `n
$global:Textfields = @{}
[string]$global:SID = "get-ActiveDirectorySID"

function get-DetailledIPInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "IP to scan", ValueFromPipeline = $true)]
        [String]$IP
    )
    Write-Host "Scanning IP: $IP"
    Test-Connection $($IP) -Count 1 -TimeoutSeconds 1 -IPv4 | Select-Object -ExcludeProperty Source -First 1 -OutVariable IPInfo
    $IPInfo | Add-Member -MemberType NoteProperty -Name "DNS" -Value (Resolve-DnsName $IP -DnsOnly -Type PTR -QuickTimeout -ErrorAction SilentlyContinue).NameHost
    $IPInfo | Add-Member -MemberType NoteProperty -Name "TTL" -Value $IPInfo.reply.Options.Ttl
    $IPInfo | Add-Member -MemberType NoteProperty -Name "RTT" -Value $IPInfo.Latency
    return $IPInfo
}

function invoke-NetScan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0, HelpMessage = "Subnet to scan")]
        [String]$Subnet = "192.168.71.",
        [Parameter(Mandatory = $false, Position = 1, HelpMessage = "IP Range Start")]
        [Int16]$RangeStart = 1,
        [Parameter(Mandatory = $false, Position = 2, HelpMessage = "IP Range End")]
        [Int16]$RangeEnd = 20,
        [Parameter(Mandatory = $false, Position = 3, HelpMessage = "Quantity of IPs to scan")]
        [Int16]$Count = 1,
        [Parameter(Mandatory = $false, Position = 4, HelpMessage = "Systemtype to scan")]
        [string]$Type = "All",
        [Parameter(Mandatory = $false, Position = 5, HelpMessage = "Resolve DNS")]
        [string]$Resolve = "True"
    )

    <#     $printerManufacturers = @(
        [PSCustomObject]@{ Manufacturer = "Hewlett Packard"; MACPrefix = "00:1E:0B" },
        [PSCustomObject]@{ Manufacturer = "Hewlett Packard"; MACPrefix = "00:50:8B" },
        [PSCustomObject]@{ Manufacturer = "Canon"; MACPrefix = "00:1E:3B" },
        [PSCustomObject]@{ Manufacturer = "Canon"; MACPrefix = "00:26:2D" },
        [PSCustomObject]@{ Manufacturer = "Epson"; MACPrefix = "00:01:38" },
        [PSCustomObject]@{ Manufacturer = "Epson"; MACPrefix = "00:26:AB" },
        [PSCustomObject]@{ Manufacturer = "Brother"; MACPrefix = "00:80:77" },
        [PSCustomObject]@{ Manufacturer = "Brother"; MACPrefix = "00:1B:A4" },
        [PSCustomObject]@{ Manufacturer = "Lexmark"; MACPrefix = "00:21:B7" },
        [PSCustomObject]@{ Manufacturer = "Lexmark"; MACPrefix = "00:21:80" }
    ) #>
    Write-Host "Scanning Subnet: $Subnet for $Type Hosts, DNS Resolution: $Resolve"
    #$RangeStart..$RangeEnd | ForEach-Object { $Subnet + $_ }
    $SubnetIPs = $RangeStart..$RangeEnd | ForEach-Object { $Subnet + $_ }
    $ScanJob = $SubnetIPs | ForEach-Object -ThrottleLimit 30 -TimeoutSeconds 60 -Parallel {
        Write-Host "Scanning: $_"
        Test-Connection "$($_)" -Count 1 -TimeoutSeconds 1 -IPv4 | Select-Object -ExcludeProperty Source -First 1
    }
    $LiveIPs = $ScanJob | Where-Object { $_.Status -eq "Success" }
    switch ($Type) {
        "All" {
            $IPs = $LiveIPs
        }
        "Windows" {
            $IPs = $LiveIPs | Where-Object { (($_[0].reply.Options.TTL -ge 65) -and ($_[0].reply.Options.TTL -le 128)) }
        }
        "Linux" {
            $IPs = $LiveIPs | Where-Object { (($_[0].reply.Options.TTL -ge 32) -and ($_[0].reply.Options.TTL -le 64)) }
        }
    }

    $Return = $IPs | ForEach-Object {
        $IP = $_[0].Address
        if ($Resolve -eq "True") {
            Write-Host "DNS Resolution for $IP`: " -NoNewline -ForegroundColor Yellow
            if (($Resolved = Resolve-DnsName $IP -DnsOnly -Type PTR -QuickTimeout -ErrorAction SilentlyContinue).NameHost) {
                $DNS = $Resolved.NameHost | Join-String -Separator ','
                Write-Host "$DNS" -ForegroundColor Green
            }
            else {
                $DNS = "failed-resolve"
                Write-Host "$DNS" -ForegroundColor Red
            }
        }
        else {
            $DNS = ""
        }
        [Int]$TTL = $_[0].reply.Options.Ttl
        $RTT = $_[0].Latency
        if (($TTL -ge 65) -and ($TTL -le 128)) {
            $OS = "Windows"
        }
        elseif ($TTL -le 64) {
            $OS = "Linux"
        }
        else {
            $OS = "Unknown"
        }
        $IP, "$DNS", $TTL, $RTT, $OS | Join-String -Separator ',' | ConvertFrom-Csv -Header IP, DNS, TTL, RTT, OS -Delimiter ','
    }
    return $Return
}

function get-ActiveDirectorySID {
    (Get-ADDomain).DomainSID.Value | Out-String -NoNewline
}

function Update-StatusBar {
    param (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Text to display in StatusBar")]
        [String]$Text
    )
    $StatusBar.Items[3].Title = $Text
    $StatusBar.SetChildNeedsDisplay()
    [Terminal.Gui.Application]::Refresh()
}

function out-Mail {
    param(
        [Parameter(Mandatory = $false, Position = 0, HelpMessage = "Mail Content")]
        [string]$MailContent = "Mail Body"
    )
    Install-Module -Name Mailozaurr -AllowPrerelease -AllowClobber
    $RegistrationFile = ".\oconsysRegistration.csv"
    $MailContent = Get-Content $RegistrationFile -Raw
    $Subject = "oconsys AD Registration"
    $smtp_host = "mailcow.infraspread.net"
    $smtp_port = "587"
    $smtp_user = "share@infraspread.net"
    $smtp_pass = "#Share@Infra2024!"
    $from = "matrix@infraspread.net"
    $MailTo = "andreas.luemkemann@gmail.com"
    # this is simple replacement (drag & drop to Send-MailMessage)
    $secureString = ConvertTo-SecureString -AsPlainText -Force -String $smtp_pass
    $encryptedString = ConvertFrom-SecureString -SecureString $secureString
    $secureString = ConvertTo-SecureString -String $encryptedString
    $MailCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "$smtp_user", $secureString
    Update-StatusBar -Text "Sending Registration Mail to $MailTo"
    Send-EmailMessage -To $MailTo -Subject $Subject -Text $MailContent -SmtpServer $smtp_host -From $from -Priority High -Credential $MailCredentials -UseSsl -Port $smtp_port -Verbose -InformationVariable global:mailInfo
    Update-StatusBar -Text "Mail to $MailTo ($mailInfo).Message ..."
}


Function ConvertTo-DataTable {
    [cmdletbinding()]
    [OutputType('System.Data.DataTable')]
    [alias('alias')]
    Param(
        [Parameter(
            Mandatory,
            Position = 0,
            ValueFromPipeline
        )]
        [ValidateNotNullOrEmpty()]
        [object]$InputObject
    )

    Begin {
        Write-Verbose "[$((Get-Date).TimeOfDay) BEGIN  ] Starting $($MyInvocation.MyCommand)"
        Write-Verbose "[$((Get-Date).TimeOfDay) BEGIN  ] Running under PowerShell version $($PSVersionTable.PSVersion)"
        $data = [System.Collections.Generic.List[object]]::New()
        $Table = [System.Data.DataTable]::New('PSData')
    } #begin

    Process {
        $Data.Add($InputObject)
    } #process

    End {
        Write-Verbose "[$((Get-Date).TimeOfDay) END    ] Building a table of $($data.count) items"
        #define columns
        foreach ($item in $data[0].PSObject.Properties) {
            Write-Verbose "[$((Get-Date).TimeOfDay) END    ] Defining column $($item.name)"
            [void]$table.Columns.Add($item.Name, $item.TypeNameOfValue)
        }
        #add rows
        for ($i = 0; $i -lt $Data.count; $i++) {
            $row = $table.NewRow()
            foreach ($item in $Data[$i].PSObject.Properties) {
                $row.Item($item.name) = $item.Value
            }
            [void]$table.Rows.Add($row)
        }
        #This is a trick to return the table object
        #as the output and not the rows
        , $table
        Write-Verbose "[$((Get-Date).TimeOfDay) END    ] Ending $($MyInvocation.MyCommand)"
    } #end

} #close ConvertTo-DataTable

function get-WlanPassword {
    (netsh wlan show profiles) | Select-String "\:(.+)$" | ForEach-Object { $name = $_.Matches.Groups[1].Value.Trim(); $_ } | ForEach-Object { (netsh wlan show profile name="$name" key=clear) } | Select-String "Schlüsselinhalt\W+\:(.+)$|Key Content\W+\:(.+)$" | ForEach-Object { $pass = $_.Matches.Groups[1].Value.Trim(); if (-not $pass) {
            $pass = $_.Matches.Groups[2].Value.Trim()
        }; $_ } | ForEach-Object { [PSCustomObject]@{ PROFILE_NAME = $name; PASSWORD = $pass } }
}

function new-TextboxAndLabel {
    param (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Key")]
        [String]$Key,
        [Parameter(Mandatory = $true, Position = 1, HelpMessage = "Value")]
        [String]$Value,
        [Parameter(Mandatory = $true, Position = 2, HelpMessage = "Text")]
        [String]$Text,
        [Parameter(Mandatory = $true, Position = 3, HelpMessage = "Label Text")]
        [String]$LabelText,
        [Parameter(Mandatory = $false, Position = 4, HelpMessage = "Label Width")]
        [Terminal.Gui.Dim]$LabelWidth = ($LabelText.Length + 1),
        [Parameter(Mandatory = $true, Position = 5, HelpMessage = "Textbox Text")]
        [String]$TextboxText,
        [Parameter(Mandatory = $false, Position = 6, HelpMessage = "Textbox Width")]
        [Terminal.Gui.Dim]$TextboxWidth = ($TextboxText.Length + 1),
        [Parameter(Mandatory = $true, Position = 7, HelpMessage = "Name of Frame")]
        [object]$Frame,
        [Parameter(Mandatory = $false, Position = 8, HelpMessage = "Line in Frame")]
        [Terminal.Gui.Pos]$X = 0,
        [Parameter(Mandatory = $false, Position = 9, HelpMessage = "Row in Frame")]
        [Terminal.Gui.Pos]$Y = 0
    )


    $TextboxAndLabel = [PSCustomObject]@{
        TextBoxAndLabelName = $LabelText.Replace(' ', '')
        TotalWidth          = $LabelText.Length + 1 + $TextboxWidth.Length + 1
    }

    $Label = [Terminal.Gui.Label]@{
        Text   = $LabelText
        Width  = $LabelWidth
        Height = 1
        X      = [Terminal.Gui.Pos]::Left($Frame)
        Y      = $Y
    }
    $Frame.Add($Label)

    $Textfield = [Terminal.Gui.Textfield]@{
        Text   = $Text
        Width  = $TextboxWidth
        Height = 1
        X      = [Terminal.Gui.Pos]::Right($Label) + 1
        Y      = $Y
    }

    $Frame.Add($Textfield)
    $global:Textfields[$TextboxAndLabel.TextBoxAndLabelName] = $Textfield
    return $Textfield
}

function new-Button {
    param (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Button Name")]
        [String]$ButtonName,
        [Parameter(Mandatory = $false, Position = 1, HelpMessage = "Button Width")]
        [Terminal.Gui.Dim]$ButtonWidth = ($ButtonName.Length + 2),
        [Parameter(Mandatory = $true, Position = 4, HelpMessage = "Name of Frame")]
        [object]$Frame,
        [Parameter(Mandatory = $false, Position = 5, HelpMessage = "Line in Frame")]
        [Terminal.Gui.Pos]$X = 0,
        [Parameter(Mandatory = $false, Position = 6, HelpMessage = "Row in Frame")]
        [Terminal.Gui.Pos]$Y = 0
    )

    $Button = [Terminal.Gui.Button]@{
        Text   = $ButtonName
        Width  = $Width
        Height = 1
        X      = [Terminal.Gui.Pos]::Left($Frame)
        Y      = $Y
    }
    $Frame.Add($Button)
}

#region Get-RemotePrinterList
function Get-RemotePrinterListTUI {
    param (
        [Parameter(Mandatory = $false)]
        [string]$Printserver = 'print.infraspread.net',
        [Parameter(Mandatory = $false)]
        [string]$SSHUser = 'administrator@infraspread.net',
        [Parameter(Mandatory = $false)]
        [string]$KeyFilePath,
        [Parameter(Mandatory = $false)]
        [string]$prepend = '__gg_',
        [Parameter(Mandatory = $false)]
        [string]$append = '__',
        [Parameter(Mandatory = $false)]
        [string]$ReplaceBlanks = '_'
    )
    $excludePrinter = @('Fax')
    $excludePrinterPort = @('Portprompt:')

    $null = $remoteSessionPrintserver = New-PSSession -HostName $Printserver -SSHTransport -UserName $SSHUser -KeyFilePath $KeyFilePath
    $null = Import-PSSession -Session $remoteSessionPrintserver -Module PrintManagement -Prefix Remote -AllowClobber -FormatTypeName *
    $null = $printersAll = Get-RemotePrinter

    $printersFiltered = $printersAll | Where-Object { ($_.Name -NotIn $excludePrinter) -and ($_.PortName -NotIn $excludePrinterPort) }
    $printerlist = foreach ($printer in $printersFiltered) {
        [Ordered]@{
            Name        = $printer.name
            Printserver = $printserver
            DriverName  = $printer.drivername
            PortName    = $printer.portname
            Location    = $printer.location
            Shared      = $printer.shared
            ShareName   = $printer.sharename
            Group       = $prepend + $printer.ShareName + $append
        }
    }
    Get-PSSession | Remove-PSSession
    return $printerlist
}
#endregion

#region Get-RemotePrinterListTUILX
function Get-RemotePrinterListTUILX {
    param (
        [Parameter(Mandatory = $false)]
        [string]$Printserver = 'print.infraspread.net',
        [Parameter(Mandatory = $false)]
        [string]$SSHUser = 'administrator@infraspread.net',
        [Parameter(Mandatory = $false)]
        [string]$KeyFilePath,
        [Parameter(Mandatory = $false)]
        [string]$prepend = '__gg_',
        [Parameter(Mandatory = $false)]
        [string]$append = '__',
        [Parameter(Mandatory = $false)]
        [string]$ReplaceBlanks = '_'
    )
    $excludePrinter = @('Fax')
    $excludePrinterPort = @('Portprompt:')

    $null = $remoteSessionPrintserver = New-PSSession -HostName $Printserver -SSHTransport -UserName $SSHUser -KeyFilePath $KeyFilePath

    #$null=Import-PSSession -Session $remoteSessionPrintserver -Module PrintManagement -Prefix Remote -AllowClobber -FormatTypeName *
    #$null=$printersAll = Get-RemotePrinter
    $null = $printersAll = Invoke-Command -Session $remoteSessionPrintserver -ScriptBlock {
        $printerlistremote = Get-Printer | Select-Object -Property Name, DriverName, PortName, Location, Shared, ShareName | Sort-Object -Property Name | ConvertTo-Csv -Delimiter ',' -NoTypeInformation
        return $printerlistremote
    }



    $printersFiltered = $printersAll | Where-Object { ($_.Name -NotIn $excludePrinter) -and ($_.PortName -NotIn $excludePrinterPort) }
    $printerlist = foreach ($printer in $printersFiltered) {
        [Ordered]@{
            Name        = $printer.name
            Printserver = $printserver
            DriverName  = $printer.drivername
            PortName    = $printer.portname
            Location    = $printer.location
            Shared      = $printer.shared
            ShareName   = $printer.sharename
            Group       = $prepend + $printer.ShareName + $append
        }
    }
    Get-PSSession | Remove-PSSession
    return $printerlist
}
#endregion


[Terminal.Gui.Application]::Init()
[Terminal.Gui.Application]::QuitKey = 27 # ESC


$window = [Terminal.Gui.Window]@{
    Title       = 'oconsysTUI'
    #ColorScheme = [Terminal.Gui.ColorScheme]::('Colors.TopLevel')
    ColorScheme = [Terminal.Gui.ColorScheme]::('Colors.Base')
}


#region status bar
$StatusBar = [Terminal.Gui.StatusBar]::New(
    @(
        [Terminal.Gui.StatusItem]::New('Unknown', $(Get-Date -Format g), {}),
        [Terminal.Gui.StatusItem]::New('Unknown', 'ESC to quit or cancel', {}),
        [Terminal.Gui.StatusItem]::New('Unknown', "v0.9", {}),
        [Terminal.Gui.StatusItem]::New('Unknown', 'Ready', {})
    )
)
#endregion

#region menu bar

$MenuItem0 = [Terminal.Gui.MenuItem]::New('_Quit', '', { [Terminal.Gui.Application]::RequestStop() })
$MenuItemMain = [Terminal.Gui.MenuItem]::New('_Main', '', { $ButtonMain.OnClicked() })
$MenuBarItem0 = [Terminal.Gui.MenuBarItem]::New('_Options', @($MenuItemMain, $MenuItem0))
$MenuItem0.ShortCut = [int][Terminal.Gui.Key]'CtrlMask' -bor [int][char]'Q'
$MenuItemNetScan = [Terminal.Gui.MenuItem]::New('_NetScan', '', { $ButtonNetwork.OnClicked() })
$MenuBarItemNetwork = [Terminal.Gui.MenuBarItem]::New('_Network', @($MenuItemNetScan))
$MenuItemRegistration = [Terminal.Gui.MenuItem]::New('_Register', '', { $ButtonPrinterDeploymentRegistration.OnClicked() })
$MenuBarItemRegistration = [Terminal.Gui.MenuBarItem]::New('_Register', @($MenuItemRegistration))
$global:MenuBar = [Terminal.Gui.MenuBar]::New(@($MenuBarItem0, $MenuBarItemNetwork, $MenuBarItemRegistration))
$MenuBar.ColorScheme = [Terminal.Gui.ColorScheme]::('Colors.Menu')
$Window.Add($global:MenuBar)

function Add-MenubarEntry {
    param (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Menu Item")]
        [object]$MenuItem,
        [Parameter(Mandatory = $true, Position = 1, HelpMessage = "Menu Bar Item")]
        [object]$MenuBarItem,
        [Parameter(Mandatory = $false, Position = 2, HelpMessage = "Menu Item Action")]
        [scriptblock]$ItemAction
    )
    #$MenuItem0 = [Terminal.Gui.MenuItem]::New('_Quit', '', { [Terminal.Gui.Application]::RequestStop() })
    #$MenuItem0.ShortCut = [int][Terminal.Gui.Key]'CtrlMask' -bor [int][char]'Q'
    #$MenuBarItem0 = [Terminal.Gui.MenuBarItem]::New('_Options', @($MenuItem0))
    #$MenuBar = [Terminal.Gui.MenuBar]::New(@($MenuBarItem0))
    $newItem = [Terminal.Gui.MenuItem]::New("_$($MenuItem)", '', $ItemAction)
    $newMenuBarItem = [Terminal.Gui.MenuBarItem]::New("_$($MenuBarItem)", @($newItem))
    $global:newMenuBar = [Terminal.Gui.MenuBar]::Add(@($newMenuBarItem))
    $Window.Add($global:newMenuBar)
}

#endregion

#region main window

<# $Panel1 = [Terminal.Gui.Panel]@{
    Title = 'Panel 1'
}

$Panel2 = [Terminal.Gui.Panel]@{
    Title = 'Panel 2'
}

$Panelview = [Terminal.Gui.PanelView]@{
    x      = [POS]::Bottom($MenuBar)
    Width  = [Terminal.Gui.Dim]::Fill()
    Height = [Terminal.Gui.Dim]::Fill()
}
$Panelview.Add($Panel1)
$Panelview.Add($Panel2)
 #>

#region Frames
$LeftFrame = [Terminal.Gui.FrameView]@{
    Title  = 'Scan Settings'
    Width  = 40
    Height = 30
    X      = 0
    Y      = [Terminal.Gui.Pos]::Bottom($MenuBar)
}

$LeftBottomFrame = [Terminal.Gui.FrameView]@{
    Title  = 'Actions'
    Width  = 40
    Height = [Terminal.Gui.Dim]::Fill()
    Y      = [Terminal.Gui.Pos]::Bottom($LeftFrame)
}

$RightFrame = [Terminal.Gui.FrameView]@{
    Title  = 'Scan Results'
    Width  = [Terminal.Gui.Dim]::Fill()
    Height = [Terminal.Gui.Dim]::Fill()
    X      = [Terminal.Gui.Pos]::Right($LeftFrame)
    Y      = [Terminal.Gui.Pos]::Bottom($MenuBar)
}

$global:PrinterDeploymentRegistrationFrameTop = [Terminal.Gui.FrameView]@{
    Title  = 'Printer Deployment Registration'
    Width  = [Terminal.Gui.Dim]::Fill()
    Height = 5
    Y      = [Terminal.Gui.Pos]::Bottom($MenuBar)
}

$global:PrinterDeploymentFrameRegistrationActions = [Terminal.Gui.FrameView]@{
    Title  = 'Printer Deployment Registration'
    Width  = [Terminal.Gui.Dim]::Fill()
    Height = 5
    Y      = [Terminal.Gui.Pos]::Bottom($PrinterDeploymentRegistrationFrameActions)
}

$global:PrinterDeploymentRegistrationFrameLeft = [Terminal.Gui.FrameView]@{
    Title  = 'Printer Deployment Registration'
    Width  = [Terminal.Gui.Dim]::Percent(35)
    Height = [Terminal.Gui.Dim]::Fill()
    Y      = [Terminal.Gui.Pos]::Bottom($PrinterDeploymentRegistrationFrameTop)
}

$global:PrinterDeploymentRegistrationFrameRight = [Terminal.Gui.FrameView]@{
    Title  = 'Registration'
    Width  = [Terminal.Gui.Dim]::Fill()
    Height = [Terminal.Gui.Dim]::Fill()
    Y      = [Terminal.Gui.Pos]::Bottom($PrinterDeploymentRegistrationFrameTop)
    X      = [Terminal.Gui.Pos]::Right($PrinterDeploymentRegistrationFrameLeft)
}

$window.Add($LeftFrame)
$window.Add($LeftBottomFrame)
$window.Add($RightFrame)
#$RightFrame.Add($Panelview)
#endregion Frames

#Region Scan Settings
$LabelSubnet = [Terminal.Gui.Label]@{
    Text   = 'Subnet'
    Width  = 12
    Height = 1
    X      = [Terminal.Gui.Pos]::Left($LeftFrame)
    Y      = [Terminal.Gui.Pos]::Top(($LeftFrame))
}
$LeftFrame.Add($LabelSubnet)

$TextfieldSubnet = [Terminal.Gui.Textfield]@{
    Text   = '192.168.2.'
    Width  = 12
    Height = 1
    X      = [Terminal.Gui.Pos]::Left($LeftFrame) + 15
    Y      = $LabelSubnet.Y
}
$LeftFrame.Add($TextfieldSubnet)

$LabelRangeStart = [Terminal.Gui.Label]@{
    Text   = 'Range Start'
    Width  = 12
    Height = 1
    X      = [Terminal.Gui.Pos]::Left($LeftFrame)
    Y      = $LabelSubnet.Y + 1

}
$LeftFrame.Add($LabelRangeStart)

$TextfieldRangeStart = [Terminal.Gui.Textfield]@{
    Text   = '1'
    Width  = 3
    Height = 1
    X      = [Terminal.Gui.Pos]::Left($LeftFrame) + 15
    Y      = $LabelRangeStart.Y
}
$LeftFrame.Add($TextfieldRangeStart)

$LabelRangeEnd = [Terminal.Gui.Label]@{
    Text   = 'Range End'
    Width  = 12
    Height = 1
    X      = [Terminal.Gui.Pos]::Left($LeftFrame)
    Y      = $LabelRangeStart.Y + 1
}
$LeftFrame.Add($LabelRangeEnd)

$TextfieldRangeEnd = [Terminal.Gui.Textfield]@{
    Text   = '254'
    Width  = 3
    Height = 1
    X      = [Terminal.Gui.Pos]::Left($LeftFrame) + 15
    Y      = $LabelRangeEnd.Y
}
$LeftFrame.Add($TextfieldRangeEnd)
#endregion Scan Settings

#region Extra Buttons
$ButtonScan = [Terminal.Gui.Button]@{
    Text     = 'Start Scan'
    AutoSize = $true
    X        = [Terminal.Gui.Pos]::Left($LeftBottomFrame)
    #    Y        = [Terminal.Gui.Pos]::Top($LeftBottomFrame)
    Y        = 0
}

$ButtonIPv4Info = [Terminal.Gui.Button]@{
    Text     = 'IPv4 Info'
    AutoSize = $true
    X        = [Terminal.Gui.Pos]::Left($LeftBottomFrame)
    Y        = $ButtonScan.Y + 1
}

$ButtonWlanPasswords = [Terminal.Gui.Button]@{
    Text     = 'WLAN Passwords'
    AutoSize = $true
    X        = [Terminal.Gui.Pos]::Left($LeftBottomFrame)
    Y        = $ButtonIPv4Info.Y + 1
}

$ButtonPrinters = [Terminal.Gui.Button]@{
    Text     = 'Local Printers'
    AutoSize = $true
    X        = [Terminal.Gui.Pos]::Left($LeftBottomFrame)
    Y        = $ButtonWlanPasswords.Y + 1
}
$ButtonShowLastError = [Terminal.Gui.Button]@{
    Text  = "Show Last Error"
    Width = 23
    X     = [Terminal.Gui.Pos]::Right($PrinterDeploymentRegistrationFrameTop) - 23
    Y     = [Terminal.Gui.Pos]::Top($PrinterDeploymentRegistrationFrameTop)
}
#endregion Extra Button

#region Printer Deployment
$ButtonPrinterDeployment = [Terminal.Gui.Button]@{
    Text     = 'Printer Deployment'
    AutoSize = $true
    X        = [Terminal.Gui.Pos]::Left($LeftBottomFrame)
    Y        = $ButtonPrinters.Y + 1
}

$ButtonPrinterDeploymentRegistration = [Terminal.Gui.Button]@{
    Text     = 'Registration'
    AutoSize = $true
    X        = [Terminal.Gui.Pos]::Left($LeftBottomFrame)
    Y        = $ButtonPrinterDeployment.Y + 1
}
#endregion Printer Deployment

#region ButtonClicks
$ButtonScan.add_Clicked({
        $global:ScanStart = New-Object -TypeName System.Text.StringBuilder
        $null = $ScanStart.Append($TextfieldSubnet.Text)
        $null = $ScanStart.Append($TextfieldRangeStart.Text)
        $ScanStart = $ScanStart.ToString()

        $global:ScanEnd = New-Object -TypeName System.Text.StringBuilder
        $null = $ScanEnd.Append($TextfieldSubnet.Text)
        $null = $ScanEnd.Append($TextfieldRangeEnd.Text)
        $ScanEnd = $ScanEnd.ToString()

        #[Terminal.Gui.MessageBox]::Query("Button: Scan Clicked!", "Scan from $ScanStart to $ScanEnd")
        $LabelIPTextVar = "Scanning IPs from $ScanStart to $ScanEnd"
        Update-StatusBar -Text $LabelIPTextVar
        #$LabelIP.Text = $LabelIPTextVar
        #$LabelIP.Update()
        $global:ScanResults = invoke-NetScan -Subnet "192.168.2." -RangeStart 1 -RangeEnd 88 -Count 1 -Resolve "false" -Type "All"
        $ScanResults | Out-GridView
        $TableView.Table = $ScanResults | ConvertTo-DataTable
        $TableView.Update()
    })

$ButtonIPv4Info.add_Clicked({
        Update-StatusBar -Text "IPv4 Info"
        $global:IPv4Info = Get-NetIPConfiguration | Select-Object -Property Interface*, IPv4Ad*
        $TableView.Table = $IPv4Info | ConvertTo-DataTable
        $TableView.Update()
    })

$ButtonWlanPasswords.add_Clicked({
        Update-StatusBar -Text "WLAN Passwords"
        $global:WlanPasswords = Get-WlanPassword
        $TableView.Table = $WlanPasswords | ConvertTo-DataTable
        $TableView.Update()
    })

$ButtonPrinters.add_Clicked({
        Update-StatusBar -Text "Printers"
        $global:Printers = Get-Printer
        $TableView.Table = $Printers | ConvertTo-DataTable
        $TableView.Update()
    })

$ButtonPrinterDeployment.add_Clicked({
        $LeftFrame.RemoveAll()
        $LeftBottomFrame.RemoveAll()
        $RightFrame.RemoveAll()

        $global:PrinterDeploymentFrameTop = [Terminal.Gui.FrameView]@{
            Title  = 'Printer Deployment Actions'
            Width  = [Terminal.Gui.Dim]::Fill()
            Height = 5
            Y      = [Terminal.Gui.Pos]::Bottom($MenuBar)
        }
        $window.Add($PrinterDeploymentFrameTop)

        $global:PrinterDeploymentFrameActions = [Terminal.Gui.FrameView]@{
            Title  = 'Printer Deployment Configuration'
            Width  = [Terminal.Gui.Dim]::Fill()
            Height = 5
            Y      = [Terminal.Gui.Pos]::Bottom($PrinterDeploymentFrameActions)

        }
        #$window.Add($PrinterDeploymentFrameActions)

        $global:PrinterDeploymentFrameLeft = [Terminal.Gui.FrameView]@{
            Title  = 'Printer Deployment'
            Width  = [Terminal.Gui.Dim]::Percent(25)
            Height = [Terminal.Gui.Dim]::Fill()
            Y      = [Terminal.Gui.Pos]::Bottom($PrinterDeploymentFrameTop)

        }
        $window.Add($PrinterDeploymentFrameLeft)

        $global:PrinterDeploymentFrameRight = [Terminal.Gui.FrameView]@{
            Title  = 'Results'
            Width  = [Terminal.Gui.Dim]::Fill()
            Height = [Terminal.Gui.Dim]::Fill()
            Y      = [Terminal.Gui.Pos]::Bottom($PrinterDeploymentFrameTop)
            X      = [Terminal.Gui.Pos]::Right($PrinterDeploymentFrameLeft)
        }
        $window.Add($PrinterDeploymentFrameRight)

        $global:TableViewPrinterDeployment = [Terminal.Gui.TableView]@{
            Width         = [Terminal.Gui.Dim]::Fill()
            Height        = [Terminal.Gui.Dim]::Fill()
            AutoSize      = $true
            TabStop       = $False
            MultiSelect   = $False
            FullRowSelect = $True
            TextAlignment = 'left'
        }
        #Keep table headers always in view
        $TableViewPrinterDeployment.Style.AlwaysShowHeaders = $True
        $TableViewPrinterDeployment.Style.ShowHorizontalHeaderOverline = $False
        $TableViewPrinterDeployment.Style.ShowHorizontalHeaderUnderline = $True
        $TableViewPrinterDeployment.Style.ShowVerticalHeaderLines = $False
        $PrinterDeploymentFrameRight.Add($TableViewPrinterDeployment)

        Update-StatusBar -Text "Printer Deployment"

        $ButtonInit = [Terminal.Gui.Button]@{
            Text  = "Init"
            Width = 10
            X     = [Terminal.Gui.Pos]::Left($PrinterDeploymentFrameTop)
            Y     = [Terminal.Gui.Pos]::Top($PrinterDeploymentFrameTop)
        }
        $PrinterDeploymentFrameTop.Add($ButtonInit)
        $ButtonInit.add_Clicked({
                $ConfigFile = ".\oconsysConfiguration.csv"
                $global:Config = Import-Csv $ConfigFile
                $i = 0
                $global:Config.ForEach({
                        $i++
                        new-TextboxAndLabel -LabelText $_.Key -TextboxText $_.Value -Frame $PrinterDeploymentFrameLeft -Y $i
                    })
            })

        $ButtonGetRemotePrinterList = [Terminal.Gui.Button]@{
            Text   = 'Get Remote Printer List'
            Width  = 29
            X      = [Terminal.Gui.Pos]::Right($ButtonInit)
            Y      = $ButtonInit.Y
            Hotkey = "null"
        }
        $PrinterDeploymentFrameTop.Add($ButtonGetRemotePrinterList)
        $ButtonGetRemotePrinterList.add_Clicked({
                $Printerlist = Get-RemotePrinterListTUI -KeyFilePath "C:\Users\andreas\.ssh\id_ed25519"
                $PrinterlistCSV = $Printerlist | ConvertTo-Csv -Delimiter ',' -NoTypeInformation
                $TableViewPrinterDeployment.Table = $PrinterlistCSV | ConvertFrom-Csv | ConvertTo-DataTable
                $TableViewPrinterDeployment.Update()
            })
    })

$ButtonShowLastError.add_Clicked({
        if ($Error[0]) {
            $errorMessage = $Error[0].ToString()
            [Terminal.Gui.MessageBox]::Query("Last Error", $errorMessage)
        }
        else {
            [Terminal.Gui.MessageBox]::Query("No Errors", "No errors have occurred.")
        }
    })



#region Main Table View
$TableView = [Terminal.Gui.TableView]@{
    Width         = [Terminal.Gui.Dim]::Fill()
    Height        = [Terminal.Gui.Dim]::Fill()
    AutoSize      = $true
    TabStop       = $False
    MultiSelect   = $False
    FullRowSelect = $True
    TextAlignment = 'left'
}
#Keep table headers always in view
$TableView.Style.AlwaysShowHeaders = $True
$TableView.Style.ShowHorizontalHeaderOverline = $False
$TableView.Style.ShowHorizontalHeaderUnderline = $True
$TableView.Style.ShowVerticalHeaderLines = $False


#endregion

#region Printer Deployment Registration
$ButtonLoadRegistration = [Terminal.Gui.Button]@{
    Text  = "Load Registration"
    Width = 23
    X     = [Terminal.Gui.Pos]::Left($PrinterDeploymentRegistrationFrameTop)
    Y     = [Terminal.Gui.Pos]::Top($PrinterDeploymentRegistrationFrameTop)
}
$ButtonSaveRegistration = [Terminal.Gui.Button]@{
    Text   = 'Save Registration'
    Width  = 23
    X      = [Terminal.Gui.Pos]::Right($ButtonLoadRegistration)
    Y      = $ButtonLoadRegistration.Y
    Hotkey = "null"
}
$ButtonDomainSID = [Terminal.Gui.Button]@{
    Text = "Get Domain SID"
    X    = [Terminal.Gui.Pos]::Right($ButtonSaveRegistration) + 1
    Y    = $ButtonSaveRegistration.Y
}
$ButtonSendRegistration = [Terminal.Gui.Button]@{
    Text   = 'Send Registration'
    Width  = 23
    X      = [Terminal.Gui.Pos]::Right($ButtonDomainSID)
    Y      = $ButtonLoadRegistration.Y
    Hotkey = "null"
}


$global:TableViewPrinterDeploymentRegistration = [Terminal.Gui.TableView]@{
    Width         = [Terminal.Gui.Dim]::Fill()
    Height        = [Terminal.Gui.Dim]::Fill()
    AutoSize      = $true
    TabStop       = $False
    MultiSelect   = $False
    FullRowSelect = $True
    TextAlignment = 'left'

}
#Keep table headers always in view
$global:TableViewPrinterDeploymentRegistration.Style.AlwaysShowHeaders = $True
$global:TableViewPrinterDeploymentRegistration.Style.ShowHorizontalHeaderOverline = $False
$global:TableViewPrinterDeploymentRegistration.Style.ShowHorizontalHeaderUnderline = $True
$global:TableViewPrinterDeploymentRegistration.Style.ShowVerticalHeaderLines = $False


$ButtonPrinterDeploymentRegistration.add_Clicked({
        $LeftFrame.RemoveAll()
        $LeftBottomFrame.RemoveAll()
        $RightFrame.RemoveAll()
        $window.Add($PrinterDeploymentRegistrationFrameTop)
        $window.Add($PrinterDeploymentRegistrationFrameLeft)
        $window.Add($PrinterDeploymentRegistrationFrameRight)
        #$window.Add($PrinterDeploymentFrameActions)
        Update-StatusBar -Text "Printer Deployment Registration"
        $PrinterDeploymentRegistrationFrameTop.Add($ButtonLoadRegistration)
        $PrinterDeploymentRegistrationFrameTop.Add($ButtonSaveRegistration)
        $PrinterDeploymentRegistrationFrameTop.Add($ButtonDomainSID)
        $PrinterDeploymentRegistrationFrameTop.Add($ButtonSendRegistration)
        $PrinterDeploymentRegistrationFrameTop.Add($ButtonShowLastError)
        $ButtonLoadRegistration.OnClicked()
        $ButtonDomainSID.OnClicked()
    })

function new-LabelPlusTextfield {
    param (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Key")]
        [String]$Key,
        [Parameter(Mandatory = $true, Position = 1, HelpMessage = "Value")]
        [String]$Value,
        [Parameter(Mandatory = $true, Position = 2, HelpMessage = "Data")]
        [String]$Data,
        [Parameter(Mandatory = $false, Position = 3, HelpMessage = "Label Width")]
        [Terminal.Gui.Dim]$LabelWidth = ($Key.Length + 1),
        [Parameter(Mandatory = $false, Position = 4, HelpMessage = "Textbox Width")]
        [Terminal.Gui.Dim]$TextfieldWidth = ($Value.Length + 1),
        [Parameter(Mandatory = $true, Position = 5, HelpMessage = "Name of Frame")]
        [object]$Frame,
        [Parameter(Mandatory = $false, Position = 6, HelpMessage = "Line in Frame")]
        [Terminal.Gui.Pos]$X = 0,
        [Parameter(Mandatory = $false, Position = 7, HelpMessage = "Row in Frame")]
        [Terminal.Gui.Pos]$Y = 0
    )
    $LabelPlusTextfield = [PSCustomObject]@{
        LabelID        = $Key.Replace(' ', '')
        LabelText      = $Key
        LabelWidth     = $Key.Length + 1
        TextfieldID    = $Key.Replace(' ', '')
        TextfieldText  = $Value
        TextfieldData  = $Data
        TextfieldWidth = $Value.Length + 1
        Width          = $Key.Length + $Value.Length
    }

    $Label = [Terminal.Gui.Label]@{
        Id     = $LabelPlusTextfield.LabelID
        Text   = $LabelPlusTextfield.LabelText
        Width  = $LabelPlusTextfield.LabelWidth
        Height = 1
        X      = [Terminal.Gui.Pos]::Left($Frame)
        Y      = $Y
    }
    $Frame.Add($Label)

    $Textfield = [Terminal.Gui.Textfield]@{
        Id     = $LabelPlusTextfield.TextfieldID
        Text   = $LabelPlusTextfield.TextfieldText
        Data   = $LabelPlusTextfield.TextfieldData
        Width  = $LabelPlusTextfield.TextfieldWidth
        Height = 1
        X      = [Terminal.Gui.Pos]::Right($Label) + 1
        Y      = $Y
    }

    $Frame.Add($Textfield)
    $global:Textfields[$LabelPlusTextfield.TextfieldID.ToString()] = $Textfield
    return $Textfield
}

function LicenseCheck {
    param (
        [string]$LICSID,
        [string]$LICDOMAIN
    )
    [string]$LICSID = "S-1-5-21-3319716681-3808972711-940181925"
    [string]$LICDOMAIN = "infraspread.net"
    [string]$CheckDom = "$($LICSID).$($LICDOMAIN)"
    Write-Warning "Checking License for $CheckDom"
    $DNSLIC = $(Resolve-DnsName $CheckDom -Type TXT -Server 9.9.9.9)

    #$DNSLIC.Strings.split(';') | ConvertFrom-StringData -Delimiter '='
    $LicenseCount = ($DNSLIC.Strings.split(';') | ConvertFrom-StringData -Delimiter '=' | Where-Object -Property LIC)['LIC']
    $ValidTill = ($DNSLIC.Strings.split(';') | ConvertFrom-StringData -Delimiter '=' | Where-Object -Property Date)['DATE']
    Write-Warning "LicenseCount: $LicenseCount, ValidTill: $ValidTill"
    #Date String: 20190928 , convert to date: in short format, no time
    $Date = [datetime]::ParseExact($ValidTill, "yyyyMMdd", $null).ToShortDateString() -replace "/", "-"
    Write-Warning "Date: $Date"
    #create a new custom object, add the license count and the date
    #$LicenseCheck = @()
    $LicenseCheck = [PSCustomObject]@{
        LicenseCount = $LicenseCount
        ValidTill    = $Date
    }
    return $LicenseCheck
}

function Update-DomainSIDTextfield {
    param (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "SID")]
        [String]$SID
    )
    if ($global:Textfields.ContainsKey("DomainSID")) {
        Write-Warning "The hashtable contains the key 'DomainSID'."
    }
    else {
        Write-Warning "The hashtable does not contain the key 'DomainSID'."
    }
    $global:Textfields["DomainSID"].Text = $($SID)
    $LicenseInfo = LicenseCheck -LICSID $SID -LICDOMAIN "infraspread.net"
    [string]$date = $($LicenseInfo).ValidTill
    [string]$count = $($LicenseInfo).LicenseCount
    $global:Textfields["LicenseCount"].Text = $($count)
    $global:Textfields["ValidTill"].Text = $($date)
    Write-Warning "LicenseCount: $($count), ValidTill: $($date)"
    #$global:Textfields["LicenseCount"].Text = "slkdjfsldkfj"
    #$global:Textfields["ValidTill"].Text = "sldkjlkdfjslkjfdjlfgjsldjflsj"

    $global:Textfields["LicenseCount"].Update()
    $global:Textfields["ValidTill"].Update()
    $global:Textfields["DomainSID"].Update()
}


$ButtonLoadRegistration.add_Clicked({
        $RegistrationFile = ".\oconsysRegistration.csv"
        $global:Registration = Import-Csv $RegistrationFile
        $i = 0
        $global:Registration.ForEach({
                $i++
                $global:CSVRegistrationContent += $_.Key, $_.Value, $_.Text | Join-String -Separator ',' -DoubleQuote -OutputSuffix `n
                Write-Warning "Key: $($_.Key), Value: $($_.Value), Text: $($_.Text)"
                new-LabelPlusTextfield -Key $_.Key -Value $_.Value -Data $_.Text -Frame $PrinterDeploymentRegistrationFrameLeft -Y $i
            })
        $global:TableViewPrinterDeploymentRegistration.Table = $Registration | ConvertTo-DataTable
        $PrinterDeploymentRegistrationFrameRight.Add($global:TableViewPrinterDeploymentRegistration)
        $global:TableViewPrinterDeploymentRegistration.Update()

    })


$ButtonSaveRegistration.add_Clicked({
        $RegistrationFile = ".\oconsysRegistration.csv"
        $global:Registration = $global:Textfields.GetEnumerator() | ForEach-Object {
            [PSCustomObject]@{
                Key   = $_.Key
                Value = $_.Value.Text
                Text  = $_.Value.Data
            }
        }
        $global:Registration | Export-Csv -Path $RegistrationFile -NoTypeInformation -Force
    })



$ButtonDomainSID.add_Clicked({
        $DomainSID = get-ActiveDirectorySID
        Update-StatusBar -Text "Domain SID: $($DomainSID)"
        Update-DomainSIDTextfield -SID $DomainSID
    })

$ButtonSendRegistration.add_Clicked({
        out-Mail
    })

#endregion Printer Deployment Registration


$ButtonMain = [Terminal.Gui.Button]@{
    Text  = "Main"
    Width = 23
    X     = [Terminal.Gui.Pos]::Top($MenuBar) + 1
    Y     = [Terminal.Gui.Pos]::Top($MenuBar) + 1
}

$ButtonNetwork = [Terminal.Gui.Button]@{
    Text  = "Network"
    Width = 23
    X     = [Terminal.Gui.Pos]::Top($MenuBar) + 1
    Y     = [Terminal.Gui.Pos]::Top($MenuBar) + 1

}
$ButtonMain.add_Clicked({
        $LeftFrame.RemoveAll()
        $LeftBottomFrame.RemoveAll()
        $RightFrame.RemoveAll()
        [Terminal.Gui.Application]::Top.add($StatusBar)
        [Terminal.Gui.Application]::Top.add($MenuBar)
        $window.Add($LeftFrame)
        $window.Add($LeftBottomFrame)
        $window.Add($RightFrame)
        #$RightFrame.Add($Panelview)

        $LeftBottomFrame.Add($ButtonPrinterDeploymentRegistration)
        $LeftBottomFrame.Add($ButtonPrinterDeployment)
        $LeftBottomFrame.Add($ButtonPrinters)
        $LeftBottomFrame.Add($ButtonScan)
        $LeftBottomFrame.Add($ButtonIPv4Info)
        $LeftBottomFrame.Add($ButtonWlanPasswords)
    })

$ButtonNetwork.add_Clicked({

        [Terminal.Gui.Application]::Top.add($StatusBar)
        [Terminal.Gui.Application]::Top.add($MenuBar)
        $window.Add($LeftFrame)
        $window.Add($LeftBottomFrame)
        $window.Add($RightFrame)
        #$RightFrame.Add($Panelview)
        $LeftFrame.Add($LabelSubnet)
        $LeftFrame.Add($TextfieldSubnet)
        $LeftFrame.Add($LabelRangeStart)
        $LeftFrame.Add($TextfieldRangeStart)
        $LeftFrame.Add($LabelRangeEnd)
        $LeftFrame.Add($TextfieldRangeEnd)
        $LeftBottomFrame.Add($ButtonPrinterDeploymentRegistration)
        $LeftBottomFrame.Add($ButtonPrinterDeployment)
        $LeftBottomFrame.Add($ButtonPrinters)
        $LeftBottomFrame.Add($ButtonScan)
        $LeftBottomFrame.Add($ButtonIPv4Info)
        $LeftBottomFrame.Add($ButtonWlanPasswords)
        $RightFrame.Add($TableView)
    })


function UpdateButtonWidths {
    param (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Window containing the buttons")]
        [object]$Window
    )

    # Recursive helper function to update button widths in a view
    function UpdateButtonWidthsInView {
        param (
            [Parameter(Mandatory = $true, Position = 0)]
            [object]$View
        )

        # Check if the view is a button
        if ($View -is [Terminal.Gui.Button]) {
            # Update the button's width based on its text length
            $View.Width = $View.Text.Length + 4
        }

        # If the view contains subviews, update button widths in each subview
        foreach ($subview in $View.Subviews) {
            UpdateButtonWidthsInView -View $subview
        }
    }

    # Update button widths in the window and all its subviews
    UpdateButtonWidthsInView -View $Window
}



[Terminal.Gui.Application]::Top.Add($window)
[Terminal.Gui.Application]::Top.add($StatusBar)
[Terminal.Gui.Application]::Top.add($MenuBar)
$window.Add($ButtonMain)
$ButtonMain.OnClicked()
#$ButtonPrinterDeploymentRegistration.OnClicked()
UpdateButtonWidths -Window $window

[Terminal.Gui.Application]::Refresh()
[Terminal.Gui.Application]::Run()
[Terminal.Gui.Application]::ShutDown()
#endregion

