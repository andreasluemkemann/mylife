Install-Module TerminalGuiDesigner -ErrorAction SilentlyContinue
Import-Module TerminalGuiDesigner  -ErrorAction SilentlyContinue
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$guiTools = (Get-Module TerminalGuiDesigner -List).ModuleBase
Set-Location -Path $scriptDir
Add-Type -Path (Join-path $guiTools Terminal.Gui.dll)
[Terminal.Gui.Application]::Init()
$Window = .\tui.ps1
[Terminal.Gui.Application]::QuitKey = 27 # ESC
[Terminal.Gui.Application]::Top.Add($Window)
[Terminal.Gui.Application]::Run()