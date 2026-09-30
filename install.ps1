param(
    [string]$DllPath = (Join-Path $PSScriptRoot 'dist\OfficeSmartFontPicker.dll')
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $DllPath)) {
    throw "DLL not found: $DllPath"
}

$dll = (Resolve-Path $DllPath).Path
$codeBase = 'file:///' + ($dll -replace '\\','/')
$assembly = 'OfficeSmartFontPicker, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null'
$base = [Microsoft.Win32.Registry]::CurrentUser

function Set-String([string]$Path,[string]$Name,[string]$Value) {
    $key = $base.CreateSubKey($Path)
    try { $key.SetValue($Name,$Value,[Microsoft.Win32.RegistryValueKind]::String) }
    finally { $key.Close() }
}

function Set-Dword([string]$Path,[string]$Name,[int]$Value) {
    $key = $base.CreateSubKey($Path)
    try { $key.SetValue($Name,$Value,[Microsoft.Win32.RegistryValueKind]::DWord) }
    finally { $key.Close() }
}
function Register-ComClass([string]$Clsid,[string]$ProgId,[string]$ClassName) {
    $ip = "Software\Classes\CLSID\$Clsid\InprocServer32"

    Set-String "Software\Classes\CLSID\$Clsid" '' $ProgId
    Set-String $ip '' 'mscoree.dll'
    Set-String $ip 'ThreadingModel' 'Both'
    Set-String $ip 'Class' $ClassName
    Set-String $ip 'Assembly' $assembly
    Set-String $ip 'RuntimeVersion' 'v4.0.30319'
    Set-String $ip 'CodeBase' $codeBase

    $ver = "$ip\0.0.0.0"
    Set-String $ver 'Class' $ClassName
    Set-String $ver 'Assembly' $assembly
    Set-String $ver 'RuntimeVersion' 'v4.0.30319'
    Set-String $ver 'CodeBase' $codeBase

    Set-String "Software\Classes\$ProgId" '' $ProgId
    Set-String "Software\Classes\$ProgId\CLSID" '' $Clsid
}

Register-ComClass '{A3D9E13C-5732-4E7E-B6C7-0A4E8D8395B1}' 'OfficeSmartFontPicker.Connect' 'OfficeSmartFontPicker.Connect'
Register-ComClass '{D1F2AE42-A212-4A7B-8E03-3C9FE4F22B77}' 'OfficeSmartFontPicker.FontPickerPane' 'OfficeSmartFontPicker.FontPickerPane'
$base.CreateSubKey('Software\Classes\CLSID\{D1F2AE42-A212-4A7B-8E03-3C9FE4F22B77}\Control').Close()

$addin = 'Software\Microsoft\Office\PowerPoint\Addins\OfficeSmartFontPicker.Connect'
Set-String $addin 'FriendlyName' ([string]([char]0x6587)+[char]0x5B57+[char]0x5DE5+[char]0x5177)
Set-String $addin 'Description' 'PowerPoint Typography Tools'
Set-Dword $addin 'LoadBehavior' 3

Write-Host ''
Write-Host 'PowerPoint 文字工具 V1.0.0 installed.' -ForegroundColor Green
Write-Host 'Please fully restart PowerPoint.'
