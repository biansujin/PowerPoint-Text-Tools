$ErrorActionPreference = 'Stop'
$base = [Microsoft.Win32.Registry]::CurrentUser

$paths = @(
    'Software\Microsoft\Office\PowerPoint\Addins\OfficeSmartFontPicker.Connect',
    'Software\Classes\CLSID\{A3D9E13C-5732-4E7E-B6C7-0A4E8D8395B1}',
    'Software\Classes\CLSID\{D1F2AE42-A212-4A7B-8E03-3C9FE4F22B77}',
    'Software\Classes\OfficeSmartFontPicker.Connect',
    'Software\Classes\OfficeSmartFontPicker.FontPickerPane'
)

foreach ($path in $paths) {
    try { $base.DeleteSubKeyTree($path, $false) } catch { }
}

Write-Host ''
Write-Host 'PowerPoint 文字工具 uninstalled.' -ForegroundColor Green
Write-Host 'Please fully restart PowerPoint.'
