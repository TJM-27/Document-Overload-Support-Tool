[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$projectRoot = $PSScriptRoot
$outputPath = Join-Path $projectRoot 'DOST.exe'
$iconPath = Join-Path $projectRoot 'DOST-PaperStack.ico'

function New-DostIconFile {
    param([Parameter(Mandatory)][string]$Path)

    Add-Type -AssemblyName System.Drawing
    $bitmap = New-Object System.Drawing.Bitmap(32, 32)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)

    try {
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.Clear([System.Drawing.Color]::Transparent)

        $outlinePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(19, 78, 52), 1.4)
        $shadowBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(85, 0, 0, 0))
        $backBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(207, 230, 211))
        $middleBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(228, 242, 231))
        $frontBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(242, 249, 243))
        $linePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(72, 154, 105), 1)

        try {
            $graphics.FillRectangle($shadowBrush, 9, 7, 17, 21)
            $graphics.FillRectangle($backBrush, 5, 3, 17, 21)
            $graphics.DrawRectangle($outlinePen, 5, 3, 17, 21)
            $graphics.FillRectangle($middleBrush, 8, 6, 17, 21)
            $graphics.DrawRectangle($outlinePen, 8, 6, 17, 21)
            $graphics.FillRectangle($frontBrush, 11, 9, 17, 21)
            $graphics.DrawRectangle($outlinePen, 11, 9, 17, 21)
            $graphics.DrawLine($linePen, 14, 15, 25, 15)
            $graphics.DrawLine($linePen, 14, 19, 25, 19)
            $graphics.DrawLine($linePen, 14, 23, 22, 23)
        } finally {
            $outlinePen.Dispose()
            $shadowBrush.Dispose()
            $backBrush.Dispose()
            $middleBrush.Dispose()
            $frontBrush.Dispose()
            $linePen.Dispose()
        }

        $icon = [System.Drawing.Icon]::FromHandle($bitmap.GetHicon())
        $stream = [System.IO.File]::Create($Path)
        try {
            $icon.Save($stream)
        } finally {
            $stream.Dispose()
            $icon.Dispose()
        }
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

$source = @'
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

public static class DostLauncher
{
    [STAThread]
    public static int Main()
    {
        string exePath = Application.ExecutablePath;
        string projectRoot = Path.GetDirectoryName(exePath);
        string scriptPath = Path.Combine(projectRoot, "DOST-ZipDocumentGui.ps1");

        if (!File.Exists(scriptPath))
        {
            MessageBox.Show(
                "Could not find DOST-ZipDocumentGui.ps1 beside DOST.exe.",
                "DOST - Document Overload Support Tool",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }

        string escapedScriptPath = scriptPath.Replace("\"", "\\\"");
        var startInfo = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = "-NoProfile -ExecutionPolicy Bypass -STA -File \"" + escapedScriptPath + "\"",
            WorkingDirectory = projectRoot,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        try
        {
            Process.Start(startInfo);
            return 0;
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Unable to start DOST.`r`n`r`n" + ex.Message,
                "DOST - Document Overload Support Tool",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }
    }
}
'@

if (Test-Path -LiteralPath $outputPath) {
    Remove-Item -LiteralPath $outputPath -Force
}

New-DostIconFile -Path $iconPath

Add-Type -AssemblyName Microsoft.CSharp

$provider = New-Object Microsoft.CSharp.CSharpCodeProvider
$parameters = New-Object System.CodeDom.Compiler.CompilerParameters
$parameters.GenerateExecutable = $true
$parameters.OutputAssembly = $outputPath
$parameters.CompilerOptions = "/target:winexe /platform:anycpu /win32icon:`"$iconPath`""
[void]$parameters.ReferencedAssemblies.Add('System.dll')
[void]$parameters.ReferencedAssemblies.Add('System.Windows.Forms.dll')

$result = $provider.CompileAssemblyFromSource($parameters, $source)
if ($result.Errors.HasErrors) {
    foreach ($errorItem in $result.Errors) {
        Write-Error ("{0}({1},{2}): {3}" -f $errorItem.FileName, $errorItem.Line, $errorItem.Column, $errorItem.ErrorText)
    }
    exit 1
}

Write-Host "Created executable: $outputPath"
