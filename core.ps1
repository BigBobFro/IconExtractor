<#
    .TITLE
        Windows Icon Extractor
    .SYNOPSIS
        Extract icons from existing app/dll/exe files and save as BMP file for use in other apps and scripts
    .DESCRIPTION

    .DEPENDANCIES
    - Known location of source file

    ========================================================================================================
    .Notes
    Author      :   Fro
    Date        :   15 May 2025
    Git         :   https://github.com/BigBobFro
    Version     :   1.0 (Final)

    ========================================================================================================
    .CHANGELOG
    1.0 - First final release

    ========================================================================================================
    .ToDo
    Add ability to use custom files to extract icons from, store them in Registry for regular use
        and manage (remove) when needed.
#>

# ====================================================
# Set file to extract from
$DefaultFileList = @{}
$DefaultFileList.add($("Shell32"),$("C:\windows\system32\shell32.dll"))
$DefaultFileList.add("imageRes","C:\windows\system32\imageres.dll")
$DefaultFileList.add("MS-Access","C:\Program Files\Microsoft Office\root\Office16\msaccess.exe")
$DefaultFileList.add("SSMS","C:\Program Files\Microsoft SQL Server Management Studio 20\Common7\IDE\Microsoft.SqlServer.CustomControls.dll")

# ====================================================
# Include .NET files for GUI
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# ====================================================
# Icon Extraction Code Base
$code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System{
    public class IconExtractor{
        public static Icon Extract(string file, int number, bool largeIcon){
            IntPtr large;
            IntPtr small;
            ExtractIconEx(file, number, out large, out small, 1);
            try{return Icon.FromHandle(largeIcon ? large : small);}
            catch{return null;}
        }
        [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);
    }
}
"@

Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing

# =====================================================
# Constants
$Global:i                   = 0     # Icon Startpoint
$srcPath                    = split-path -path $MyInvocation.MyCommand.path
$regRoot                    = "HKCU:\SOFTWARE\FAB\IconExtractor"
$Debug                      = $false

# Visual Element Constants
$LargeButtonSize            = [System.Drawing.Size]::new(50,50)
$SmallButtonSize            = [System.Drawing.Size]::new(50,25)
$DefaultFont                = [System.Drawing.Font]::new("Microsoft Sans Serif",10, [System.Drawing.FontStyle]::Regular)
$TitleFont                  = [System.Drawing.Font]::new("Microsoft Sans Serif",18, [System.Drawing.FontStyle]::Bold)

function UpdateIcon {           # Just got tired of typing this out all the time
    $idLabel.Text       = $global:i
    $testButton.Image   = [system.IconExtractor]::Extract($Global:iconSource, $global:i, $true)
}


# ======================================================
# Window Control Objects

$testButton                 = New-Object System.Windows.Forms.Button
$testButton.Location        = [System.Drawing.Size]::new(25,25)
$testButton.Size            = $LargeButtonSize
$testButton.add_click({ 
    if($debug){write-host "Test Button pressed"}
})

$idLabel                    = New-Object System.Windows.Forms.Label
$idLabel.Location           = [System.Drawing.Size]::new(25,100)
$idLabel.Size               = [System.Drawing.Size]::new(100,25)
$idLabel.Font               = $TitleFont

UpdateIcon

$Jump2Box                   = New-Object System.Windows.Forms.TextBox
$Jump2Box.Location          = [System.Drawing.Size]::new(25,150)
$Jump2Box.Size              = [System.Drawing.Size]::new(100,25)

$Jump2Button                = New-Object System.Windows.Forms.Button
$Jump2Button.Location       = [System.Drawing.Size]::new(25,175)
$Jump2Button.Size           = [System.Drawing.Size]::new(100,25)
$Jump2Button.Font           = $DefaultFont
$Jump2Button.Text           = "Jump to"
$Jump2Button.add_click({
    if($debug){write-host "Jump to button pressed"}

    if([string]::isnullorempty($Jump2Box.Text)){write-host "No jump provided"}
    else{
        if($($Jump2Box.Text) -match "^[\d\.]+$"){
            if($debug){write-host "Jump entry is numeric"}
            $global:i           = [int]$($Jump2Box.text)
            UpdateIcon
        }
        else{write-host "Jump 2 is not numeric"}
    }
})

$PlusButton                 = New-Object System.Windows.Forms.Button
$PlusButton.Location        = [System.Drawing.Size]::new(75,200)
$PlusButton.Size            = $SmallButtonSize
$PlusButton.Text            = "+"
$PlusButton.Font            = $TitleFont
$PlusButton.add_click({
    $global:i = $global:i + 1
    UpdateIcon
})

$MinusButton                = New-Object System.Windows.Forms.Button
$MinusButton.Location       = [System.Drawing.Size]::new(25,200)
$MinusButton.Size           = $SmallButtonSize
$MinusButton.Text           = "-"
$MinusButton.Font           = $TitleFont
$MinusButton.add_click({
    $global:i               = $global:i - 1
    UpdateIcon
})

$ExportName                 = New-Object System.Windows.Forms.TextBox
$ExportName.Location        = [System.Drawing.Size]::new(350,25)
$ExportName.Size            = [System.Drawing.Size]::new(100,25)


$ExtractButton              = New-Object System.Windows.Forms.Button
$ExtractButton.Location     = [System.Drawing.Size]::new(350,50)
$ExtractButton.Size         = [System.Drawing.Size]::new(75,50)
$ExtractButton.Text         = "Extract"
$ExtractButton.Font         = $DefaultFont
$ExtractButton.add_click({
    if([string]::isnullorempty($exportname.text)){
        write-host "No extract name provided"
    }
    else{
        if($debug){write-host "Name provided: $($exportname.text)"}
        if($($exportname.text) -notlike "*.bmp"){
            if($debug){write-host "File extention BMP not provided.  Adding."}
            $exportname.text = "$($exportname.text)" + ".bmp"
        }
        $Icon = [System.IconExtractor]::Extract($Global:iconSource,$($global:i), $true)
        $Icon.tobitmap().save("$srcpath\$($exportname.text)")
    }
})

# File Source objects

$FileSourceList             = $DefaultFileList # Add custom file locations from Registry
$FileSelectDD               = New-Object System.Windows.Forms.ComboBox
$FileSelectDD.Location      = [System.Drawing.Size]::New(150,25)
$FileSelectDD.Size          = [System.Drawing.Size]::new(150,25)
foreach($f in $filesourcelist.Keys){
    write-host "File to add: $f"
    [void]$FileSelectDD.Items.Add($($f))
}
# Pull last file used for extraction from Registry
$filename = Get-ItemPropertyValue -path $regRoot -name "IconSource"
# Check if file is in file list
if($FileSourceList.Keys -match $filename){$FileSelectDD.SelectedItem  = $filename}
else{$FileSelectDD.SelectedItem  = "Shell32"}

$Global:iconSource                 = $($FileSourceList[$($FileSelectDD.SelectedItem)])
$FileSelectDD.add_selectedIndexChanged({
    $Global:iconSource             = $($FileSourceList[$($FileSelectDD.SelectedItem)])
    UpdateIcon
})

# ===========================================
# Principal Window Object(s)

$exitButton                 = New-Object System.Windows.Forms.Button
$exitButton.Location        = [System.Drawing.Size]::new($($MainWindow.width - 100),$($MainWindow.Height - 100))
$exitbutton.Size            = [System.Drawing.Size]::new(75,50)
$exitButton.Text            = "Exit"
$exitButton.Font            = $TitleFont
$exitbutton.add_click({
    # set current setting to registry for later retrieval
    if($debug){write-host "Exit button pressed"}
    if(!$(test-path $regRoot)){
        New-Item -Path $regRoot -Force
        New-ItemProperty -path $regRoot -name "IconSource"
    }
    Set-ItemProperty -path $regRoot -name "IconSource" -value "$($Global:iconSource)"
    $MainWindow.Close()
})


$MainWindow                 = New-Object System.Windows.Forms.Form
$MainWindow.Text            = "Windows Icon Extractor"
$MainWindow.Size            = [System.Drawing.Size]::new(500,300)
$MainWindow.KeyPreview      = $true
$MainWindow.TopMost         = $true
$MainWindow.StartPosition   = "CenterScreen"

$MainWindow.Controls.add($testButton)
$MainWindow.Controls.add($idLabel)
$MainWindow.Controls.add($ExtractButton)
$MainWindow.Controls.add($ExportName)
$MainWindow.Controls.add($PlusButton)
$MainWindow.Controls.add($MinusButton)
$MainWindow.Controls.add($Jump2Box)
$MainWindow.Controls.add($Jump2Button)
$MainWindow.Controls.Add($FileSelectDD)

$MainWindow.Controls.add($exitButton)
[void] $MainWindow.ShowDialog()

