Add-Type -AssemblyName System.Drawing

$Path = Join-Path (Get-Location) "screen.png"
$OutPath = Join-Path (Get-Location) "guide_taskbar.png"

$bmp = [System.Drawing.Image]::FromFile($Path)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Red, 15)

# Draw arrow pointing to the Chrome icon in the taskbar
# Chrome icon is roughly at x=580, y=930 based on the screenshot (centered taskbar)
$g.DrawRectangle($pen, 560, 910, 40, 40) # Box around Chrome icon
$g.DrawLine($pen, 580, 850, 580, 910) # Vertical line
$g.DrawLine($pen, 580, 910, 560, 890) # Arrow head left
$g.DrawLine($pen, 580, 910, 600, 890) # Arrow head right

$g.Dispose()
$bmp.Save($OutPath, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
echo "Taskbar guide saved to $OutPath"
