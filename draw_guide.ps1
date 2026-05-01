Add-Type -AssemblyName System.Drawing

$Path = Join-Path (Get-Location) "screen.png"
$OutPath = Join-Path (Get-Location) "guide.png"

$bmp = [System.Drawing.Image]::FromFile($Path)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Red, 10)

# Draw arrow pointing to the Chrome tab "Deploy an app"
$g.DrawRectangle($pen, 280, 5, 200, 40)
$g.DrawLine($pen, 380, 150, 380, 50)
$g.DrawLine($pen, 380, 50, 360, 70)
$g.DrawLine($pen, 380, 50, 400, 70)

$g.Dispose()
$bmp.Save($OutPath, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
echo "Guide image saved to $OutPath"
