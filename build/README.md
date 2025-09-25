# Цей файл містить базову іконку програми для Windows
# Якщо у вас є Adobe Illustrator, GIMP або онлайн конвертер,
# конвертуйте icon.svg в icon.ico розміром 256x256 пікселів

# Альтернативно, використайте цю команду PowerShell для створення базової іконки:
# Add-Type -AssemblyName System.Drawing
# $bitmap = New-Object System.Drawing.Bitmap(256, 256)
# $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
# $graphics.Clear([System.Drawing.Color]::Blue)
# $bitmap.Save("$PWD\build\icon.ico", [System.Drawing.Imaging.ImageFormat]::Icon)

# Або скачайте готову іконку з інтернету та помістіть сюди як icon.ico