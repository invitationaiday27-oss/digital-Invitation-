$path = 'c:\Users\USER\Downloads\KADDIGITAL'
$file = 'WhatsApp Video 2026-10-12 at 4.28.59 PM.mp4'
$shell = New-Object -ComObject Shell.Application
$folder = $shell.Namespace($path)
$item = $folder.ParseName($file)
# Dapatkan beberapa sifat video
0..320 | ForEach-Object {
  $name = $folder.GetDetailsOf($null, $_)
  $val  = $folder.GetDetailsOf($item, $_)
  if ($val -and $name) { "{0}: {1}" -f $name, $val }
}
