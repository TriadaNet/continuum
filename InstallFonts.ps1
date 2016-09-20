$path = "c:\!installs\fonts"
$fontpath = "c:\!installs\fonts\fontdownload"

New-Item -ItemType directory -Path $path -force

$shell = new-object -com shell.application

$zip = $shell.NameSpace( c:\!installs\fontdownload.zip")

foreach($item in $zip.items()) {

   $shell.Namespace($path).copyhere($item,0x14)

}

$Fontdir = dir $fontpath
$sa =  new-object -comobject shell.application
$Fonts =  $sa.NameSpace(0x14)
foreach($File in $Fontdir) {
  if ((Test-Path "C:\Windows\Fonts\$File") -eq $False) {
    $Fonts.CopyHere($File.fullname)
  }
}
