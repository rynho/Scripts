# This script find .PNG files in the PSX manual: Document.dat that
#   associated with EBOOT.pbp, and extract all of them one by one.
#
# Usage:
#   1. Put the Document.dat and this script in same folder.
#   2. Right click in folder to "Open in Terminal".
#   3. Run ".\dat2png" in powershell to extract .png into
#      "_png" subfolder.

if (!(test-path -path ".\document.dat")) {
  write-host "DOCUMENT.DAT not presenting."
  pause
  exit
}

write-host "starting..."

# Define the path
$path = ".\document.dat"

# Read binary file using ISO-8859-1 encoding (Codepage 28591)
$encoding = [System.Text.Encoding]::GetEncoding(28591)
$binaryText = [System.IO.File]::ReadAllText($path, $encoding)

# Define your pattern (e.g., matching the hex sequence 0x4D 0x5A)
$pattern = "(?s)\x89\x50\x4E\x47.+?\x49\x45\x4E\x44\xAE\x42\x60\x82"

# Match and save the results
$matches = [regex]::Matches($binaryText, $pattern)

if ($matches.count -gt 0) {
  New-Item -Path "_png" -ItemType Directory -Force
}

$i=1
foreach ($match in $matches) {
  $pg = "{0:d3}" -f $i
  write-host "found $pg page"
  $bytes = [System.Text.Encoding]::GetEncoding(28591).GetBytes($match)
  Set-Content -Path "_png\ext$pg.png" -Value $bytes -Encoding Byte
  $i++
}

pause
