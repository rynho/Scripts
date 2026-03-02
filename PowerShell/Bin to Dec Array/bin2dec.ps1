$file = $args[0]

$hex = (Format-Hex -Path $file `
  | Select-Object -Expand Bytes `
  | ForEach-Object { '{0:D}' -f $_ } `
  | ForEach-Object { $i = 0 } `
    { $i++; @{$true="`r`n";$false=""}[($i - 1) % 16 -eq 0] + $_ } `
) -join ', '

$hex > "$($file).txt"
