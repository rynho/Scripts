$hex = Get-Content -Path $args[0] -Raw

# split the input string using ',' as delimiter, and trims spaces and line breaker;
# $hex may contain decimal (no prefix, e.g., '10') or hexadecimal (prefixed with '0x', e.g., '0x0A') representing a byte.

# casting the result to [byte[]] then recognizes this hex format directly.

[byte[]]$bytes = ($hex -split ',').Trim()

[System.IO.File]::WriteAllBytes("$($args[0]).bin", $bytes)