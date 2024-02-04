$hex = Get-Content -Path $args[0] -Raw

# split the input string using ',' as delimiter, and ignor space and line break;
# casting the result to [byte[]] then recognizes this hex format directly.

[byte[]]$bytes = ($hex -split ',').Trim()

[System.IO.File]::WriteAllBytes("$($args[0]).bin", $bytes)