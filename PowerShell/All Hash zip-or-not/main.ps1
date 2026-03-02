param($Work)
# restart PowerShell with -noexit, the same script, and 1
if (!$Work) {
    powershell -noexit -file $MyInvocation.MyCommand.Path 1
    return
}

# now the script does something
@'

This script is running below then wait.
$output = '_HashList'

$target = <full-path of target>

all-hash $target $output

'@

#$ErrorActionPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'

Add-Type -TypeDefinition @"
// Copyright (c) Damien Guard.  All rights reserved.
// Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. 
// You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0
using System;
using System.Collections.Generic;
using System.Security.Cryptography;
/// <summary>
/// Implements a 32-bit CRC hash algorithm compatible with Zip etc.
/// </summary>
/// <remarks>
/// Crc32 should only be used for backward compatibility with older file formats
/// and algorithms. It is not secure enough for new applications.
/// If you need to call multiple times for the same data either use the HashAlgorithm
/// interface or remember that the result of one Compute call needs to be ~ (XOR) before
/// being passed in as the seed for the next Compute call.
/// </remarks>
public sealed class Crc32 : HashAlgorithm
{
    public const UInt32 DefaultPolynomial = 0xedb88320u;
    public const UInt32 DefaultSeed = 0xffffffffu;
    static UInt32[] defaultTable;
    readonly UInt32 seed;
    readonly UInt32[] table;
    UInt32 hash;
    public Crc32()
        : this(DefaultPolynomial, DefaultSeed)
    {
    }
    public Crc32(UInt32 polynomial, UInt32 seed)
    {
        table = InitializeTable(polynomial);
        this.seed = hash = seed;
    }
    public override void Initialize()
    {
        hash = seed;
    }
    protected override void HashCore(byte[] array, int ibStart, int cbSize)
    {
        hash = CalculateHash(table, hash, array, ibStart, cbSize);
    }
    protected override byte[] HashFinal()
    {
        var hashBuffer = UInt32ToBigEndianBytes(~hash);
        HashValue = hashBuffer;
        return hashBuffer;
    }
    public override int HashSize { get { return 32; } }
    public static UInt32 Compute(byte[] buffer)
    {
        return Compute(DefaultSeed, buffer);
    }
    public static UInt32 Compute(UInt32 seed, byte[] buffer)
    {
        return Compute(DefaultPolynomial, seed, buffer);
    }
    public static UInt32 Compute(UInt32 polynomial, UInt32 seed, byte[] buffer)
    {
        return ~CalculateHash(InitializeTable(polynomial), seed, buffer, 0, buffer.Length);
    }
    static UInt32[] InitializeTable(UInt32 polynomial)
    {
        if (polynomial == DefaultPolynomial && defaultTable != null)
            return defaultTable;
        var createTable = new UInt32[256];
        for (var i = 0; i < 256; i++)
        {
            var entry = (UInt32)i;
            for (var j = 0; j < 8; j++)
                if ((entry & 1) == 1)
                    entry = (entry >> 1) ^ polynomial;
                else
                    entry = entry >> 1;
            createTable[i] = entry;
        }
        if (polynomial == DefaultPolynomial)
            defaultTable = createTable;
        return createTable;
    }
    static UInt32 CalculateHash(UInt32[] table, UInt32 seed, IList<byte> buffer, int start, int size)
    {
        var hash = seed;
        for (var i = start; i < start + size; i++)
            hash = (hash >> 8) ^ table[buffer[i] ^ hash & 0xff];
        return hash;
    }
    static byte[] UInt32ToBigEndianBytes(UInt32 uint32)
    {
        var result = BitConverter.GetBytes(uint32);
        if (BitConverter.IsLittleEndian)
            Array.Reverse(result);
        return result;
    }
}
"@ -PassThru | Out-Null

Function get-crc32
{
    Param (
        [Parameter(Mandatory=$true)][String]$path
    )

#   $ErrorActionPreference = "Stop"
    $crc32 = New-Object Crc32 
    $stream = New-Object IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
#   $date_size = gi $path | %{'{0:yyyyMMdd_HH:mm:ss} | {1,13:N0}' -f $_.lastWriteTime, $_.length}
    $hash = [String]::Empty

    foreach ($byte in $crc32.ComputeHash($stream))
    {
        $hash += $byte.toString('x2').toUpper()
    }

    $stream.Close()

#   echo "$hash : $date_size : $path"
    echo "$hash"
}

function sum-crc32 ($file) {
  dir -LiteralPath $file | %{'file|{0}|{1:yyyyMMdd HH:mm:ss}|{2,15:N0}|{3}' -f $(get-crc32 $_.FullName), $_.lastWriteTime, $_.length, $_.fullname}
}

function Get-FolderCRC ($folder) {
  dir -LiteralPath $folder -Recurse | ?{!$_.psiscontainer} | %{sum-crc32 $_.FullName}
}

function 7z-crc32 ($file) {
  $dir = $PSScriptRoot
  echo "`n>>>>>>>>> $(date -format 'yyyy-mm-dd HH:mm:ss') >>>>>>>>>>"
  & "$dir\7z.exe" l -slt $file | ?{$_ -match '^(?:Path|Size|Modified|CRC|--)'}
}

function all-hash ($folder, $output) {
  $7zType = ".7z",".zip",".rar"

  if ($output -eq $null) {
     dir -LiteralPath $folder -Recurse | ?{!$_.psiscontainer} | `
        %{ sum-crc32 $_.FullName
           if ($_.Extension -in $7zType) {7z-crc32 $_.FullName}
        }
  } else {
     $output_file = $output+"_file.txt"
     $output_7z = $output+"_7z.txt"
     
     if (test-path $output_file) {del $output_file}
     if (test-path $output_7z) {del $output_7z}

     dir -LiteralPath $folder -Recurse | ?{!$_.psiscontainer} | `
        %{ sum-crc32 $_.FullName >> $output_file
           if ($_.Extension -in $7zType) {7z-crc32 $_.FullName >> $output_7z}
        }
     sort-HL7z ($output_7z, $output_file)
  }
}

function sort-HL7z ($file, $output) {
  $FileContent = Get-Content -Path $file

  $array = @('','','','')
  $flag = 0
  $package = ''

  foreach ($line in $FileContent) {
    switch -regex ($line) {
      '^>{8}.+>{8}$' {
        $flag = 0
        countinue
      }

      '^--$' {
        $flag = 1
        countinue
      }

      '^-{10}$' {
        $flag = 2
        countinue
      }

      '^Path = (.*)$' {
        if ($flag -eq 1) {
          $package = $Matches[1]
          continue
        }
        if ($flag -eq 2) {
          $array[3] = $package+"\"+$Matches[1]
          continue
        }
      }

      '^Size = (.*)$' {
        $array[2] = '{0,15:N0}' -f [int]$Matches[1]
        continue
      }

      '^Modified = (.{19}).*$' {
        $array[1] = $Matches[1].replace('-','')
        continue
      }

      '^CRC = (.*)$' {
        if ($Matches[1] -eq '') {$array[0] = '00000000'} else {$array[0] = $Matches[1]}
        "7zip|"+$array[0]+"|"+$array[1]+"|"+$array[2]+"|"+$array[3] >> $output
        continue
      }
    }
  }
}

##### Main Program #####
$output = '_HashList'

do {
  $target = Read-Host -Prompt "Please provide full-path of the folder for analysis`n"
} while ([string]::IsNullOrEmpty($target))

all-hash $target $output

"All completed.`n"





