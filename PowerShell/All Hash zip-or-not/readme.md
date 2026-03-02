Run main.ps1 using "." operator prefix to register CRC32 algorithm,
as well as all functions for calculating hashes:
`. .\main.ps1`

The five functions can then be called:
  `sum-crc32 <file name>`
   Calculate CRC32 for one file, align with 7z calculation.

  `Get-FolderCRC <folder name>`
   Calculate CRC32 for every file in a folder recursively.

   `7z-crc32 <file name>`
   List CRC32 checksum of all content from single compressed package,
   would it be either .zip, .7z, or .rar format.

   `all-hash <folder name> [output]`
   Calculate CRC32 for every file in a folder recursively;
   for any compressed package amongst them, list the CRC32 checksum
   of the content as well.
   When define file name as [output] parameter, the CRC32 of every file
   and compressed package prints into the output file.

  `sort-HL7z <file name> <output file>`
   Process the 7z-crc32 output file into same format as sum-crc32
   (output in all-hash function), and append to "output file"
   for further analysis.


A typical use case would be running below in order:
1. Right click the "main.ps1" script, and "Run with PowerShell".
   "7z.exe" should be in same folder of the script.

2. The script sets below output files for result:
   `$output = "_HashList"`

3. The script prompts for the target folder. Copy the path of target
   folder (target full-path), and paste to the prompt, and type Enter.

4. The script will run below to get CRC32 of all files and content of
   compressed package, and then process the "_HashList_7z.txt" (7z-crc32
   result) to compatible format, and append to "_HashList.txt".
   `all-hash $target $output`

5. After the script completed, it will stay in PowerShell console.
   Use below to change to other target folder, and manually run all-hash
   function to process more files.
   ```
   $output = "_HashList1"     # This avoids overwriting last result.
   $target = '<target full-path>'`
   all-hash $target $output
   ```



