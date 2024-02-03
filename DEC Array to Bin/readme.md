- This .vbs code is to convert a text file containing an Array of 10-based integer into a binary file.
- Input file format: comma separated array of 10-based integer representing a byte. E.g., '10,' in Input file will be converted to '0A' in binary in Output file.  
All the Carriage-Return (0x0D) / Line-Feed (0x0A) / Space (0x20) characters in Input file are removed before conversion. 

Reason for this is to revert the format used in Arduino files, especially the attachment used in ESP32 Server 900U.
