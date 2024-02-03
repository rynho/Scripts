
   Const ForReading = 1, ForWriting = 2
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile(".\testfile.txt", ForReading)
   ReadAllTextFile = Replace(f.ReadAll,VbCrLf,"")
   ReadAllTextFile = Replace(ReadAllTextFile," ","")


file = split(readalltextfile,",")

   Set f = fso.OpenTextFile(".\testfile.bin", ForWriting, True)
   for i = 0 to UBound(file)
     f.write Chr(file(i))
   Next
	 
	 
