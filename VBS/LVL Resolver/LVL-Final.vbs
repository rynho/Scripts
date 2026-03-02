Dim i1,j1
Dim S(9,9)
Dim Fit(22,2)

For i1 = 0 to 8
  For j1 = 0 to 8
    S(i1,j1) = " - "
  Next
Next

For i1 = 0 to 21
  For j1 = 0 to 1
    Fit(i1,j1) = -1
  Next
Next


Dim blocks(22)
blocks(0)  = Arr(" A ", 5, "0001102021")
blocks(1)  = Arr(" B ", 5, "0110111221")
blocks(2)  = Arr(" C ", 4, "00011020")
blocks(3)  = Arr(" D ", 4, "00011020")
blocks(4)  = Arr(" E ", 4, "00011121")
blocks(5)  = Arr(" F ", 1, "00")
blocks(6)  = Arr(" G ", 4, "01112021")
blocks(7)  = Arr(" H ", 4, "01101112")
blocks(8)  = Arr(" I ", 2, "0001")
blocks(9)  = Arr(" J ", 4, "02101112")
blocks(10) = Arr(" K ", 3, "000102")
blocks(11) = Arr(" L ", 3, "000102")
blocks(12) = Arr(" M ", 4, "00011020")
blocks(13) = Arr(" N ", 4, "00010210")
blocks(14) = Arr(" O ", 4, "01101112")
blocks(15) = Arr(" P ", 4, "00011121")
blocks(16) = Arr(" Q ", 5, "0002101112")
blocks(17) = Arr(" R ", 4, "02101112")
blocks(18) = Arr(" S ", 2, "0010")
blocks(19) = Arr(" T ", 3, "011011")
blocks(20) = Arr(" U ", 4, "00011011")
blocks(21) = Arr(" V ", 4, "00011011")


n = 0
Test 0
'Verify


Sub Verify
  For i = 0 to 21
    If Put(i,0,0) Then Draw S
	Tak i
  Next
End Sub



Sub Test (num)
  Dim i,j
  For i = 0 to 8
	For j = 0 to 8
	  If Put(num,i,j) Then
'		  Draw S
		If num = 21 Then
		  Draw S
		  n = n+1
		  Tak num
		Else
		  Test (num+1)
		End If
	  ElseIf i=8 and j=8 and num>0 Then
	    Tak num-1
	  End If
'	  If num=0 Then MsgBox i&j
	Next
  Next

  If num=0 Then MsgBox "Completed with "& n &" results."
End Sub

Function Arr (sym,dm,corr)
  Dim i
  ReDim arr_t(dm,2)
  arr_t(0,0) = sym
  For i = 1 to dm
    arr_t(i,0)= Cint(Mid(corr,i*2-1,1))
    arr_t(i,1)= Cint(Mid(corr,i*2,1))
  Next
  Arr = arr_t  
End Function

Function Put (num,x,y)
  Dim i
  For i = 1 to UBound(blocks(num),1)
'   msgbox blocks(num)(0,0) &"("& i &") is ("& x+blocks(num)(i,0) &","& y+blocks(num)(i,1) &")."
    If (x+blocks(num)(i,0) > 8) or (y+blocks(num)(i,1) > 8) Then
	  Put = False
'	  msgbox blocks(num)(0,0) &"("& i &") is ("& x+blocks(num)(i,0) &","& y+blocks(num)(i,1) &")."
	  Exit Function
	ElseIf (S(x+blocks(num)(i,0),y+blocks(num)(i,1)) <> " - ")  Then
	  Put = False
'	  msgbox "S("& x+blocks(num)(i,0) &","& y+blocks(num)(i,1) &") is occupied."
	  Exit Function
	End If
  Next
  For i = 1 to UBound(blocks(num),1)
    S(x+blocks(num)(i,0),y+blocks(num)(i,1)) = blocks(num)(0,0)
	Fit(num,0)=x:Fit(num,1)=y
  Next
  Put = True
End Function

Sub Tak (num)
  Dim i
  x=Fit(num,0):y=Fit(num,1)
  For i = 1 to UBound(blocks(num),1)
    S(x+blocks(num)(i,0),y+blocks(num)(i,1)) = " - "
  Next
  Fit(num,0)=-1:Fit(num,1)=-1
End Sub


Sub Draw (s)
  Dim i,j
  str = "   0  1  2  3  4  5  6  7  8" &vbCrLf&_
        "  --------------------------"
  For i = 0 to 8
    str = str &vbCrLf& i &"|"
    For j = 0 to 8
	  str = str & S(i,j)
	Next
  Next
  CusMsgBox str

End Sub

Sub CusMsgBox(msg)
  Set ie = CreateObject("InternetExplorer.Application")
  ie.Navigate "about:blank"

  While ie.ReadyState <> 4 : WScript.Sleep 100 : Wend

  ie.ToolBar   = False
  ie.StatusBar = True
  ie.Width     = 700
  ie.Height    = 800

  ie.document.body.innerHTML = "<p class='msg'><pre>" & msg & "</pre></p>" & _
    "<p class='ctrl'><input type='hidden' id='OK' name='OK' value='0'>" & _
    "<input type='submit' value='OK' id='OKButton' " &_
    "onclick='document.all.OK.value=1'></p>"

  Set style = ie.document.CreateStyleSheet
  style.AddRule "p.msg", "font-family: 'Arial'; font-size: 80px;"
  style.AddRule "body", "font-family:san-serif; font-size:30px; text-align:center;"
  style.AddRule "p.ctrl", "text-align:rightf;"

  ie.Visible = True

  On Error Resume Next
  Do While ie.Document.all.OK.value = 0 
    WScript.Sleep 200
  Loop
  ie.Quit
End Sub
