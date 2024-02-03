Dim i,j
Dim S(5,5)
Dim Fit(7,2)

For i = 0 to 4
  For j = 0 to 4
    S(i,j) = " - "
  Next
Next

For i = 0 to 6
  For j = 0 to 1
    Fit(i,j) = -1
  Next
Next


Dim blocks(7)
blocks(0) = Arr(" A ","001020",3)
blocks(1) = Arr(" B ","00010211",4)
blocks(2) = Arr(" C ","00010212",4)
blocks(3) = Arr(" D ","02101112",4)
blocks(4) = Arr(" E ","01112021",4)
blocks(5) = Arr(" F ","001011",3)
blocks(6) = Arr(" G ","011011",3)


n = 0
Test 0

Sub Test (num)
  Dim i,j
  i=0:j=0
  For i = 0 to 4
	For j = 0 to 4
	  If Put(blocks(num),i,j) Then
	      Fit(num,0)=i:Fit(num,1)=j
'		  Draw S
		If num = 6 Then
		  Draw S
		  n = n+1
		  Tak num
		Else
		  Test (num+1)
		End If
	  ElseIf i=4 and j=4 and num>0 Then
	    Tak num-1
	  End If
	Next
  Next

  If num=0 Then MsgBox "Completed with "& n &" results."
End Sub

Function Arr (sym,corr,dm)
  Dim i
  ReDim arr_t(dm,2)
  arr_t(0,0) = sym
  For i = 1 to dm
    arr_t(i,0)= Cint(Mid(corr,i*2-1,1))
    arr_t(i,1)= Cint(Mid(corr,i*2,1))
  Next
  Arr = arr_t  
End Function

Function Put (block,x,y)
  Dim i
  For i = 1 to UBound(block,1)
'   msgbox block(0,0) &"("& i &") is ("& x+block(i,0) &","& y+block(i,1) &")."
    If (x+block(i,0) > 4) or (y+block(i,1) > 4) Then
	  Put = False
'	  msgbox block(0,0) &"("& i &") is ("& x+block(i,0) &","& y+block(i,1) &")."
	  Exit Function
	ElseIf (S(x+block(i,0),y+block(i,1)) <> " - ")  Then
	  Put = False
'	  msgbox "S("& x+block(i,0) &","& y+block(i,1) &") is occupied."
	  Exit Function
	End If
  Next
  For i = 1 to UBound(block,1)
    S(x+block(i,0),y+block(i,1)) = block(0,0)
  Next
  Put = True
End Function

Sub Tak (num1)
  Dim i
  x=Fit(num1,0):y=Fit(num1,1)
  For i = 1 to UBound(blocks(num1),1)
    S(x+blocks(num1)(i,0),y+blocks(num1)(i,1)) = " - "
  Next
  Fit(num1,0)=-1:Fit(num1,1)=-1
End Sub


Sub Draw (s)
  Dim i,j
  str = "   0  1  2  3  4" &vbCrLf&_
        "  --------------"
  For i = 0 to 4
    str = str &vbCrLf& i &"|"
    For j = 0 to 4
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
  ie.Width     = 350
  ie.Height    = 400

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
