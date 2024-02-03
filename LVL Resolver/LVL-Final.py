import numpy as np

S = np.array([['-']*9]*9)
Fit = np.array([[-1]*2]*21)
#print(S)
#print(Fit)

blocks = []
blocks.append(['A',(0,0),(0,1),(1,0),(2,0),(2,1)])
blocks.append(['B',(0,1),(1,0),(1,1),(1,2),(2,1)])
blocks.append(['C',(0,0),(0,1),(1,0),(2,0)])
blocks.append(['D',(0,0),(0,1),(1,0),(2,0)])
blocks.append(['E',(0,0),(0,1),(1,1),(2,1)])
blocks.append(['F',(0,0)])
blocks.append(['G',(0,1),(1,1),(2,0),(2,1)])
blocks.append(['H',(0,1),(1,0),(1,1),(1,2)])
blocks.append(['I',(0,0),(0,1)])
blocks.append(['J',(0,2),(1,0),(1,1),(1,2)])
blocks.append(['K',(0,0),(0,1),(0,2)])
blocks.append(['L',(0,0),(0,1),(0,2)])
blocks.append(['M',(0,0),(0,1),(1,0),(2,0)])
blocks.append(['N',(0,0),(0,1),(0,2),(1,0)])
blocks.append(['O',(0,1),(1,0),(1,1),(1,2)])
blocks.append(['P',(0,0),(0,1),(1,1),(2,1)])
blocks.append(['Q',(0,0),(0,2),(1,0),(1,1),(1,2)])
blocks.append(['R',(0,2),(1,0),(1,1),(1,2)])
blocks.append(['S',(0,0),(1,0)])
blocks.append(['T',(0,1),(1,0),(1,1)])
blocks.append(['U',(0,0),(0,1),(1,0),(1,1)])
blocks.append(['V',(0,0),(0,1),(1,0),(1,1)])
#print(blocks)

#Verify

n = 0; m = 0
Test 0

def Verify():
  for i in range(21):
    if Put(i,0,0):
	  Draw S
	Tak i


def Test(num):
  Dim i,j
  For i = 0 to 8
	For j = 0 to 8
	  m=m+1
	  If Put(num,i,j) Then
		If (m mod 10000000)=0 Then Draw S
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
	Next
  Next

  If num=0 Then MsgBox "Completed with "& n &" results."
End Sub


def Put(num,x,y):
  Dim i
  For i = 1 to UBound(blocks(num),1)
'   msgbox blocks(num)(0,0) &"("& i &") is ("& x+blocks(num)(i,0) &_
'          ","& y+blocks(num)(i,1) &")."
    If (x+blocks(num)(i,0) > 8) or (y+blocks(num)(i,1) > 8) Then
	  Put = False
'	  msgbox blocks(num)(0,0) &"("& i &") is ("& x+blocks(num)(i,0) &_
'            ","& y+blocks(num)(i,1) &")."
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
  return True


def Tak(num):
  Dim i
  x=Fit(num,0):y=Fit(num,1)
  For i = 1 to UBound(blocks(num),1)
    S(x+blocks(num)(i,0),y+blocks(num)(i,1)) = " - "
  Next
  Fit(num,0)=-1:Fit(num,1)=-1


def Draw(s):
  Dim i,j
  str = "   0  1  2  3  4  5  6  7  8" &vbCrLf&_
        "  --------------------------"
  For i = 0 to 8
    str = str &vbCrLf& i &"|"
    For j = 0 to 8
	  str = str & S(i,j)
	Next
  Next
  CusMsgBox m &vbCrLf&str

