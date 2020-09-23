VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Function Drawer 1.0 by SMA Soft"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program draws the chart of the math functions
'I have written this program when I was 15 years old
'This Program is downloaded from http://www.partiasoft.com
'For contacting me, send email to sma_soft@yahoo.com
Const P = 3.14159265358979
Public Beg, Ende, YMin, YMAx, XA, YA

Private Sub Form_Click()
Cls
DrawFN
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then Call Form_Click
End Sub

Private Sub Form_Load()
MsgBox "Click on the form or press F5 to draw the chart"
End Sub

Function G(X)
On Error Resume Next
G = Sin(X) * 5
End Function

Function F(X)
On Error Resume Next
F = X * Sin(X) 'You define your own math function here
End Function

Function H(X)
On Error Resume Next
H = F(G(X))
End Function

Function J(X)
J = Int(X)
End Function
Function Q(X)
Q = Abs(X)
End Function
Function R(X)
R = Sqr(X)
End Function
Sub DrawFN()
On Error Resume Next
Beg = -10
Ende = 10
YMin = Beg
YMAx = Ende
Me.Scale (Beg, YMin)-(Ende, YMAx)
Me.DrawWidth = 1
For n = Int(Beg) To Int(Ende)
    Line (n, YMin)-(n, YMAx), QBColor(8)
Next n
For n = Int(YMin) To Int(YMAx)
    Line (Beg, n)-(Ende, n), QBColor(8)
Next n
Me.DrawWidth = 1
mojo = F(3E+35)
Line (Beg, mojo)-(Ende, mojo), QBColor(9)
mojo = F(-3E+35)
Line (Beg, mojo)-(Ende, mojo), QBColor(9)

Line (Beg, 0)-(Ende, 0), QBColor(15)
Line (0, YMin)-(0, YMAx), QBColor(15)

For X = Beg To Ende Step 0.0025
'Change this number compatible with your CPU Speed.
'Choose a smaller number if you want the chart to be drawn better.
'But if you choose a small value, the chart will be drawn slower
    y = F(X)
    y = -y
    PSet (X, y), vbRed
    y = -H(X)
    ''PSet (X, y), vbBlue
    If Err Then
        Line (X, YMin)-(X, YMAx), vbGreen
    End If
Next X
End Sub
