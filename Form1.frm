VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8688
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17736
   LinkTopic       =   "Form1"
   ScaleHeight     =   8688
   ScaleWidth      =   17736
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstTable 
      Height          =   2928
      Left            =   180
      TabIndex        =   1
      Top             =   660
      Width           =   1512
   End
   Begin VB.PictureBox pic1 
      Height          =   8592
      Left            =   2400
      ScaleHeight     =   8544
      ScaleWidth      =   15144
      TabIndex        =   0
      Top             =   0
      Width           =   15192
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim slope As Single
Dim slope2 As Single

Dim extendedpoint1 As Single


Dim distance, distance2 As Single
Dim i As Integer
Dim p As Integer



'Dim X1, X2, Y1, Y2, B2, B1, A2, A1, x, Y, TRY1, TRY2, W, T, Point1, Point2, U, V As Single

Dim g, o As Single



Private Sub Form_Activate()
pic1.Scale (-10, 10)-(10, -10)
pic1.Line (-10, 0)-(10, 0), vbRed
pic1.Line (0, -10)-(0, 10), vbBlue

For i = -10 To 10
    pic1.Line (i, 0.5)-(i, -0.5), vbBlue
    pic1.Line (0.5, i)-(-0.5, i), vbRed
    
Next i
    
End Sub

Private Sub Form_Load()
p = 0
End Sub

Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'p = p + 1


Dim PointNumber As Integer
Const increment = 0.1
ReDim YValues(1 To 21 / increment) As Single
PointNumber = 1

For x = -10 To 10 Step increment
y = a * x * x + b * x + c
'- store y value in the array
YValues(PointNumber) = y
PointNumber = PointNumber + 1
lstTable.AddItem Format(x, "Fixed") & Chr(9) & Format$( y, "Fixed)
Next x
'- set up pic box w/ user defined coordiates
pic1.Scale (-10, 20)-(10, -20)
'- diagnoatic code
Dim n1 As String, msg As String
'- New line character
n1 = Chr$(13) & Chr$(10)





End Sub
