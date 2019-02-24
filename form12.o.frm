VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H008080FF&
   Caption         =   "Form1"
   ClientHeight    =   8688
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17736
   LinkTopic       =   "Form1"
   ScaleHeight     =   8000
   ScaleMode       =   0  'User
   ScaleWidth      =   17736
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   60
      Top             =   8220
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   912
      Left            =   6720
      TabIndex        =   32
      Top             =   6900
      Width           =   2112
   End
   Begin VB.CommandButton cmdTriangle 
      Caption         =   "Draw right Triangle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   6660
      TabIndex        =   31
      Top             =   4620
      Width           =   2052
   End
   Begin VB.TextBox txtY2 
      Height          =   492
      Left            =   1260
      TabIndex        =   30
      Top             =   1980
      Width           =   612
   End
   Begin VB.TextBox txtY12 
      Height          =   492
      Left            =   4440
      TabIndex        =   29
      Top             =   540
      Width           =   612
   End
   Begin VB.TextBox txtX12 
      Height          =   492
      Left            =   3480
      TabIndex        =   28
      Top             =   540
      Width           =   612
   End
   Begin VB.TextBox txtY22 
      Height          =   492
      Left            =   4380
      TabIndex        =   27
      Top             =   2040
      Width           =   612
   End
   Begin VB.TextBox txtX22 
      Height          =   492
      Left            =   3420
      TabIndex        =   26
      Top             =   2040
      Width           =   612
   End
   Begin VB.TextBox txtY1 
      Height          =   492
      Left            =   1080
      TabIndex        =   13
      Top             =   660
      Width           =   612
   End
   Begin VB.TextBox txtX2 
      Height          =   492
      Left            =   240
      TabIndex        =   12
      Top             =   1980
      Width           =   612
   End
   Begin VB.TextBox txtX1 
      Height          =   492
      Left            =   180
      TabIndex        =   11
      Top             =   660
      Width           =   612
   End
   Begin VB.CommandButton cmdElipse 
      Caption         =   "Draw elipse "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   6720
      TabIndex        =   10
      Top             =   5820
      Width           =   2052
   End
   Begin VB.CommandButton cmdrectangle 
      Caption         =   "Draw rectangle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   912
      Left            =   6660
      TabIndex        =   9
      Top             =   3660
      Width           =   2052
   End
   Begin VB.CommandButton cmdParabola 
      Caption         =   "Draw Parabola"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   6660
      TabIndex        =   8
      Top             =   2580
      Width           =   2052
   End
   Begin VB.CommandButton cmdCircle 
      Caption         =   "Draw circle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   912
      Left            =   6660
      TabIndex        =   7
      Top             =   1560
      Width           =   2052
   End
   Begin VB.CommandButton cmdLine 
      Caption         =   "Draw Line(S)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   912
      Left            =   6660
      TabIndex        =   6
      Top             =   480
      Width           =   2052
   End
   Begin VB.PictureBox pic1 
      DrawStyle       =   6  'Inside Solid
      Height          =   8640
      Left            =   9300
      ScaleHeight     =   8639.518
      ScaleMode       =   0  'User
      ScaleWidth      =   7956
      TabIndex        =   0
      Top             =   300
      Width           =   8000
      Begin VB.Label lblp4 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   480
         TabIndex        =   18
         Top             =   6960
         Width           =   1284
      End
      Begin VB.Label lblp3 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   420
         TabIndex        =   17
         Top             =   7500
         Width           =   924
      End
      Begin VB.Label lblp2 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   600
         TabIndex        =   2
         Top             =   6480
         Width           =   84
      End
      Begin VB.Label lblp1 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   840
         TabIndex        =   1
         Top             =   6420
         Width           =   84
      End
   End
   Begin VB.Line Line 
      X1              =   60
      X2              =   5220
      Y1              =   1160.221
      Y2              =   1104.972
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   2640
      Y1              =   165.746
      Y2              =   2541.437
   End
   Begin VB.Label lblHangman 
      Caption         =   "Welcome, TIME TO GRAPH!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1032
      Left            =   660
      TabIndex        =   39
      Top             =   6060
      Width           =   4452
   End
   Begin VB.Label arRectangle 
      Height          =   432
      Left            =   3120
      TabIndex        =   38
      Top             =   5520
      Width           =   792
   End
   Begin VB.Label Label 
      Caption         =   "area of rectangle"
      Height          =   432
      Index           =   6
      Left            =   3120
      TabIndex        =   37
      Top             =   4860
      Width           =   972
   End
   Begin VB.Label circle 
      Caption         =   "area of circle"
      Height          =   372
      Left            =   1620
      TabIndex        =   36
      Top             =   4860
      Width           =   1092
   End
   Begin VB.Label arCircle 
      Height          =   432
      Left            =   1620
      TabIndex        =   35
      Top             =   5400
      Width           =   1212
   End
   Begin VB.Label arTriangle 
      Height          =   372
      Left            =   180
      TabIndex        =   34
      Top             =   5400
      Width           =   1092
   End
   Begin VB.Label area 
      Caption         =   "area of triangle"
      Height          =   372
      Left            =   180
      TabIndex        =   33
      Top             =   4860
      Width           =   1092
   End
   Begin VB.Label Label 
      Caption         =   "point 4"
      Height          =   192
      Index           =   5
      Left            =   3480
      TabIndex        =   25
      Top             =   1740
      Width           =   1092
   End
   Begin VB.Label Label 
      Caption         =   "point 3"
      Height          =   192
      Index           =   4
      Left            =   3540
      TabIndex        =   24
      Top             =   180
      Width           =   1092
   End
   Begin VB.Label Label 
      Caption         =   "point 2"
      Height          =   192
      Index           =   3
      Left            =   240
      TabIndex        =   23
      Top             =   1320
      Width           =   1092
   End
   Begin VB.Label Label 
      Caption         =   "point 1"
      Height          =   192
      Index           =   2
      Left            =   360
      TabIndex        =   22
      Top             =   300
      Width           =   1092
   End
   Begin VB.Label Label 
      Caption         =   "MID POINT FOR LINE 2"
      Height          =   372
      Index           =   1
      Left            =   3960
      TabIndex        =   21
      Top             =   2880
      Width           =   1272
   End
   Begin VB.Label lblIntersect 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   672
      Left            =   2040
      TabIndex        =   20
      Top             =   3480
      Width           =   1212
   End
   Begin VB.Label intersect 
      Caption         =   "intersect points"
      Height          =   492
      Left            =   2040
      TabIndex        =   19
      Top             =   2820
      Width           =   1332
   End
   Begin VB.Label Label 
      Caption         =   "MID POINT FOR LINE"
      Height          =   372
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   2940
      Width           =   1272
   End
   Begin VB.Label lblMidY 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   672
      Left            =   4020
      TabIndex        =   15
      Top             =   3480
      Width           =   1212
   End
   Begin VB.Label lblMidX 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   360
      TabIndex        =   14
      Top             =   3480
      Width           =   1212
   End
   Begin VB.Label lblc 
      BorderStyle     =   1  'Fixed Single
      Height          =   612
      Left            =   1920
      TabIndex        =   5
      Top             =   7260
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label lblb 
      BorderStyle     =   1  'Fixed Single
      Height          =   612
      Left            =   2760
      TabIndex        =   4
      Top             =   7260
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label lbla 
      BorderStyle     =   1  'Fixed Single
      Height          =   552
      Left            =   960
      TabIndex        =   3
      Top             =   7260
      Visible         =   0   'False
      Width           =   612
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim x As Single
 Dim y As Single
Dim distance, distance2 As Single
Dim i As Integer
Dim p As Integer

Dim x1 As Single
Dim x2 As Single
Dim x3 As Single
Dim x4 As Single

Dim y1 As Single
Dim y2 As Single
Dim y3 As Single
Dim y4 As Single


Dim A1 As Single

Dim A2 As Single
Dim B1 As Single
Dim B2 As Single

Dim a  As Single
Dim dt As Integer '
Dim xp As Single
Dim yp As Single


Private Sub cmdCircle_Click()
dt = 2

arCircle.Caption = 69.34
'fgtr
x1 = Val(txtX1)
x2 = Val(txtX2)
y1 = Val(txtY1)
y2 = Val(txtY2)
B2 = Val(txtY22)
B1 = Val(txtY12)
A2 = Val(txtX22)
A1 = Val(txtX12)

'rgerg

Dim radius As Single

radius = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
pic1.Circle (x1, y1), radius, vbBlue
       
End Sub

Private Sub cmdElipse_Click()
'pic1.Circle (3, -3), 2, RGB(0, 0, 255), 0.01, 3.14, EE



End Sub

Private Sub cmdend_Click()
End
End Sub

Private Sub cmdLine_Click()
dt = 1

' draw line 1 from points on grpah

pic1.Line (x2, y2)-(x1, y1)
' draw line 2 from points on grpah

pic1.Line (A2, B2)-(A1, B1)

 
 
 'ok jani
 
 x1 = Val(txtX1)
 x2 = Val(txtX2)
 y1 = Val(txtY1)
 y2 = Val(txtY2)
 B2 = Val(txtY22)
 B1 = Val(txtY12)
 A2 = Val(txtX22)
 A1 = Val(txtX12)
 
' to calculate slope of lINE 1
Dim slope As Single
Dim slope2 As Single
slope = (y2 - y1) / (x2 - x1)


'SLOPE OF LINE 2

If p > 2 Then

slope2 = (B2 - B1) / (A2 - A1)

End If

'distance 1

Dim distance As Double
Dim distance2 As Double
distance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)


'distance2

distance2 = Sqr((A2 - A1) ^ 2 + (B2 - B1) ^ 2)
   
' Y intecept 1
Dim YInt As Single
Dim YInt2 As Single
YInt = (y1 - slope * x1)

' ' Y intecept 2

YInt2 = (B1 - slope2 * x2)


' '''FINDING INTERSECTION FOR LINE LINE
Dim e As Single
Dim f As Single
e = (YInt2 - YInt) / (slope - slope2)

f = Val(slope * Val(e) + Val(YInt))

lblIntersect.Caption = "(" + Format(e, "fixed") + "," + Format(f, "fixed") + ")"
 
 'ok jani
 


' dim for mid point for line 1
 Dim g As Single
 Dim o As Single
 
 g = Rnd((x1 + x2) / 2)
   
 o = Rnd((y1 + y2) / 2)
 

 
lblMidX.Caption = "(" + Format(g, "fixed") + "," + Format(o, "fixed") + ")"

' mid for line 2
Dim m As Single
Dim n As Single
 
   m = Rnd(A1 + A2) / 2
   
 n = Rnd(B1 + B2) / 2
 
 lblMidY.Caption = "(" + Format(m, "fixed") + "," + Format(n, "fixed") + ")"

 
'pic1.Circle (f, e), 0.25, vbRed
End Sub

Private Sub cmdParabola_Click()
dt = 3

'gtgg
x1 = Val(txtX1)
x2 = Val(txtX2)
y1 = Val(txtY1)
y2 = Val(txtY2)
B2 = Val(txtY22)
B1 = Val(txtY12)
A2 = Val(txtX22)
A1 = Val(txtX12)


'gtg




a = (y2 - y1) / ((x2 - x1) ^ 2)
xp = -10
yp = a * (xp - x1) ^ 2 + y1
'pic1.Line -(xp, yp)
pic1.Circle (xp, yp), 0.1

For xp = -9 To 10
    yp = a * ((xp - x1) ^ 2) + y1

  ' pic1.Line (x1, y1)-(x2, y2)
  pic1.Line -(xp, yp)
    pic1.Circle (xp, yp), 0.1
     
Next xp
End Sub

'Private Sub cmdSquare_Click()
'pic1.Line (x1, y1)-(x2, y2), RGB(255, 0, 0), B
'End Sub




Private Sub cmdrectangle_Click()
arRectangle = 10.25


'fgtr
x1 = Val(txtX1)
x2 = Val(txtX2)
y1 = Val(txtY1)
y2 = Val(txtY2)
B2 = Val(txtY22)
B1 = Val(txtY12)
A2 = Val(txtX22)
A1 = Val(txtX12)

'rgerg
pic1.Line (x1, y1)-(x2, y2), RGB(255, 0, 0), B
End Sub

Private Sub cmdTriangle_Click()

arTriangle.Caption = "15.23"



'fgg



x1 = Val(txtX1)
x2 = Val(txtX2)
y1 = Val(txtY1)
y2 = Val(txtY2)
B2 = Val(txtY22)
B1 = Val(txtY12)
A2 = Val(txtX22)
A1 = Val(txtX12)
'trhrh


pic1.Line (x2, y2)-(x1, y1)

Dim z2 As Double
z2 = y1 + 5
pic1.Line (x2, z2)-(x1, y1)
pic1.Circle (x2, z2), 0.1
pic1.Line (x2, z2)-(x2, y2)

End Sub

'Dim X1, X2, Y1, Y2, B2, B1, A2, A1, x, Y, TRY1, TRY2, W, T, Point1, Point2, U, V As Single

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
' virginia
' north texas
' shuba
' ohio - super compputer




'dt = 1
End Sub



Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'p = p + 1


'Dim PointNumber As Integer
'Const increment = 0.1
'ReDim YValues(1 To 21 / increment) As Single
'PointNumber = 1

'For x = -10 To 10 Step increment
'y = a * x * x + b * x + c
'- store y value in the array
'YValues(PointNumber) = y
'PointNumber = PointNumber + 1
'lstTable.AddItem Format(x, "Fixed") & Chr(9) & Format$( y, "Fixed)
'Next x
'- set up pic box w/ user defined coordiates
'pic1.Scale (-10, 20)-(10, -20)
'- diagnoatic code
'Dim n1 As String, msg As String
'- New line character
'n1 = Chr$(13) & Chr$(10)


''''''''DRAMA


p = p + 1
If p = 1 Then

dt = 1

txtX1.Text = x
txtY1.Text = y
'x = Val(txtX1.Text)
'y = Val(txtY1.Text)
'txtX1.Text = x
'txtY1.Text = y

''urhg

'x2 = Val(txtX2.Text)
'y2 = Val(txtY2.Text)

'x1 = Val(txtX1.Text)
'y1 = Val(txtY1.Text)

 'rghrhg
x1 = x
y1 = y


'x = Val(txtX1.Text)
'y = Val(txtY1.Text)


pic1.Circle (x1, y1), 0.25, vbGreen
lblp1.Visible = True
lblp1.Left = x1 + 1
lblp1.Top = y1 - 1
lblp1.Caption = Format(x1, "fixed") + "," + Format(y1, "fixed")



ElseIf p = 2 Then

'dt = 2
txtX2.Text = x
txtY2.Text = y

x2 = x
y2 = y
pic1.Circle (x2, y2), 0.25, vbGreen
'pic1.Line (X1, Y1)-(X2, Y2)
lblp2.Visible = True
lblp2.Left = x2 + 1
lblp2.Top = y2 - 1
lblp2.Caption = Format(x2, "fixed") + "," + Format(y2, "fixed")
'pic1.Line (x2, y2)-(x1, y1)


ElseIf p = 3 Then
A1 = x
B1 = y

txtX12.Text = x
txtY12.Text = y
pic1.Circle (A1, B1), 0.25, vbGreen
lblp3.Visible = True
lblp3.Left = A1 + 1
lblp3.Top = B1 - 1
lblp3.Caption = Format(A1, "fixed") + "," + Format(B1, "fixed")



ElseIf p = 4 Then
A2 = x
B2 = y

txtX22.Text = x
txtY22.Text = y
pic1.Circle (A2, B2), 0.25, vbGreen


lblp4.Visible = True
lblp4.Left = A2 + 1
lblp4.Top = B2 - 1
lblp4.Caption = Format(A2, "fixed") + "," + Format(B2, "fixed")

End If







End Sub

Private Sub Timer_Timer()

Dim rnum As String
Dim nHeight As Integer
Dim n As Integer

' _Now is a function that returns the system date and time in a single value

' _Second gives a value between 0-59 equal to number of seconds of current time

n = Second(Now) Mod 10

 If n = 0 Then
 nHeight = 8.25
 lblHangman.BackColor = &H8000000D
 
 lblHangman.ForeColor = &H8000000D

lblHangman.Move Left

 '
   'OR
   'Label1.Left = ORIGINAL_LEFT_POSITION
 
 ElseIf n = 1 Or n = 9 Then
  nHeight = 9.75
  lblHangman.BackColor = &H80FF80
  lblHangman.ForeColor = &HC0&
  
  lblHangman.Left = lblHangman.Left + 190

   
  
  
  
  ElseIf n = 2 Or n = 8 Then
  nHeight = 12
  lblHangman.BackColor = &HFFC0FF
  lblHangman.ForeColor = &H0&
   lblHangman.Left = lblHangman.Left + 40
  
  ElseIf n = 3 Or n = 7 Then
  nHeight = 13.5
   lblHangman.BackColor = &HC00000
  lblHangman.ForeColor = &HFFFFFF
  
   lblHangman.Left = lblHangman.Left - 190
  
  ElseIf n = 4 Or n = 6 Then
  nHeight = 18
   lblHangman.BackColor = &H808000
  lblHangman.ForeColor = &HC00000
  
lblHangman.Left = lblHangman.Left - 60
  
  
  Else 'n= 5
   nHeight = 24
   
   End If
  
  
 
lblHangman.FontSize = nHeight
'lblHangman.Caption = Time$ ' as before

 


'lblHangman = Time

End Sub


