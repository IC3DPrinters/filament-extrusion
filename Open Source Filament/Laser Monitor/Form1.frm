VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   DrawWidth       =   20
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   11280
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   6600
      Width           =   9975
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Text            =   "0"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Text            =   "0"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Text            =   "0"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Text            =   "0"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Text            =   "0"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Text            =   "0"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Text            =   "0"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Text            =   "0"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   6720
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "AddPoint"
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear"
         Height          =   375
         Left            =   4200
         TabIndex        =   2
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "d"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "dtm"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "fi"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "hte"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "hre"
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "hm"
         Height          =   255
         Left            =   1800
         TabIndex        =   16
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "tou"
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "p"
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label ResultLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4080
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   480
      ScaleHeight     =   6345
      ScaleWidth      =   10665
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   10695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Me.ScaleMode = vbPixels
Me.AutoRedraw = True
Me.Show
Me.Caption = "Graph"
Picture1.Left = 0
Picture1.Top = 0
Picture1.Width = Me.ScaleWidth
Frame1.Caption = "Input Data"
Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
Frame1.Top = Me.ScaleHeight - Frame1.Height
Picture1.Height = Frame1.Top
Picture1.Scale (-30, 720)-(520, -120)
Picture1.BackColor = &HFFFFFF
Picture1.ForeColor = 0
Picture1.DrawWidth = 1
Picture1.FontName = "Arial"
Picture1.FontSize = Picture1.ScaleX(9, vbPixels, vbPoints)

Picture1.Cls
Draw_Axes
'Draw_Curve


End Sub

Sub Draw_Axes()
Dim i As Integer
Picture1.AutoRedraw = True
For i = 0 To 500 Step 20
    Picture1.Line (i, 0)-(i, 700), &HD0D0D0 'draw y axis
    If i = 0 Then Picture1.Line (i, 0)-(i, 700), &H0
    Picture1.Line (i, 0)-(i, -10), &H0 'hashlines
    Picture1.CurrentX = i - Picture1.TextWidth(Format(Exp((i / 100 - 3) * 2.302585), "#.0##")) / 2
    Picture1.CurrentY = -20
    Picture1.Print Format(Exp((i / 100 - 3) * 2.302585), "#.0##")
Next

For i = 0 To 700 Step 50
    Picture1.Line (0, i)-(500, i), &HD0D0D0
    If i = 0 Then Picture1.Line (0, i)-(500, i), &H0
    Picture1.Line (0, i)-(-3, i), &H0
    Picture1.CurrentX = -5 - Picture1.TextWidth(Trim(Val(i - 100)))
    Picture1.CurrentY = i - Picture1.TextHeight(Trim(Val(i - 100))) / 2
    Picture1.Print Trim(Val(i - 100))
Next
Picture1.AutoRedraw = False
End Sub


Private Sub Command1_Click()
Dim result As Double
Dim i As Integer
'=============
Dim d As Double           'user input distance
Dim dtm As Double         'user input distance(overland+inland)
Dim dlm As Double
Dim fi As Double          'user input path centre latitude in degrees
Dim hte As Double         'user input effective height of the transmitting antenna
Dim hre As Double         'user input effective height of the receiver antenna
Dim hm As Double          'user input height terrain roughness
Dim tou As Double         'user input if pol is H then its 0, if V then 90
Dim u1 As Double
Dim u2 As Double
Dim u3 As Double
Dim u4 As Double
Dim b0 As Double
Dim alpha As Double
Dim beta As Double
Dim di As Double
Dim t As Double
Dim p As Double           'user input small percentage from 0.001 to 50%
Dim e As Double
Dim pol As String
Dim h As Double
Dim v As Double
Dim ae As Double
Dim result1 As Double
d = Val(Text1.Text)
dtm = Val(Text2.Text)
fi = Val(Text3.Text)
hte = Val(Text4.Text)
hre = Val(Text5.Text)
hm = Val(Text6.Text)
pol = Val(Text7.Text)
p = Val(Text8.Text)
ae = 6370
'defining the first equation as tou
If LCase(Text7) = "h" Then
   tou = 0
   u1 = (((10 ^ (-dtm / 16 - (6.6 * tou))) + ((10 ^ -(0.496 + 0.354 * tou)) ^ 5))) ^ 0.2
ElseIf LCase(Text7) = "v" Then
   tou = 90
    u1 = ((10 ^ (-dtm / 16 - (6.6 * tou)) + (10 ^ -(0.496 + 0.354 * tou)) ^ 5)) ^ 0.2
Else
    MsgBox "pol value must be h or v"
    Exit Sub
End If
If fi <= 70 Then
       u4 = (10 ^ ((-0.935 + 0.0176 * fi) * Log(u1) / Log(10)))
ElseIf fi > 70 Then
       u4 = 10 ^ (0.3 * ((Log(u1) / Log(10))))
End If
'now defining the b0
If fi <= 70 Then
   b0 = (10 ^ (-0.015 * fi + 1.67)) * u1 * u4
ElseIf fi > 70 Then
   b0 = 4.17 * u1 * u4
End If
'now defining u2 and u3
alpha = -0.6 - 3.5 * (10 ^ -9) * (d ^ 3.1) * tou
u2 = ((500 / ae) * ((d ^ 2) / ((Sqr(hte) + Sqr(hre)) ^ 2)) ^ alpha)
'now u3
If hm <= 10 Then
   u3 = 1
ElseIf hm > 10 Then
   u3 = Exp((-4.6 * (10 ^ -5) * (hm - 10) * (43)))
End If
'now defining beta
beta = b0 * u2 * u3
'now defining t
t = (1.076 / ((2.0058 - Log(beta)) ^ 1.012)) * Exp(-(9.51 - 4.8 * Log(beta) + 0.198 * Log(beta) ^ 2) * (10 ^ -6) * (d ^ 1.13))
'the final equation for time percentage (cumulative distribution)
result = -12 + ((1.2 + 3.7 * 10 ^ -3 * d * Log(p / beta))) + (12 * (p / beta) ^ t)
result = result
ResultLabel.Caption = CStr(result) + "dB"
List1.AddItem result & " (" & p & ")"

'translate scale
Picture1.Cls
For i = 0 To List1.ListCount - 1
Picture1.Circle (X2Log(Val(Mid(List1.List(i), InStr(List1.List(i), "(") + 1))), Val(List1.List(i)) + 100), 1, vbGreen
Next
End Sub

Function X2Log(x As Double) As Double

X2Log = (Log(x) / 2.302585 + 3) * 100
End Function
