VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Laser Measurement"
   ClientHeight    =   10710
   ClientLeft      =   11745
   ClientTop       =   5205
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic2Xaxis 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9100
      ScaleHeight     =   615
      ScaleMode       =   0  'User
      ScaleWidth      =   9700
      TabIndex        =   20
      Top             =   8280
      Width           =   9700
   End
   Begin VB.PictureBox pic2Yaxis 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8400
      Left            =   8600
      ScaleHeight     =   8400
      ScaleMode       =   0  'User
      ScaleWidth      =   900
      TabIndex        =   21
      Top             =   200
      Width           =   900
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   10440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pic3Xaxis 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2960
      ScaleHeight     =   615
      ScaleMode       =   0  'User
      ScaleWidth      =   5600
      TabIndex        =   22
      Top             =   8160
      Width           =   5600
   End
   Begin VB.PictureBox pic3Yaxis 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   2640
      ScaleHeight     =   2700
      ScaleMode       =   0  'User
      ScaleWidth      =   761.194
      TabIndex        =   23
      Top             =   5628
      Width           =   765
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FF8080&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6840
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0FFC0&
      ClipControls    =   0   'False
      Height          =   2300
      Left            =   3360
      ScaleHeight     =   2300
      ScaleMode       =   0  'User
      ScaleWidth      =   3000
      TabIndex        =   18
      Top             =   5880
      Width           =   4800
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFC0&
      ClipControls    =   0   'False
      Height          =   7800
      Left            =   9500
      ScaleHeight     =   8000
      ScaleMode       =   0  'User
      ScaleWidth      =   3000
      TabIndex        =   17
      Top             =   500
      Width           =   8900
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      Caption         =   "Nominal Diameter (mm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   14
      Top             =   3240
      Visible         =   0   'False
      Width           =   2535
      Begin VB.OptionButton opt29 
         BackColor       =   &H0080FF80&
         Caption         =   "2.9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   615
      End
      Begin VB.OptionButton opt2875 
         BackColor       =   &H0080FF80&
         Caption         =   "2.875"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton opt28 
         BackColor       =   &H0080FF80&
         Caption         =   "2.8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton opt175 
         BackColor       =   &H0080FF80&
         Caption         =   "1.75"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.ComboBox cboTolerance 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   12
      Text            =   "0.1"
      Top             =   5040
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Audible Alarm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   1815
      Begin VB.OptionButton optOff 
         BackColor       =   &H00FFFF80&
         Caption         =   "Off"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton optOn 
         BackColor       =   &H00FFFF80&
         Caption         =   "On"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   3015
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   8
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtDiameter 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtOvality 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtY 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtX 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   2880
      Top             =   10440
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2160
      Top             =   10320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
      BaudRate        =   2400
      InputMode       =   1
   End
   Begin VB.Label lblSelect1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select COM port for laser"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   240
      TabIndex        =   29
      Top             =   8880
      Width           =   4935
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exceed Tolerance Alarm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3825
      TabIndex        =   28
      Top             =   240
      Width           =   2190
   End
   Begin VB.Shape shpAlarm 
      BackColor       =   &H000000FF&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ovality (%)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5265
      TabIndex        =   27
      Top             =   5520
      Width           =   990
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Diameter (mm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   13225
      TabIndex        =   26
      Top             =   120
      Width           =   1410
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tolerance (mm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   5100
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Diameter (mm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3360
      TabIndex        =   7
      Top             =   2400
      Width           =   2220
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ovality (%)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      TabIndex        =   6
      Top             =   3120
      Width           =   1590
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Y Value (mm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      TabIndex        =   3
      Top             =   4920
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "X Value (mm)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      TabIndex        =   2
      Top             =   4200
      Width           =   1995
   End
   Begin VB.Menu COMPort 
      Caption         =   "COM Port"
      Begin VB.Menu Laser 
         Caption         =   "Laser"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gstrBuffer2 As String
Dim gstrBuffer1(20) As Byte
Dim recordFlag As Boolean
Dim transferFlag As Boolean
Dim rcvFlag As Boolean
Dim intTimer1 As Integer
Dim ptr As Integer
Dim xDiam As Single
Dim yDiam As Single
Dim blnFlag As Boolean
Dim startFlag As Boolean
Dim intCounter As Long
Dim nominalTolerance As Single
Dim stringLengthFlag As Boolean
Dim drawDiameterBuffer(3000) As Single
Dim drawOvalityBuffer(3000) As Single
Dim drawColorBuffer(3000) As Boolean
Dim drawDiameter As Single
Dim drawOvality As Single
Dim intPoint As Integer
Dim startTime As Date
Dim endTime As Date
Dim drawFlag As Boolean
Dim speedBuffer(1000) As Single
Dim speedPtr As Integer
Dim speedAverage As Single
'1.393225805 4.714500753
'0.153763441 0.207526882

Const m = 0.126628628

Const b = 6.802436136


'Const m = 0.0946
'Const b = 0.0541


Private Sub cboTolerance_Click()
    DrawAxis
    

End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub



Private Sub cmdStart_Click()
    Dim strFileName As String
    
    If cmdStart.Caption = "Start" Then
        cmdStart.Caption = "Stop"
        cmdStart.BackColor = &HFF&
        
        ' Set CancelError is True
        CommonDialog1.CancelError = True
        On Error GoTo ErrHandler
        
        CommonDialog1.FileName = "Data" & Format(Now, "mmddhhmmss") & ".xls"
        ' Set filters.
        CommonDialog1.Filter = "Data Files (*.xls) | *.xls"
        ' Specify default filter.
        CommonDialog1.FilterIndex = 0
        
        ' Display the Open dialog box.
        CommonDialog1.ShowOpen
        ' Call the open file procedure.
        Open (CommonDialog1.FileName) For Output As #1
        
        'strFileName = "Data" & Format(Now, "mmddhhmmss") & ".xls"
        
        'Open strFileName For Output As 1
        
        Print #1, "Time" & vbTab & "X" & vbTab & "Y" & vbTab & "Diameter" & vbTab & "%Ovality"
        startFlag = True
        
        startTime = Now
        startTime = Format(Now, "hh:mm")
        
        DrawAxis
        
    Else
        startFlag = False

        cmdStart.Caption = "Start"
        cmdStart.BackColor = &HFF00&
        Close 1
        
    
    End If
    
    Exit Sub
    
ErrHandler:
    Resume Next
    
    'User pressed the Cancel button
    Exit Sub
    
End Sub


Private Sub Form_Load()
    Dim lineSng As Single
    Dim strTime As String
    
    stringLengthFlag = False
    drawFlag = False
    speedPtr = 0
    
    cboTolerance.AddItem 0.1
    cboTolerance.AddItem 0.09
    cboTolerance.AddItem 0.08
    cboTolerance.AddItem 0.07
    cboTolerance.AddItem 0.06
    cboTolerance.AddItem 0.05
    
    If opt175.Value = True Then
        nominalTolerance = 1.75
        
    ElseIf opt28.Value = True Then
        nominalTolerance = 2.8
    
    ElseIf opt2875.Value = True Then
        nominalTolerance = 2.875
    
    ElseIf opt29.Value = True Then
        nominalTolerance = 2.9
    
    
    End If
    
    startTime = Now
    startTime = Format(Now, "hh:mm")
    strTime = Format(Now, "hh:mm")
    
    Picture2.AutoRedraw = True
    Picture3.AutoRedraw = True
    
    DrawAxis
    
    drawFlag = True
    

End Sub


Private Sub Laser_Click()
    frmComm.Show vbModal

    If MSComm1.PortOpen Then MSComm1.PortOpen = False
    MSComm1.CommPort = portNumber

    MSComm1.PortOpen = True
    
    lblSelect1.Visible = False


End Sub

Public Sub MSComm1_OnComm()

    Dim Buffer
    Dim ret As Integer
    Dim length As Integer
    Dim intI As Integer
    Dim intJ As Integer
    Dim strData As String
    Dim xDiameter As Single
    Dim yDiameter As Single
    Dim ovality As Single
    Dim percentOvality As Single
    Dim avgDiameter As Single
    Dim mSng As Single
    Dim lineSng As Single
    
    On Error GoTo ErrHandler1
    
    Select Case MSComm1.CommEvent
        Case comEvReceive
            Buffer = MSComm1.Input  'read buffer
            
            If stringLengthFlag = True Then
                rcvFlag = True
                gstrBuffer1(0) = Buffer(0)
                gstrBuffer1(1) = Buffer(1)
                gstrBuffer1(2) = Buffer(2)
                gstrBuffer1(3) = Buffer(3)
                gstrBuffer1(4) = Buffer(4)
                gstrBuffer1(5) = Buffer(5)
                gstrBuffer1(6) = Buffer(6)
                gstrBuffer1(7) = Buffer(7)
                gstrBuffer1(8) = Buffer(8)
                
                intI = gstrBuffer1(1) \ 16
                strData = CStr(intI)

                intI = gstrBuffer1(1) - (intI * 16)
                strData = strData & CStr(intI)

                intI = gstrBuffer1(2) \ 16
                strData = strData & CStr(intI)

                strData = strData & "."

                intI = gstrBuffer1(2) - (intI * 16)
                strData = strData & CStr(intI)

                intI = gstrBuffer1(3) \ 16
                strData = strData & CStr(intI)

                intI = gstrBuffer1(3) - (intI * 16)
                strData = strData & CStr(intI)

                txtX.Text = Format(strData, "0.000")

                xDiameter = CSng(strData)

                intI = gstrBuffer1(4) \ 16
                strData = CStr(intI)

                intI = gstrBuffer1(4) - (intI * 16)
                strData = strData & CStr(intI)

                intI = gstrBuffer1(5) \ 16
                strData = strData & CStr(intI)

                strData = strData & "."

                intI = gstrBuffer1(5) - (intI * 16)
                strData = strData & CStr(intI)

                intI = gstrBuffer1(6) \ 16
                strData = strData & CStr(intI)

                intI = gstrBuffer1(6) - (intI * 16)
                strData = strData & CStr(intI)

                txtY.Text = Format(strData, "0.000")

                yDiameter = CSng(strData)

        'X -Value
        'Y -Value
        'Avg diam = (X + Y) / 2
        'Ovality = absolute(X - Y)
        'Ovality % = ovality / avg diam * 100
                avgDiameter = (xDiameter + yDiameter) / 2
                txtDiameter.Text = Format(avgDiameter, "0.000")

                ovality = Abs(xDiameter - yDiameter)

                percentOvality = (ovality / avgDiameter) * 100

                txtOvality.Text = Format(percentOvality, "0.00")

                If startFlag = True Then
                   Print #1, Format(Now, "hh:mm:ss") & vbTab & xDiameter & vbTab & yDiameter & vbTab & avgDiameter & vbTab & percentOvality
'                   Print #1, Format(Now, "hh:mm:ss") & vbTab & CStr(gstrBuffer1(1))

                End If
'                If opt175.Value = True Then
'                    mSng = 1.75
'                Else
'                    mSng = 2.95
'
'                End If

                drawDiameter = 4000 - ((avgDiameter - nominalTolerance) * 1600 / CSng(cboTolerance.Text))
                drawDiameterBuffer(intCounter) = drawDiameter
                
                If drawFlag = True Then
                    Picture2.DrawWidth = 5
                    
                    If avgDiameter > nominalTolerance + CSng(cboTolerance.Text) Then
                        If optOn.Value = True Then
                            Beep
                        End If
                        shpAlarm.FillColor = &HFF&
                        
                        Picture2.PSet (intCounter, CInt(drawDiameter)), vbRed
                        drawColorBuffer(intCounter) = True
    
                    ElseIf avgDiameter < nominalTolerance - CSng(cboTolerance.Text) Then
                        If optOn.Value = True Then
                            Beep
                        End If
                        shpAlarm.FillColor = &HFF&
                        drawColorBuffer(intCounter) = True
                        
                        Picture2.PSet (intCounter, CInt(drawDiameter)), vbRed
    
                    Else
                        Picture2.PSet (intCounter, CInt(drawDiameter)), vbBlack
                        shpAlarm.FillColor = &HFF00&
                        drawColorBuffer(intCounter) = False
                    
    
                    End If
                End If

                drawOvality = (5# - percentOvality) * 460
                
                drawOvalityBuffer(intCounter) = drawOvality

                If drawFlag = True Then
                    Picture3.DrawWidth = 5
    
                    Picture3.PSet (intCounter, CInt(drawOvality)), vbBlack
                    
                    intCounter = intCounter + 1
                End If
                If intCounter >= 3000 Then  'move back one minute
                    intCounter = 2700   '1830
                    
                    startTime = DateAdd("n", 1, startTime)
                    
                    intJ = 0
                    For intI = 300 To 2999  '203 To 2033
                        drawDiameterBuffer(intJ) = drawDiameterBuffer(intI)
                        drawOvalityBuffer(intJ) = drawOvalityBuffer(intI)
                        drawColorBuffer(intJ) = drawColorBuffer(intI)


                        intJ = intJ + 1
                    
                    Next intI
                    
                    Picture2.Cls
                    Picture3.Cls
                    
                    DrawAxis
                    
                    
                    If drawFlag = True Then
                        Picture2.DrawWidth = 5
                        Picture3.DrawWidth = 5
                        
                        For intI = 0 To 2699    '1830    'redraw
                            If drawColorBuffer(intI) = True Then
                                Picture2.PSet (intI, drawDiameterBuffer(intI)), vbRed
            
                            Else
                                Picture2.PSet (intI, drawDiameterBuffer(intI)), vbBlack
            
                            End If
                        
                            Picture3.PSet (intI, drawOvalityBuffer(intI)), vbBlack
                        
                        Next intI
                    End If
                    
                End If
            
            Else
                If blnFlag = False Then
                    If Buffer(0) = &HA7 Then
                        blnFlag = True
                
                    End If
                Else
                    If Buffer(0) = &HA8 Then
                        blnFlag = False
                        stringLengthFlag = True
                        
                        MSComm1.InputLen = 9
                        MSComm1.RThreshold = 9
                    
                    End If
                End If
                
            End If
    End Select
    
    Exit Sub
    
ErrHandler1:
 '   Resume Next
    blnFlag = False
    stringLengthFlag = False 'look for next correct string
    
    Exit Sub
    
End Sub


Private Sub Timer1_Timer()
    If intTimer1 <> 0 Then
        intTimer1 = intTimer1 - 1
    End If
    
    
End Sub


Public Sub DrawAxis()
    Dim intI As Integer
    Dim intJ As Integer
    Dim intK As Single
    Dim intL As Integer
    Dim sngM As Single
    Dim yPos As Integer
    Dim strTemp As String           'temporary string
    Dim intHalfWidth As Integer
    Dim sngX As Single
    Dim sngFactor As Single
    Dim intHalfHeight As Integer
    
    drawFlag = False
    Picture2.DrawWidth = 1
    Picture3.DrawWidth = 1
    
    pic2Xaxis.Cls
    pic2Yaxis.Cls
    pic2Xaxis.AutoRedraw = True
    pic2Yaxis.AutoRedraw = True
    
    strTemp = Format(startTime, "hh:mm")
    pic2Xaxis.CurrentX = TextWidth(strTemp) / 2 ' Set X to center header
    pic2Xaxis.CurrentY = 50
    pic2Xaxis.Print CStr(strTemp)
        
    endTime = CStr(DateAdd("n", 10, startTime))
    strTemp = Format(endTime, "hh:mm")
    pic2Xaxis.CurrentX = 9300 - TextWidth(strTemp) / 2 ' Set X to center header
    pic2Xaxis.CurrentY = 50
    pic2Xaxis.Print strTemp
        

    strTemp = "Time"
    intHalfWidth = TextWidth(strTemp) / 2    ' Calculate one-half width of header.
    pic2Xaxis.CurrentX = pic2Xaxis.Width / 2 - intHalfWidth   ' Set X to center header
    pic2Xaxis.CurrentY = 200
    pic2Xaxis.Print strTemp
    
    
'y axis
    intL = TextHeight(strTemp) '/ 2  8400 - 7800
    intK = 0
    
    sngM = nominalTolerance + (CSng(cboTolerance.Text) * 2.5)
    
    For intJ = 300 To 8100 Step 780
        pic2Yaxis.Line (565, intJ)-(900, intJ), vbGreen 'draw y axis

        strTemp = Format(sngM - (intK * CSng(cboTolerance.Text)), "0.00")     '*2
        intK = intK + 0.5

        intHalfHeight = TextHeight(strTemp) / 2    ' Calculate one-half width of header.
        pic2Yaxis.CurrentY = intJ - intHalfHeight   ' Set X to center header

        pic2Yaxis.CurrentX = 100
        pic2Yaxis.Print strTemp


    Next intJ
    
    Picture2.Line (0, 2400)-(3000, 2400), vbRed
    
    Picture2.Line (0, 5600)-(3000, 5600), vbRed
    
    Picture2.Line (0, 4000)-(3000, 4000), vbBlack
    
    pic3Xaxis.Cls
    pic3Yaxis.Cls
    pic3Xaxis.AutoRedraw = True
    pic3Yaxis.AutoRedraw = True
    
    strTemp = Format(startTime, "hh:mm")
    pic3Xaxis.CurrentX = TextWidth(strTemp) / 2 ' Set X to center header
    pic3Xaxis.CurrentY = 50
    pic3Xaxis.Print CStr(strTemp)
        
    endTime = CStr(DateAdd("n", 10, startTime))
    strTemp = Format(endTime, "hh:mm")
    pic3Xaxis.CurrentX = 5200 - TextWidth(strTemp) / 2 ' Set X to center header
    pic3Xaxis.CurrentY = 50
    pic3Xaxis.Print strTemp

    strTemp = "Time"
    intHalfWidth = TextWidth(strTemp) / 2    ' Calculate one-half width of header.
    pic3Xaxis.CurrentX = pic3Xaxis.Width / 2 - intHalfWidth   ' Set X to center header
    pic3Xaxis.CurrentY = 200
    pic3Xaxis.Print strTemp
    'Picture3.Line (0, 2000)-(2034, 2000), vbBlack
    
    
       

    strTemp = "Time"
    intHalfWidth = TextWidth(strTemp) / 2    ' Calculate one-half width of header.
    pic2Xaxis.CurrentX = pic2Xaxis.Width / 2 - intHalfWidth   ' Set X to center header
    pic2Xaxis.CurrentY = 200
    pic2Xaxis.Print strTemp
    
    
    
    
    
    
    
   For intJ = 5 To 0 Step -1
        pic3Yaxis.Line (565, 260 + (intJ * 430))-(765, 260 + (intJ * 430)), vbGreen 'draw y axis
        
        strTemp = CStr(intK)
        intK = intK + 1
        
        intHalfHeight = TextHeight(strTemp) / 2    ' Calculate one-half width of header.
        pic3Yaxis.CurrentY = 260 + (intJ * 430) - intHalfHeight ' Set X to center header
        
        pic3Yaxis.CurrentX = 0
        pic3Yaxis.Print strTemp
        
        
    Next intJ

    drawFlag = True

End Sub

