VERSION 5.00
Begin VB.Form frmComm 
   BackColor       =   &H00FFFFC0&
   Caption         =   "COM port"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000FF00&
      Caption         =   "Return"
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox cboComm 
      Height          =   1350
      Left            =   1080
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "COM Ports"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   765
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// API constants
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const OPEN_EXISTING = 3
Const FILE_ATTRIBUTE_NORMAL = &H80

Private Sub cmdReturn_Click()
    Dim lineNumber As Integer
    
    lineNumber = cboComm.ListIndex
    
    portNumber = CInt(cboComm.List(lineNumber))


    
    Unload Me

End Sub

Private Sub Form_Load()
    Call ListComPorts

End Sub

'// Return TRUE if the COM exists, FALSE if the COM does not exist
Public Function COMAvailable(COMNum As Integer) As Boolean
    Dim hCOM As Long
    Dim ret As Long
    Dim sec As SECURITY_ATTRIBUTES

    '// try to open the COM port
    hCOM = CreateFile("\.\COM" & COMNum & "", 0, FILE_SHARE_READ + _
        FILE_SHARE_WRITE, sec, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hCOM = -1 Then
        COMAvailable = False
    Else
        COMAvailable = True
        '// close the COM port
        ret = CloseHandle(hCOM)
    End If
End Function


Private Sub ListComPorts()
    Dim i As Integer
    
    cboComm.Clear
    For i = 1 To 16
        If COMAvailable(i) Then
            cboComm.AddItem i
        End If
    Next
    cboComm.ListIndex = 0
End Sub


