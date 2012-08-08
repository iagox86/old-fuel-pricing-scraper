VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Fuel Pricing"
   ClientHeight    =   4425
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock ws 
      Left            =   0
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "204.110.224.126"
      RemotePort      =   80
      LocalPort       =   69
   End
   Begin VB.TextBox txtWindow 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   600
      Width           =   5415
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame fraStates 
      Caption         =   "States"
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2655
      Begin VB.ListBox lstStates 
         Height          =   2400
         ItemData        =   "frmMain.frx":030A
         Left            =   120
         List            =   "frmMain.frx":030C
         MultiSelect     =   1  'Simple
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdNone 
         Caption         =   "&None"
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "&All"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2760
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Select the states that you want to print Fuel Prices for:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mBuffer As String
Dim mWholePage As String
Dim mGoodPage As String
Dim States(0 To 100) As String
Dim Prices(0 To 200) As String
Dim miNumber As Integer
Dim mFinalProduct As String
Const TheBeginning = "<TR height=20>"
Const Beginning2 = "<b>"
Const Beginning3 = "<FONT face=Arial,Helvetica,Geneva,Swiss,SunSans-Regular size=2>"
Const Beginning4 = "<FONT face=Arial,Helvetica,Geneva,Swiss,SunSans-Regular size=2>"
Const TheEnd = "</TR>"
Const End2 = "</b>"
Const End3 = "</FONT>"
Const end4 = "</font>"

Private Sub GetStringReady()
    Dim iindex As Integer
    Dim iindex2 As Integer
    Dim State As String
    Dim Use As Boolean
    mFinalProduct = ""
    For iindex = LBound(Prices) To UBound(Prices)
        State = Left(Prices(iindex), 2)
        Use = False
        
        For iindex2 = 0 To lstStates.ListCount - 1
            If lstStates.List(iindex2) = State And lstStates.Selected(iindex2) = True Then
                Use = True
            End If
        Next
        
        If Use Then
            mFinalProduct = mFinalProduct & Prices(iindex) & vbCrLf

        End If
    Next
    If mFinalProduct <> "" Then
        mFinalProduct = Left(mFinalProduct, Len(mFinalProduct) - 2)
    End If
    
End Sub

Private Sub cmdAll_Click()
    Dim iindex As Integer
    For iindex = lstStates.ListCount - 1 To 0 Step -1
        lstStates.Selected(iindex) = True
    Next
End Sub

Private Sub cmdNone_Click()
    Dim iindex As Integer
    
    For iindex = lstStates.ListCount - 1 To 0 Step -1
        lstStates.Selected(iindex) = False
    Next
End Sub

Private Sub cmdPreview_Click()
    Call GetStringReady
    
    txtWindow.Text = mFinalProduct
End Sub

Private Sub cmdPrint_Click()
    Call GetStringReady

    

    Printer.FontSize = 10
    Printer.FontName = "Courier"
    Printer.Print mFinalProduct
    Printer.EndDoc
End Sub

Private Sub Form_Load()
    'randomize the port, because that's what I like doing
    Randomize Timer
    ws.LocalPort = Int(Rnd(10000))
    ws.RemoteHost = "204.110.224.126"
    ws.RemotePort = 80
    
    'connect
    ws.Connect
End Sub

Private Sub Form_Terminate()
    ws.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ws.Close
    
End Sub

Private Sub txtWindow_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub ws_Close()
    ws.Close
    cmdPreview.Enabled = True

    Call FilterOutTheCrap
    
    Unload frmSplash
    frmMain.Show
End Sub

Private Sub ws_Connect()
    'this should send the HTTP request to the server
    ws.SendData _
    ( _
        "GET /fuel/diesel_CF.cfm HTTP/1.1" & vbCrLf & _
        "host:serv2.flyingj.com" & vbCrLf & vbCrLf _
    )
    
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    Dim Temp As String
    
    ws.GetData Temp
    
 
    mBuffer = mBuffer & Temp
    Temp = ""
    Call ProcessData
    
    
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Winsock caused an error." & vbCrLf & _
    "Number: " & Number & vbCrLf & _
    "Description: " & Description & vbCrLf & vbCrLf & _
    "Port " & ws.LocalPort & vbCrLf & _
    "Remoteaddr " & ws.RemoteHost _
    , vbCritical, "Winsock Error"
    
    ws.Close
End Sub

Private Sub ProcessData()

    mWholePage = mWholePage & mBuffer
    
    mBuffer = ""
    
    If (InStr(1, mWholePage, "</html>")) Then
        Call ws_Close
    End If
    
End Sub

Private Sub FilterOutTheCrap()
     Dim Num As Integer
     Dim iindex As Integer
     Dim Index1 As Integer
     Dim Index2 As Integer
     Dim Temp As String
     Dim ToAdd As String
     Dim NewState As String
     Dim Good As Boolean
     
     mWholePage = Replace(mWholePage, Chr(&HA), "")
     mWholePage = Replace(mWholePage, Chr(&HD), "")
     mWholePage = Replace(mWholePage, Chr(&H9), "")
     mWholePage = Replace(mWholePage, "&nbsp;", "")
     mWholePage = Replace(mWholePage, "1000", "")
    
    'Index1 is the position of TheBeginning (<tr height=30> in the string)
    'Index2 is the position of TheEnd ("/tr") in the string
    Index1 = 1
    Index2 = 1
    Num = 0
    
    While (Index1 > 0 And Index2 > 0)
        Num = Num + 1
        Index1 = InStr(1, mWholePage, TheBeginning)
        If Index1 > 0 Then
            ToAdd = ""
            Temp = Between(mWholePage, TheBeginning, TheEnd)
            'Get rid of the left portion of wholepage up to the end
            mWholePage = Replace(mWholePage, Left(mWholePage, miNumber + Len(TheEnd)), "", , 1)
    
            ToAdd = ToAdd & Between(Temp, Beginning2, End2) & vbCrLf
            Temp = Replace(Temp, Left(Temp, miNumber + Len(End2)), "", , 1)
            ToAdd = ToAdd & Between(Temp, Beginning4, end4) & vbCrLf
            Temp = Replace(Temp, Left(Temp, miNumber + Len(end4)), "", , 1)
            ToAdd = ToAdd & Between(Temp, Beginning4, end4) & vbCrLf
    
            NewState = Left(ToAdd, 2)
            
            Good = True
            For iindex = 0 To lstStates.ListCount
                If lstStates.List(iindex) = NewState Then
                    Good = False
                End If
            Next
            
            If (Good) Then
                lstStates.AddItem NewState
            End If
        End If
    
        Prices(Num) = ToAdd
    
    Wend
 
End Sub

Private Function Between(str As String, a As String, b As String)
    Dim Index1, Index2 As Integer
    
    On Error Resume Next

    'First, set index1 to the first instance of the beginning
    Index1 = InStr(1, str, a)
    
    
    If Index1 > 0 Then
        'Set index2 to the first instance of the end after Index1
        Index2 = InStr(Index1, str, b)
        miNumber = Index2
        'Add the new piece to the goodpage
        Between = Mid(str, Index1 + Len(a), Index2 - Index1 - Len(a))

    End If
End Function


