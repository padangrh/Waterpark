VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Main"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Force Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   2880
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2160
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmr_Jam 
      Interval        =   1000
      Left            =   960
      Top             =   2880
   End
   Begin VB.CommandButton cmdC1_3 
      BackColor       =   &H000000FF&
      Height          =   735
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdC1_2 
      BackColor       =   &H000000FF&
      Height          =   735
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cmdC1_1 
      BackColor       =   &H000000FF&
      Height          =   735
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.Timer Tmr_RFIDEX 
      Interval        =   60000
      Left            =   360
      Top             =   2880
   End
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   735
      Left            =   240
      OleObjectBlob   =   "Main.frx":0000
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin zkemkeeperCtl.CZKEM CZKEM2 
      Height          =   735
      Left            =   240
      OleObjectBlob   =   "Main.frx":0024
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin zkemkeeperCtl.CZKEM CZKEM3 
      Height          =   735
      Left            =   240
      OleObjectBlob   =   "Main.frx":0048
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblDT 
      Alignment       =   2  'Center
      Caption         =   "dd/MM/yyyy HH:mm:ss"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   6
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timerCount As Integer

Private Sub cmdC1_1_Click()
    Tmr_RFIDEX.Enabled = False
    DoEvents
    If cmdC1_1.BackColor <> &HFF00& Then
        cmdC1_1.BackColor = &HFFFF&
        If confirmC1(Setting_Object("C1_1")) Then
            Dim C1_1Con As Boolean
            CZKEM1.BASE64 = 1
            C1_1Con = False
            C1_1Con = CZKEM1.Connect_Net(Setting_Object("C1_1"), 4370)
            If C1_1Con Then CZKEM1.Beep 150
            refillC1 1
            cmdC1_1.BackColor = &HFF00&
        Else
            cmdC1_1.BackColor = &HFF&
        End If
    End If
    Tmr_RFIDEX.Enabled = True
End Sub

Private Sub cmdC1_2_Click()
    Tmr_RFIDEX.Enabled = False
    DoEvents
    If cmdC1_2.BackColor <> &HFF00& Then
        cmdC1_2.BackColor = &HFFFF&
        If confirmC1(Setting_Object("C1_2")) Then
            Dim C1_2Con As Boolean
            CZKEM2.BASE64 = 1
            C1_2Con = False
            C1_2Con = CZKEM2.Connect_Net(Setting_Object("C1_2"), 4370)
            If C1_2Con Then CZKEM2.Beep 150
            refillC1 2
            cmdC1_2.BackColor = &HFF00&
        Else
            cmdC1_2.BackColor = &HFF&
        End If
    End If
    Tmr_RFIDEX.Enabled = True
End Sub

Private Sub cmdC1_3_Click()
    Tmr_RFIDEX.Enabled = False
    DoEvents
    If cmdC1_3.BackColor <> &HFF00& Then
        cmdC1_3.BackColor = &HFFFF&
        If confirmC1(Setting_Object("C1_3")) Then
            Dim C1_3Con As Boolean
            CZKEM3.BASE64 = 1
            C1_3Con = False
            C1_3Con = CZKEM3.Connect_Net(Setting_Object("C1_3"), 4370)
            If C1_3Con Then CZKEM3.Beep 150
            refillC1 3
            cmdC1_3.BackColor = &HFF00&
        Else
            cmdC1_3.BackColor = &HFF&
        End If
    End If
    Tmr_RFIDEX.Enabled = True
End Sub

Private Sub cmdUpdate_Click()
    Dim yy As Boolean
    Tmr_RFIDEX.Enabled = False
    cmdUpdate.BackColor = &HFFFF&
    DoEvents
    
    yy = confirmC1(Setting_Object("C1_1"))
    If Not yy Then
        frmMain.cmdC1_1.BackColor = &HFF&
    End If
    StatusC1_1 = yy
        
    yy = confirmC1(Setting_Object("C1_2"))
    If Not yy Then
        frmMain.cmdC1_2.BackColor = &HFF&
    End If
    StatusC1_2 = yy
    
    yy = confirmC1(Setting_Object("C1_3"))
    If Not yy Then
        frmMain.cmdC1_3.BackColor = &HFF&
    End If
    StatusC1_3 = yy
    
    
'    StatusC1_1 = confirmC1(Setting_Object("C1_1"))
'    StatusC1_2 = confirmC1(Setting_Object("C1_2"))
'    StatusC1_3 = confirmC1(Setting_Object("C1_3"))
    
'    If StatusC1_1 Then refillC1 1
'    If StatusC1_2 Then refillC1 2
'    If StatusC1_3 Then refillC1 3
    
    If StatusC1_1 Then
        cmdC1_1.BackColor = &HFF00&
        Dim C1_1Con As Boolean
        CZKEM1.BASE64 = 1
        C1_1Con = False
        C1_1Con = CZKEM1.Connect_Net(Setting_Object("C1_1"), 4370)
        If C1_1Con Then CZKEM1.Beep 150
        refillC1 1
    End If
    If StatusC1_2 Then
        cmdC1_2.BackColor = &HFF00&
        Dim C1_2Con As Boolean
        CZKEM2.BASE64 = 1
        C1_2Con = False
        C1_2Con = CZKEM2.Connect_Net(Setting_Object("C1_2"), 4370)
        If C1_2Con Then CZKEM2.Beep 150
        refillC1 2
    End If
    If StatusC1_3 Then
        cmdC1_3.BackColor = &HFF00&
        Dim C1_3Con As Boolean
        CZKEM3.BASE64 = 1
        C1_3Con = False
        C1_3Con = CZKEM3.Connect_Net(Setting_Object("C1_3"), 4370)
        If C1_3Con Then CZKEM3.Beep 150
        refillC1 3
    End If
    Tmr_RFIDEX.Enabled = True
    cmdUpdate.BackColor = &HFFC0C0
End Sub

Private Sub Form_Load()
    lblDT.Caption = Format(Now, "dd/MM/yyyy HH:mm:ss")
    timerCount = 0
    Dim namafile, file_data, huruf As String
    namafile = App.Path & "\DataReset.txt"
    IFile = FreeFile
    Open namafile For Input As #IFile
    file_data = StrConv(InputB(LOF(IFile), IFile), vbUnicode)
    Close #IFile
    Dim xx As Integer
    xx = DateDiff("d", file_data, Now)
    If DateDiff("d", file_data, Now) > 0 Then
        con.Execute ("Delete from tbreader")
        con.Execute ("alter table tbreader auto_increment = 1")
        Open namafile For Output As #1
        Print #1, Now
        Close #1
    End If
    
    con.Execute ("alter table tbreader auto_increment = 1")
    
    StatusC1_1 = Setting_Object("C1_1Status")
    StatusC1_2 = Setting_Object("C1_2Status")
    StatusC1_3 = Setting_Object("C1_3Status")
    
    If StatusC1_1 Then
        If confirmC1(Setting_Object("C1_1")) Then
            cmdC1_1.BackColor = &HFF00&
            Dim C1_1Con As Boolean
            CZKEM1.BASE64 = 1
            C1_1Con = False
            C1_1Con = CZKEM1.Connect_Net(Setting_Object("C1_1"), 4370)
            If C1_1Con Then CZKEM1.Beep 150
        Else
            StatusC1_1 = False
        End If
    End If
    If StatusC1_2 Then
        If confirmC1(Setting_Object("C1_2")) Then
            cmdC1_2.BackColor = &HFF00&
            Dim C1_2Con As Boolean
            CZKEM2.BASE64 = 1
            C1_2Con = False
            C1_2Con = CZKEM2.Connect_Net(Setting_Object("C1_2"), 4370)
            If C1_2Con Then CZKEM2.Beep 150
        Else
            StatusC1_2 = False
        End If
    End If
    If StatusC1_3 Then
        If confirmC1(Setting_Object("C1_3")) Then
            cmdC1_3.BackColor = &HFF00&
            Dim C1_3Con As Boolean
            CZKEM3.BASE64 = 1
            C1_3Con = False
            C1_3Con = CZKEM3.Connect_Net(Setting_Object("C1_3"), 4370)
            If C1_3Con Then CZKEM3.Beep 150
        Else
            StatusC1_3 = False
        End If
    End If
End Sub

Private Sub Form_unload(Cancel As Integer)
    DoEvents
    con.Close
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
End Sub

Private Sub tmr_Jam_Timer()
    lblDT.Caption = Format(Now, "dd/MM/yyyy HH:mm:ss")
End Sub

Private Sub Tmr_RFIDEX_Timer()
    confirmAllC1
    timerCount = timerCount + 1
    If timerCount >= 30 Then
    Dim rsTbAktif As ADODB.Recordset
'    Dim rsReader As ADODB.Recordset
        Set rsTbAktif = con.Execute("select * from tbaktif where time_to_sec(timeDIFF(now(),concat(tanggal, ' ' , jam)))/3600 > 6 and status <> 0")
        Do While Not rsTbAktif.EOF
            con.Execute ("Update tbaktif set status = '0' where rfid = '" & rsTbAktif!rfid & "'")
    '        Set rsReader = con.Execute("select * from tbreader where rfid = '" & rsTbAktif!rfid & "'")
            deleteC1 rsTbAktif!rfid
            con.Execute ("Delete from tbreader where rfid = '" & rsTbAktif!rfid & "'")
    '        Set rsReader = Nothing
            rsTbAktif.MoveNext
        Loop
        timerCount = 0
    End If
    
End Sub

'Private Sub connectC1(ip As String)
'    Dim bconn As Boolean
'
'    CZKEM1.BASE64 = 1
'    CZKEM2.BASE64 = 1
'    CZKEM3.BASE64 = 1
'    bconn = False
'
'    bconn = CZKEM1.Connect_Net("192.168.1.250", 4370)
'    If bconn Then CZKEM1.Beep 150
'
'    bconn = CZKEM2.Connect_Net("192.168.1.251", 4370)
'    If bconn Then CZKEM2.Beep 150
'
'    bconn = CZKEM3.Connect_Net("192.168.1.252", 4370)
'    If bconn Then CZKEM3.Beep 150
'End Sub

'Private Sub disconnectC1()
'    CZKEM1.Beep 150
'    DoEvents
'    CZKEM1.Disconnect
'End Sub


