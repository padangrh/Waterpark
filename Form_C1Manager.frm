VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_C1Manager 
   BackColor       =   &H00004080&
   Caption         =   "C1 Manager"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_Refresh 
      BackColor       =   &H0080FFFF&
      Caption         =   "Refresh"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmd_Hapus 
      BackColor       =   &H000000FF&
      Caption         =   "Hapus"
      Enabled         =   0   'False
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
      Left            =   15720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmd_Tambah 
      BackColor       =   &H000000FF&
      Caption         =   "Tambah"
      Enabled         =   0   'False
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
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmd_Open3 
      BackColor       =   &H0080FF80&
      Caption         =   "Pintu 3"
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
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmd_Open2 
      BackColor       =   &H0080FF80&
      Caption         =   "Pintu 2"
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
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmd_Open1 
      BackColor       =   &H0080FF80&
      Caption         =   "Pintu 1"
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
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmd_Push 
      BackColor       =   &H00808080&
      Caption         =   " Push >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmd_Load 
      BackColor       =   &H00808080&
      Caption         =   "Load Data"
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
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ComboBox Cmb_Mesin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form_C1Manager.frx":0000
      Left            =   11400
      List            =   "Form_C1Manager.frx":000D
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00004040&
      Caption         =   "Mesin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   9720
      TabIndex        =   12
      Top             =   1680
      Width           =   7875
      Begin VB.Frame frame_Tambah 
         Caption         =   "Tambah RFID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   4200
         TabIndex        =   16
         Top             =   3000
         Visible         =   0   'False
         Width           =   3015
         Begin VB.CommandButton cmd_TambahCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   1680
            TabIndex        =   22
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton cmd_TambahOK 
            Caption         =   "OK"
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txt_RFID 
            Height          =   375
            Left            =   720
            TabIndex        =   18
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox txt_ID 
            Height          =   375
            Left            =   720
            TabIndex        =   17
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "RFID"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "ID"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.CommandButton cmd_RefreshMesin 
         BackColor       =   &H0080FFFF&
         Caption         =   "Refresh"
         Enabled         =   0   'False
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5640
         Width           =   1455
      End
      Begin MSComctlLib.ListView lv_Mesin 
         Height          =   4815
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "RFID"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004040&
      Caption         =   "Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   7875
      Begin MSComctlLib.ListView lv_Database 
         Height          =   4815
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "RFID"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "Buka Pintu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   17880
      TabIndex        =   14
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Pengaturan User C1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   13
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "Form_C1Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Hapus_Click()
    If MsgBox("Anda yakin akan menghapus ID " & lv_Mesin.SelectedItem.Text, vbYesNo, "Konfirmasi") Then
        If lv_Mesin.SelectedItem.Text <> "65534" Then
            Dim bDltUser As Boolean
            If Frame2.Caption = "Mesin 1" Then
                bDltUser = FrmMain.CZKEM1.DeleteEnrollData(1, lv_Mesin.SelectedItem.Text, 1, 12)
                Call refresh_LV_Mesin(FrmMain.CZKEM1, StatusC1_1, FrmMain.cmdC1_1, Setting_Object("C1_1"))
            ElseIf Frame2.Caption = "Mesin 2" Then
                bDltUser = FrmMain.CZKEM2.DeleteEnrollData(1, lv_Mesin.SelectedItem.Text, 1, 12)
                Call refresh_LV_Mesin(FrmMain.CZKEM2, StatusC1_2, FrmMain.cmdC1_2, Setting_Object("C1_2"))
            ElseIf Frame2.Caption = "Mesin 3" Then
                bDltUser = FrmMain.CZKEM3.DeleteEnrollData(1, lv_Mesin.SelectedItem.Text, 1, 12)
                Call refresh_LV_Mesin(FrmMain.CZKEM3, StatusC1_3, FrmMain.cmdC1_3, Setting_Object("C1_3"))
            End If
        Else
            MsgBox ("Admin tidak bisa dihapus")
        End If
    End If
End Sub

Private Sub cmd_Load_Click()
    If MsgBox("Load data " & Cmb_Mesin.Text, vbYesNo, "Konfirmasi") = vbYes Then
        If frame_Tambah.Visible = True Then
            cmd_TambahCancel_Click
        End If
        
        If Cmb_Mesin.ListIndex = 0 Then
            Call refresh_LV_Mesin(FrmMain.CZKEM1, StatusC1_1, FrmMain.cmdC1_1, Setting_Object("C1_1"))
        ElseIf Cmb_Mesin.ListIndex = 1 Then
            Call refresh_LV_Mesin(FrmMain.CZKEM2, StatusC1_2, FrmMain.cmdC1_2, Setting_Object("C1_2"))
        ElseIf Cmb_Mesin.ListIndex = 2 Then
            Call refresh_LV_Mesin(FrmMain.CZKEM3, StatusC1_3, FrmMain.cmdC1_3, Setting_Object("C1_3"))
        Else
            MsgBox "Mesin tidak ditemukan"
            Cmb_Mesin.ListIndex = 0
            Exit Sub
        End If
        
        cmd_RefreshMesin.Enabled = True
        cmd_Tambah.Enabled = True
        cmd_Hapus.Enabled = True
            
    End If
End Sub

Private Sub cmd_Open1_Click()
    If MsgBox("Buka pintu 1?", vbYesNo, "Konfirmasi") = vbYes Then
        If StatusC1_1 Then
            FrmMain.CZKEM1.ACUnlock CLng(1), 100
        End If
    End If
End Sub

Private Sub cmd_Open2_Click()
    If MsgBox("Buka pintu 2?", vbYesNo, "Konfirmasi") = vbYes Then
        If StatusC1_2 Then
            FrmMain.CZKEM2.ACUnlock CLng(1), 100
        End If
    End If
End Sub

Private Sub cmd_Open3_Click()
    If MsgBox("Buka pintu 3?", vbYesNo, "Konfirmasi") = vbYes Then
        If StatusC1_3 Then
            FrmMain.CZKEM3.ACUnlock CLng(1), 100
        End If
    End If
End Sub

Private Sub cmd_Push_Click()
    If Frame2.Caption = "Mesin" Then
        cmd_Load_Click
    End If
    If Frame2.Caption <> "Mesin" Then
        If MsgBox("Push data ke " & Frame2.Caption & " ?", vbYesNo, "Konfirmasi") = vbYes Then
            If Frame2.Caption = "Mesin 1" And StatusC1_1 = True Then
                refillC1 (1)
                Call refresh_LV_Mesin(FrmMain.CZKEM1, StatusC1_1, FrmMain.cmdC1_1, Setting_Object("C1_1"))
            ElseIf Frame2.Caption = "Mesin 2" And StatusC1_2 = True Then
                refillC1 (2)
                Call refresh_LV_Mesin(FrmMain.CZKEM2, StatusC1_2, FrmMain.cmdC1_2, Setting_Object("C1_2"))
            ElseIf Frame2.Caption = "Mesin 3" And StatusC1_3 = True Then
                refillC1 (3)
                Call refresh_LV_Mesin(FrmMain.CZKEM3, StatusC1_3, FrmMain.cmdC1_3, Setting_Object("C1_3"))
            End If
        End If
        Cmb_Mesin.Text = Frame2.Caption
    End If
End Sub

Private Sub cmd_Refresh_Click()
    Call refresh_LV_Database
End Sub

Private Sub cmd_RefreshMesin_Click()
    If Frame2.Caption = "Mesin 1" Then
        Call refresh_LV_Mesin(FrmMain.CZKEM1, StatusC1_1, FrmMain.cmdC1_1, Setting_Object("C1_1"))
    ElseIf Frame2.Caption = "Mesin 2" Then
        Call refresh_LV_Mesin(FrmMain.CZKEM2, StatusC1_2, FrmMain.cmdC1_2, Setting_Object("C1_2"))
    ElseIf Frame2.Caption = "Mesin 3" Then
        Call refresh_LV_Mesin(FrmMain.CZKEM3, StatusC1_3, FrmMain.cmdC1_3, Setting_Object("C1_3"))
    End If
End Sub

Private Sub cmd_Tambah_Click()
    frame_Tambah.Visible = True
End Sub

Private Sub cmd_TambahCancel_Click()
    txt_id.Text = ""
    txt_RFID.Text = ""
    frame_Tambah.Visible = False
End Sub

Private Sub cmd_TambahOK_Click()
    If txt_id.Text < 65534 And txt_id.Text > 0 And Len(txt_RFID.Text) = 8 Then
        If Frame2.Caption = "Mesin 1" Then
            FrmMain.CZKEM1.CardNumber(0) = txt_RFID.Text
            bcrtUser = FrmMain.CZKEM1.SetUserInfo(1, txt_id.Text, "Tamu", "0751", 0, 1)
            Call refresh_LV_Mesin(FrmMain.CZKEM1, StatusC1_1, FrmMain.cmdC1_1, Setting_Object("C1_1"))
        ElseIf Frame2.Caption = "Mesin 2" Then
            FrmMain.CZKEM2.CardNumber(0) = txt_RFID.Text
            bcrtUser = FrmMain.CZKEM2.SetUserInfo(1, txt_id.Text, "Tamu", "0751", 0, 1)
            Call refresh_LV_Mesin(FrmMain.CZKEM2, StatusC1_2, FrmMain.cmdC1_2, Setting_Object("C1_2"))
        ElseIf Frame2.Caption = "Mesin 3" Then
            FrmMain.CZKEM3.CardNumber(0) = txt_RFID.Text
            bcrtUser = FrmMain.CZKEM3.SetUserInfo(1, txt_id.Text, "Tamu", "0751", 0, 1)
            Call refresh_LV_Mesin(FrmMain.CZKEM3, StatusC1_3, FrmMain.cmdC1_3, Setting_Object("C1_3"))
        End If
        txt_id.Text = ""
        txt_RFID.Text = ""
        frame_Tambah.Visible = False
    Else
        MsgBox "ID tidak boleh lebih dari 65534, tidak boleh lebih kecil dari 0" & vbNewLine & "RFID harus 8 digit"
    End If

End Sub

Private Sub Form_Load()
    Cmb_Mesin.ListIndex = 0
    refresh_LV_Database
End Sub

Private Sub refresh_LV_Database()
    lv_Database.ListItems.Clear
    Dim rsReader As ADODB.Recordset
    Dim litem As ListItem
    Set rsReader = con.Execute("Select * from tbreader")
    Do While Not rsReader.EOF
        Set litem = lv_Database.ListItems.Add(, , rsReader!id)
        litem.SubItems(1) = rsReader!rfid
        rsReader.MoveNext
    Loop
    Set rsReader = Nothing
End Sub

Private Sub refresh_LV_Mesin(ByRef tempCZKEM As CZKEM, ByRef tempStatus As Boolean, ByRef tempButton As CommandButton, ByVal tempIP As String)
    If tempStatus = False Then
        If MsgBox("Check status mesin?", vbYesNo, "Konfirmasi") = vbYes Then
            tempStatus = confirmC1(tempIP)
            If tempStatus Then
                tempButton.BackColor = &HFF00&
            Else
                tempButton.BackColor = &HFF&
                MsgBox "Mesin tidak ditemukan."
                Exit Sub
            End If
        Else
            Exit Sub
        End If
        Dim C1Con As Boolean
        tempCZKEM.BASE64 = 1
        C1Con = False
        C1Con = tempCZKEM.Connect_Net(tempIP, 4370)
        If C1Con Then tempCZKEM.Beep 150
    End If
    
    Dim dwEnrollNmber As Long
    Dim name As String
    Dim pwd As String
    Dim privilege As Long
    Dim sEnabled As Boolean
    Dim tempRFID As String
    
    lv_Mesin.ListItems.Clear
    Frame2.Caption = Cmb_Mesin.Text
    
    If tempCZKEM.ReadAllUserID(1) Then
        dwEnrollNmber = 0
        Dim litem As ListItem
        Do While tempCZKEM.GetAllUserInfo(CLng(1), dwEnrollNmber, name, pwd, privilege, sEnabled)
            Set litem = lv_Mesin.ListItems.Add(, , dwEnrollNmber)
            litem.SubItems(1) = tempCZKEM.CardNumber(0)
        Loop
    End If
End Sub

Private Sub txt_ID_KeyPress(KeyAscii As Integer)
    Call validateKey(KeyAscii, 1)
End Sub

Private Sub txt_RFID_KeyPress(KeyAscii As Integer)
    Call validateKey(KeyAscii, 1)
End Sub
