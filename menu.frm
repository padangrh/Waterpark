VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Begin VB.MDIForm FrmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Main Menu"
   ClientHeight    =   10635
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15705
   Icon            =   "menu.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "menu.frx":628A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu.frx":294AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu.frx":2999A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu.frx":29C4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu.frx":29F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu.frx":2A2A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu.frx":2A65D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15705
      _ExtentX        =   27702
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Penjualan"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Laporan"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Barang"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Admin"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Logout"
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin zkemkeeperCtl.CZKEM CZKEM3 
         Height          =   375
         Left            =   12240
         OleObjectBlob   =   "menu.frx":4D891
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin zkemkeeperCtl.CZKEM CZKEM2 
         Height          =   375
         Left            =   11760
         OleObjectBlob   =   "menu.frx":4D8B5
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdC1_3 
         BackColor       =   &H000000FF&
         Height          =   615
         Left            =   3720
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdC1_2 
         BackColor       =   &H000000FF&
         Height          =   615
         Left            =   3360
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdC1_1 
         BackColor       =   &H000000FF&
         Height          =   615
         Left            =   3000
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   375
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   10680
         Top             =   120
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   10200
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin zkemkeeperCtl.CZKEM CZKEM1 
         Height          =   375
         Left            =   11280
         OleObjectBlob   =   "menu.frx":4D8D9
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H0080FF80&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   7080
         TabIndex        =   2
         Text            =   "0"
         Top             =   0
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
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
         Left            =   4320
         TabIndex        =   1
         Text            =   "0"
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Menu p 
      Caption         =   "Penjualan"
      Begin VB.Menu pjl 
         Caption         =   "Trans. Penjualan"
      End
      Begin VB.Menu ad 
         Caption         =   "Ambil Deposit"
      End
      Begin VB.Menu kk 
         Caption         =   "Kehilangan Kartu"
      End
      Begin VB.Menu lgf 
         Caption         =   "Logoff"
      End
   End
   Begin VB.Menu l 
      Caption         =   "Laporan"
      Begin VB.Menu lpr 
         Caption         =   "Laporan"
      End
   End
   Begin VB.Menu b 
      Caption         =   "Barang"
      Begin VB.Menu sp 
         Caption         =   "Entri Suplier"
      End
      Begin VB.Menu ebr 
         Caption         =   "Entri Barang / Stock"
      End
   End
   Begin VB.Menu a 
      Caption         =   "Admin"
      Begin VB.Menu tu 
         Caption         =   "User Manager"
      End
      Begin VB.Menu rm 
         Caption         =   "RFID Manager"
      End
      Begin VB.Menu hs 
         Caption         =   "History"
      End
   End
   Begin VB.Menu lg 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim active_form As Form

Public Sub changeForm(new_form As Form)
    If active_form Is new_form Then
        Exit Sub
    End If
    
    new_form.Show
    
    If Not active_form Is Nothing Then
        Unload active_form
    End If
    Set active_form = new_form
End Sub

Private Sub edb_Click()
    Call changeForm(Form_List_barang)
End Sub

Private Sub eds_Click()
    Call changeForm(Form_List_Supplier)
End Sub

Private Sub k_Click()
    Unload Me
End Sub

Private Sub bd_Click()
    Call changeForm(Form_List_Supplier)
End Sub

Private Sub ad_Click()
    'Form_Deposit.Show vbModal, Me
    Call changeForm(Form_List_Deposit)
End Sub

Private Sub cmdC1_1_Click()
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
End Sub

Private Sub cmdC1_2_Click()
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
End Sub

Private Sub cmdC1_3_Click()
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
End Sub

Private Sub ebr_Click()
    Call changeForm(Form_List_barang)
End Sub

Private Sub hs_Click()
    Form_List_NonAktif.Show
End Sub

Private Sub kk_Click()
    Form_ReplaceRFID.Show (1)
End Sub

Private Sub lg_Click()
    Unload Me
End Sub

Private Sub lgf_Click()
    logoff
End Sub

Public Sub logoff()
    'Unload active_form
    Unload Me
    Set active_form = Nothing
    username = ""
    status = ""
    frmlogin.Show (1)
End Sub

Private Sub lpr_Click()
    Form_Laporan.Show (1)
End Sub

Private Sub pb_Click()
    Call changeForm(Form_List_beli)
End Sub

Private Sub pjl_Click()
    Call changeForm(Form_List_Jual)
End Sub

Private Sub rbl_Click()
    Call changeForm(Form_List_Supplier)
End Sub

Private Sub rjl_Click()
    Call changeForm(Form_List_Supplier)
End Sub

Private Sub rm_Click()
    Form_List_RFID.Show
End Sub

Private Sub sp_Click()
    Call changeForm(Form_List_Supplier)
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
    Case 1: PopupMenu p
    Case 2: PopupMenu l
    Case 3: PopupMenu b
    Case 4: PopupMenu a
    Case 5: Unload Me
    End Select
End Sub

Private Sub MDIForm_Activate()
    If username = "" Then
       frmlogin.Show 1
    End If
End Sub

Private Sub MDIForm_Load()
    Set active_form = Nothing
    Dim temp_stringX, File_StringX As String
    Dim fso As FileSystemObject
    
    Set fso = New FileSystemObject
    If fso.FileExists(App.Path & "\Settings.json") Then
        'load JSON file
        temp_stringX = ReadTextFile(App.Path & "\Settings.json")
        'Decode file
        File_StringX = Base64DecodeString(temp_stringX)
        'Generate variables
        Set Setting_Object = JSON.parse(File_StringX)
        
        If con.State = adStateClosed Then
            connect
        End If
     Else
        MsgBox "Settings file is missing."
        Unload Me
    End If
'    connectC1 ("192.168.1.250")
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

Private Sub MDIForm_Unload(Cancel As Integer)
'    disconnectC1
    If StatusC1_1 Then
        CZKEM1.Beep 150
        DoEvents
        CZKEM1.Disconnect
    End If
    If StatusC1_2 Then
        CZKEM2.Beep 150
        DoEvents
        CZKEM2.Disconnect
    End If
    If StatusC1_3 Then
        CZKEM3.Beep 150
        DoEvents
        CZKEM3.Disconnect
    End If
    Dim Form As VB.Form
    For Each Form In VB.Forms
        Unload Form
    Next
End Sub

Private Sub tu_Click()
    Call changeForm(Form_User)
End Sub

Public Function ReadTextFile(sFilePath As String) As String
   On Error Resume Next
   
   Dim handle As Integer
   If LenB(Dir$(sFilePath)) > 0 Then
   
      handle = FreeFile
      Open sFilePath For Binary As #handle
      ReadTextFile = Space$(LOF(handle))
      Get #handle, , ReadTextFile
      Close #handle
      
   End If
   
End Function

