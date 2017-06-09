VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm FrmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Main Menu"
   ClientHeight    =   10635
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15060
   Icon            =   "x.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "x.frx":628A
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
            Picture         =   "x.frx":71A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "x.frx":768E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "x.frx":7940
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "x.frx":7C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "x.frx":7F95
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "x.frx":8351
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
      Width           =   15060
      _ExtentX        =   26564
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
      Begin VB.PictureBox picStretched 
         Height          =   255
         Left            =   10080
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   3
         Top             =   240
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
      Begin VB.Image Ori_Image 
         Height          =   375
         Left            =   11280
         Picture         =   "x.frx":2B585
         Stretch         =   -1  'True
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Menu p 
      Caption         =   "Penjualan"
      Begin VB.Menu pjl 
         Caption         =   "Trans. Penjualan"
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
      Begin VB.Menu pb 
         Caption         =   "Entri Pembelian"
      End
   End
   Begin VB.Menu a 
      Caption         =   "Admin"
      Begin VB.Menu tu 
         Caption         =   "User Manager"
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

Private Sub changeForm(new_form As Form)
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

Private Sub ebr_Click()
    Call changeForm(Form_List_barang)
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


Private Sub MDIForm_Resize()
Dim client_rect As Rect
Dim client_hwnd As Long

    picStretched.Move 0, 0, ScaleWidth, ScaleHeight

    ' Copy the original picture into picStretched.
'    picStretched.PaintPicture _
'        Ori_Image.Picture, _
'        0, 0, _
'        picStretched.ScaleWidth, _
'        picStretched.ScaleHeight, _
'        0, 0, _
'        Ori_Image.Width, _
'        Ori_Image.Height
    

    ' Set the MDI form's picture.
    'FrmMain.Picture = picStretched.Picture
    FrmMain.Picture = Ori_Image.Picture
    ' Invalidate the picture.
'    client_hwnd = FindWindowEx(Me.hWnd, 0, "MDIClient", _
'        vbNullChar)
'    GetClientRect client_hwnd, client_rect
'    InvalidateRect client_hwnd, client_rect, 1

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

Private Sub sp_Click()
    Call changeForm(Form_List_Supplier)
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
End Sub

Private Sub MDIForm_Unload(cancel As Integer)
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

