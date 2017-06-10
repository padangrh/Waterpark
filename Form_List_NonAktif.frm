VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_List_NonAktif 
   Caption         =   "History"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   14130
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chk_Sampai 
      Height          =   255
      Left            =   10320
      TabIndex        =   5
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txt_Search 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   1920
      Width           =   2895
   End
   Begin VB.ComboBox cb_Search 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1680
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1935
   End
   Begin MSComctlLib.ListView lv_RFID 
      Height          =   4815
      Left            =   960
      TabIndex        =   0
      Top             =   2880
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImageListOrder"
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode RFID"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tanggal"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Jam"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Keterangan"
         Object.Width           =   6456
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Login"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListOrder 
      Left            =   3720
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   11
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_List_NonAktif.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_List_NonAktif.frx":037B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dt_start 
      Height          =   495
      Left            =   8160
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   95092739
      CurrentDate     =   42810
   End
   Begin MSComCtl2.DTPicker dt_end 
      Height          =   495
      Left            =   11760
      TabIndex        =   7
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   95092739
      CurrentDate     =   42810
   End
   Begin VB.Label Label4 
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Sampai"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   1
      Top             =   480
      Width           =   8055
   End
End
Attribute VB_Name = "Form_List_NonAktif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cb_Search_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 97 To 122, 8 ' A-Z, 0-9, a-z and backspace
        'Let these key codes pass through
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub

Private Sub chk_Sampai_Click()
    If chk_Sampai.Value = 0 Then
        Label3.Enabled = False
        dt_end.Enabled = False
    Else
        Label3.Enabled = True
        dt_end.Enabled = True
    End If
    reload_list
End Sub

Private Sub Command1_Click()
    Dim x As Integer
    For x = 1 To lv_RFID.ColumnHeaders.count
        MsgBox lv_RFID.ColumnHeaders(x).Text & " " & lv_RFID.ColumnHeaders(x).Width
    Next
End Sub

Private Sub dt_end_Change()
    reload_list
End Sub

Private Sub dt_end_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 97 To 122, 8 ' A-Z, 0-9, a-z and backspace
        'Let these key codes pass through
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub

Private Sub dt_start_Change()
    reload_list
End Sub

Private Sub dt_start_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 97 To 122, 8 ' A-Z, 0-9, a-z and backspace
        'Let these key codes pass through
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub

Private Sub Form_Load()
    Me.Top = 615
    Me.Left = 2385
    cb_Search.Clear
    Dim x As Integer
    For x = 1 To lv_RFID.ColumnHeaders.count
        cb_Search.AddItem (lv_RFID.ColumnHeaders(x).Text)
    Next
    cb_Search.ListIndex = 0
    dt_start.Value = Format(Now, "yyyy-MM-dd")
    dt_end.Value = Format(Now, "yyyy-MM-dd")
    reload_list
End Sub

Private Sub lv_RFID_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lv_RFID.Sorted = True
    Dim i As Byte
    For i = 1 To lv_RFID.ColumnHeaders.count
      lv_RFID.ColumnHeaders.item(i).Icon = 0
    Next
    If lv_RFID.SortKey <> ColumnHeader.index - 1 Then
      lv_RFID.SortOrder = lvwAscending
      ColumnHeader.Icon = 1
      lv_RFID.SortKey = ColumnHeader.index - 1
    Else
      If lv_RFID.SortOrder = lvwAscending Then
        lv_RFID.SortOrder = lvwDescending
        ColumnHeader.Icon = 2
      Else
        lv_RFID.SortOrder = lvwAscending
        ColumnHeader.Icon = 1
      End If
    End If
End Sub

Private Sub txt_Search_Change()
    reload_list
End Sub

'    Dim rsNonAktif As ADODB.Recordset
'    Dim bitem As ListItem
'    Set rsNonAktif = con.Execute("Select * from tbnonaktif")
'    If Not rsNonAktif.EOF Then
'        Do While Not rsNonAktif.EOF
'            Set bitem = lv_RFID.ListItems.Add(, , rsNonAktif!rfid)
'            bitem.SubItems(1) = rsNonAktif!tanggal
'            bitem.SubItems(2) = rsNonAktif!jam
'            bitem.SubItems(3) = rsNonAktif!status
'            bitem.SubItems(4) = rsNonAktif!keterangan
'            bitem.SubItems(5) = rsNonAktif!login
'            rsNonAktif.MoveNext
'        Loop
'    End If
'    Set rsNonAktif = Nothing

Sub reload_list()
    lv_RFID.ListItems.Clear
    Dim rsRFID As ADODB.Recordset
    Dim query As String
    Dim aitem As ListItem
    query = "Select * from tbnonaktif"
    
    If chk_Sampai.Value = 0 Then
        query = query & " where tanggal = '" & Format(dt_start.Value, "yyyy-MM-dd") & "'"
    Else
        query = query & " where tanggal >= '" & Format(dt_start.Value, "yyyy-MM-dd") & "' and tanggal <= '" & Format(dt_end.Value, "yyyy-MM-dd") & "'"
    End If
    
    If txt_Search.Text <> "" Then
        Select Case cb_Search.ListIndex
            Case 0
                query = query & " and rfid like '%" & txt_Search.Text & "%'"
            Case 1
                query = query & " and cast(tanggal as char(20)) like '%" & txt_Search.Text & "%'"
            Case 2
                query = query & " and jam like '%" & txt_Search.Text & "%'"
            Case 3
                query = query & " and status like '%" & txt_Search.Text & "%'"
            Case 4
                query = query & " and keterangan like '%" & txt_Search.Text & "%'"
        End Select
    End If
    
    query = query & " order by keterangan desc"
    
    Set rsRFID = con.Execute(query)
    
    Do While Not rsRFID.EOF
        Set aitem = lv_RFID.ListItems.Add(, , rsRFID!rfid)
        aitem.SubItems(1) = Format(rsRFID!tanggal, "yyyy-mm-dd")
        aitem.SubItems(2) = rsRFID!jam
        aitem.SubItems(3) = rsRFID!status
        aitem.SubItems(4) = rsRFID!keterangan
        aitem.SubItems(5) = rsRFID!login
        rsRFID.MoveNext
    Loop
End Sub

Private Sub txt_Search_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 97 To 122, 8 ' A-Z, 0-9, a-z and backspace
        'Let these key codes pass through
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub
