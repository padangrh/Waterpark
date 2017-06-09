VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_List_RFID 
   BackColor       =   &H0080C0FF&
   Caption         =   "RFID Manager"
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   11730
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
      Left            =   1920
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CheckBox chk_Nonaktif 
      BackColor       =   &H0080C0FF&
      Caption         =   "Tampilkan non-aktif"
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
      Left            =   7080
      TabIndex        =   7
      Top             =   1080
      Value           =   1  'Checked
      Width           =   3735
   End
   Begin VB.CommandButton btn_Aktivasi 
      BackColor       =   &H000080FF&
      Caption         =   "Deaktivasi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   1575
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
      Left            =   3960
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton btn_Tambah 
      BackColor       =   &H000080FF&
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton btn_Hapus 
      BackColor       =   &H000080FF&
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   1575
   End
   Begin MSComctlLib.ListView lv_RFID 
      Height          =   4815
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   9675
      _ExtentX        =   17066
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode RFID"
         Object.Width           =   5644
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
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListOrder 
      Left            =   3000
      Top             =   240
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
            Picture         =   "Form_List_RFID.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_List_RFID.frx":037B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Daftar Kartu RFID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Width           =   11775
   End
   Begin VB.Label Label1 
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
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "Form_List_RFID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Aktivasi_Click()
    If MsgBox("Ubah status RFID " & lv_RFID.SelectedItem.Text & "?", vbYesNo, "Konfirmasi " & btn_Aktivasi.Caption) = vbYes Then
        'con.Execute ("insert into tbnonaktif values('" & lv_RFID.SelectedItem.Text & "','" & Format(lv_RFID.SelectedItem.SubItems(1), "yyyy-mm-dd") & "','" & lv_RFID.SelectedItem.SubItems(2) & "','" & lv_RFID.SelectedItem.SubItems(3) & "','" & btn_Aktivasi.Caption & "-list rfid','" & username & "')")
        Call backupAktif(lv_RFID.SelectedItem.Text, btn_Aktivasi.Caption & " - list RFID")
        If lv_RFID.SelectedItem.SubItems(3) = 1 Then
            con.Execute ("update tbaktif set status = '0' where rfid = '" & lv_RFID.SelectedItem.Text & "'")
        ElseIf lv_RFID.SelectedItem.SubItems(3) = 0 Then
            con.Execute ("update tbaktif set tanggal = '" & Format(Now, "yyyy-mm-dd") & "', jam = '" & Format(Now, "hh:mm:ss") & "', status = '1', keterangan = '" & username & "' where rfid = '" & lv_RFID.SelectedItem.Text & "'")
        End If
        reload_list
    End If
End Sub

Private Sub btn_Hapus_Click()
    If lv_RFID.ListItems.count > 0 And MsgBox("Yakin akan menghapus RFID " & lv_RFID.SelectedItem.Text & " ?", vbYesNo, "Konfirmasi Hapus") = vbYes Then
        con.Execute ("insert into tbnonaktif values('" & lv_RFID.SelectedItem.Text & "','" & Format(lv_RFID.SelectedItem.SubItems(1), "yyyy-mm-dd") & "','" & lv_RFID.SelectedItem.SubItems(2) & "','" & lv_RFID.SelectedItem.SubItems(3) & "','dihapus-list rfid','" & username & "')")
        con.Execute ("Delete from tbaktif where rfid = '" & lv_RFID.SelectedItem.Text & "'")
        reload_list
    End If
End Sub

Private Sub cb_Search_Click()
    txt_Search.Text = ""
End Sub

Private Sub chk_Nonaktif_Click()
    reload_list
End Sub

Private Sub Form_Load()
    reload_list
    cb_Search.Clear
    Dim x As Integer
    For x = 1 To lv_RFID.ColumnHeaders.count
        cb_Search.AddItem (lv_RFID.ColumnHeaders(x).Text)
    Next
    cb_Search.ListIndex = 0
End Sub

Sub reload_list()
    lv_RFID.ListItems.Clear
    Dim rsRFID As ADODB.Recordset
    Dim query As String
    Dim aitem As ListItem
    ''query = "Select a.rfid as rfid,a.tanggal as tanggal,a.jam as jam, a.status as status, b.nobukti as nobukti from tbaktif a, tbrfid b where a.rfid = b.rfid"
    
    'old query
'    query = "Select a.rfid as rfid,a.tanggal as tanggal,a.jam as jam, a.status as status, c.nobukti as nobukti from tbaktif a left join( select b.rfid, max(b.nobukti) as nobukti from tbrfid b group by b.rfid ) c on a.rfid = c.rfid"
'    If txt_Search.Text <> "" Then
'        Select Case cb_Search.ListIndex
'            Case 0
'                query = query & " where a.rfid like '%" & txt_Search.Text & "%'"
'            Case 1
'                query = query & " where cast(a.tanggal as char(20)) like '%" & txt_Search.Text & "%'"
'            Case 2
'                query = query & " where a.jam like '%" & txt_Search.Text & "%'"
'            Case 3
'                query = query & " where a.status like '%" & txt_Search.Text & "%'"
'            Case 4
'                query = query & " where c.nobukti like '%" & txt_Search.Text & "%'"
'        End Select
'    End If
'    '"and a.rfid like '%" & txt_Search.Text & "%'"
'    If chk_Nonaktif.Value = 0 Then
'        If InStr(1, query, "where") > 0 Then
'            query = query & " and "
'        Else
'            query = query & " where "
'        End If
'        query = query & "a.status = '1'"
'    End If
'    query = query & " order by c.nobukti desc"
    
    query = "Select * from tbaktif"
    If txt_Search.Text <> "" Then
        Select Case cb_Search.ListIndex
            Case 0
                query = query & " where rfid like '%" & txt_Search.Text & "%'"
            Case 1
                query = query & " where cast(a.tanggal as char(20)) like '%" & txt_Search.Text & "%'"
            Case 2
                query = query & " where jam like '%" & txt_Search.Text & "%'"
            Case 3
                query = query & " where status like '%" & txt_Search.Text & "%'"
            Case 4
                query = query & " where keterangan like '%" & txt_Search.Text & "%'"
        End Select
    End If
    '"and a.rfid like '%" & txt_Search.Text & "%'"
    If chk_Nonaktif.Value = 0 Then
        If InStr(1, query, "where") > 0 Then
            query = query & " and "
        Else
            query = query & " where "
        End If
        query = query & "status = '1'"
    End If
    query = query & " order by keterangan desc"
    
    
    Set rsRFID = con.Execute(query)
    
    Do While Not rsRFID.EOF
        Set aitem = lv_RFID.ListItems.Add(, , rsRFID!rfid)
        aitem.SubItems(1) = Format(rsRFID!tanggal, "yyyy-mm-dd")
        aitem.SubItems(2) = rsRFID!jam
        aitem.SubItems(3) = rsRFID!status
        If IsNull(rsRFID!keterangan) Then
            aitem.SubItems(4) = ""
        Else
            aitem.SubItems(4) = rsRFID!keterangan
        End If
        rsRFID.MoveNext
    Loop
    
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

Private Sub lv_RFID_DblClick()
    If Not (lv_RFID.SelectedItem Is Nothing) Then
        Form_EditRFID.Show
        Form_EditRFID.Init lv_RFID.SelectedItem.Text, lv_RFID.SelectedItem.SubItems(1), lv_RFID.SelectedItem.SubItems(2), lv_RFID.SelectedItem.SubItems(3)
    End If
End Sub

Private Sub lv_RFID_ItemClick(ByVal item As MSComctlLib.ListItem)
    If lv_RFID.SelectedItem.SubItems(3) = 1 Then
        btn_Aktivasi.Caption = "Deaktivasi"
    ElseIf lv_RFID.SelectedItem.SubItems(3) = 0 Then
        btn_Aktivasi.Caption = "Aktivasi"
    End If
End Sub

Private Sub txt_Search_Change()
    reload_list
End Sub

Private Sub btn_Tambah_Click()
    Dim temp_RFID As String
    temp_RFID = ""
    temp_RFID = InputBox("Masukkan RFID", "Tambah RFID")
    If temp_RFID = "" Then
        Exit Sub
    End If
    
    If Len(temp_RFID) = 10 Then
        Dim flagY As Boolean
        flagY = False
'        Dim rsCheck As ADODB.Recordset
'        Set rsCheck = con.Execute("Select * from tbaktif")
'        Do While Not rsCheck.EOF
'            If rsCheck!rfid = temp_RFID Then
'                flagY = True
'                Exit Do
'            End If
'            rsCheck.MoveNext
'        Loop
        flagY = isInTBAktif(temp_RFID)
        If flagY = False Then
            con.Execute ("Insert into tbaktif values ('" & temp_RFID & "','" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','1','" & username & "')")
            reload_list
        Else
            MsgBox ("Kartu sudah terdaftar")
        End If
    Else
        MsgBox ("RFID tidak valid")
    End If
End Sub

Private Sub txt_Search_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 97 To 122, 8, 45, 58, 13 ' A-Z, 0-9, a-z, backspace, -, :, enter
        'Let these key codes pass through
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub
