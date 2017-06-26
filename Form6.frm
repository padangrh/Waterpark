VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form_Entri_Barang 
   BackColor       =   &H0000C000&
   Caption         =   "Entri Data Barang"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8640
   ControlBox      =   0   'False
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Entri "
   ScaleHeight     =   5580
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_nama_supplier 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   3240
      Width           =   3855
   End
   Begin VB.ComboBox cb_kategori 
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
      Left            =   2160
      TabIndex        =   3
      Top             =   2040
      Width           =   3495
   End
   Begin MSComctlLib.ListView list_supplier 
      Height          =   2055
      Left            =   3120
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "KODE"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "NAMA SUPLIER"
         Object.Width           =   5999
      EndProperty
   End
   Begin VB.CommandButton btn_kategori 
      Appearance      =   0  'Flat
      Caption         =   "+"
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
      Left            =   5640
      TabIndex        =   14
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txt_kode_supplier 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   3240
      Width           =   735
   End
   Begin MSMask.MaskEdBox txt_jual 
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   2640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txt_nama 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   5175
   End
   Begin VB.TextBox txt_kode 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton btn_cancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton btn_save 
      Appearance      =   0  'Flat
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C000&
      Caption         =   "Supplier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C000&
      Caption         =   "Harga Jual"
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
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C000&
      Caption         =   "Kategori"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "Nama Barang"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "Kode Barang"
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
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "ENTRI dan UPDATE DATA BARANG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8415
   End
End
Attribute VB_Name = "Form_Entri_Barang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsbarang As New ADODB.Recordset
Dim txt_sup_toggle As Boolean

Private Sub btn_kategori_Click()
    Dim new_kategori As String
    new_kategori = InputBox("Kategori Baru: ", "Kategori")
    
    If new_kategori = "" Then
        Exit Sub
    End If
    
    Dim i As Integer
    i = 0
    Do While i < cb_kategori.ListCount
        If Trim(UCase(new_kategori)) = Trim(UCase(cb_kategori.List(i))) Then
            MsgBox "Kategori telah terdaftar"
            Exit Sub
        End If
        i = i + 1
    Loop
    
    cb_kategori.Text = new_kategori
    cb_kategori.AddItem (new_kategori)
    con.Execute ("insert into tbkategori values('" & new_kategori & "')")
End Sub

Private Sub cb_kategori_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_jual.SetFocus
    KeyAscii = validateKey(KeyAscii, 2)
End Sub

Private Sub Form_Activate()
    If txt_kode = "" Then
        txt_kode.SetFocus
    Else
        txt_nama.SetFocus
    End If
End Sub

Private Sub btn_Save_Click()
    Dim a As New ADODB.Recordset
    
    'kerjakan cek kategori
    
    If cek_Kategori = False Then
        MsgBox "Kategori tidak ditemukan"
        Exit Sub
    End If
    
    'kerjakan
    
    'If Trim(txt_kode.Text) = "" Or txt_nama.Text = "" Or txt_modal = "" Or txt_jual = "" Or txt_kode_supplier = "" Then
    If Trim(txt_kode.Text) = "" Or txt_nama.Text = "" Or txt_jual = "" Or txt_kode_supplier = "" Then
        MsgBox "Isi Data dengan Lengkap.....!"
        Exit Sub
    End If
    
    If getBarang(txt_kode) Then
        'disabled, hapus jumlah_akhir
        'con.Execute ("Update tbbarang set nama='" & txt_nama & "',kategori='" & cb_kategori.Text & "',harga_modal='" & Val(txt_modal) & "',harga_jual='" & Val(txt_jual) & "',kdsuplier='" & Val(txt_kode_supplier) & "',tgl_masuk='" & Format(dp_masuk, "yyyy-MM-dd") & "',ketahanan='" & Val(txt_ketahanan) & "', jumlah_akhir=" & Val(txt_stok) & " where kode='" & Trim(txt_kode) & "' ")
        'con.Execute ("Update tbbarang set nama='" & txt_nama & "', kategori='" & cb_kategori.Text & "', harga_modal = " & Val(txt_modal) & ", harga_jual = " & Val(txt_jual) & ", kdsuplier='" & Val(txt_kode_supplier) & "' where kode='" & Trim(txt_kode.Text) & "' ")
        con.Execute ("Update tbbarang set nama='" & txt_nama & "', kategori='" & cb_kategori.Text & "', harga_jual = " & Val(txt_jual) & ", kdsuplier='" & Val(txt_kode_supplier) & "' where kode='" & Trim(txt_kode.Text) & "' ")
    Else
        'disabled, hapus jumlah_akhir
        'con.Execute ("Insert into tbbarang values('" & Trim(txt_kode) & "' ,'" & txt_nama & "','" & cb_kategori.Text & "','" & Val(txt_modal) & "','" & Val(txt_jual) & "'," & Val(txt_stok) & ",'" & Val(txt_kode_supplier) & "','" & Format(dp_masuk, "yyyy-MM-dd") & "', '" & Val(txt_ketahanan) & "')")
        'con.Execute ("Insert into tbbarang values('" & Trim(txt_kode) & "' ,'" & txt_nama & "','" & cb_kategori.Text & "' ,'" & Val(txt_modal) & "','" & Val(txt_jual) & "','" & Val(txt_kode_supplier) & "')")
        con.Execute ("Insert into tbbarang values('" & Trim(txt_kode) & "' ,'" & txt_nama & "' ,'" & cb_kategori.Text & "' ,'" & Val(txt_jual) & "' ,'" & Val(txt_kode_supplier) & "')")
    
    End If
    kosongkan
    
    Form_List_barang.refreshlist
    Unload Me
End Sub

Sub kosongkan()
    txt_kode = ""
    txt_nama = ""
    cb_kategori.ListIndex = -1
    txt_kode_supplier = ""
    txt_nama_supplier = ""
    'txt_ketahanan = ""
    'txt_modal = 0
    txt_jual = 0
End Sub
Private Sub btn_Cancel_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Set rsbarang = con.Execute("select * from tbbarang")
    
    kosongkan
    'dp_masuk = Date
    txt_sup_toggle = True
    reload_Kategori
    reload_Supplier
    
    list_supplier.Visible = False
    
    'lbl_stok.Visible = isMaster
    'txt_stok.Visible = isMaster
End Sub

Private Function getBarang(kode As String) As Boolean
    'If rsbarang.EOF Or rsbarang.BOF Then
     '   getBarang = False
      '  Exit Function
    'End If
    
    Dim found As Boolean
    found = False
    Set rsbarang = con.Execute("select * from tbbarang")
    If Not rsbarang.EOF Then
        rsbarang.MoveFirst
        Do While Not rsbarang.EOF
            If rsbarang!kode = kode Then
                found = True
                Exit Do
            End If
            rsbarang.MoveNext
        Loop
    End If
    getBarang = found
End Function

Private Sub txt_kode_change()
    
    If getBarang(txt_kode) Then
        txt_nama = rsbarang!nama
        cb_kategori.Text = rsbarang!kategori
        'txt_modal.Text = rsbarang!harga_modal
        txt_jual = rsbarang!harga_jual
        txt_kode_supplier.Text = rsbarang!kdsuplier
        Set rsSupplier = con.Execute("select * from tbsuplier")
        If getSupplier(rsbarang!kdsuplier) Then
            txt_nama_supplier = rsSupplier!nmsuplier
        End If
        'txt_ketahanan.Text = rsbarang!ketahanan
        'dp_masuk = rsbarang!tgl_masuk
        'txt_stok = rsbarang!jumlah_akhir
    Else
        txt_nama.Text = ""
        cb_kategori.ListIndex = -1
        'txt_modal = 0
        txt_jual = 0
        txt_kode_supplier = ""
        txt_nama_supplier = ""
        'txt_ketahanan = ""
        'txt_stok = 0
    End If
    txt_sup_toggle = True

End Sub

Private Sub txt_kode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txt_nama.SetFocus
    End If
End Sub

Private Sub txt_kode_supplier_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 2)
End Sub

Private Sub txt_nama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cb_kategori.SetFocus
    KeyAscii = validateKey(KeyAscii, 3)
End Sub

Private Sub txt_jual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt_kode_supplier.SetFocus
    KeyAscii = validateKey(KeyAscii, 1)
End Sub

Private Sub txt_kode_supplier_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        txt_sup_toggle = True
        
        Set rsSupplier = con.Execute("select * from tbsuplier")
        
        If getSupplier(txt_kode_supplier) Then
            txt_nama_supplier.Text = rsSupplier!nmsuplier
            'txt_ketahanan.SetFocus
            btn_save.SetFocus
        Else
            MsgBox "Supplier tidak terdaftar"
            txt_kode_supplier.Text = ""
        End If
    Else
        txt_nama_supplier = ""
    End If
End Sub

Private Sub txt_nama_supplier_Change()
    If txt_nama_supplier.Text <> "" And txt_sup_toggle = False Then
        list_supplier.Visible = True
        reload_Supplier
    Else
        list_supplier.Visible = False
        txt_sup_toggle = False
    End If
End Sub

Private Sub txt_nama_supplier_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 3)
End Sub

'Private Sub txt_ketahanan_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Val(txt_ketahanan) > 0 Then
'        btn_save.SetFocus
'    Else
'        MsgBox ("Ketahanan barang tidak valid")
'    End If
'End If
'End Sub

Private Sub txt_nama_supplier_LostFocus()
    If Not Me.ActiveControl Is Nothing Then
        If Not Me.ActiveControl.Name = "list_supplier" Then
            list_supplier.Visible = False
        End If
    End If
End Sub

Private Sub txt_nama_supplier_KeyDown(key As Integer, Shift As Integer)
    
    If key = 40 Then
        list_supplier.Visible = True
        list_supplier.SetFocus
        Exit Sub
    ElseIf key = 13 And list_supplier.Visible = True Then
        'txt_kode_supplier = list_supplier.ListItems(0).Text
        'txt_nama_supplier = list_supplier.ListItems(0).SubItems(1)
        list_supplier.SetFocus
    ElseIf key = 13 And list_supplier.Visible = False And txt_kode_supplier.Text <> "" Then
        btn_save.SetFocus
    Else
        txt_kode_supplier = ""
    End If
End Sub

Private Sub list_supplier_LostFocus()
    list_supplier.Visible = False
End Sub

Private Sub list_supplier_dblclick()
    If list_supplier.SelectedItem.index >= 0 Then
        txt_kode_supplier = list_supplier.SelectedItem.Text
        txt_nama_supplier = list_supplier.SelectedItem.SubItems(1)
        'txt_ketahanan.SetFocus
        btn_save.SetFocus
    End If
End Sub

Private Sub list_supplier_keydown(key As Integer, Shift As Integer)
    If key = 13 Then
        list_supplier_dblclick
    End If
End Sub

Private Sub reload_Supplier()
    'list_supplier.Visible = True
    list_supplier.ListItems.Clear
    Dim rsSup As ADODB.Recordset
    Set rsSup = con.Execute("select * from tbsuplier where nmsuplier like '%" & txt_nama_supplier & "%'")
    If rsSup.EOF Then
        list_supplier.Visible = False
        Exit Sub
    End If
    
    rsSup.MoveFirst
    Do While Not rsSup.EOF
        list_supplier.ListItems.Add(, , rsSup!kdsuplier).SubItems(1) = rsSup!nmsuplier
        rsSup.MoveNext
    Loop
    
    Set rsSup = Nothing
End Sub

Private Sub reload_Kategori()
    Dim rsKategori As ADODB.Recordset
    Set rsKategori = con.Execute("select * from tbkategori")
    If Not (rsKategori.EOF Or rsKategori.BOF) Then
        rsKategori.MoveFirst
        Do While Not rsKategori.EOF
            cb_kategori.AddItem (rsKategori!kode)
            rsKategori.MoveNext
        Loop
    End If
End Sub

Private Function cek_Kategori() As Boolean
    cek_Kategori = False
    Dim i As Integer
    Do While i < cb_kategori.ListCount
        If Trim(UCase(cb_kategori.Text)) = Trim(UCase(cb_kategori.List(i))) Then
            cek_Kategori = True
        End If
        i = i + 1
    Loop
End Function
