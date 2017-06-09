VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_ReplaceRFID 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Kartu Hilang"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_Close 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Tutup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton cmd_Search 
      Caption         =   "Cari"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton btn_Tambah 
      BackColor       =   &H00FF8080&
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
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton btn_Hapus 
      BackColor       =   &H00FF8080&
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
      Height          =   615
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton btn_Save 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Simpan"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton btn_Cancel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Batal"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8640
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
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   4335
   End
   Begin MSComctlLib.ListView lv_RFID 
      Height          =   4815
      Left            =   720
      TabIndex        =   6
      Top             =   2760
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nomor"
         Object.Width           =   2219
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kode RFID"
         Object.Width           =   5917
      EndProperty
   End
   Begin VB.Label lbl_Nobukti 
      BackStyle       =   0  'Transparent
      Caption         =   "Kosong"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Tiket : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lbl_JumlahTiket 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari berdasarkan Tiket/RFID : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "Form_ReplaceRFID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Cancel_Click()
    kosongkan
End Sub

Private Sub btn_Hapus_Click()
    If lv_RFID.ListItems.count > 0 Then
        lv_RFID.ListItems.Remove (lv_RFID.SelectedItem.index)
        Dim l As Integer
        For l = 1 To lv_RFID.ListItems.count
            lv_RFID.ListItems.item(l).Text = l
        Next
    End If
End Sub

Private Sub btn_Save_Click()
    If lv_RFID.ListItems.count <> lbl_JumlahTiket.Caption Then
        MsgBox ("Jumlah tiket dan RFID tidak sama")
        Exit Sub
    End If
    
    Dim m As Integer
    For m = 1 To lv_RFID.ListItems.count
        If RFIDinUse(lv_RFID.ListItems(m).SubItems(1)) = True Then
            MsgBox ("RFID " & lv_RFID.ListItems(m).SubItems(1) & " sedang digunakan." & vbNewLine & "Gunakan kartu RFID lain")
            Exit Sub
        End If
    Next
    

    Dim rsCompare As ADODB.Recordset
    m = lv_RFID.ListItems.count
    Dim flagX() As Integer
    Dim foundX As Boolean
    ReDim Preserve flagX(m)
    Dim perubahan As Integer
    perubahan = 0
    Set rsCompare = con.Execute("Select * from tbrfid where nobukti = '" & lbl_Nobukti.Caption & "'")
    Do While Not rsCompare.EOF
        foundX = False
        For m = 1 To lv_RFID.ListItems.count
            If rsCompare!rfid = lv_RFID.ListItems(m).SubItems(1) Then
                foundX = True
                Exit For
            End If
        Next
        If foundX = False Then
            perubahan = perubahan + 1
        End If
        rsCompare.MoveNext
    Loop
    Form_Print.init_ReplaceRFID perubahan, lbl_Nobukti.Caption
    Form_Print.Show vbModal, Me
    
    
    
'    If konfirmasi_Perubahan(perubahan) = False Then
'        Set rsCompare = Nothing
'        MsgBox ("Perubahan dibatalkan")
'        Exit Sub
'    End If

'    rsCompare.MoveFirst
'    Do While Not rsCompare.EOF
'        foundX = False
'        For m = 1 To lv_RFID.ListItems.count
'            If rsCompare!rfid = lv_RFID.ListItems(m).SubItems(1) Then
'                flagX(m) = 1
'                foundX = True
'                Exit For
'            End If
'        Next
'        If foundX = False Then
'            con.Execute ("Delete from tbrfid where nobukti = '" & lbl_Nobukti.Caption & "' and rfid = '" & rsCompare!rfid & "'")
'            Call backupAktif(rsCompare!rfid, "perubahan kartu hilang - ReplaceRFID")
'            con.Execute ("delete from tbaktif where rfid = '" & rsCompare!rfid & "'")
'        End If
'        rsCompare.MoveNext
'    Loop
'    Set rsCompare = con.Execute("select tanggal, jam, nobukti from bill where nobukti = '" & lbl_Nobukti.Caption & "'")
'    For m = 1 To lv_RFID.ListItems.count
'        If flagX(m) = 0 Then
'
'            con.Execute ("insert into tbaktif values('" & lv_RFID.ListItems(m).SubItems(1) & "','" & Format(rsCompare!tanggal, "yyyy-mm-dd") & "','" & rsCompare!jam & "','1','" & rsCompare!nobukti & "')")
'            con.Execute ("insert into tbrfid values('" & lbl_Nobukti.Caption & "','" & lv_RFID.ListItems(m).SubItems(1) & "')")
'        End If
'    Next
    
'    Printer.Font = "Times new roman"
'    Printer.FontSize = 12
'    Printer.Print Tab(6);
'    Printer.Print Tab(4); Format(Now, "dd-MM-yyyy  hh:mm:ss");
'    Printer.Print Tab(4); "No Faktur"; Tab(18); ": "; Form_Print.txt_bon
'    Printer.Print Tab(4); "Supervisor"; Tab(18); ": "; txt_spv
'    Printer.Print Tab(4); "Status"; Tab(18); ": "; cb_status.Text
'    Printer.Print Tab(4); "Customer"; Tab(18); ": "; txt_customer
'    Printer.Print Tab(4); "Diskon"; Tab(18); ": Rp."; txt_diskon
'    Printer.EndDoc
'
'    Set rsCompare = Nothing
'    Printer.Font = "dotumche"
'
'    Printer.FontSize = 18
'    Printer.FontBold = True
'    Printer.Print Tab(2); Setting_Object("Toko");
'
'    Printer.FontSize = 10
'    Printer.FontBold = False
'    Printer.Print Tab(1); "                                                            "; 'baris kosong
'    Printer.Print Tab(1); Setting_Object("Alamat1");
'    Printer.Print Tab(1); "No. FAKTUR : "; lbl_Nobukti.Caption
'    Printer.Print Tab(1); Format(Now, "dd-MM-yyyy HH:mm:ss"); "  ";
'    Printer.Print Tab(1); "Nama Kasir : "; username;
'    Printer.Print Tab(1); "                                                                  ";
'    Printer.Print Tab(1); "------------------------------------------------------------------";
''    Do While Not rsJual.EOF
''        Printer.Print Tab(1); rsJual!nama_Barang
''        Dim bayar As Long
''        bayar = Val(rsJual!jumlah_jual) * Val(rsJual!harga_jual)
''        Printer.Print Tab(2); rsJual!jumlah_jual; Tab(9); "x"; Tab(21 - Len(Format(rsJual!harga_jual, "###,###,##0"))); Format(rsJual!harga_jual, "###,###,##0"); Tab(35 - Len(Format(bayar, "###,###,##0"))); Format(bayar, "###,###,##0")
''        rsJual.MoveNext
''    Loop
''    Printer.Print Tab(1); "                                                                  ";
''    If priceToNum(txt_diskon) > 0 Then
''        Printer.Print Tab(1); "Diskon"; Tab(20); "Rp."; Tab(35 - Len(Format(txt_diskon, "###,###,##0"))); Format(txt_diskon, "###,###,##0")
''    End If
'    Printer.Print Tab(1); "------------------------------------------------------------------";
'    Printer.FontSize = 12
'    Printer.Print Tab(2); "Grand Total"; Tab(15); "Rp."; Tab(30 - Len(Format(txt_grandTotal, "###,###,##0"))); Format(txt_grandTotal, "###,###,##0")
'
'    Printer.CurrentX = 0
'    Printer.FontSize = 10
'    Printer.Print Tab(3); "                                                             ";
'    If (tunai = 1) Then
'        Printer.Print Tab(1); "Jumlah Uang"; Tab(15); "Rp. "; Tab(31 - Len(Format(txt_uang, "###,###,##0"))); Format(txt_uang, "###,###,##0");
'        Printer.Print Tab(1); "Kembalian  "; Tab(15); "Rp. "; Tab(31 - Len(Format(txt_kembali, "###,###,##0"))); Format(txt_kembali, "###,###,##0");
'    Else
'        Printer.Print Tab(3); "-NON TUNAI-";
'    End If
'    Printer.Print Tab(1); "                                                             ";
'    Printer.Print Tab(1); "                                                             ";
'    Printer.Print Tab((38 - Len("*Simpanlah struk ini*")) / 2); "*Simpanlah struk ini*";
'    Printer.Print Tab((38 - Len("*hingga anda meninggalkan lokasi*")) / 2); "*hingga anda meninggalkan lokasi*";
'
'    Printer.EndDoc
'
'    MsgBox ("Data berhasil disimpan")
    kosongkan
End Sub

Private Sub btn_Tambah_Click()
    Dim temp_RFID As String
    temp_RFID = ""
    temp_RFID = InputBox("Masukkan RFID", "Tambah RFID")
    If temp_RFID = "" Then
        Exit Sub
    End If
    If Len(temp_RFID) = 10 Then
        Dim n As Integer
        Dim kitem As ListItem
        For n = 1 To lv_RFID.ListItems.count
            If temp_RFID = lv_RFID.ListItems(n).SubItems(1) Then
                MsgBox ("RFID sudah terisi")
                Exit Sub
            End If
        Next
        Set kitem = lv_RFID.ListItems.Add(, , lv_RFID.ListItems.count + 1)
        kitem.SubItems(1) = temp_RFID
    Else
        MsgBox ("RFID tidak valid")
    End If
End Sub

Private Sub cmd_Close_Click()
    Unload Me
'    MsgBox (Me.Top & " " & Me.Left)
End Sub

Private Sub cmd_Search_Click()
    Dim rsCekRFID As ADODB.Recordset
    If Len(txt_Search.Text) = 10 Then
        Set rsCekRFID = con.Execute("Select a.* from tbrfid a, bill b where a.nobukti = b.nobukti and left(a.nobukti,1) <> 'R' and rfid = '" & txt_Search.Text & "' order by concat(b.tanggal, ' ', b.jam) desc ")

    Else
        Set rsCekRFID = con.Execute("Select * from bill where nobukti = '" & txt_Search.Text & "' and left(nobukti,1) <> 'R'")
      
    End If
    
    If Not rsCekRFID.EOF Then
        lbl_Nobukti.Caption = rsCekRFID!nobukti
        Dim rsRFID As ADODB.Recordset
        Dim litem As ListItem
        lv_RFID.ListItems.Clear
        Set rsRFID = con.Execute("select * from tbrfid where nobukti = '" & lbl_Nobukti.Caption & "'")
        Do While Not rsRFID.EOF
            Set litem = lv_RFID.ListItems.Add(, , lv_RFID.ListItems.count + 1)
            litem.SubItems(1) = rsRFID!rfid
            rsRFID.MoveNext
        Loop
        Set rsRFID = Nothing
        lbl_JumlahTiket.Caption = lv_RFID.ListItems.count
        lv_RFID.Enabled = True
        btn_Save.Enabled = True
        btn_Cancel.Enabled = True
        btn_Hapus.Enabled = True
        btn_Tambah.Enabled = True
    Else
        MsgBox ("Data tidak ditemukan")
        kosongkan
    End If
    
    Set rsCekRFID = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    ElseIf KeyCode = 46 Then
        btn_Hapus_Click
    End If
End Sub

Private Function RFIDinUse(noRFID As String) As Boolean
    RFIDinUse = False
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = con.Execute("Select a.rfid as rfid,a.tanggal as tanggal,a.jam as jam, a.status as status, c.nobukti as nobukti from tbaktif a left join( select b.rfid, max(b.nobukti) as nobukti from tbrfid b group by b.rfid ) c on a.rfid = c.rfid where a.status = '1' and c.nobukti <> '" & lbl_Nobukti.Caption & "' and a.rfid = '" & noRFID & "'")
    If Not rsTemp.EOF Then RFIDinUse = True
    Set rsTemp = Nothing
End Function

'Private Function konfirmasi_Perubahan(jumlah As Integer) As Boolean
'    konfirmasi_Perubahan = False
'    If jumlah > 0 Then
'        Dim rsbarang As ADODB.Recordset
'        Set rsbarang = con.Execute("Select * from tbbarang where kode = '2'")
'        If Not rsbarang.EOF Then
'
'            If MsgBox("Terdapat " & jumlah & " perubahan." & vbNewLine & "Biaya perubahan RFID : " & Format(rsbarang!harga_jual * jumlah, "###,###,##0"), vbYesNo, "Konfirmasi") = vbYes Then
'                konfirmasi_Perubahan = True
'            End If
'        End If
'        Set rsbarang = Nothing
'    Else
'        MsgBox ("Tidak terdapat perubahan")
'    End If
'End Function

Private Sub Form_Load()
    Me.Top = 330
    Me.Left = 7080
End Sub

Private Sub txt_Search_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 97 To 122, 8 ' A-Z, 0-9, a-z and backspace
        'Let these key codes pass through
        Case 13
            If Len(txt_Search.Text) >= 20 Then
                txt_Search.Text = Right(txt_Search.Text, 10)
                txt_Search.SelStart = Len(txt_Search.Text)
            End If
            cmd_Search_Click
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub

Private Sub kosongkan()
    lv_RFID.ListItems.Clear
    txt_Search.Text = ""
    lbl_Nobukti.Caption = "Kosong"
    lbl_JumlahTiket.Caption = 0
    lv_RFID.Enabled = False
    btn_Tambah.Enabled = False
    btn_Hapus.Enabled = False
    btn_Save.Enabled = False
    btn_Cancel.Enabled = False
    txt_Search.SetFocus
End Sub
