VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_Deposit 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Ambil Deposit"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form11"
   ScaleHeight     =   13046.83
   ScaleMode       =   0  'User
   ScaleWidth      =   14399.27
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_Tambah 
      Caption         =   "Tambah"
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
      Left            =   7680
      TabIndex        =   18
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton btn_Hapus 
      Caption         =   "Hapus"
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
      Left            =   9480
      TabIndex        =   17
      Top             =   3720
      Width           =   1335
   End
   Begin MSComctlLib.ListView lv_jual 
      Height          =   5655
      Left            =   840
      TabIndex        =   12
      Top             =   4560
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   9975
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nomor"
         Object.Width           =   2515
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "RFID"
         Object.Width           =   10265
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Deposit"
         Object.Width           =   4710
      EndProperty
   End
   Begin VB.TextBox txt_jumlah 
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
      Height          =   614
      Left            =   4080
      TabIndex        =   11
      Text            =   "12345678901234"
      Top             =   2280
      Width           =   1239
   End
   Begin VB.TextBox txt_harga 
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
      Left            =   960
      TabIndex        =   10
      Text            =   "12345678901234"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txt_Nama 
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
      Left            =   3240
      TabIndex        =   7
      Text            =   "12345678901234"
      Top             =   3720
      Width           =   3975
   End
   Begin VB.TextBox txt_total 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   3
      Text            =   "00.000.000"
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   16
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0FF&
      Caption         =   "[F1 -> Print]  [Delete -> Hapus 1 baris]  [Shift + Delete -> Hapus Semua]  [F4 -> Tutup]"
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
      Left            =   0
      TabIndex        =   13
      Top             =   10320
      Width           =   12135
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah"
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
      Left            =   4080
      TabIndex        =   9
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Harga"
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
      Left            =   960
      TabIndex        =   8
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "RFID"
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
      Left            =   840
      TabIndex        =   6
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label lbl_user 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Richard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   5
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Kasir:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   4
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6240
      TabIndex        =   2
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Label lbl_faktur 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "A123456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "No. Faktur:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0FF&
      Height          =   1695
      Left            =   360
      TabIndex        =   14
      Top             =   1560
      Width           =   11355
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0FF&
      Height          =   1095
      Left            =   360
      TabIndex        =   15
      Top             =   120
      Width           =   11355
   End
End
Attribute VB_Name = "Form_Deposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsbarang As ADODB.Recordset
Dim hargaTiket As Long
Public is_new As Boolean

Private Sub btn_Hapus_Click()
    If lv_jual.ListItems.count > 0 Then
        lv_jual.ListItems.Remove (lv_jual.SelectedItem.index)
        Dim j As Integer
        For j = 1 To lv_jual.ListItems.count
            lv_jual.ListItems.item(j).Text = j
        Next
        txt_jumlah.Text = lv_jual.ListItems.count
        txt_total.Text = Format(txt_jumlah * txt_harga, "###,###,##0")
    End If
'    MsgBox lv_jual.ColumnHeaders(1).Width & " " & lv_jual.ColumnHeaders(2).Width & " " & lv_jual.ColumnHeaders(3).Width
End Sub

Private Sub btn_Tambah_Click()
    Call tambah
End Sub

Private Sub Form_Load()
    lbl_user = username
    txt_total = 0
    txt_jumlah.Text = 0
    is_new = True
    Dim rsHarga As ADODB.Recordset
    Set rsHarga = con.Execute("select * from tbbarang where kode = '2'")
    hargaTiket = CLng(rsHarga!harga_jual)
    txt_harga.Text = hargaTiket
    kosongkan
    Dim namafile, file_data, huruf As String
    Dim angka As Long
    namafile = App.Path & "\fakturdeposit.txt"
    Open namafile For Input As #1
    While Not EOF(1)
        Input #1, data
        file_data = data
        huruf = Left(file_data, 1)
        angka = Val(Mid(file_data, 2, 20))
        lbl_faktur = huruf + CStr(angka + 1)
    Wend
    Close #1
End Sub

Public Sub loadDeposit(nobukti As String)
    is_new = False
    lbl_faktur = nobukti
    Dim rsJual As ADODB.Recordset
    Set rsJual = con.Execute("select * from tbrfiddeposit where nodeposit = '" & nobukti & "'")
    
    'Set rsJual = con.Execute("select a.rfid, b.`harga_jual` from tbaktif a, tbjual b, tbrfid c where a.tanggal = b.tglbukti and b.`nobukti` = c.nobukti  and a.rfid = c.rfid and kode = '2' and c.nobukti = '" & nobukti & "'")
        
    If Not rsJual.EOF Then
        Dim l As Integer
        Dim litem As ListItem
        l = 0
        Do While Not rsJual.EOF
            l = l + 1
            Set litem = lv_jual.ListItems.Add(, , l)
            litem.SubItems(1) = rsJual!rfid
            litem.SubItems(2) = rsJual!hargarfid
            txt_total = Format(txt_total + rsJual!hargarfid, "###,###,##0")
            rsJual.MoveNext
        Loop
    End If
    Set rsJual = Nothing
    txt_jumlah.Text = lv_jual.ListItems.count
'    Set rsJual = con.Execute("Select * from tbdeposit where nodeposit = '" & nobukti & "'")
'    txt_harga = rsJual!hargarfid
'    txt_jumlah = rsJual!deposit / rsJual!hargarfid
'    txt_total = rsJual!deposit
'    Set rsJual = Nothing
End Sub

Private Sub Form_KeyDown(key As Integer, Shift As Integer)

    If key = 112 Then
        If is_new Then
            If lv_jual.ListItems.count > 0 Then
                If validateRFID = True Then
                    If MsgBox("Apakah data yg diisi sudah lengkap?", vbYesNo, "Konfirmasi") = vbYes Then
                        Call print_bon
                    End If
                End If
            Else
                MsgBox "Faktur masih kosong"
            End If
        Else
            print_bon
        End If
    End If
    
    If key = 46 Then
        If Shift = 1 Then
            lv_jual.ListItems.Clear
            txt_jumlah.Text = 0
            txt_total.Text = 0
        Else
'            If lv_jual.ListItems.count > 0 Then
'                lv_jual.ListItems.Remove (lv_jual.SelectedItem.index)
'                Dim k As Integer
'                For k = 1 To lv_jual.ListItems.count
'                    lv_jual.ListItems.item(k).Text = k
'                Next
'                txt_jumlah.Text = lv_jual.ListItems.count
'                txt_total.Text = Format(txt_jumlah * txt_harga, "###,###,##0")
'            End If
            Call btn_Hapus_Click
        End If
    End If
    If key = 115 Then
        If MsgBox("Tutup form transaksi?", vbYesNo) = vbYes Then
            Unload Me
        End If
    End If
End Sub

Private Sub kosongkan()
    txt_Nama.Text = ""
    txt_jumlah.Text = 0
    txt_total.Text = 0
End Sub

Private Sub lv_jual_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
    Else
        txt_Nama = Chr(KeyAscii)
        txt_Nama.SetFocus
        txt_Nama.SelStart = Len(txt_Nama.Text)
    End If
End Sub

Private Sub txt_jumlah_Change()
'    If Len(txt_jumlah) < 4 And Val(txt_jumlah.Text) > 0 Then
'        Dim subtotal As String
'        subtotal = Format(hargaTiket * Val(txt_jumlah), "###,###,###")
'        txt_total.Text = Format(priceToNum(subtotal), "###,###,###")
'    Else
'        txt_jumlah = "0"
'    End If
    Dim x As Integer
    Dim subtotal As Long
    subtotal = 0
    If lv_jual.ListItems.count > 0 Then
        For x = 1 To lv_jual.ListItems.count
            subtotal = subtotal + lv_jual.ListItems(x).SubItems(2)
        Next
    End If
    txt_total.Text = Format(subtotal, "###,###,##0")
    
End Sub

Private Sub txt_jumlah_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        If Len(txt_jumlah) > 3 Then
            txt_jumlah = "0"
            Exit Sub
        End If
            
        If Val(txt_jumlah.Text) < 1 Then
            MsgBox "Jumlah tidak valid"
            Exit Sub
        End If
        
        Dim subtotal As String
        subtotal = Format(hargaTiket * Val(txt_jumlah), "###,###,###")
        
        txt_Nama.SetFocus
    End If
End Sub

Private Sub txt_jumlah_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(txt_jumlah.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub

Private Sub txt_nama_KeyDown(key As Integer, Shift As Integer)

    If key = 13 Then
        If Len(txt_Nama.Text) = 10 Then
            Call tambah
        Else
            MsgBox "Masukkan nomor RFID yg benar"
            txt_Nama.Text = ""
        End If
    End If
    
End Sub

Public Sub nextFaktur()
    Dim namafile, huruf As String
    Dim angka As Long
    Me.Enabled = True
    huruf = Left(lbl_faktur, 1)
    angka = Val(Mid(lbl_faktur, 2, 20))
    
    namafile = App.Path & "\fakturdeposit.txt"
    Open namafile For Output As #1
    Print #1, lbl_faktur
    Close #1
    
    lbl_faktur = huruf + CStr(angka + 1)
    lv_jual.ListItems.Clear
    txt_total = "0"
    kosongkan
    Form_List_Deposit.refreshlist
    
End Sub

Sub tambah()
    
    If txt_Nama.Text <> "" And Len(txt_Nama.Text) > 5 And cekRFID = True Then
        Dim mitem As ListItem
        Dim query As String
        query = "select a.rfid, b.`harga_jual` from tbaktif a, tbjual b, tbrfid c where a.tanggal = b.tglbukti and b.`nobukti` = c.nobukti  and a.rfid = c.rfid and kode = '2' and a.rfid = '" & txt_Nama.Text & "'"
        If isInTBAktif(txt_Nama.Text) Then
            'get deposit
            Dim rsDep As ADODB.Recordset
            Set rsDep = con.Execute(query)
            If Not rsDep.EOF Then
                Set mitem = lv_jual.ListItems.Add(, , lv_jual.ListItems.count + 1)
                mitem.SubItems(1) = txt_Nama.Text
                mitem.SubItems(2) = rsDep!harga_jual
            End If
        Else
            MsgBox ("RFID " & txt_Nama.Text & " tidak ditemukan")
        End If
        
    Else
        MsgBox ("RFID tidak valid atau sudah terdaftar")
    End If
    txt_Nama.Text = ""
    txt_Nama.SetFocus
    If lv_jual.ListItems.count > txt_jumlah.Text Then
        txt_jumlah.Text = lv_jual.ListItems.count
    End If
End Sub

Function cekRFID() As Boolean
    cekRFID = True
    If lv_jual.ListItems.count > 0 Then
        Dim i As Integer
        For i = 1 To lv_jual.ListItems.count
            If txt_Nama.Text = lv_jual.ListItems.item(i).SubItems(1) Then
                cekRFID = False
            End If
        Next
    End If
End Function

Private Sub print_bon()

    If is_new Then
        Dim i As Integer
        i = 1
        Dim tanggal As String
        tanggal = Format(Now, "yyyy-mm-dd")
        
        ' find kdsuplier1
'        Dim kdsuplier_Temp As ADODB.Recordset
        

        For i = 1 To lv_jual.ListItems.count
            Call backupAktif(lv_jual.ListItems(i).SubItems(1), lbl_faktur.Caption)
            con.Execute ("delete from tbaktif where rfid = '" & lv_jual.ListItems(i).SubItems(1) & "'")
            deleteC1 lv_jual.ListItems(i).SubItems(1)
            con.Execute ("delete from tbreader where rfid = '" & lv_jual.ListItems(i).SubItems(1) & "'")
            'y
            
            con.Execute ("insert into tbrfiddeposit values('" & lbl_faktur & "', '" & lv_jual.ListItems(i).SubItems(1) & "','" & priceToNum(lv_jual.ListItems(i).SubItems(2)) & "')")
        Next
        
        con.Execute ("insert into tbdeposit values('" & lbl_faktur.Caption & "','" & username & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:mm:ss") & "'  ," & priceToNum(txt_total.Text) & " )")
        
        
              
'    Else
'        'con.Execute ("update bill set cash = " & tunai & ", bayar = " & priceToNum(txt_uang) & ", diskon = " & priceToNum(txt_diskon) & " where nobukti = '" & txt_bon & "'")
'        con.Execute ("update bill set cash = " & tunai & ", bayar = " & priceToNum(txt_uang) & ", diskon = " & priceToNum(txt_diskon) & ", jumlah = " & priceToNum(txt_subTotal) & " ,total = " & priceToNum(txt_grandTotal) & " where nobukti = '" & txt_bon & "'")
'
'        Set rsDis = con.Execute("select * from tbdiskon where nobukti = '" & txt_bon & "'")
'        If Val(txt_diskon) > 0 Then
'            If rsDis.EOF = True Then
'                con.Execute ("insert into tbdiskon values('" & txt_bon & "', '" & dis_spv & "', '" & dis_status & "', '" & dis_cust & "', " & priceToNum(txt_diskon) & ")")
'            Else
'                con.Execute ("update tbdiskon set supervisor = '" & dis_spv & "', customer = '" & dis_cust & "', status = '" & dis_status & "', nilai = " & priceToNum(txt_diskon) & " where nobukti = '" & txt_bon.Text & "'")
'            End If
'        Else
'            If Not rsDis.EOF Then
'                con.Execute ("Delete from tbdiskon where nobukti = '" & txt_bon & "'")
'            End If
'        End If
'
'        Dim rsDeposit As ADODB.Recordset
'        Set rsDeposit = con.Execute("Select * from tbjual where kode = '2' and nobukti = '" & txt_bon & "'")
'        If Val(txt_LoanedRFID) > 0 Then
'            If rsDis.EOF = True Then
'                con.Execute ("insert into tbjual values('" & txt_bon & "','" & Format(tanggal, "yyyy-mm-dd") & "','" & 2 & "', '" & "Deposit" & "'," & harga_Deposit & "," & priceToNum(txt_BesarDeposit) / harga_Deposit & ", '2')")
'            Else
'                con.Execute ("update tbjual set jumlah_jual = '" & priceToNum(txt_BesarDeposit) / harga_Deposit & "' where kode = '2' and nobukti = '" & txt_bon & "'")
'            End If
'        Else
'            If Not rsDeposit.EOF Then
'                con.Execute ("Delete from tbjual where kode = '2' and nobukti = '" & txt_bon & "'")
'            End If
'        End If
    End If
    
    Set rsBill = con.Execute("select * from tbdeposit where nodeposit = '" & lbl_faktur.Caption & "'")
'    Set rsJual = con.Execute("select * from tbjual where nobukti = '" & txt_bon & "'")
'    If rsJual.EOF Then
'        MsgBox "data tidak ditemukan"
'        Exit Sub
'    End If
'    rsJual.MoveFirst
'    Dim temp_value As Boolean
'    temp_value = False
    
    If is_new = False Then
        If MsgBox("Cetak struk pembelian?", vbYesNo) = vbYes Then temp_value = True
    End If
        
    If is_new = True Or temp_value = True Then
        Printer.CurrentX = 0
        Printer.CurrentY = 0
        Printer.Font = "dotumche"
        'Printer.Font = tempFont
        Printer.FontSize = 18
        Printer.FontBold = True
        'Printer.Print Tab(2); Printer.PaintPicture(App.Path & "\CHIP.jpg");
        'Printer.Print Tab(2); "CHRISTINE HAKIM";
        
'        Printer.PaintPicture LoadPicture(App.Path & "\chip.jpg"), 300, 0, 2774, 1510
'        Printer.Print Tab(1); "                                                                  ";
'        Printer.Print Tab(1); "                                                                  ";
'        Printer.Print Tab(1); "                                                                  ";
'        Printer.Print Tab(1); "                                                                  ";
'        Printer.Print Tab(1); "                                                                  ";
        
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.Print Tab(1); "                                                            "; 'baris kosong
        Printer.Print Tab(1); Setting_Object("Alamat1");
        Printer.Print Tab(1); "No. FAKTUR : "; lbl_faktur.Caption
        Printer.Print Tab(1); Format(rsBill!tanggal, "dd-MM-yyyy"); "  "; rsBill!jam;
        Printer.Print Tab(1); "Nama Kasir : "; rsBill!kasir;
        Printer.Print Tab(1); "                                                                  ";
        Printer.Print Tab(1); "------------------------------------------------------------------";
'        Do While Not rsJual.EOF
'            Printer.Print Tab(1); rsJual!nama_Barang
'            Dim bayar As Long
'            bayar = Val(rsJual!jumlah_jual) * Val(rsJual!harga_jual)
'            Printer.Print Tab(2); rsJual!jumlah_jual; Tab(9); "x"; Tab(21 - Len(Format(rsJual!harga_jual, "###,###,##0"))); Format(rsJual!harga_jual, "###,###,##0"); Tab(35 - Len(Format(bayar, "###,###,##0"))); Format(bayar, "###,###,##0")
'            rsJual.MoveNext
'        Loop
        Printer.Print Tab(1); "                                                                  ";
        Printer.FontSize = 10
        Printer.Print Tab(1); "Jumlah Kartu RFID : "; txt_jumlah.Text; " lembar"
        Printer.Print Tab(1); "Total Deposit : "; Tab(20); "Rp."; Tab(35 - Len(Format(txt_total.Text, "###,###,##0"))); Format(txt_total.Text, "###,###,##0")
'        If priceToNum(txt_diskon) > 0 Then
'            Printer.Print Tab(1); "Diskon"; Tab(20); "Rp."; Tab(35 - Len(Format(txt_diskon, "###,###,##0"))); Format(txt_diskon, "###,###,##0")
'        End If
        'Printer.Print Tab(1); "Pajak Restoran 10%"; Tab(20); "Rp."; Tab(35 - Len(Format(txt_ppn, "###,###,##0"))); Format(txt_ppn, "###,###,##0")
        Printer.Print Tab(1); "------------------------------------------------------------------";
'        Printer.FontSize = 12
'        Printer.Print Tab(2); "Grand Total"; Tab(15); "Rp."; Tab(30 - Len(Format(txt_grandTotal, "###,###,##0"))); Format(txt_grandTotal, "###,###,##0")
'
'        Printer.CurrentX = 0
'        Printer.FontSize = 10
'        Printer.Print Tab(3); "                                                             ";
'        If (tunai = 1) Then
'            Printer.Print Tab(1); "Jumlah Uang"; Tab(15); "Rp. "; Tab(31 - Len(Format(txt_uang, "###,###,##0"))); Format(txt_uang, "###,###,##0");
'            Printer.Print Tab(1); "Kembalian  "; Tab(15); "Rp. "; Tab(31 - Len(Format(txt_kembali, "###,###,##0"))); Format(txt_kembali, "###,###,##0");
'        Else
'            Printer.Print Tab(3); "-NON TUNAI-";
'        End If
        Printer.Print Tab(1); "                                                             ";
'        'Printer.FontSize = 10
        Printer.Print Tab(1); "                                                             ";
'        'Printer.FontSize = 10
        Printer.Print Tab((38 - Len("*Periksalah kembali uang anda*")) / 2); "*Periksalah kembali uang anda*";
        Printer.Print Tab((38 - Len("*sebelum meninggalkan kasir*")) / 2); "*sebelum meninggalkan kasir*";
    
        Printer.EndDoc
    End If
      
    Close #1
    
    
    If is_new Then
        nextFaktur
    'Else
        Form_List_Deposit.refreshlist
    End If
    
End Sub

Function validateRFID() As Boolean
    validateRFID = False
    Dim x As Boolean
    Dim i As Integer
    For i = 1 To lv_jual.ListItems.count
        validateRFID = isInTBAktif(lv_jual.ListItems(i).SubItems(1))
        If validateRFID = False Then
            MsgBox ("RFID " & lv_jual.ListItems(i).SubItems(1) & " tidak ditemukan")
            Exit For
        End If
    Next
End Function

'Function cekRFID2(inRFID As String) As Boolean
'    cekRFID2 = False
'    Dim rsAktif As ADODB.Recordset
'    Set rsAktif = con.Execute("select * from tbaktif where rfid = '" & inRFID & "'")
'    If Not rsAktif.EOF Then cekRFID2 = True
'End Function
Private Sub txt_nama_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 2)
End Sub
