VERSION 5.00
Begin VB.Form form_TestBon 
   Caption         =   "Bon"
   ClientHeight    =   9750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   9750
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tunai 
      Height          =   285
      Left            =   6360
      TabIndex        =   1
      Text            =   "1"
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "form_TestBon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBill, rsJual As ADODB.Recordset

Private Sub Command1_Click()
    Set rsBill = con.Execute("select * from bill where nobukti = '" & Form_Print.txt_bon.Text & "'")
    If priceToNum(Form_Print.txt_diskon.Text) > 0 Then
        'con.Execute ("insert into tbdiskon values('" & Form_Print.txt_bon.Text & "', '" & Form_Print.dis_spv & "', '" & Form_Print.dis_status & "', '" & Form_Print.dis_cust & "', " & priceToNum(txt_diskon.Text) & ")")
           
        Font = "Times new roman"
        FontSize = 10
        Print Tab(10); Format(Now, "dd-MM-yyyy  hh:mm:ss");
        Print Tab(10); "No Faktur"; Tab(25); ":  "; Form_Print.txt_bon.Text
        Print Tab(10); "Supervisor"; Tab(25); ":  "; Form_Print.dis_spv
        Print Tab(10); "Status"; Tab(25); ":  "; Form_Print.dis_status
        Print Tab(10); "Customer"; Tab(25); ":  "; Form_Print.dis_cust
        Print Tab(10); "Diskon"; Tab(25); ":  Rp. "; Format(Form_Print.txt_diskon.Text, "###,###,##0")
        Print Tab(3); "                                                            ";
        Print Tab(3); "                                                            ";
        Print Tab(3); "                                                            ";
        Printer.EndDoc
    End If

    ''end diskon
    
    Set rsJual = con.Execute("select * from tbjual where nobukti = '" & Form_Print.txt_bon.Text & "'")
    If rsJual.EOF Then
        MsgBox "data tidak ditemukan"
        Exit Sub
    End If
    rsJual.MoveFirst
    
    Font = "times new roman"
    'CurrentX = 0
    'CurrentY = 0
    FontSize = 18
    FontBold = True
    Print Tab(3); " KRIPIK BALADO";
    Print Tab(2); "CHRISTINE HAKIM";
    FontSize = 10
    FontBold = False
    Print Tab(3); "                                                            "; 'baris kosong
    Print Tab(3); "Jl. Adinegoro No. 11A Padang";
    Print Tab(3); "                                                             ";
    Print Tab(3); "No. FAKTUR :"; Form_Print.txt_bon.Text
    Print Tab(3); "Nama Kasir :"; rsBill!kasir;
    Print Tab(3); Format(rsBill!tanggal, "dd-MM-yyyy"); "  "; rsBill!jam;
    Print Tab(3); "                                                                  ";
    Do While Not rsJual.EOF
        Print Tab(3); rsJual!nama_barang
        Dim bayar As Long
        bayar = Val(rsJual!jumlah_jual) * Val(rsJual!harga_jual)
        Print Tab(5); rsJual!jumlah_jual; Tab(11); "x"; Tab(28 - Len(Format(rsJual!harga_jual, "###,###,###"))); Format(rsJual!harga_jual, "###,###,###"); Tab(42 - Len(Format(bayar, "###,###,###"))); Format(bayar, "###,###,###")
        rsJual.MoveNext
    Loop
    Print Tab(3); "------------------------------------------------------------------";
    FontSize = 10
    Print Tab(5); "Total"; Tab(30); ": Rp."; Tab(42 - Len(Format(Form_Print.txt_subTotal.Text, "###,###,###"))); Format(Form_Print.txt_subTotal.Text, "###,###,###")
    If priceToNum(Form_Print.txt_diskon.Text) > 0 Then
        Print Tab(5); "Diskon"; Tab(30); ": Rp."; Tab(42 - Len(Format(Form_Print.txt_diskon.Text, "###,###,###"))); Format(Form_Print.txt_diskon.Text, "###,###,###")
    End If
    ''test
    Print Tab(5); "Pajak Restoran 10%"; Tab(30); ": Rp."; Tab(42 - Len(Format(Form_Print.txt_ppn.Text, "###,###,###"))); Format(Form_Print.txt_ppn.Text, "###,###,###")
    FontSize = 12
    Print Tab(2); "Grand Total"; Tab(30 - Len(Format(Form_Print.txt_grandTotal.Text, "###,###,###"))); "Rp.  " & Format(Form_Print.txt_grandTotal.Text, "###,###,###")
    ' ; Tab(18); ": Rp."
''    Dim diskon_total As Long
''    diskon_total = priceToNum(txt_subtotal) - priceToNum(txt_diskon)
''diskon
''    Printer.Print Tab(8); "Total"; Tab(16); ": Rp."; Format(diskon_total, "###,###,###")
''nyelip di sini
    CurrentX = 0
    FontSize = 10
    Print Tab(3); "                                                             ";
    If (tunai = 1) Then
        Print Tab(3); "Jumlah Uang  Rp. "; Tab(31 - Len(Format(Form_Print.txt_uang.Text, "###,###,##0"))); Format(Form_Print.txt_uang.Text, "###,###,##0");
        Print Tab(3); "Kembalian    Rp. "; Tab(31 - Len(Format(Form_Print.txt_kembali.Text, "###,###,##0"))); Format(Form_Print.txt_kembali.Text, "###,###,##0");
    Else
        Print Tab(3); "-NON TUNAI-";
        tunai = 0
    End If
    Print Tab(3); "                                                             ";
    FontSize = 10
    Print Tab(3); "                                                             ";
    Print Tab(3); "Customer Service: (0751)483518";
    Print Tab(3); "HP Pemesanan: 0811 668 5000";
    Print Tab(3); "Website: www.christinehakimideapark.com";
    Print Tab(3); "                                                             ";
    FontSize = 8
    Print Tab(3); "*Barang yang sudah dibeli tidak dapat dikembalikan*";
    Print Tab(25); "*Terima Kasih*";
End Sub

