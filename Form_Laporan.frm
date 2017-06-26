VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form_Laporan 
   Caption         =   "Laporan"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chk_Sampai 
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   360
      Width           =   255
   End
   Begin Crystal.CrystalReport cr 
      Left            =   6120
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton btn_hutang 
      Caption         =   "Laporan Hutang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6360
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton btn_pembayaran 
      Caption         =   "Laporan Pembayaran"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5760
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton btn_stok 
      Caption         =   "Laporan Stok"
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
      Left            =   6360
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton btn_pengeluaran 
      Caption         =   "Laporan Pengeluaran"
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
      Left            =   5760
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton btn_penjualan 
      Caption         =   "Laporan Penjualan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      TabIndex        =   4
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton btn_harian 
      Caption         =   "Laporan Harian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker dt_start 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
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
      Format          =   115343361
      CurrentDate     =   42810
   End
   Begin MSComCtl2.DTPicker dt_end 
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
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
      Format          =   115343361
      CurrentDate     =   42810
   End
   Begin VB.CommandButton btn_Terlaris 
      Caption         =   "Terlaris"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CommandButton btn_LaporanDeposit 
      Caption         =   "Laporan Deposit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      TabIndex        =   10
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CommandButton btn_Tiket 
      Caption         =   "Laporan Tiket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   12
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label Label2 
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
      Left            =   4440
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form_Laporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NO_DATE As Integer = 0
Const ONE_DAY As Integer = 1
Const DURATION As Integer = 2
Dim txt_sup_toggle As Boolean

'Private Sub btn_detailRekap_Click()
'    Call openReport("detailrekaptenant.rpt", "bill.tanggal", ONE_DAY)
'End Sub

Private Sub btn_harian_Click()
    If chk_Sampai.Value = 0 Then
        Call openReport("laporanharian.rpt", "bill.tanggal", ONE_DAY)
    Else
        Call openReport("Laporanharian.rpt", "bill.tanggal", DURATION)
    End If
    'Call openReport("laporanharian_test.rpt", "bill.tanggal", DURATION)
End Sub

Private Sub openReport(file_name As String, date_parameter As String, report_type As Integer)
    'cr.connect = "Provider=MSDASQL.1;Pwd=" & Setting_Object("DB_Pw") & ";Persist Security Info=True;User ID=" & Setting_Object("DB_Id") & ";Data Source=Data"
    cr.connect = "Provider=MSDASQL.1;Pwd=" & Setting_Object("DB_Pw") & ";Persist Security Info=True;User ID=" & Setting_Object("DB_Id") & ";Data Source=" & Setting_Object("DB_Name")
    
    cr.ReportFileName = App.Path + "\" + file_name
    
    If file_name = "laporanTenant.rpt" Then
        cr.SelectionFormula = "{ tbjual.tglbukti }= #" & Format(dt_start.Value, "yyyy-MM-dd") & "# and {tbjual.kdsuplier} = '" & txt_kode_supplier.Text & "'"
        cr.Formulas(0) = "tgl='" & Format(dt_start.Value, "dd/MM/yyyy") & "'"
        cr.Formulas(1) = "supplier='" & txt_nama_supplier.Text & "'"
    Else
        If report_type = ONE_DAY Then
            cr.SelectionFormula = "{" & date_parameter & "}= #" & Format(dt_start.Value, "yyyy-MM-dd") & "#"
            'cr.SelectionFormula = "{ tbjual.tglbukti }= #" & Format(dt_start.Value, "yyyy-MM-dd") & "#"
            cr.Formulas(0) = "tgl1='" & "Tanggal : " & Format(dt_start.Value, "dd/MM/yyyy") & "'"
            cr.Formulas(1) = "Header1='" & Setting_Object("Toko") & "'"
            cr.Formulas(2) = "Header2='" & Setting_Object("Alamat1") & "'"
            cr.Formulas(3) = "Header3='" & Setting_Object("Alamat2") & "'"
            
        ElseIf report_type = DURATION Then
            cr.SelectionFormula = "{" & date_parameter & "}>= #" & Format(dt_start.Value, "yyyy-MM-dd") & "# and {" & date_parameter & "}<= #" & Format(dt_end.Value, "yyyy-MM-dd") & "#"
            'cr.SelectionFormula = "{" & tanggal & "}>= #" & Format(dt_start.Value, "yyyy-MM-dd") & "# and {" & date_parameter & "}<= #" & Format(dt_end.Value, "yyyy-MM-dd") & "#"
            cr.Formulas(0) = "tgl1='" & "Dari Tanggal  : " & Format(dt_start.Value, "dd/MM/yyyy") & "'"
            cr.Formulas(1) = "tgl2='" & "S/D Tanggal : " & Format(dt_end.Value, "dd/MM/yyyy") & "'"
            cr.Formulas(2) = "Header1='" & Setting_Object("Toko") & "'"
            cr.Formulas(3) = "Header2='" & Setting_Object("Alamat1") & "'"
            cr.Formulas(4) = "Header3='" & Setting_Object("Alamat2") & "'"
        Else
            cr.Formulas(0) = "Header1='" & Setting_Object("Toko") & "'"
            cr.Formulas(1) = "Header2='" & Setting_Object("Alamat1") & "'"
            cr.Formulas(2) = "Header3='" & Setting_Object("Alamat2") & "'"
        End If
    End If
    cr.WindowState = crptMaximized
    cr.RetrieveDataFiles
    cr.Action = 1
    cr.reset
End Sub

Private Sub btn_LaporanDeposit_Click()
    If chk_Sampai.Value = 0 Then
        Call openReport("laporandeposit.rpt", "v_deposit1.tanggal", ONE_DAY)
    Else
        Call openReport("laporandeposit.rpt", "v_deposit1.tanggal", DURATION)
    End If
End Sub

'Private Sub btn_hutang_Click()
'    Call openReport("laporanhutang.rpt", "", NO_DATE)
'End Sub

'Private Sub btn_pembayaran_Click()
'    Call openReport("laporanpembayaran.rpt", "", NO_DATE)
'End Sub

'Private Sub btn_pengeluaran_Click()
'    Call openReport("laporanpengeluaran.rpt", "bill_beli.tanggal_lunas", DURATION)
'End Sub

Private Sub btn_penjualan_Click()
    If chk_Sampai.Value = 0 Then
        Call openReport("laporanpenjualan.rpt", "tbjual.tglbukti", ONE_DAY)
    Else
        Call openReport("laporanpenjualan.rpt", "tbjual.tglbukti", DURATION)
    End If
End Sub

'Private Sub btn_RekapTenant_Click()
'    Call openReport("rekaptenant.rpt", "tbjual.tglbukti", ONE_DAY)
'End Sub

'Private Sub btn_stok_Click()
'    Call openReport("laporanstok.rpt", "", NO_DATE)
'End Sub

'Private Sub btn_Tenant_Click()
'    If txt_kode_supplier <> "" And txt_nama_supplier <> "" Then
'        Call openReport("laporanTenant.rpt", "tbjual.tglbukti", ONE_DAY)
'    Else
'        MsgBox ("Nama dan Kode tidak boleh kosong")
'    End If
'End Sub

Private Sub btn_Terlaris_Click()
    Call openReport("terlaris.rpt", "", NO_DATE)
End Sub



Private Sub btn_Tiket_Click()
    If chk_Sampai.Value = 0 Then
        Call openReport("laporantiket.rpt", "tbjual.tglbukti", ONE_DAY)
    Else
        Call openReport("laporantiket.rpt", "tbjual.tglbukti", DURATION)
    End If
End Sub

Private Sub chk_Sampai_Click()
    If chk_Sampai.Value = 0 Then
        Label2.Enabled = False
        dt_end.Enabled = False
    Else
        Label2.Enabled = True
        dt_end.Enabled = True
    End If
End Sub

Private Sub dt_end_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 2)
End Sub

Private Sub dt_start_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 2)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    dt_start.Value = Now
    dt_end.Value = Now
    txt_sup_toggle = False
End Sub



