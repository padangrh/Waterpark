VERSION 5.00
Begin VB.Form Form_Diskon 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Diskon"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4755
   FillColor       =   &H0000FF00&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btn_cancel 
      Caption         =   "Cancel"
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
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton btn_ok 
      Caption         =   "OK"
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
      Left            =   3240
      TabIndex        =   10
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txt_diskon 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2040
      TabIndex        =   9
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox txt_customer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2040
      TabIndex        =   7
      Top             =   2040
      Width           =   2535
   End
   Begin VB.ComboBox cb_status 
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
      ItemData        =   "Form_Diskon.frx":0000
      Left            =   2040
      List            =   "Form_Diskon.frx":0013
      TabIndex        =   5
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txt_password 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txt_spv 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Diskon"
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
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Customer"
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
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Status"
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
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Password"
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
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Supervisor"
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
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form_Diskon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_cancel_Click()
    Unload Me
End Sub

Private Sub btn_ok_Click()
    
    If cek_Status = False Then
        MsgBox "Status tidak valid"
        Exit Sub
    End If
    
    If priceToNum(Form_Print.txt_subTotal) < priceToNum(txt_diskon.Text) Then
        MsgBox "Diskon yg diberikan lebih besar dari harga barang"
    Else
        Dim rsUser As ADODB.Recordset
        Set rsUser = con.Execute("select * from tblogin where userid = '" & txt_spv & "'")
        If rsUser.EOF Or rsUser.BOF Then
            MsgBox "Supervisor tidak terdaftar"
            'Exit Function
            Exit Sub
        End If
        
        If rsUser!posisi = "Karyawan" Then
            MsgBox "Hanya supervisor yang bisa memberi diskon"
            'Exit Function
            Exit Sub
        End If
        
        If rsUser!pass <> txt_password Then
            MsgBox "Password salah"
            'Exit Function
            Exit Sub
        End If
        
        Form_Print.txt_diskon = Format(txt_diskon, "###,###,##0")
        
        Form_Print.diskon_query
        Form_Print.hitung
    
    ''commit database diskon
    ''con.Execute ("insert into tbdiskon values('" & Form_Print.txt_bon & "', '" & txt_spv & "', '" & cb_status.Text & "', '" & txt_customer & "', " & priceToNum(txt_diskon) & ")")
    
    ''Printer.Font = "Times new roman"
    ''Printer.FontSize = 12
    ''Printer.Print Tab(4); Format(Now, "dd-MM-yyyy  hh:mm:ss");
    ''Printer.Print Tab(4); "No Faktur"; Tab(18); ": "; Form_Print.txt_bon
    ''Printer.Print Tab(4); "Supervisor"; Tab(18); ": "; txt_spv
    ''Printer.Print Tab(4); "Status"; Tab(18); ": "; cb_status.Text
    ''Printer.Print Tab(4); "Customer"; Tab(18); ": "; txt_customer
    ''Printer.Print Tab(4); "Diskon"; Tab(18); ": Rp."; txt_diskon
    ''Printer.EndDoc
    
    ''sampai disini dan kalaupun dicancel, tetap ada di database dan 1 faktur bisa memperoleh lebih dari 1 diskon
        If priceToNum(Form_Print.txt_uang.Text) > 0 Then
            Form_Print.txt_kembali = Format(priceToNum(Form_Print.txt_uang) - priceToNum(Form_Print.txt_grandTotal.Text), "###,###,##0")
        End If
        Unload Me
        Form_Print.txt_uang.SetFocus
    End If
End Sub
'End Function


Private Sub Form_unload(cancel As Integer)
    Form_Print.Enabled = True
End Sub

Private Sub txt_diskon_LostFocus()
    txt_diskon = Format(txt_diskon, "###,###,##0")
End Sub

Private Sub txt_spv_keypress(key As Integer)
    If key = 13 Then
        txt_password.SetFocus
    End If
End Sub

Private Sub txt_password_keypress(key As Integer)
    If key = 13 Then
        cb_status.SetFocus
    End If
End Sub

Private Sub cb_status_keypress(key As Integer)
    If key = 13 Then
        txt_customer.SetFocus
    End If
End Sub

Private Sub txt_customer_keypress(key As Integer)
    If key = 13 Then
        txt_diskon.SetFocus
    End If
End Sub

Private Sub txt_diskon_keypress(key As Integer)
    If key = 13 Then
        btn_ok.SetFocus
    End If
End Sub

Private Function cek_Status() As Boolean
    cek_Status = False
    Dim i As Integer
    Do While i < cb_status.ListCount
        If Trim(UCase(cb_status.Text)) = Trim(UCase(cb_status.List(i))) Then
            cek_Status = True
        End If
        i = i + 1
    Loop
End Function
