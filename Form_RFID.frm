VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_RFID 
   BackColor       =   &H00F0F0FF&
   Caption         =   "RFID"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   6495
   Begin VB.CommandButton btn_Cancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton btn_Save 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Save"
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton btn_Hapus 
      BackColor       =   &H00D0D0FF&
      Caption         =   "Hapus"
      Height          =   615
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton btn_Tambah 
      BackColor       =   &H00D0D0FF&
      Caption         =   "Tambah"
      Height          =   615
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   975
   End
   Begin MSComctlLib.ListView lv_RFID 
      Height          =   4815
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   4755
      _ExtentX        =   8387
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
   Begin VB.Label lbl_JumlahTiket 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Tiket : "
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
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lbl_Nobukti 
      BackStyle       =   0  'Transparent
      Caption         =   "W12345"
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
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form_RFID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim is_new As Boolean

Public Sub Start(isnew As Boolean, num_Ticket As Integer, nofaktur As String)
    is_new = isnew
    lbl_Nobukti.Caption = nofaktur
    lbl_JumlahTiket.Caption = num_Ticket
    lv_RFID.ListItems.Clear
    If is_new Then
        Dim i As Integer
        Dim iitem As ListItem
        For i = 1 To Form_Penjualan.lv_RFID.ListItems.count
            Set iitem = lv_RFID.ListItems.Add(, , Form_Penjualan.lv_RFID.ListItems(i).Text)
            iitem.SubItems(1) = Form_Penjualan.lv_RFID.ListItems(i).SubItems(1)
        Next
    Else
        Dim rsRFID As ADODB.Recordset
        Dim jitem As ListItem
        Set rsRFID = con.Execute("select * from tbrfid where nobukti = '" & lbl_Nobukti.Caption & "'")
        Do While Not rsRFID.EOF
            Set jitem = lv_RFID.ListItems.Add(, , lv_RFID.ListItems.count + 1)
            jitem.SubItems(1) = rsRFID!rfid
            rsRFID.MoveNext
        Loop
        Set rsRFID = Nothing
    End If
End Sub

Private Sub btn_Cancel_Click()
    Unload Me
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
    
    If is_new = True Then
        Dim n As Integer
        Dim nitem As ListItem
        Form_Penjualan.lv_RFID.ListItems.Clear
        For n = 1 To lv_RFID.ListItems.count
            Set nitem = Form_Penjualan.lv_RFID.ListItems.Add(, , lv_RFID.ListItems(n).Text)
            nitem.SubItems(1) = lv_RFID.ListItems(n).SubItems(1)
        Next
    Else
        Dim rsCompare As ADODB.Recordset
        m = lv_RFID.ListItems.count
        Dim flagX() As Integer
        Dim foundX As Boolean
        ReDim Preserve flagX(m)
        Set rsCompare = con.Execute("Select * from tbrfid where nobukti = '" & lbl_Nobukti.Caption & "'")
        Do While Not rsCompare.EOF
            foundX = False
            For m = 1 To lv_RFID.ListItems.count
                If rsCompare!rfid = lv_RFID.ListItems(m).SubItems(1) Then
                    flagX(m) = 1
                    foundX = True
                    Exit For
                End If
            Next
            If foundX = False Then
                con.Execute ("Delete from tbrfid where nobukti = '" & lbl_Nobukti.Caption & "' and rfid = '" & rsCompare!rfid & "'")
                Call backupAktif(rsCompare!rfid, "perubahan kartu - RFID")
                con.Execute ("delete from tbaktif where rfid = '" & rsCompare!rfid & "'")
                deleteC1 rsCompare!rfid
                con.Execute ("delete from tbreader where rfid = '" & rsCompare!rfid & "'")
                'y
            End If
            rsCompare.MoveNext
        Loop
        Set rsCompare = con.Execute("select tanggal, jam, nobukti from bill where nobukti = '" & lbl_Nobukti.Caption & "'")
        For m = 1 To lv_RFID.ListItems.count
            If flagX(m) = 0 Then
                
                con.Execute ("insert into tbaktif values('" & lv_RFID.ListItems(m).SubItems(1) & "','" & Format(rsCompare!tanggal, "yyyy-mm-dd") & "','" & rsCompare!jam & "','1','" & rsCompare!nobukti & "')")
                con.Execute ("insert into tbreader (rfid) values ('" & lv_RFID.ListItems(m).SubItems(1) & "')")
                pushC1 lv_RFID.ListItems(m).SubItems(1)
                'y
                con.Execute ("insert into tbrfid values('" & lbl_Nobukti.Caption & "','" & lv_RFID.ListItems(m).SubItems(1) & "')")
            End If
        Next
        Set rsCompare = Nothing
    End If
    Unload Me

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
