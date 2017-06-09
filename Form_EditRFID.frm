VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_EditRFID 
   BackColor       =   &H00A0E0FF&
   Caption         =   "Edit"
   ClientHeight    =   4455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_Save 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Save"
      Height          =   615
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton btn_Cancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txt_status 
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
      Left            =   2520
      TabIndex        =   3
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txt_RFID 
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
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker dt_Tanggal 
      Height          =   480
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   847
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   37748739
      CurrentDate     =   42922
   End
   Begin MSComCtl2.DTPicker dt_Jam 
      Height          =   480
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   847
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "HH:mm:ss"
      Format          =   37748739
      UpDown          =   -1  'True
      CurrentDate     =   42922
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   480
      TabIndex        =   9
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Jam"
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
      Left            =   480
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
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
      Left            =   480
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RFID"
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
      Left            =   480
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form_EditRFID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Cancel_Click()
    Unload Me
End Sub

Public Sub init(numRFID As String, tanggalRFID As Date, jamRFID As String, statusRFID As String)
    txt_RFID.Text = numRFID
    dt_Tanggal.Value = tanggalRFID
    dt_Jam.Value = jamRFID
    txt_status.Text = statusRFID
End Sub

Private Sub btn_Save_Click()
    Call backupAktif(txt_RFID.Text, "perubahan - EditRFID")
    con.Execute ("Update tbaktif set tanggal = '" & Format(dt_Tanggal.Value, "yyyy-mm-dd") & "', jam = '" & Format(dt_Jam.Value, "HH:mm:ss") & "', status = '" & txt_status.Text & "', keterangan = '" & username & "' where rfid = '" & txt_RFID.Text & "'")
    Form_List_RFID.reload_list
    Unload Me
End Sub

Private Sub txt_status_KeyPress(KeyAscii As Integer)
    txt_status.Text = ""
    Select Case KeyAscii
        Case 48 To 49, 8 ' 0-1, backspace
        'Let these key codes pass through
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub

Private Sub txt_status_LostFocus()
    If txt_status.Text = "" Then txt_status.Text = "0"
End Sub
