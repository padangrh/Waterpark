VERSION 5.00
Begin VB.Form frmlogin 
   BackColor       =   &H0000C000&
   Caption         =   "Security"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5520
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   2895
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Commandbatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Commandlogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtpass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtuser 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "User Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5280
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Login"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandBatal_Click()
    End
End Sub

Private Sub CommandLogin_Click()
On Error GoTo salah:
    Dim Rec As ADODB.Recordset
    Set Rec = con.Execute("select * from tblogin where UserID='" & Trim(txtuser.Text) & "'")
    If Not Rec.EOF Then
        If UCase(Rec.Fields("UserID")) = UCase(Trim(txtuser)) And Rec.Fields("pass") = Trim(txtpass) Then
'            username = Rec!userid
            If Rec!posisi = "Master" Then
                Unload Me
                frmMain.Show
            Else
                MsgBox "Hanya admin yg boleh masuk"
            End If
        Else
            MsgBox "Nama user atau password anda tidak cocok!"
            txtuser.SetFocus
        End If
    Else
      MsgBox "Nama user atau password anda tidak cocok!"
      txtuser.SetFocus
    End If
    Exit Sub

'    frmMain.Visible = True
'    frmMain.Enabled = True
'
'    Unload Me
    Exit Sub
salah:
MsgBox "Periksa komputer server hidup atau tidak, kabel internet tercolok di komputer atau tidak, coba restart modem"
End Sub

Private Sub Form_Activate()
    Dim temp_stringX, File_StringX As String
    Dim fso As FileSystemObject
    
    Set fso = New FileSystemObject
    If fso.FileExists(App.Path & "\Settings.json") Then
        'load JSON file
        temp_stringX = ReadTextFile(App.Path & "\Settings.json")
        'Decode file
        File_StringX = Base64DecodeString(temp_stringX)
        'Generate variables
        Set Setting_Object = JSON.parse(File_StringX)
        
        connect
     Else
        MsgBox "Settings file is missing."
        Unload Me
    End If
    txtuser.SetFocus
End Sub

Private Sub txtpass_keyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        CommandLogin_Click
    End If
End Sub

Private Sub Form_unload(Cancel As Integer)
    'con.Close
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
    KeyAscii = validateKey(KeyAscii, 2)
End Sub

Private Sub txtuser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtpass.SetFocus
    KeyAscii = validateKey(KeyAscii, 2)
End Sub

Public Function ReadTextFile(sFilePath As String) As String
   On Error Resume Next
   
   Dim handle As Integer
   If LenB(Dir$(sFilePath)) > 0 Then
   
      handle = FreeFile
      Open sFilePath For Binary As #handle
      ReadTextFile = Space$(LOF(handle))
      Get #handle, , ReadTextFile
      Close #handle
      
   End If
End Function
