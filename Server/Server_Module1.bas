Attribute VB_Name = "Server_Module1"
Public con As New ADODB.Connection
Public status
Public username As String
Public Setting_Object As Object
Public StatusC1_1 As Boolean
Public StatusC1_2 As Boolean
Public StatusC1_3 As Boolean

Public Sub connect()
    con.ConnectionString = "Provider=MSDASQL.1;Password=" & Setting_Object.Item("DB_Pw") & ";Persist Security Info=True;User ID=" & Setting_Object.Item("DB_Id") & ";Data Source=" & Setting_Object.Item("DB_Name")
    con.Open
End Sub

Function validateKey(KeyAscii As Integer, lim As Integer)
    If lim = 1 Then 'for number
        Select Case KeyAscii
            Case 48 To 57, 44, 45, 46, 8 '0-9, comma, minus, dot and backspace
            'Let these key codes pass through
            Case Else
            'All others get trapped
            KeyAscii = 0 ' set ascii 0 to trap others input
        End Select
    ElseIf lim = 2 Then 'for password/kode
        Select Case KeyAscii
            Case 65 To 90, 48 To 57, 97 To 122, 8 ' A-Z, 0-9, a-z and backspace
            'Let these key codes pass through
            Case Else
            'All others get trapped
            KeyAscii = 0 ' set ascii 0 to trap others input
        End Select
    ElseIf lim = 3 Then 'for general textbox
        Select Case KeyAscii
            Case 8, 32 To 38, 40 To 58, 60 To 126 ' Allow all except ' and ;
            'Let these key codes pass through
            Case Else
            'All others get trapped
            KeyAscii = 0 ' set ascii 0 to trap others input
        End Select
    End If
    validateKey = KeyAscii
End Function
