Attribute VB_Name = "Module1"
Public fMainForm As FrmMain
Public con As New ADODB.Connection
Public rsSupplier As ADODB.Recordset
Public status
Public username As String
Public Setting_Object As Object
Public StatusC1_1 As Boolean
Public StatusC1_2 As Boolean
Public StatusC1_3 As Boolean

' **********************************************
' Posiflex usbpd.dll DLL
' **********************************************
Public Declare Function WritePD _
    Lib "usbpd.dll" _
    (ByVal data As String, ByVal Length As Long) _
As Long

Public Declare Function WritePD80 _
    Lib "usbpd.dll" Alias "WritePD" _
    (ByRef data As Any, ByVal Length As Long) _
As Long

Public Declare Function PdState _
    Lib "usbpd.dll" _
    () _
As Long

Public Declare Function OpenUSBpd _
    Lib "usbpd.dll" _
    () _
As Long

Public Declare Function CloseUSBpd _
    Lib "usbpd.dll" _
    () _
As Long

Declare Sub Sleep Lib "kernel32" _
   (ByVal dwMilliseconds As Long)
   
Public Sub connect()
    
    
'    con.ConnectionString = "Provider=MSDASQL.1;Password=yuyu;Persist Security Info=True;User ID=root;Data Source=foodcourt1"
    'con.ConnectionString = "Provider=MSDASQL.1;Password=yuyu;Persist Security Info=True;User ID=root;Data Source=iceskating"
    con.ConnectionString = "Provider=MSDASQL.1;Password=" & Setting_Object.item("DB_Pw") & ";Persist Security Info=True;User ID=" & Setting_Object.item("DB_Id") & ";Data Source=" & Setting_Object.item("DB_Name")
    con.Open
End Sub

Public Function getSupplier(kode As String) As Boolean
    If rsSupplier Is Nothing Then
        Set rsSupplier = con.Execute("select * from tbsuplier")
    'Else
        'rsSupplier.MoveFirst
    End If
    
    Dim found As Boolean
    found = False
    If Not rsSupplier.EOF Then
        rsSupplier.MoveFirst
        Do While Not rsSupplier.EOF
            If kode = rsSupplier!kdsuplier Then
                found = True
                Exit Do
            End If
            rsSupplier.MoveNext
        Loop
    End If
    getSupplier = found
End Function

Public Function priceToNum(price As String) As Long
    price = Replace(price, ".", "")
    price = Replace(price, ",", "")
    priceToNum = Val(price)
End Function

Public Function isMaster() As Boolean
    isMaster = (status = "Master")
End Function

Public Function isSPV() As Boolean
    isSPV = (status = "Supervisor")
End Function

Public Function isInTBAktif(kodeRFID As String) As Boolean
    isInTBAktif = False
    Dim rsAktif As ADODB.Recordset
    Set rsAktif = con.Execute("select * from tbaktif where rfid = '" & kodeRFID & "'")
    If Not rsAktif.EOF Then isInTBAktif = True
    Set rsAktif = Nothing
End Function

Public Sub backupAktif(kodeRFID As String, reason As String)
    Dim rsAktif As ADODB.Recordset
    Set rsAktif = con.Execute("select * from tbaktif where rfid = '" & kodeRFID & "'")
    reason = rsAktif!keterangan & " - " & reason
    'editV2
    con.Execute ("insert into tbnonaktif (rfid, tanggal, jam, status, keterangan, userid) values ('" & rsAktif!rfid & "','" & Format(rsAktif!tanggal, "yyyy-mm-dd") & "','" & rsAktif!jam & "','" & rsAktif!status & "','" & reason & "','" & username & "')")
    Set rsAktif = Nothing
End Sub

Sub Main()
    Set fMainForm = New FrmMain
    fMainForm.Show
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

Public Sub pushC1(rfid As String)
    Dim rsReader As ADODB.Recordset
    Dim bcrtUser As Boolean
    confirmAllC1
    Set rsReader = con.Execute("select * from tbreader where rfid = '" & rfid & "'")
    If StatusC1_1 Then
        FrmMain.CZKEM1.CardNumber(0) = CLng(rfid)
        bcrtUser = FrmMain.CZKEM1.SetUserInfo(1, rsReader!id, "Tamu", "0751", 0, 1)
    End If
    If StatusC1_2 Then
        FrmMain.CZKEM2.CardNumber(0) = CLng(rfid)
        bcrtUser = FrmMain.CZKEM2.SetUserInfo(1, rsReader!id, "Tamu", "0751", 0, 1)
    End If
    If StatusC1_3 Then
        FrmMain.CZKEM3.CardNumber(0) = CLng(rfid)
        bcrtUser = FrmMain.CZKEM3.SetUserInfo(1, rsReader!id, "Tamu", "0751", 0, 1)
    End If
End Sub

Public Sub deleteC1(rfid As String)
    Dim rsReader As ADODB.Recordset
    Dim bdltUser As Boolean
    confirmAllC1
    Set rsReader = con.Execute("select * from tbreader where rfid = '" & rfid & "'")
    If StatusC1_1 Then
        bdltUser = FrmMain.CZKEM1.DeleteEnrollData(1, rsReader!id, 1, 12)
    End If
    If StatusC1_2 Then
        bdltUser = FrmMain.CZKEM2.DeleteEnrollData(1, rsReader!id, 1, 12)
    End If
    If StatusC1_3 Then
        bdltUser = FrmMain.CZKEM3.DeleteEnrollData(1, rsReader!id, 1, 12)
    End If
End Sub


Public Function confirmC1(ip As String) As Boolean
    Dim yy As Boolean
    yy = False
'    FrmMain.Winsock1.LocalPort = 0
    If FrmMain.Winsock1.State = sckClosed Then
        FrmMain.Winsock1.Protocol = sckTCPProtocol
        FrmMain.Winsock1.connect ip, 80
        FrmMain.Timer1.Enabled = True
        Do
            DoEvents
            If FrmMain.Winsock1.State = 7 Then
                FrmMain.Timer1.Enabled = False
                yy = True
            End If
        Loop While FrmMain.Timer1.Enabled
        FrmMain.Winsock1.Close
    ElseIf FrmMain.Winsock1.State = 7 Then
        yy = True
    End If
    confirmC1 = yy
End Function

Public Sub confirmAllC1()
    Dim xx As Boolean
    xx = False
    If StatusC1_1 Then
        xx = confirmC1(Setting_Object("C1_1"))
        If xx = False Then
            FrmMain.cmdC1_1.BackColor = &HFF&
            StatusC1_1 = xx
        End If
    End If
    If StatusC1_2 Then
        xx = confirmC1(Setting_Object("C1_2"))
        If xx = False Then
            FrmMain.cmdC1_2.BackColor = &HFF&
            StatusC1_2 = xx
        End If
    End If
    If StatusC1_3 Then
        xx = confirmC1(Setting_Object("C1_3"))
        If xx = False Then
            FrmMain.cmdC1_3.BackColor = &HFF&
            StatusC1_3 = xx
        End If
    End If
    DoEvents
End Sub

Public Sub refillC1(C1_ID As Integer)
    Dim i As Long
    Dim tempCZKEM As CZKEM
    Dim rsReader As ADODB.Recordset
    Dim dwEnrollNmber As Long
    Dim name As String
    Dim pwd As String
    Dim privilege As Long
    Dim sEnabled As Boolean
    
    If C1_ID = 1 Then
        Set tempCZKEM = FrmMain.CZKEM1
    ElseIf C1_ID = 2 Then
        Set tempCZKEM = FrmMain.CZKEM2
    ElseIf C1_ID = 3 Then
        Set tempCZKEM = FrmMain.CZKEM3
    End If
    
    Set rsReader = con.Execute("select * from tbreader")

    If tempCZKEM.ReadAllUserID(1) Then
        dwEnrollNmber = 0
        Do While tempCZKEM.GetAllUserInfo(CLng(1), dwEnrollNmber, name, pwd, privilege, sEnabled)
            If dwEnrollNmber <> 65534 Then
                bdltUser = tempCZKEM.DeleteEnrollData(1, dwEnrollNmber, 1, 12)
            End If
        Loop
    End If
        
    Do While Not rsReader.EOF
        tempCZKEM.CardNumber(0) = CLng(rsReader!rfid)
        bcrtUser = tempCZKEM.SetUserInfo(1, rsReader!id, "Tamu", "0751", 0, 1)
        rsReader.MoveNext
    Loop
    
    Set tempCZKEM = Nothing
    Set rsReader = Nothing
End Sub

Sub disableC1(C1_ID As Integer)
    Select Case C1_ID
        Case 1
            StatusC1_1 = False
            FrmMain.cmdC1_1.BackColor = &HFF&
        Case 2
            StatusC1_2 = False
            FrmMain.cmdC1_2.BackColor = &HFF&
        Case 3
            StatusC1_3 = False
            FrmMain.cmdC1_3.BackColor = &HFF&
    End Select
End Sub

